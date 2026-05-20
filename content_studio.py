#!/usr/bin/env python3
"""
The Wellness Collection — Content Studio
=========================================
A branded desktop tool for embedding narration audio into course presentations.

Runs a local web server and opens a beautiful browser-based interface
branded to The Wellness Collection / Gracefully Redefined.

Usage:
    python content_studio.py
    
    Then open http://localhost:5500 in your browser (opens automatically).
"""

import os
import sys
import gc
import json
import wave
import shutil
import zipfile
import tempfile
import re
import threading
import webbrowser
from http.server import HTTPServer, SimpleHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs
from pathlib import Path
import io
import subprocess

# Import core engine
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from embed_audio import embed_audio_into_pptx, find_audio_files, get_wav_duration_ms, count_slides, analyze_wav

PORT = int(os.environ.get('PORT', 5500))
UPLOAD_DIR = os.path.join(tempfile.gettempdir(), 'wellness_content_studio')
os.makedirs(UPLOAD_DIR, exist_ok=True)


# ---------------------------------------------------------------------------
# FFmpeg / LibreOffice helpers for MP4 export
# ---------------------------------------------------------------------------

def _find_ffmpeg():
    candidates = [shutil.which('ffmpeg')]
    if sys.platform == 'win32':
        candidates += [
            r'C:\ffmpeg\bin\ffmpeg.exe',
            r'C:\Program Files\ffmpeg\bin\ffmpeg.exe',
            r'C:\Program Files (x86)\ffmpeg\bin\ffmpeg.exe',
        ]
    return next((c for c in candidates if c and os.path.isfile(c)), None)


def _find_soffice():
    candidates = [shutil.which('soffice'), shutil.which('libreoffice')]
    if sys.platform == 'win32':
        candidates += [
            r'C:\Program Files\LibreOffice\program\soffice.exe',
            r'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
        ]
    elif sys.platform == 'darwin':
        candidates += ['/Applications/LibreOffice.app/Contents/MacOS/soffice']
    else:
        candidates += ['/usr/bin/libreoffice', '/usr/bin/soffice']
    return next((c for c in candidates if c and os.path.isfile(c)), None)


def _extract_slide_images(pptx_path, slides_dir):
    """Export each slide to PNG via PowerPoint COM (Windows) or LibreOffice."""
    os.makedirs(slides_dir, exist_ok=True)
    for f in os.listdir(slides_dir):
        if f.lower().endswith('.png'):
            os.remove(os.path.join(slides_dir, f))

    if sys.platform == 'win32':
        try:
            import win32com.client
            import pythoncom

            com_result = {'paths': [], 'done': False}

            def _com_export():
                try:
                    pythoncom.CoInitialize()
                    pptx_abs = os.path.abspath(pptx_path)
                    app = win32com.client.Dispatch('PowerPoint.Application')
                    app.Visible = False
                    prs = app.Presentations.Open(pptx_abs, WithWindow=False)
                    for i in range(1, prs.Slides.Count + 1):
                        img = os.path.abspath(os.path.join(slides_dir, f'slide_{i:04d}.png'))
                        prs.Slides(i).Export(img, 'PNG', 1920, 1080)
                        com_result['paths'].append(img)
                    prs.Close()
                    app.Quit()
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
                finally:
                    com_result['done'] = True

            t = threading.Thread(target=_com_export, daemon=True)
            t.start()
            t.join(timeout=90)

            if com_result['paths']:
                return com_result['paths']
        except ImportError:
            pass

    soffice = _find_soffice()
    if soffice:
        try:
            lo_profile = os.path.join(slides_dir, 'lo_profile')
            os.makedirs(lo_profile, exist_ok=True)
            user_install = 'file://' + lo_profile.replace('\\', '/')
            subprocess.run(
                [soffice,
                 f'-env:UserInstallation={user_install}',
                 '--headless', '--norestore', '--nofirststartwizard',
                 '--nojava',
                 '--convert-to', 'png',
                 '--outdir', slides_dir, pptx_path],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                timeout=300, check=False
            )
            gc.collect()
            pngs = sorted(
                (os.path.join(slides_dir, f)
                 for f in os.listdir(slides_dir)
                 if f.lower().endswith('.png')),
                key=lambda p: [int(c) if c.isdigit() else c.lower()
                               for c in re.split(r'(\d+)', os.path.basename(p))]
            )
            if pngs:
                return pngs
        except Exception:
            pass

    raise RuntimeError(
        'Could not render slide images. Install Microsoft PowerPoint (Windows) '
        'or LibreOffice and try again.'
    )


def _create_silence_wav(path, duration_seconds, sample_rate=48000):
    num_frames = int(duration_seconds * sample_rate)
    with wave.open(path, 'wb') as wf:
        wf.setnchannels(1)
        wf.setsampwidth(2)
        wf.setframerate(sample_rate)
        wf.writeframes(b'\x00' * num_frames * 2)


def _export_mp4(pptx_path, audio_dir, output_path, buffer_seconds=1.5,
                norm_mode='off', norm_target=-1.0):
    """
    Render PPTX + narration audio to MP4.
    Video: H.264 (libx264), slow preset, 8000k bitrate, 30 fps.
    Audio: AAC, 48000 Hz sample rate.
    """
    ffmpeg = _find_ffmpeg()
    if not ffmpeg:
        raise RuntimeError(
            'FFmpeg not found. Install FFmpeg and add it to your PATH. '
            'Download: https://ffmpeg.org/download.html'
        )

    work_dir = os.path.join(UPLOAD_DIR, 'mp4_work')
    slides_dir = os.path.join(work_dir, 'slides')
    os.makedirs(work_dir, exist_ok=True)

    num_slides = count_slides(pptx_path)
    audio_map = find_audio_files(audio_dir)

    slide_images = _extract_slide_images(pptx_path, slides_dir)
    slide_img = {i + 1: p for i, p in enumerate(slide_images)}

    silence_path = os.path.join(work_dir, 'silence.wav')
    _create_silence_wav(silence_path, buffer_seconds)

    # Merge all clips + silence into one WAV track
    audio_concat_txt = os.path.join(work_dir, 'audio_concat.txt')
    with open(audio_concat_txt, 'w', encoding='utf-8') as f:
        for s in range(1, num_slides + 1):
            if s in audio_map:
                for wav in audio_map[s]:
                    f.write("file '{}'\n".format(wav.replace(os.sep, '/')))
            f.write("file '{}'\n".format(silence_path.replace(os.sep, '/')))

    merged_wav = os.path.join(work_dir, 'merged_audio.wav')
    r = subprocess.run(
        [ffmpeg, '-y', '-f', 'concat', '-safe', '0', '-i', audio_concat_txt,
         '-ar', '48000', '-ac', '1', '-sample_fmt', 's16', merged_wav],
        stdout=subprocess.DEVNULL, stderr=subprocess.PIPE
    )
    if r.returncode != 0:
        raise RuntimeError(
            'FFmpeg audio merge failed: ' +
            r.stderr.decode('utf-8', errors='replace')[-600:]
        )

    # Build FFmpeg command: each slide as a timed loop input, concat filter, mux with audio
    encode_cmd = [ffmpeg, '-y']
    for s in range(1, num_slides + 1):
        img = slide_img.get(s) or slide_img[max(slide_img)]
        audio_dur = sum(
            get_wav_duration_ms(w) for w in audio_map.get(s, [])
        ) / 1000.0
        duration = audio_dur + buffer_seconds
        encode_cmd += ['-loop', '1', '-t', f'{duration:.3f}', '-i', img]

    encode_cmd += ['-i', merged_wav]

    filter_parts = ''.join(f'[{i}:v]' for i in range(num_slides))
    filter_complex = f'{filter_parts}concat=n={num_slides}:v=1:a=0[v]'

    encode_cmd += [
        '-filter_complex', filter_complex,
        '-map', '[v]',
        '-map', f'{num_slides}:a',
        '-c:v', 'libx264',
        '-preset', 'veryfast',
        '-b:v', '1500k',
        '-r', '30',
        '-c:a', 'aac',
        '-ar', '48000',
        '-pix_fmt', 'yuv420p',
        output_path,
    ]

    r = subprocess.run(encode_cmd, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)
    if r.returncode != 0:
        raise RuntimeError(
            'FFmpeg encode failed: ' +
            r.stderr.decode('utf-8', errors='replace')[-600:]
        )


# ---------------------------------------------------------------------------
# HTML Template — The Wellness Collection branded interface
# ---------------------------------------------------------------------------
HTML_TEMPLATE = r'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Content Studio — The Wellness Collection</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,300;0,400;0,500;0,600;1,300;1,400&family=DM+Sans:wght@300;400;500&display=swap" rel="stylesheet">
<style>
  *, *::before, *::after { margin: 0; padding: 0; box-sizing: border-box; }

  :root {
    --cream: #EAE2D7;
    --taupe: #D1C8BF;
    --blush: #EDE0DC;
    --near-white: #F7F4F2;
    --white: #FFFFFF;
    --dark: #2a2a2a;
    --dark-soft: #444444;
    --gold: #c8922a;
    --gold-light: #e8b04a;
    --text: #1a1a1a;
    --text-muted: #8a7a6a;
    --success: #5a8f6b;
    --success-bg: #eef5f0;
    --error: #a84040;
    --error-bg: #f5eeee;
    --font-display: 'Cormorant Garamond', Georgia, serif;
    --font-body: 'DM Sans', system-ui, -apple-system, sans-serif;
    --radius: 12px;
    --radius-sm: 8px;
    --shadow: 0 4px 24px rgba(0,0,0,0.06);
    --shadow-lg: 0 12px 48px rgba(0,0,0,0.10);
  }

  body {
    font-family: var(--font-body);
    background: var(--cream);
    color: var(--text);
    min-height: 100vh;
    display: flex;
    flex-direction: column;
    -webkit-font-smoothing: antialiased;
  }

  /* --- Header --- */
  header {
    background: var(--taupe);
    padding: 28px 40px;
    text-align: center;
    position: relative;
    overflow: hidden;
  }

  header::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0; bottom: 0;
    background: radial-gradient(ellipse at 50% 0%, rgba(255,255,255,0.15) 0%, transparent 60%);
    pointer-events: none;
  }

  .brand-label {
    font-family: var(--font-body);
    font-size: 10px;
    letter-spacing: 0.3em;
    text-transform: uppercase;
    color: var(--dark-soft);
    margin-bottom: 6px;
    font-weight: 400;
  }

  .brand-title {
    font-family: var(--font-display);
    font-size: 32px;
    font-weight: 300;
    color: var(--dark);
    letter-spacing: 0.02em;
    line-height: 1.1;
  }

  .brand-title em {
    font-style: italic;
    font-weight: 300;
  }

  .brand-sub {
    font-family: var(--font-body);
    font-size: 11px;
    letter-spacing: 0.15em;
    text-transform: uppercase;
    color: var(--text-muted);
    margin-top: 10px;
    font-weight: 400;
  }

  /* --- Divider line --- */
  .divider {
    width: 48px;
    height: 1px;
    background: var(--dark);
    margin: 10px auto 0;
    opacity: 0.2;
  }

  /* --- Main --- */
  main {
    flex: 1;
    max-width: 760px;
    width: 100%;
    margin: 0 auto;
    padding: 36px 24px 60px;
  }

  /* --- Steps --- */
  .step {
    background: var(--white);
    border-radius: var(--radius);
    box-shadow: var(--shadow);
    padding: 28px 32px;
    margin-bottom: 20px;
    border: 1px solid rgba(209,200,191,0.4);
    transition: box-shadow 0.3s ease;
  }

  .step:hover {
    box-shadow: var(--shadow-lg);
  }

  .step-header {
    display: flex;
    align-items: center;
    gap: 14px;
    margin-bottom: 18px;
  }

  .step-number {
    width: 32px;
    height: 32px;
    border-radius: 50%;
    background: var(--blush);
    display: flex;
    align-items: center;
    justify-content: center;
    font-family: var(--font-display);
    font-size: 16px;
    font-weight: 500;
    color: var(--dark);
    flex-shrink: 0;
  }

  .step-title {
    font-family: var(--font-display);
    font-size: 22px;
    font-weight: 400;
    color: var(--dark);
  }

  .step-desc {
    font-size: 13px;
    color: var(--text-muted);
    margin-bottom: 16px;
    line-height: 1.6;
  }

  /* --- File inputs --- */
  .file-row {
    display: flex;
    align-items: center;
    gap: 12px;
    margin-bottom: 12px;
  }

  .file-label {
    font-size: 13px;
    font-weight: 500;
    color: var(--dark);
    min-width: 130px;
    text-align: right;
  }

  .file-input-wrap {
    flex: 1;
    position: relative;
  }

  input[type="file"] {
    width: 100%;
    padding: 10px 14px;
    border: 1.5px dashed var(--taupe);
    border-radius: var(--radius-sm);
    background: var(--near-white);
    font-family: var(--font-body);
    font-size: 13px;
    color: var(--text);
    cursor: pointer;
    transition: border-color 0.2s, background 0.2s;
  }

  input[type="file"]:hover {
    border-color: var(--gold);
    background: var(--white);
  }

  input[type="file"]::file-selector-button {
    background: var(--blush);
    border: 1px solid var(--taupe);
    padding: 6px 16px;
    border-radius: 6px;
    font-family: var(--font-body);
    font-size: 12px;
    font-weight: 500;
    color: var(--dark);
    cursor: pointer;
    margin-right: 12px;
    transition: all 0.2s;
  }

  input[type="file"]::file-selector-button:hover {
    background: var(--taupe);
  }

  /* --- Buffer setting --- */
  .setting-row {
    display: flex;
    align-items: center;
    gap: 12px;
    flex-wrap: wrap;
  }

  .setting-row label {
    font-size: 13px;
    font-weight: 500;
    color: var(--dark);
  }

  .setting-row input[type="number"] {
    width: 80px;
    padding: 8px 12px;
    border: 1.5px solid var(--taupe);
    border-radius: var(--radius-sm);
    background: var(--near-white);
    font-family: var(--font-body);
    font-size: 14px;
    color: var(--text);
    text-align: center;
  }

  .setting-row input[type="number"]:focus {
    outline: none;
    border-color: var(--gold);
    background: var(--white);
  }

  .setting-hint {
    font-size: 12px;
    color: var(--text-muted);
    font-style: italic;
  }

  /* --- Buttons --- */
  .btn {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    gap: 8px;
    padding: 12px 28px;
    border-radius: var(--radius-sm);
    font-family: var(--font-body);
    font-size: 13px;
    font-weight: 500;
    letter-spacing: 0.04em;
    border: none;
    cursor: pointer;
    transition: all 0.25s ease;
    text-decoration: none;
  }

  .btn-secondary {
    background: var(--blush);
    color: var(--dark);
    border: 1px solid var(--taupe);
  }

  .btn-secondary:hover {
    background: var(--taupe);
  }

  .btn-primary {
    background: var(--dark);
    color: var(--white);
    font-size: 14px;
    padding: 14px 36px;
    letter-spacing: 0.06em;
    text-transform: uppercase;
  }

  .btn-primary:hover {
    background: var(--gold);
    box-shadow: 0 4px 16px rgba(200,146,42,0.25);
  }

  .btn-primary:disabled {
    background: var(--taupe);
    color: var(--text-muted);
    cursor: not-allowed;
    box-shadow: none;
  }

  .btn-gold-outline {
    background: transparent;
    color: var(--gold);
    border: 1.5px solid var(--gold);
  }

  .btn-gold-outline:hover {
    background: var(--gold);
    color: var(--white);
  }

  .action-row {
    text-align: center;
    margin-top: 8px;
  }

  /* --- Preview table --- */
  .preview-area {
    background: var(--near-white);
    border: 1px solid rgba(209,200,191,0.5);
    border-radius: var(--radius-sm);
    padding: 20px;
    margin-top: 16px;
    min-height: 60px;
    font-size: 13px;
    line-height: 1.7;
  }

  .preview-area table {
    width: 100%;
    border-collapse: collapse;
  }

  .preview-area th {
    text-align: left;
    font-size: 11px;
    font-weight: 500;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    color: var(--text-muted);
    padding: 6px 8px;
    border-bottom: 1px solid var(--taupe);
  }

  .preview-area td {
    padding: 8px;
    font-size: 13px;
    color: var(--text);
    border-bottom: 1px solid rgba(209,200,191,0.3);
  }

  .preview-area tr:last-child td {
    border-bottom: none;
  }

  .preview-summary {
    margin-top: 14px;
    padding-top: 14px;
    border-top: 1px solid var(--taupe);
    font-size: 13px;
    color: var(--text-muted);
  }

  .preview-summary strong {
    color: var(--dark);
    font-weight: 500;
  }

  .badge-ok {
    display: inline-block;
    background: var(--success-bg);
    color: var(--success);
    font-size: 11px;
    padding: 2px 10px;
    border-radius: 10px;
    font-weight: 500;
  }

  .badge-skip {
    display: inline-block;
    background: var(--blush);
    color: var(--text-muted);
    font-size: 11px;
    padding: 2px 10px;
    border-radius: 10px;
    font-weight: 500;
  }

  /* --- Progress --- */
  .progress-wrap {
    margin: 20px 0 6px;
    display: none;
  }

  .progress-wrap.active {
    display: block;
  }

  .progress-bar-track {
    width: 100%;
    height: 6px;
    background: var(--blush);
    border-radius: 3px;
    overflow: hidden;
  }

  .progress-bar-fill {
    height: 100%;
    background: linear-gradient(90deg, var(--gold), var(--gold-light));
    border-radius: 3px;
    width: 0%;
    transition: width 0.4s ease;
  }

  .progress-status {
    font-size: 12px;
    color: var(--text-muted);
    margin-top: 6px;
    font-style: italic;
  }

  /* --- Success / Error messages --- */
  .message {
    padding: 16px 20px;
    border-radius: var(--radius-sm);
    margin-top: 16px;
    font-size: 13px;
    line-height: 1.6;
    display: none;
  }

  .message.success {
    display: block;
    background: var(--success-bg);
    color: var(--success);
    border: 1px solid rgba(90,143,107,0.2);
  }

  .message.error {
    display: block;
    background: var(--error-bg);
    color: var(--error);
    border: 1px solid rgba(168,64,64,0.2);
  }

  .message h4 {
    font-family: var(--font-display);
    font-size: 18px;
    font-weight: 500;
    margin-bottom: 8px;
  }

  .next-steps {
    margin-top: 12px;
    padding-left: 0;
    list-style: none;
    counter-reset: steps;
  }

  .next-steps li {
    counter-increment: steps;
    padding: 4px 0 4px 28px;
    position: relative;
  }

  .next-steps li::before {
    content: counter(steps);
    position: absolute;
    left: 0;
    width: 20px;
    height: 20px;
    border-radius: 50%;
    background: rgba(90,143,107,0.15);
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 11px;
    font-weight: 500;
    top: 5px;
  }

  /* --- Footer --- */
  footer {
    text-align: center;
    padding: 24px;
    font-size: 11px;
    color: var(--text-muted);
    letter-spacing: 0.1em;
    text-transform: uppercase;
  }

  footer a {
    color: var(--gold);
    text-decoration: none;
  }

  /* --- Animations --- */
  @keyframes fadeUp {
    from { opacity: 0; transform: translateY(12px); }
    to { opacity: 1; transform: translateY(0); }
  }

  .step {
    animation: fadeUp 0.5s ease both;
  }
  .step:nth-child(1) { animation-delay: 0.05s; }
  .step:nth-child(2) { animation-delay: 0.15s; }
  .step:nth-child(3) { animation-delay: 0.25s; }
  .step:nth-child(4) { animation-delay: 0.35s; }
  .step:nth-child(5) { animation-delay: 0.45s; }
  .step:nth-child(6) { animation-delay: 0.55s; }

  /* --- Responsive --- */
  @media (max-width: 640px) {
    main { padding: 20px 16px 40px; }
    .step { padding: 20px; }
    .file-row { flex-direction: column; align-items: stretch; }
    .file-label { text-align: left; min-width: auto; }
    .brand-title { font-size: 26px; }
  }
</style>
</head>
<body>

<header>
  <div class="brand-label">Gracefully Redefined Counseling + Wellness</div>
  <div class="brand-title">The Wellness <em>Collection</em></div>
  <div class="divider"></div>
  <div class="brand-sub">Content Studio</div>
</header>

<main>

  <!-- Step 1: Upload -->
  <div class="step">
    <div class="step-header">
      <div class="step-number">1</div>
      <div class="step-title">Select Your Files</div>
    </div>
    <p class="step-desc">
      Choose your presentation file and the folder of narration audio clips.
      Audio files should be named <strong>slide_01.wav</strong>, <strong>slide_02.wav</strong>, etc.
    </p>
    <div class="file-row">
      <span class="file-label">Presentation</span>
      <div class="file-input-wrap">
        <input type="file" id="pptxFile" accept=".pptx">
      </div>
    </div>
    <div class="file-row">
      <span class="file-label">Audio Clips</span>
      <div class="file-input-wrap">
        <input type="file" id="audioFiles" accept=".wav" multiple>
      </div>
    </div>
  </div>

  <!-- Step 2: Timing Settings -->
  <div class="step">
    <div class="step-header">
      <div class="step-number">2</div>
      <div class="step-title">Timing Settings</div>
    </div>
    <div class="setting-row">
      <label for="buffer">Pause between slides</label>
      <input type="number" id="buffer" value="1.5" min="0" max="10" step="0.5">
      <span class="setting-hint">seconds of silence after each narration ends</span>
    </div>
  </div>

  <!-- Step 3: Audio Leveling -->
  <div class="step">
    <div class="step-header">
      <div class="step-number">3</div>
      <div class="step-title">Audio Leveling</div>
    </div>
    <p class="step-desc">
      Ensure all narration clips play at the same volume. Recommended for recordings
      made across multiple sessions.
    </p>
    <div class="setting-row" style="margin-bottom: 14px;">
      <label for="normMode">Leveling mode</label>
      <select id="normMode" style="padding:8px 12px; border:1.5px solid var(--taupe); border-radius:var(--radius-sm); background:var(--near-white); font-family:var(--font-body); font-size:13px; color:var(--text); min-width:180px;" onchange="updateTargetDefaults()">
        <option value="off">Off — use original audio</option>
        <option value="peak" selected>Peak — match loudest point</option>
        <option value="rms">RMS — match average loudness</option>
      </select>
    </div>
    <div class="setting-row" id="targetRow">
      <label for="normTarget">Target level</label>
      <input type="number" id="normTarget" value="-1.0" min="-40" max="0" step="0.5">
      <span class="setting-hint">dB (peak: -1.0 recommended, RMS: -16.0 recommended)</span>
    </div>
    <div style="margin-top: 12px; padding: 12px 16px; background: var(--near-white); border-radius: var(--radius-sm); font-size: 12px; color: var(--text-muted); line-height: 1.6;">
      <strong style="color: var(--dark);">Peak</strong> — makes every clip's loudest moment hit the same level. Best for consistent voiceover.<br>
      <strong style="color: var(--dark);">RMS</strong> — makes every clip's average volume match. Best when some clips have dynamic range differences.
    </div>
  </div>

  <!-- Step 4: Preview -->
  <div class="step">
    <div class="step-header">
      <div class="step-number">4</div>
      <div class="step-title">Preview &amp; Confirm</div>
    </div>
    <p class="step-desc">
      Review how your audio clips map to each slide before processing.
    </p>
    <div class="action-row" style="text-align:left; margin-bottom: 12px;">
      <button class="btn btn-secondary" onclick="scanPreview()">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
        Scan &amp; Preview
      </button>
    </div>
    <div class="preview-area" id="previewArea">
      <span style="color: var(--text-muted); font-style: italic;">
        Upload your files above, then click Scan &amp; Preview to see the mapping.
      </span>
    </div>
  </div>

  <!-- Step 5: Process -->
  <div class="step">
    <div class="step-header">
      <div class="step-number">5</div>
      <div class="step-title">Create Presentation</div>
    </div>
    <p class="step-desc">
      Embed all narration audio and generate your presentation with automatic slide timing.
    </p>

    <div class="progress-wrap" id="progressWrap">
      <div class="progress-bar-track">
        <div class="progress-bar-fill" id="progressFill"></div>
      </div>
      <div class="progress-status" id="progressStatus">Preparing...</div>
    </div>

    <div id="resultMessage"></div>

    <div class="action-row" style="margin-top: 16px;">
      <button class="btn btn-primary" id="processBtn" onclick="processFiles()">
        Embed Audio &amp; Save
      </button>
    </div>
  </div>

  <!-- Step 6: Export to MP4 -->
  <div class="step">
    <div class="step-header">
      <div class="step-number">6</div>
      <div class="step-title">Export to MP4</div>
    </div>
    <div id="ffmpegWarning" style="display:none; padding:14px 18px; background:#f5eeee; border:1px solid rgba(168,64,64,0.3); border-radius:8px; margin-bottom:16px; font-size:13px; color:#a84040; line-height:1.6;">
      <strong style="font-size:15px;">&#9888; FFmpeg not found on this computer.</strong><br>
      Export to MP4 requires FFmpeg to be installed.<br><br>
      <strong>To install:</strong><br>
      1. Download FFmpeg from <strong>ffmpeg.org/download.html</strong><br>
      2. Extract the zip and copy the <strong>bin</strong> folder contents to <code style="background:rgba(168,64,64,0.1);padding:1px 6px;border-radius:4px;">C:\ffmpeg\bin\</code><br>
      3. Restart this app (close the black window and re-run the .bat file)
    </div>
    <p class="step-desc">
      Convert your presentation directly to an MP4 video using FFmpeg.
      Requires <strong>FFmpeg</strong> in your PATH and either
      <strong>Microsoft PowerPoint</strong> (Windows) or
      <strong>LibreOffice</strong> for slide rendering.
    </p>
    <div class="progress-wrap" id="mp4ProgressWrap">
      <div class="progress-bar-track">
        <div class="progress-bar-fill" id="mp4ProgressFill"></div>
      </div>
      <div class="progress-status" id="mp4ProgressStatus">Preparing&hellip;</div>
    </div>
    <div id="mp4ResultMessage"></div>
    <div class="action-row" style="margin-top: 16px;">
      <button class="btn btn-primary" id="mp4Btn" onclick="exportMp4()">
        Export to MP4
      </button>
    </div>
  </div>

</main>

<footer>
  The Wellness Collection &mdash; Content Studio &nbsp;&middot;&nbsp;
  Powered by <a href="#">AI Visionaries Studio</a>
</footer>

<script>
  // --- Client-side logic ---

  // Check FFmpeg availability on load
  fetch('/api/check_ffmpeg').then(r => r.json()).then(data => {
    if (!data.available) {
      document.getElementById('ffmpegWarning').style.display = 'block';
    }
  }).catch(() => {});

  function updateTargetDefaults() {
    const mode = document.getElementById('normMode').value;
    const targetInput = document.getElementById('normTarget');
    const targetRow = document.getElementById('targetRow');
    if (mode === 'off') {
      targetRow.style.opacity = '0.4';
      targetRow.style.pointerEvents = 'none';
    } else {
      targetRow.style.opacity = '1';
      targetRow.style.pointerEvents = 'auto';
      if (mode === 'peak') targetInput.value = '-1.0';
      if (mode === 'rms') targetInput.value = '-16.0';
    }
  }
  // Initialize on load
  document.addEventListener('DOMContentLoaded', updateTargetDefaults);

  function getNormParams() {
    return {
      mode: document.getElementById('normMode').value,
      target: document.getElementById('normTarget').value,
    };
  }

  async function scanPreview() {
    const pptxInput = document.getElementById('pptxFile');
    const audioInput = document.getElementById('audioFiles');
    const preview = document.getElementById('previewArea');

    if (!pptxInput.files.length) {
      preview.innerHTML = '<span style="color: var(--error);">Please select a PowerPoint file first.</span>';
      return;
    }
    if (!audioInput.files.length) {
      preview.innerHTML = '<span style="color: var(--error);">Please select audio files first.</span>';
      return;
    }

    const formData = new FormData();
    formData.append('pptx', pptxInput.files[0]);
    for (const f of audioInput.files) {
      formData.append('audio', f);
    }
    formData.append('buffer', document.getElementById('buffer').value);
    const norm = getNormParams();
    formData.append('norm_mode', norm.mode);
    formData.append('norm_target', norm.target);

    preview.innerHTML = '<span style="color: var(--text-muted); font-style: italic;">Scanning files...</span>';

    try {
      const resp = await fetch('/api/scan', { method: 'POST', body: formData });
      const data = await resp.json();

      if (data.error) {
        preview.innerHTML = `<span style="color: var(--error);">${data.error}</span>`;
        return;
      }

      const showLevels = data.levels && data.levels.length > 0;
      let html = '<table><tr><th>Slide</th><th>Audio File</th><th>Duration</th>';
      if (showLevels) html += '<th>Peak dB</th><th>RMS dB</th>';
      html += '<th>Advance After</th><th>Status</th></tr>';

      for (const d of data.details) {
        const lvl = showLevels ? data.levels.find(l => l.slide === d.slide) : null;
        html += `<tr>
          <td>${d.slide}</td>
          <td>${d.audio_file}</td>
          <td>${d.duration_s}s</td>`;
        if (showLevels && lvl) {
          html += `<td>${lvl.peak_db} dB</td><td>${lvl.rms_db} dB</td>`;
        } else if (showLevels) {
          html += '<td>—</td><td>—</td>';
        }
        html += `<td>${d.advance_s}s</td>
          <td><span class="badge-ok">Ready</span></td>
        </tr>`;
      }
      html += '</table>';

      if (showLevels) {
        const peaks = data.levels.map(l => l.peak_db);
        const rmses = data.levels.map(l => l.rms_db);
        const peakRange = (Math.max(...peaks) - Math.min(...peaks)).toFixed(1);
        const rmsRange = (Math.max(...rmses) - Math.min(...rmses)).toFixed(1);
        const normLabel = norm.mode === 'off' ? 'leveling off' : `${norm.mode} → ${norm.target} dB`;
        html += `<div style="margin-top:10px; padding:10px 14px; background:var(--blush); border-radius:var(--radius-sm); font-size:12px; color:var(--dark);">
          <strong>Volume spread:</strong> ${peakRange} dB peak range, ${rmsRange} dB RMS range &nbsp;
          ${parseFloat(peakRange) > 3 && norm.mode === 'off'
            ? '<span style="color:var(--error); font-weight:500;">— consider enabling audio leveling</span>'
            : norm.mode !== 'off'
              ? '<span style="color:var(--success); font-weight:500;">— leveling will be applied (' + normLabel + ')</span>'
              : '<span style="color:var(--success);">— volumes are consistent</span>'}
        </div>`;
      }

      if (data.skipped_slides.length > 0) {
        html += '<div style="margin-top:10px;">';
        for (const s of data.skipped_slides) {
          html += `<span class="badge-skip" style="margin:2px 4px;">Slide ${s} — no audio</span> `;
        }
        html += '</div>';
      }

      html += `<div class="preview-summary">
        <strong>${data.total_files}</strong> audio clips &middot;
        <strong>${data.total_duration_s}s</strong> total narration
        (${(data.total_duration_s / 60).toFixed(1)} min) &middot;
        <strong>${data.num_slides}</strong> slides total
      </div>`;

      preview.innerHTML = html;

    } catch (e) {
      preview.innerHTML = `<span style="color: var(--error);">Error: ${e.message}</span>`;
    }
  }

  async function processFiles() {
    const pptxInput = document.getElementById('pptxFile');
    const audioInput = document.getElementById('audioFiles');
    const btn = document.getElementById('processBtn');
    const progressWrap = document.getElementById('progressWrap');
    const progressFill = document.getElementById('progressFill');
    const progressStatus = document.getElementById('progressStatus');
    const resultDiv = document.getElementById('resultMessage');

    if (!pptxInput.files.length || !audioInput.files.length) {
      resultDiv.className = 'message error';
      resultDiv.innerHTML = 'Please select both a PowerPoint file and audio files first.';
      return;
    }

    btn.disabled = true;
    btn.textContent = 'Processing...';
    progressWrap.classList.add('active');
    progressFill.style.width = '10%';
    resultDiv.className = 'message';
    resultDiv.style.display = 'none';

    // Start elapsed timer
    const startTime = Date.now();
    let timerInterval = setInterval(() => {
      const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
      const currentLabel = progressStatus.getAttribute('data-label') || 'Processing';
      progressStatus.textContent = `${currentLabel}  \u2014  ${elapsed}s elapsed`;
    }, 100);

    function setStatus(label) {
      progressStatus.setAttribute('data-label', label);
      const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
      progressStatus.textContent = `${label}  \u2014  ${elapsed}s elapsed`;
    }

    setStatus('Uploading files...');

    const formData = new FormData();
    formData.append('pptx', pptxInput.files[0]);
    for (const f of audioInput.files) {
      formData.append('audio', f);
    }
    formData.append('buffer', document.getElementById('buffer').value);
    const norm = getNormParams();
    formData.append('norm_mode', norm.mode);
    formData.append('norm_target', norm.target);

    try {
      progressFill.style.width = '20%';
      const normLabel = norm.mode !== 'off' ? 'Leveling audio & embedding...' : 'Embedding audio into slides...';
      setStatus(normLabel);

      const resp = await fetch('/api/process', { method: 'POST', body: formData });

      progressFill.style.width = '80%';
      setStatus('Generating output file...');

      if (resp.ok) {
        const blob = await resp.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        const origName = pptxInput.files[0].name.replace('.pptx', '');
        a.href = url;
        a.download = `${origName}_with_audio.pptx`;
        a.click();
        URL.revokeObjectURL(url);

        // Stop timer and show final time
        clearInterval(timerInterval);
        const totalTime = ((Date.now() - startTime) / 1000).toFixed(1);
        progressFill.style.width = '100%';
        progressStatus.textContent = `Complete!  \u2014  ${totalTime}s total`;

        const levelNote = norm.mode !== 'off'
          ? `<p style="margin-top:8px;">Audio leveled to <strong>${norm.target} dB</strong> (${norm.mode} mode).</p>`
          : '';

        resultDiv.className = 'message success';
        resultDiv.innerHTML = `
          <h4>Presentation Ready</h4>
          <p>Your file has been saved with embedded narration and automatic slide timing.</p>
          <p style="margin-top:4px; color: var(--text-muted); font-size:12px;">Processed in ${totalTime} seconds</p>
          ${levelNote}
          <p style="margin-top:10px; font-weight:500;">Next Steps:</p>
          <ol class="next-steps">
            <li>Open the downloaded file in PowerPoint</li>
            <li>Go to <strong>File &rarr; Export &rarr; Create a Video</strong></li>
            <li>Select <strong>Full HD (1080p)</strong> resolution</li>
            <li>Choose <strong>"Use Recorded Timings and Narrations"</strong></li>
            <li>Save as MP4</li>
            <li>Upload to VdoCipher</li>
          </ol>
        `;
      } else {
        clearInterval(timerInterval);
        const errData = await resp.json();
        throw new Error(errData.error || 'Processing failed');
      }

    } catch (e) {
      clearInterval(timerInterval);
      const totalTime = ((Date.now() - startTime) / 1000).toFixed(1);
      progressFill.style.width = '0%';
      resultDiv.className = 'message error';
      resultDiv.innerHTML = `<h4>Something went wrong</h4><p>${e.message}</p><p style="font-size:12px; color:var(--text-muted);">Failed after ${totalTime}s</p>`;
      resultDiv.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }

    btn.disabled = false;
    btn.textContent = 'Embed Audio & Save';
  }
  async function exportMp4() {
    const pptxInput = document.getElementById('pptxFile');
    const audioInput = document.getElementById('audioFiles');
    const btn = document.getElementById('mp4Btn');
    const progressWrap = document.getElementById('mp4ProgressWrap');
    const progressFill = document.getElementById('mp4ProgressFill');
    const progressStatus = document.getElementById('mp4ProgressStatus');
    const resultDiv = document.getElementById('mp4ResultMessage');

    if (!pptxInput.files.length || !audioInput.files.length) {
      resultDiv.className = 'message error';
      resultDiv.innerHTML = 'Please select both a PowerPoint file and audio files first.';
      return;
    }

    btn.disabled = true;
    btn.textContent = 'Exporting…';
    progressWrap.classList.add('active');
    progressFill.style.width = '10%';
    resultDiv.className = 'message';
    resultDiv.style.display = 'none';

    const startTime = Date.now();
    let timerInterval = setInterval(() => {
      const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
      const label = progressStatus.getAttribute('data-label') || 'Processing';
      progressStatus.textContent = `${label}  —  ${elapsed}s elapsed`;
    }, 100);

    function setStatus(label) {
      progressStatus.setAttribute('data-label', label);
      const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
      progressStatus.textContent = `${label}  —  ${elapsed}s elapsed`;
    }

    setStatus('Uploading files…');

    const formData = new FormData();
    formData.append('pptx', pptxInput.files[0]);
    for (const f of audioInput.files) formData.append('audio', f);
    formData.append('buffer', document.getElementById('buffer').value);
    const norm = getNormParams();
    formData.append('norm_mode', norm.mode);
    formData.append('norm_target', norm.target);

    try {
      progressFill.style.width = '20%';
      setStatus('Rendering slides & encoding video…');

      const resp = await fetch('/api/export_mp4', { method: 'POST', body: formData });

      progressFill.style.width = '90%';
      setStatus('Finalizing MP4…');

      if (resp.ok) {
        const blob = await resp.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        const origName = pptxInput.files[0].name.replace(/\.pptx$/i, '');
        a.href = url;
        a.download = `${origName}.mp4`;
        a.click();
        URL.revokeObjectURL(url);

        clearInterval(timerInterval);
        const totalTime = ((Date.now() - startTime) / 1000).toFixed(1);
        progressFill.style.width = '100%';
        progressStatus.textContent = `Complete!  —  ${totalTime}s total`;

        resultDiv.className = 'message success';
        resultDiv.innerHTML = `
          <h4>MP4 Ready</h4>
          <p>Your video has been exported and downloaded.</p>
          <p style="margin-top:4px; color:var(--text-muted); font-size:12px;">
            H.264 &middot; AAC &middot; 30 fps &middot; 8000k &middot; 48 kHz &middot; Exported in ${totalTime}s
          </p>`;
      } else {
        clearInterval(timerInterval);
        const errData = await resp.json().catch(() => ({ error: 'Export failed' }));
        throw new Error(errData.error || 'Export failed');
      }
    } catch (e) {
      clearInterval(timerInterval);
      const totalTime = ((Date.now() - startTime) / 1000).toFixed(1);
      progressFill.style.width = '0%';
      resultDiv.className = 'message error';
      resultDiv.innerHTML = `<h4>Export Failed</h4><p>${e.message}</p>` +
        `<p style="font-size:12px; color:var(--text-muted);">Failed after ${totalTime}s</p>`;
      resultDiv.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }

    btn.disabled = false;
    btn.textContent = 'Export to MP4';
  }

</script>
</body>
</html>'''


# ---------------------------------------------------------------------------
# HTTP Request Handler
# ---------------------------------------------------------------------------
class StudioHandler(SimpleHTTPRequestHandler):
    
    def do_GET(self):
        if self.path == '/' or self.path == '/index.html':
            self.send_response(200)
            self.send_header('Content-type', 'text/html; charset=utf-8')
            self.end_headers()
            self.wfile.write(HTML_TEMPLATE.encode('utf-8'))
        elif self.path == '/api/check_ffmpeg':
            ffmpeg = _find_ffmpeg()
            self._json_response({'available': ffmpeg is not None, 'path': ffmpeg})
        else:
            self.send_error(404)
    
    def do_POST(self):
        if self.path == '/api/scan':
            self._handle_scan()
        elif self.path == '/api/process':
            self._handle_process()
        elif self.path == '/api/export_mp4':
            self._handle_export_mp4()
        else:
            self.send_error(404)
    
    def _parse_multipart(self):
        """Parse multipart form data with binary-safe handling (no cgi module)."""
        content_type = self.headers['Content-Type']
        content_length = int(self.headers['Content-Length'])
        
        # Clear previous uploads
        for f in os.listdir(UPLOAD_DIR):
            fp = os.path.join(UPLOAD_DIR, f)
            if os.path.isfile(fp):
                os.remove(fp)
            elif os.path.isdir(fp):
                shutil.rmtree(fp)
        
        # Read the entire body as binary
        body = self.rfile.read(content_length)
        
        # Extract boundary from Content-Type
        boundary = None
        for part in content_type.split(';'):
            part = part.strip()
            if part.startswith('boundary='):
                boundary = part[len('boundary='):].strip().strip('"')
                break
        
        if not boundary:
            raise ValueError("No boundary found in Content-Type")
        
        boundary_bytes = ('--' + boundary).encode('utf-8')
        end_boundary = (boundary_bytes + b'--')
        
        # Split body by boundary
        parts = body.split(boundary_bytes)
        
        pptx_path = None
        audio_dir = os.path.join(UPLOAD_DIR, 'audio')
        os.makedirs(audio_dir, exist_ok=True)
        buffer = 1.5
        norm_mode = 'off'
        norm_target = -1.0
        
        for part in parts:
            # Skip empty parts and the final boundary marker
            if not part or part.strip() == b'' or part.strip() == b'--':
                continue
            if part.startswith(b'--'):
                continue
            
            # Remove leading \r\n
            if part.startswith(b'\r\n'):
                part = part[2:]
            # Remove trailing \r\n
            if part.endswith(b'\r\n'):
                part = part[:-2]
            
            # Split headers from body at \r\n\r\n
            header_end = part.find(b'\r\n\r\n')
            if header_end == -1:
                continue
            
            header_block = part[:header_end].decode('utf-8', errors='replace')
            file_data = part[header_end + 4:]
            
            # Parse the Content-Disposition header
            name = None
            filename = None
            for line in header_block.split('\r\n'):
                if 'Content-Disposition' in line:
                    for token in line.split(';'):
                        token = token.strip()
                        if token.startswith('name='):
                            name = token[len('name='):].strip('"')
                        elif token.startswith('filename='):
                            filename = token[len('filename='):].strip('"')
            
            if not name:
                continue
            
            if name == 'pptx' and filename:
                # Sanitize filename
                safe_name = re.sub(r'[^\w\s\-\.]', '_', filename)
                pptx_path = os.path.join(UPLOAD_DIR, safe_name)
                with open(pptx_path, 'wb') as f:
                    f.write(file_data)
            elif name == 'audio' and filename:
                safe_name = re.sub(r'[^\w\s\-\.]', '_', filename)
                audio_path = os.path.join(audio_dir, safe_name)
                with open(audio_path, 'wb') as f:
                    f.write(file_data)
            elif name == 'buffer':
                try:
                    buffer = float(file_data.decode('utf-8').strip())
                except:
                    buffer = 1.5
            elif name == 'norm_mode':
                val = file_data.decode('utf-8').strip()
                norm_mode = val if val in ('off', 'peak', 'rms') else 'off'
            elif name == 'norm_target':
                try:
                    norm_target = float(file_data.decode('utf-8').strip())
                except:
                    norm_target = -1.0
        
        return pptx_path, audio_dir, buffer, norm_mode, norm_target
    
    def _handle_scan(self):
        """Scan files and return preview data with audio levels."""
        try:
            pptx_path, audio_dir, buffer, norm_mode, norm_target = self._parse_multipart()
            
            if not pptx_path:
                self._json_response({'error': 'No PowerPoint file uploaded'})
                return
            
            num_slides = count_slides(pptx_path)
            audio_map = find_audio_files(audio_dir)
            
            details = []
            levels = []
            total_clips = 0
            for slide_num, wav_paths in audio_map.items():
                slide_total_ms = 0
                clip_names = []
                for wav_path in wav_paths:
                    dur_ms = get_wav_duration_ms(wav_path)
                    slide_total_ms += dur_ms
                    clip_names.append(os.path.basename(wav_path))
                    total_clips += 1
                    # Analyze audio levels per clip
                    stats = analyze_wav(wav_path)
                    levels.append({
                        'slide': slide_num,
                        'audio_file': os.path.basename(wav_path),
                        'peak_db': stats['peak_db'],
                        'rms_db': stats['rms_db'],
                    })
                
                slide_dur_s = slide_total_ms / 1000.0
                details.append({
                    'slide': slide_num,
                    'audio_file': ', '.join(clip_names),
                    'clips': len(wav_paths),
                    'duration_s': round(slide_dur_s, 1),
                    'advance_s': round(slide_dur_s + buffer, 1),
                })
            
            covered = set(audio_map.keys())
            skipped = [i for i in range(1, num_slides + 1) if i not in covered]
            total_dur = sum(d['duration_s'] for d in details)
            
            self._json_response({
                'num_slides': num_slides,
                'total_files': total_clips,
                'total_duration_s': round(total_dur, 1),
                'skipped_slides': skipped,
                'details': details,
                'levels': levels,
            })
        except Exception as e:
            self._json_response({'error': str(e)})
    
    def _handle_process(self):
        """Process files and return the output PPTX."""
        try:
            pptx_path, audio_dir, buffer, norm_mode, norm_target = self._parse_multipart()
            
            if not pptx_path:
                self._json_response({'error': 'No PowerPoint file uploaded'}, status=400)
                return
            
            output_path = os.path.join(UPLOAD_DIR, 'output_with_audio.pptx')
            
            results = embed_audio_into_pptx(
                pptx_path=pptx_path,
                audio_dir=audio_dir,
                output_path=output_path,
                buffer_seconds=buffer,
                normalize_mode=norm_mode,
                normalize_target_db=norm_target,
            )
            
            # Send file back
            file_size = os.path.getsize(output_path)
            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation')
            self.send_header('Content-Disposition', 'attachment; filename="presentation_with_audio.pptx"')
            self.send_header('Content-Length', str(file_size))
            self.end_headers()
            with open(output_path, 'rb') as f:
                shutil.copyfileobj(f, self.wfile, 65536)
        
        except Exception as e:
            self._json_response({'error': str(e)}, status=500)
    
    def _handle_export_mp4(self):
        """Render PPTX + audio to MP4 and stream the file back."""
        try:
            pptx_path, audio_dir, buffer, norm_mode, norm_target = self._parse_multipart()

            if not pptx_path:
                self._json_response({'error': 'No PowerPoint file uploaded'}, status=400)
                return

            output_path = os.path.join(UPLOAD_DIR, 'output_export.mp4')

            _export_mp4(
                pptx_path=pptx_path,
                audio_dir=audio_dir,
                output_path=output_path,
                buffer_seconds=buffer,
                norm_mode=norm_mode,
                norm_target=norm_target,
            )

            file_size = os.path.getsize(output_path)
            self.send_response(200)
            self.send_header('Content-Type', 'video/mp4')
            self.send_header('Content-Disposition', 'attachment; filename="presentation_export.mp4"')
            self.send_header('Content-Length', str(file_size))
            self.end_headers()
            with open(output_path, 'rb') as f:
                shutil.copyfileobj(f, self.wfile, 65536)

        except Exception as e:
            self._json_response({'error': str(e)}, status=500)


    def _json_response(self, data, status=200):
        body = json.dumps(data).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)
    
    def log_message(self, format, *args):
        """Suppress default logging noise."""
        pass


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    is_cloud = 'PORT' in os.environ
    host = '0.0.0.0' if is_cloud else '127.0.0.1'
    server = ThreadingHTTPServer((host, PORT), StudioHandler)
    url = f'http://localhost:{PORT}'

    print()
    print("  ┌─────────────────────────────────────────────┐")
    print("  │                                             │")
    print("  │   The Wellness Collection                   │")
    print("  │   Content Studio                            │")
    print("  │                                             │")
    print(f"  │   Running at: {url:<28} │")
    print("  │                                             │")
    print("  │   Press Ctrl+C to stop                      │")
    print("  │                                             │")
    print("  └─────────────────────────────────────────────┘")
    print()

    if not is_cloud:
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n  Shutting down Content Studio...")
        server.shutdown()


if __name__ == '__main__':
    main()
