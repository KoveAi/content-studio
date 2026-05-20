#!/usr/bin/env python3
"""
PPTX Audio Embedder
===================
Embeds WAV audio files into PowerPoint slides with auto-play and auto-advance timing.

Usage (CLI):
    python embed_audio.py presentation.pptx audio_folder/ [--buffer 1.5] [--output output.pptx]

Usage (GUI):
    python embed_audio_gui.py

Audio files should be named to match slide numbers:
    slide_01.wav, slide_02.wav, etc.
    OR: slide_1.wav, slide_2.wav, etc.
    OR: 01.wav, 02.wav, etc.
    OR: 1.wav, 2.wav, etc.

The tool reads each WAV duration, embeds it into the matching slide,
configures auto-play on slide load, and sets slide advance timing
to (audio duration + buffer).
"""

import os
import re
import sys
import wave
import math
import array
import struct
import shutil
import zipfile
import tempfile
import argparse
import xml.etree.ElementTree as ET
from pathlib import Path
from copy import deepcopy


# ---------------------------------------------------------------------------
# Audio Normalization (stdlib only — no pip installs needed)
# Uses struct.unpack/pack for bulk operations instead of per-sample loops
# ---------------------------------------------------------------------------
def analyze_wav(wav_path: str) -> dict:
    """
    Analyze a WAV file and return its audio statistics.
    Uses struct for fast bulk unpacking.
    """
    with wave.open(wav_path, 'r') as w:
        params = w.getparams()
        raw = w.readframes(params.nframes)

    if params.sampwidth == 2:
        samples = array.array('h', raw)
        if sys.byteorder != 'little':
            samples.byteswap()
        max_possible = 32767
    elif params.sampwidth == 1:
        samples = tuple(b - 128 for b in raw)
        max_possible = 127
    elif params.sampwidth == 3:
        samples = []
        for i in range(0, len(raw), 3):
            val = int.from_bytes(raw[i:i+3], byteorder='little', signed=True)
            samples.append(val)
        max_possible = 8388607
    else:
        raise ValueError(f"Unsupported sample width: {params.sampwidth} bytes")

    if len(samples) == 0:
        return {'peak_db': -96.0, 'rms_db': -96.0, 'peak_linear': 0,
                'rms_linear': 0, 'sample_rate': params.framerate,
                'channels': params.nchannels, 'sample_width': params.sampwidth,
                'n_frames': params.nframes}

    # Use min/max on the tuple directly — much faster than abs() loop
    peak_pos = max(samples)
    peak_neg = min(samples)
    peak_linear = max(peak_pos, -peak_neg)

    # RMS via sum of squares
    sum_sq = sum(s * s for s in samples)
    rms_linear = math.sqrt(sum_sq / len(samples))

    peak_db = 20 * math.log10(peak_linear / max_possible) if peak_linear > 0 else -96.0
    rms_db = 20 * math.log10(rms_linear / max_possible) if rms_linear > 0 else -96.0

    return {
        'peak_db': round(peak_db, 1),
        'rms_db': round(rms_db, 1),
        'peak_linear': peak_linear,
        'rms_linear': rms_linear,
        'max_possible': max_possible,
        'sample_rate': params.framerate,
        'channels': params.nchannels,
        'sample_width': params.sampwidth,
        'n_frames': params.nframes,
    }


def normalize_wav(
    input_path: str,
    output_path: str,
    mode: str = 'peak',
    target_db: float = -1.0,
) -> dict:
    """
    Normalize a WAV file using fast bulk struct operations.
    """
    stats_before = analyze_wav(input_path)

    with wave.open(input_path, 'r') as w:
        params = w.getparams()
        raw = w.readframes(params.nframes)

    sw = params.sampwidth
    if sw == 2:
        samples = array.array('h', raw)
        if sys.byteorder != 'little':
            samples.byteswap()
        max_possible = 32767
    elif sw == 1:
        samples = tuple(b - 128 for b in raw)
        max_possible = 127
    else:
        raise ValueError(f"Unsupported sample width: {sw} bytes")

    if len(samples) == 0 or stats_before['peak_linear'] == 0:
        shutil.copy2(input_path, output_path)
        return {'gain_db': 0, 'before': stats_before, 'after': stats_before, 'clipped': False}

    # Calculate gain
    target_linear = max_possible * (10 ** (target_db / 20))

    if mode == 'peak':
        current = stats_before['peak_linear']
    elif mode == 'rms':
        current = stats_before['rms_linear']
    else:
        raise ValueError(f"Unknown normalization mode: {mode}")

    if current <= 0:
        shutil.copy2(input_path, output_path)
        return {'gain_db': 0, 'before': stats_before, 'after': stats_before, 'clipped': False}

    gain = target_linear / current
    gain_db = 20 * math.log10(gain) if gain > 0 else 0

    # Apply gain with clipping — use fast integer math + bulk pack
    clipped = False
    if sw == 2:
        min_val, max_val = -32768, 32767
        # Single list comprehension with inline clamping — much faster than per-append loop
        normalized = []
        for s in samples:
            val = int(s * gain)
            if val > max_val:
                val = max_val
                clipped = True
            elif val < min_val:
                val = min_val
                clipped = True
            normalized.append(val)
        out_arr = array.array('h', normalized)
        if sys.byteorder != 'little':
            out_arr.byteswap()
        out_bytes = out_arr.tobytes()
    else:
        min_val, max_val = -128, 127
        normalized = []
        for s in samples:
            val = int(s * gain)
            if val > max_val:
                val = max_val
                clipped = True
            elif val < min_val:
                val = min_val
                clipped = True
            normalized.append(val)
        out_bytes = bytes([max(0, min(255, s + 128)) for s in normalized])

    # Write output
    with wave.open(output_path, 'w') as w:
        w.setparams(params)
        w.writeframes(out_bytes)

    stats_after = analyze_wav(output_path)

    return {
        'gain_db': round(gain_db, 1),
        'before': stats_before,
        'after': stats_after,
        'clipped': clipped,
    }


def normalize_batch(
    audio_map: dict,
    output_dir: str,
    mode: str = 'peak',
    target_db: float = -1.0,
    progress_callback=None,
) -> dict:
    """
    Normalize all WAV files in a batch to a consistent level.

    Args:
        audio_map: {slide_num: [wav_path, ...]} from find_audio_files()
        output_dir: Directory to write normalized files
        mode: 'peak' or 'rms'
        target_db: Target level in dB
        progress_callback: Optional callable(message, percent)

    Returns:
        dict with per-file results and new audio_map pointing to normalized files
    """
    os.makedirs(output_dir, exist_ok=True)
    results = {'files': [], 'normalized_map': {}}

    # Count total clips for progress
    all_clips = [(slide_num, wp) for slide_num, paths in audio_map.items() for wp in paths]
    total = len(all_clips)
    
    for idx, (slide_num, wav_path) in enumerate(all_clips):
        if progress_callback:
            pct = idx / max(total, 1)
            progress_callback(f"Leveling slide {slide_num} audio...", pct)

        filename = os.path.basename(wav_path)
        out_path = os.path.join(output_dir, filename)

        result = normalize_wav(wav_path, out_path, mode=mode, target_db=target_db)
        result['slide'] = slide_num
        result['filename'] = filename
        results['files'].append(result)
        
        if slide_num not in results['normalized_map']:
            results['normalized_map'][slide_num] = []
        results['normalized_map'][slide_num].append(out_path)

    return results


# ---------------------------------------------------------------------------
# XML Namespace Map
# ---------------------------------------------------------------------------
NSMAP = {
    'a':   'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r':   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'p':   'http://schemas.openxmlformats.org/presentationml/2006/main',
    'p14': 'http://schemas.microsoft.com/office/powerpoint/2010/main',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
    'ct':  'http://schemas.openxmlformats.org/package/2006/content-types',
}

# Register all namespaces so ET doesn't mangle them
for prefix, uri in NSMAP.items():
    ET.register_namespace(prefix if prefix not in ('rel', 'ct') else '', uri)

# Also register common namespaces that appear in PPTX
ET.register_namespace('', 'http://schemas.openxmlformats.org/package/2006/relationships')


REL_TYPE_AUDIO = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio'
REL_TYPE_MEDIA = 'http://schemas.microsoft.com/office/2007/relationships/media'


def _create_transparent_png(path: str):
    """Create a minimal 1x1 transparent PNG file (stdlib only, no Pillow needed)."""
    import zlib
    # Minimal valid PNG: 1x1 pixel, RGBA, fully transparent
    def _chunk(chunk_type, data):
        c = chunk_type + data
        crc = zlib.crc32(c) & 0xFFFFFFFF
        return struct.pack('>I', len(data)) + c + struct.pack('>I', crc)

    signature = b'\x89PNG\r\n\x1a\n'
    ihdr = struct.pack('>IIBBBBB', 1, 1, 8, 6, 0, 0, 0)  # 1x1, 8bit, RGBA
    raw_data = b'\x00' + b'\x00\x00\x00\x00'  # filter byte + 1 transparent RGBA pixel
    idat = zlib.compress(raw_data)

    with open(path, 'wb') as f:
        f.write(signature)
        f.write(_chunk(b'IHDR', ihdr))
        f.write(_chunk(b'IDAT', idat))
        f.write(_chunk(b'IEND', b''))


# ---------------------------------------------------------------------------
# WAV Duration Reader
# ---------------------------------------------------------------------------
def get_wav_duration_ms(wav_path: str) -> int:
    """Return duration of a WAV file in milliseconds."""
    with wave.open(wav_path, 'r') as w:
        frames = w.getnframes()
        rate = w.getframerate()
        duration_s = frames / float(rate)
    return int(duration_s * 1000)


# ---------------------------------------------------------------------------
# Audio File Discovery
# ---------------------------------------------------------------------------
def find_audio_files(audio_dir: str) -> dict:
    """
    Scan directory for WAV files and map them to slide numbers.
    Supports single clips: slide_01.wav, slide_1.wav, 01.wav, 1.wav
    Supports multi clips: slide_01a.wav, slide_01b.wav (played sequentially)
    
    Returns: {slide_number: [filepath, ...]} (1-indexed, list per slide, sorted alphabetically)
    """
    audio_map = {}
    patterns = [
        # slide_01a.wav, slide_01b.wav, slide_01.wav
        re.compile(r'slide[_\-]?(\d+)([a-z])?\.wav', re.IGNORECASE),
        # 01a.wav, 01.wav, 1.wav
        re.compile(r'^(\d+)([a-z])?\.wav$', re.IGNORECASE),
    ]
    
    for filename in os.listdir(audio_dir):
        if not filename.lower().endswith('.wav'):
            continue
        for pattern in patterns:
            match = pattern.match(filename)
            if match:
                slide_num = int(match.group(1))
                filepath = os.path.join(audio_dir, filename)
                if slide_num not in audio_map:
                    audio_map[slide_num] = []
                audio_map[slide_num].append(filepath)
                break
    
    # Sort clips within each slide alphabetically (a, b, c order)
    for slide_num in audio_map:
        audio_map[slide_num].sort(key=lambda p: os.path.basename(p).lower())
    
    return dict(sorted(audio_map.items()))


# ---------------------------------------------------------------------------
# Count Slides in PPTX
# ---------------------------------------------------------------------------
def count_slides(pptx_path: str) -> int:
    """Count slides in a PPTX file."""
    with zipfile.ZipFile(pptx_path, 'r') as z:
        slide_files = [n for n in z.namelist() if re.match(r'ppt/slides/slide\d+\.xml$', n)]
    return len(slide_files)


# ---------------------------------------------------------------------------
# XML Builders
# ---------------------------------------------------------------------------
def build_audio_shape_xml(rid_audio: str, rid_media: str, rid_image: str, audio_name: str, shape_id: int) -> str:
    """
    Build the <p:pic> XML element for an embedded audio shape.
    No xmlns declarations — inherits from parent <p:sld> element.
    """
    return f'''<p:pic><p:nvPicPr><p:cNvPr id="{shape_id}" name="{audio_name}"><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr><a:audioFile r:link="{rid_audio}"/><p:extLst><p:ext uri="{{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}}"><p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="{rid_media}"/></p:ext></p:extLst></p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="{rid_image}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="609600" cy="609600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'''


def build_timing_xml(shape_ids: list, durations_ms: list = None) -> str:
    """
    Build the <p:timing> XML for auto-play audio.
    
    For a single clip: plays immediately on slide load.
    For multiple clips: plays sequentially — clip 2 starts after clip 1's duration, etc.
    
    shape_ids: list of shape IDs (one per audio clip)
    durations_ms: list of clip durations in ms (needed for sequential timing of 2+ clips)
    """
    if len(shape_ids) == 1:
        # Single clip — simple auto-play
        spid = shape_ids[0]
        return f'''<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst><p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:cmd type="call" cmd="playFrom(0)"><p:cBhvr><p:cTn id="5" dur="1" fill="hold"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:cmd></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'''
    
    # Multiple clips — sequential playback
    # Each clip gets its own <p:par> block inside the main sequence.
    # First clip starts at delay=0, subsequent clips start after previous clips finish.
    ctn_id = 1
    
    clip_pars = []
    cumulative_delay = 0
    for idx, spid in enumerate(shape_ids):
        ctn_id += 1; outer_id = ctn_id
        ctn_id += 1; inner_id = ctn_id
        ctn_id += 1; cmd_id = ctn_id
        
        clip_par = (
            f'<p:par><p:cTn id="{outer_id}" fill="hold">'
            f'<p:stCondLst><p:cond delay="{cumulative_delay}"/></p:stCondLst>'
            f'<p:childTnLst><p:par><p:cTn id="{inner_id}" fill="hold">'
            f'<p:stCondLst><p:cond delay="0"/></p:stCondLst>'
            f'<p:childTnLst><p:cmd type="call" cmd="playFrom(0)">'
            f'<p:cBhvr><p:cTn id="{cmd_id}" dur="1" fill="hold"/>'
            f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
            f'</p:cBhvr></p:cmd></p:childTnLst>'
            f'</p:cTn></p:par></p:childTnLst>'
            f'</p:cTn></p:par>'
        )
        clip_pars.append(clip_par)
        
        # Add this clip's duration to cumulative delay for next clip
        if durations_ms and idx < len(durations_ms):
            cumulative_delay += durations_ms[idx]
    
    clips_xml = ''.join(clip_pars)
    ctn_id += 1; seq_id = ctn_id
    
    return (
        f'<p:timing><p:tnLst><p:par>'
        f'<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
        f'<p:seq concurrent="1" nextAc="seek">'
        f'<p:cTn id="{seq_id}" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
        f'{clips_xml}'
        f'</p:childTnLst></p:cTn>'
        f'<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
        f'<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
        f'</p:seq></p:childTnLst></p:cTn>'
        f'</p:par></p:tnLst></p:timing>'
    )


def build_transition_xml(advance_ms: int) -> str:
    """
    Build <p:transition> XML that auto-advances the slide after the given time.
    No xmlns declarations — inherits from parent <p:sld> element.
    """
    return f'''<p:transition advClick="1" advTm="{advance_ms}"/>'''


# ---------------------------------------------------------------------------
# Core Embedding Logic
# ---------------------------------------------------------------------------
def embed_audio_into_pptx(
    pptx_path: str,
    audio_dir: str,
    output_path: str,
    buffer_seconds: float = 1.5,
    hide_icon: bool = True,
    normalize_mode: str = 'off',
    normalize_target_db: float = -1.0,
    progress_callback=None
) -> dict:
    """
    Main function: embed WAV files into PPTX slides.
    
    Args:
        pptx_path: Path to input .pptx
        audio_dir: Directory containing WAV files
        output_path: Path for output .pptx
        buffer_seconds: Extra seconds after audio ends before slide advances
        hide_icon: If True, position audio icon off-slide (hidden)
        normalize_mode: 'off', 'peak', or 'rms'
        normalize_target_db: Target level in dB (default: -1.0 for peak, -16.0 for RMS)
        progress_callback: Optional callable(message: str, percent: float)
    
    Returns:
        dict with summary info: {slides_processed, slides_skipped, total_duration_s, details: [...], normalization: {...}}
    """
    
    def log(msg, pct=None):
        if progress_callback:
            progress_callback(msg, pct)
        else:
            print(msg)
    
    # Validate inputs
    if not os.path.isfile(pptx_path):
        raise FileNotFoundError(f"PPTX not found: {pptx_path}")
    if not os.path.isdir(audio_dir):
        raise NotADirectoryError(f"Audio directory not found: {audio_dir}")
    
    # Discover audio files
    audio_map = find_audio_files(audio_dir)
    if not audio_map:
        raise ValueError(f"No WAV files found in {audio_dir}. "
                        f"Expected names like: slide_01.wav, slide_1.wav, 01.wav, 1.wav")
    
    num_slides = count_slides(pptx_path)
    log(f"Found {len(audio_map)} audio files for {num_slides} slides")
    
    # Validate slide numbers
    for slide_num in audio_map:
        if slide_num < 1 or slide_num > num_slides:
            raise ValueError(f"Audio file maps to slide {slide_num}, "
                           f"but presentation only has {num_slides} slides")
    
    # Work in a temp directory
    with tempfile.TemporaryDirectory() as tmpdir:

        # --- Audio Normalization (if enabled) ---
        norm_results = None
        if normalize_mode in ('peak', 'rms'):
            log("Leveling audio volumes...", 0.03)
            norm_dir = os.path.join(tmpdir, 'normalized')
            norm_results = normalize_batch(
                audio_map, norm_dir,
                mode=normalize_mode,
                target_db=normalize_target_db,
                progress_callback=lambda msg, pct: log(msg, 0.03 + pct * 0.07),
            )
            # Use normalized files from here on
            audio_map = norm_results['normalized_map']
            log(f"Audio leveled: {len(norm_results['files'])} files normalized to {normalize_target_db} dB ({normalize_mode})", 0.10)

        extract_dir = os.path.join(tmpdir, 'extracted')
        
        # Extract PPTX
        log("Extracting presentation...", 0.12)
        with zipfile.ZipFile(pptx_path, 'r') as z:
            z.extractall(extract_dir)
        
        # Ensure media directory exists
        media_dir = os.path.join(extract_dir, 'ppt', 'media')
        os.makedirs(media_dir, exist_ok=True)
        
        # Process Content_Types to add WAV if not present
        ct_path = os.path.join(extract_dir, '[Content_Types].xml')
        ct_tree = ET.parse(ct_path)
        ct_root = ct_tree.getroot()
        ct_ns = 'http://schemas.openxmlformats.org/package/2006/content-types'
        
        # Check if .wav extension is already registered
        has_wav = False
        has_png = False
        for default in ct_root.findall(f'{{{ct_ns}}}Default'):
            ext = default.get('Extension', '').lower()
            if ext == 'wav':
                has_wav = True
            if ext == 'png':
                has_png = True
        if not has_wav:
            wav_default = ET.SubElement(ct_root, f'{{{ct_ns}}}Default')
            wav_default.set('Extension', 'wav')
            wav_default.set('ContentType', 'audio/wav')
        if not has_png:
            png_default = ET.SubElement(ct_root, f'{{{ct_ns}}}Default')
            png_default.set('Extension', 'png')
            png_default.set('ContentType', 'image/png')
        
        ct_tree.write(ct_path, xml_declaration=True, encoding='UTF-8')
        
        # Process each slide
        results = {
            'slides_processed': 0,
            'slides_skipped': 0,
            'total_duration_s': 0.0,
            'details': []
        }
        
        total_steps = len(audio_map)
        for step_idx, (slide_num, wav_paths) in enumerate(audio_map.items()):
            pct = 0.1 + (0.8 * step_idx / max(total_steps, 1))
            clip_count = len(wav_paths)
            log(f"Processing slide {slide_num} ({clip_count} clip{'s' if clip_count > 1 else ''})...", pct)
            
            # --- Update slide .rels ---
            slide_rels_path = os.path.join(
                extract_dir, 'ppt', 'slides', '_rels', f'slide{slide_num}.xml.rels'
            )
            
            rels_ns = 'http://schemas.openxmlformats.org/package/2006/relationships'
            ET.register_namespace('', rels_ns)
            
            rels_tree = ET.parse(slide_rels_path)
            rels_root = rels_tree.getroot()
            
            # Find highest existing rId
            existing_rids = []
            for rel in rels_root.findall(f'{{{rels_ns}}}Relationship'):
                rid = rel.get('Id', '')
                rid_match = re.match(r'rId(\d+)', rid)
                if rid_match:
                    existing_rids.append(int(rid_match.group(1)))
            
            next_rid = max(existing_rids, default=0) + 1
            
            # Read slide XML and find max shape ID
            slide_xml_path = os.path.join(
                extract_dir, 'ppt', 'slides', f'slide{slide_num}.xml'
            )
            with open(slide_xml_path, 'r', encoding='utf-8') as f:
                slide_content = f.read()
            
            slide_tree = ET.parse(slide_xml_path)
            slide_root = slide_tree.getroot()
            max_id = 0
            for elem in slide_root.iter():
                cid = elem.get('id')
                if cid and cid.isdigit():
                    max_id = max(max_id, int(cid))
            
            # Process each audio clip for this slide
            shape_ids = []
            durations_ms = []
            total_duration_ms = 0
            clip_details = []
            
            for clip_idx, wav_path in enumerate(wav_paths):
                clip_suffix = chr(ord('a') + clip_idx) if clip_count > 1 else ''
                
                # Read duration
                duration_ms = get_wav_duration_ms(wav_path)
                durations_ms.append(duration_ms)
                total_duration_ms += duration_ms
                
                # Copy WAV into media folder
                audio_filename = f'audio_slide{slide_num}{clip_suffix}.wav'
                dest_audio = os.path.join(media_dir, audio_filename)
                shutil.copy2(wav_path, dest_audio)
                
                # Create thumbnail PNG
                icon_filename = f'audio_icon{slide_num}{clip_suffix}.png'
                icon_path = os.path.join(media_dir, icon_filename)
                if not os.path.exists(icon_path):
                    _create_transparent_png(icon_path)
                
                # Assign relationship IDs (3 per clip: audio, media, image)
                rid_audio = f'rId{next_rid}'
                rid_media = f'rId{next_rid + 1}'
                rid_image = f'rId{next_rid + 2}'
                next_rid += 3
                
                # Add relationships
                audio_rel = ET.SubElement(rels_root, f'{{{rels_ns}}}Relationship')
                audio_rel.set('Id', rid_audio)
                audio_rel.set('Type', REL_TYPE_AUDIO)
                audio_rel.set('Target', f'../media/{audio_filename}')
                
                media_rel = ET.SubElement(rels_root, f'{{{rels_ns}}}Relationship')
                media_rel.set('Id', rid_media)
                media_rel.set('Type', REL_TYPE_MEDIA)
                media_rel.set('Target', f'../media/{audio_filename}')
                
                image_rel = ET.SubElement(rels_root, f'{{{rels_ns}}}Relationship')
                image_rel.set('Id', rid_image)
                image_rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')
                image_rel.set('Target', f'../media/{icon_filename}')
                
                # Build audio shape and insert into spTree
                shape_id = max_id + 1 + clip_idx
                shape_ids.append(shape_id)
                
                audio_shape = build_audio_shape_xml(rid_audio, rid_media, rid_image, audio_filename, shape_id)
                slide_content = slide_content.replace('</p:spTree>', audio_shape + '\n</p:spTree>')
                
                clip_details.append(os.path.basename(wav_path))
            
            # Write rels
            rels_tree.write(slide_rels_path, xml_declaration=True, encoding='UTF-8')
            
            # Build timing XML (sequential playback for multi-clip)
            timing_xml = build_timing_xml(shape_ids, durations_ms)
            
            # Build transition XML (advance after ALL clips finish + buffer)
            advance_ms = total_duration_ms + int(buffer_seconds * 1000)
            transition_xml = build_transition_xml(advance_ms)
            
            # Remove existing timing/transition, then add ours
            slide_content = re.sub(r'<p:timing[^>]*>.*?</p:timing>', '', slide_content, flags=re.DOTALL)
            slide_content = re.sub(r'<p:transition[^/]*?/>', '', slide_content, flags=re.DOTALL)
            slide_content = re.sub(r'<p:transition[^>]*>.*?</p:transition>', '', slide_content, flags=re.DOTALL)
            
            insert_block = transition_xml + '\n' + timing_xml + '\n'
            slide_content = slide_content.replace('</p:sld>', insert_block + '</p:sld>')
            
            # Write modified slide
            with open(slide_xml_path, 'w', encoding='utf-8') as f:
                f.write(slide_content)
            
            total_duration_s = total_duration_ms / 1000.0
            results['slides_processed'] += 1
            results['total_duration_s'] += total_duration_s
            results['details'].append({
                'slide': slide_num,
                'audio_file': ', '.join(clip_details),
                'clips': clip_count,
                'duration_s': round(total_duration_s, 1),
                'advance_after_s': round(advance_ms / 1000.0, 1),
            })
        
        results['slides_skipped'] = num_slides - results['slides_processed']
        results['normalization'] = norm_results
        
        # Repack PPTX
        log("Repacking presentation...", 0.92)
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for root, dirs, files in os.walk(extract_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, extract_dir)
                    zout.write(file_path, arcname)
        
        log(f"Done! Output saved to: {output_path}", 1.0)
    
    return results


# ---------------------------------------------------------------------------
# CLI Entry Point
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(
        description='Embed WAV audio into PowerPoint slides with auto-play and timing.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Audio File Naming:
  slide_01.wav, slide_02.wav   (recommended)
  slide_1.wav, slide_2.wav     (also works)
  01.wav, 02.wav               (also works)
  1.wav, 2.wav                 (also works)

Example:
  python embed_audio.py presentation.pptx ./audio_clips/ --buffer 2.0
        """
    )
    parser.add_argument('pptx', help='Input PowerPoint file')
    parser.add_argument('audio_dir', help='Directory containing WAV files')
    parser.add_argument('--buffer', type=float, default=1.5,
                       help='Seconds of buffer after audio ends before slide advances (default: 1.5)')
    parser.add_argument('--output', '-o', default=None,
                       help='Output file path (default: input_with_audio.pptx)')
    parser.add_argument('--normalize', choices=['off', 'peak', 'rms'], default='off',
                       help='Audio normalization mode (default: off)')
    parser.add_argument('--target-db', type=float, default=None,
                       help='Normalization target in dB (default: -1.0 for peak, -16.0 for rms)')
    
    args = parser.parse_args()
    
    if args.output is None:
        base = os.path.splitext(args.pptx)[0]
        args.output = f"{base}_with_audio.pptx"
    
    # Set default target based on mode
    if args.target_db is None:
        if args.normalize == 'rms':
            args.target_db = -16.0
        else:
            args.target_db = -1.0
    
    try:
        results = embed_audio_into_pptx(
            pptx_path=args.pptx,
            audio_dir=args.audio_dir,
            output_path=args.output,
            buffer_seconds=args.buffer,
            normalize_mode=args.normalize,
            normalize_target_db=args.target_db,
        )
        
        print("\n" + "=" * 50)
        print("SUMMARY")
        print("=" * 50)
        print(f"Slides with audio: {results['slides_processed']}")
        print(f"Slides skipped:    {results['slides_skipped']}")
        print(f"Total audio:       {results['total_duration_s']:.1f}s")
        if results.get('normalization'):
            print(f"Audio leveling:    {args.normalize} → {args.target_db} dB")
        print()
        for d in results['details']:
            print(f"  Slide {d['slide']:>2}: {d['audio_file']:<25} "
                  f"({d['duration_s']}s audio, advances after {d['advance_after_s']}s)")
        print(f"\nOutput: {args.output}")
        
    except Exception as e:
        print(f"\nERROR: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()
