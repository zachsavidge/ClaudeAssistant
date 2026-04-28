"""
Transcribe a Japanese MP4 video using OpenAI Whisper API, then translate to English.
Handles files > 25 MB by extracting audio and splitting into chunks.
"""
import _cloud_creds  # noqa: F401  -- bootstraps GCP creds for cloud sandbox (used by translation step)

import os, sys, tempfile, subprocess, math, json
from pathlib import Path

import imageio_ffmpeg
from openai import OpenAI

# --- Config ---
try:
    API_KEY = os.environ["OPENAI_API_KEY"]
except KeyError:
    sys.stderr.write(
        "ERROR: OPENAI_API_KEY env var not set. This script requires an OpenAI API key "
        "for Whisper transcription. In cloud sessions, set it in the workspace env vars; "
        "locally, `set OPENAI_API_KEY=...` (Windows) or `export OPENAI_API_KEY=...` (Unix).\n"
    )
    sys.exit(2)
FFMPEG = imageio_ffmpeg.get_ffmpeg_exe()
MAX_CHUNK_MB = 24  # Whisper API limit is 25 MB; leave margin

client = OpenAI(api_key=API_KEY)

def extract_audio(mp4_path: str, out_path: str) -> float:
    """Extract audio as mp3, return duration in seconds."""
    subprocess.run(
        [FFMPEG, "-y", "-i", mp4_path, "-vn", "-acodec", "libmp3lame", "-q:a", "4", out_path],
        check=True, capture_output=True,
    )
    # Get duration
    result = subprocess.run(
        [FFMPEG, "-i", out_path],
        capture_output=True, text=True,
    )
    for line in result.stderr.split("\n"):
        if "Duration:" in line:
            t = line.split("Duration:")[1].split(",")[0].strip()
            parts = t.split(":")
            return float(parts[0]) * 3600 + float(parts[1]) * 60 + float(parts[2])
    return 0

def split_audio(audio_path: str, chunk_dir: str, max_mb: int = MAX_CHUNK_MB) -> list:
    """Split audio into chunks under max_mb. Returns list of chunk paths."""
    file_size = os.path.getsize(audio_path)
    if file_size <= max_mb * 1024 * 1024:
        return [audio_path]

    # Get duration
    result = subprocess.run(
        [FFMPEG, "-i", audio_path], capture_output=True, text=True,
    )
    duration = 0
    for line in result.stderr.split("\n"):
        if "Duration:" in line:
            t = line.split("Duration:")[1].split(",")[0].strip()
            parts = t.split(":")
            duration = float(parts[0]) * 3600 + float(parts[1]) * 60 + float(parts[2])
            break

    num_chunks = math.ceil(file_size / (max_mb * 1024 * 1024))
    chunk_duration = duration / num_chunks
    chunks = []

    for i in range(num_chunks):
        start = i * chunk_duration
        chunk_path = os.path.join(chunk_dir, f"chunk_{i:03d}.mp3")
        subprocess.run(
            [FFMPEG, "-y", "-i", audio_path, "-ss", str(start), "-t", str(chunk_duration),
             "-acodec", "libmp3lame", "-q:a", "4", chunk_path],
            check=True, capture_output=True,
        )
        chunks.append(chunk_path)
        print(f"  Created chunk {i+1}/{num_chunks}: {os.path.getsize(chunk_path)/1024/1024:.1f} MB")

    return chunks

def transcribe_chunks(chunks: list) -> str:
    """Send each chunk to Whisper API, return combined Japanese transcript."""
    full_text = []
    for i, chunk_path in enumerate(chunks):
        print(f"  Transcribing chunk {i+1}/{len(chunks)}...")
        with open(chunk_path, "rb") as f:
            result = client.audio.transcriptions.create(
                model="whisper-1",
                file=f,
                language="ja",
                response_format="verbose_json",
            )
        # Extract text with timestamps
        if hasattr(result, 'segments') and result.segments:
            for seg in result.segments:
                start = seg.start if hasattr(seg, 'start') else seg['start']
                text = seg.text if hasattr(seg, 'text') else seg['text']
                mins = int(start // 60)
                secs = int(start % 60)
                full_text.append(f"[{mins:02d}:{secs:02d}] {text}")
        else:
            full_text.append(result.text)

    return "\n".join(full_text)

def translate_text(japanese_text: str) -> str:
    """Translate Japanese transcript to English using GPT-4o-mini."""
    # Split into manageable sections (~3000 chars each)
    lines = japanese_text.split("\n")
    sections = []
    current = []
    current_len = 0
    for line in lines:
        current.append(line)
        current_len += len(line)
        if current_len > 3000:
            sections.append("\n".join(current))
            current = []
            current_len = 0
    if current:
        sections.append("\n".join(current))

    translated = []
    for i, section in enumerate(sections):
        print(f"  Translating section {i+1}/{len(sections)}...")
        resp = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are a professional Japanese-to-English translator. "
                 "Translate the following transcript accurately, preserving timestamps and speaker meaning. "
                 "This is a business due diligence interview. Keep business/financial terms precise. "
                 "Preserve the [MM:SS] timestamps exactly as they appear."},
                {"role": "user", "content": section},
            ],
            temperature=0.2,
        )
        translated.append(resp.choices[0].message.content)

    return "\n".join(translated)


def main():
    mp4_path = sys.argv[1]
    output_dir = sys.argv[2]
    os.makedirs(output_dir, exist_ok=True)

    base_name = Path(mp4_path).stem
    print(f"Processing: {Path(mp4_path).name} ({os.path.getsize(mp4_path)/1024/1024:.1f} MB)")

    with tempfile.TemporaryDirectory() as tmpdir:
        # Step 1: Extract audio
        print("\n[1/4] Extracting audio...")
        audio_path = os.path.join(tmpdir, "audio.mp3")
        extract_audio(mp4_path, audio_path)
        audio_mb = os.path.getsize(audio_path) / 1024 / 1024
        print(f"  Audio extracted: {audio_mb:.1f} MB")

        # Step 2: Split if needed
        print("\n[2/4] Splitting audio if needed...")
        chunks = split_audio(audio_path, tmpdir)
        print(f"  {len(chunks)} chunk(s) ready")

        # Step 3: Transcribe
        print("\n[3/4] Transcribing (Japanese)...")
        japanese_text = transcribe_chunks(chunks)

        # Save Japanese transcript
        ja_path = os.path.join(output_dir, f"{base_name}_transcript_JA.txt")
        with open(ja_path, "w", encoding="utf-8") as f:
            f.write(japanese_text)
        print(f"  Japanese transcript saved: {ja_path}")

        # Step 4: Translate
        print("\n[4/4] Translating to English...")
        english_text = translate_text(japanese_text)

        # Save English transcript
        en_path = os.path.join(output_dir, f"{base_name}_transcript_EN.txt")
        with open(en_path, "w", encoding="utf-8") as f:
            f.write(english_text)
        print(f"  English transcript saved: {en_path}")

    print("\nDone!")

if __name__ == "__main__":
    main()
