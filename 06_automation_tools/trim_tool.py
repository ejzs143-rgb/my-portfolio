import os
import re
import json
import shutil
import subprocess
import tempfile
import tkinter as tk
from typing import Optional

from tkinter import filedialog, simpledialog, messagebox, ttk


# -----------------------------
# Time parser
# -----------------------------
def parse_time(s: str) -> float:
    """
    Accept:
    - seconds: "240" or "240.5"
    - mm:ss: "04:00"
    - hh:mm:ss: "01:17:00"
    Return seconds (float).
    """
    s = (s or "").strip()
    if not s:
        raise ValueError("empty input")

    if re.fullmatch(r"\d+(\.\d+)?", s):
        return float(s)

    parts = s.split(":")
    if len(parts) == 2:
        mm, ss = parts
        return float(mm) * 60 + float(ss)

    if len(parts) == 3:
        hh, mm, ss = parts
        return float(hh) * 3600 + float(mm) * 60 + float(ss)

    raise ValueError(f"Invalid time format: {s}")


# -----------------------------
# ffmpeg / ffprobe finder
# -----------------------------
def find_ffmpeg() -> Optional[str]:
    p = shutil.which("ffmpeg")
    if p:
        return p

    candidates = [
        os.path.join(os.environ.get("PROGRAMFILES", ""), "ffmpeg", "bin", "ffmpeg.exe"),
        os.path.join(os.environ.get("PROGRAMFILES(X86)", ""), "ffmpeg", "bin", "ffmpeg.exe"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs", "ffmpeg", "bin", "ffmpeg.exe"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "WindowsApps", "ffmpeg.exe"),
    ]
    for c in candidates:
        if c and os.path.isfile(c):
            return c
    return None


def find_ffprobe(ffmpeg_path: Optional[str]) -> Optional[str]:
    p = shutil.which("ffprobe")
    if p:
        return p

    if ffmpeg_path:
        candidate = os.path.join(os.path.dirname(ffmpeg_path), "ffprobe.exe")
        if os.path.isfile(candidate):
            return candidate

    return None


def ask_ffmpeg_path(initial_dir: Optional[str] = None) -> Optional[str]:
    path = filedialog.askopenfilename(
        title="Select ffmpeg.exe",
        initialdir=initial_dir if initial_dir and os.path.isdir(initial_dir) else None,
        filetypes=[("ffmpeg executable", "ffmpeg.exe"), ("Executable", "*.exe"), ("All files", "*.*")],
    )
    return path or None


# -----------------------------
# Media probe
# -----------------------------
def probe_media(ffprobe_path: Optional[str], in_file: str) -> dict:
    info = {
        "has_video": False,
        "has_audio": False,
        "has_attached_pic": False,
    }

    if not ffprobe_path or not os.path.isfile(ffprobe_path):
        return info

    cmd = [
        ffprobe_path,
        "-v", "error",
        "-show_streams",
        "-of", "json",
        in_file,
    ]

    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, check=True)
        data = json.loads(proc.stdout or "{}")
        for stream in data.get("streams", []):
            codec_type = (stream.get("codec_type") or "").lower()
            disposition = stream.get("disposition") or {}
            attached_pic = bool(disposition.get("attached_pic", 0))

            if codec_type == "audio":
                info["has_audio"] = True
            elif codec_type == "video":
                if attached_pic:
                    info["has_attached_pic"] = True
                else:
                    info["has_video"] = True
    except Exception:
        pass

    return info


# -----------------------------
# UI: progress dialog
# -----------------------------
class ProgressDialog:
    def __init__(self, parent, title="Processing"):
        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.top.geometry("560x170")
        self.top.resizable(False, False)
        self.top.attributes("-topmost", True)

        self.label = ttk.Label(self.top, text="Starting...", anchor="w")
        self.label.pack(fill="x", padx=12, pady=(12, 6))

        self.pbar = ttk.Progressbar(self.top, mode="determinate", maximum=100)
        self.pbar.pack(fill="x", padx=12, pady=6)

        self.detail = ttk.Label(self.top, text="", anchor="w")
        self.detail.pack(fill="x", padx=12, pady=(6, 12))

        self.top.update_idletasks()

    def update(self, text=None, percent=None, detail=None):
        if text is not None:
            self.label.config(text=text)
        if percent is not None:
            self.pbar.config(mode="determinate")
            self.pbar["value"] = max(0, min(100, float(percent)))
        if detail is not None:
            self.detail.config(text=detail)
        self.top.update_idletasks()

    def close(self):
        try:
            self.top.destroy()
        except Exception:
            pass


# -----------------------------
# ffmpeg helpers
# -----------------------------
def run_ffmpeg_with_progress(ffmpeg_path, cmd, segment_seconds, progress_cb):
    """
    Run ffmpeg with -progress pipe:1 and parse out_time_ms.
    progress_cb(percent, detail)
    """
    full_cmd = [ffmpeg_path, "-hide_banner", "-y"] + cmd + ["-progress", "pipe:1", "-nostats"]

    proc = subprocess.Popen(
        full_cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        bufsize=1
    )

    last_percent = -1
    try:
        if proc.stdout:
            for line in proc.stdout:
                line = line.strip()

                if line.startswith("out_time_ms="):
                    ms = int(line.split("=", 1)[1])
                    sec = ms / 1_000_000.0
                    pct = (sec / segment_seconds) * 100.0 if segment_seconds > 0 else 0.0

                    if int(pct) != int(last_percent):
                        last_percent = pct
                        progress_cb(pct, f"{sec:,.1f}s / {segment_seconds:,.1f}s")

                elif line.startswith("progress=") and line.endswith("end"):
                    break

    finally:
        stderr = proc.stderr.read() if proc.stderr else ""
        rc = proc.wait()

    if rc != 0:
        raise RuntimeError(stderr.strip() or "ffmpeg failed")


def safe_remove(path: str):
    try:
        if path and os.path.isfile(path):
            os.remove(path)
    except Exception:
        pass


def ffmpeg_trim_audio_best_effort(ffmpeg_path, in_file, start_sec, end_sec, out_file, progress_cb, aac_bitrate="128k"):
    """
    ExampleBrand Bo-only output (.m4a)
    1) Try stream copy
    2) Fallback to AAC re-encode
    """
    segment_seconds = max(0.01, end_sec - start_sec)

    try:
        progress_cb(0, "Trying audio stream copy...")
        cmd_copy = [
            "-ss", str(start_sec),
            "-to", str(end_sec),
            "-i", in_file,
            "-vn",
            "-c:a", "copy",
            out_file,
        ]
        run_ffmpeg_with_progress(ffmpeg_path, cmd_copy, segment_seconds, progress_cb)
        return "audio-copy"

    except Exception:
        safe_remove(out_file)

    progress_cb(0, f"Copy failed. Re-encoding AAC {aac_bitrate}...")
    cmd_enc = [
        "-i", in_file,
        "-ss", str(start_sec),
        "-to", str(end_sec),
        "-vn",
        "-c:a", "aac",
        "-b:a", aac_bitrate,
        out_file,
    ]
    run_ffmpeg_with_progress(ffmpeg_path, cmd_enc, segment_seconds, progress_cb)
    return "audio-reencode"


def ffmpeg_trim_media_best_effort(ffmpeg_path, in_file, start_sec, end_sec, out_file, progress_cb, aac_bitrate="128k"):
    """
    Media output (.mp4)
    1) Try stream copy while preserving video/audio/subtitles/cover-art where possible
    2) Fallback to H.264 + AAC re-encode
    """
    segment_seconds = max(0.01, end_sec - start_sec)

    try:
        progress_cb(0, "Trying media stream copy...")
        cmd_copy = [
            "-ss", str(start_sec),
            "-to", str(end_sec),
            "-i", in_file,
            "-map", "0:v?",
            "-map", "0:a?",
            "-map", "0:s?",
            "-map_metadata", "0",
            "-c", "copy",
            "-movflags", "+faststart",
            out_file,
        ]
        run_ffmpeg_with_progress(ffmpeg_path, cmd_copy, segment_seconds, progress_cb)
        return "media-copy"

    except Exception:
        safe_remove(out_file)

    progress_cb(0, f"Copy failed. Re-encoding video/audio (AAC {aac_bitrate})...")
    cmd_enc = [
        "-i", in_file,
        "-ss", str(start_sec),
        "-to", str(end_sec),
        "-map", "0:v?",
        "-map", "0:a?",
        "-map", "0:s?",
        "-map_metadata", "0",
        "-c:v", "libx264",
        "-preset", "medium",
        "-crf", "20",
        "-c:a", "aac",
        "-b:a", aac_bitrate,
        "-c:s", "copy",
        "-movflags", "+faststart",
        out_file,
    ]
    run_ffmpeg_with_progress(ffmpeg_path, cmd_enc, segment_seconds, progress_cb)
    return "media-reencode"


# -----------------------------
# PowerPoint fallback (audio-oriented)
# -----------------------------
def write_ps1_and_run(in_file, start_sec, end_sec, out_hint):
    """
    Fallback: PowerPoint automation.
    This is mainly useful for audio-oriented trimming when ffmpeg is unavailable.
    """
    ps1 = r"""
param(
    [Parameter(Mandatory=$true)][string]$InFile,
    [Parameter(Mandatory=$true)][double]$StartSec,
    [Parameter(Mandatory=$true)][double]$EndSec,
    [Parameter(Mandatory=$true)][string]$OutHint
)

Add-Type -AssemblyName System.Windows.Forms | Out-Null

try {
    $ppt = New-Object -ComObject PowerPoint.Application
} catch {
    [System.Windows.Forms.MessageBox]::Show("PowerPoint could not be started (COM). Is desktop PowerPoint installed?","Trim Tool")
    exit 2
}

$ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$pres = $ppt.Presentations.Add()
$slide = $pres.Slides.Add(1, 12)

$shape = $slide.Shapes.AddMediaObject2($InFile, $false, $true, 50, 50, 400, 50)

$mf = $shape.MediaFormat
$mf.StartPoint = [Math]::Round($StartSec * 1000)
$mf.EndPoint   = [Math]::Round($EndSec   * 1000)

$shape.Select()

$msg = "Trim is already set in PowerPoint." + [Environment]::NewLine + [Environment]::NewLine +
"To export the trimmed media manually:" + [Environment]::NewLine +
"1) File > Info > Compress Media" + [Environment]::NewLine +
"2) Right-click media object > Save Media As" + [Environment]::NewLine + [Environment]::NewLine +
"Suggested save path/name:" + [Environment]::NewLine + $OutHint

[System.Windows.Forms.MessageBox]::Show($msg, "Trim Tool")
"""

    with tempfile.TemporaryDirectory() as td:
        ps1_path = os.path.join(td, "ppt_trim.ps1")

        with open(ps1_path, "w", encoding="utf-8") as f:
            f.write(ps1)

        cmd = [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-File", ps1_path,
            "-InFile", in_file,
            "-StartSec", str(start_sec),
            "-EndSec", str(end_sec),
            "-OutHint", out_hint,
        ]
        subprocess.run(cmd)


# -----------------------------
# Main
# -----------------------------
def main():
    root = tk.Tk()
    root.withdraw()

    in_file = filedialog.askopenfilename(
        title="Select media file",
        filetypes=[
            ("Media files", "*.m4a;*.mp4;*.m4v;*.mov;*.mp3;*.aac;*.wav;*.mkv;*.avi"),
            ("All files", "*.*"),
        ],
    )
    if not in_file:
        messagebox.showinfo("Cancelled", "Input file selection was cancelled.")
        return

    s = simpledialog.askstring("Trim", "Start time (sec or mm:ss or hh:mm:ss)", initialvalue="240")
    if s is None:
        messagebox.showinfo("Cancelled", "Start time input was cancelled.")
        return

    e = simpledialog.askstring("Trim", "End time (sec or mm:ss or hh:mm:ss)", initialvalue="4620")
    if e is None:
        messagebox.showinfo("Cancelled", "End time input was cancelled.")
        return

    try:
        start_sec = parse_time(s)
        end_sec = parse_time(e)
        if end_sec <= start_sec:
            raise ValueError("End must be greater than Start")
    except Exception as ex:
        messagebox.showerror("FAILED", f"Invalid time input: {ex}")
        return

    ffmpeg_path = find_ffmpeg()

    if not ffmpeg_path:
        messagebox.showwarning(
            "ffmpeg not found",
            "ffmpeg was not found on this PC.\n\n"
            "You can choose ffmpeg.exe manually now.\n"
            "If you cancel, PowerPoint fallback will be used."
        )
        ffmpeg_path = ask_ffmpeg_path()

    ffprobe_path = find_ffprobe(ffmpeg_path) if ffmpeg_path else None
    info = probe_media(ffprobe_path, in_file)

    suggested_ext = ".mp4" if (info["has_video"] or info["has_attached_pic"]) else ".m4a"
    save_title = "Save trimmed media as"
    default_name = os.path.splitext(os.path.basename(in_file))[0] + "_cut" + suggested_ext

    out_file = filedialog.asksaveasfilename(
        title=save_title,
        defaultextension=suggested_ext,
        initialfile=default_name,
        filetypes=[
            ("MP4 media", "*.mp4"),
            ("M4A audio", "*.m4a"),
            ("All files", "*.*"),
        ],
    )
    if not out_file:
        messagebox.showinfo("Cancelled", "Output selection was cancelled.")
        return

    out_ext = os.path.splitext(out_file)[1].lower()

    if ffmpeg_path:
        if not (os.path.isfile(ffmpeg_path) and ffmpeg_path.lower().endswith("ffmpeg.exe")):
            messagebox.showerror("FAILED", f"Selected file is not ffmpeg.exe:\n{ffmpeg_path}")
            return

        br = simpledialog.askstring(
            "ExampleBrand Bo Bitrate",
            "AAC bitrate for re-encode fallback (e.g. 96k / 128k / 192k)",
            initialvalue="128k"
        )
        if br is None or not re.fullmatch(r"\d+k", br.strip()):
            br = "128k"
        br = br.strip()

        dlg = ProgressDialog(root, "Trimming with ffmpeg")

        try:
            dlg.update(text="Processing...", percent=0, detail="Preparing...")

            def cb(pct, detail):
                dlg.update(text="Processing with ffmpeg...", percent=pct, detail=detail)

            if out_ext == ".mp4":
                mode = ffmpeg_trim_media_best_effort(
                    ffmpeg_path, in_file, start_sec, end_sec, out_file, cb, aac_bitrate=br
                )
            else:
                mode = ffmpeg_trim_audio_best_effort(
                    ffmpeg_path, in_file, start_sec, end_sec, out_file, cb, aac_bitrate=br
                )

            dlg.close()

            if not os.path.isfile(out_file) or os.path.getsize(out_file) == 0:
                raise RuntimeError("Output file was not created or is 0 bytes.")

            size_mb = os.path.getsize(out_file) / (1024 * 1024)
            messagebox.showinfo(
                "SUCCESS",
                f"Completed ({mode}).\nSaved:\n{out_file}\nSize: {size_mb:.2f} MB"
            )
            return

        except Exception as ex:
            dlg.close()
            try:
                if os.path.isfile(out_file) and os.path.getsize(out_file) == 0:
                    os.remove(out_file)
            except Exception:
                pass

            messagebox.showerror("FAILED", f"Trimming failed.\n\n{ex}")
            return

    if out_ext == ".mp4":
        messagebox.showerror(
            "FAILED",
            "ffmpeg is required to preserve video/cover art into MP4.\n\n"
            "Please install ffmpeg or select ffmpeg.exe manually."
        )
        return

    messagebox.showinfo(
        "Fallback: PowerPoint",
        "ffmpeg was not provided.\n\n"
        "PowerPoint will open with trim already set.\n\n"
        "IMPORTANT:\n"
        "File > Info > Compress Media (wait)\n"
        "then Right-click media object > Save Media As"
    )
    write_ps1_and_run(in_file, start_sec, end_sec, out_file)

    messagebox.showinfo(
        "HANDOFF",
        "PowerPoint has been opened and trim is set.\n"
        "Please complete:\n"
        "Compress Media -> Save Media As\n\n"
        f"Suggested output:\n{out_file}"
    )


if __name__ == "__main__":
    main()