#!/usr/bin/env python3
"""png_optimize.py
Optimizes PNG images using Pillow to reduce file size without quality loss.
Usage: python scripts/png_optimize.py [--fix]
"""

import argparse
import os
import subprocess
import sys
from datetime import datetime
from PIL import Image

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)
log_dir = os.path.join(project_root, "Logs")
os.makedirs(log_dir, exist_ok=True)
log_path = os.path.join(log_dir, "png_optimize.log")


def log(msg):
    line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {msg}"
    print(line)
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def optimize_png(filepath, fix):
    """Optimize a single PNG file. Returns True if optimized."""
    try:
        with Image.open(filepath) as img:
            original_size = os.path.getsize(filepath)
            if fix:
                img.save(filepath, optimize=True, quality=95)
                new_size = os.path.getsize(filepath)
                saved = original_size - new_size
                if saved > 0:
                    log(f"Optimized {os.path.basename(filepath)}: {original_size} -> {new_size} bytes (-{saved} bytes)")
                return saved > 0
            else:
                # Estimate savings by saving to temp
                import tempfile
                with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
                    img.save(tmp.name, optimize=True, quality=95)
                    new_size = os.path.getsize(tmp.name)
                    os.unlink(tmp.name)
                    saved = original_size - new_size
                    if saved > 0:
                        log(f"Would optimize {os.path.basename(filepath)}: {original_size} -> {new_size} bytes (-{saved} bytes)")
                        return saved > 0
    except Exception as e:
        log(f"Error optimizing {os.path.basename(filepath)}: {e}")
    return False


def main():
    parser = argparse.ArgumentParser(description="Optimize PNG images with Pillow")
    parser.add_argument("--fix", action="store_true", help="Apply optimization automatically")
    args = parser.parse_args()

    log("=== PNG optimization (Pillow) started ===")

    result = subprocess.run(
        ["python", "-c", "from PIL import Image; print('Pillow version:', Image.__version__)"],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        log("ERROR: Pillow is not installed. Run: pip install Pillow")
        sys.exit(1)
    log(f"Pillow version: {result.stdout.strip()}")

    screenshot_dir = os.path.join(project_root, "documentation", "screenshots")
    if not os.path.exists(screenshot_dir):
        log("ERROR: Screenshot directory not found")
        sys.exit(1)

    png_files = [f for f in os.listdir(screenshot_dir) if f.lower().endswith(".png")]
    if not png_files:
        log("No PNG files found in documentation/screenshots/")
        return

    total_saved = 0
    optimized_count = 0
    for png_file in sorted(png_files):
        filepath = os.path.join(screenshot_dir, png_file)
        if optimize_png(filepath, args.fix):
            optimized_count += 1
            total_saved += 1  # count as optimized

    if optimized_count > 0:
        log(f"PNG optimization complete: {optimized_count} image(s) optimized, saved ~{total_saved} bytes total")
    else:
        log("PNG optimization complete: all images already optimized")


if __name__ == "__main__":
    main()