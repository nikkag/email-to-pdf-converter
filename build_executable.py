#!/usr/bin/env python3
"""
Build script for EmailToPDFConverter executable with Playwright support.

This script ensures that Playwright browsers are properly included in the PyInstaller build.
"""

import os
import platform
import subprocess
import sys
from pathlib import Path


def find_playwright_browsers():
    """Find the Playwright browsers installation directory."""
    if platform.system() == "Windows":
        browsers_path = Path(os.environ.get("USERPROFILE", "")) / "AppData" / "Local" / "ms-playwright"
    else:
        browsers_path = Path.home() / ".cache" / "ms-playwright"

    return browsers_path


def install_playwright_browsers():
    """Install Playwright browsers if not already installed."""
    print("🔧 Installing Playwright browsers...")
    try:
        subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=True, capture_output=True)
        print("✅ Playwright browsers installed successfully")
    except subprocess.CalledProcessError as e:
        print(f"❌ Failed to install Playwright browsers: {e}")
        return False
    return True


def build_executable():
    """Build the executable using PyInstaller."""
    print("🔧 Building executable...")

    # Check if spec file exists
    spec_file = Path("EmailToPDFConverter.spec")
    if spec_file.exists():
        print("📋 Using existing spec file")
        cmd = [sys.executable, "-m", "PyInstaller", "EmailToPDFConverter.spec"]
    else:
        print("📋 Creating new build with Playwright support")
        cmd = [
            sys.executable,
            "-m",
            "PyInstaller",
            "--onefile",
            "--windowed",
            "--name",
            "EmailToPDFConverter",
            "--additional-hooks-dir",
            "hooks",
            "--runtime-hook",
            "hooks/runtime-hook-playwright.py",
            "eml_to_pdf_converter.py",
        ]

    try:
        subprocess.run(cmd, check=True)
        print("✅ Executable built successfully")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Failed to build executable: {e}")
        return False


def main():
    """Main build process."""
    print("🚀 Starting EmailToPDFConverter build process...")

    # Install Playwright browsers
    if not install_playwright_browsers():
        print("⚠️ Continuing without Playwright browsers (HTML rendering may not work)")

    # Build executable
    if build_executable():
        print("🎉 Build completed successfully!")
        print("📁 Executable location: dist/EmailToPDFConverter.exe")

        # Check file size
        exe_path = Path("dist/EmailToPDFConverter.exe")
        if exe_path.exists():
            size_mb = exe_path.stat().st_size / (1024 * 1024)
            print(f"📊 Executable size: {size_mb:.1f} MB")

            if size_mb < 50:
                print("⚠️ Warning: Executable seems small. Playwright browsers may not be included.")
            else:
                print("✅ Executable size looks good (includes Playwright browsers)")
    else:
        print("❌ Build failed")
        sys.exit(1)


if __name__ == "__main__":
    main()
