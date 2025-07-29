# Runtime hook for Playwright
# This ensures Playwright browsers work correctly in the bundled executable

import os
import sys
from pathlib import Path


def setup_playwright_browsers():
    """Set up Playwright browsers for the bundled executable."""
    if getattr(sys, "frozen", False):
        # Running in a bundle
        bundle_dir = Path(sys._MEIPASS)

        # Look for Playwright browsers in the bundle
        playwright_dirs = ["playwright-win32", "playwright-darwin", "playwright-linux"]

        for dir_name in playwright_dirs:
            browser_dir = bundle_dir / dir_name
            if browser_dir.exists():
                # Set environment variable to point to bundled browsers
                os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(browser_dir)
                break


# Set up browsers when module is imported
setup_playwright_browsers()
