#!/usr/bin/env python3
"""
EML to PDF Converter

This module provides functionality to convert .eml files to PDFs with date-based naming
following the same logic as the amdate.py script. It includes a file selector dialog
for cross-platform directory selection.

Author: Dima Ghalili
Email: dghalili@fdmattorneys.com
"""

import asyncio
import html
import os
import re
import sys
import tkinter as tk
import unicodedata
from datetime import datetime
from email import policy
from email.parser import Parser
from email.utils import parsedate_to_datetime
from pathlib import Path

from bs4 import BeautifulSoup
from fpdf import FPDF
from playwright.async_api import async_playwright


class DirectorySelector:
    """Cross-platform directory selector using tkinter."""

    def __init__(self) -> None:
        """Initialize the directory selector."""
        self.root = tk.Tk()
        self.root.withdraw()  # Hide the main window
        self.root.attributes("-topmost", True)  # Keep dialog on top

    def select_directory(self) -> Path | None:
        """
        Open a directory selection dialog.

        Returns:
            Optional[Path]: Selected directory path or None if cancelled
        """
        try:
            from tkinter import filedialog

            directory = filedialog.askdirectory(
                title="Select directory containing .eml files", mustexist=True, initialdir=os.getcwd()
            )
            return Path(directory) if directory else None
        except Exception as e:
            print(f"❌ Error opening directory selector: {e}")
            return None
        finally:
            self.root.destroy()


class EMLToPDFConverter:
    """Convert EML files to PDFs with date-based naming."""

    def __init__(self) -> None:
        """Initialize the converter."""
        self.converted_files: list[str] = []
        self.failed_files: list[str] = []
        self.browser = None
        self.playwright = None

    async def _initialize_browser(self) -> None:
        """Initialize the browser instance for HTML rendering."""
        try:
            self.playwright = await async_playwright().start()
            self.browser = await self.playwright.chromium.launch()
        except Exception as e:
            print(f"⚠️ Failed to initialize browser: {e}")
            self.browser = None

    async def _cleanup_browser(self) -> None:
        """Clean up browser resources."""
        if self.browser:
            try:
                await self.browser.close()
            except Exception:
                pass
        if self.playwright:
            try:
                await self.playwright.stop()
            except Exception:
                pass

    def _extract_filename_prefix(self, filename: str) -> str:
        """
        Extract the first 3 words from filename for use in new filename.

        Args:
            filename: Original filename without extension

        Returns:
            str: Sanitized prefix from first 3 words
        """
        words = filename.split()
        first_three = "_".join(words[:3]) if words else "NoName"
        return re.sub(r"[^\w_]", "", first_three)

    def _parse_email_date(self, msg) -> datetime | None:
        """
        Parse date from email headers.

        Args:
            msg: Parsed email message object

        Returns:
            Optional[datetime]: Parsed datetime or None if not found
        """
        date_header = msg.get("Date")
        if not date_header:
            return None

        try:
            return parsedate_to_datetime(date_header)
        except Exception:
            return None

    def _extract_email_content(self, msg) -> tuple[str, str]:
        """
        Extract and render email content similar to email client display.

        Args:
            msg: Parsed email message object

        Returns:
            tuple[str, str]: (Plain text content, HTML content)
        """
        # Extract email headers
        subject = msg.get("Subject", "No Subject")
        sender = msg.get("From", "Unknown Sender")
        recipient = msg.get("To", "Unknown Recipient")
        date_header = msg.get("Date", "No Date")

        # Format headers
        header_content = f"Subject: {subject}\n"
        header_content += f"From: {sender}\n"
        header_content += f"To: {recipient}\n"
        header_content += f"Date: {date_header}\n"
        header_content += "=" * 80 + "\n\n"

        # Extract body content
        text_content = self._extract_text_content(msg)
        html_content = self._extract_html_content(msg)

        return header_content + text_content, html_content

    def _extract_text_content(self, msg) -> str:
        """
        Extract plain text content from email message.

        Args:
            msg: Parsed email message object

        Returns:
            str: Plain text content if found
        """
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    try:
                        return part.get_content()
                    except Exception:
                        continue
        else:
            if msg.get_content_type() == "text/plain":
                try:
                    return msg.get_content()
                except Exception:
                    pass
        return ""

    def _extract_html_content(self, msg) -> str | None:
        """
        Extract HTML content from email message.

        Args:
            msg: Parsed email message object

        Returns:
            Optional[str]: HTML content if found
        """
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/html":
                    try:
                        return part.get_content()
                    except Exception:
                        continue
        else:
            if msg.get_content_type() == "text/html":
                try:
                    return msg.get_content()
                except Exception:
                    pass
        return None

    def _render_html_content(self, html_content: str) -> str:
        """
        Convert HTML content to readable text similar to email client rendering.

        Args:
            html_content: Raw HTML content

        Returns:
            str: Rendered text content
        """
        # Parse HTML with BeautifulSoup
        soup = BeautifulSoup(html_content, "html.parser")

        # Remove script and style elements
        for script in soup(["script", "style"]):
            script.decompose()

        # Decode HTML entities
        content = html.unescape(str(soup))

        # Clean and format the content
        content = self._clean_html_tags(content)
        content = self._clean_text_formatting(content)

        return content

    def _clean_html_tags(self, html_content: str) -> str:
        """
        Remove HTML tags while preserving text structure and formatting.

        Args:
            html_content: HTML content with tags

        Returns:
            str: Clean text content with preserved formatting
        """
        # Replace common HTML tags with appropriate text formatting
        replacements = [
            (r"<br\s*/?>", "\n"),  # Line breaks
            (r"<p[^>]*>", "\n\n"),  # Paragraphs
            (r"</p>", "\n"),
            (r"<div[^>]*>", "\n"),  # Divs
            (r"</div>", "\n"),
            (r"<h[1-6][^>]*>", "\n\n"),  # Headers
            (r"</h[1-6]>", "\n"),
            (r"<li[^>]*>", "\n• "),  # List items
            (r"</li>", ""),
            (r"<ul[^>]*>", "\n"),  # Lists
            (r"</ul>", "\n"),
            (r"<ol[^>]*>", "\n"),  # Ordered lists
            (r"</ol>", "\n"),
            (r"<strong[^>]*>", "**"),  # Bold
            (r"</strong>", "**"),
            (r"<b[^>]*>", "**"),  # Bold
            (r"</b>", "**"),
            (r"<em[^>]*>", "*"),  # Italic
            (r"</em>", "*"),
            (r"<i[^>]*>", "*"),  # Italic
            (r"</i>", "*"),
            (r"<u[^>]*>", "_"),  # Underline
            (r"</u>", "_"),
            (r'<a[^>]*href="([^"]*)"[^>]*>', r"[\1]"),  # Links
            (r"</a>", ""),
            (r"<blockquote[^>]*>", "\n> "),  # Blockquotes
            (r"</blockquote>", "\n"),
            (r"<hr[^>]*>", "\n" + "-" * 40 + "\n"),  # Horizontal rules
            (r"<table[^>]*>", "\n"),  # Tables
            (r"</table>", "\n"),
            (r"<tr[^>]*>", "\n"),  # Table rows
            (r"</tr>", "\n"),
            (r"<td[^>]*>", " | "),  # Table cells
            (r"</td>", ""),
            (r"<th[^>]*>", " | "),  # Table headers
            (r"</th>", ""),
        ]

        content = html_content
        for pattern, replacement in replacements:
            content = re.sub(pattern, replacement, content, flags=re.IGNORECASE)

        # Remove any remaining HTML tags
        content = re.sub(r"<[^>]+>", "", content)

        return content

    def _clean_text_formatting(self, text: str) -> str:
        """
        Clean up text formatting and whitespace.

        Args:
            text: Raw text content

        Returns:
            str: Cleaned text content
        """
        # Remove excessive whitespace
        text = re.sub(r"\n\s*\n\s*\n", "\n\n", text)  # Multiple blank lines
        text = re.sub(r" +", " ", text)  # Multiple spaces
        text = re.sub(r"\t+", "    ", text)  # Tabs to spaces

        # Clean up line breaks
        text = re.sub(r"\n\s*\n\s*\n", "\n\n", text)

        # Strip leading/trailing whitespace
        text = text.strip()

        return text

    def _generate_pdf_filename(self, date: datetime, prefix: str, output_directory: Path) -> Path:
        """
        Generate PDF filename with conflict resolution.

        Args:
            date: Email date
            prefix: Filename prefix from original name
            output_directory: Output directory for PDFs

        Returns:
            Path: Final PDF file path
        """
        formatted_date = date.strftime("%Y-%m-%d")
        base_filename = f"{formatted_date}_Email_{prefix}.pdf"
        pdf_path = output_directory / base_filename

        # Handle filename conflicts
        counter = 1
        while pdf_path.exists():
            base_filename = f"{formatted_date}_Email_{prefix}_{counter}.pdf"
            pdf_path = output_directory / base_filename
            counter += 1

        return pdf_path

    def _prepare_html_for_pdf(self, html_content: str, subject: str, sender: str, recipient: str, date: str) -> str:
        """
        Prepare HTML content for PDF rendering with proper styling.

        Args:
            html_content: Raw HTML content
            subject: Email subject
            sender: Email sender
            recipient: Email recipient
            date: Email date

        Returns:
            str: Complete HTML document with styling
        """
        # Parse and clean the HTML
        soup = BeautifulSoup(html_content, "html.parser")

        # Remove script tags and other problematic elements
        for script in soup(["script", "style", "meta", "link"]):
            script.decompose()

        # Create a complete HTML document with email client-like styling
        html_template = f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>{subject}</title>
    <style>
        body {{
            font-family: Arial, sans-serif;
            line-height: 1.6;
            color: #333;
            max-width: 800px;
            margin: 0 auto;
            padding: 20px;
            background-color: #f9f9f9;
        }}
        .email-header {{
            background-color: #ffffff;
            padding: 20px;
            margin-bottom: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }}
        .email-header h1 {{
            color: #2c3e50;
            margin: 0 0 10px 0;
            font-size: 18px;
        }}
        .email-header p {{
            margin: 5px 0;
            color: #666;
            font-size: 14px;
        }}
        .email-content {{
            background-color: #ffffff;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }}
        .email-content img {{
            max-width: 100%;
            height: auto;
            border-radius: 4px;
        }}
        .email-content a {{
            color: #3498db;
            text-decoration: none;
        }}
        .email-content a:hover {{
            text-decoration: underline;
        }}
        .email-content h1, .email-content h2, .email-content h3 {{
            color: #2c3e50;
        }}
        .email-content ul, .email-content ol {{
            padding-left: 20px;
        }}
        .email-content blockquote {{
            border-left: 4px solid #3498db;
            margin: 10px 0;
            padding-left: 15px;
            color: #666;
        }}
        .separator {{
            border-top: 1px solid #eee;
            margin: 20px 0;
        }}
    </style>
</head>
<body>
    <div class="email-header">
        <h1>{subject}</h1>
        <p><strong>From:</strong> {sender}</p>
        <p><strong>To:</strong> {recipient}</p>
        <p><strong>Date:</strong> {date}</p>
    </div>
    <div class="email-content">
        {str(soup)}
    </div>
</body>
</html>
        """

        return html_template

    async def _create_pdf_from_html(self, html_content: str, output_path: Path) -> None:
        """
        Create PDF from HTML content using Playwright.

        Args:
            html_content: Complete HTML document
            output_path: Path where PDF should be saved
        """
        if not self.browser:
            print("⚠️ Browser not initialized, falling back to text-based PDF")
            self._create_text_pdf(html_content, output_path)
            return

        try:
            # Create a new page in the existing browser
            page = await self.browser.new_page()
            await page.set_content(html_content)
            await page.pdf(path=str(output_path), format="A4")
            await page.close()
        except Exception as e:
            print(f"⚠️ HTML rendering failed: {e}")
            # Fallback to text-based PDF
            self._create_text_pdf(html_content, output_path)

    def _create_text_pdf(self, content: str, output_path: Path) -> None:
        """
        Create PDF from text content using FPDF (fallback method).

        Args:
            content: Text content
            output_path: Path where PDF should be saved
        """
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)

        # Handle Unicode characters more robustly
        try:
            # Clean the content by removing problematic characters
            content_cleaned = ""
            for char in content:
                # Replace problematic Unicode characters with ASCII equivalents
                if ord(char) > 127:
                    # Replace common Unicode characters
                    if char == "—":
                        content_cleaned += "-"
                    elif char == "–":
                        content_cleaned += "-"
                    elif char == '"':
                        content_cleaned += '"'
                    elif char == '"':
                        content_cleaned += '"'
                    elif char == "":
                        content_cleaned += "'"
                    elif char == "":
                        content_cleaned += "'"
                    elif char == "…":
                        content_cleaned += "..."
                    elif char == "\u200c":  # Zero-width non-joiner
                        continue  # Skip this character
                    else:
                        # For other Unicode characters, try to normalize or replace
                        try:
                            normalized = unicodedata.normalize("NFKD", char)
                            if normalized.isascii():
                                content_cleaned += normalized
                            else:
                                content_cleaned += "?"  # Replace with question mark
                        except:
                            content_cleaned += "?"  # Replace with question mark
                else:
                    content_cleaned += char

            pdf.multi_cell(0, 10, content_cleaned)
        except Exception:
            # Final fallback: encode to ASCII with replacement
            try:
                content_safe = content.encode("ascii", errors="replace").decode("ascii")
                pdf.multi_cell(0, 10, content_safe)
            except Exception:
                # Last resort: create a minimal PDF
                pdf.multi_cell(0, 10, "PDF creation failed due to encoding issues.")

        pdf.output(str(output_path))

    async def convert_eml_files(self, input_directory: Path, output_directory: Path) -> None:
        """
        Convert all .eml files in directory to PDFs with date-based naming.

        Args:
            input_directory: Directory containing .eml files
            output_directory: Directory where PDFs will be saved
        """
        if not input_directory.exists():
            print(f"❌ Error: The input directory '{input_directory}' does not exist.")
            return

        # Create output directory if it doesn't exist
        output_directory.mkdir(exist_ok=True)
        print(f"📁 Output directory: {output_directory}")

        eml_files = list(input_directory.glob("*.eml"))
        if not eml_files:
            print(f"⚠️ No .eml files found in {input_directory}")
            return

        print(f"📁 Found {len(eml_files)} .eml files in {input_directory}")

        # Initialize browser for HTML rendering
        print("🚀 Initializing browser for HTML rendering...")
        await self._initialize_browser()

        try:
            # Create semaphore to limit concurrency
            semaphore = asyncio.Semaphore(50)

            # Create tasks for concurrent processing
            tasks = []
            for eml_path in eml_files:
                task = self._process_single_eml_file_async(eml_path, output_directory, semaphore)
                tasks.append(task)

            # Wait for all tasks to complete
            await asyncio.gather(*tasks, return_exceptions=True)

        finally:
            # Clean up browser resources
            print("🧹 Cleaning up browser resources...")
            await self._cleanup_browser()

    async def _process_single_eml_file_async(
        self, eml_path: Path, output_directory: Path, semaphore: asyncio.Semaphore
    ) -> None:
        """
        Process a single EML file and convert it to PDF (async version with semaphore).

        Args:
            eml_path: Path to the EML file
            output_directory: Directory where PDF will be saved
            semaphore: Semaphore to limit concurrency
        """
        async with semaphore:
            try:
                await self._process_single_eml_file_async_internal(eml_path, output_directory)
            except Exception as e:
                print(f"❌ Error processing {eml_path.name}: {e}")
                self.failed_files.append(eml_path.name)

    async def _process_single_eml_file_async_internal(self, eml_path: Path, output_directory: Path) -> None:
        """
        Internal async method to process a single EML file.

        Args:
            eml_path: Path to the EML file
            output_directory: Directory where PDF will be saved
        """
        print(f"\n📄 Processing: {eml_path.name}")

        # Extract filename prefix
        base_name = eml_path.stem
        prefix = self._extract_filename_prefix(base_name)

        # Parse EML file
        parser = Parser(policy=policy.default)
        with open(eml_path, encoding="utf-8", errors="ignore") as f:
            msg = parser.parse(f)

        # Extract and parse date
        date_obj = self._parse_email_date(msg)
        if not date_obj:
            print(f"⚠️ No valid date found in: {eml_path.name}")
            self.failed_files.append(eml_path.name)
            return

        print(f"📅 Found date: {date_obj}")

        # Extract email content
        subject = msg.get("Subject", "No Subject")
        sender = msg.get("From", "Unknown Sender")
        recipient = msg.get("To", "Unknown Recipient")
        date_header = msg.get("Date", "No Date")

        text_content, html_content = self._extract_email_content(msg)

        # Generate PDF filename and path
        pdf_path = self._generate_pdf_filename(date_obj, prefix, output_directory)

        # Create PDF
        if html_content:
            # Use HTML rendering for better formatting
            html_document = self._prepare_html_for_pdf(html_content, subject, sender, recipient, date_header)
            await self._create_pdf_from_html(html_document, pdf_path)
        else:
            # Fallback to text-based PDF
            self._create_text_pdf(text_content, pdf_path)

        self.converted_files.append(pdf_path.name)
        print(f"✅ Converted to: {pdf_path.name}")

    def print_summary(self) -> None:
        """Print conversion summary."""
        print("\n📊 Summary:")
        print(f"Converted: {len(self.converted_files)} files")
        for name in self.converted_files:
            print(f"  ➤ {name}")

        if self.failed_files:
            print(f"\n⚠️ Failed to convert {len(self.failed_files)} files:")
            for name in self.failed_files:
                print(f"  ✗ {name}")


async def main() -> None:
    """Main function to run the EML to PDF converter."""
    print("📧 EML to PDF Converter")
    print("=" * 50)

    # Check for headless flag
    headless = "--headless" in sys.argv

    if headless:
        # Check if directory path is provided as argument
        if len(sys.argv) < 3:
            print("❌ Error: Headless mode requires a directory path.")
            print("Usage: python eml_to_pdf_converter.py --headless <directory_path>")
            sys.exit(1)

        input_directory = Path(sys.argv[2])
        if not input_directory.exists():
            print(f"❌ Error: The directory '{input_directory}' does not exist.")
            sys.exit(1)

        # Create output directory in same location as input
        output_directory = input_directory / "PDFs"
        print(f"🔧 Headless mode: Using input directory: {input_directory}")
        print(f"🔧 Headless mode: Using output directory: {output_directory}")
    else:
        # Use file dialog for interactive mode
        selector = DirectorySelector()
        input_directory = selector.select_directory()

        if not input_directory:
            print("❌ No directory selected. Exiting.")
            sys.exit(1)

        # Create output directory in same location as input
        output_directory = input_directory / "PDFs"

    print(f"📁 Input directory: {input_directory}")

    # Convert files
    converter = EMLToPDFConverter()
    await converter.convert_eml_files(input_directory, output_directory)
    converter.print_summary()


def run_main() -> None:
    """Run the async main function."""
    asyncio.run(main())


if __name__ == "__main__":
    run_main()
