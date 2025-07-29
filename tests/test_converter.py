"""
Tests for the EML to PDF converter.

This module contains comprehensive tests for the EmailToPDFConverter class
and related functionality.
"""

from datetime import datetime
from pathlib import Path
from unittest.mock import AsyncMock, Mock, patch

import pytest

from eml_to_pdf_converter import DirectorySelector, EmailToPDFConverter


class TestEmailToPDFConverter:
    """Test cases for EmailToPDFConverter class."""

    def setup_method(self) -> None:
        """Set up test fixtures."""
        self.converter = EmailToPDFConverter()

    def test_extract_filename_prefix_basic(self) -> None:
        """Test basic filename prefix extraction."""
        result = self.converter._extract_filename_prefix("John Doe Inquiry")
        assert result == "John_Doe_Inquiry"

    def test_extract_filename_prefix_with_special_chars(self) -> None:
        """Test filename prefix extraction with special characters."""
        result = self.converter._extract_filename_prefix("John@Doe!Inquiry#")
        assert result == "JohnDoeInquiry"

    def test_extract_filename_prefix_short(self) -> None:
        """Test filename prefix extraction with fewer than 3 words."""
        result = self.converter._extract_filename_prefix("Short")
        assert result == "Short"

    def test_extract_filename_prefix_empty(self) -> None:
        """Test filename prefix extraction with empty string."""
        result = self.converter._extract_filename_prefix("")
        assert result == "NoName"

    def test_parse_email_date_valid(self) -> None:
        """Test parsing valid email date."""
        msg_data = {"Date": "Mon, 15 Jan 2024 10:30:00 +0000"}

        result = self.converter._parse_email_date(msg_data)
        assert result is not None
        assert isinstance(result, datetime)
        assert result.year == 2024
        assert result.month == 1
        assert result.day == 15

    def test_parse_email_date_missing(self) -> None:
        """Test parsing email with missing date."""
        msg_data = {"Date": None}

        result = self.converter._parse_email_date(msg_data)
        assert result is None

    def test_parse_email_date_invalid(self) -> None:
        """Test parsing email with invalid date format."""
        msg_data = {"Date": "Invalid Date String"}

        result = self.converter._parse_email_date(msg_data)
        assert result is None

    def test_extract_email_content(self) -> None:
        """Test email content extraction from unified message data."""
        msg_data = {
            "Subject": "Test Subject",
            "From": "test@example.com",
            "To": "recipient@example.com",
            "Date": "Mon, 15 Jan 2024 10:30:00 +0000",
            "body": "Test email body content",
            "htmlBody": "<p>Test HTML content</p>",
        }

        text_content, html_content = self.converter._extract_email_content(msg_data)

        assert "Subject: Test Subject" in text_content
        assert "From: test@example.com" in text_content
        assert "To: recipient@example.com" in text_content
        assert "Date: Mon, 15 Jan 2024 10:30:00 +0000" in text_content
        assert "Test email body content" in text_content
        assert html_content == "<p>Test HTML content</p>"

    def test_extract_email_content_missing_fields(self) -> None:
        """Test email content extraction with missing fields."""
        msg_data = {
            "Subject": None,
            "From": None,
            "To": None,
            "Date": None,
            "body": "",
            "htmlBody": "",
        }

        text_content, html_content = self.converter._extract_email_content(msg_data)

        assert "Subject: No Subject" in text_content
        assert "From: Unknown Sender" in text_content
        assert "To: Unknown Recipient" in text_content
        assert "Date: No Date" in text_content
        assert html_content == ""

    def test_generate_pdf_filename_basic(self, tmp_path: Path) -> None:
        """Test basic PDF filename generation."""
        date = datetime(2024, 1, 15)
        prefix = "Test_Prefix"

        result = self.converter._generate_pdf_filename(date, prefix, tmp_path)

        expected_name = "2024-01-15_Email_Test_Prefix.pdf"
        assert result.name == expected_name
        assert result.parent == tmp_path

    def test_generate_pdf_filename_conflict(self, tmp_path: Path) -> None:
        """Test PDF filename generation with conflict resolution."""
        date = datetime(2024, 1, 15)
        prefix = "Test_Prefix"

        # Create a file with the same name
        existing_file = tmp_path / "2024-01-15_Email_Test_Prefix.pdf"
        existing_file.touch()

        result = self.converter._generate_pdf_filename(date, prefix, tmp_path)

        expected_name = "2024-01-15_Email_Test_Prefix_1.pdf"
        assert result.name == expected_name

    def test_generate_pdf_filename_multiple_conflicts(self, tmp_path: Path) -> None:
        """Test PDF filename generation with multiple conflicts."""
        date = datetime(2024, 1, 15)
        prefix = "Test_Prefix"

        # Create multiple files with similar names
        for i in range(3):
            existing_file = tmp_path / f"2024-01-15_Email_Test_Prefix{'_' + str(i) if i > 0 else ''}.pdf"
            existing_file.touch()

        result = self.converter._generate_pdf_filename(date, prefix, tmp_path)

        expected_name = "2024-01-15_Email_Test_Prefix_3.pdf"
        assert result.name == expected_name

    @patch("eml_to_pdf_converter.FPDF")
    def test_create_text_pdf(self, mock_fpdf: Mock, tmp_path: Path) -> None:
        """Test text-based PDF creation."""
        content = "Test PDF content"
        output_path = tmp_path / "test.pdf"

        self.converter._create_text_pdf(content, output_path)

        # Verify FPDF was called correctly
        mock_fpdf.assert_called_once()
        mock_pdf = mock_fpdf.return_value
        mock_pdf.add_page.assert_called_once()
        mock_pdf.set_font.assert_called_once_with("Arial", size=12)
        mock_pdf.multi_cell.assert_called_once()
        mock_pdf.output.assert_called_once_with(str(output_path))

    def test_clean_html_tags(self) -> None:
        """Test HTML tag cleaning."""
        html_content = "<p>This is a <strong>test</strong> with <br/>line breaks.</p>"
        result = self.converter._clean_html_tags(html_content)

        # Should remove HTML tags and preserve text
        assert "<p>" not in result
        assert "<strong>" not in result
        assert "<br/>" not in result
        assert "This is a" in result
        assert "test" in result
        assert "line breaks" in result

    def test_clean_text_formatting(self) -> None:
        """Test text formatting cleanup."""
        text = "  Multiple    spaces\n\n\n\nand\t\ttabs  "
        result = self.converter._clean_text_formatting(text)

        # Should clean up excessive whitespace
        assert "Multiple spaces" in result
        assert "and    tabs" in result  # tabs converted to spaces
        assert result.strip() == result  # no leading/trailing whitespace

    def test_prepare_html_for_pdf(self) -> None:
        """Test HTML preparation for PDF rendering."""
        html_content = "<p>Test content</p>"
        subject = "Test Subject"
        sender = "test@example.com"
        recipient = "recipient@example.com"
        date = "Mon, 15 Jan 2024 10:30:00 +0000"

        result = self.converter._prepare_html_for_pdf(html_content, subject, sender, recipient, date)

        assert "<!DOCTYPE html>" in result
        assert "<html>" in result
        assert subject in result
        assert sender in result
        assert recipient in result
        assert date in result
        assert "<p>Test content</p>" in result

    @patch("eml_to_pdf_converter.extract_msg")
    def test_parse_msg_file_success(self, mock_extract_msg: Mock) -> None:
        """Test successful MSG file parsing."""
        mock_msg = Mock()
        mock_msg.subject = "Test Subject"
        mock_msg.sender = "test@example.com"
        mock_msg.to = "recipient@example.com"
        mock_msg.date = "Mon, 15 Jan 2024 10:30:00 +0000"
        mock_msg.body = "Test body"
        mock_msg.htmlBody = "<p>Test HTML</p>"
        mock_msg.attachments = []

        mock_extract_msg.Message.return_value = mock_msg

        msg_path = Path("test.msg")
        result = self.converter._parse_msg_file(msg_path)

        assert result is not None
        assert result["Subject"] == "Test Subject"
        assert result["From"] == "test@example.com"
        assert result["To"] == "recipient@example.com"
        assert result["Date"] == "Mon, 15 Jan 2024 10:30:00 +0000"
        assert result["body"] == "Test body"
        assert result["htmlBody"] == "<p>Test HTML</p>"
        assert result["attachments"] == []

    @patch("eml_to_pdf_converter.extract_msg")
    def test_parse_msg_file_error(self, mock_extract_msg: Mock) -> None:
        """Test MSG file parsing with error."""
        mock_extract_msg.Message.side_effect = Exception("Test error")

        msg_path = Path("test.msg")
        result = self.converter._parse_msg_file(msg_path)

        assert result is None

    @patch("eml_to_pdf_converter.Parser")
    def test_parse_eml_file_success(self, mock_parser: Mock, tmp_path: Path) -> None:
        """Test successful EML file parsing."""
        mock_msg = Mock()

        def mock_get(key, default=""):
            return {
                "Subject": "Test Subject",
                "From": "test@example.com",
                "To": "recipient@example.com",
                "Date": "Mon, 15 Jan 2024 10:30:00 +0000",
            }.get(key, default)

        mock_msg.get.side_effect = mock_get

        # Mock the text and HTML content extraction
        with (
            patch.object(self.converter, "_extract_text_content", return_value="Test body"),
            patch.object(self.converter, "_extract_html_content", return_value="<p>Test HTML</p>"),
        ):
            mock_parser.return_value.parse.return_value = mock_msg

            # Create a temporary EML file
            eml_path = tmp_path / "test.eml"
            eml_path.write_text("From: test@example.com\nSubject: Test\n\nBody")

            result = self.converter._parse_eml_file(eml_path)

            assert result is not None
            assert result["Subject"] == "Test Subject"
            assert result["From"] == "test@example.com"
            assert result["To"] == "recipient@example.com"
            assert result["Date"] == "Mon, 15 Jan 2024 10:30:00 +0000"
            assert result["body"] == "Test body"
            assert result["htmlBody"] == "<p>Test HTML</p>"
            assert result["attachments"] == []

    @patch("eml_to_pdf_converter.Parser")
    def test_parse_eml_file_error(self, mock_parser: Mock) -> None:
        """Test EML file parsing with error."""
        mock_parser.return_value.parse.side_effect = Exception("Test error")

        eml_path = Path("test.eml")
        result = self.converter._parse_eml_file(eml_path)

        assert result is None

    def test_print_summary_empty(self, capsys: pytest.CaptureFixture) -> None:
        """Test summary printing with no conversions."""
        self.converter.print_summary()

        captured = capsys.readouterr()
        assert "Converted: 0 files" in captured.out
        assert "Failed to convert" not in captured.out

    def test_print_summary_with_conversions(self, capsys: pytest.CaptureFixture) -> None:
        """Test summary printing with successful conversions."""
        self.converter.converted_files = ["file1.pdf", "file2.pdf"]
        self.converter.failed_files = ["file3.eml"]

        self.converter.print_summary()

        captured = capsys.readouterr()
        assert "Converted: 2 files" in captured.out
        assert "➤ file1.pdf" in captured.out
        assert "➤ file2.pdf" in captured.out
        assert "Failed to convert 1 files" in captured.out
        assert "✗ file3.eml" in captured.out


class TestDirectorySelector:
    """Test cases for DirectorySelector class."""

    def test_select_directory_success(self) -> None:
        """Test successful directory selection."""
        with patch("tkinter.filedialog.askdirectory") as mock_askdirectory, patch("tkinter.Tk") as mock_tk:
            mock_askdirectory.return_value = "/test/path"
            mock_tk_instance = Mock()
            mock_tk.return_value = mock_tk_instance

            selector = DirectorySelector()
            result = selector.select_directory()

            assert result == Path("/test/path")
            mock_askdirectory.assert_called_once()
            mock_tk_instance.withdraw.assert_called_once()
            mock_tk_instance.attributes.assert_called_once()
            mock_tk_instance.destroy.assert_called_once()

    def test_select_directory_cancelled(self) -> None:
        """Test directory selection when cancelled."""
        with patch("tkinter.filedialog.askdirectory") as mock_askdirectory, patch("tkinter.Tk") as mock_tk:
            mock_askdirectory.return_value = ""
            mock_tk_instance = Mock()
            mock_tk.return_value = mock_tk_instance

            selector = DirectorySelector()
            result = selector.select_directory()

            assert result is None

    def test_select_directory_error(self) -> None:
        """Test directory selection with error."""
        with patch("tkinter.filedialog.askdirectory") as mock_askdirectory, patch("tkinter.Tk") as mock_tk:
            mock_askdirectory.side_effect = Exception("Test error")
            mock_tk_instance = Mock()
            mock_tk.return_value = mock_tk_instance

            selector = DirectorySelector()
            result = selector.select_directory()

            assert result is None


# Integration tests
class TestIntegration:
    """Integration tests for the complete workflow."""

    def setup_method(self) -> None:
        """Set up test fixtures."""
        self.converter = EmailToPDFConverter()

    @pytest.mark.asyncio
    async def test_full_conversion_workflow_eml(self, tmp_path: Path) -> None:
        """Test the complete conversion workflow with EML file."""
        # Create a mock .eml file
        eml_content = """From: test@example.com
To: recipient@example.com
Subject: Test Email
Date: Mon, 15 Jan 2024 10:30:00 +0000

This is a test email body.
"""

        eml_file = tmp_path / "Test Email.eml"
        eml_file.write_text(eml_content)

        # Mock the async browser methods and PDF creation
        with (
            patch.object(self.converter, "_initialize_browser"),
            patch.object(self.converter, "_cleanup_browser"),
            patch.object(self.converter, "_create_pdf_from_html", new_callable=AsyncMock),
            patch.object(self.converter, "_parse_eml_file") as mock_parse_eml,
        ):
            # Mock successful EML parsing
            mock_parse_eml.return_value = {
                "Subject": "Test Email",
                "From": "test@example.com",
                "To": "recipient@example.com",
                "Date": "Mon, 15 Jan 2024 10:30:00 +0000",
                "body": "This is a test email body.",
                "htmlBody": "<p>Test HTML</p>",
                "attachments": [],
            }

            # Run conversion
            converter = EmailToPDFConverter()
            await converter.convert_email_files(tmp_path, tmp_path)

            # Check results
            assert len(converter.converted_files) == 1
            assert converter.converted_files[0] == "2024-01-15_Email_Test_Email.pdf"
            assert len(converter.failed_files) == 0

            # Since we have HTML content, it should use HTML rendering
            # The HTML PDF creation is mocked, so we just verify the test passes

    @pytest.mark.asyncio
    async def test_full_conversion_workflow_msg(self, tmp_path: Path) -> None:
        """Test the complete conversion workflow with MSG file."""
        # Create a mock .msg file (just a placeholder since we mock the parsing)
        msg_file = tmp_path / "Test Message.msg"
        msg_file.touch()

        # Mock the MSG parsing and async browser methods
        with (
            patch.object(self.converter, "_initialize_browser"),
            patch.object(self.converter, "_cleanup_browser"),
            patch.object(self.converter, "_create_pdf_from_html", new_callable=AsyncMock),
        ):
            # Mock successful MSG parsing at the class level
            with patch.object(EmailToPDFConverter, "_parse_msg_file") as mock_parse_msg:
                mock_parse_msg.return_value = {
                    "Subject": "Test Message",
                    "From": "test@example.com",
                    "To": "recipient@example.com",
                    "Date": "Mon, 15 Jan 2024 10:30:00 +0000",
                    "body": "Test message body",
                    "htmlBody": "<p>Test HTML</p>",
                    "attachments": [],
                }

                # Run conversion
                converter = EmailToPDFConverter()
                await converter.convert_email_files(tmp_path, tmp_path)

            # Check results
            assert len(converter.converted_files) == 1
            assert converter.converted_files[0] == "2024-01-15_Email_Test_Message.pdf"
            assert len(converter.failed_files) == 0

            # Since we have HTML content, it should use HTML rendering
            # The HTML PDF creation is mocked, so we just verify the test passes

    @pytest.mark.asyncio
    async def test_convert_email_files_no_directory(self) -> None:
        """Test conversion with non-existent directory."""
        non_existent_path = Path("/non/existent/path")

        await self.converter.convert_email_files(non_existent_path, non_existent_path)

        assert len(self.converter.converted_files) == 0
        assert len(self.converter.failed_files) == 0

    @pytest.mark.asyncio
    async def test_convert_email_files_no_files(self, tmp_path: Path) -> None:
        """Test conversion with directory containing no email files."""
        # Create some non-email files
        (tmp_path / "test.txt").touch()
        (tmp_path / "test.pdf").touch()

        await self.converter.convert_email_files(tmp_path, tmp_path)

        assert len(self.converter.converted_files) == 0
        assert len(self.converter.failed_files) == 0

    @pytest.mark.asyncio
    async def test_browser_initialization_failure(self, tmp_path: Path) -> None:
        """Test handling of browser initialization failure."""
        # Create a test file
        eml_file = tmp_path / "test.eml"
        eml_file.write_text("From: test@example.com\nSubject: Test\nDate: Mon, 15 Jan 2024 10:30:00 +0000\n\nBody")

        # Mock browser initialization to fail
        with (
            patch("eml_to_pdf_converter.async_playwright") as mock_playwright,
            patch.object(self.converter, "_create_text_pdf") as mock_create_pdf,
        ):
            mock_playwright.side_effect = Exception("Browser init failed")

            await self.converter.convert_email_files(tmp_path, tmp_path)

            # Should still attempt conversion with text fallback
            mock_create_pdf.assert_called_once()
