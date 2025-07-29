# Building Windows Executable using Virtual Machine

## Prerequisites
- VirtualBox, Parallels Desktop, or VMware Fusion installed on your Mac
- Windows 10/11 VM or ISO
- At least 4GB RAM allocated to the VM

## Steps

1. **Set up Windows VM**
   - Install Windows 10/11 in your VM
   - Install Python 3.9+ from python.org
   - Install Git for Windows

2. **Transfer your code**
   - Use shared folders or copy files to VM
   - Or clone your repository in the VM

3. **Install dependencies**
   ```cmd
   # Install uv (recommended)
   pip install uv

   # Or install with pip
   pip install pyinstaller
   pip install -r requirements.txt
   ```

4. **Install Playwright browsers (required for HTML rendering)**
   ```cmd
   python -m playwright install chromium
   ```

5. **Build the executable**
   ```cmd
   # Using uv (recommended)
   uv run pyinstaller --onefile --windowed --name "EmailToPDFConverter" eml_to_pdf_converter.py

   # Or using pip
   pyinstaller --onefile --windowed --name "EmailToPDFConverter" eml_to_pdf_converter.py
   ```

6. **Copy the executable back to Mac**
   - The executable will be in the `dist/` folder
   - Copy it to your Mac via shared folders

## Advantages
- Native Windows build environment
- More reliable than Wine
- Can test the executable in the VM

## Disadvantages
- Requires Windows license
- More resource intensive
- Slower build process