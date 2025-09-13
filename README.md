## 📄 Advanced Document Converter
A powerful Streamlit-based web application for converting between various document formats. This tool provides a user-friendly interface for all your document conversion needs.

https://img.shields.io/badge/Streamlit-FF4B4B?style=for-the-badge&logo=Streamlit&logoColor=white
https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white
https://img.shields.io/badge/PDF-FF0000?style=for-the-badge&logo=adobeacrobatreader&logoColor=white

# ✨ Features
🔄 Conversion Tools
PDF to Word - Extract text from PDF to Word documents

Word to PDF - Convert Word documents to PDF format

Merge PDFs - Combine multiple PDF files into one

Split PDF - Split PDF into single pages or extract page ranges

Compress PDF - Reduce PDF file size with adjustable compression

PDF to PowerPoint - Convert PDF content to presentations

PDF to JPG - Extract pages from PDF as images

JPG to PDF - Combine multiple images into a PDF

PDF to Excel - Extract text from PDF to spreadsheet format


# 🎨 User Interface
Clean, modern design with custom CSS styling

Responsive layout that works on all screen sizes

Intuitive tabbed interface for easy navigation

Real-time file preview and information

Progress indicators during conversion

Download buttons for converted files

# 🚀 Installation
Prerequisites
Python 3.7 or higher

LibreOffice (for Word to PDF and Excel to PDF conversion)

Step 1: Clone or Download the Project
bash
git clone <your-repository-url>
cd advanced-document-converter
Step 2: Install Python Dependencies
bash
pip install -r requirements.txt
If you don't have a requirements.txt file, install packages individually:

bash
pip install streamlit PyMuPDF python-docx python-pptx pillow pandas reportlab openpyxl img2pdf
Step 3: Install LibreOffice
Windows:

Download from LibreOffice Official Website

Follow the installation wizard

macOS:

bash
brew install libreoffice
Ubuntu/Debian:

bash
sudo apt-get update
sudo apt-get install libreoffice
Other Linux Distributions:

Use your distribution's package manager or download from the official website

# 🏃‍♂️ Usage
Running the Application
bash
streamlit run app.py
The application will open in your default web browser at http://localhost:8501

How to Use
Select Conversion Type: Choose from the sidebar options

Upload File(s): Use the file uploader to select your document(s)

Configure Options: Adjust settings if needed (compression level, page range, etc.)

Convert: Click the conversion button

Download: Use the download button to save your converted file

Supported File Formats
Input: PDF, DOCX, DOC, JPG, JPEG, PNG, XLSX, XLS

Output: PDF, DOCX, PPTX, JPG, XLSX, ZIP

# 🛠️ Technical Details
Architecture
text
advanced-document-converter/
│
├── app.py              # Main Streamlit application
├── docs.py             # Additional PDF manipulation functions
├── temp/               # Temporary files directory (auto-created)
├── output/             # Output files directory (auto-created)
└── requirements.txt    # Python dependencies
Key Libraries Used
Streamlit: Web application framework

PyMuPDF (fitz): PDF manipulation and text extraction

python-docx: Word document creation and manipulation

python-pptx: PowerPoint presentation creation

Pillow (PIL): Image processing

pandas: Data manipulation for Excel conversion

ReportLab: PDF generation from images

img2pdf: Image to PDF conversion

openpyxl: Excel file manipulation

Platform Support
✅ Windows (Full functionality with LibreOffice)

✅ macOS (Full functionality with LibreOffice)

✅ Linux (Full functionality with LibreOffice)

# 🔧 Troubleshooting
Common Issues
LibreOffice not found error

Ensure LibreOffice is installed and accessible in your system PATH

Restart your terminal/command prompt after installation

Conversion fails for large files

The application may time out for very large files (>100MB)

Try compressing files first or splitting them into smaller parts

Memory errors

The application may struggle with very large files on systems with limited RAM

Close other applications to free up memory

Formatting issues

Complex formatting may not be preserved perfectly in conversions

Some advanced PDF features (forms, annotations) won't be converted

Performance Tips
For large PDFs, use compression before other operations

Split large files into smaller chunks for better performance

Close other browser tabs to improve application responsiveness

# 📁 Project Structure
text
advanced-document-converter/
│
├── app.py                 # Main application file
├── docs.py                # PDF manipulation functions
├── temp/                  # Temporary storage for processing
├── output/                # Storage for converted files
├── requirements.txt       # Python dependencies
├── README.md             # This file
└── .gitignore           # Git ignore file

# 🤝 Contributing
We welcome contributions! Please feel free to submit issues, feature requests, or pull requests.

Development Setup
Fork the repository

Create a feature branch: git checkout -b feature-name

Make your changes and test thoroughly

Commit your changes: git commit -m 'Add some feature'

Push to the branch: git push origin feature-name

Submit a pull request

- Testing
  - Test all conversion types with various file formats
  
  - Verify cross-platform compatibility
  
  - Check error handling with invalid files

# 📄 License
This project is open source and available under the MIT License.

# 🙏 Acknowledgments
Streamlit for the amazing web framework

PyMuPDF for PDF manipulation capabilities

LibreOffice for document conversion capabilities

All the open-source libraries that make this project possible

# 📞 Support
If you encounter any issues or have questions:

Check the Troubleshooting section

Search existing GitHub Issues

Create a new issue with detailed information about your problem

# 🔄 Version History
v1.0.0 (Current)

Initial release with all core conversion features

Cross-platform support

Responsive web interface


Note: This application processes files on your local machine. No documents are uploaded to external servers, ensuring your data privacy and security.
