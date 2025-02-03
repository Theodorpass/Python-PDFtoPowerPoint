# PDF to PowerPoint Conversion

This script extracts text and images from a PDF file and converts them into a PowerPoint presentation. It uses the PyMuPDF library for extracting PDF content, Pillow for handling images, and python-pptx for generating the PowerPoint slides.

## Requirements

To run this script, you need to install the following Python libraries:

``bash
pip install -r requirements.txt


1 ) Clone or download this repository to your local machine.
  git clone https://github.com/Theodorpass/PDF-to-PowerPoint.git



2) Navigate to the project folder:
``bash 
  cd PDF-to-PowerPoint

3) Install the required dependencies:
  pip install -r requirements.txt


4) Open the pdf_to_pptx.py script and modify the pdf_file and pptx_file paths to your specific files:

  pdf_file = r'C:\path\to\your\input.pdf'  # Replace with your input PDF file path
  pptx_file = r'C:\path\to\your\output.pptx'  # Replace with your desired output PowerPoint file path

5) Run the script

python pdf_to_pptx.py

6) After execution, check the destination folder for the newly created PowerPoint file containing the extracted text and images from the PDF.

