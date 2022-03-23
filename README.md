# PDF-To-Excel
This script utilizes OpenCV for image recognition and PyTesseract for object character recognition to convert tables in a PDF to Excel.

In order to use this script, you need to install 4 libraries. Pdf2image, Opencv, openpyxl, and pytesseract.

Install pdf2image by using their guide as you will need poppler. Same thing to Pytesseract, but make sure to install it to the script folder and name the extracted folder "Tesseract"

Once you run the script and everything is functional, it will ask you to provide a path to an excel file. Simply write the path and press the convert button. It will then convert each page in the PDF to a JPEG. Once converted, drag and drop a box over the table it is you want to copy and press ENTER to confirm. Press C to cancel the selection.
