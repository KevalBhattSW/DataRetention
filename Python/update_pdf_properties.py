import os
import sys
import zipfile

from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject
import shutil
import re


# Function to check if PDF has a digital signature. If so, skip.
# Taken with thanks from https://stackoverflow.com/questions/4226479/scan-for-secured-pdf-documents
def is_signed(pdf_file_path):
    reader = PdfReader(pdf_file_path)
    root = reader.trailer['/Root'];l_signed=False
    if acroform := root.get('/AcroForm'):
        if acroform and (sig := acroform.get('/SigFlags')):
            l_signed = bool(sig & 1)
    return l_signed


# Function to check if PDF is encrypted. If so, skip.
# Taken from same page as above
def is_encrypted(pdf_file_path):
    reader = PdfReader(pdf_file_path)
    return reader.is_encrypted


# Function to create/update a custom property in a PDF file
# Reads the pdf using the pypdf2 library, outputs to a new PDF page by page, including properties, adds a new property then creates a new PDF
# Makes a back-up stored on the executing machine, for testing against the original document
def update_pdf_properties(property_name, property_value, pdf_file):
    pdf_file_copy = "c:\\temp\\unstructured\\backup\\2"
    if not os.path.exists(pdf_file):
        # print(f"Error: {pdf file} does not exist!", flush-True)
        sys.exit(1)
    if is_encrypted(pdf_file):
        return 1
    if is_signed(pdf_file):
        return 2
    # Check the path name of the original file to derive the path of the back-up location, to help with checking pdf integrity
    if (pdf_file.upper()).startswith("C:"):
        pdf_file_copy = pdf_file_copy + "\\" + pdf_file.replace(":", "")
    elif pdf_file.startswith("\\\\"):
        pdf_file_copy = pdf_file_copy + pdf_file.replace("\\\\", "\\")
    pdf_file_zip_copy = re.sub(r"\.pdf", r".zip", pdf_file_copy, flags=re.IGNORECASE)
    pdf_file_zip = re.sub(r"\.pdf", r".zip", pdf_file, flags=re.IGNORECASE)
    pdf_file_output = re.sub(r"\.pdf", r"_updated.pdf", pdf_file, flags=re.IGNORECASE)

    # Back up original PDF
    if not os.path.isfile(pdf_file_zip_copy):
        with zipfile.ZipFile(pdf_file_zip, "w", zipfile.ZIP_DEFLATED) as zip_file:
            zip_file.write(pdf_file)

    if not os.path.exists(os.path.dirname(pdf_file_zip_copy)):
        os.makedirs(os.path.dirname(pdf_file_zip_copy), exist_ok=True)
    shutil.move(pdf_file_zip, pdf_file_zip_copy)

    try:
        reader = PdfReader(pdf_file)
        writer = PdfWriter()

        # Copy pages
        for page in reader.pages:
            writer.add_page(page)

        # Copy and update metadata
        metadata = reader.metadata or {}
        new_metadata = {NameObject(k): str(v) for k, v in metadata.items() if v is not None}
        new_metadata[NameObject(f"/{property_name}")] = str(property_value)
        writer.add_metadata(new_metadata)

        # Write updated PDF
        with open(pdf_file_output, "wb") as f_out:
            writer.write(f_out)
        # Replace original file
        os.remove(pdf_file)
        os.rename(pdf_file_output, pdf_file)
        # Remove backup
        # if os.path.exists (pdf file zip_copy):
        #

        os.remove(pdf_file_zip_copy)
        return 0

    # In case of error, remove the zipped/updated copies
    except Exception as e:
        print(f" [Error] {e}")
        if os.path.exists(pdf_file_output):
            os.remove(pdf_file_output)
        if os.path.exists(pdf_file_zip):
            os.remove(pdf_file_zip)
        if os.path.exists(pdf_file_zip_copy):
            os.remove(pdf_file_zip_copy)
        return -1

retval = update_pdf_properties(sys.argv[1], sys.argv[2], sys.argv[3])
#retval = update_pdf_properties("LastAccessedThreshold","True","C:\\Temp\\Labelling\\TestPDF1.pdf")
print(retval)
