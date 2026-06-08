import os
import sys
import zipfile
import shutil
import re

from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject


def is_signed(pdf_file_path):
    """Check if PDF has a digital signature."""
    reader = PdfReader(pdf_file_path)
    root = reader.trailer['/Root']
    l_signed = False
    if acroform := root.get('/AcroForm'):
        if acroform and (sig := acroform.get('/SigFlags')):
            l_signed = bool(sig & 1)
    return l_signed


def is_encrypted(pdf_file_path):
    """Check if PDF is encrypted."""
    reader = PdfReader(pdf_file_path)
    return reader.is_encrypted


def update_pdf_properties(pdf_file, properties: dict):
    """
    Write one or more custom properties to a PDF in a single read/write pass.

    parameters
    ----------
    pdf_file   : str   full path to the PDF file
    properties : dict  {property_name: property_value, ...}

    return codes
    ------------
     0  success
     1  encrypted
     2  digitally signed
    -1  unexpected error
    """
    backup_root = "c:\\temp\\unstructured\\backup\\2"

    if not os.path.exists(pdf_file):
        sys.exit(1)

    if is_encrypted(pdf_file):
        return 1

    if is_signed(pdf_file):
        return 2

    # Derive paths — all derived once, used consistently throughout
    if (pdf_file.upper()).startswith("C:"):
        pdf_file_copy = backup_root + "\\" + pdf_file.replace(":", "")
    elif pdf_file.startswith("\\\\"):
        pdf_file_copy = backup_root + pdf_file.replace("\\\\", "\\")
    else:
        pdf_file_copy = backup_root + "\\" + pdf_file.replace(":", "")

    pdf_file_zip_copy = re.sub(r"\.pdf$", ".zip", pdf_file_copy, flags=re.IGNORECASE)

    # Use a process-ID-tagged temp name to avoid collisions between concurrent
    # runspaces operating in the same directory
    pid = os.getpid()
    pdf_file_output = re.sub(r"\.pdf$", f"_updated_{pid}.pdf", pdf_file, flags=re.IGNORECASE)
    pdf_file_zip    = re.sub(r"\.pdf$", f"_backup_{pid}.zip",  pdf_file, flags=re.IGNORECASE)

    # Back up original PDF as ZIP (only if backup doesn't already exist)
    if not os.path.isfile(pdf_file_zip_copy):
        try:
            with zipfile.ZipFile(pdf_file_zip, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.write(pdf_file, arcname=os.path.basename(pdf_file))

            os.makedirs(os.path.dirname(pdf_file_zip_copy), exist_ok=True)
            shutil.move(pdf_file_zip, pdf_file_zip_copy)
        except Exception as e:
            # Backup failure is non-fatal — clean up and continue
            print(f"[Warning] Backup failed, continuing without backup: {e}", flush=True)
            if os.path.exists(pdf_file_zip):
                os.remove(pdf_file_zip)

    try:
        reader = PdfReader(pdf_file)
        writer = PdfWriter()

        # Copy all pages
        for page in reader.pages:
            writer.add_page(page)

        # Copy existing metadata, then apply all new properties in one pass
        existing_metadata = reader.metadata or {}
        new_metadata = {NameObject(k): str(v) for k, v in existing_metadata.items() if v is not None}

        for prop_name, prop_value in properties.items():
            # Ensure property name has leading slash as required by PDF spec
            key = prop_name if prop_name.startswith("/") else f"/{prop_name}"
            new_metadata[NameObject(key)] = str(prop_value)

        writer.add_metadata(new_metadata)

        # Write to PID-tagged temp file — safe for concurrent runspaces
        with open(pdf_file_output, "wb") as f_out:
            writer.write(f_out)

        # Atomic swap: remove original then rename temp into place
        os.remove(pdf_file)
        os.rename(pdf_file_output, pdf_file)

        # Remove backup now that write succeeded
        if os.path.exists(pdf_file_zip_copy):
            os.remove(pdf_file_zip_copy)

        return 0

    except Exception as e:
        print(f"[Error] {e}", flush=True)
        # Clean up temp files on failure; leave backup in place for recovery
        if os.path.exists(pdf_file_output):
            os.remove(pdf_file_output)
        return -1


def parse_args(args):
    """
    Expected call:
        python update_pdf_properties.py <pdf_file> <Name=Value> [<Name=Value> ...]

    Example:
        python update_pdf_properties.py "C:\\docs\\file.pdf" \
            "OriginalPath=C:\\docs\\file.pdf" \
            "LastAccessed18Months=True" \
            "Created3Years=False"
    """
    if len(args) < 3:
        print("Usage: update_pdf_properties.py <pdf_file> <Name=Value> [<Name=Value> ...]")
        sys.exit(1)

    pdf_file   = args[1]
    properties = {}

    for arg in args[2:]:
        if "=" not in arg:
            print(f"[Error] Property argument must be in Name=Value format, got: {arg}")
            sys.exit(1)
        name, _, value = arg.partition("=")
        properties[name.strip()] = value.strip()

    return pdf_file, properties


if __name__ == "__main__":
    pdf_file, properties = parse_args(sys.argv)
    retval = update_pdf_properties(pdf_file, properties)
    print(retval, flush=True)
