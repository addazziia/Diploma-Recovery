
import os
import zipfile
from io import BytesIO

DOCX_SIG = b'\x50\x4B\x03\x04'
MAX_FILES = 20


def is_valid_docx(chunk):
    """Check if a chunk contains a valid .docx by checking ZIP structure and presence of document.xml"""
    try:
        with zipfile.ZipFile(BytesIO(chunk)) as zf:
            return "word/document.xml" in zf.namelist()
    except:
        return False


def recover_docx_from_fragment(fragment_path, output_dir="recovered_docs_from_fragment"):
    """Scans fragment and extracts valid .docx files from it"""
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    with open(fragment_path, "rb") as f:
        data = f.read()

    offset = 0
    count = 0
    recovered_files = []

    while count < MAX_FILES:
        sig_offset = data.find(DOCX_SIG, offset)
        if sig_offset == -1:
            break

        chunk = data[sig_offset:sig_offset + 200000]  # ~200 KB
        if is_valid_docx(chunk):
            file_path = os.path.join(output_dir, f"recovered_{count + 1}.docx")
            with open(file_path, "wb") as out:
                out.write(chunk)
            recovered_files.append(file_path)
            count += 1
            offset = sig_offset + len(chunk)
        else:
            offset = sig_offset + 1

    return recovered_files
