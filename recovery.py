# RECOVERY MODULE — recovery.py

import os
from io import BytesIO
import zipfile
import olefile
from utils import DOCX_SIG, DOCX_END, DOC_SIG, MIN_DOCX_SIZE, BLOCK_SIZE

def recover_documents_from_dump(dump_path, output_dir="recovered_docs", max_files=20):
    # Ensure output directory exists
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Load entire binary content of the dump
    with open(dump_path, "rb") as f:
        data = f.read()

    offset = 0        # Current search position
    count = 0         # Number of recovered files
    recovered_files = []

    while count < max_files:
        # Find the next signature for .docx or .doc
        docx_pos = data.find(DOCX_SIG, offset)
        doc_pos = data.find(DOC_SIG, offset)

        # If no signatures are found, stop the loop
        if docx_pos == -1 and doc_pos == -1:
            break

        # If .docx is found earlier or only .docx exists
        if docx_pos != -1 and (doc_pos == -1 or docx_pos < doc_pos):
            start = docx_pos
            ext = ".docx"
            end = data.find(DOCX_END, start)

            # If end of .docx found, include full archive
            if end != -1:
                end += 22  # include ZIP EOCD structure
                chunk = data[start:end]
            else:
                chunk = data[start:start + BLOCK_SIZE]

            # Check minimum size and valid ZIP structure
            if len(chunk) < MIN_DOCX_SIZE or not zipfile.is_zipfile(BytesIO(chunk)):
                offset = start + 1
                continue

        else:
            # Detected .doc (OLE format)
            start = doc_pos
            ext = ".doc"
            chunk = data[start:start + BLOCK_SIZE]

            # Validate using olefile
            if not olefile.isOleFile(BytesIO(chunk)):
                offset = start + 1
                continue

        # Save recovered file
        file_path = os.path.join(output_dir, f"recovered_{count + 1}{ext}")
        with open(file_path, "wb") as out:
            out.write(chunk)

        recovered_files.append(file_path)
        count += 1
        offset = start + len(chunk)

    return recovered_files
