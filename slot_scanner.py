import os
from utils import DOCX_SIG

CHUNK_SIZE = 2 * 1024 * 1024  # 2MB 
OVERLAP_SIZE = 64            # Overlap to catch split signatures
OUTPUT_DIR = "fragments"


def scan_and_extract_fragments(dump_path, sig=DOCX_SIG, output_dir=OUTPUT_DIR):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    total_slots = 0
    valid_count = 0
    buffer = b""

    with open(dump_path, "rb") as f:
        while True:
            chunk = f.read(CHUNK_SIZE)
            if not chunk:
                break

            total_slots += 1
            buffer = buffer[-OVERLAP_SIZE:] + chunk
            pos = buffer.find(sig)
            if pos != -1:
                fragment = buffer[pos:pos + CHUNK_SIZE]
                if b"word/document.xml" in fragment:
                    out_path = os.path.join(output_dir, f"fragment_{valid_count + 1}.bin")
                    with open(out_path, "wb") as out:
                        out.write(fragment)
                    valid_count += 1

    return valid_count, total_slots
