import os
import zipfile
import olefile
import json
from xml.etree import ElementTree

# Extract plain text from a .docx file by parsing the document.xml file
def extract_docx_text(path):
    try:
        with zipfile.ZipFile(path, 'r') as zf:
            if "word/document.xml" in zf.namelist():
                xml_data = zf.read("word/document.xml")
                tree = ElementTree.fromstring(xml_data)
                paragraphs = tree.findall(
                    './/w:t', 
                    namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
                )
                return " ".join([p.text for p in paragraphs if p.text])
    except Exception as e:
        return f"[Error]: {e}"
    return ""

# Extract names of embedded images from a .docx file (if any)
def extract_docx_images(path):
    images = []
    try:
        with zipfile.ZipFile(path, 'r') as zf:
            for name in zf.namelist():
                if name.startswith("word/media/"):
                    images.append(name)
    except:
        pass
    return images

# Analyze the .docx file: structure, encryption, text content, images
def analyze_docx(path):
    result = {
        "type": "docx",
        "valid_zip": False,
        "has_document_xml": False,
        "encrypted": False,
        "extracted_text": "",
        "images": []
    }
    try:
        with zipfile.ZipFile(path, 'r') as zf:
            result["valid_zip"] = True
            names = zf.namelist()
            result["has_document_xml"] = "word/document.xml" in names
            result["encrypted"] = any("EncryptedPackage" in n for n in names)
            result["images"] = extract_docx_images(path)
            if result["has_document_xml"] and not result["encrypted"]:
                result["extracted_text"] = extract_docx_text(path)
    except Exception as e:
        result["error"] = str(e)
    return result

# Analyze a legacy .doc (OLE) file
def analyze_doc(path):
    result = {
        "type": "doc",
        "valid_ole": False,
        "has_word_stream": False,
        "encrypted": False
    }
    try:
        if olefile.isOleFile(path):
            result["valid_ole"] = True
            ole = olefile.OleFileIO(path)
            streams = ole.listdir()
            result["has_word_stream"] = any('WordDocument' in '/'.join(s) for s in streams)
            if ole.exists('EncryptionInfo') or ole.exists('EncryptedPackage'):
                result["encrypted"] = True
            ole.close()
    except Exception as e:
        result["error"] = str(e)
    return result

# Analyze a list of .doc/.docx files and generate a report
def analyze_files(file_paths):
    report = {}
    for path in file_paths:
        filename = os.path.basename(path)
        if path.endswith(".docx"):
            report[filename] = analyze_docx(path)
        elif path.endswith(".doc"):
            report[filename] = analyze_doc(path)
        else:
            report[filename] = {"error": "Unsupported format"}
    return report
