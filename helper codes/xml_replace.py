#!/usr/bin/env python3
# xml_replace.py (no CLI arguments)
# Automatically replaces placeholders across ALL XML parts inside a .docx
# including textboxes, shapes, drawings, headers, footers, etc.

import zipfile, shutil, os, json

# ----------------------------------------------------------
# CONFIG — SET YOUR OWN PATHS AND PLACEHOLDERS HERE
# ----------------------------------------------------------

INPUT_DOCX  = r"D:\Dixon\Automation_software\Office Software\template_formatted.docx"
OUTPUT_DOCX = r"D:\Dixon\Automation_software\Office Software\template_fixed.docx"

# Example: define placeholder → replacement text
PLACEHOLDER_MAP = {
    "(shift_1_signin)": "[SHIFT1_IMAGE]",
    "(shift_1_signout)": "[SHIFT1_OUT]",
    "(shift_2_signin)": "[SHIFT2_IN]",
    "(shift_2_signout)": "[SHIFT2_OUT]",
    "(date_with_month)": "24 November 2025",
    "(date)": "2025-11-24"
}

# ----------------------------------------------------------
# CORE FUNCTIONS
# ----------------------------------------------------------

def replace_in_bytes(data_bytes, mapping):
    text = data_bytes.decode("utf-8")
    for k, v in mapping.items():
        text = text.replace(k, v)
    return text.encode("utf-8")

def xml_replace_docx(input_path, output_path, mapping):
    tmpdir = input_path + "_unzipped"

    if os.path.exists(tmpdir):
        shutil.rmtree(tmpdir)
    os.makedirs(tmpdir)

    # Unzip
    with zipfile.ZipFile(input_path, "r") as zin:
        zin.extractall(tmpdir)

    # Replace in ALL XML files (document, header, footer, drawings…)
    for root, dirs, files in os.walk(tmpdir):
        for fname in files:
            if fname.endswith(".xml"):
                fpath = os.path.join(root, fname)
                with open(fpath, "rb") as f:
                    data = f.read()
                newdata = replace_in_bytes(data, mapping)
                with open(fpath, "wb") as f:
                    f.write(newdata)

    # Rezip
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for root, dirs, files in os.walk(tmpdir):
            for fname in files:
                fpath = os.path.join(root, fname)
                arcname = os.path.relpath(fpath, tmpdir)
                zout.write(fpath, arcname)

    shutil.rmtree(tmpdir)
    print(f"[OK] Saved replaced document → {output_path}")


# ----------------------------------------------------------
# MAIN EXECUTION
# ----------------------------------------------------------

if __name__ == "__main__":
    print("[INFO] Running XML placeholder replacement...")
    xml_replace_docx(INPUT_DOCX, OUTPUT_DOCX, PLACEHOLDER_MAP)
    print("[DONE]")
