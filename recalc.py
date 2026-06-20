"""Recalculate the workbook with LibreOffice so that every formula cell carries a
cached computed value.

openpyxl writes formulas but no cached results, so spreadsheet viewers that do not
themselves recalculate (GitHub preview, Quick Look, Google Sheets import, pandas
read with data_only, etc.) show blanks. Running this step opens the workbook in
LibreOffice with "always recalculate on load" forced on, then re-saves it — keeping
the live formulas *and* embedding the computed values.

Requires LibreOffice (`libreoffice`/`soffice`) on PATH.

Run:  python recalc.py [workbook.xlsx]
"""
import os
import shutil
import subprocess
import sys
import tempfile

XL = sys.argv[1] if len(sys.argv) > 1 else "SpaceX_Cursor_Pro_Forma_Model.xlsx"

# LibreOffice user profile that forces recalculation of OOXML/ODF files on load.
RECALC_PROFILE = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
 <item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
 <item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="ODFRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>
"""


def main():
    soffice = shutil.which("libreoffice") or shutil.which("soffice")
    if not soffice:
        sys.exit("LibreOffice not found on PATH; cannot cache values.")

    with tempfile.TemporaryDirectory() as prof, tempfile.TemporaryDirectory() as out:
        os.makedirs(os.path.join(prof, "user"), exist_ok=True)
        with open(os.path.join(prof, "user", "registrymodifications.xcu"), "w") as f:
            f.write(RECALC_PROFILE)
        cmd = [
            soffice, "--headless", "--norestore", "--invisible",
            f"-env:UserInstallation=file://{prof}",
            "--convert-to", "xlsx:Calc MS Excel 2007 XML",
            "--outdir", out, os.path.abspath(XL),
        ]
        subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        produced = os.path.join(out, os.path.basename(XL))
        if not os.path.exists(produced):
            sys.exit("LibreOffice did not produce an output file.")
        shutil.move(produced, XL)
    print(f"Recalculated and cached values into {XL}")


if __name__ == "__main__":
    main()
