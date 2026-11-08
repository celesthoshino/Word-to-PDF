"""
Word to PDF
-----------
Put Word documents in the "input" folder, run this script, and PDF versions
show up in the "output" folder. The original documents are never changed.

    python word_to_pdf.py                         # convert everything
    python word_to_pdf.py --engine libreoffice    # use LibreOffice instead of Word

Needs Microsoft Word (Windows) or LibreOffice (free, any system) installed.
See README.md for all the details.
"""

import argparse
import shutil
import subprocess
import sys
import tempfile
from collections import Counter
from pathlib import Path

# The input/output folders live next to this script, so it works no matter
# which folder you run it from.
BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"

INPUT_EXTENSIONS = {".docx", ".doc", ".docm", ".dotx", ".dot", ".rtf", ".odt"}

LIBREOFFICE_LOCATIONS = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/usr/bin/soffice",
    "/usr/bin/libreoffice",
]

# Word asks for a password when opening a protected document, which would freeze
# the script. Giving it a made-up password makes it fail straight away instead.
# Documents without a password ignore it.
FAKE_PASSWORD = "no-password-was-given"


class ConversionError(Exception):
    """A document that couldn't be converted, with a message that explains why."""


class WordEngine:
    """Converts using Microsoft Word (Windows only), via a hidden copy of Word."""
    name = "Microsoft Word"

    def __init__(self):
        import pythoncom
        import pywintypes
        import win32com.client
        self.pythoncom = pythoncom
        self.com_error = pywintypes.com_error
        pythoncom.CoInitialize()
        try:
            # DispatchEx starts a separate copy of Word, so documents you have open are not touched.
            self.word = win32com.client.DispatchEx("Word.Application")
        except Exception:
            pythoncom.CoUninitialize()
            raise
        self.word.Visible = False
        self.word.DisplayAlerts = 0  # never show pop-ups that would wait for a click

    def convert(self, src: Path, out_path: Path):
        try:
            doc = self.word.Documents.Open(
                FileName=str(src), ConfirmConversions=False, ReadOnly=True,
                AddToRecentFiles=False, PasswordDocument=FAKE_PASSWORD, Visible=False,
            )
        except self.com_error as error:
            raise ConversionError(self.explain(error)) from None
        try:
            doc.ExportAsFixedFormat(OutputFileName=str(out_path), ExportFormat=17)  # 17 = PDF
        except self.com_error as error:
            raise ConversionError(self.explain(error)) from None
        finally:
            doc.Close(SaveChanges=0)

    @staticmethod
    def explain(error) -> str:
        """Pull Word's own message out of the error, e.g. 'The password is incorrect.'"""
        details = error.args[2] if len(error.args) > 2 else None
        message = details[2] if details and len(details) > 2 and details[2] else error.args[1]
        return " ".join(str(message).split())

    def close(self):
        try:
            self.word.Quit(SaveChanges=0)
        except Exception:
            pass
        self.pythoncom.CoUninitialize()


class LibreOfficeEngine:
    """Converts using LibreOffice (free, works on Windows, Mac and Linux)."""
    name = "LibreOffice"

    def __init__(self):
        self.program = find_libreoffice()
        if not self.program:
            raise FileNotFoundError("LibreOffice not found")
        self.temp = tempfile.TemporaryDirectory()
        # A private settings folder, so this works even while LibreOffice is open.
        self.profile = Path(self.temp.name, "profile").as_uri()

    def convert(self, src: Path, out_path: Path):
        work_dir = Path(self.temp.name, "out")
        work_dir.mkdir(exist_ok=True)
        command = [self.program, f"-env:UserInstallation={self.profile}", "--headless",
                   "--convert-to", "pdf", "--outdir", str(work_dir), str(src)]
        try:
            result = subprocess.run(command, capture_output=True, encoding="utf-8", errors="replace", timeout=300)
        except subprocess.TimeoutExpired:
            raise ConversionError("LibreOffice took longer than 5 minutes and was stopped") from None
        produced = work_dir / (src.stem + ".pdf")
        if not produced.exists():
            message = (result.stderr or result.stdout).strip() or "LibreOffice could not convert it"
            raise ConversionError(" ".join(message.split()))
        shutil.move(str(produced), str(out_path))

    def close(self):
        self.temp.cleanup()


def find_libreoffice():
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    for location in LIBREOFFICE_LOCATIONS:
        if Path(location).exists():
            return location
    return None


def start_engine(choice: str):
    """Start Word or LibreOffice, depending on what's installed and what was asked for."""
    problems = []
    if choice in ("auto", "word"):
        if sys.platform != "win32":
            problems.append("Microsoft Word can only be used on Windows")
        else:
            try:
                return WordEngine()
            except ImportError:
                problems.append("pywin32 is not installed. Run:  pip install -r requirements.txt")
            except Exception:
                problems.append("Microsoft Word is not installed, or it could not be started")
    if choice in ("auto", "libreoffice"):
        try:
            return LibreOfficeEngine()
        except FileNotFoundError:
            problems.append("LibreOffice is not installed (free download: https://www.libreoffice.org/download/)")
    sys.exit("Error: couldn't find a program to convert the documents with.\n"
             + "\n".join(f"  - {problem}" for problem in problems))


def find_input_files(input_dir: Path) -> list[Path]:
    """Return the documents in input_dir. Other files are listed as ignored."""
    documents = []
    for path in sorted(input_dir.iterdir(), key=lambda p: p.name.lower()):
        if not path.is_file() or path.name.startswith((".", "~$")):
            continue  # "~$" files are temporary files Word makes while a document is open
        if path.suffix.lower() in INPUT_EXTENSIONS:
            documents.append(path)
        else:
            print(f"  [ignored] {path.name} (not a supported document)")
    return documents


def main():
    parser = argparse.ArgumentParser(description="Convert every Word document in the input folder to PDF.")
    parser.add_argument("--engine", choices=["auto", "word", "libreoffice"], default="auto",
                        help="program to convert with (default: auto, which tries Word first, then LibreOffice)")
    parser.add_argument("-i", "--input", type=Path, default=INPUT_DIR,
                        help="folder to read documents from (default: the input folder)")
    parser.add_argument("-o", "--output", type=Path, default=OUTPUT_DIR,
                        help="folder to save PDFs to (default: the output folder)")
    args = parser.parse_args()

    args.input.mkdir(parents=True, exist_ok=True)
    args.output.mkdir(parents=True, exist_ok=True)

    print(f"Looking in: {args.input}")
    files = find_input_files(args.input)
    if not files:
        print("No documents found. Put the documents you want converted in that folder, then run this again.")
        return
    print(f"Found {len(files)} document(s).")

    # "report.docx" and "report.doc" would both become "report.pdf",
    # so those keep their old extension in the name: "report_docx.pdf", "report_doc.pdf".
    stem_counts = Counter(f.stem.lower() for f in files)
    duplicate_stems = {stem for stem, count in stem_counts.items() if count > 1}

    engine = start_engine(args.engine)
    print(f"Converting with {engine.name}...\n")
    converted = failed = 0
    try:
        for i, src in enumerate(files, start=1):
            stem = f"{src.stem}_{src.suffix.lstrip('.').lower()}" if src.stem.lower() in duplicate_stems else src.stem
            out_path = (args.output / (stem + ".pdf")).resolve()
            try:
                engine.convert(src.resolve(), out_path)
            except ConversionError as error:
                print(f"  [fail] [{i}/{len(files)}] {src.name}: {error}")
                failed += 1
                continue
            size_kb = out_path.stat().st_size / 1024
            print(f"  [ok]   [{i}/{len(files)}] {src.name} -> {out_path.name}  ({size_kb:,.0f} KB)")
            converted += 1
    finally:
        engine.close()  # always close Word/LibreOffice, even if something went wrong

    print(f"\nDone: {converted} converted, {failed} failed.")
    print(f"PDFs are in: {args.output}")
    if failed:
        sys.exit(1)


if __name__ == "__main__":
    try:
        main()
    except (KeyboardInterrupt, EOFError):
        print("\nCancelled.")
        sys.exit(1)
