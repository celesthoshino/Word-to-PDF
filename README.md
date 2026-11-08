# Word to PDF

Converts Word documents to PDF. Put your documents in the `input` folder, run the script, and get the PDFs in the
`output` folder.

The conversion is done by **Microsoft Word** itself (or the free **LibreOffice**), so the PDFs look exactly like the
documents do when you open them.

## Folder layout

```
word_to_pdf/
├── word_to_pdf.py     the script
├── requirements.txt   Python packages to install
├── README.md          this file
├── input/             put the documents you want converted here
└── output/            PDFs appear here, with the same names
```

## Setup (one time)

1. Install **Python 3.9 or newer** from https://www.python.org/downloads/
   (on Windows, tick "Add Python to PATH" during install).
2. You need one of these installed:
   - **Microsoft Word** (Windows), or
   - **LibreOffice**, free from https://www.libreoffice.org/download/ (Windows, Mac or Linux)
3. Open a terminal in the `word_to_pdf` folder and run:

   ```
   pip install -r requirements.txt
   ```

   This installs pywin32, which lets Python control Word. On Mac and Linux it installs nothing, since LibreOffice
   is used there.

## How to use

1. Copy the documents you want converted into the `input` folder.
2. Open a terminal in the `word_to_pdf` folder and run:

   ```
   python word_to_pdf.py
   ```

3. The PDFs are saved in the `output` folder. For example, `input/Report.docx` becomes `output/Report.pdf`.

| Command | What it does |
|---|---|
| `python word_to_pdf.py` | Converts every document, using Word if installed, otherwise LibreOffice |
| `python word_to_pdf.py --engine libreoffice` | Uses LibreOffice even if Word is installed |
| `python word_to_pdf.py --engine word` | Only uses Word |
| `python word_to_pdf.py -i "D:\Docs" -o "D:\PDFs"` | Uses different input and output folders |
| `python word_to_pdf.py --help` | Shows all options |

## Options

| Option | Default | Meaning |
|---|---|---|
| `--engine` | `auto` | `auto` tries Word first, then LibreOffice. `word` or `libreoffice` uses only that one |
| `-i`, `--input` | `input` folder | Folder to read documents from |
| `-o`, `--output` | `output` folder | Folder to save PDFs to |

**Documents it accepts:** `.docx .doc .docm .dotx .dot .rtf .odt`.
Other files are listed as "ignored" and left alone.

## Good to know

- **Original documents are never changed.** They're opened read-only.
- **Your open Word documents are safe.** The script uses its own hidden copy of Word, separate from any Word window
  you have open, and closes it when it's done.
- **Running it again overwrites earlier PDFs** with the same name in `output`.
- **Password-protected documents are skipped** with a message, instead of the script stopping to wait for a password.
- **Documents with the same name but different extensions**, like `Report.docx` and `Report.doc`, keep their old
  extension in the new name so neither overwrites the other: `Report_docx.pdf` and `Report_doc.pdf`.
- **Word's temporary `~$` files** (made while a document is open in Word) are skipped automatically.
- **Fonts:** LibreOffice may swap a font for a similar one if a document uses a font only Word has. Word doesn't have
  this problem.
- **Subfolders inside `input` are not scanned.** Put the documents directly in `input`.

## Troubleshooting

**`couldn't find a program to convert the documents with`**
Neither Word nor LibreOffice was found. The lines below the error say what's missing. Install LibreOffice (free), or run
`pip install -r requirements.txt` if it says pywin32 is missing.

**`The password is incorrect. Word cannot open the document.`**
The document has a password. Open it in Word, remove the password (File → Info → Protect Document), save it, and try again.

**`The file appears to be corrupted.`**
The document is damaged. Try opening it in Word. Word can often repair it, and you can then save a fixed copy.

**It seems stuck on one document**
Very large documents can take a minute. If Word is truly stuck, press Ctrl+C to stop. Then open Task Manager, end any
leftover "Microsoft Word" task, and try again.

**`'python' is not recognized`**
Python isn't on your PATH. On Windows, try `py word_to_pdf.py` instead, or reinstall Python with "Add Python to PATH" ticked.
