# Matrix PDF-Driver

A virtual printer for Windows that turns anything you print into a PDF file.
It can ask where to save each time, or save straight into a folder you choose.

[한국어 설명](README.ko.md)

- Publisher: Hihon Inc.
- Runs on: Windows 7 SP1 to Windows 11 (32-bit and 64-bit)
- Languages: English (default) and Korean
- License: GNU AGPL v3 (`LICENSE.txt`). Bundled third-party software is listed in `THIRD-PARTY.txt`.

## Features

- **Print to PDF from any program.** A printer named "Matrix PDF-Driver" is added; whatever you print to it becomes a PDF.
- **Two ways to save.** Ask where to save each time, or save automatically into a folder without any question.
- **File names built from rules.** Combine a fixed prefix with the document name, date, time, user, computer name or a running number.
- **Folders created for you.** Add the date or the user to the save folder and the subfolders are made as you print.
- **Archive formats.** Plain PDF, or PDF/A-1b, PDF/A-2b and PDF/A-3b for long-term storage.
- **Quality options.** Keep images as they are or reduce the file size; color or grayscale; optional automatic page rotation.
- **After printing.** Show a short progress window, open the PDF right away, or send it to a real printer at the same time.
- **Searchable text.** Text stays as text, including Korean, so it can be selected, searched and extracted.
- **English and Korean.** Switch the language in the settings; the setup can run in either.
- **No ads, no account, free.** Open source under the GNU AGPL v3.

## Download

[Matrix-PDF-Driver-Setup.exe](https://github.com/hihon-dev/Matrix-PDF-Driver/releases/latest/download/Matrix-PDF-Driver-Setup.exe) (about 30 MB) —
or see all versions on the [Releases](https://github.com/hihon-dev/Matrix-PDF-Driver/releases) page.

The installer is not code-signed yet, so Windows may show a warning before it runs.
Choose "More info", then "Run anyway".

## Install

1. Run `Matrix-PDF-Driver-Setup.exe` and allow it to make changes.
2. The setup starts in English. Pick another language on the first screen if you prefer.
3. Accept the license, check the install folder and click **Install**.
4. A printer named **Matrix PDF-Driver** is added. Leave "Open Setting now" ticked to adjust the settings right away.

To install without any window, for example on many PCs:

```
Matrix-PDF-Driver-Setup.exe /Q /C:"cmd.exe /d /c .\setup.cmd --mode auto --output C:\MatrixPDF"
```

| Option | Meaning |
|---|---|
| `--mode auto` or `--mode ask` | Save automatically, or ask every time. Default is `auto`. |
| `--output <folder>` | Save folder |
| `--pattern <rule>` | File name rule, for example `Cube_{datetime}` |
| `--dir <folder>` | Install folder |
| `--language en` or `--language ko` | Screen language |

Give at least one option. With no options the setup window opens instead.

## Use

Print from any program and choose **Matrix PDF-Driver** as the printer.

- **Ask every time**: a save dialog opens in the save folder a second or two later, with the file name already filled in.
- **Save automatically**: the PDF appears in the save folder without any question.

It keeps working after the PC restarts.

## Settings

Open **Matrix PDF-Driver Setting** from the Start menu. Changes apply from the next printout.

| Menu | Item | What it does |
|---|---|---|
| Save | Save mode | Ask every time, or save automatically |
| Save | Save folder | A folder on this PC. "Add subfolder" appends an item such as the date, and that folder is created when you print. Example: `C:\MatrixPDF\{date}` |
| File name | Prefix | The fixed start of the name. Default `Cube_` |
| File name | Name items | What follows the prefix. Default `{datetime}` |
| File name | Full path | Folder, prefix and name items on one line. Editing it updates the other fields. |
| Quality | Format | PDF, PDF/A-1b, PDF/A-2b (recommended for archiving), PDF/A-3b |
| Quality | Image quality | Keep original, or reduce file size (images are downsampled to 300 dpi and compressed) |
| Quality | Page rotation | As printed, or rotate to match the text direction |
| Quality | Color | As printed, or grayscale |
| Output | Show conversion progress | A small window with a progress bar appears briefly in the middle of the screen |
| Output | Open the PDF when it is ready | Opens each PDF in your PDF viewer |
| Output | Also print to | Sends each PDF to a real printer as well |
| About | Language | English or Korean |

Items you can use in the save folder and the file name:

| Item | Meaning | Example |
|---|---|---|
| `{title}` | Name of the printed document | Report |
| `{datetime}` | Date and time | 20261008_150405 |
| `{date}` `{time}` | Date, time | 20261008, 150405 |
| `{year}` `{month}` `{day}` | Year, month, day | 2026, 10, 08 |
| `{user}` | Who printed | hong |
| `{computer}` | Computer name | LAB-PC01 |
| `{printer}` | Printer name | Matrix PDF-Driver |
| `{counter}` | Running number. `{counter:04d}` pads it. | 12, 0012 |

If a file with the same name already exists, `_2`, `_3` is added.

## Remove

Remove **Matrix PDF-Driver** from Windows "Installed apps".
Your PDF files and settings are kept. Only the files this program installed are deleted from the install folder.

## Good to know

- **Text inside the PDF.** Some programs print without passing the text along, for example the new Notepad in Windows 11.
  Those PDFs look right but the text cannot be selected or extracted. Print once from your program and check that you can select text in the PDF.
- **PDF/A.** The files are marked as PDF/A-1b, 2b or 3b, but they have not been checked with a validator such as veraPDF.
  Validate your own output before relying on it for compliance. PDF/A-3b output contains no attachments, so its content is the same as 2b.
- **If the program is not running**, printouts wait in the Windows print queue and are converted when it starts again.

## Troubleshooting

| Symptom | What to check |
|---|---|
| No PDF appears | The top right of the Setting window should say "Running". If it says "Stopped", restart the PC or run the installer again. |
| The save dialog does not open | Sign out and sign in again. |
| Printed, but nothing was saved | About > "View failed jobs", then read the `.reason.txt` file. |
| Details | `C:\ProgramData\Matrix PDF-Driver\logs\driver.log` (written in Korean) |

## How it works

```
Print from any program
  → printer "Matrix PDF-Driver" (built-in Windows PostScript driver "MS Publisher Color Printer")
  → port 127.0.0.1:9100 (open only inside this PC; 9101-9119 if 9100 is taken)
  → MatrixPdfDriver.exe --run   (starts with Windows, runs as SYSTEM)
  → Ghostscript converts PostScript to PDF
  → save automatically: written as .tmp, then renamed to .pdf
  → ask every time: handed to MatrixPdfDriver.exe --agent, which shows the save dialog
```

The receiving program runs as SYSTEM and cannot show windows, so a small helper (`--agent`) runs for each signed-in user.
The helper shows the save dialog and the progress window, opens the PDF and prints to a real printer.

| Location | Contents |
|---|---|
| `C:\Program Files\Matrix PDF-Driver\` | Program, Ghostscript (`gs\`), licenses |
| `C:\ProgramData\Matrix PDF-Driver\config.json` | Settings |
| `C:\ProgramData\Matrix PDF-Driver\logs\` | Logs (5 MB × 5) |
| `C:\ProgramData\Matrix PDF-Driver\failed\` | Print data that could not be converted, with the reason |
| `C:\ProgramData\Matrix PDF-Driver\pending\` | PDFs waiting for the save dialog |

## Build

You need:

- Python 3.8, 32-bit. It is the last version that supports Windows 7.
  `winget install Python.Python.3.8 --architecture x86 --scope user`
- The files in `vendor\ghostscript\`: `gswin32c.exe`, `gsdll32.dll` and `vcredist_x86.exe` from the official
  32-bit Ghostscript 10.08.0 installer (`gs10080w32.exe`, extract it with 7-Zip). They are included in this repository.

```powershell
& "$env:LOCALAPPDATA\Programs\Python\Python38-32\python.exe" -m pip install -r requirements-build.txt
& "$env:LOCALAPPDATA\Programs\Python\Python38-32\python.exe" -m unittest test_matrix_pdf_driver
.\build.ps1
```

The result is `dist\Matrix-PDF-Driver-Setup.exe`.
Pass `-SigningCertificateThumbprint` to sign it if you have a code-signing certificate.
Pass `-PublishDirectory <folder>` to also copy the installer to another folder, for example the download folder of a web server.

On-screen text is written in Korean in the source and wrapped in `tr()`; English comes from the `TRANSLATIONS` table.
When you add text, do both. The tests fail if an English entry is missing.

## Tested

| Item | Status |
|---|---|
| Windows 11 64-bit: install, print, repeated prints, reinstall, remove | Checked |
| Korean text extraction, Korean file and folder names | Checked |
| Windows 7, Windows 10, 32-bit Windows | **Not checked yet** |
| Installing through the setup window (only the silent install was checked) | **Not checked yet** |
| Automatic start after a restart | **Not checked yet** |
| PDF/A conformance with a validator | **Not checked yet** |
