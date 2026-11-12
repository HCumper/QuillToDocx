QuillToDocx
QuillToDocx converts legacy QL and PC Quill `.doc` files into `.docx`, `.rtf`, `.pdf`, `.html`, and `.md`.
A Word import converter (WordConverter/) lets Microsoft Word open Quill files directly.

Usage
QuillToDocx basename
QuillToDocx input.doc [output.docx|output.rtf|output.pdf|output.html|output.md]

Examples:

QuillToDocx libcport
QuillToDocx libcport.doc libcport.md
QuillToDocx sample.doc sample.html
QuillToDocx report_doc report.docx

Formats
.docx: Word document via Open XML
.rtf: Rich Text Format with the same layout as .docx (used by the Word converter)
.pdf: lightweight native PDF export
.html: standalone HTML with embedded CSS
.md: Markdown with inline HTML for styles Markdown cannot represent directly

Opening Quill files in Word
WordConverter/ holds a Word text converter (QuillCnv.cnv) for 64-bit Word. Once installed,
double-clicking a Quill `.doc` opens it in Word, converted. Word detects Quill files by their
contents, so real Word `.doc` files are unaffected. The converter does not write Quill: Save
on a converted document opens Save As, so the Quill original is not overwritten silently.

Install it with QuillToDocx-<version>-x64.msi: close Word, run the MSI and accept the
administrator prompt (Word only reads converters from HKLM). It installs to
Program Files\QuillToDocx and is removed from Apps & Features. It needs 64-bit Word.

Building: run WordConverter\build.cmd (needs Visual Studio with the C++ x64 tools and the .NET
SDK). It builds the converter, a self-contained QuillToDocx.exe and the MSI, which lands in
Installer\bin\x64\Release. The version comes from Directory.Build.props. For development without
the MSI, WordConverter\Install-QuillConverter.ps1 and Uninstall-QuillConverter.ps1 (elevated)
install from WordConverter\bin instead.

License: MIT (see LICENSE); bundled third-party components are listed in THIRD-PARTY-NOTICES.txt.

QL-style names such as `report_doc` have no extension, so Windows can't send them to Word on
double-click, and Word only uses the converter for them when it is asked for explicitly
(Documents.Open with the converter's OpenFormat); otherwise Word opens them as plain text.
Convert them with the command line instead, or rename them to `.doc`.

To trace what Word asks the converter, create an empty `%TEMP%\QuillCnv.log`.
WordConverter\bin\CnvTest.exe drives the converter the way Word does, without Word:
CnvTest QuillCnv.cnv input.doc output.rtf

Structure
QuillParser.fs reads the binary Quill file structures
QuillDecode.fs converts raw records into decoded paragraphs and inline runs
QuillLayout.fs holds the page and paragraph geometry shared by the DOCX and RTF writers
DocxWriter.fs, RtfWriter.fs, PdfWriter.fs, HtmlWriter.fs, and MarkdownWriter.fs render the decoded document
Program.fs resolves CLI arguments and dispatches to the selected writer
WordConverter/QuillCnv.cpp is the Word converter; it runs QuillToDocx.exe to produce RTF
Installer/ is the WiX project for the MSI
