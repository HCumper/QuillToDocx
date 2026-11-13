QuillToDocx
QuillToDocx converts legacy QL and PC Quill `.doc` files into `.docx`, `.rtf`, `.pdf`, `.html`, and `.md`,
and writes Quill documents from RTF.
A Word text converter (WordConverter/) lets Microsoft Word open Quill files directly and save them as Quill again.

Usage
QuillToDocx basename
QuillToDocx input.doc [output.docx|output.rtf|output.pdf|output.html|output.md]
QuillToDocx input.rtf output.doc
QuillToDocx input.rtf output_doc --quill

Examples:

QuillToDocx libcport
QuillToDocx libcport.doc libcport.md
QuillToDocx sample.doc sample.html
QuillToDocx report_doc report.docx
QuillToDocx report.rtf report_doc --quill

Formats
.docx: Word document via Open XML
.rtf: Rich Text Format with the same layout as .docx (used by the Word converter)
.pdf: lightweight native PDF export
.html: standalone HTML with embedded CSS
.md: Markdown with inline HTML for styles Markdown cannot represent directly
.doc (or --quill): a QL Quill document, from RTF or from another Quill document

Writing Quill
The input is RTF, such as Word saves, or a Quill document. What Quill can hold comes across: text in
the QL character set, bold, underline, subscript and superscript, tabs and soft hyphens, page breaks,
left/indent/right margins and tab stops in character columns (120 twips each), left, centred and
justified paragraphs, a one-line header and footer (a PAGE field becomes Quill's nnn), and the page
length and margins. Quill has no fonts, sizes, italic, right alignment, tables, pictures, lists or
styles: tables become tab-separated lines, list numbers and bullets become text, italic and right
alignment are dropped. Characters outside the QL set become their nearest ASCII form (smart quotes,
dashes) or "?".

A Quill file opened through the converter carries its layout through Word in a document variable
(QuillLayout): the tab tables, display settings, and a header or footer that is switched off come
back unchanged when it is saved as Quill again. Rewriting the 1,497 test documents directly gives the
same structure for 1,417; through Word's RTF, the rest differ only where Quill shows no difference
(which of two identical tab tables a paragraph uses, the unused indent of a centred paragraph).

The first time a Quill document is replaced, the original is kept beside it as <name>.bak.
Documents written this way, including one edited and saved in Word, load in real Quill (tested
under Q-emuLator). Quill itself only opens names of up to 8 characters before `_doc`, counting any
folders below the drive (`mdv1_skatru_doc` loads; `mdv1_skat_skatru_doc` gives "Not a valid QUILL
file", even for Quill's own files), so keep names short or attach the drive to the document's folder.

Opening and saving Quill files in Word
WordConverter/ holds a Word text converter (QuillCnv.cnv) for 64-bit Word. Once installed,
double-clicking a Quill `.doc` opens it in Word, converted. Word detects Quill files by their
contents, so real Word `.doc` files are unaffected. "Sinclair QL Quill" appears under Save As, and
saving a document that was opened from a Quill file writes Quill again.

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
CnvTest QuillCnv.cnv --export input.rtf output.doc

Structure
QuillParser.fs reads the binary Quill file structures
QuillCharset.fs maps the QL character set to and from Unicode
QuillDecode.fs converts raw records into decoded paragraphs and inline runs
QuillLayout.fs holds the page and paragraph geometry shared by the DOCX and RTF writers
QuillWriter.fs writes Quill documents, and carries a Quill layout through Word (QuillOriginal)
DocxWriter.fs, RtfWriter.fs, PdfWriter.fs, HtmlWriter.fs, and MarkdownWriter.fs render the decoded document
RtfReader.fs reads RTF; QuillFromRtf.fs turns it into a Quill document
Program.fs resolves CLI arguments and dispatches to the selected writer
WordConverter/QuillCnv.cpp is the Word converter; it runs QuillToDocx.exe to convert in both directions
Installer/ is the WiX project for the MSI
