QuillToDocx
QuillToDocx converts legacy QL and PC Quill `.doc` files into `.docx`, `.pdf`, `.html`, and `.md`.

Usage
QuillToDocx basename
QuillToDocx input.doc [output.docx|output.pdf|output.html|output.md]

Examples:

QuillToDocx libcport
QuillToDocx libcport.doc libcport.md
QuillToDocx sample.doc sample.html

Formats
.docx: Word document via Open XML
.pdf: lightweight native PDF export
.html: standalone HTML with embedded CSS
.md: Markdown with inline HTML for styles Markdown cannot represent directly
Structure
QuillParser.fs reads the binary Quill file structures
QuillDecode.fs converts raw records into decoded paragraphs and inline runs
DocxWriter.fs, PdfWriter.fs, HtmlWriter.fs, and MarkdownWriter.fs render the decoded document
Program.fs resolves CLI arguments and dispatches to the selected writer
