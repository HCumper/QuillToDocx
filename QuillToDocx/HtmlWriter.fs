namespace QuillToDocx

open System
open System.IO
open System.Text

module HtmlWriter =

    [<Literal>]
    let private DefaultBodyWidthCols = 80

    let private lineSpacingPercent (layout: LayoutTable) =
        max 110 (min 220 (100 + (int layout.LineGap * 10)))

    let private paragraphAfterEm (layout: LayoutTable) =
        float (max 0 (min 24 (int layout.LineGap + 2))) / 10.0

    let private pageWidthPx =
        900

    let private htmlEncode (text: string) =
        text
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;")

    let private cssClassForStyle (style: TextStyle) =
        let parts = ResizeArray<string>()
        if style.Bold then parts.Add("bold")
        if style.Underline then parts.Add("underline")
        if style.Sub then parts.Add("sub")
        if style.Super then parts.Add("super")
        String.Join(" ", parts)

    let private appendStyledSpan (builder: StringBuilder) (text: string) (style: TextStyle) =
        let encoded = htmlEncode text
        let classes = cssClassForStyle style

        if String.IsNullOrEmpty(classes) then
            builder.Append(encoded) |> ignore
        else
            builder.AppendFormat("<span class=\"{0}\">{1}</span>", classes, encoded) |> ignore

    let private renderParagraphContent (content: Inline list) =
        let builder = StringBuilder()

        for item in content do
            match item with
            | RunText (text, style) -> appendStyledSpan builder text style
            | RunTab _ -> builder.Append("<span class=\"tab\"></span>") |> ignore
            | RunSoftHyphen _ -> builder.Append("&shy;") |> ignore

        builder.ToString()

    let private textAlignment justification =
        match int justification with
        | 1
        | 5 -> "center"
        | 2
        | 6 -> "right"
        | _ -> "left"

    let private rightColumns (p: ParaTable) =
        if int p.RightMarg > int p.LeftMarg then int p.RightMarg else DefaultBodyWidthCols

    let private paragraphStyle (layout: LayoutTable) (p: ParaTable) =
        let leftCols = int p.LeftMarg
        let firstCols = max leftCols (int p.IndentMarg)
        let rightCols = rightColumns p
        let widthCols = max 20 (rightCols - leftCols)
        let leftPercent = (float leftCols / float DefaultBodyWidthCols) * 100.0
        let widthPercent = (float widthCols / float DefaultBodyWidthCols) * 100.0
        let indentEm = float (max 0 (firstCols - leftCols))
        let lineHeight = lineSpacingPercent layout
        let marginBottom = paragraphAfterEm layout

        sprintf
            "margin-left:%.2f%%;width:%.2f%%;text-align:%s;text-indent:%.2fch;line-height:%d%%;margin-top:0;margin-bottom:%.2fem;"
            leftPercent
            widthPercent
            (textAlignment p.Justif)
            indentEm
            lineHeight
            marginBottom

    let private renderParagraph (layout: LayoutTable) (paragraph: DecodedParagraph) =
        sprintf
            "<p class=\"quill-paragraph\" style=\"%s\">%s</p>"
            (paragraphStyle layout paragraph.Descriptor)
            (renderParagraphContent paragraph.Content)

    let private renderOptionalSection className title layout paragraph =
        match paragraph with
        | None -> ""
        | Some p ->
            sprintf
                "<section class=\"%s\"><h2>%s</h2>%s</section>"
                className
                title
                (renderParagraph layout p)

    let private htmlDocument title (qf: QuillFile) (doc: DecodedDocument) =
        let bodyParagraphs =
            doc.BodyParagraphs
            |> List.map (renderParagraph qf.Layout)
            |> String.concat "\n"

        let headerSection = renderOptionalSection "document-header" "Header" qf.Layout doc.HeaderParagraph
        let footerSection = renderOptionalSection "document-footer" "Footer" qf.Layout doc.FooterParagraph

        sprintf
            """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>%s</title>
  <style>
    :root {
      --page-width: %dpx;
      --page-bg: #f6f1e8;
      --paper-bg: #fffdf8;
      --ink: #1f1b16;
      --rule: #d9cfbf;
      --accent: #8b6f47;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      background:
        linear-gradient(180deg, #ede4d6 0%%, #f6f1e8 25%%, #efe6d8 100%%);
      color: var(--ink);
      font-family: Georgia, "Times New Roman", serif;
      padding: 32px 16px;
    }
    .page {
      width: min(100%%, var(--page-width));
      margin: 0 auto;
      background: var(--paper-bg);
      border: 1px solid var(--rule);
      box-shadow: 0 20px 45px rgba(52, 39, 21, 0.12);
      padding: 48px 56px;
    }
    .document-header,
    .document-footer {
      border-bottom: 1px solid var(--rule);
      margin-bottom: 24px;
      padding-bottom: 12px;
    }
    .document-footer {
      border-top: 1px solid var(--rule);
      border-bottom: none;
      margin-top: 24px;
      margin-bottom: 0;
      padding-top: 12px;
      padding-bottom: 0;
    }
    .document-header h2,
    .document-footer h2 {
      margin: 0 0 10px 0;
      font-size: 0.75rem;
      letter-spacing: 0.18em;
      text-transform: uppercase;
      color: var(--accent);
      font-family: Arial, Helvetica, sans-serif;
    }
    .quill-body {
      font-family: "Courier New", Courier, monospace;
      font-size: 0.95rem;
      white-space: normal;
    }
    .quill-paragraph {
      margin-right: 0;
    }
    .tab {
      display: inline-block;
      width: 4ch;
    }
    .bold { font-weight: 700; }
    .underline { text-decoration: underline; }
    .sub { vertical-align: sub; font-size: 0.8em; }
    .super { vertical-align: super; font-size: 0.8em; }
    @media (max-width: 720px) {
      body { padding: 12px; }
      .page { padding: 24px 18px; }
    }
  </style>
</head>
<body>
  <main class="page">
    %s
    <section class="quill-body">
      %s
    </section>
    %s
  </main>
</body>
</html>
"""
            (htmlEncode title)
            pageWidthPx
            headerSection
            bodyParagraphs
            footerSection

    let writeHtml (qf: QuillFile) (doc: DecodedDocument) (outPath: string) =
        let title =
            match Path.GetFileNameWithoutExtension(outPath) with
            | null
            | "" -> "Quill Document"
            | fileName -> fileName

        let html = htmlDocument title qf doc
        File.WriteAllText(outPath, html, Encoding.UTF8)
