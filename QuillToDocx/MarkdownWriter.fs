namespace QuillToDocx

open System
open System.IO
open System.Text

module MarkdownWriter =

    let private repeatText count (text: string) =
        if count <= 0 then "" else String.replicate count text

    let private escapeMarkdownText (text: string) =
        let builder = StringBuilder(text.Length)

        for ch in text do
            match ch with
            | '\\'
            | '`'
            | '*'
            | '_'
            | '{'
            | '}'
            | '['
            | ']'
            | '('
            | ')'
            | '#'
            | '+'
            | '-'
            | '.'
            | '!'
            | '|'
            | '>' ->
                builder.Append('\\').Append(ch) |> ignore
            | _ -> builder.Append(ch) |> ignore

        builder.ToString()

    let private applyHtmlStyleTags (style: TextStyle) (text: string) =
        let withUnderline =
            if style.Underline then sprintf "<u>%s</u>" text else text

        let withItalic =
            if style.Italic then sprintf "*%s*" withUnderline else withUnderline

        let withSub =
            if style.Sub then sprintf "<sub>%s</sub>" withItalic else withItalic

        if style.Super then sprintf "<sup>%s</sup>" withSub else withSub

    let private renderInline (item: Inline) =
        match item with
        | RunText (text, style) ->
            let escaped = escapeMarkdownText text
            let withBold =
                if style.Bold then sprintf "**%s**" escaped else escaped

            applyHtmlStyleTags style withBold
        | RunTab _ -> "    "
        | RunSoftHyphen _ -> "&shy;"

    let private renderParagraphContent (content: Inline list) =
        content
        |> List.map renderInline
        |> String.concat ""

    let private paragraphPrefix (p: ParaTable) =
        let leftIndent = repeatText (max 0 (int p.LeftMarg)) " "
        let firstIndentCount = max 0 (int p.IndentMarg - int p.LeftMarg)
        let firstIndent = repeatText firstIndentCount " "
        leftIndent + firstIndent

    let private renderParagraph (paragraph: DecodedParagraph) =
        paragraphPrefix paragraph.Descriptor + renderParagraphContent paragraph.Content

    let private renderSection heading paragraph =
        match paragraph with 
        | None -> None
        | Some value ->
            Some(
                String.concat
                    "\n\n"
                    [ sprintf "## %s" heading
                      renderParagraph value ]
            )

    let private markdownDocument title (doc: DecodedDocument) =
        let sections = ResizeArray<string>()
        sections.Add(sprintf "# %s" (escapeMarkdownText title))

        renderSection "Header" doc.HeaderParagraph |> Option.iter sections.Add
        if not doc.BodyParagraphs.IsEmpty then
            let body =
                doc.BodyParagraphs
                |> List.map renderParagraph
                |> String.concat "\n\n"

            sections.Add(body)

        renderSection "Footer" doc.FooterParagraph |> Option.iter sections.Add

        String.concat "\n\n" sections

    let writeMarkdown (_: QuillFile) (doc: DecodedDocument) (outPath: string) =
        let title =
            match Path.GetFileNameWithoutExtension(outPath) with
            | null
            | "" -> "Quill Document"
            | fileName -> fileName

        let markdown = markdownDocument title doc
        File.WriteAllText(outPath, markdown, Encoding.UTF8)
