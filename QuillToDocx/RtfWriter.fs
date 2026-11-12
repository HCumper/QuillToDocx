namespace QuillToDocx

open System
open System.IO
open System.Text

/// Writes RTF with the same layout decisions as DocxWriter. RTF is the format Word's
/// text-converter interface accepts, so this is what the Word import converter streams.
module RtfWriter =

    open QuillLayout

    [<Literal>]
    let private FontSizeHalfPoints = 20

    let private appendEscaped (sb: StringBuilder) (text: string) =
        for ch in text do
            match ch with
            | '\\' -> sb.Append(@"\\") |> ignore
            | '{' -> sb.Append(@"\{") |> ignore
            | '}' -> sb.Append(@"\}") |> ignore
            | c when c >= ' ' && int c < 0x7F -> sb.Append(c) |> ignore
            // \uN takes a signed 16-bit value; '?' is the fallback for readers without Unicode.
            | c -> sb.Append(@"\u").Append(int (int16 c)).Append('?') |> ignore

    let private styleControls (s: TextStyle) =
        String.concat
            ""
            [ if s.Bold then @"\b"
              if s.Italic then @"\i"
              if s.Underline then @"\ul"
              if s.Sub then @"\sub"
              if s.Super then @"\super" ]

    let private appendRun (sb: StringBuilder) item =
        let openGroup style =
            let controls = styleControls style
            sb.Append('{').Append(controls) |> ignore
            controls

        match item with
        | RunText (text, style) ->
            // A space after a control word is its delimiter, not text.
            if openGroup style <> "" then sb.Append(' ') |> ignore
            appendEscaped sb text
            sb.Append('}') |> ignore
        | RunTab style ->
            openGroup style |> ignore
            sb.Append(@"\tab}") |> ignore
        | RunSoftHyphen style ->
            openGroup style |> ignore
            sb.Append(@"\-}") |> ignore

    let private alignmentControl alignment =
        match alignment with
        | AlignLeft -> @"\ql"
        | AlignCenter -> @"\qc"
        | AlignRight -> @"\qr"

    let private appendParagraphStart (sb: StringBuilder) (layout: LayoutTable) alignment =
        sb.Append(@"\pard\plain").Append(alignmentControl alignment) |> ignore
        sb.AppendFormat(@"\sb0\sa{0}\sl{1}\slmult1", paragraphAfterTwips layout, lineSpacingTwips layout) |> ignore

    let private appendBodyParagraph (sb: StringBuilder) (qf: QuillFile) (p: DecodedParagraph) =
        appendParagraphStart sb qf.Layout (paragraphAlignment p.Descriptor.Justif)

        if p.PageBreakBefore then sb.Append(@"\pagebb") |> ignore

        let indents = paragraphIndents p.Descriptor
        sb.AppendFormat(@"\li{0}", indents.LeftIndent) |> ignore
        if indents.RightIndent > 0 then sb.AppendFormat(@"\ri{0}", indents.RightIndent) |> ignore
        if indents.FirstLineIndent > 0 then sb.AppendFormat(@"\fi{0}", indents.FirstLineIndent) |> ignore

        match tabStops qf p.Descriptor with
        | Some stops ->
            for position, kind in stops do
                match kind with
                | TabCenter -> sb.Append(@"\tqc") |> ignore
                | TabRight -> sb.Append(@"\tqr") |> ignore
                | TabLeft -> ()

                sb.AppendFormat(@"\tx{0}", position) |> ignore
        | None -> ()

        sb.AppendFormat(@"\f0\fs{0} ", FontSizeHalfPoints) |> ignore

        for item in p.Content do
            appendRun sb item

    let private appendHeaderFooter (sb: StringBuilder) (layout: LayoutTable) (destination: string) (para: DecodedParagraph) justification boldFlag =
        sb.Append('{').Append(destination) |> ignore
        appendParagraphStart sb layout (headerFooterAlignment justification)
        sb.AppendFormat(@"\f0\fs{0} ", FontSizeHalfPoints) |> ignore

        for item in para.Content do
            let item =
                if not boldFlag then
                    item
                else
                    match item with
                    | RunText (t, s) -> RunText(t, { s with Bold = true })
                    | RunTab s -> RunTab { s with Bold = true }
                    | RunSoftHyphen s -> RunSoftHyphen { s with Bold = true }

            appendRun sb item

        sb.Append(@"\par}").Append("\r\n") |> ignore

    let buildRtf (qf: QuillFile) (doc: DecodedDocument) =
        let sb = StringBuilder()
        let page = pageGeometry qf.Layout

        sb.Append(@"{\rtf1\ansi\ansicpg1252\deff0\uc1").Append("\r\n") |> ignore
        sb.Append(@"{\fonttbl{\f0\fmodern\fcharset0 Courier New;}}").Append("\r\n") |> ignore

        sb.AppendFormat(
            @"\paperw{0}\paperh{1}\margl{2}\margr{3}\margt{4}\margb{5}",
            page.Width,
            page.Height,
            page.LeftMargin,
            page.RightMargin,
            page.TopMargin,
            page.BottomMargin
        )
        |> ignore

        sb.AppendFormat(@"\sectd\headery{0}\footery{1}", page.HeaderMargin, page.FooterMargin) |> ignore

        if qf.Layout.StartPage > 1uy then
            sb.AppendFormat(@"\pgnstarts{0}\pgnrestart", int qf.Layout.StartPage) |> ignore

        sb.Append("\r\n") |> ignore

        match doc.HeaderParagraph with
        | Some hp when qf.Layout.HeaderF <> 0uy ->
            appendHeaderFooter sb qf.Layout @"\header" hp qf.Layout.HeaderF (qf.Layout.HeaderBold <> 0uy)
        | _ -> ()

        match doc.FooterParagraph with
        | Some fp when qf.Layout.FooterF <> 0uy ->
            appendHeaderFooter sb qf.Layout @"\footer" fp qf.Layout.FooterF (qf.Layout.FooterBold <> 0uy)
        | _ -> ()

        doc.BodyParagraphs
        |> List.iteri (fun i p ->
            if i > 0 then sb.Append(@"\par").Append("\r\n") |> ignore
            appendBodyParagraph sb qf p)

        sb.Append('}').Append("\r\n") |> ignore
        sb.ToString()

    let writeRtf (qf: QuillFile) (doc: DecodedDocument) (outPath: string) =
        // Everything outside 7-bit ASCII is escaped, so the file is pure ASCII.
        File.WriteAllText(outPath, buildRtf qf doc, Encoding.ASCII)
