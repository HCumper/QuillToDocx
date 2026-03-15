namespace QuillToDocx

open System
open System.Text
open System.Text.RegularExpressions

module QuillDecode =

    [<Literal>]
    let EndText = 0x0Euy

    [<Literal>]
    let EndPara = 0x00uy

    [<Literal>]
    let Tab = 0x09uy

    [<Literal>]
    let Bold = 0x0Fuy

    [<Literal>]
    let Underline = 0x10uy

    [<Literal>]
    let SubScript = 0x11uy

    [<Literal>]
    let SuperScript = 0x12uy

    [<Literal>]
    let SoftHyphen = 0x1Euy

    let private defaultStyle =
        { Bold = false
          Underline = false
          Sub = false
          Super = false }

    let private qlCharToUnicode (b: byte) =
        match int b with
        | x when x >= 0x20 && x <= 0x7E -> string (char x)
        | 0x7F -> "\u00A3"
        | 0x80 -> "\u00C4"
        | 0x81 -> "\u00A3"
        | 0x82 -> "\u00C5"
        | 0x83 -> "\u00C9"
        | 0x84 -> "\u00F6"
        | 0x85 -> "\u00F5"
        | 0x86 -> "\u00F8"
        | 0x87 -> "\u00FC"
        | 0x88 -> "\u00E7"
        | 0x89 -> "\u00F1"
        | 0x8A -> "\u00FD"
        | 0x8B -> "\u0152"
        | 0x8C -> "\u00E1"
        | 0x8D -> "\u00E0"
        | 0x8E -> "\u00E2"
        | 0x8F -> "\u00EB"
        | 0x90 -> "\u00E8"
        | 0x91 -> "\u00EA"
        | 0x92 -> "\u00EF"
        | 0x93 -> "\u00ED"
        | 0x94 -> "\u00EC"
        | 0x95 -> "\u00EE"
        | 0x96 -> "\u00F3"
        | 0x97 -> "\u00F2"
        | 0x98 -> "\u00F4"
        | 0x99 -> "\u00FA"
        | 0x9A -> "\u00F9"
        | 0x9B -> "\u00FB"
        | 0x9C -> "\u00DF"
        | 0x9D -> "\u00A2"
        | 0x9E -> "\u00A5"
        | _ when b < 0x20uy -> ""
        | _ ->
            Encoding.GetEncoding("ISO-8859-1").GetString([| b |])

    let private getParagraphBytes (qf: QuillFile) (p: ParaTable) =
        let start = int p.Offset
        let len = int p.ParaLen
        if len = 0 then
            [||]
        else
            if start < 0 || start + len > qf.TextBuffer.Length then
                failwithf "Paragraph out of range: offset=%d len=%d" start len
            qf.TextBuffer[start .. start + len - 1]

    let private decodeParagraphContent (bytes: byte[]) =
        let mutable style = defaultStyle
        let sb = StringBuilder()
        let items = ResizeArray<Inline>()

        let flush () =
            if sb.Length > 0 then
                items.Add(RunText(sb.ToString(), style))
                sb.Clear() |> ignore

        for b in bytes do
            match b with
            | EndPara ->
                flush ()
                style <- defaultStyle

            | Tab ->
                flush ()
                items.Add(RunTab style)

            | Bold ->
                flush ()
                style <- { style with Bold = not style.Bold }

            | Underline ->
                flush ()
                style <- { style with Underline = not style.Underline }

            | SubScript ->
                flush ()
                style <- { style with Sub = not style.Sub; Super = false }

            | SuperScript ->
                flush ()
                style <- { style with Super = not style.Super; Sub = false }

            | SoftHyphen ->
                flush ()
                items.Add(RunSoftHyphen style)

            | _ ->
                let s = qlCharToUnicode b
                if not (String.IsNullOrEmpty s) then
                    sb.Append(s) |> ignore

        flush ()
        items |> Seq.toList

    let private paragraphText (content: Inline list) =
        let sb = StringBuilder()

        for item in content do
            match item with
            | RunText (text, _) -> sb.Append(text) |> ignore
            | RunTab _ -> sb.Append('\t') |> ignore
            | RunSoftHyphen _ -> sb.Append('\u00AD') |> ignore

        sb.ToString().Trim()

    let private trailingMarkerPattern = Regex("^[-0-9]+\u00FF{2,}$", RegexOptions.Compiled)

    let private isTrailingMarkerParagraph (paragraph: DecodedParagraph) =
        trailingMarkerPattern.IsMatch(paragraphText paragraph.Content)

    let rec private trimTrailingMarkerParagraphs (paragraphs: DecodedParagraph list) =
        match List.rev paragraphs with
        | last :: rest when isTrailingMarkerParagraph last -> trimTrailingMarkerParagraphs (List.rev rest)
        | _ -> paragraphs

    let private fallbackParagraphDescriptor index =
        { Offset = uint32 index
          ParaLen = 0us
          Dummy = 0uy
          LeftMarg = 0uy
          IndentMarg = 0uy
          RightMarg = 80uy
          Justif = 0uy
          TabTable = 0uy
          Dummy2 = 0s }

    let private decodeFromTextBuffer (qf: QuillFile) =
        let paras = ResizeArray<DecodedParagraph>()
        let current = ResizeArray<byte>()
        let mutable index = 0
        let flushParagraph () =
            let content = decodeParagraphContent (current.ToArray())
            if not content.IsEmpty then
                paras.Add(
                    { Index = index
                      Descriptor = fallbackParagraphDescriptor index
                      Content = content })
                index <- index + 1
            current.Clear()

        for b in qf.TextBuffer do
            match b with
            | EndText ->
                flushParagraph ()
            | EndPara ->
                flushParagraph ()
            | _ ->
                current.Add(b)

        if current.Count > 0 then
            flushParagraph ()

        paras |> Seq.toList

    let decodeDocument (qf: QuillFile) =
        let paras =
            qf.Paras
            |> Array.mapi (fun i p ->
                try
                    let bytes = getParagraphBytes qf p
                    Some
                        { Index = i
                          Descriptor = p
                          Content = decodeParagraphContent bytes }
                with _ ->
                    None)
            |> Array.choose id
            |> Array.toList

        let parsedParasHaveContent = paras |> List.exists (fun p -> not p.Content.IsEmpty)
        let paras =
            if parsedParasHaveContent then
                paras
            else
                decodeFromTextBuffer qf

        let hasHeader = parsedParasHaveContent && qf.Layout.HeaderF <> 0uy
        let hasFooter = parsedParasHaveContent && qf.Layout.FooterF <> 0uy

        let headerPara =
            if hasHeader && paras.Length >= 1 then Some paras[0] else None

        let footerIndex = if hasHeader then 1 else 0

        let footerPara =
            if hasFooter && paras.Length > footerIndex then Some paras[footerIndex] else None

        let bodyStart =
            (if hasHeader then 1 else 0) +
            (if hasFooter then 1 else 0)

        let body =
            if paras.Length <= bodyStart then [] else paras |> List.skip bodyStart

        let body = trimTrailingMarkerParagraphs body

        { HeaderParagraph = headerPara
          FooterParagraph = footerPara
          BodyParagraphs = body
          Layout = qf.Layout }
