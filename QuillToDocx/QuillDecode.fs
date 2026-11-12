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
    let Italic = 0x13uy

    [<Literal>]
    let SoftHyphen = 0x1Euy

    [<Literal>]
    let FormFeed = 0x0Cuy

    // Paragraph-table entry 0 is not a paragraph. Entries 1 and 2 are always the header and
    // footer text, stored even when the layout table switches them off; the body starts at 3.
    [<Literal>]
    let private HeaderEntry = 1

    [<Literal>]
    let private FooterEntry = 2

    [<Literal>]
    let private FirstBodyEntry = 3

    // .NET only ships code page 437 (PC Quill's character set) through the code-pages provider.
    let private cp437 =
        lazy
            (Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)
             Encoding.GetEncoding(437))

    let private defaultStyle =
        { Bold = false
          Italic = false
          Underline = false
          Sub = false
          Super = false }

    /// QL characters 0x80-0xBF: accented letters, Greek letters and arrows. 0xB5 was an odd
    /// script mark in the original ROM; SMSQ/E made it the euro sign, which is how documents use it.
    let private qlHighCharacters =
        "\u00E4\u00E3\u00E5\u00E9\u00F6\u00F5\u00F8\u00FC\u00E7\u00F1\u00E6\u0153\u00E1\u00E0\u00E2\u00EB" // 80
        + "\u00E8\u00EA\u00EF\u00ED\u00EC\u00EE\u00F3\u00F2\u00F4\u00FA\u00F9\u00FB\u00DF\u00A2\u00A5`" // 90
        + "\u00C4\u00C3\u00C5\u00C9\u00D6\u00D5\u00D8\u00DC\u00C7\u00D1\u00C6\u0152\u03B1\u03B4\u03B8\u03BB" // A0
        + "\u00B5\u03C0\u03C6\u00A1\u00BF\u20AC\u00A7\u00A4\u00AB\u00BB\u00B0\u00F7\u2190\u2192\u2191\u2193" // B0

    let private qlCharToUnicode (isPcFile: bool) (b: byte) =
        if b < 0x20uy then
            ""
        elif isPcFile then
            cp437.Value.GetString([| b |])
        else
            // The QL set is ASCII except for 0x60 (pound) and 0x7F (copyright).
            match int b with
            | 0x60 -> "\u00A3"
            | 0x7F -> "\u00A9"
            | x when x < 0x7F -> string (char x)
            | x when x <= 0xBF -> string qlHighCharacters[x - 0x80]
            // Not QL characters; the few files that use them hold Latin-1 text (0xE9 for e-acute).
            | x -> string (char x)

    let private getParagraphBytes (qf: QuillFile) (p: ParaTable) =
        let start = int p.Offset - int qf.Header.HeaderLen
        let len = int p.ParaLen
        if len = 0 then
            [||]
        else
            if start < 0 || start + len > qf.TextBuffer.Length then
                failwithf "Paragraph out of range: offset=%d len=%d" start len
            qf.TextBuffer[start .. start + len - 1]

    let private decodeParagraphContent (qf: QuillFile) (bytes: byte[]) =
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

            | Italic ->
                flush ()
                style <- { style with Italic = not style.Italic }

            | SoftHyphen ->
                flush ()
                items.Add(RunSoftHyphen style)

            | FormFeed ->
                flush ()
                items.Add(RunPageBreak)

            | _ ->
                let s = qlCharToUnicode qf.IsPcFile b
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
            | RunPageBreak -> ()

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
            let content = decodeParagraphContent qf (current.ToArray())
            if not content.IsEmpty then
                paras.Add(
                    { Index = index
                      Descriptor = fallbackParagraphDescriptor index
                      Content = content
                      PageBreakBefore = false })
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

    let private tryDecodeEntry (qf: QuillFile) index =
        try
            let p = qf.Paras[index]

            Some
                { Index = index
                  Descriptor = p
                  Content = decodeParagraphContent qf (getParagraphBytes qf p)
                  PageBreakBefore = false }
        with _ ->
            None

    let private isBlank item =
        match item with
        | RunText (text, _) -> String.IsNullOrWhiteSpace text
        | RunTab _
        | RunSoftHyphen _
        | RunPageBreak -> true

    /// Quill's form feed is a page break. A paragraph holding nothing else is dropped. A break
    /// before a paragraph's text, or after it, becomes a break before that paragraph or the next;
    /// one inside the text stays where it is.
    let private applyPageBreaks (paragraphs: DecodedParagraph list) =
        let folder (pending, acc) (para: DecodedParagraph) =
            let content = para.Content

            if List.forall isBlank content then
                if List.contains RunPageBreak content then
                    true, acc
                else
                    false, { para with PageBreakBefore = pending } :: acc
            else
                let leading = content |> List.takeWhile isBlank
                let trailing = content |> List.rev |> List.takeWhile isBlank |> List.rev
                let text = content |> List.skip leading.Length |> List.take (content.Length - leading.Length - trailing.Length)
                let withoutBreaks = List.filter ((<>) RunPageBreak)

                List.contains RunPageBreak trailing,
                { para with
                    Content = withoutBreaks leading @ text @ withoutBreaks trailing
                    PageBreakBefore = pending || List.contains RunPageBreak leading }
                :: acc

        paragraphs |> List.fold folder (false, []) |> snd |> List.rev

    let decodeDocument (qf: QuillFile) =
        let entries =
            [ for i in HeaderEntry .. qf.Paras.Length - 1 do
                  match tryDecodeEntry qf i with
                  | Some entry -> yield entry
                  | None -> () ]

        let parsedParasHaveContent = entries |> List.exists (fun p -> not p.Content.IsEmpty)

        if not parsedParasHaveContent then
            { HeaderParagraph = None
              FooterParagraph = None
              BodyParagraphs = trimTrailingMarkerParagraphs (decodeFromTextBuffer qf)
              Layout = qf.Layout }
        else
            let entryAt index =
                entries |> List.tryFind (fun p -> p.Index = index)

            let body =
                entries
                |> List.filter (fun p -> p.Index >= FirstBodyEntry)
                |> applyPageBreaks
                |> trimTrailingMarkerParagraphs

            { HeaderParagraph = if qf.Layout.HeaderF <> 0uy then entryAt HeaderEntry else None
              FooterParagraph = if qf.Layout.FooterF <> 0uy then entryAt FooterEntry else None
              BodyParagraphs = body
              Layout = qf.Layout }
