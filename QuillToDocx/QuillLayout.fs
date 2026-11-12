namespace QuillToDocx

open System.Text.RegularExpressions

/// Page and paragraph geometry shared by the word-processor writers (DOCX and RTF).
/// Quill measures in character columns; one column is 120 twips (12 characters per inch).
module QuillLayout =

    type Alignment =
        | AlignLeft
        | AlignCenter
        | AlignRight
        | AlignJustify

    type TabKind =
        | TabLeft
        | TabCenter
        | TabRight
        | TabDecimal

    type PageNumberStyle =
        | PageArabic
        | PageAlphabetic
        | PageRoman

    type HeaderFooterPiece =
        | FixedText of string
        | PageNumberField of PageNumberStyle

    type PageGeometry =
        { Width: int
          Height: int
          TopMargin: int
          BottomMargin: int
          LeftMargin: int
          RightMargin: int
          HeaderMargin: int
          FooterMargin: int }

    type Indents =
        { LeftIndent: int
          RightIndent: int
          FirstLineIndent: int }

    [<Literal>]
    let private DefaultPageMargin = 1440

    [<Literal>]
    let private DefaultHeaderFooterMargin = 720

    [<Literal>]
    let private DefaultPageWidth = 12480

    [<Literal>]
    let private DefaultPageHeight = 15840

    [<Literal>]
    let private DefaultBodyWidthCols = 80

    let twipsFromCols cols =
        cols * 120

    let private clampTwips defaultValue minValue maxValue rawValue =
        if rawValue >= minValue && rawValue <= maxValue then
            rawValue
        else
            defaultValue

    let lineSpacingTwips (layout: LayoutTable) =
        clampTwips 240 200 480 (240 + (int layout.LineGap * 10))

    let paragraphAfterTwips (layout: LayoutTable) =
        clampTwips 80 0 240 (lineSpacingTwips layout / 3)

    /// Quill counts the page length and the top and bottom margins in printer lines, six to the inch.
    [<Literal>]
    let private PrinterLineTwips = 240

    let pageGeometry (layout: LayoutTable) =
        let topMargin = clampTwips DefaultPageMargin 720 2880 (int layout.TopMargin * PrinterLineTwips)
        let bottomMargin = clampTwips DefaultPageMargin 720 2880 (int layout.BottomMarg * PrinterLineTwips)
        let leftMargin = DefaultPageMargin
        let rightMargin = DefaultPageMargin

        { Width = clampTwips DefaultPageWidth 10080 15840 (leftMargin + rightMargin + twipsFromCols DefaultBodyWidthCols)
          // The page length is the whole sheet, margins included: 66 lines is 11 inches, 70 is A4.
          Height = clampTwips DefaultPageHeight 10080 20160 (int layout.PageLen * PrinterLineTwips)
          TopMargin = topMargin
          BottomMargin = bottomMargin
          LeftMargin = leftMargin
          RightMargin = rightMargin
          HeaderMargin = clampTwips DefaultHeaderFooterMargin 240 1440 (int layout.HeaderMarg * 240)
          FooterMargin = clampTwips DefaultHeaderFooterMargin 240 1440 (int layout.FooterMarg * 240) }

    /// QL Quill uses 0-2 for left/centre/right, PC Quill 4-6 (the top four bits hold line
    /// spacing). Quill's "right" justification is the old word-processor sense: lines flush to
    /// both margins, i.e. justified. Prose paragraphs use it throughout the corpus, with
    /// headings centred; true right alignment would put whole manuals against the right margin.
    let paragraphAlignment (justif: byte) =
        match int (justif &&& 0x0Fuy) with
        | 1
        | 5 -> AlignCenter
        | 2
        | 6 -> AlignJustify
        | _ -> AlignLeft

    /// Columns the first line starts to the right of the left margin. The indent margin is an
    /// absolute column, so it can lie left of the left margin: a hanging indent, which glossaries
    /// use for "keyword at column 14, description from column 24". The result is then negative.
    /// Centred paragraphs are centred between the left and right margins: applying the indent
    /// margin as well would push headings off centre (some documents set it to 20+ columns).
    let firstLineIndentCols (p: ParaTable) =
        match paragraphAlignment p.Justif with
        | AlignCenter -> 0
        | _ -> int p.IndentMarg - int p.LeftMarg

    let paragraphIndents (p: ParaTable) =
        { LeftIndent = twipsFromCols (int p.LeftMarg)
          RightIndent = twipsFromCols (max 0 (DefaultBodyWidthCols - int p.RightMarg))
          FirstLineIndent = twipsFromCols (firstLineIndentCols p) }

    /// Layout-table header/footer justification: 0 none (not printed), 1 left, 2 centre, 3 right.
    let headerFooterAlignment (flag: byte) =
        match int flag with
        | 2 -> AlignCenter
        | 3 -> AlignRight
        | _ -> AlignLeft

    let private pageNumberPattern = Regex("nnn|aaa|rrr", RegexOptions.Compiled)

    /// Quill prints the page number where header or footer text says nnn, as a letter for aaa
    /// and in roman numerals for rrr.
    let headerFooterPieces (text: string) =
        let pieces = ResizeArray<HeaderFooterPiece>()
        let mutable start = 0

        for m in pageNumberPattern.Matches(text) do
            if m.Index > start then pieces.Add(FixedText(text.Substring(start, m.Index - start)))

            let style =
                match m.Value with
                | "aaa" -> PageAlphabetic
                | "rrr" -> PageRoman
                | _ -> PageArabic

            pieces.Add(PageNumberField style)
            start <- m.Index + m.Length

        if start < text.Length then pieces.Add(FixedText(text.Substring(start)))
        List.ofSeq pieces

    /// The Word field that prints the page number in the given style.
    let pageFieldInstruction style =
        match style with
        | PageArabic -> "PAGE"
        | PageAlphabetic -> @"PAGE \* alphabetic"
        | PageRoman -> @"PAGE \* roman"

    let tabStops (qf: QuillFile) (p: ParaTable) =
        qf.TabTables
        |> List.tryFind (fun t -> t.EntryNumber = p.TabTable)
        |> Option.map (fun t ->
            [ for e in t.Entries ->
                  let kind =
                      match int e.Type with
                      | 1 -> TabCenter
                      | 2 -> TabRight
                      | 3 -> TabDecimal
                      | _ -> TabLeft

                  twipsFromCols (int e.Pos), kind ])
