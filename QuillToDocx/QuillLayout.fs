namespace QuillToDocx

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

    let pageGeometry (layout: LayoutTable) =
        let topMargin = clampTwips DefaultPageMargin 720 2880 (int layout.TopMargin * 240)
        let bottomMargin = clampTwips DefaultPageMargin 720 2880 (int layout.BottomMarg * 240)
        let leftMargin = DefaultPageMargin
        let rightMargin = DefaultPageMargin

        { Width = clampTwips DefaultPageWidth 10080 15840 (leftMargin + rightMargin + twipsFromCols DefaultBodyWidthCols)
          Height =
            clampTwips DefaultPageHeight 10080 20160 (topMargin + bottomMargin + (int layout.PageLen * lineSpacingTwips layout))
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

    /// Columns the first line starts to the right of the left margin. Centred paragraphs are
    /// centred between the left and right margins: applying the indent margin as well would push
    /// headings off centre (by half the indent, which some documents set to 20+ columns).
    let firstLineIndentCols (p: ParaTable) =
        match paragraphAlignment p.Justif with
        | AlignCenter -> 0
        | _ -> max 0 (int p.IndentMarg - int p.LeftMarg)

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

    let tabStops (qf: QuillFile) (p: ParaTable) =
        qf.TabTables
        |> List.tryFind (fun t -> t.EntryNumber = p.TabTable)
        |> Option.map (fun t ->
            [ for e in t.Entries ->
                  let kind =
                      match int e.Type with
                      | 1 -> TabCenter
                      | 2 -> TabRight
                      | _ -> TabLeft

                  twipsFromCols (int e.Pos), kind ])
