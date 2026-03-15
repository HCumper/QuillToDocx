#nowarn "3261"
#nowarn "3391"

namespace QuillToDocx

open System
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Wordprocessing

module DocxWriter =

    [<Literal>]
    let private DefaultFontName = "Courier New"

    [<Literal>]
    let private DefaultFontSize = "20"

    [<Literal>]
    let private DefaultParagraphAfter = "80"

    [<Literal>]
    let private DefaultLineSpacing = "240"

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

    let private twipsFromCols cols =
        cols * 120

    let private clampTwips defaultValue minValue maxValue rawValue =
        if rawValue >= minValue && rawValue <= maxValue then
            rawValue
        else
            defaultValue

    let private validatedPageMargin rawValue =
        clampTwips DefaultPageMargin 720 2880 rawValue

    let private validatedHeaderFooterMargin rawValue =
        clampTwips DefaultHeaderFooterMargin 240 1440 rawValue

    let private lineSpacingTwips (layout: LayoutTable) =
        clampTwips 240 200 480 (240 + (int layout.LineGap * 10))

    let private paragraphAfterTwips (layout: LayoutTable) =
        clampTwips 80 0 240 (lineSpacingTwips layout / 3)

    let private paragraphSpacing (layout: LayoutTable) =
        SpacingBetweenLines(
            Before = StringValue("0"),
            After = StringValue((paragraphAfterTwips layout).ToString()),
            Line = StringValue((lineSpacingTwips layout).ToString()),
            LineRule = EnumValue(LineSpacingRuleValues.Auto))

    let private pageHeightTwips (layout: LayoutTable) topMargin bottomMargin =
        let candidate = topMargin + bottomMargin + (int layout.PageLen * lineSpacingTwips layout)
        clampTwips DefaultPageHeight 10080 20160 candidate

    let private pageWidthTwips leftMargin rightMargin =
        clampTwips DefaultPageWidth 10080 15840 (leftMargin + rightMargin + twipsFromCols DefaultBodyWidthCols)

    let private mapJustification j =
        match int j with
        | 0
        | 4 -> JustificationValues.Left
        | 1
        | 5 -> JustificationValues.Center
        | 2
        | 6 -> JustificationValues.Right
        | _ -> JustificationValues.Left

    let private tabStopsForParagraph (qf: QuillFile) (p: ParaTable) =
        qf.TabTables
        |> List.tryFind (fun t -> t.EntryNumber = p.TabTable)
        |> Option.map (fun t ->
            let tabs = Tabs()
            for e in t.Entries do
                let stop =
                    TabStop(
                        Position = Int32Value(twipsFromCols (int e.Pos)),
                        Val =
                            match int e.Type with
                            | 1 -> TabStopValues.Center
                            | 2 -> TabStopValues.Right
                            | _ -> TabStopValues.Left)
                tabs.AppendChild(stop) |> ignore
            tabs)

    let private runProps (s: TextStyle) =
        let rp = RunProperties()
        rp.AppendChild(RunFonts(Ascii = DefaultFontName, HighAnsi = DefaultFontName, ComplexScript = DefaultFontName)) |> ignore
        rp.AppendChild(FontSize(Val = StringValue(DefaultFontSize))) |> ignore
        if s.Bold then rp.AppendChild(Bold()) |> ignore
        if s.Underline then rp.AppendChild(Underline(Val = UnderlineValues.Single)) |> ignore
        if s.Sub then rp.AppendChild(VerticalTextAlignment(Val = VerticalPositionValues.Subscript)) |> ignore
        if s.Super then rp.AppendChild(VerticalTextAlignment(Val = VerticalPositionValues.Superscript)) |> ignore
        rp

    let private makeRun inlineItem =
        match inlineItem with
        | RunText (txt, s) ->
            let r = Run()
            let rp = runProps s
            if rp.HasChildren then r.AppendChild(rp) |> ignore
            let t = Text(txt)
            t.Space <- SpaceProcessingModeValues.Preserve
            r.AppendChild(t) |> ignore
            r

        | RunTab s ->
            let r = Run()
            let rp = runProps s
            if rp.HasChildren then r.AppendChild(rp) |> ignore
            r.AppendChild(TabChar()) |> ignore
            r

        | RunSoftHyphen s ->
            let r = Run()
            let rp = runProps s
            if rp.HasChildren then r.AppendChild(rp) |> ignore
            r.AppendChild(SoftHyphen()) |> ignore
            r

    let private paraProps (qf: QuillFile) (p: ParaTable) =
        let pp = ParagraphProperties()

        pp.AppendChild(ParagraphStyleId(Val = "Normal")) |> ignore
        pp.AppendChild(Justification(Val = mapJustification p.Justif)) |> ignore
        pp.AppendChild(paragraphSpacing qf.Layout) |> ignore

        let left = twipsFromCols (int p.LeftMarg)
        let right = twipsFromCols (max 0 (80 - int p.RightMarg))
        let firstLine = twipsFromCols (max 0 (int p.IndentMarg - int p.LeftMarg))

        let ind = Indentation(Left = StringValue(left.ToString()))
        if right > 0 then ind.Right <- StringValue(right.ToString())
        if firstLine > 0 then ind.FirstLine <- StringValue(firstLine.ToString())
        pp.AppendChild(ind) |> ignore

        match tabStopsForParagraph qf p with
        | Some tabs -> pp.AppendChild(tabs) |> ignore
        | None -> ()

        pp

    let private createStylesPart (layout: LayoutTable) (main: MainDocumentPart) =
        let stylesPart = main.AddNewPart<StyleDefinitionsPart>()
        let styles = Styles()

        let docDefaults = DocDefaults()

        let runDefaults = RunPropertiesDefault()
        let defaultRunProps = RunPropertiesBaseStyle()
        defaultRunProps.AppendChild(RunFonts(Ascii = DefaultFontName, HighAnsi = DefaultFontName, ComplexScript = DefaultFontName)) |> ignore
        defaultRunProps.AppendChild(FontSize(Val = StringValue(DefaultFontSize))) |> ignore
        runDefaults.AppendChild(defaultRunProps) |> ignore

        let paragraphDefaults = ParagraphPropertiesDefault()
        let defaultParagraphProps = ParagraphPropertiesBaseStyle()
        defaultParagraphProps.AppendChild(paragraphSpacing layout) |> ignore
        paragraphDefaults.AppendChild(defaultParagraphProps) |> ignore

        docDefaults.AppendChild(runDefaults) |> ignore
        docDefaults.AppendChild(paragraphDefaults) |> ignore
        styles.AppendChild(docDefaults) |> ignore

        let normalStyle = Style()
        normalStyle.Type <- StyleValues.Paragraph
        normalStyle.StyleId <- StringValue("Normal")
        normalStyle.Default <- OnOffValue.FromBoolean(true)
        normalStyle.AppendChild(StyleName(Val = StringValue("Normal"))) |> ignore

        let normalParagraphProps = StyleParagraphProperties()
        normalParagraphProps.AppendChild(paragraphSpacing layout) |> ignore
        normalStyle.AppendChild(normalParagraphProps) |> ignore

        let normalRunProps = StyleRunProperties()
        normalRunProps.AppendChild(RunFonts(Ascii = DefaultFontName, HighAnsi = DefaultFontName, ComplexScript = DefaultFontName)) |> ignore
        normalRunProps.AppendChild(FontSize(Val = StringValue(DefaultFontSize))) |> ignore
        normalStyle.AppendChild(normalRunProps) |> ignore

        styles.AppendChild(normalStyle) |> ignore
        stylesPart.Styles <- styles

    let private buildParagraph (qf: QuillFile) (p: DecodedParagraph) =
        let para = Paragraph()
        para.AppendChild(paraProps qf p.Descriptor) |> ignore
        for item in p.Content do
            para.AppendChild(makeRun item) |> ignore
        para

    let private buildHeaderFooterParagraph (layout: LayoutTable) (items: Inline list) justify boldFlag =
        let para = Paragraph()
        let pp = ParagraphProperties()
        pp.AppendChild(ParagraphStyleId(Val = "Normal")) |> ignore
        pp.AppendChild(Justification(Val = justify)) |> ignore
        pp.AppendChild(paragraphSpacing layout) |> ignore
        para.AppendChild(pp) |> ignore

        let items2 =
            if not boldFlag then items
            else
                items
                |> List.map (function
                    | RunText (t, s) -> RunText(t, { s with Bold = true })
                    | RunTab s -> RunTab { s with Bold = true }
                    | RunSoftHyphen s -> RunSoftHyphen { s with Bold = true })

        for i in items2 do
            para.AppendChild(makeRun i) |> ignore

        para

    let private headerJustify flag =
        match int flag with
        | 1 -> JustificationValues.Left
        | 2 -> JustificationValues.Center
        | 3 -> JustificationValues.Right
        | _ -> JustificationValues.Left

    let writeDocx (qf: QuillFile) (doc: DecodedDocument) (outPath: string) =
        if IO.File.Exists outPath then
            IO.File.Delete outPath

        use pkg = WordprocessingDocument.Create(outPath, WordprocessingDocumentType.Document)
        let main = pkg.AddMainDocumentPart()
        createStylesPart qf.Layout main
        let body = Body()

        for p in doc.BodyParagraphs do
            body.AppendChild(buildParagraph qf p) |> ignore

        let sect = SectionProperties()

        match doc.HeaderParagraph with
        | Some hp when qf.Layout.HeaderF <> 0uy ->
            let part = main.AddNewPart<HeaderPart>()
            let hdr = Header()
            hdr.AppendChild(
                buildHeaderFooterParagraph
                    qf.Layout
                    hp.Content
                    (headerJustify qf.Layout.HeaderF)
                    (qf.Layout.HeaderBold <> 0uy))
            |> ignore
            part.Header <- hdr
            sect.AppendChild(HeaderReference(Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(part))) |> ignore
        | _ -> ()

        match doc.FooterParagraph with
        | Some fp when qf.Layout.FooterF <> 0uy ->
            let part = main.AddNewPart<FooterPart>()
            let ftr = Footer()
            ftr.AppendChild(
                buildHeaderFooterParagraph
                    qf.Layout
                    fp.Content
                    (headerJustify qf.Layout.FooterF)
                    (qf.Layout.FooterBold <> 0uy))
            |> ignore
            part.Footer <- ftr
            sect.AppendChild(FooterReference(Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(part))) |> ignore
        | _ -> ()

        let topMargin = validatedPageMargin (int qf.Layout.TopMargin * 240)
        let bottomMargin = validatedPageMargin (int qf.Layout.BottomMarg * 240)
        let leftMargin = DefaultPageMargin
        let rightMargin = DefaultPageMargin
        let headerMargin = validatedHeaderFooterMargin (int qf.Layout.HeaderMarg * 240)
        let footerMargin = validatedHeaderFooterMargin (int qf.Layout.FooterMarg * 240)
        let pageWidth = pageWidthTwips leftMargin rightMargin
        let pageHeight = pageHeightTwips qf.Layout topMargin bottomMargin

        let pageMargin =
            PageMargin(
                Top = Int32Value(topMargin),
                Bottom = Int32Value(bottomMargin),
                Left = UInt32Value(uint32 leftMargin),
                Right = UInt32Value(uint32 rightMargin),
                Header = UInt32Value(uint32 headerMargin),
                Footer = UInt32Value(uint32 footerMargin)
            )

        let pageSize =
            PageSize(
                Width = UInt32Value(uint32 pageWidth),
                Height = UInt32Value(uint32 pageHeight))

        sect.AppendChild(pageSize) |> ignore
        sect.AppendChild(pageMargin) |> ignore

        if qf.Layout.StartPage > 1uy then
            sect.AppendChild(PageNumberType(Start = Int32Value(int qf.Layout.StartPage))) |> ignore

        body.AppendChild(sect) |> ignore
        let document = Document()
        document.Body <- body
        main.Document <- document
        main.Document.Save()
