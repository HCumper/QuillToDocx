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

    let private paragraphSpacing (layout: LayoutTable) =
        SpacingBetweenLines(
            Before = StringValue("0"),
            After = StringValue((QuillLayout.paragraphAfterTwips layout).ToString()),
            Line = StringValue((QuillLayout.lineSpacingTwips layout).ToString()),
            LineRule = EnumValue(LineSpacingRuleValues.Auto))

    let private justification alignment =
        match alignment with
        | QuillLayout.AlignLeft -> JustificationValues.Left
        | QuillLayout.AlignCenter -> JustificationValues.Center
        | QuillLayout.AlignRight -> JustificationValues.Right
        | QuillLayout.AlignJustify -> JustificationValues.Both

    let private tabStopsForParagraph (qf: QuillFile) (p: ParaTable) =
        QuillLayout.tabStops qf p
        |> Option.map (fun stops ->
            let tabs = Tabs()
            for position, kind in stops do
                let stop =
                    TabStop(
                        Position = Int32Value(position),
                        Val =
                            match kind with
                            | QuillLayout.TabCenter -> TabStopValues.Center
                            | QuillLayout.TabRight -> TabStopValues.Right
                            | QuillLayout.TabDecimal -> TabStopValues.Decimal
                            | QuillLayout.TabLeft -> TabStopValues.Left)
                tabs.AppendChild(stop) |> ignore
            tabs)

    let private runProps (s: TextStyle) =
        let rp = RunProperties()
        rp.AppendChild(RunFonts(Ascii = DefaultFontName, HighAnsi = DefaultFontName, ComplexScript = DefaultFontName)) |> ignore
        if s.Bold then rp.AppendChild(Bold()) |> ignore
        if s.Italic then rp.AppendChild(Italic()) |> ignore
        rp.AppendChild(FontSize(Val = StringValue(DefaultFontSize))) |> ignore
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

        | RunPageBreak ->
            let r = Run()
            r.AppendChild(Break(Type = BreakValues.Page)) |> ignore
            r

    let private makePageField style pageStyle =
        let field = SimpleField(Instruction = StringValue(" " + QuillLayout.pageFieldInstruction pageStyle + " "))
        let r = Run()
        r.AppendChild(runProps style) |> ignore
        r.AppendChild(Text("1")) |> ignore
        field.AppendChild(r) |> ignore
        field

    let private paraProps (qf: QuillFile) bodyCols (p: ParaTable) pageBreakBefore =
        let pp = ParagraphProperties()

        pp.AppendChild(ParagraphStyleId(Val = "Normal")) |> ignore

        // Schema order: pageBreakBefore follows pStyle and precedes tabs.
        if pageBreakBefore then pp.AppendChild(PageBreakBefore()) |> ignore

        let indents = QuillLayout.paragraphIndents p bodyCols

        let ind = Indentation(Left = StringValue(indents.LeftIndent.ToString()))
        if indents.RightIndent > 0 then ind.Right <- StringValue(indents.RightIndent.ToString())
        if indents.FirstLineIndent > 0 then
            ind.FirstLine <- StringValue(indents.FirstLineIndent.ToString())
        elif indents.FirstLineIndent < 0 then
            ind.Hanging <- StringValue((-indents.FirstLineIndent).ToString())

        match tabStopsForParagraph qf p with
        | Some tabs -> pp.AppendChild(tabs) |> ignore
        | None -> ()

        pp.AppendChild(paragraphSpacing qf.Layout) |> ignore
        pp.AppendChild(ind) |> ignore
        pp.AppendChild(Justification(Val = justification (QuillLayout.paragraphAlignment p.Justif))) |> ignore

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

    let private buildParagraph (qf: QuillFile) bodyCols (p: DecodedParagraph) =
        let para = Paragraph()
        para.AppendChild(paraProps qf bodyCols p.Descriptor p.PageBreakBefore) |> ignore
        for item in p.Content do
            para.AppendChild(makeRun item) |> ignore
        para

    let private buildHeaderFooterParagraph (layout: LayoutTable) (items: Inline list) justify boldFlag =
        let para = Paragraph()
        let pp = ParagraphProperties()
        pp.AppendChild(ParagraphStyleId(Val = "Normal")) |> ignore
        pp.AppendChild(paragraphSpacing layout) |> ignore
        pp.AppendChild(Justification(Val = justify)) |> ignore
        para.AppendChild(pp) |> ignore

        let items2 =
            if not boldFlag then items
            else
                items
                |> List.map (function
                    | RunText (t, s) -> RunText(t, { s with Bold = true })
                    | RunTab s -> RunTab { s with Bold = true }
                    | RunSoftHyphen s -> RunSoftHyphen { s with Bold = true }
                    | RunPageBreak -> RunPageBreak)

        // Quill's page-number placeholders become PAGE fields. A header can't break the page.
        for i in items2 do
            match i with
            | RunText (text, style) ->
                for piece in QuillLayout.headerFooterPieces text do
                    match piece with
                    | QuillLayout.FixedText t -> para.AppendChild(makeRun (RunText(t, style))) |> ignore
                    | QuillLayout.PageNumberField pageStyle -> para.AppendChild(makePageField style pageStyle) |> ignore
            | RunPageBreak -> ()
            | other -> para.AppendChild(makeRun other) |> ignore

        para

    let private headerJustify flag =
        justification (QuillLayout.headerFooterAlignment flag)

    let writeDocx (qf: QuillFile) (doc: DecodedDocument) (outPath: string) =
        if IO.File.Exists outPath then
            IO.File.Delete outPath

        use pkg = WordprocessingDocument.Create(outPath, WordprocessingDocumentType.Document)
        let main = pkg.AddMainDocumentPart()
        createStylesPart qf.Layout main
        let body = Body()
        let bodyCols = QuillLayout.bodyWidthCols qf

        for p in doc.BodyParagraphs do
            body.AppendChild(buildParagraph qf bodyCols p) |> ignore

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

        let page = QuillLayout.pageGeometry qf.Layout bodyCols

        let pageMargin =
            PageMargin(
                Top = Int32Value(page.TopMargin),
                Bottom = Int32Value(page.BottomMargin),
                Left = UInt32Value(uint32 page.LeftMargin),
                Right = UInt32Value(uint32 page.RightMargin),
                Header = UInt32Value(uint32 page.HeaderMargin),
                Footer = UInt32Value(uint32 page.FooterMargin)
            )

        let pageSize =
            PageSize(
                Width = UInt32Value(uint32 page.Width),
                Height = UInt32Value(uint32 page.Height))

        sect.AppendChild(pageSize) |> ignore
        sect.AppendChild(pageMargin) |> ignore

        if qf.Layout.StartPage > 1uy then
            sect.AppendChild(PageNumberType(Start = Int32Value(int qf.Layout.StartPage))) |> ignore

        body.AppendChild(sect) |> ignore
        let document = Document()
        document.Body <- body
        main.Document <- document
        main.Document.Save()
