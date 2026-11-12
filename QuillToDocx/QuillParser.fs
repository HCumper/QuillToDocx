namespace QuillToDocx

open System
open System.IO
open System.Text

module QuillParser =

    open QuillBinary

    let private readU16At isPcFile (bytes: byte[]) offset =
        if isPcFile then
            uint16 bytes[offset] ||| (uint16 bytes[offset + 1] <<< 8)
        else
            (uint16 bytes[offset] <<< 8) ||| uint16 bytes[offset + 1]

    let private looksLikeHtml (bytes: byte[]) =
        let len = min bytes.Length 128
        let prefix = Encoding.ASCII.GetString(bytes, 0, len).TrimStart()

        prefix.StartsWith("<!DOCTYPE", StringComparison.OrdinalIgnoreCase)
        || prefix.StartsWith("<html", StringComparison.OrdinalIgnoreCase)

    let private validateHeader (bytes: byte[]) (h: Header) =
        if h.HeaderLen <> 20us then
            if looksLikeHtml bytes then
                failwith "Input appears to be an HTML page, not a raw Quill document."
            else
                failwithf "Unexpected header length: %d" h.HeaderLen
        if h.Id <> "vrm1qdf0" then
            failwithf "Not a Quill document. Identifier was '%s'" h.Id

    let private tryParagraphTableShape isPcFile (bytes: byte[]) paraLen offset =
        let paraHeadSize = 8
        let paraEntrySize = 14

        if paraLen < paraHeadSize || offset < 0 || offset + paraHeadSize > bytes.Length || offset + paraLen > bytes.Length then
            None
        else
            let usedParas = int (readU16At isPcFile bytes (offset + 4))
            let allocParas = int (readU16At isPcFile bytes (offset + 6))

            if allocParas >= usedParas && paraHeadSize + (allocParas * paraEntrySize) <= paraLen then
                Some(usedParas, allocParas)
            else
                None

    let private resolveParagraphSection isPcFile (bytes: byte[]) (header: Header) =
        let paraLen = int header.ParaLen
        let primaryOffset = int header.TextLen
        let inferredOffset =
            bytes.Length - paraLen - int header.FreeLen - int header.LayoutLen

        match tryParagraphTableShape isPcFile bytes paraLen primaryOffset with
        | Some shape -> primaryOffset, shape
        | None ->
            match tryParagraphTableShape isPcFile bytes paraLen inferredOffset with
            | Some shape -> inferredOffset, shape
            | None ->
                let usedParas =
                    if primaryOffset + 6 < bytes.Length then
                        int (readU16At isPcFile bytes (primaryOffset + 4))
                    else
                        -1

                let allocParas =
                    if primaryOffset + 8 <= bytes.Length then
                        int (readU16At isPcFile bytes (primaryOffset + 6))
                    else
                        -1

                failwithf "Unexpected paragraph table shape: used=%d alloc=%d len=%d" usedParas allocParas paraLen

    let private parseTabTables (qf: QuillFile) (bytes: byte[]) (layoutOffset: int) (layout: LayoutTable) =
        let start = layoutOffset + (if qf.IsPcFile then 22 else 20)
        let used = int layout.TabSize

        if used <= 0 then
            []
        else
            let mutable pos = start
            let endPos = start + used
            let results = ResizeArray<TabTableDef>()

            while pos + 1 < endPos do
                let entry = bytes[pos]
                let len = bytes[pos + 1]

                if entry = 0uy || len < 2uy then
                    pos <- endPos
                else
                    let entryCount = (int len / 2) - 1
                    let items =
                        [ for i in 0 .. entryCount - 1 do
                            let off = pos + 2 + (i * 2)
                            if off + 1 < endPos then
                                yield
                                    { Pos = bytes[off]
                                      Type = bytes[off + 1] } ]

                    results.Add
                        { EntryNumber = entry
                          Entries = items }

                    pos <- pos + int len

            results |> Seq.toList

    let parseFile (path: string) =
        let bytes = File.ReadAllBytes path

        if bytes.Length < 20 then
            failwith "File too small."

        let isPcFile, header = readHeader bytes
        validateHeader bytes header

        let textOffset = int header.HeaderLen
        let paraLen = int header.ParaLen
        let paraHeadOffset, (usedParas, allocParas) = resolveParagraphSection isPcFile bytes header

        if paraHeadOffset > bytes.Length || paraHeadOffset < textOffset then
            failwith "Text area length exceeds file length."

        let textBuffer = bytes[textOffset .. paraHeadOffset - 1]

        let paraHeadSize = 8
        let paraEntrySize = 14

        let paraHead =
            { Size = uint16 paraEntrySize
              Gran = 0us
              Used = uint16 usedParas
              Alloc = uint16 allocParas }
        // Entry 0 is a dummy, so the table holds Used + 1 entries: the dummy, then the header,
        // the footer and the body paragraphs. Reading only Used entries drops the last paragraph.
        let entryCount = min (usedParas + 1) allocParas

        let paras =
            Array.init entryCount (fun i ->
                let off = paraHeadOffset + paraHeadSize + (i * paraEntrySize)
                readPara isPcFile bytes off)

        let layoutOffset = paraHeadOffset + int header.ParaLen + int header.FreeLen
        let layout = readLayout isPcFile bytes layoutOffset

        let quillFile =
            { Header = header
              IsPcFile = isPcFile
              TextBuffer = textBuffer
              ParaHead = paraHead
              Paras = paras
              Layout = layout
              TabTables = [] }

        let tabTables = parseTabTables quillFile bytes layoutOffset layout

        { quillFile with
            TabTables = tabTables }
