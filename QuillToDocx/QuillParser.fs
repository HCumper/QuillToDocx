namespace QuillToDocx

open System
open System.IO

module QuillParser =

    open QuillBinary

    let private readU16BEAt (bytes: byte[]) offset =
        (uint16 bytes[offset] <<< 8) ||| uint16 bytes[offset + 1]

    let private validateHeader (h: Header) =
        if h.HeaderLen <> 20us then
            failwithf "Unexpected header length: %d" h.HeaderLen
        if h.Id <> "vrm1qdf0" then
            failwithf "Not a Quill document. Identifier was '%s'" h.Id

    let private parseTabTables (bytes: byte[]) (layoutOffset: int) (layout: LayoutTable) =
        let start = layoutOffset + 20
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

        let header = readHeader bytes
        validateHeader header

        let textOffset = int header.HeaderLen
        let textLen = int header.TextLen
        if textOffset + textLen > bytes.Length then
            failwith "Text area length exceeds file length."

        let textBuffer = bytes[textOffset .. textOffset + textLen - 1]

        let paraHeadOffset = textOffset + textLen
        let paraHeadSize = 8
        let paraEntrySize = 14
        let paraLen = int header.ParaLen
        if paraLen < paraHeadSize then
            failwithf "Unexpected paragraph section length: %d" paraLen

        let usedParas = int (readU16BEAt bytes (paraHeadOffset + 4))
        let allocParas = int (readU16BEAt bytes (paraHeadOffset + 6))
        if usedParas < 0 || allocParas < usedParas || paraHeadSize + (allocParas * paraEntrySize) > paraLen then
            failwithf "Unexpected paragraph table shape: used=%d alloc=%d len=%d" usedParas allocParas paraLen

        let paraHead =
            { Size = uint16 paraEntrySize
              Gran = 0us
              Used = uint16 usedParas
              Alloc = uint16 allocParas }
        let paras =
            Array.init usedParas (fun i ->
                let off = paraHeadOffset + paraHeadSize + (i * paraEntrySize)
                readPara bytes off)

        let layoutOffset = textOffset + textLen + int header.ParaLen + int header.FreeLen
        let layout = readLayout bytes layoutOffset
        let tabTables = parseTabTables bytes layoutOffset layout

        { Header = header
          TextBuffer = textBuffer
          ParaHead = paraHead
          Paras = paras
          Layout = layout
          TabTables = tabTables }
