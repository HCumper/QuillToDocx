namespace QuillToDocx

open System
open System.Text

module QuillBinary =

    let private readU16BE (bytes: byte[]) offset =
        (uint16 bytes[offset] <<< 8) ||| uint16 bytes[offset + 1]

    let private readU16LE (bytes: byte[]) offset =
        uint16 bytes[offset] ||| (uint16 bytes[offset + 1] <<< 8)

    let private readU32BE (bytes: byte[]) offset =
        (uint32 bytes[offset] <<< 24) |||
        (uint32 bytes[offset + 1] <<< 16) |||
        (uint32 bytes[offset + 2] <<< 8) |||
        uint32 bytes[offset + 3]

    let private readU32LE (bytes: byte[]) offset =
        uint32 bytes[offset] |||
        (uint32 bytes[offset + 1] <<< 8) |||
        (uint32 bytes[offset + 2] <<< 16) |||
        (uint32 bytes[offset + 3] <<< 24)

    let readAscii (bytes: byte[]) offset len =
        Encoding.ASCII.GetString(bytes, offset, len)

    let readHeader (bytes: byte[]) =
        let beHeaderLen = readU16BE bytes 0
        let isPcFile = beHeaderLen = 5120us
        let readU16 = if isPcFile then readU16LE else readU16BE
        let readU32 = if isPcFile then readU32LE else readU32BE

        isPcFile,
        { HeaderLen = readU16 bytes 0
          Id = readAscii bytes 2 8
          TextLen = readU32 bytes 10
          ParaLen = readU16 bytes 14
          FreeLen = readU16 bytes 16
          LayoutLen = readU16 bytes 18 }

    let readParaHead isPcFile (bytes: byte[]) offset =
        let readU16 = if isPcFile then readU16LE else readU16BE
        { Size = readU16 bytes offset
          Gran = readU16 bytes (offset + 2)
          Used = readU16 bytes (offset + 4)
          Alloc = readU16 bytes (offset + 6) }

    let readPara isPcFile (bytes: byte[]) offset =
        let readU16 = if isPcFile then readU16LE else readU16BE
        let readU32 = if isPcFile then readU32LE else readU32BE
        { Offset = readU32 bytes offset
          ParaLen = readU16 bytes (offset + 4)
          Dummy = bytes[offset + 6]
          LeftMarg = bytes[offset + 7]
          IndentMarg = bytes[offset + 8]
          RightMarg = bytes[offset + 9]
          Justif = bytes[offset + 10]
          TabTable = bytes[offset + 11]
          Dummy2 = int16 (readU16 bytes (offset + 12)) }

    let readLayout isPcFile (bytes: byte[]) offset =
        if isPcFile then
            { BottomMarg = bytes[offset + 0]
              DispMode = 0uy
              LineGap = bytes[offset + 1]
              PageLen = bytes[offset + 2]
              StartPage = bytes[offset + 3]
              Color = 0uy
              TopMargin = bytes[offset + 4]
              Dummy1 = bytes[offset + 5]
              WordCount = readU16LE bytes (offset + 10)
              MaxTabSize = readU16LE bytes (offset + 12)
              TabSize = readU16LE bytes (offset + 14)
              HeaderF = bytes[offset + 16]
              FooterF = bytes[offset + 17]
              HeaderMarg = bytes[offset + 18]
              FooterMarg = bytes[offset + 19]
              HeaderBold = bytes[offset + 20]
              FooterBold = bytes[offset + 21] }
        else
            { BottomMarg = bytes[offset + 0]
              DispMode = bytes[offset + 1]
              LineGap = bytes[offset + 2]
              PageLen = bytes[offset + 3]
              StartPage = bytes[offset + 4]
              Color = bytes[offset + 5]
              TopMargin = bytes[offset + 6]
              Dummy1 = bytes[offset + 7]
              WordCount = readU16BE bytes (offset + 8)
              MaxTabSize = readU16BE bytes (offset + 10)
              TabSize = readU16BE bytes (offset + 12)
              HeaderF = bytes[offset + 14]
              FooterF = bytes[offset + 15]
              HeaderMarg = bytes[offset + 16]
              FooterMarg = bytes[offset + 17]
              HeaderBold = bytes[offset + 18]
              FooterBold = bytes[offset + 19] }
