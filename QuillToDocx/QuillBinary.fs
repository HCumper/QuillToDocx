namespace QuillToDocx

open System
open System.Text

module QuillBinary =

    let private readU16BE (bytes: byte[]) offset =
        (uint16 bytes[offset] <<< 8) ||| uint16 bytes[offset + 1]

    let private readU32BE (bytes: byte[]) offset =
        (uint32 bytes[offset] <<< 24) |||
        (uint32 bytes[offset + 1] <<< 16) |||
        (uint32 bytes[offset + 2] <<< 8) |||
        uint32 bytes[offset + 3]

    let readAscii (bytes: byte[]) offset len =
        Encoding.ASCII.GetString(bytes, offset, len)

    let readHeader (bytes: byte[]) =
        { HeaderLen = readU16BE bytes 0
          Id = readAscii bytes 2 8
          TextLen = readU32BE bytes 10
          ParaLen = readU16BE bytes 14
          FreeLen = readU16BE bytes 16
          LayoutLen = readU16BE bytes 18 }

    let readParaHead (bytes: byte[]) offset =
        { Size = readU16BE bytes offset
          Gran = readU16BE bytes (offset + 2)
          Used = readU16BE bytes (offset + 4)
          Alloc = readU16BE bytes (offset + 6) }

    let readPara (bytes: byte[]) offset =
        { Offset = readU32BE bytes offset
          ParaLen = readU16BE bytes (offset + 4)
          Dummy = bytes[offset + 6]
          LeftMarg = bytes[offset + 7]
          IndentMarg = bytes[offset + 8]
          RightMarg = bytes[offset + 9]
          Justif = bytes[offset + 10]
          TabTable = bytes[offset + 11]
          Dummy2 = int16 (readU16BE bytes (offset + 12)) }

    let readLayout (bytes: byte[]) offset =
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
