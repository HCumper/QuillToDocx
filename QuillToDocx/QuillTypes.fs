namespace QuillToDocx

type Header =
    { HeaderLen: uint16
      Id: string
      TextLen: uint32
      ParaLen: uint16
      FreeLen: uint16
      LayoutLen: uint16 }

type ParaTableHead =
    { Size: uint16
      Gran: uint16
      Used: uint16
      Alloc: uint16 }

type ParaTable =
    { Offset: uint32
      ParaLen: uint16
      Dummy: byte
      LeftMarg: byte
      IndentMarg: byte
      RightMarg: byte
      Justif: byte
      TabTable: byte
      Dummy2: int16 }

type LayoutTable =
    { BottomMarg: byte
      DispMode: byte
      LineGap: byte
      PageLen: byte
      StartPage: byte
      Color: byte
      TopMargin: byte
      Dummy1: byte
      WordCount: uint16
      MaxTabSize: uint16
      TabSize: uint16
      HeaderF: byte
      FooterF: byte
      HeaderMarg: byte
      FooterMarg: byte
      HeaderBold: byte
      FooterBold: byte }

type TabHeader =
    { Entry: byte
      Length: byte }

type TabEntry =
    { Pos: byte
      Type: byte }

type TabTableDef =
    { EntryNumber: byte
      Entries: TabEntry list }

type QuillFile =
    { Header: Header
      IsPcFile: bool
      TextBuffer: byte[]
      ParaHead: ParaTableHead
      Paras: ParaTable[]
      Layout: LayoutTable
      TabTables: TabTableDef list }

type TextStyle =
    { Bold: bool
      Italic: bool
      Underline: bool
      Sub: bool
      Super: bool }

type Inline =
    | RunText of string * TextStyle
    | RunTab of TextStyle
    | RunSoftHyphen of TextStyle

type DecodedParagraph =
    { Index: int
      Descriptor: ParaTable
      Content: Inline list }

type DecodedDocument =
    { HeaderParagraph: DecodedParagraph option
      FooterParagraph: DecodedParagraph option
      BodyParagraphs: DecodedParagraph list
      Layout: LayoutTable }
