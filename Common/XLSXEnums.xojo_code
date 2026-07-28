#tag Module
Protected Module XLSXEnums
	#tag Enum, Name = eCellType, Type = Integer, Flags = &h0, Description = 546865206B696E64206F662076616C75652073746F72656420696E206120776F726B73686565742063656C6C2E0A
		Empty = 0
		  Str = 1
		  Number = 2
		  DateValue = 3
		  Bool = 4
		  FormulaCached = 5
		ErrorVal = 6
	#tag EndEnum

	#tag Enum, Name = eParseError, Type = Integer, Flags = &h0, Description = 4572726F722063617465676F726965732072616973656420696E7369646520616E20584C5358457863657074696F6E2E0A
		NotAZip = 0
		  MissingPart = 1
		  MalformedXML = 2
		  Encrypted = 3
		Unsupported = 4
	#tag EndEnum

	#tag Enum, Name = eOpenMode, Type = Integer, Flags = &h0, Description = 584C53585A6970206261636B656E642073656C6563746F723A204175746F202F204D656D6F7279202F204469736B2E0A
		Auto = 0
		  Memory = 1
		Disk = 2
	#tag EndEnum

	#tag Enum, Name = eAlignH, Type = Integer, Flags = &h0
		General = 0
		  Left = 1
		  Center = 2
		Right = 3
	#tag EndEnum

	#tag Enum, Name = eAlignV, Type = Integer, Flags = &h0
		Bottom = 0
		  Top = 1
		Middle = 2
	#tag EndEnum

	#tag Enum, Name = eBorderStyle, Type = Integer, Flags = &h0
		None = 0
		  Thin = 1
		  Medium = 2
		Thick = 3
	#tag EndEnum

	#tag Note, Name = About
		Enums shared by the XLSX parser.

		eCellType describes how a worksheet cell value should be interpreted.
		  Empty / Str / Number / DateValue / Bool / FormulaCached / ErrorVal

		eParseError tags errors raised by the parser inside an XLSXException:
		  NotAZip       - file does not start with the ZIP magic 50 4B 03 04
		  MissingPart   - required part (e.g. xl/workbook.xml) absent
		  MalformedXML  - an XmlDocument.LoadXml call failed
		  Encrypted     - workbook is in the OLE-wrapped encrypted form
		  Unsupported   - feature out of V1 scope or environment limitation

		eOpenMode picks the XLSXZip backend at Open time:
		  Auto    - Memory on Xojo 2024r3+, else Disk
		  Memory  - parse zip directory in-memory, decompress via MemoryBlock.Decompress
		  Disk    - FolderItem.Unzip into SpecialFolder.Temporary, read parts as FolderItems

		Always reference cases via the namespace, e.g. XLSXEnums.eCellType.Number.
	#tag EndNote

	#tag ViewBehavior
	#tag EndViewBehavior
End Module
#tag EndModule
