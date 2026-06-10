#tag Module
Protected Module ODSWriter
	#tag Method, Flags = &h0, Description = 53657269616C697A652074686520776F726B626F6F6B20746F20616E202E6F64732066696C65206F6E206469736B2E0A
		Sub Save(wb As XLSXWorkbook, file As FolderItem)
		  Build(wb).SaveTo(file)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 53657269616C697A652074686520776F726B626F6F6B20746F20726177202E6F64732062797465732028666F72207468652057656220646F776E6C6F61642070617468292E0A
		Function ToMemoryBlock(wb As XLSXWorkbook) As MemoryBlock
		  Return Build(wb).ToMemoryBlock
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 456D6974206D696D657479706520286669727374202B2073746F7265642C2070657220746865204F44462073706563292C204D4554412D494E462F6D616E69666573742E786D6C2C20616E6420636F6E74656E742E786D6C20696E746F20612053707265616473686565745A69705772697465722E0A
		Private Function Build(wb As XLSXWorkbook) As SpreadsheetZipWriter
		  If wb Is Nil Or wb.SheetCount = 0 Then
		    Raise New XLSXException(XLSXEnums.eParseError.Unsupported, "empty workbook")
		  End If
		  Var zip As New SpreadsheetZipWriter
		  ' The ODF spec requires mimetype FIRST and STORED (uncompressed).
		  zip.AddPart("mimetype", "application/vnd.oasis.opendocument.spreadsheet", True)
		  zip.AddPart("META-INF/manifest.xml", ManifestXml)
		  zip.AddPart("content.xml", ContentXml(wb))
		  Return zip
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 456D6974204D4554412D494E462F6D616E69666573742E786D6C206465636C6172696E6720746865207061636B61676520616E6420636F6E74656E742E786D6C2E0A
		Private Function ManifestXml() As String
		  Var parts() As String
		  parts.Add "<?xml version=""1.0"" encoding=""UTF-8""?>"
		  parts.Add "<manifest:manifest xmlns:manifest=""urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"" manifest:version=""1.2"">"
		  parts.Add "<manifest:file-entry manifest:full-path=""/"" manifest:media-type=""application/vnd.oasis.opendocument.spreadsheet""/>"
		  parts.Add "<manifest:file-entry manifest:full-path=""content.xml"" manifest:media-type=""text/xml""/>"
		  parts.Add "</manifest:manifest>"
		  Return String.FromArray(parts, "")
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 456D697420636F6E74656E742E786D6C3A206175746F6D61746963207374796C6573202864617461207374796C65732066726F6D2074686520776F726B626F6F6B277320666F726D617420636F6465732920706C7573206F6E65207461626C65207065722073686565742E0A
		Private Function ContentXml(wb As XLSXWorkbook) As String
		  ' Collect the distinct format codes; each that converts to an ODS data
		  ' style gets a data style N<k> and a cell style ce<k>.
		  Var codes() As String
		  Var codeToStyle As New Dictionary   ' format code -> cell style name
		  Var styleXml() As String
		  For s As Integer = 1 To wb.SheetCount
		    Var sheet As XLSXSheet = wb.SheetAt(s)
		    For r As Integer = 1 To sheet.RowCount
		      For c As Integer = 1 To sheet.ColCount
		        Var cell As XLSXCell = sheet.CellAt(r, c)
		        If cell.IsEmpty Then Continue
		        Var code As String = XLSXHelpers.WriteFormatCode(cell, wb.Styles)
		        If code = "" Or codes.IndexOf(code) >= 0 Then Continue
		        codes.Add code
		        Var k As Integer = codes.Count
		        Var dataXml As String = DataStyleXml(code, "N" + Str(k))
		        If dataXml <> "" Then
		          styleXml.Add dataXml
		          styleXml.Add "<style:style style:name=""ce" + Str(k) + """ style:family=""table-cell"" style:data-style-name=""N" + Str(k) + """/>"
		          codeToStyle.Value(code) = "ce" + Str(k)
		        End If
		      Next
		    Next
		  Next

		  Var parts() As String
		  parts.Add "<?xml version=""1.0"" encoding=""UTF-8""?>"
		  parts.Add "<office:document-content" _
		    + " xmlns:office=""urn:oasis:names:tc:opendocument:xmlns:office:1.0""" _
		    + " xmlns:style=""urn:oasis:names:tc:opendocument:xmlns:style:1.0""" _
		    + " xmlns:text=""urn:oasis:names:tc:opendocument:xmlns:text:1.0""" _
		    + " xmlns:table=""urn:oasis:names:tc:opendocument:xmlns:table:1.0""" _
		    + " xmlns:number=""urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0""" _
		    + " xmlns:fo=""urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0""" _
		    + " office:version=""1.2"">"
		  parts.Add "<office:automatic-styles>"
		  parts.Add "<style:style style:name=""co1"" style:family=""table-column""><style:table-column-properties style:column-width=""3.0cm""/></style:style>"
		  For Each sx As String In styleXml
		    parts.Add sx
		  Next
		  parts.Add "</office:automatic-styles>"
		  parts.Add "<office:body><office:spreadsheet>"
		  For s As Integer = 1 To wb.SheetCount
		    parts.Add TableXml(wb.SheetAt(s), wb.Styles, codeToStyle)
		  Next
		  parts.Add "</office:spreadsheet></office:body>"
		  parts.Add "</office:document-content>"
		  Return String.FromArray(parts, "")
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 456D6974206F6E65207461626C653A7461626C6520E2809420726F7773206F662063656C6C73207769746820636F76657265642D7461626C652D63656C6C20666F72206D6572676520666F6C6C6F7765727320616E64207370616E2061747472696275746573206F6E206D6572676520616E63686F72732E0A
		Private Function TableXml(sheet As XLSXSheet, styles As XLSXStyles, codeToStyle As Dictionary) As String
		  ' Index merged-range anchors so spans can be emitted on the anchor cell.
		  Var anchors As New Dictionary
		  For i As Integer = 0 To sheet.MergedRangeCount - 1
		    Var rng As XLSXCellRange = sheet.MergedRangeAt(i)
		    If rng <> Nil Then anchors.Value((rng.FirstRow * 16384) + rng.FirstCol) = rng
		  Next

		  Var cols As Integer = Max(1, sheet.ColCount)
		  Var parts() As String
		  parts.Add "<table:table table:name=""" + XLSXHelpers.XmlEscape(sheet.Name) + """>"
		  parts.Add "<table:table-column table:style-name=""co1"" table:number-columns-repeated=""" + Str(cols) + """/>"

		  For r As Integer = 1 To sheet.RowCount
		    Var rowParts() As String
		    rowParts.Add "<table:table-row>"
		    For c As Integer = 1 To cols
		      If sheet.IsCellMergedFollower(r, c) Then
		        rowParts.Add "<table:covered-table-cell/>"
		        Continue
		      End If
		      Var cell As XLSXCell = sheet.CellAt(r, c)
		      Var attrs As String = ""
		      Var key As Integer = (r * 16384) + c
		      If anchors.HasKey(key) Then
		        Var rng As XLSXCellRange = anchors.Value(key)
		        attrs = attrs + " table:number-columns-spanned=""" + Str(rng.LastCol - rng.FirstCol + 1) _
		          + """ table:number-rows-spanned=""" + Str(rng.LastRow - rng.FirstRow + 1) + """"
		      End If
		      If cell.IsEmpty Then
		        rowParts.Add "<table:table-cell" + attrs + "/>"
		        Continue
		      End If
		      Var code As String = XLSXHelpers.WriteFormatCode(cell, styles)
		      If code <> "" And codeToStyle.HasKey(code) Then
		        attrs = " table:style-name=""" + codeToStyle.Value(code).StringValue + """" + attrs
		      End If
		      rowParts.Add CellXml(cell, attrs)
		    Next
		    rowParts.Add "</table:table-row>"
		    parts.Add String.FromArray(rowParts, "")
		  Next

		  parts.Add "</table:table>"
		  Return String.FromArray(parts, "")
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 456D6974206F6E65207461626C653A7461626C652D63656C6C206279206543656C6C547970653A20737472696E672028746578743A7020706572206C696E65292C20666C6F61742C20626F6F6C65616E2C2064617465202849534F29206F722074696D65202850546E486E4D6E53206475726174696F6E292E0A
		Private Function CellXml(cell As XLSXCell, attrs As String) As String
		  Select Case cell.eType
		  Case XLSXEnums.eCellType.Bool
		    Var bv As String = If(cell.BooleanValue, "true", "false")
		    Return "<table:table-cell office:value-type=""boolean"" office:boolean-value=""" + bv + """" + attrs + "/>"

		  Case XLSXEnums.eCellType.DateValue
		    Var d As Double = cell.NumberValue
		    If d < 1.0 Then
		      ' Pure time of day -> ISO 8601 duration.
		      Var total As Integer = Round(d * 86400.0)
		      Var hh As Integer = total \ 3600
		      Var mm As Integer = (total Mod 3600) \ 60
		      Var ss As Integer = total Mod 60
		      Return "<table:table-cell office:value-type=""time"" office:time-value=""PT" _
		        + Str(hh) + "H" + Str(mm) + "M" + Str(ss) + "S""" + attrs + "/>"
		    End If
		    Var dt As DateTime = XLSXFormatter.ExcelSerialToDateTime(d)
		    Var iso As String = dt.SQLDate
		    If d - Floor(d) > 0.000001 Then
		      iso = iso + "T" + Pad2(dt.Hour) + ":" + Pad2(dt.Minute) + ":" + Pad2(dt.Second)
		    End If
		    Return "<table:table-cell office:value-type=""date"" office:date-value=""" + iso + """" + attrs + "/>"

		  Case XLSXEnums.eCellType.Number, XLSXEnums.eCellType.FormulaCached
		    Return "<table:table-cell office:value-type=""float"" office:value=""" _
		      + XLSXHelpers.XmlEscape(cell.RawString) + """" + attrs + "/>"

		  Else
		    ' Str, ErrorVal: inline text, one <text:p> per line.
		    Var lines() As String = cell.RawString.ReplaceAll(Chr(13) + Chr(10), Chr(10)).ReplaceAll(Chr(13), Chr(10)).Split(Chr(10))
		    Var body() As String
		    For Each ln As String In lines
		      body.Add "<text:p>" + XLSXHelpers.XmlEscape(ln) + "</text:p>"
		    Next
		    Return "<table:table-cell office:value-type=""string""" + attrs + ">" + String.FromArray(body, "") + "</table:table-cell>"
		  End Select
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 436F6E76657274206F6E6520666F726D617420636F646520696E746F20616E204F4453206E756D6265723A2A2D7374796C6520747265652028646174652F74696D652C2070657263656E746167652C2063757272656E63792C20706C61696E206E756D626572292E20456D70747920666F7220756E737570706F7274656420636F6465732E0A
		Private Function DataStyleXml(code As String, styleName As String) As String
		  ' Convert one Excel-style format code into an ODS <number:*-style> tree.
		  ' Returns "" for codes outside the supported subset (cell saves unstyled).
		  Var first As String = code
		  Var semi As Integer = first.IndexOf(";")
		  If semi >= 0 Then first = first.Left(semi)
		  If first = "" Then Return ""
		  If XLSXFormatter.IsDateFormatCode(first) Then Return DateStyleXml(first, styleName)
		  If first.IndexOf("%") >= 0 Then Return PercentageStyleXml(first, styleName)
		  If first.IndexOf("[$") >= 0 Then Return CurrencyStyleXml(first, styleName)
		  Return NumberStyleXml(first, styleName)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 54686520736861726564203C6E756D6265723A6E756D6265722F3E20656C656D656E743A20646563696D616C20706C616365732066726F6D2074686520636F64652773206672616374696F6E206469676974732C2067726F7570696E672066726F6D20612074686F7573616E647320736570617261746F722E0A
		Private Function NumberElementXml(c As String) As String
		  ' The shared <number:number .../> element: decimals from the digits after
		  ' the decimal point, grouping from a thousands separator in the code.
		  If c.IndexOf("0") < 0 And c.IndexOf("#") < 0 Then Return ""
		  Var decimals As Integer = 0
		  Var dot As Integer = c.IndexOf(".")
		  If dot >= 0 Then
		    Var i As Integer = dot + 1
		    While i < c.Length
		      Var ch As String = c.Middle(i, 1)
		      If ch = "0" Or ch = "#" Then
		        decimals = decimals + 1
		        i = i + 1
		      Else
		        Exit
		      End If
		    Wend
		  End If
		  Var grouping As String = ""
		  If c.IndexOf("#,#") >= 0 Or c.IndexOf("0,0") >= 0 Then grouping = " number:grouping=""true"""
		  Return "<number:number number:decimal-places=""" + Str(decimals) _
		    + """ number:min-integer-digits=""1""" + grouping + "/>"
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 57726170204E756D626572456C656D656E74586D6C20696E2061203C6E756D6265723A6E756D6265722D7374796C653E2E0A
		Private Function NumberStyleXml(c As String, styleName As String) As String
		  Var numEl As String = NumberElementXml(c)
		  If numEl = "" Then Return ""
		  Return "<number:number-style style:name=""" + styleName + """>" + numEl + "</number:number-style>"
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 456D69742061203C6E756D6265723A70657263656E746167652D7374796C653E20E28094206E756D62657220656C656D656E7420706C75732061206C69746572616C202520746578742E0A
		Private Function PercentageStyleXml(c As String, styleName As String) As String
		  Var numEl As String = NumberElementXml(c)
		  If numEl = "" Then numEl = "<number:number number:decimal-places=""0"" number:min-integer-digits=""1""/>"
		  Return "<number:percentage-style style:name=""" + styleName + """>" + numEl _
		    + "<number:text>%</number:text></number:percentage-style>"
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 456D69742061203C6E756D6265723A63757272656E63792D7374796C653E2077697468207468652073796D626F6C2066726F6D20746865205B2453594D2D6C6F63616C655D2074616720706F736974696F6E6564206265666F7265206F7220616674657220746865206E756D6265722E0A
		Private Function CurrencyStyleXml(c As String, styleName As String) As String
		  ' Extract the symbol from the [$SYM-locale] tag; emit it before or after
		  ' the number depending on where the tag sits relative to the digits.
		  Var tagStart As Integer = c.IndexOf("[$")
		  If tagStart < 0 Then Return NumberStyleXml(c, styleName)
		  Var tagEnd As Integer = c.IndexOf(tagStart, "]")
		  If tagEnd < 0 Then Return NumberStyleXml(c, styleName)
		  Var inner As String = c.Middle(tagStart + 2, tagEnd - tagStart - 2)
		  Var dash As Integer = inner.IndexOf("-")
		  Var sym As String = If(dash >= 0, inner.Left(dash), inner)
		  If sym = "" Then Return NumberStyleXml(c, styleName)

		  Var numEl As String = NumberElementXml(c)
		  If numEl = "" Then numEl = "<number:number number:decimal-places=""2"" number:min-integer-digits=""1""/>"
		  Var symEl As String = "<number:currency-symbol>" + XLSXHelpers.XmlEscape(sym) + "</number:currency-symbol>"

		  ' Symbol is a prefix when the tag appears before the first digit placeholder.
		  Var firstDigit As Integer = c.IndexOf("0")
		  Var firstHash As Integer = c.IndexOf("#")
		  If firstHash >= 0 And (firstDigit < 0 Or firstHash < firstDigit) Then firstDigit = firstHash
		  Var body As String
		  If firstDigit < 0 Or tagStart < firstDigit Then
		    body = symEl + numEl
		  Else
		    body = numEl + symEl
		  End If
		  Return "<number:currency-style style:name=""" + styleName + """>" + body + "</number:currency-style>"
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 546F6B656E697A65206120646174652F74696D6520666F726D617420636F64652028792F6D2F642F682F732072756E732C20414D2F504D2C206C69746572616C733B206D6F6E74682D76732D6D696E7574657320646973616D626967756174696F6E2920696E746F203C6E756D6265723A646174652D7374796C653E206F72203C6E756D6265723A74696D652D7374796C653E2E0A
		Private Function DateStyleXml(c As String, styleName As String) As String
		  ' Tokenize the date code into letter runs (y/m/d/h/s), AM/PM markers, and
		  ' literal text, then emit the matching <number:date-style> (or time-style
		  ' when only time tokens are present).
		  Var kinds() As String
		  Var texts() As String
		  Var i As Integer = 0
		  Var n As Integer = c.Length
		  While i < n
		    Var ch As String = c.Middle(i, 1)
		    Var lc As String = ch.Lowercase
		    If lc = "a" And c.Middle(i, 5).Lowercase = "am/pm" Then
		      kinds.Add "ap"
		      texts.Add ""
		      i = i + 5
		    ElseIf lc = "y" Or lc = "m" Or lc = "d" Or lc = "h" Or lc = "s" Then
		      Var runLen As Integer = 1
		      While i + runLen < n And c.Middle(i + runLen, 1).Lowercase = lc
		        runLen = runLen + 1
		      Wend
		      kinds.Add lc
		      texts.Add c.Middle(i, runLen)
		      i = i + runLen
		    Else
		      Var litStart As Integer = i
		      While i < n
		        Var c2 As String = c.Middle(i, 1).Lowercase
		        If c2 = "y" Or c2 = "m" Or c2 = "d" Or c2 = "h" Or c2 = "s" Then Exit
		        If c2 = "a" And c.Middle(i, 5).Lowercase = "am/pm" Then Exit
		        i = i + 1
		      Wend
		      kinds.Add "lit"
		      texts.Add c.Middle(litStart, i - litStart)
		    End If
		  Wend

		  ' Disambiguate m: minutes when the nearest letter token before is h,
		  ' or the nearest after is s / AM-PM; otherwise month.
		  For t As Integer = 0 To kinds.LastIndex
		    If kinds(t) <> "m" Then Continue
		    Var isMin As Boolean = False
		    For p As Integer = t - 1 DownTo 0
		      If kinds(p) = "lit" Then Continue
		      If kinds(p) = "h" Then isMin = True
		      Exit
		    Next
		    If Not isMin Then
		      For q As Integer = t + 1 To kinds.LastIndex
		        If kinds(q) = "lit" Then Continue
		        If kinds(q) = "s" Or kinds(q) = "ap" Then isMin = True
		        Exit
		      Next
		    End If
		    If isMin Then kinds(t) = "min"
		  Next

		  Var hasDatePart As Boolean = False
		  For Each k As String In kinds
		    If k = "y" Or k = "m" Or k = "d" Then hasDatePart = True
		  Next

		  Var body() As String
		  For t As Integer = 0 To kinds.LastIndex
		    Var isLong As String = If(texts(t).Length >= 2, " number:style=""long""", "")
		    Select Case kinds(t)
		    Case "y"
		      body.Add "<number:year" + If(texts(t).Length >= 4, " number:style=""long""", "") + "/>"
		    Case "m"
		      body.Add "<number:month" + isLong + "/>"
		    Case "d"
		      body.Add "<number:day" + isLong + "/>"
		    Case "h"
		      body.Add "<number:hours" + isLong + "/>"
		    Case "min"
		      body.Add "<number:minutes" + isLong + "/>"
		    Case "s"
		      body.Add "<number:seconds" + isLong + "/>"
		    Case "ap"
		      body.Add "<number:am-pm/>"
		    Case "lit"
		      If texts(t) <> "" Then
		        body.Add "<number:text>" + XLSXHelpers.XmlEscape(texts(t)) + "</number:text>"
		      End If
		    End Select
		  Next
		  If body.Count = 0 Then Return ""

		  Var elName As String = If(hasDatePart, "number:date-style", "number:time-style")
		  Return "<" + elName + " style:name=""" + styleName + """>" + String.FromArray(body, "") + "</" + elName + ">"
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 54776F2D6469676974207A65726F2D70616464656420646563696D616C20737472696E6720666F7220302E2E39392E0A
		Private Function Pad2(v As Integer) As String
		  If v < 10 Then Return "0" + Str(v)
		  Return Str(v)
		End Function
	#tag EndMethod

	#tag Note, Name = About
		Serializes the shared XLSXWorkbook model into an .ods file (OpenDocument
		Spreadsheet) — the write-side mirror of ODSReader.

		Entry points:
		  Save(wb, file)        - write to disk
		  ToMemoryBlock(wb)     - raw bytes (Web download)

		Archive layout (via SpreadsheetZipWriter): mimetype FIRST and STORED as the
		ODF packaging spec requires, then META-INF/manifest.xml and content.xml.

		Number formats: each distinct format code in the workbook is converted back
		into an ODS data-style element tree — the exact inverse of what ODSReader
		parses: date/time codes -> <number:date-style>/<number:time-style> token
		elements (with the m = month-vs-minutes disambiguation), percent codes ->
		<number:percentage-style>, [$SYM-...] currency tags ->
		<number:currency-style> with the symbol positioned before/after the number,
		plain numeric codes -> <number:number-style> with decimals + grouping.
		Codes outside that subset save unstyled (value preserved, format dropped).

		Cells: office:value-type string / float / boolean / date / time; date
		serials convert back to ISO dates (or PTnHnMnS durations when < 1 day);
		multi-line text becomes one <text:p> per line. Merged ranges emit
		column/row-spanned attributes on the anchor plus <table:covered-table-cell>
		for followers — again the inverse of the reader.

		Out of scope: cell colors / fonts / borders, formulas (cached value saved),
		charts, images, conditional formatting.
	#tag EndNote

	#tag ViewBehavior
	#tag EndViewBehavior
End Module
#tag EndModule
