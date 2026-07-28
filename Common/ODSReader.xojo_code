#tag Module
Protected Module ODSReader
	#tag Method, Flags = &h0, Description = 4F70656E20616E202E6F64732066726F6D206120466F6C6465724974656D2E2052657475726E73207468652073686172656420584C5358576F726B626F6F6B206D6F64656C2E205468726561647320654F70656E4D6F646520746F20584C53585A69702E0A
		Function Open(file As FolderItem, mode As XLSXEnums.eOpenMode = XLSXEnums.eOpenMode.Auto) As XLSXWorkbook
		  If file Is Nil Or Not file.Exists Then
		    Raise New XLSXException(XLSXEnums.eParseError.MissingPart, "file does not exist")
		  End If
		  Var t0 As Double = System.Microseconds
		  Var zip As XLSXZip = XLSXZip.Open(file, mode)
		  Var t1 As Double = System.Microseconds
		  Var wb As XLSXWorkbook = OpenFromZip(zip, file.Name)
		  Var t2 As Double = System.Microseconds
		  wb.OpenMode = zip.Mode
		  wb.ZipMicroseconds = t1 - t0
		  wb.XmlMicroseconds = t2 - t1
		  Return wb
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 4F70656E20616E202E6F64732066726F6D2072617720627974657320285765622075706C6F6164292E2052657475726E73207468652073686172656420584C5358576F726B626F6F6B206D6F64656C2E0A
		Function Open(data As MemoryBlock, sourceName As String, mode As XLSXEnums.eOpenMode = XLSXEnums.eOpenMode.Auto) As XLSXWorkbook
		  If data Is Nil Or data.Size = 0 Then
		    Raise New XLSXException(XLSXEnums.eParseError.NotAZip, "empty data")
		  End If
		  Var t0 As Double = System.Microseconds
		  Var zip As XLSXZip = XLSXZip.Open(data, mode)
		  Var t1 As Double = System.Microseconds
		  Var wb As XLSXWorkbook = OpenFromZip(zip, sourceName)
		  Var t2 As Double = System.Microseconds
		  wb.OpenMode = zip.Mode
		  wb.ZipMicroseconds = t1 - t0
		  wb.XmlMicroseconds = t2 - t1
		  Return wb
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 506172736520636F6E74656E742E786D6C2028616E64206F7074696F6E616C207374796C65732E786D6C2920696E746F20616E20584C5358576F726B626F6F6B3A207265736F6C7665207374796C65732C207468656E206275696C64206F6E6520584C5358536865657420706572207461626C653A7461626C652E0A
		Private Function OpenFromZip(zip As XLSXZip, sourceName As String) As XLSXWorkbook
		  If Not zip.HasPart("content.xml") Then
		    Raise New XLSXException(XLSXEnums.eParseError.MissingPart, "content.xml")
		  End If

		  Var wb As New XLSXWorkbook(sourceName)

		  Var contentXml As String = zip.ReadPart("content.xml")
		  Var doc As New XmlDocument
		  Try
		    doc.LoadXml(contentXml)
		  Catch
		    Raise New XLSXException(XLSXEnums.eParseError.MalformedXML, "content.xml")
		  End Try

		  ' Resolve styles: cellStyleName -> dataStyleName, dataStyleName -> formatCode,
		  ' columnStyleName -> width (pt), rowStyleName -> height (pt).
		  Var cellStyles As New Dictionary
		  Var dataStyles As New Dictionary
		  Var colStyleWidths As New Dictionary
		  Var rowStyleHeights As New Dictionary
		  ParseStyles(doc, cellStyles, dataStyles, colStyleWidths, rowStyleHeights)

		  ' styles.xml may carry additional named / data styles.
		  Var stylesXml As String = zip.ReadPart("styles.xml")
		  If stylesXml <> "" Then
		    Var sdoc As New XmlDocument
		    Try
		      sdoc.LoadXml(stylesXml)
		      ParseStyles(sdoc, cellStyles, dataStyles, colStyleWidths, rowStyleHeights)
		    Catch
		      ' styles.xml is optional polish; ignore a malformed one.
		    End Try
		  End If

		  Var tables As XmlNodeList = doc.Xql("//*[local-name()='spreadsheet']/*[local-name()='table']")
		  For i As Integer = 0 To tables.Length - 1
		    Var sheet As XLSXSheet = ParseTable(tables.Item(i), i + 1, cellStyles, dataStyles, colStyleWidths, rowStyleHeights)
		    wb.AddSheet(sheet)
		  Next

		  Return wb
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 4275696C6420616E20584C535853686565742066726F6D206F6E65207461626C653A7461626C652C20686F6E6F72696E67206E756D6265722D636F6C756D6E732F726F77732D726570656174656420616E6420636F6C756D6E2F726F77207370616E7320286D65726765642063656C6C73292E0A
		Private Function ParseTable(tableNode As XmlNode, tabIndex As Integer, cellStyles As Dictionary, dataStyles As Dictionary, colStyleWidths As Dictionary, rowStyleHeights As Dictionary) As XLSXSheet
		  Var name As String = tableNode.GetAttribute("table:name")
		  Var sheet As New XLSXSheet(name, tabIndex)

		  ' Column widths: <table:table-column> elements (in document order, honoring
		  ' number-columns-repeated) carry a style whose column-width we resolved.
		  Var colIndex As Integer = 0
		  Var colNode As XmlNode = tableNode.FirstChild
		  While colNode <> Nil
		    If LocalName(colNode) = "table-column" Then
		      Var rep As Integer = IntAttr(colNode, "table:number-columns-repeated", 1)
		      If rep < 1 Then rep = 1
		      Var styleName As String = colNode.GetAttribute("table:style-name")
		      Var wpt As Double = 0.0
		      If styleName <> "" And colStyleWidths.HasKey(styleName) Then wpt = colStyleWidths.Value(styleName)
		      If wpt <= 0 Then
		        colIndex = colIndex + rep
		      Else
		        Var reps As Integer = Min(rep, 4096)
		        For k As Integer = 1 To reps
		          colIndex = colIndex + 1
		          sheet.SetColumnWidth(colIndex, wpt)
		        Next
		        If rep > reps Then colIndex = colIndex + (rep - reps)
		      End If
		    End If
		    colNode = colNode.NextSibling
		  Wend

		  ' Descendant axis catches rows wrapped in <table:table-header-rows> / row groups.
		  Var rows As XmlNodeList = tableNode.Xql(".//*[local-name()='table-row']")
		  Var rowIdx As Integer = 0
		  For ri As Integer = 0 To rows.Length - 1
		    Var rowNode As XmlNode = rows.Item(ri)
		    Var rowRepeat As Integer = IntAttr(rowNode, "table:number-rows-repeated", 1)
		    If rowRepeat < 1 Then rowRepeat = 1
		    ' Custom row height (points) from the row's style, when present.
		    Var rowHpt As Double = 0.0
		    Var rowStyleName As String = rowNode.GetAttribute("table:style-name")
		    If rowStyleName <> "" And rowStyleHeights.HasKey(rowStyleName) Then rowHpt = rowStyleHeights.Value(rowStyleName)

		    ' Build this row's cell plan once.
		    Var planCols() As Integer
		    Var planCells() As XLSXCell
		    Var planColSpan() As Integer
		    Var planRowSpan() As Integer
		    Var colIdx As Integer = 0
		    Var hasContent As Boolean = False

		    ' Iterate the row's children directly (in document order) so we don't
		    ' depend on an XPath 'or' predicate, and so table-cell / covered-table-cell
		    ' positions stay correct for column counting.
		    Var cellNode As XmlNode = rowNode.FirstChild
		    While cellNode <> Nil
		      Var ln As String = LocalName(cellNode)
		      If ln = "table-cell" Or ln = "covered-table-cell" Then
		        Var isCovered As Boolean = (ln = "covered-table-cell")
		        Var colRepeat As Integer = IntAttr(cellNode, "table:number-columns-repeated", 1)
		        If colRepeat < 1 Then colRepeat = 1
		        Var colSpan As Integer = IntAttr(cellNode, "table:number-columns-spanned", 1)
		        Var rowSpan As Integer = IntAttr(cellNode, "table:number-rows-spanned", 1)

		        Var cell As XLSXCell
		        If Not isCovered Then cell = MakeCell(cellNode, cellStyles, dataStyles)

		        If cell Is Nil Then
		          ' Empty or covered cell: advance the column counter arithmetically.
		          colIdx = colIdx + colRepeat
		        Else
		          Var reps As Integer = Min(colRepeat, 4096)
		          For k As Integer = 1 To reps
		            colIdx = colIdx + 1
		            planCols.Add colIdx
		            planCells.Add cell
		            planColSpan.Add colSpan
		            planRowSpan.Add rowSpan
		            hasContent = True
		          Next
		          If colRepeat > reps Then colIdx = colIdx + (colRepeat - reps)
		        End If
		      End If

		      cellNode = cellNode.NextSibling
		    Wend

		    If hasContent Then
		      Var emitRows As Integer = Min(rowRepeat, 4096)
		      For iter As Integer = 0 To emitRows - 1
		        Var rr As Integer = rowIdx + 1 + iter
		        If rowHpt > 0 Then sheet.SetRowHeight(rr, rowHpt)
		        For p As Integer = 0 To planCols.LastIndex
		          sheet.PutCell(rr, planCols(p), planCells(p))
		          If planColSpan(p) > 1 Or planRowSpan(p) > 1 Then
		            sheet.AddMergedRange(rr, planCols(p), rr + planRowSpan(p) - 1, planCols(p) + planColSpan(p) - 1)
		          End If
		        Next
		      Next
		    End If

		    rowIdx = rowIdx + rowRepeat
		  Next

		  Return sheet
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 5475726E206F6E65207461626C653A7461626C652D63656C6C20696E746F20616E20584C535843656C6C20286F72204E696C20696620656D707479292E204D617073206F66666963653A76616C75652D7479706520746F206543656C6C5479706520616E64207265736F6C7665732074686520666F726D617420636F64652E0A
		Private Function MakeCell(cellNode As XmlNode, cellStyles As Dictionary, dataStyles As Dictionary) As XLSXCell
		  ' Returns a populated XLSXCell, or Nil for a truly-empty cell.
		  Var valueType As String = cellNode.GetAttribute("office:value-type")
		  Var styleName As String = cellNode.GetAttribute("table:style-name")
		  Var formatCode As String = ResolveFormat(styleName, cellStyles, dataStyles)
		  Var hasFormula As Boolean = cellNode.GetAttribute("table:formula") <> ""

		  Var rawValue As String = ""
		  Var eType As XLSXEnums.eCellType = XLSXEnums.eCellType.Empty

		  Select Case valueType
		  Case "float", "currency"
		    rawValue = cellNode.GetAttribute("office:value")
		    If rawValue = "" Then Return Nil
		    eType = If(hasFormula, XLSXEnums.eCellType.FormulaCached, XLSXEnums.eCellType.Number)

		  Case "percentage"
		    rawValue = cellNode.GetAttribute("office:value")
		    If rawValue = "" Then Return Nil
		    eType = If(hasFormula, XLSXEnums.eCellType.FormulaCached, XLSXEnums.eCellType.Number)
		    If formatCode = "" Then formatCode = "0%"

		  Case "date"
		    Var dv As String = cellNode.GetAttribute("office:date-value")
		    If dv = "" Then Return Nil
		    rawValue = Str(ISODateToSerial(dv))
		    eType = XLSXEnums.eCellType.DateValue
		    If formatCode = "" Then
		      formatCode = If(dv.IndexOf("T") >= 0, "yyyy-mm-dd hh:mm", "yyyy-mm-dd")
		    End If

		  Case "time"
		    Var tv As String = cellNode.GetAttribute("office:time-value")
		    If tv = "" Then Return Nil
		    rawValue = Str(ISOTimeToFraction(tv))
		    eType = XLSXEnums.eCellType.DateValue
		    If formatCode = "" Then formatCode = "hh:mm:ss"

		  Case "boolean"
		    Var bv As String = cellNode.GetAttribute("office:boolean-value")
		    rawValue = If(bv = "true", "1", "0")
		    eType = XLSXEnums.eCellType.Bool

		  Else
		    ' string or untyped — text lives in <text:p> children.
		    rawValue = ParagraphText(cellNode)
		    If rawValue = "" Then Return Nil
		    eType = XLSXEnums.eCellType.Str
		  End Select

		  Var cell As New XLSXCell(eType, rawValue, -1)
		  cell.FormatCode = formatCode
		  ' Preserve the formula (ODF -> A1) for numeric formula cells so it survives
		  ' a round-trip, matching the XLSX reader. (Date/bool/string formula results
		  ' keep their computed value only.)
		  If hasFormula And eType = XLSXEnums.eCellType.FormulaCached Then
		    cell.Formula = XLSXHelpers.OdfFormulaToA1(cellNode.GetAttribute("table:formula"))
		  End If
		  Return cell
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 5265736F6C766520612063656C6C2773207461626C653A7374796C652D6E616D6520746F206120666F726D61742D636F646520737472696E6720766961207468652063656C6C5374796C65202D3E20646174615374796C65202D3E20666F726D61742D636F6465206D6170732E0A
		Private Function ResolveFormat(styleName As String, cellStyles As Dictionary, dataStyles As Dictionary) As String
		  If styleName = "" Then Return ""
		  If Not cellStyles.HasKey(styleName) Then Return ""
		  Var dataName As String = cellStyles.Value(styleName)
		  If dataStyles.HasKey(dataName) Then Return dataStyles.Value(dataName)
		  Return ""
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 4275696C642063656C6C5374796C654E616D65202D3E20646174615374796C654E616D6520616E6420646174615374796C654E616D65202D3E20666F726D61742D636F6465206D6170732066726F6D206120636F6E74656E742E786D6C206F72207374796C65732E786D6C20646F63756D656E742E0A
		Private Sub ParseStyles(doc As XmlDocument, cellStyles As Dictionary, dataStyles As Dictionary, colStyleWidths As Dictionary, rowStyleHeights As Dictionary)
		  ' Data styles: number / date / time / currency / percentage styles -> format code.
		  ' One query per type (avoids relying on an XPath 'or' predicate).
		  Var typeNames() As String = Array("number-style", "date-style", "time-style", "currency-style", "percentage-style")
		  For Each tn As String In typeNames
		    Var dataNodes As XmlNodeList = doc.Xql("//*[local-name()='" + tn + "']")
		    For i As Integer = 0 To dataNodes.Length - 1
		      Var n As XmlNode = dataNodes.Item(i)
		      Var styName As String = n.GetAttribute("style:name")
		      If styName = "" Then Continue
		      dataStyles.Value(styName) = BuildFormatCode(n)
		    Next
		  Next

		  ' <style:style> entries: table-cell -> data-style-name; table-column ->
		  ' column-width (points); table-row -> row-height (points).
		  Var styleNodes As XmlNodeList = doc.Xql("//*[local-name()='style']")
		  For i As Integer = 0 To styleNodes.Length - 1
		    Var s As XmlNode = styleNodes.Item(i)
		    Var nm As String = s.GetAttribute("style:name")
		    If nm = "" Then Continue
		    Select Case s.GetAttribute("style:family")
		    Case "table-cell"
		      Var dn As String = s.GetAttribute("style:data-style-name")
		      If dn <> "" Then cellStyles.Value(nm) = dn
		    Case "table-column"
		      Var props As XmlNodeList = s.Xql("./*[local-name()='table-column-properties']")
		      If props.Length > 0 Then
		        Var w As String = props.Item(0).GetAttribute("style:column-width")
		        If w <> "" Then colStyleWidths.Value(nm) = XLSXHelpers.OdsLengthToPoints(w)
		      End If
		    Case "table-row"
		      Var props As XmlNodeList = s.Xql("./*[local-name()='table-row-properties']")
		      If props.Length > 0 Then
		        Var h As String = props.Item(0).GetAttribute("style:row-height")
		        If h <> "" Then rowStyleHeights.Value(nm) = XLSXHelpers.OdsLengthToPoints(h)
		      End If
		    End Select
		  Next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 44697370617463682061203C6E756D6265723A2A2D7374796C653E206E6F646520746F20746865206D61746368696E6720646174652F6E756D6265722F70657263656E746167652F63757272656E637920636F6465206275696C6465722E0A
		Private Function BuildFormatCode(styleNode As XmlNode) As String
		  Select Case LocalName(styleNode)
		  Case "date-style", "time-style"
		    Return BuildDateCode(styleNode)
		  Case "number-style"
		    Return BuildNumberCode(styleNode)
		  Case "percentage-style"
		    Return BuildPercentageCode(styleNode)
		  Case "currency-style"
		    Return BuildCurrencyCode(styleNode)
		  End Select
		  Return ""
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 436F6E766572742061203C6E756D6265723A646174652D7374796C653E202F203C6E756D6265723A74696D652D7374796C653E20656C656D656E74207472656520696E746F206120646174652F74696D6520666F726D61742D636F646520737472696E672028652E672E2064642F6D6D2F79797979292E0A
		Private Function BuildDateCode(styleNode As XmlNode) As String
		  Var parts() As String
		  Var child As XmlNode = styleNode.FirstChild
		  While child <> Nil
		    Var isLong As Boolean = (child.GetAttribute("number:style") = "long")
		    Var isTextual As Boolean = (child.GetAttribute("number:textual") = "true")
		    Select Case LocalName(child)
		    Case "year"
		      parts.Add If(isLong, "yyyy", "yy")
		    Case "month"
		      If isTextual Then
		        parts.Add If(isLong, "mmmm", "mmm")
		      Else
		        parts.Add If(isLong, "mm", "m")
		      End If
		    Case "day"
		      parts.Add If(isLong, "dd", "d")
		    Case "day-of-week"
		      parts.Add If(isLong, "dddd", "ddd")
		    Case "hours"
		      parts.Add If(isLong, "hh", "h")
		    Case "minutes"
		      parts.Add If(isLong, "mm", "m")
		    Case "seconds"
		      parts.Add If(isLong, "ss", "s")
		    Case "am-pm"
		      parts.Add "AM/PM"
		    Case "text"
		      parts.Add CollectText(child)
		    End Select
		    child = child.NextSibling
		  Wend
		  Return String.FromArray(parts, "")
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 436F6E766572742061203C6E756D6265723A6E756D6265722D7374796C653E207472656520696E746F2061206E756D6572696320666F726D617420636F64652028707265666978202B20636F7265202B20737566666978292E0A
		Private Function BuildNumberCode(styleNode As XmlNode) As String
		  Var prefix As String = ""
		  Var suffix As String = ""
		  Var core As String = ""
		  Var child As XmlNode = styleNode.FirstChild
		  While child <> Nil
		    Select Case LocalName(child)
		    Case "number"
		      Var dec As Integer = IntAttr(child, "number:decimal-places", 0)
		      Var grouping As Boolean = (child.GetAttribute("number:grouping") = "true")
		      core = NumberCore(dec, grouping)
		    Case "scientific-number"
		      Var dec As Integer = IntAttr(child, "number:decimal-places", 2)
		      core = ScientificCore(dec)
		    Case "text"
		      Var t As String = CollectText(child)
		      If core = "" Then
		        prefix = prefix + t
		      Else
		        suffix = suffix + t
		      End If
		    End Select
		    child = child.NextSibling
		  Wend
		  Return prefix + core + suffix
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 436F6E766572742061203C6E756D6265723A70657263656E746167652D7374796C653E207472656520696E746F2022302522206F722022302E30302522202874686520666F726D7320584C5358466F726D617474657220756E6465727374616E6473292E0A
		Private Function BuildPercentageCode(styleNode As XmlNode) As String
		  ' Collapse to the two percentage forms XLSXFormatter understands.
		  Var dec As Integer = 0
		  Var child As XmlNode = styleNode.FirstChild
		  While child <> Nil
		    If LocalName(child) = "number" Then
		      dec = IntAttr(child, "number:decimal-places", 0)
		    End If
		    child = child.NextSibling
		  Wend
		  If dec > 0 Then Return "0.00%"
		  Return "0%"
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 436F6E766572742061203C6E756D6265723A63757272656E63792D7374796C653E207472656520696E746F206120636F64652077697468207468652073796D626F6C2061732061205B2453594D2D305D2074616720736F20584C5358466F726D61747465722072656E646572732069742E0A
		Private Function BuildCurrencyCode(styleNode As XmlNode) As String
		  ' Emit the symbol as a [$SYM-0] tag so XLSXFormatter's currency-tag path
		  ' renders it (prefix or suffix based on position relative to the number).
		  Var prefix As String = ""
		  Var suffix As String = ""
		  Var core As String = ""
		  Var child As XmlNode = styleNode.FirstChild
		  While child <> Nil
		    Select Case LocalName(child)
		    Case "number"
		      Var dec As Integer = IntAttr(child, "number:decimal-places", 2)
		      Var grouping As Boolean = (child.GetAttribute("number:grouping") = "true")
		      core = NumberCore(dec, grouping)
		    Case "currency-symbol"
		      Var sym As String = CollectText(child)
		      Var tag As String = "[$" + sym + "-0]"
		      If core = "" Then
		        prefix = prefix + tag
		      Else
		        suffix = suffix + tag
		      End If
		    Case "text"
		      Var t As String = CollectText(child)
		      If core = "" Then
		        prefix = prefix + t
		      Else
		        suffix = suffix + t
		      End If
		    End Select
		    child = child.NextSibling
		  Wend
		  Return prefix + core + suffix
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 4275696C6420746865206E756D6572696320636F7265206F66206120666F726D617420636F64652066726F6D206120646563696D616C2D706C6163657320636F756E7420616E6420612067726F7570696E6720666C61672028652E672E2022232C2323302E303022292E0A
		Private Function NumberCore(decimals As Integer, grouping As Boolean) As String
		  Var intPart As String = If(grouping, "#,##0", "0")
		  If decimals <= 0 Then Return intPart
		  Var zeros As String = ""
		  For i As Integer = 1 To decimals
		    zeros = zeros + "0"
		  Next
		  Return intPart + "." + zeros
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 4275696C64206120736369656E746966696320666F726D617420636F7265206C696B652022302E3030452B3030222066726F6D206120646563696D616C2D706C6163657320636F756E742E0A
		Private Function ScientificCore(decimals As Integer) As String
		  If decimals <= 0 Then Return "0E+00"
		  Var zeros As String = ""
		  For i As Integer = 1 To decimals
		    zeros = zeros + "0"
		  Next
		  Return "0." + zeros + "E+00"
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 436F6E636174656E617465207468652074657874206F6620612063656C6C2773203C746578743A703E206368696C6472656E2C206A6F696E696E67206D756C7469706C65207061726167726170687320776974682061206E65776C696E652E0A
		Private Function ParagraphText(cellNode As XmlNode) As String
		  Var ps As XmlNodeList = cellNode.Xql("./*[local-name()='p']")
		  Var parts() As String
		  For i As Integer = 0 To ps.Length - 1
		    parts.Add CollectText(ps.Item(i))
		  Next
		  Return String.FromArray(parts, EndOfLine)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 5265637572736976656C7920636F6C6C6563742061206E6F64652773207465787420636F6E74656E743B203C746578743A733E2F3C746578743A7461623E2F3C746578743A6C696E652D627265616B3E206265636F6D652073706163652F7461622F6E65776C696E652E0A
		Private Function CollectText(node As XmlNode) As String
		  ' Recursive text content. Text nodes contribute their Value; <text:s> /
		  ' <text:tab> / <text:line-break> become a space / tab / newline.
		  Var result As String = ""
		  Var child As XmlNode = node.FirstChild
		  While child <> Nil
		    If child.FirstChild <> Nil Then
		      result = result + CollectText(child)
		    Else
		      Select Case LocalName(child)
		      Case "s"
		        result = result + " "
		      Case "tab"
		        result = result + Chr(9)
		      Case "line-break"
		        result = result + EndOfLine
		      Else
		        ' a text node contributes its literal Value; an empty element
		        ' harmlessly contributes "" (XmlNode.Value is never Nil).
		        result = result + child.Value
		      End Select
		    End If
		    child = child.NextSibling
		  Wend
		  Return result
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 436F6E7665727420616E2049534F2064617465206F7220646174652D74696D6520737472696E6720286F66666963653A646174652D76616C75652920746F20616E20457863656C2073657269616C206E756D6265722E0A
		Private Function ISODateToSerial(s As String) As Double
		  ' "2026-05-11" or "2026-05-11T14:30:00" -> Excel serial (days since 1899-12-30).
		  If s = "" Then Return 0.0
		  Var datePart As String = s
		  Var timePart As String = ""
		  Var t As Integer = s.IndexOf("T")
		  If t >= 0 Then
		    datePart = s.Left(t)
		    timePart = s.Middle(t + 1)
		  End If
		  Var dp() As String = datePart.Split("-")
		  If dp.LastIndex < 2 Then Return 0.0
		  Var y As Integer = dp(0).ToInteger
		  Var mo As Integer = dp(1).ToInteger
		  Var d As Integer = dp(2).ToInteger
		  Var hh As Integer = 0
		  Var mm As Integer = 0
		  Var ss As Integer = 0
		  If timePart <> "" Then
		    Var tp() As String = timePart.Split(":")
		    If tp.LastIndex >= 0 Then hh = tp(0).ToInteger
		    If tp.LastIndex >= 1 Then mm = tp(1).ToInteger
		    If tp.LastIndex >= 2 Then ss = tp(2).ToInteger
		  End If
		  Var dt As New DateTime(y, mo, d, hh, mm, ss, 0, TimeZone.Current)
		  Return XLSXFormatter.DateTimeToExcelSerial(dt)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 436F6E7665727420616E2049534F2038363031206475726174696F6E20286F66666963653A74696D652D76616C75652C20652E672E20505431344833304D3135532920746F2061206672616374696F6E206F662061206461792E0A
		Private Function ISOTimeToFraction(s As String) As Double
		  ' ISO 8601 duration "PT14H30M15S" -> fraction of a day.
		  If Not s.BeginsWith("PT") Then Return 0.0
		  Var body As String = s.Middle(2)
		  Var h As Double = 0.0
		  Var m As Double = 0.0
		  Var sec As Double = 0.0
		  Var num As String = ""
		  For i As Integer = 0 To body.Length - 1
		    Var ch As String = body.Middle(i, 1)
		    Select Case ch
		    Case "H"
		      h = num.ToDouble
		      num = ""
		    Case "M"
		      m = num.ToDouble
		      num = ""
		    Case "S"
		      sec = num.ToDouble
		      num = ""
		    Else
		      num = num + ch
		    End Select
		  Next
		  Return ((h * 3600.0) + (m * 60.0) + sec) / 86400.0
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 526561642061206E616D65642061747472696275746520617320616E20496E74656765722C2072657475726E696E6720612064656661756C74207768656E20616273656E742E0A
		Private Function IntAttr(node As XmlNode, name As String, defaultValue As Integer) As Integer
		  Var v As String = node.GetAttribute(name)
		  If v = "" Then Return defaultValue
		  Return v.ToInteger
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 537472697020746865206E616D657370616365207072656669782066726F6D2061206E6F64652773207175616C6966696564206E616D652028652E672E20226E756D6265723A7965617222202D3E20227965617222292E0A
		Private Function LocalName(node As XmlNode) As String
		  ' Strip the namespace prefix from a node's qualified name.
		  Var nm As String = node.Name
		  Var i As Integer = nm.IndexOf(":")
		  If i >= 0 Then Return nm.Middle(i + 1)
		  Return nm
		End Function
	#tag EndMethod

	#tag Note, Name = About
		Reads OpenDocument Spreadsheet (.ods) files into the shared XLSXWorkbook
		model, so the UI / listbox fillers / format engine work unchanged.

		Two entry points (mirror XLSXReader), both returning XLSXWorkbook:
		  Open(file As FolderItem [, mode])
		  Open(data As MemoryBlock, sourceName [, mode])

		ODS layout differs from XLSX:
		  - All sheets live in a single content.xml (vs per-sheet sheetN.xml).
		  - Cell values carry office:value-type + a value attribute (vs <v>/<t>).
		  - No shared-strings table — text is inline in <text:p>.
		  - Number formats are <number:*-style> element trees, not code strings.
		  - Merged cells use table:number-rows/columns-spanned on the anchor;
		    covered cells appear as <table:covered-table-cell>.
		  - Repeated empty cells/rows are compressed via number-columns/rows-repeated.

		This reader converts each <number:date-style> / <number:number-style> /
		<number:currency-style> / <number:percentage-style> into the format-code
		strings XLSXFormatter already understands, and stamps the result onto each
		cell's FormatCode. Dates (ISO strings) are converted to Excel serials so the
		existing date-formatting path applies.

		Out of V1 scope: cell colors / fonts / borders, conditional formatting,
		charts, pivots, formula re-evaluation (cached value shown).
	#tag EndNote

	#tag ViewBehavior
	#tag EndViewBehavior
End Module
#tag EndModule
