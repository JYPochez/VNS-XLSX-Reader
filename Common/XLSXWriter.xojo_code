#tag Module
Protected Module XLSXWriter
	#tag Method, Flags = &h0, Description = 53657269616C697A652074686520776F726B626F6F6B20746F20616E202E786C73782066696C65206F6E206469736B2E0A
		Sub Save(wb As XLSXWorkbook, file As FolderItem)
		  Build(wb).SaveTo(file)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 53657269616C697A652074686520776F726B626F6F6B20746F20726177202E786C73782062797465732028666F72207468652057656220646F776E6C6F61642070617468292E0A
		Function ToMemoryBlock(wb As XLSXWorkbook) As MemoryBlock
		  Return Build(wb).ToMemoryBlock
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 436F6C6C6563742074686520776F726B626F6F6B27732064697374696E637420666F726D617420636F6465732C207468656E20656D6974206576657279204F5043207061727420696E746F20612053707265616473686565745A69705772697465722E0A
		Private Function Build(wb As XLSXWorkbook) As SpreadsheetZipWriter
		  If wb Is Nil Or wb.SheetCount = 0 Then
		    Raise New XLSXException(XLSXEnums.eParseError.Unsupported, "empty workbook")
		  End If

		  ' Build the full style tables (numFmts + fonts + fills + borders + cellXfs)
		  ' from every cell's format code AND visual style, keyed by a signature so
		  ' identical appearances share one xf. sigToXf maps signature -> xf index.
		  Var sigToXf As New Dictionary
		  Var stylesXml As String = BuildStyleTables(wb, sigToXf)

		  Var zip As New SpreadsheetZipWriter
		  zip.AddPart("[Content_Types].xml", ContentTypesXml(wb))
		  zip.AddPart("_rels/.rels", RootRelsXml)
		  zip.AddPart("xl/workbook.xml", WorkbookXml(wb))
		  zip.AddPart("xl/_rels/workbook.xml.rels", WorkbookRelsXml(wb))
		  zip.AddPart("xl/styles.xml", stylesXml)
		  Var tableCounter As Integer = 0
		  For i As Integer = 1 To wb.SheetCount
		    Var sheet As XLSXSheet = wb.SheetAtRaw(i)
		    ' Excel Table parts: one xl/tables/tableN.xml per table (globally numbered)
		    ' plus a worksheet rels part linking them.
		    Var tableGlobalIds() As Integer
		    For t As Integer = 0 To sheet.TableCount - 1
		      tableCounter = tableCounter + 1
		      tableGlobalIds.Add tableCounter
		      zip.AddPart("xl/tables/table" + Str(tableCounter) + ".xml", TableXml(sheet.TableAt(t), tableCounter, sheet))
		    Next
		    If tableGlobalIds.Count > 0 Then
		      zip.AddPart("xl/worksheets/_rels/sheet" + Str(i) + ".xml.rels", SheetRelsXml(tableGlobalIds))
		    End If
		    zip.AddPart("xl/worksheets/sheet" + Str(i) + ".xml", SheetXml(sheet, wb.Styles, sigToXf, tableGlobalIds.Count))
		  Next
		  Return zip
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 456D6974205B436F6E74656E745F54797065735D2E786D6C2077697468204F766572726964657320666F7220776F726B626F6F6B2C207374796C65732C20616E6420656163682073686565742E0A
		Private Function ContentTypesXml(wb As XLSXWorkbook) As String
		  Var parts() As String
		  parts.Add "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
		  parts.Add "<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">"
		  parts.Add "<Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>"
		  parts.Add "<Default Extension=""xml"" ContentType=""application/xml""/>"
		  parts.Add "<Override PartName=""/xl/workbook.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/>"
		  parts.Add "<Override PartName=""/xl/styles.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml""/>"
		  For i As Integer = 1 To wb.SheetCount
		    parts.Add "<Override PartName=""/xl/worksheets/sheet" + Str(i) + ".xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>"
		  Next
		  Var tableCount As Integer = 0
		  For i As Integer = 1 To wb.SheetCount
		    tableCount = tableCount + wb.SheetAtRaw(i).TableCount
		  Next
		  For k As Integer = 1 To tableCount
		    parts.Add "<Override PartName=""/xl/tables/table" + Str(k) + ".xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml""/>"
		  Next
		  parts.Add "</Types>"
		  Return String.FromArray(parts, "")
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 456D6974205F72656C732F2E72656C7320706F696E74696E6720617420786C2F776F726B626F6F6B2E786D6C2E0A
		Private Function RootRelsXml() As String
		  Var parts() As String
		  parts.Add "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
		  parts.Add "<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">"
		  parts.Add "<Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""xl/workbook.xml""/>"
		  parts.Add "</Relationships>"
		  Return String.FromArray(parts, "")
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 456D697420786C2F776F726B626F6F6B2E786D6C2077697468206F6E65203C73686565743E20656E7472792070657220736865657420286E616D65202B20724964292E0A
		Private Function WorkbookXml(wb As XLSXWorkbook) As String
		  Var parts() As String
		  parts.Add "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
		  parts.Add "<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">"
		  parts.Add "<sheets>"
		  For i As Integer = 1 To wb.SheetCount
		    Var nm As String = XLSXHelpers.XmlEscape(wb.SheetAtRaw(i).Name)
		    parts.Add "<sheet name=""" + nm + """ sheetId=""" + Str(i) + """ r:id=""rId" + Str(i) + """/>"
		  Next
		  parts.Add "</sheets>"
		  parts.Add "</workbook>"
		  Return String.FromArray(parts, "")
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 456D697420786C2F5F72656C732F776F726B626F6F6B2E786D6C2E72656C733A206F6E6520776F726B73686565742072656C6174696F6E736869702070657220736865657420706C757320746865207374796C65732072656C6174696F6E736869702E0A
		Private Function WorkbookRelsXml(wb As XLSXWorkbook) As String
		  Var parts() As String
		  parts.Add "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
		  parts.Add "<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">"
		  For i As Integer = 1 To wb.SheetCount
		    parts.Add "<Relationship Id=""rId" + Str(i) + """ Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet" + Str(i) + ".xml""/>"
		  Next
		  parts.Add "<Relationship Id=""rId" + Str(wb.SheetCount + 1) + """ Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.xml""/>"
		  parts.Add "</Relationships>"
		  Return String.FromArray(parts, "")
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 57616C6B2065766572792063656C6C206F6E63652C20646564757020666F6E74732F66696C6C732F626F72646572732F6E756D466D747320616E6420656D6974206F6E652063656C6C5866207065722064697374696E637420617070656172616E63652E20506F70756C6174657320736967546F586620287374796C65207369676E6174757265202D3E20786620696E6465782920616E642072657475726E73207374796C65732E786D6C2E0A
		Private Function BuildStyleTables(wb As XLSXWorkbook, sigToXf As Dictionary) As String
		  ' Walk every cell once, allocating deduped fonts / fills / borders / numFmts
		  ' and one cellXf per distinct appearance (format code + visual style).
		  ' Populates sigToXf (style signature -> cellXf index) for SheetXml, and
		  ' returns the assembled styles.xml. Index 0 in each reserved table is the
		  ' default; fills reserve 0=none and 1=gray125 as Excel requires.
		  Var fontXml() As String
		  fontXml.Add "<font><sz val=""11""/><name val=""Calibri""/><family val=""2""/></font>"
		  Var fontKey As New Dictionary

		  Var fillXml() As String
		  fillXml.Add "<fill><patternFill patternType=""none""/></fill>"
		  fillXml.Add "<fill><patternFill patternType=""gray125""/></fill>"
		  Var fillKey As New Dictionary

		  Var borderXml() As String
		  borderXml.Add "<border><left/><right/><top/><bottom/><diagonal/></border>"
		  Var borderKey As New Dictionary

		  Var numFmtCodes() As String
		  Var numFmtKey As New Dictionary

		  Var xfXml() As String
		  xfXml.Add "<xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0"" xfId=""0""/>"

		  For s As Integer = 1 To wb.SheetCount
		    Var sheet As XLSXSheet = wb.SheetAtRaw(s)
		    For r As Integer = 1 To sheet.RowCount
		      For c As Integer = 1 To sheet.ColCount
		        Var cell As XLSXCell = sheet.CellAtRaw(r, c)
		        If cell.IsEmpty Then Continue
		        Var code As String = XLSXHelpers.WriteFormatCode(cell, wb.Styles)
		        Var cst As XLSXCellStyle = cell.ResolvedStyle(wb.Styles)
		        Var sig As String = StyleSignature(code, cst)
		        If sig = "" Or sigToXf.HasKey(sig) Then Continue

		        Var numFmtId As Integer = 0
		        If code <> "" And code <> "General" Then
		          If numFmtKey.HasKey(code) Then
		            numFmtId = numFmtKey.Value(code)
		          Else
		            numFmtId = 164 + numFmtCodes.Count
		            numFmtCodes.Add code
		            numFmtKey.Value(code) = numFmtId
		          End If
		        End If
		        Var fontId As Integer = FontIdFor(cst, fontXml, fontKey)
		        Var fillId As Integer = FillIdFor(cst, fillXml, fillKey)
		        Var borderId As Integer = BorderIdFor(cst, borderXml, borderKey)
		        sigToXf.Value(sig) = xfXml.Count
		        xfXml.Add XfElement(numFmtId, fontId, fillId, borderId, cst)
		      Next
		    Next
		  Next

		  Var parts() As String
		  parts.Add "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
		  parts.Add "<styleSheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">"
		  If numFmtCodes.Count > 0 Then
		    parts.Add "<numFmts count=""" + Str(numFmtCodes.Count) + """>"
		    For i As Integer = 0 To numFmtCodes.LastIndex
		      parts.Add "<numFmt numFmtId=""" + Str(164 + i) + """ formatCode=""" + XLSXHelpers.XmlEscape(numFmtCodes(i)) + """/>"
		    Next
		    parts.Add "</numFmts>"
		  End If
		  parts.Add "<fonts count=""" + Str(fontXml.Count) + """>" + String.FromArray(fontXml, "") + "</fonts>"
		  parts.Add "<fills count=""" + Str(fillXml.Count) + """>" + String.FromArray(fillXml, "") + "</fills>"
		  parts.Add "<borders count=""" + Str(borderXml.Count) + """>" + String.FromArray(borderXml, "") + "</borders>"
		  parts.Add "<cellStyleXfs count=""1""><xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0""/></cellStyleXfs>"
		  parts.Add "<cellXfs count=""" + Str(xfXml.Count) + """>" + String.FromArray(xfXml, "") + "</cellXfs>"
		  parts.Add "<cellStyles count=""1""><cellStyle name=""Normal"" xfId=""0"" builtinId=""0""/></cellStyles>"
		  parts.Add "</styleSheet>"
		  Return String.FromArray(parts, "")
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 537461626C65206B657920666F7220612063656C6C27732066756C6C20617070656172616E63652028666F726D617420636F6465202B2076697375616C207374796C65292E20456D707479207768656E207468652063656C6C20697320706C61696E20616E64206E65656473206E6F2063656C6C58662E0A
		Private Function StyleSignature(code As String, cst As XLSXCellStyle) As String
		  ' A stable key for a cell's full appearance. Empty when the cell is plain
		  ' (no number format and a default visual style) so it needs no cellXf.
		  If (code = "" Or code = "General") And cst.IsDefault Then Return ""
		  Var p() As String
		  p.Add code
		  p.Add If(cst.Bold, "b", "")
		  p.Add If(cst.Italic, "i", "")
		  p.Add If(cst.Underline, "u", "")
		  p.Add cst.FontName
		  p.Add Str(cst.FontSize)
		  p.Add If(cst.HasFontColor, ColorToHex(cst.FontColor), "")
		  p.Add If(cst.HasBackground, ColorToHex(cst.BackgroundColor), "")
		  p.Add Str(Integer(cst.AlignH)) + Str(Integer(cst.AlignV)) + If(cst.WrapText, "w", "")
		  p.Add Str(Integer(cst.BorderLeft)) + Str(Integer(cst.BorderRight)) + Str(Integer(cst.BorderTop)) + Str(Integer(cst.BorderBottom))
		  p.Add If(cst.HasBorderColor, ColorToHex(cst.BorderColor), "")
		  Return String.FromArray(p, "|")
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 44656475702F616C6C6F636174652061203C666F6E743E20666F722061207374796C65277320666F6E7420617370656374733B2072657475726E732030202864656661756C7420666F6E7429207768656E206E6F6E6520617265207365742E0A
		Private Function FontIdFor(cst As XLSXCellStyle, fontXml() As String, fontKey As Dictionary) As Integer
		  ' Default font when the cell sets no font aspect.
		  If Not cst.Bold And Not cst.Italic And Not cst.Underline And cst.FontName = "" _
		    And cst.FontSize <= 0 And Not cst.HasFontColor Then Return 0
		  Var sig As String = If(cst.Bold, "b", "") + If(cst.Italic, "i", "") + If(cst.Underline, "u", "") _
		    + "|" + cst.FontName + "|" + Str(cst.FontSize) + "|" + If(cst.HasFontColor, ColorToHex(cst.FontColor), "")
		  If fontKey.HasKey(sig) Then Return fontKey.Value(sig)
		  Var f() As String
		  If cst.Bold Then f.Add "<b/>"
		  If cst.Italic Then f.Add "<i/>"
		  If cst.Underline Then f.Add "<u/>"
		  ' Only emit size / name when the model actually specifies them. Inventing a
		  ' default here turns "unspecified" into an explicit 11pt Calibri, which the
		  ' reader then treats as a deliberate override — so an italic-only cell came
		  ' back smaller than it went in. A font may legally carry just <i/>; the
		  ' complete default font stays at index 0 for renderers that need one.
		  If cst.FontSize > 0 Then f.Add "<sz val=""" + Str(cst.FontSize) + """/>"
		  If cst.FontName <> "" Then f.Add "<name val=""" + XLSXHelpers.XmlEscape(cst.FontName) + """/>"
		  If cst.HasFontColor Then f.Add "<color rgb=""" + ColorToHex(cst.FontColor) + """/>"
		  Var id As Integer = fontXml.Count
		  fontXml.Add "<font>" + String.FromArray(f, "") + "</font>"
		  fontKey.Value(sig) = id
		  Return id
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 44656475702F616C6C6F63617465206120736F6C6964203C66696C6C3E20666F722061207374796C652773206261636B67726F756E643B2072657475726E732030207768656E2074686572652773206E6F2066696C6C2E0A
		Private Function FillIdFor(cst As XLSXCellStyle, fillXml() As String, fillKey As Dictionary) As Integer
		  If Not cst.HasBackground Then Return 0
		  Var hex As String = ColorToHex(cst.BackgroundColor)
		  If fillKey.HasKey(hex) Then Return fillKey.Value(hex)
		  Var id As Integer = fillXml.Count
		  fillXml.Add "<fill><patternFill patternType=""solid""><fgColor rgb=""" + hex + """/></patternFill></fill>"
		  fillKey.Value(hex) = id
		  Return id
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 44656475702F616C6C6F636174652061203C626F726465723E20666F722061207374796C652773207065722D6564676520626F72646572733B2072657475726E732030207768656E20746865726520617265206E6F6E652E0A
		Private Function BorderIdFor(cst As XLSXCellStyle, borderXml() As String, borderKey As Dictionary) As Integer
		  If Not cst.HasAnyBorder Then Return 0
		  Var colorHex As String = If(cst.HasBorderColor, ColorToHex(cst.BorderColor), "")
		  Var sig As String = Str(Integer(cst.BorderLeft)) + Str(Integer(cst.BorderRight)) _
		    + Str(Integer(cst.BorderTop)) + Str(Integer(cst.BorderBottom)) + "|" + colorHex
		  If borderKey.HasKey(sig) Then Return borderKey.Value(sig)
		  Var b As String = "<border>" _
		    + BorderEdgeXml("left", cst.BorderLeft, colorHex) _
		    + BorderEdgeXml("right", cst.BorderRight, colorHex) _
		    + BorderEdgeXml("top", cst.BorderTop, colorHex) _
		    + BorderEdgeXml("bottom", cst.BorderBottom, colorHex) _
		    + "<diagonal/></border>"
		  Var id As Integer = borderXml.Count
		  borderXml.Add b
		  borderKey.Value(sig) = id
		  Return id
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 456D6974206F6E6520626F72646572206564676520283C6C6566747C72696768747C746F707C626F74746F6D207374796C653D227468696E7C6D656469756D7C746869636B223E3C636F6C6F722F3E3C2F3E292E0A
		Private Function BorderEdgeXml(edge As String, style As XLSXEnums.eBorderStyle, colorHex As String) As String
		  If style = XLSXEnums.eBorderStyle.None Then Return "<" + edge + "/>"
		  Var s As String = "thin"
		  If style = XLSXEnums.eBorderStyle.Medium Then s = "medium"
		  If style = XLSXEnums.eBorderStyle.Thick Then s = "thick"
		  Var col As String = If(colorHex <> "", "<color rgb=""" + colorHex + """/>", "")
		  Return "<" + edge + " style=""" + s + """>" + col + "</" + edge + ">"
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 4275696C64206F6E65203C78663E2077697468206E756D466D742F666F6E742F66696C6C2F626F72646572206964732C206170706C792A20666C6167732C20616E6420616E206F7074696F6E616C203C616C69676E6D656E743E206368696C642E0A
		Private Function XfElement(numFmtId As Integer, fontId As Integer, fillId As Integer, borderId As Integer, cst As XLSXCellStyle) As String
		  Var attrs As String = "numFmtId=""" + Str(numFmtId) + """ fontId=""" + Str(fontId) _
		    + """ fillId=""" + Str(fillId) + """ borderId=""" + Str(borderId) + """ xfId=""0"""
		  If numFmtId <> 0 Then attrs = attrs + " applyNumberFormat=""1"""
		  If fontId <> 0 Then attrs = attrs + " applyFont=""1"""
		  If fillId <> 0 Then attrs = attrs + " applyFill=""1"""
		  If borderId <> 0 Then attrs = attrs + " applyBorder=""1"""
		  Var align As String = AlignmentXml(cst)
		  If align = "" Then Return "<xf " + attrs + "/>"
		  Return "<xf " + attrs + " applyAlignment=""1"">" + align + "</xf>"
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 4275696C6420746865203C616C69676E6D656E743E20656C656D656E7420666F72206E6F6E2D64656661756C7420686F72697A6F6E74616C2F766572746963616C2F777261703B20656D707479206F74686572776973652E0A
		Private Function AlignmentXml(cst As XLSXCellStyle) As String
		  Var a() As String
		  Select Case cst.AlignH
		  Case XLSXEnums.eAlignH.Left
		    a.Add "horizontal=""left"""
		  Case XLSXEnums.eAlignH.Center
		    a.Add "horizontal=""center"""
		  Case XLSXEnums.eAlignH.Right
		    a.Add "horizontal=""right"""
		  End Select
		  Select Case cst.AlignV
		  Case XLSXEnums.eAlignV.Top
		    a.Add "vertical=""top"""
		  Case XLSXEnums.eAlignV.Middle
		    a.Add "vertical=""center"""
		  End Select
		  If cst.WrapText Then a.Add "wrapText=""1"""
		  If a.Count = 0 Then Return ""
		  Return "<alignment " + String.FromArray(a, " ") + "/>"
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 466F726D6174206120436F6C6F72206173204F4F584D4C202246465252474742422220286F70617175652C2075707065726361736520686578292E0A
		Private Function ColorToHex(c As Color) As String
		  ' "FFRRGGBB" (opaque) from a Color, uppercase hex.
		  Return "FF" + Pad2Hex(c.Red) + Pad2Hex(c.Green) + Pad2Hex(c.Blue)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 54776F2D6469676974207570706572636173652068657820666F72206120302E2E32353520627974652E0A
		Private Function Pad2Hex(v As Integer) As String
		  Var h As String = Hex(v)
		  If h.Length < 2 Then h = "0" + h
		  Return h
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 456D6974206F6E6520776F726B73686565743A20726F7773206F66203C633E2063656C6C732028696E6C696E6520737472696E67732C206E756D626572732F73657269616C732C20626F6F6C65616E732920706C7573203C6D6572676543656C6C733E2E0A
		Private Function SheetXml(sheet As XLSXSheet, styles As XLSXStyles, sigToXf As Dictionary, tablePartCount As Integer = 0) As String
		  Var parts() As String
		  parts.Add "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
		  parts.Add "<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">"
		  ' dimension + a default row height: Excel recomputes these anyway, but
		  ' minimal renderers (e.g. macOS Quick Look) blow up row heights without
		  ' an explicit defaultRowHeight.
		  Var lastRef As String = "A1"
		  If sheet.RowCount > 0 And sheet.ColCount > 0 Then
		    lastRef = XLSXCellRef.IndexToColLettersRaw(sheet.ColCount) + Str(sheet.RowCount)
		  End If
		  parts.Add "<dimension ref=""A1:" + lastRef + """/>"
		  parts.Add "<sheetFormatPr defaultRowHeight=""15"" customHeight=""false""/>"
		  ' <cols> (column widths) must come after <sheetFormatPr> and before <sheetData>.
		  If sheet.HasColumnWidths Then
		    Var colParts() As String
		    For c As Integer = 1 To sheet.ColCount
		      Var wpt As Double = sheet.ColumnWidthRaw(c)
		      If wpt <= 0 Then Continue
		      Var chars As Double = XLSXHelpers.ColumnPointsToChars(wpt)
		      colParts.Add "<col min=""" + Str(c) + """ max=""" + Str(c) + """ width=""" + Str(chars) + """ customWidth=""1""/>"
		    Next
		    If colParts.Count > 0 Then parts.Add "<cols>" + String.FromArray(colParts, "") + "</cols>"
		  End If

		  parts.Add "<sheetData>"

		  For r As Integer = 1 To sheet.RowCount
		    Var rowCells() As String
		    For c As Integer = 1 To sheet.ColCount
		      Var cell As XLSXCell = sheet.CellAtRaw(r, c)
		      If cell.IsEmpty Then Continue
		      Var ref As String = XLSXCellRef.IndexToColLettersRaw(c) + Str(r)
		      Var sAttr As String = ""
		      Var sig As String = StyleSignature(XLSXHelpers.WriteFormatCode(cell, styles), cell.ResolvedStyle(styles))
		      If sig <> "" And sigToXf.HasKey(sig) Then
		        sAttr = " s=""" + Str(sigToXf.Value(sig).IntegerValue) + """"
		      End If
		      Select Case cell.eType
		      Case XLSXEnums.eCellType.Str, XLSXEnums.eCellType.ErrorVal
		        rowCells.Add "<c r=""" + ref + """ t=""inlineStr""" + sAttr + "><is><t xml:space=""preserve"">" _
		          + XLSXHelpers.XmlEscape(cell.RawString) + "</t></is></c>"
		      Case XLSXEnums.eCellType.Bool
		        Var bv As String = If(cell.BooleanValue, "1", "0")
		        rowCells.Add "<c r=""" + ref + """ t=""b""" + sAttr + "><v>" + bv + "</v></c>"
		      Else
		        If cell.eType = XLSXEnums.eCellType.FormulaCached And cell.Formula <> "" Then
		          ' Formula cell: emit <f> (R1C1 authored formulas convert to A1 here,
		          ' anchored at this cell) plus the cached <v> when we have one.
		          Var fA1 As String = cell.Formula
		          If cell.FormulaIsR1C1 Then fA1 = XLSXHelpers.FormulaToA1(cell.Formula, r, c)
		          Var fBody As String = "<f>" + XLSXHelpers.XmlEscape(fA1) + "</f>"
		          If cell.RawString <> "" Then fBody = fBody + "<v>" + XLSXHelpers.XmlEscape(cell.RawString) + "</v>"
		          rowCells.Add "<c r=""" + ref + """" + sAttr + ">" + fBody + "</c>"
		        Else
		          ' Number or DateValue (serial), cached value only.
		          rowCells.Add "<c r=""" + ref + """" + sAttr + "><v>" + XLSXHelpers.XmlEscape(cell.RawString) + "</v></c>"
		        End If
		      End Select
		    Next
		    Var hpt As Double = sheet.RowHeightRaw(r)
		    Var rowAttrs As String = " r=""" + Str(r) + """"
		    If hpt > 0 Then rowAttrs = rowAttrs + " ht=""" + Str(hpt) + """ customHeight=""1"""
		    If rowCells.Count > 0 Then
		      parts.Add "<row" + rowAttrs + ">" + String.FromArray(rowCells, "") + "</row>"
		    ElseIf hpt > 0 Then
		      ' A custom-height row with no cells still needs an element to carry ht.
		      parts.Add "<row" + rowAttrs + "/>"
		    End If
		  Next

		  parts.Add "</sheetData>"

		  If sheet.MergedRangeCount > 0 Then
		    parts.Add "<mergeCells count=""" + Str(sheet.MergedRangeCount) + """>"
		    For i As Integer = 0 To sheet.MergedRangeCount - 1
		      Var rng As XLSXCellRange = sheet.MergedRangeAt(i)
		      If rng Is Nil Then Continue
		      Var ref As String = XLSXCellRef.IndexToColLettersRaw(rng.FirstColRaw) + Str(rng.FirstRowRaw) _
		        + ":" + XLSXCellRef.IndexToColLettersRaw(rng.LastColRaw) + Str(rng.LastRowRaw)
		      parts.Add "<mergeCell ref=""" + ref + """/>"
		    Next
		    parts.Add "</mergeCells>"
		  End If

		  If tablePartCount > 0 Then
		    parts.Add "<tableParts count=""" + Str(tablePartCount) + """>"
		    For k As Integer = 1 To tablePartCount
		      parts.Add "<tablePart r:id=""rId" + Str(k) + """/>"
		    Next
		    parts.Add "</tableParts>"
		  End If
		  parts.Add "</worksheet>"
		  Return String.FromArray(parts, "")
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Function TableXml(table As XLSXTable, globalId As Integer, sheet As XLSXSheet) As String
		  ' Emit one xl/tables/tableN.xml for an Excel Table (ListObject): range,
		  ' autoFilter, tableColumns (names taken from the header row), tableStyleInfo.
		  Var ref As String = XLSXCellRef.IndexToColLettersRaw(table.FirstColRaw) + Str(table.FirstRowRaw) _
		    + ":" + XLSXCellRef.IndexToColLettersRaw(table.LastColRaw) + Str(table.LastRowRaw)
		  Var nm As String = XLSXHelpers.XmlEscape(SanitizeTableName(table.Name, globalId))
		  Var parts() As String
		  parts.Add "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
		  Var attrs As String = "id=""" + Str(globalId) + """ name=""" + nm + """ displayName=""" + nm _
		    + """ ref=""" + ref + """ headerRowCount=""" + Str(table.HeaderRowCount) + """"
		  If table.TotalsRowCount > 0 Then
		    attrs = attrs + " totalsRowShown=""1"" totalsRowCount=""" + Str(table.TotalsRowCount) + """"
		  Else
		    attrs = attrs + " totalsRowShown=""0"""
		  End If
		  parts.Add "<table xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" " + attrs + ">"
		  If table.HeaderRowCount > 0 Then
		    Var afLast As Integer = table.LastRowRaw - table.TotalsRowCount
		    Var afRef As String = XLSXCellRef.IndexToColLettersRaw(table.FirstColRaw) + Str(table.FirstRowRaw) _
		      + ":" + XLSXCellRef.IndexToColLettersRaw(table.LastColRaw) + Str(afLast)
		    parts.Add "<autoFilter ref=""" + afRef + """/>"
		  End If
		  Var names() As String = TableColumnNames(table, sheet)
		  parts.Add "<tableColumns count=""" + Str(names.Count) + """>"
		  For j As Integer = 0 To names.LastIndex
		    parts.Add "<tableColumn id=""" + Str(j + 1) + """ name=""" + XLSXHelpers.XmlEscape(names(j)) + """/>"
		  Next
		  parts.Add "</tableColumns>"
		  Var tsi As String = "<tableStyleInfo"
		  If table.StyleName <> "" Then tsi = tsi + " name=""" + XLSXHelpers.XmlEscape(table.StyleName) + """"
		  tsi = tsi + " showFirstColumn=""" + If(table.ShowFirstColumn, "1", "0") + """"
		  tsi = tsi + " showLastColumn=""" + If(table.ShowLastColumn, "1", "0") + """"
		  tsi = tsi + " showRowStripes=""" + If(table.ShowRowStripes, "1", "0") + """"
		  tsi = tsi + " showColumnStripes=""" + If(table.ShowColumnStripes, "1", "0") + """/>"
		  parts.Add tsi
		  parts.Add "</table>"
		  Return String.FromArray(parts, "")
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Function TableColumnNames(table As XLSXTable, sheet As XLSXSheet) As String()
		  ' Column names from the table's header row; blanks get ColumnN; duplicates get
		  ' a numeric suffix (Excel requires unique, non-empty column names).
		  Var names() As String
		  Var seen As New Dictionary
		  For c As Integer = table.FirstColRaw To table.LastColRaw
		    Var nm As String = ""
		    ' Verbatim — Excel requires each tableColumn name to match its header cell
		    ' EXACTLY, and flags the workbook as needing repair otherwise. Trimming
		    ' would rename a real-world header like " Sales" (Microsoft's own Financial
		    ' Sample has one) to "Sales" while the cell keeps its leading space.
		    If table.HeaderRowCount > 0 Then nm = sheet.CellAtRaw(table.FirstRowRaw, c).RawString
		    If nm = "" Then nm = "Column" + Str(c - table.FirstColRaw + 1)
		    Var base As String = nm
		    Var n As Integer = 2
		    While seen.HasKey(nm)
		      nm = base + Str(n)
		      n = n + 1
		    Wend
		    seen.Value(nm) = True
		    names.Add nm
		  Next
		  Return names
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Function SanitizeTableName(rawName As String, globalId As Integer) As String
		  ' Excel's name/displayName must be a valid identifier: letters/digits/_ only,
		  ' starting with a letter or underscore. Fall back to TableN.
		  Var out As String = ""
		  For i As Integer = 0 To rawName.Length - 1
		    Var ch As String = rawName.Middle(i, 1)
		    If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Or ch = "_" Then
		      out = out + ch
		    Else
		      out = out + "_"
		    End If
		  Next
		  If out = "" Then Return "Table" + Str(globalId)
		  Var first As String = out.Middle(0, 1)
		  If Not ((first >= "A" And first <= "Z") Or (first >= "a" And first <= "z") Or first = "_") Then out = "_" + out
		  Return out
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Function SheetRelsXml(tableGlobalIds() As Integer) As String
		  ' Worksheet rels linking each rIdK to ../tables/table{globalId}.xml.
		  Var parts() As String
		  parts.Add "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
		  parts.Add "<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">"
		  For k As Integer = 0 To tableGlobalIds.LastIndex
		    parts.Add "<Relationship Id=""rId" + Str(k + 1) + """ Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"" Target=""../tables/table" + Str(tableGlobalIds(k)) + ".xml""/>"
		  Next
		  parts.Add "</Relationships>"
		  Return String.FromArray(parts, "")
		End Function
	#tag EndMethod

	#tag Note, Name = About
		Serializes the shared XLSXWorkbook model back into an .xlsx file (OPC zip).

		Entry points:
		  Save(wb, file)        - write to disk
		  ToMemoryBlock(wb)     - raw bytes (Web download)

		Parts emitted: [Content_Types].xml, _rels/.rels, xl/workbook.xml,
		xl/_rels/workbook.xml.rels, xl/styles.xml, xl/worksheets/sheetN.xml.

		Design choices:
		  - Strings are written as INLINE strings (t="inlineStr"), so no
		    sharedStrings.xml is needed. Excel, LibreOffice, and our own reader all
		    accept inline strings.
		  - Number formats: every distinct format code in the workbook (from
		    XLSXHelpers.WriteFormatCode) becomes one custom numFmt (id 164+) and one
		    cellXf. Date cells with no explicit code get a default ISO date code so
		    they don't degrade to bare serials.
		  - Formula cells are written as their cached VALUE (the model does not
		    keep formula text). Error cells are written as inline text.
		  - Merged ranges are emitted as <mergeCells>.

		Out of scope: fonts / fills / borders styling fidelity, formulas, charts,
		images, conditional formatting (none of these exist in the model).
	#tag EndNote

	#tag ViewBehavior
	#tag EndViewBehavior
End Module
#tag EndModule
