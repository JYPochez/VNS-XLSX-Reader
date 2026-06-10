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

		  ' Pass 1: collect the distinct format codes used anywhere in the workbook.
		  ' Each becomes one custom numFmt (id 164+) + one cellXf (index 1+; 0 = General).
		  Var codes() As String
		  Var codeToXf As New Dictionary
		  For s As Integer = 1 To wb.SheetCount
		    Var sheet As XLSXSheet = wb.SheetAt(s)
		    For r As Integer = 1 To sheet.RowCount
		      For c As Integer = 1 To sheet.ColCount
		        Var cell As XLSXCell = sheet.CellAt(r, c)
		        If cell.IsEmpty Then Continue
		        Var code As String = XLSXHelpers.WriteFormatCode(cell, wb.Styles)
		        If code <> "" And Not codeToXf.HasKey(code) Then
		          codes.Add code
		          codeToXf.Value(code) = codes.Count
		        End If
		      Next
		    Next
		  Next

		  Var zip As New SpreadsheetZipWriter
		  zip.AddPart("[Content_Types].xml", ContentTypesXml(wb))
		  zip.AddPart("_rels/.rels", RootRelsXml)
		  zip.AddPart("xl/workbook.xml", WorkbookXml(wb))
		  zip.AddPart("xl/_rels/workbook.xml.rels", WorkbookRelsXml(wb))
		  zip.AddPart("xl/styles.xml", StylesXml(codes))
		  For i As Integer = 1 To wb.SheetCount
		    zip.AddPart("xl/worksheets/sheet" + Str(i) + ".xml", SheetXml(wb.SheetAt(i), wb.Styles, codeToXf))
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
		    Var nm As String = XLSXHelpers.XmlEscape(wb.SheetAt(i).Name)
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

	#tag Method, Flags = &h21, Description = 456D697420786C2F7374796C65732E786D6C3A20656163682064697374696E637420666F726D617420636F6465206265636F6D6573206120637573746F6D206E756D466D7420286964203136342B2920616E6420612063656C6C58663B20696E64657820302073746179732047656E6572616C2E0A
		Private Function StylesXml(codes() As String) As String
		  ' Custom numFmt ids start at 164 (below that is reserved for built-ins).
		  ' cellXf index 0 is the General style; code i (0-based) gets cellXf i+1.
		  Var parts() As String
		  parts.Add "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
		  parts.Add "<styleSheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">"
		  If codes.Count > 0 Then
		    parts.Add "<numFmts count=""" + Str(codes.Count) + """>"
		    For i As Integer = 0 To codes.LastIndex
		      parts.Add "<numFmt numFmtId=""" + Str(164 + i) + """ formatCode=""" + XLSXHelpers.XmlEscape(codes(i)) + """/>"
		    Next
		    parts.Add "</numFmts>"
		  End If
		  parts.Add "<fonts count=""1""><font><sz val=""11""/><name val=""Calibri""/><family val=""2""/></font></fonts>"
		  parts.Add "<fills count=""2""><fill><patternFill patternType=""none""/></fill><fill><patternFill patternType=""gray125""/></fill></fills>"
		  parts.Add "<borders count=""1""><border/></borders>"
		  parts.Add "<cellStyleXfs count=""1""><xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0""/></cellStyleXfs>"
		  parts.Add "<cellXfs count=""" + Str(codes.Count + 1) + """>"
		  parts.Add "<xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0"" xfId=""0""/>"
		  For i As Integer = 0 To codes.LastIndex
		    parts.Add "<xf numFmtId=""" + Str(164 + i) + """ fontId=""0"" fillId=""0"" borderId=""0"" xfId=""0"" applyNumberFormat=""1""/>"
		  Next
		  parts.Add "</cellXfs>"
		  parts.Add "<cellStyles count=""1""><cellStyle name=""Normal"" xfId=""0"" builtinId=""0""/></cellStyles>"
		  parts.Add "</styleSheet>"
		  Return String.FromArray(parts, "")
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 456D6974206F6E6520776F726B73686565743A20726F7773206F66203C633E2063656C6C732028696E6C696E6520737472696E67732C206E756D626572732F73657269616C732C20626F6F6C65616E732920706C7573203C6D6572676543656C6C733E2E0A
		Private Function SheetXml(sheet As XLSXSheet, styles As XLSXStyles, codeToXf As Dictionary) As String
		  Var parts() As String
		  parts.Add "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
		  parts.Add "<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">"
		  ' dimension + a default row height: Excel recomputes these anyway, but
		  ' minimal renderers (e.g. macOS Quick Look) blow up row heights without
		  ' an explicit defaultRowHeight.
		  Var lastRef As String = "A1"
		  If sheet.RowCount > 0 And sheet.ColCount > 0 Then
		    lastRef = XLSXCellRef.IndexToColLetters(sheet.ColCount) + Str(sheet.RowCount)
		  End If
		  parts.Add "<dimension ref=""A1:" + lastRef + """/>"
		  parts.Add "<sheetFormatPr defaultRowHeight=""15"" customHeight=""false""/>"
		  parts.Add "<sheetData>"

		  For r As Integer = 1 To sheet.RowCount
		    Var rowCells() As String
		    For c As Integer = 1 To sheet.ColCount
		      Var cell As XLSXCell = sheet.CellAt(r, c)
		      If cell.IsEmpty Then Continue
		      Var ref As String = XLSXCellRef.IndexToColLetters(c) + Str(r)
		      Var sAttr As String = ""
		      Var code As String = XLSXHelpers.WriteFormatCode(cell, styles)
		      If code <> "" And codeToXf.HasKey(code) Then
		        sAttr = " s=""" + Str(codeToXf.Value(code).IntegerValue) + """"
		      End If
		      Select Case cell.eType
		      Case XLSXEnums.eCellType.Str, XLSXEnums.eCellType.ErrorVal
		        rowCells.Add "<c r=""" + ref + """ t=""inlineStr""" + sAttr + "><is><t xml:space=""preserve"">" _
		          + XLSXHelpers.XmlEscape(cell.RawString) + "</t></is></c>"
		      Case XLSXEnums.eCellType.Bool
		        Var bv As String = If(cell.BooleanValue, "1", "0")
		        rowCells.Add "<c r=""" + ref + """ t=""b""" + sAttr + "><v>" + bv + "</v></c>"
		      Else
		        ' Number, FormulaCached (cached value only), DateValue (serial).
		        rowCells.Add "<c r=""" + ref + """" + sAttr + "><v>" + XLSXHelpers.XmlEscape(cell.RawString) + "</v></c>"
		      End Select
		    Next
		    If rowCells.Count > 0 Then
		      parts.Add "<row r=""" + Str(r) + """>" + String.FromArray(rowCells, "") + "</row>"
		    End If
		  Next

		  parts.Add "</sheetData>"

		  If sheet.MergedRangeCount > 0 Then
		    parts.Add "<mergeCells count=""" + Str(sheet.MergedRangeCount) + """>"
		    For i As Integer = 0 To sheet.MergedRangeCount - 1
		      Var rng As XLSXCellRange = sheet.MergedRangeAt(i)
		      If rng Is Nil Then Continue
		      Var ref As String = XLSXCellRef.IndexToColLetters(rng.FirstCol) + Str(rng.FirstRow) _
		        + ":" + XLSXCellRef.IndexToColLetters(rng.LastCol) + Str(rng.LastRow)
		      parts.Add "<mergeCell ref=""" + ref + """/>"
		    Next
		    parts.Add "</mergeCells>"
		  End If

		  parts.Add "</worksheet>"
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
