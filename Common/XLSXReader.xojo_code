#tag Module
Protected Module XLSXReader
	#tag Method, Flags = &h0, Description = 4F70656E20616E20584C53582066726F6D20612066696C65206F6E206469736B2E2052616973657320584C5358457863657074696F6E206F6E206661696C7572652E0A
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

	#tag Method, Flags = &h0, Description = 4F70656E20616E20584C53582066726F6D2072617720627974657320285765622075706C6F6164292E2054686520736F757263654E616D65206973207573656420666F7220646961676E6F737469637320616E64205549207469746C652E0A
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

	#tag Method, Flags = &h21, Description = 496E7465726E616C20706970656C696E653A20706172736520736861726564537472696E67732C207374796C65732C2072656C732C207468656E2065616368207368656574206C697374656420696E20776F726B626F6F6B2E786D6C2E0A
		Private Function OpenFromZip(zip As XLSXZip, sourceName As String) As XLSXWorkbook
		  If Not zip.HasPart("xl/workbook.xml") Then
		    Raise New XLSXException(XLSXEnums.eParseError.MissingPart, "xl/workbook.xml")
		  End If

		  Var wb As New XLSXWorkbook(sourceName)
		  wb.SharedStrings = ParseSharedStrings(zip.ReadPart("xl/sharedStrings.xml"))
		  wb.Styles = New XLSXStyles(zip.ReadPart("xl/styles.xml"), zip.ReadPart("xl/theme/theme1.xml"))

		  Var sheetMap As Dictionary = ParseRelsToTargets(zip.ReadPart("xl/_rels/workbook.xml.rels"))
		  Var workbookXml As String = zip.ReadPart("xl/workbook.xml")
		  Var doc As New XmlDocument
		  Try
		    doc.LoadXml(workbookXml)
		  Catch
		    Raise New XLSXException(XLSXEnums.eParseError.MalformedXML, "xl/workbook.xml")
		  End Try

		  ' 1904 date system (<workbookPr date1904="1"/>): common in Mac-authored
		  ' files. Detect it here and let each sheet normalize its date serials.
		  Var date1904 As Boolean = False
		  Var prNodes As XmlNodeList = doc.Xql("//*[local-name()='workbookPr']")
		  If prNodes.Length > 0 Then
		    Var v As String = prNodes.Item(0).GetAttribute("date1904").Lowercase
		    date1904 = (v = "1" Or v = "true")
		  End If

		  Var sheetNodes As XmlNodeList = doc.Xql("//*[local-name()='sheets']/*[local-name()='sheet']")
		  For i As Integer = 0 To sheetNodes.Length - 1
		    Var sn As XmlNode = sheetNodes.Item(i)
		    Var name As String = sn.GetAttribute("name")
		    Var rid As String = sn.GetAttribute("r:id")
		    If rid = "" Then rid = sn.GetAttribute("id")
		    Var target As String = If(sheetMap.HasKey(rid), sheetMap.Value(rid), "")
		    If target = "" Then Continue

		    ' rels target is relative to the rels owner — for xl/_rels/workbook.xml.rels
		    ' that means relative to xl/. So "worksheets/sheet1.xml" -> "xl/worksheets/sheet1.xml".
		    Var partPath As String = target
		    If Not partPath.BeginsWith("/") Then
		      partPath = "xl/" + partPath
		    Else
		      partPath = partPath.Middle(1)
		    End If

		    Var sheetXml As String = zip.ReadPart(partPath)
		    Var sheet As New XLSXSheet(name, i + 1, sheetXml, wb.SharedStrings, wb.Styles, date1904)
		    LoadTablesForSheet(zip, partPath, sheetXml, sheet)
		    wb.AddSheet(sheet)
		  Next

		  Return wb
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 506172736520786C2F736861726564537472696E67732E786D6C20696E746F206120537472696E672829206F66207265736F6C76656420737472696E677320286F6E6520656E74727920706572203C73693E2C206A6F696E696E6720616C6C203C743E206368696C6472656E20666F7220726963682D746578742072756E73292E0A
		Private Function ParseSharedStrings(xml As String) As String()
		  Var arr() As String
		  If xml = "" Then Return arr
		  Var doc As New XmlDocument
		  Try
		    doc.LoadXml(xml)
		  Catch
		    Raise New XLSXException(XLSXEnums.eParseError.MalformedXML, "xl/sharedStrings.xml")
		  End Try
		  Var siNodes As XmlNodeList = doc.Xql("//*[local-name()='sst']/*[local-name()='si']")
		  For i As Integer = 0 To siNodes.Length - 1
		    arr.Add(JoinSiText(siNodes.Item(i)))
		  Next
		  Return arr
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 436F6E636174656E617465207468652074657874206F66206576657279203C743E2064657363656E64616E74206F6620616E203C73693E206E6F646520E280942068616E646C657320626F74682073696E676C65203C743E20616E6420726963682D74657874203C723E3C743E2072756E732E0A
		Private Function JoinSiText(siNode As XmlNode) As String
		  ' <si> may be a single <t> or several <r><t>...</t></r> rich-text runs.
		  Var result As String = ""
		  Var ts As XmlNodeList = siNode.Xql(".//*[local-name()='t']")
		  For i As Integer = 0 To ts.Length - 1
		    Var n As XmlNode = ts.Item(i)
		    If n.FirstChild <> Nil Then result = result + n.FirstChild.Value
		  Next
		  Return result
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 526561642074686520776F726B73686565742773203C7461626C6550617274733E2C207265736F6C76652065616368207669612074686520776F726B73686565742072656C732C20616E6420617474616368207468652070617273656420457863656C205461626C657320746F207468652073686565742E2053696C656E74206F6E20616E79206D6973732E0A
		Private Sub LoadTablesForSheet(zip As XLSXZip, sheetPartPath As String, sheetXml As String, sheet As XLSXSheet)
		  ' Read the worksheet's <tableParts>, resolve each via the worksheet rels,
		  ' and attach the parsed Excel Tables to the sheet. Silent on any miss —
		  ' tables are decorative, so a malformed part must not fail the open.
		  If sheetXml = "" Then Return
		  Var doc As New XmlDocument
		  Try
		    doc.LoadXml(sheetXml)
		  Catch
		    Return
		  End Try
		  Var parts As XmlNodeList = doc.Xql("//*[local-name()='tableParts']/*[local-name()='tablePart']")
		  If parts.Length = 0 Then Return

		  ' Owner dir + rels path: xl/worksheets/sheet1.xml -> xl/worksheets/_rels/sheet1.xml.rels
		  Var segs() As String = sheetPartPath.Split("/")
		  Var fileName As String = segs(segs.LastIndex)
		  segs.RemoveAt(segs.LastIndex)
		  Var ownerDir As String = String.FromArray(segs, "/")
		  Var relsPath As String = ownerDir + "/_rels/" + fileName + ".rels"
		  Var relMap As Dictionary = ParseRelsToTargets(zip.ReadPart(relsPath))

		  For i As Integer = 0 To parts.Length - 1
		    Var rid As String = parts.Item(i).GetAttribute("r:id")
		    If rid = "" Then rid = parts.Item(i).GetAttribute("id")
		    If Not relMap.HasKey(rid) Then Continue
		    Var tablePath As String = ResolvePartPath(ownerDir, relMap.Value(rid))
		    Var tbl As XLSXTable = ParseTable(zip.ReadPart(tablePath))
		    If tbl <> Nil Then sheet.AddTable(tbl)
		  Next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 5061727365206F6E6520786C2F7461626C65732F7461626C654E2E786D6C20696E746F20616E20584C53585461626C65202872616E67652C206865616465722F746F74616C7320726F7720636F756E74732C206275696C742D696E207374796C65206E616D65202B2062616E64696E6720666C616773292E204E696C206F6E20616E792070726F626C656D2E0A
		Private Function ParseTable(tableXml As String) As XLSXTable
		  ' Parse one xl/tables/tableN.xml into an XLSXTable (range, header/totals row
		  ' counts, built-in style name + banding flags). Nil on any problem.
		  If tableXml = "" Then Return Nil
		  Var doc As New XmlDocument
		  Try
		    doc.LoadXml(tableXml)
		  Catch
		    Return Nil
		  End Try
		  Var tnodes As XmlNodeList = doc.Xql("//*[local-name()='table']")
		  If tnodes.Length = 0 Then Return Nil
		  Var tn As XmlNode = tnodes.Item(0)

		  Var fr, fc, lr, lc As Integer
		  If Not ParseRangeRef(tn.GetAttribute("ref"), fr, fc, lr, lc) Then Return Nil

		  Var name As String = tn.GetAttribute("displayName")
		  If name = "" Then name = tn.GetAttribute("name")
		  Var tbl As XLSXTable = XLSXTable.NewRaw(name, fr, fc, lr, lc)

		  Var hrc As String = tn.GetAttribute("headerRowCount")
		  If hrc <> "" Then tbl.HeaderRowCount = hrc.ToInteger   ' absent => default 1
		  Var trc As String = tn.GetAttribute("totalsRowCount")
		  If trc <> "" Then tbl.TotalsRowCount = trc.ToInteger

		  Var si As XmlNodeList = tn.Xql("./*[local-name()='tableStyleInfo']")
		  If si.Length > 0 Then
		    Var s As XmlNode = si.Item(0)
		    tbl.StyleName = s.GetAttribute("name")
		    tbl.ShowRowStripes = BoolAttr(s, "showRowStripes", False)
		    tbl.ShowColumnStripes = BoolAttr(s, "showColumnStripes", False)
		    tbl.ShowFirstColumn = BoolAttr(s, "showFirstColumn", False)
		    tbl.ShowLastColumn = BoolAttr(s, "showLastColumn", False)
		  End If
		  Return tbl
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 506172736520616E2041312D7374796C652072616E6765206C696B65202241313A503730312220286F7220612073696E676C652063656C6C2920696E746F20312D626173656420636F726E6572732E0A
		Private Function ParseRangeRef(ref As String, ByRef fr As Integer, ByRef fc As Integer, ByRef lr As Integer, ByRef lc As Integer) As Boolean
		  ' Parse an A1-style range like "A1:P701" (or a single cell) into 1-based corners.
		  Var parts() As String = ref.Split(":")
		  If parts.LastIndex < 0 Then Return False
		  Var r1, c1 As Integer
		  If Not XLSXCellRef.A1ToRowColRaw(parts(0), r1, c1) Then Return False
		  If parts.LastIndex = 0 Then
		    fr = r1
		    fc = c1
		    lr = r1
		    lc = c1
		    Return True
		  End If
		  Var r2, c2 As Integer
		  If Not XLSXCellRef.A1ToRowColRaw(parts(1), r2, c2) Then Return False
		  fr = Min(r1, r2)
		  fc = Min(c1, c2)
		  lr = Max(r1, r2)
		  lc = Max(c1, c2)
		  Return True
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 5265736F6C766520612072656C732054617267657420616761696E737420746865206F776E657220706172742773206469726563746F72792C206E6F726D616C697A696E67202E2F20616E64202E2E2F207365676D656E74733B2061206C656164696E67202F206D65616E73207061636B6167652D6162736F6C7574652E0A
		Private Function ResolvePartPath(ownerDir As String, target As String) As String
		  ' Resolve a rels Target (e.g. "../tables/table1.xml") against the owner part's
		  ' directory, normalizing "." / ".." segments. A leading "/" means package-absolute.
		  If target.BeginsWith("/") Then Return target.Middle(1)
		  Var combined As String = ownerDir + "/" + target
		  Var segs() As String = combined.Split("/")
		  Var outp() As String
		  For Each s As String In segs
		    If s = "" Or s = "." Then Continue
		    If s = ".." Then
		      If outp.LastIndex >= 0 Then outp.RemoveAt(outp.LastIndex)
		    Else
		      outp.Add s
		    End If
		  Next
		  Return String.FromArray(outp, "/")
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 52656164206120626F6F6C65616E20584D4C2061747472696275746520282231222F227472756522203D3E2054727565292C2072657475726E696E6720612064656661756C74207768656E20616273656E742E0A
		Private Function BoolAttr(node As XmlNode, name As String, defaultValue As Boolean) As Boolean
		  Var v As String = node.GetAttribute(name).Lowercase
		  If v = "" Then Return defaultValue
		  Return v = "1" Or v = "true"
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 506172736520786C2F5F72656C732F776F726B626F6F6B2E786D6C2E72656C7320696E746F20612044696374696F6E617279206B657965642062792052656C6174696F6E73686970204964202D3E2054617267657420706174682E0A
		Private Function ParseRelsToTargets(relsXml As String) As Dictionary
		  Var d As New Dictionary
		  If relsXml = "" Then Return d
		  Var doc As New XmlDocument
		  Try
		    doc.LoadXml(relsXml)
		  Catch
		    Return d
		  End Try
		  Var rels As XmlNodeList = doc.Xql("//*[local-name()='Relationships']/*[local-name()='Relationship']")
		  For i As Integer = 0 To rels.Length - 1
		    Var n As XmlNode = rels.Item(i)
		    Var id As String = n.GetAttribute("Id")
		    Var target As String = n.GetAttribute("Target")
		    If id <> "" And target <> "" Then d.Value(id) = target
		  Next
		  Return d
		End Function
	#tag EndMethod

	#tag Note, Name = About
		Top-level entry point for the XLSX parser.

		Two overloads:
		  Open(file As FolderItem) As XLSXWorkbook
		    For Desktop file dialogs.
		  Open(data As MemoryBlock, sourceName As String) As XLSXWorkbook
		    For Web uploads (WebFileUploader.UploadComplete provides bytes).

		Both return a fully-populated XLSXWorkbook (sharedStrings + styles +
		every sheet parsed eagerly). Callers can then iterate via SheetCount /
		SheetAt(i) and bind cells to Listbox / WebListbox.

		Errors raise XLSXException with one of XLSXEnums.eParseError:
		  NotAZip       - magic bytes wrong, or empty MemoryBlock
		  MissingPart   - file doesn't exist, or xl/workbook.xml absent
		  MalformedXML  - workbook.xml / sharedStrings.xml / sheetN.xml
		                  failed XmlDocument.LoadXml

		Pipeline (private OpenFromZip):
		  1. ZipReader extracts archive (XLSXZip).
		  2. ParseSharedStrings -> wb.SharedStrings()
		  3. New XLSXStyles(...) -> wb.Styles
		  4. ParseRelsToTargets reads xl/_rels/workbook.xml.rels (rId -> target).
		  5. For each <sheet> in workbook.xml, look up its rId target and
		     construct an XLSXSheet from xl/<target>.

		Out of V1 scope: defined names, external links, theme/colors, formulas
		(cached value only), encrypted workbooks.
	#tag EndNote

	#tag ViewBehavior
	#tag EndViewBehavior
End Module
#tag EndModule
