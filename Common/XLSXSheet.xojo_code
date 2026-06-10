#tag Class
Protected Class XLSXSheet
	#tag Method, Flags = &h0, Description = 5061727365206F6E6520776F726B7368656574277320584D4C20696E746F2063656C6C73202B206D65726765642072616E6765732E0A
		Sub Constructor(name As String, tabIndex As Integer, sheetXml As String, sharedStrings() As String)
		  Me.Name = name
		  Me.TabIndex = tabIndex
		  mCells = New Dictionary
		  mMergeFollowers = New Dictionary
		  mMergeRanges = New Dictionary
		  mSharedStrings = sharedStrings
		  ParseSheetXml(sheetXml)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(name As String, tabIndex As Integer)
		  ' Empty sheet — no XLSX XML parsing. The caller populates it via PutCell
		  ' and AddMergedRange. Used by ODSReader (and any non-XLSX source) to build
		  ' the shared workbook model directly.
		  Me.Name = name
		  Me.TabIndex = tabIndex
		  mCells = New Dictionary
		  mMergeFollowers = New Dictionary
		  mMergeRanges = New Dictionary
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 496E736572742061207072652D6275696C7420584C535843656C6C206174206120312D62617365642028726F772C20636F6C292C207570646174696E672074686520736865657420657874656E742E2055736564206279206E6F6E2D584C535820736F757263657320284F4453526561646572292E0A
		Sub PutCell(row As Integer, col As Integer, cell As XLSXCell)
		  ' Insert a pre-built cell at a 1-based (row, col). Updates the sheet extent.
		  If row <= 0 Or col <= 0 Or cell Is Nil Then Return
		  Var key As Integer = (row * 16384) + col
		  mCells.Value(key) = cell
		  If row > mMaxRow Then mMaxRow = row
		  If col > mMaxCol Then mMaxCol = col
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 5265636F72642061206D65726765642072656374616E676C652028312D62617365642C20696E636C75736976652920616E6420666C616720657665727920636F76657265642063656C6C206578636570742074686520616E63686F722E0A
		Sub AddMergedRange(firstRow As Integer, firstCol As Integer, lastRow As Integer, lastCol As Integer)
		  ' Record a merged rectangle and flag every covered cell except the anchor.
		  If firstRow <= 0 Or firstCol <= 0 Or lastRow < firstRow Or lastCol < firstCol Then Return
		  Var range As New XLSXCellRange(firstRow, firstCol, lastRow, lastCol)
		  mMergeRanges.Value(mMergeRanges.KeyCount) = range
		  For r As Integer = firstRow To lastRow
		    For c As Integer = firstCol To lastCol
		      If r = firstRow And c = firstCol Then Continue
		      mMergeFollowers.Value(r.ToString + "," + c.ToString) = True
		    Next
		  Next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 47726F772074686520736865657420657874656E74206279206F6E6520656D70747920726F772061742074686520626F74746F6D20286E6F2063656C6C732073746F726564292E0A
		Sub AppendRow()
		  ' Grow the sheet extent by one (empty) row at the bottom. No cells are
		  ' stored — CellAt returns the empty sentinel until something is put there.
		  mMaxRow = mMaxRow + 1
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 44726F702074686520626F74746F6D20726F773A2064656C657465206974732073746F7265642063656C6C7320616E6420736872696E6B2074686520657874656E742E0A
		Sub RemoveLastRow()
		  ' Drop the bottom row: delete its stored cells and shrink the extent.
		  If mMaxRow < 1 Then Return
		  For c As Integer = 1 To mMaxCol
		    Var key As Integer = (mMaxRow * 16384) + c
		    If mCells.HasKey(key) Then mCells.Remove(key)
		  Next
		  mMaxRow = mMaxRow - 1
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 47726F772074686520736865657420657874656E74206279206F6E6520656D70747920636F6C756D6E2061742074686520726967687420286E6F2063656C6C732073746F726564292E0A
		Sub AppendColumn()
		  ' Grow the sheet extent by one (empty) column at the right.
		  mMaxCol = mMaxCol + 1
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 44726F70207468652072696768746D6F737420636F6C756D6E3A2064656C657465206974732073746F7265642063656C6C7320616E6420736872696E6B2074686520657874656E742E0A
		Sub RemoveLastColumn()
		  ' Drop the rightmost column: delete its stored cells and shrink the extent.
		  If mMaxCol < 1 Then Return
		  For r As Integer = 1 To mMaxRow
		    Var key As Integer = (r * 16384) + mMaxCol
		    If mCells.HasKey(key) Then mCells.Remove(key)
		  Next
		  mMaxCol = mMaxCol - 1
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 4869676865737420726F77206E756D626572207769746820612073746F7265642063656C6C2E0A
		Function RowCount() As Integer
		  Return mMaxRow
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 4869676865737420636F6C756D6E206E756D626572207769746820612073746F7265642063656C6C2E0A
		Function ColCount() As Integer
		  Return mMaxCol
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 52657475726E73207468652063656C6C2061742028726F772C20636F6C2920E2809420312D62617365642E204E65766572204E696C3B20616273656E742063656C6C732072657475726E20612073686172656420656D7074792073656E74696E656C2E0A
		Function CellAt(row As Integer, col As Integer) As XLSXCell
		  Var key As Integer = (row * 16384) + col
		  If mCells.HasKey(key) Then Return mCells.Value(key)
		  If mEmptyCell Is Nil Then mEmptyCell = New XLSXCell(XLSXEnums.eCellType.Empty, "", -1)
		  Return mEmptyCell
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 4E756D626572206F66203C6D6572676543656C6C3E2072616E676573207061727365642066726F6D20746869732073686565742E0A
		Function MergedRangeCount() As Integer
		  Return mMergeRanges.KeyCount
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 52657475726E732074686520692D7468206D65726765642072616E67652028302D6261736564292C206F72204E696C2069662069206973206F7574206F662072616E67652E0A
		Function MergedRangeAt(i As Integer) As XLSXCellRange
		  If Not mMergeRanges.HasKey(i) Then Return Nil
		  Return mMergeRanges.Value(i)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 52657475726E73205472756520666F722063656C6C7320696E736964652061206D65726765642072616E6765207468617420617265204E4F542074686520746F702D6C65667420616E63686F722E20557365207468697320746F206C656176652063656C6C7320626C616E6B207768656E2066696C6C696E672061204C697374626F782E0A
		Function IsCellMergedFollower(row As Integer, col As Integer) As Boolean
		  Var key As String = row.ToString + "," + col.ToString
		  Return mMergeFollowers.HasKey(key)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 50617273652074686520776F726B736865657420584D4C20E2809420726F77732C207468656E206D65726765642063656C6C732E0A
		Private Sub ParseSheetXml(sheetXml As String)
		  If sheetXml = "" Then Return
		  Var doc As New XmlDocument
		  Try
		    doc.LoadXml(sheetXml)
		  Catch
		    Raise New XLSXException(XLSXEnums.eParseError.MalformedXML, "sheet xml")
		  End Try
		  Var rows As XmlNodeList = doc.Xql("//*[local-name()='sheetData']/*[local-name()='row']")
		  For i As Integer = 0 To rows.Length - 1
		    ParseRow(rows.Item(i))
		  Next
		  ParseMergedCells(doc)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 5061727365206F6E65203C726F773E20656C656D656E7420E28094206974657261746520697473203C633E206368696C6472656E2E0A
		Private Sub ParseRow(rowNode As XmlNode)
		  Var rAttr As String = rowNode.GetAttribute("r")
		  Var rowIndex As Integer = If(rAttr <> "", Integer.FromString(rAttr), 0)
		  Var cells As XmlNodeList = rowNode.Xql("./*[local-name()='c']")
		  For i As Integer = 0 To cells.Length - 1
		    ParseCell(cells.Item(i), rowIndex)
		  Next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 5061727365206F6E65203C633E20656C656D656E7420E28094207265736F6C76657320746865207479706520636F64652C207468652076616C75652028696E636C7564696E67207368617265642D737472696E67206C6F6F6B7570292C20616E6420746865207374796C6520696E6465782E0A
		Private Sub ParseCell(cellNode As XmlNode, rowFallback As Integer)
		  Var refAttr As String = cellNode.GetAttribute("r")
		  Var typeAttr As String = cellNode.GetAttribute("t")
		  Var styleAttr As String = cellNode.GetAttribute("s")

		  Var row As Integer = rowFallback
		  Var col As Integer = 0
		  If refAttr <> "" Then
		    Call XLSXCellRef.A1ToRowCol(refAttr, row, col)
		  End If
		  If row <= 0 Or col <= 0 Then Return

		  Var styleIndex As Integer = If(styleAttr <> "", Integer.FromString(styleAttr), -1)

		  Var rawValue As String = ""
		  Var cellType As XLSXEnums.eCellType = XLSXEnums.eCellType.Empty

		  Var vList As XmlNodeList = cellNode.Xql("./*[local-name()='v']")
		  Var fList As XmlNodeList = cellNode.Xql("./*[local-name()='f']")
		  Var isList As XmlNodeList = cellNode.Xql("./*[local-name()='is']")
		  Var vNode As XmlNode = If(vList.Length > 0, vList.Item(0), Nil)
		  Var fNode As XmlNode = If(fList.Length > 0, fList.Item(0), Nil)
		  Var isNode As XmlNode = If(isList.Length > 0, isList.Item(0), Nil)

		  Select Case typeAttr
		  Case "s"
		    cellType = XLSXEnums.eCellType.Str
		    If vNode <> Nil And vNode.FirstChild <> Nil Then
		      Var idx As Integer = Integer.FromString(vNode.FirstChild.Value)
		      If idx >= 0 And idx <= mSharedStrings.LastIndex Then
		        rawValue = mSharedStrings(idx)
		      End If
		    End If
		  Case "b"
		    cellType = XLSXEnums.eCellType.Bool
		    If vNode <> Nil And vNode.FirstChild <> Nil Then rawValue = vNode.FirstChild.Value
		  Case "e"
		    cellType = XLSXEnums.eCellType.ErrorVal
		    If vNode <> Nil And vNode.FirstChild <> Nil Then rawValue = vNode.FirstChild.Value
		  Case "str", "inlineStr"
		    cellType = XLSXEnums.eCellType.Str
		    If isNode <> Nil Then
		      Var tList As XmlNodeList = isNode.Xql("./*[local-name()='t']")
		      If tList.Length > 0 And tList.Item(0).FirstChild <> Nil Then
		        rawValue = tList.Item(0).FirstChild.Value
		      End If
		    ElseIf vNode <> Nil And vNode.FirstChild <> Nil Then
		      rawValue = vNode.FirstChild.Value
		    End If
		  Else
		    If fNode <> Nil Then
		      cellType = XLSXEnums.eCellType.FormulaCached
		    ElseIf vNode <> Nil Then
		      cellType = XLSXEnums.eCellType.Number
		    Else
		      cellType = XLSXEnums.eCellType.Empty
		    End If
		    If vNode <> Nil And vNode.FirstChild <> Nil Then rawValue = vNode.FirstChild.Value
		  End Select

		  ' Skip truly-empty cells without a style — saves memory on sparse sheets.
		  If cellType = XLSXEnums.eCellType.Empty And styleIndex < 0 Then Return

		  Var cell As New XLSXCell(cellType, rawValue, styleIndex)
		  Var key As Integer = (row * 16384) + col
		  mCells.Value(key) = cell
		  If row > mMaxRow Then mMaxRow = row
		  If col > mMaxCol Then mMaxCol = col
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 5061727365203C6D6572676543656C6C733E2F3C6D6572676543656C6C3E20656E747269657320E280942073746F72652072616E6765732C206D61726B20666F6C6C6F7765722063656C6C732E0A
		Private Sub ParseMergedCells(doc As XmlDocument)
		  Var nodes As XmlNodeList = doc.Xql("//*[local-name()='mergeCells']/*[local-name()='mergeCell']")
		  For i As Integer = 0 To nodes.Length - 1
		    Var refAttr As String = nodes.Item(i).GetAttribute("ref")
		    If refAttr = "" Or refAttr.IndexOf(":") < 0 Then Continue
		    Var parts() As String = refAttr.Split(":")
		    If parts.Count <> 2 Then Continue
		    Var r1, c1, r2, c2 As Integer
		    If Not XLSXCellRef.A1ToRowCol(parts(0), r1, c1) Then Continue
		    If Not XLSXCellRef.A1ToRowCol(parts(1), r2, c2) Then Continue
		    Var range As New XLSXCellRange(r1, c1, r2, c2)
		    mMergeRanges.Value(i) = range
		    For r As Integer = r1 To r2
		      For c As Integer = c1 To c2
		        If r = r1 And c = c1 Then Continue
		        mMergeFollowers.Value(r.ToString + "," + c.ToString) = True
		      Next
		    Next
		  Next
		End Sub
	#tag EndMethod

	#tag Property, Flags = &h0, Description = 5368656574206E616D652066726F6D203C776F726B626F6F6B3E2F3C7368656574733E2F3C73686565743E406E616D652E0A
		Name As String
	#tag EndProperty

	#tag Property, Flags = &h0, Description = 312D626173656420706F736974696F6E206F66207468697320736865657420696E2074686520776F726B626F6F6B2E0A
		TabIndex As Integer
	#tag EndProperty

	#tag Property, Flags = &h21, Description = 44696374696F6E617279206B657965642062792028726F77202A20313633383429202B20636F6C202D3E20584C535843656C6C2E205370617273652E0A
		Private mCells As Dictionary
	#tag EndProperty

	#tag Property, Flags = &h21, Description = 536574206B657965642062792022726F772C636F6C22202D3E205472756520666F722063656C6C7320636F76657265642062792061206D65726765642072616E676520627574206E6F742069747320616E63686F722E0A
		Private mMergeFollowers As Dictionary
	#tag EndProperty

	#tag Property, Flags = &h21, Description = 4D617020496E746567657220696E646578202D3E20584C535843656C6C52616E67652E20536F75726365206F6620747275746820666F722074686520706172736564203C6D6572676543656C6C3E20656E74726965732E0A
		Private mMergeRanges As Dictionary
	#tag EndProperty

	#tag Property, Flags = &h21, Description = 4C61726765737420726F7720696E646578207365656E20647572696E672070617273652E0A
		Private mMaxRow As Integer = 0
	#tag EndProperty

	#tag Property, Flags = &h21, Description = 4C61726765737420636F6C756D6E20696E646578207365656E20647572696E672070617273652E0A
		Private mMaxCol As Integer = 0
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mSharedStrings() As String
	#tag EndProperty

	#tag Property, Flags = &h21, Description = 5368617265642073656E74696E656C2072657475726E65642062792043656C6C4174207768656E206E6F2063656C6C20657869737473206174207468617420706F736974696F6E2E0A
		Private mEmptyCell As XLSXCell
	#tag EndProperty

	#tag Note, Name = About
		One parsed worksheet from an XLSX archive.

		Construction:
		  New XLSXSheet(name, tabIndex, sheetXml, sharedStrings)
		    name           - sheet name from <workbook>/<sheets>/<sheet>@name
		    tabIndex       - 1-based position in the workbook
		    sheetXml       - UTF-8 text of xl/worksheets/sheetN.xml
		    sharedStrings  - the workbook's resolved sharedStrings array

		Cell access (1-based, like Excel):
		  CellAt(row, col)  -> XLSXCell  (never Nil; absent cells return a
		                                  shared empty sentinel)
		  RowCount, ColCount - max row / col with any value present

		Merged cells:
		  MergedRangeCount() As Integer
		  MergedRangeAt(i)   As XLSXCellRange
		  IsCellMergedFollower(row, col) As Boolean
		    True for cells inside a merged range that are NOT the top-left
		    anchor. Listbox fillers should leave those cells blank so values
		    don't appear duplicated visually.

		Internal storage: a Dictionary keyed by (row * 16384) + col -> XLSXCell.
		Sparse — only non-empty (or styled) cells are stored.

		Out of V1 scope: cell colors / fonts / borders, conditional formatting,
		formulas (cached value only is shown), pivots, charts, frozen panes
		(we read the data but do not preserve the visual freeze).
	#tag EndNote

	#tag ViewBehavior
	#tag EndViewBehavior
End Class
#tag EndClass
