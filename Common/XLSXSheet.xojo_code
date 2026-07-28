#tag Class
Protected Class XLSXSheet
	#tag Method, Flags = &h0, Description = 5061727365206F6E6520776F726B7368656574277320584D4C20696E746F2063656C6C73202B206D65726765642072616E6765732E0A
		Sub Constructor(name As String, tabIndex As Integer, sheetXml As String, sharedStrings() As String, styles As XLSXStyles = Nil, date1904 As Boolean = False)
		  Me.Name = name
		  Me.TabIndex = tabIndex
		  mCells = New Dictionary
		  mMergeFollowers = New Dictionary
		  mMergeRanges = New Dictionary
		  mColWidths = New Dictionary
		  mRowHeights = New Dictionary
		  mSharedStrings = sharedStrings
		  mStyles = styles
		  mDate1904 = date1904
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
		  mColWidths = New Dictionary
		  mRowHeights = New Dictionary
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 5769647468206F66206120312D626173656420636F6C756D6E20696E20706F696E74732C206F7220302069662069742075736573207468652073686565742064656661756C742E20506F696E7473206973207468652063616E6F6E6963616C20776964746820756E6974206163726F737320626F746820666F726D6174732E0A
		Function ColumnWidth(col As Integer) As Double
		  ' Width of a 1-based column in points, or 0 if the column uses the
		  ' sheet default. Canonical unit is points across both formats.
		  If mColWidths.HasKey(col) Then Return mColWidths.Value(col)
		  Return 0.0
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 53657420286F7220636C6561722C207768656E203C3D203029206120312D626173656420636F6C756D6E277320776964746820696E20706F696E74732E0A
		Sub SetColumnWidth(col As Integer, widthPoints As Double)
		  ' Set (or clear, when widthPoints <= 0) a 1-based column's width in points.
		  If col <= 0 Then Return
		  If widthPoints <= 0 Then
		    If mColWidths.HasKey(col) Then mColWidths.Remove(col)
		  Else
		    mColWidths.Value(col) = widthPoints
		  End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 486569676874206F66206120312D626173656420726F7720696E20706F696E74732C206F7220302069662069742075736573207468652073686565742064656661756C742E0A
		Function RowHeight(row As Integer) As Double
		  ' Height of a 1-based row in points, or 0 if the row uses the sheet default.
		  If mRowHeights.HasKey(row) Then Return mRowHeights.Value(row)
		  Return 0.0
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 53657420286F7220636C6561722C207768656E203C3D203029206120312D626173656420726F7727732068656967687420696E20706F696E74732E0A
		Sub SetRowHeight(row As Integer, heightPoints As Double)
		  ' Set (or clear, when heightPoints <= 0) a 1-based row's height in points.
		  If row <= 0 Then Return
		  If heightPoints <= 0 Then
		    If mRowHeights.HasKey(row) Then mRowHeights.Remove(row)
		  Else
		    mRowHeights.Value(row) = heightPoints
		  End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 5472756520696620616E7920636F6C756D6E2068617320616E206578706C6963697420776964746820286C65747320777269746572732F66696C6C65727320736B697020746865203C636F6C733E20776F726B206F7468657277697365292E0A
		Function HasColumnWidths() As Boolean
		  Return mColWidths.KeyCount > 0
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 5472756520696620616E7920726F772068617320616E206578706C69636974206865696768742E0A
		Function HasRowHeights() As Boolean
		  Return mRowHeights.KeyCount > 0
		End Function
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

	#tag Method, Flags = &h0
		Function Cell(row As Integer, col As Integer) As XLSXCell
		  ' Getter half of the indexed accessor: ws.Cell(r, c). Never Nil.
		  Return CellAt(row, col)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Cell(row As Integer, col As Integer, Assigns c As XLSXCell)
		  ' Setter half: ws.Cell(r, c) = someCell.
		  PutCell(row, col, c)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function PutText(row As Integer, col As Integer, text As String) As XLSXCell
		  ' Place a string cell and return it for fluent chaining (.Bold etc.).
		  Var c As XLSXCell = XLSXCell.TextCell(text)
		  PutCell(row, col, c)
		  Return c
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function PutNumber(row As Integer, col As Integer, value As Double) As XLSXCell
		  Var c As XLSXCell = XLSXCell.NumberCell(value)
		  PutCell(row, col, c)
		  Return c
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function PutMoney(row As Integer, col As Integer, value As Double) As XLSXCell
		  Var c As XLSXCell = XLSXCell.MoneyCell(value)
		  PutCell(row, col, c)
		  Return c
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function PutDate(row As Integer, col As Integer, dt As DateTime) As XLSXCell
		  Var c As XLSXCell = XLSXCell.DateCell(dt)
		  PutCell(row, col, c)
		  Return c
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function PutDateTime(row As Integer, col As Integer, dt As DateTime) As XLSXCell
		  Var c As XLSXCell = XLSXCell.DateTimeCell(dt)
		  PutCell(row, col, c)
		  Return c
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function PutBool(row As Integer, col As Integer, value As Boolean) As XLSXCell
		  Var c As XLSXCell = XLSXCell.BoolCell(value)
		  PutCell(row, col, c)
		  Return c
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function PutFormula(row As Integer, col As Integer, formula As String) As XLSXCell
		  ' Place an R1C1 formula (converted to A1 at write time). Returns it for chaining.
		  Var c As XLSXCell = XLSXCell.FormulaCell(formula)
		  PutCell(row, col, c)
		  Return c
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function PutFormulaA1(row As Integer, col As Integer, formula As String) As XLSXCell
		  ' Place an A1 formula (written verbatim). Returns it for chaining.
		  Var c As XLSXCell = XLSXCell.FormulaCellA1(formula)
		  PutCell(row, col, c)
		  Return c
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

	#tag Method, Flags = &h0, Description = 526567697374657220616E20457863656C205461626C6520284C6973744F626A65637429206F7665726C616964206F6E20746869732073686565742E0A
		Sub AddTable(t As XLSXTable)
		  ' Register an Excel Table (ListObject) overlaid on this sheet.
		  mTables.Add t
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 4E756D626572206F6620457863656C205461626C657320617474616368656420746F20746869732073686565742E0A
		Function TableCount() As Integer
		  Return mTables.Count
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 52657475726E2074686520692D746820457863656C205461626C652028302D6261736564292C206F72204E696C206966206F7574206F662072616E67652E0A
		Function TableAt(i As Integer) As XLSXTable
		  If i < 0 Or i > mTables.LastIndex Then Return Nil
		  Return mTables(i)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 52656E646572207374796C6520666F722028726F772C20636F6C293A20616E206578706C696369742063656C6C207374796C652077696E732C20656C736520616E20656E636C6F73696E6720457863656C205461626C652773206275696C742D696E207374796C6520737570706C6965732061206865616465722066696C6C202F2062616E6465642D726F772074696E742E204E65766572204E696C2E0A
		Function EffectiveStyle(row As Integer, col As Integer) As XLSXCellStyle
		  ' The style used to render (row, col). Inside an Excel Table, the table's
		  ' header fill / banded-row tint applies — but a cell's own explicit fill (a
		  ' deliberate background) always wins over the table band. A bare default
		  ' font / number-format on the cell must NOT suppress the table style, so we
		  ' key only on whether the cell brings a background of its own.
		  Var own As XLSXCellStyle = CellAt(row, col).ResolvedStyle(mStyles)
		  If mStyles <> Nil And Not own.HasBackground Then
		    Var ts As XLSXCellStyle = TableStyleForCell(row, col)
		    If ts <> Nil Then Return ts
		  End If
		  Return own
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 49662028726F772C20636F6C2920697320696E736964652061207461626C652C20746865207461626C652D7374796C652063656C6C20666F722069747320726F6C652028686561646572202F2073747269706564202F20706C61696E293B204E696C206F74686572776973652E0A
		Private Function TableStyleForCell(row As Integer, col As Integer) As XLSXCellStyle
		  ' If (row, col) is inside a table, return the table-style cell for its role
		  ' (header / striped body / plain). Nil when no table covers the cell or the
		  ' role carries no styling.
		  For Each t As XLSXTable In mTables
		    If Not t.Contains(row, col) Then Continue
		    Var isHeader As Boolean = t.IsHeaderRow(row)
		    Var striped As Boolean = False
		    If Not isHeader And t.ShowRowStripes Then
		      ' Band every other body row (the first data row is unstriped).
		      striped = ((row - t.FirstDataRow) Mod 2) = 1
		    End If
		    Var st As XLSXCellStyle = mStyles.TableStyleCell(t.StyleName, isHeader, striped)
		    If st <> Nil And Not st.IsDefault Then Return st
		    Return Nil
		  Next
		  Return Nil
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
		  ParseColumns(doc)   ' after rows: mMaxCol is known, so we can clamp ranges
		  ParseMergedCells(doc)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 5061727365203C636F6C733E2F3C636F6C3E207769647468732028457863656C2063686172616374657220756E697473202D3E20706F696E7473292C20636C616D70656420746F207468652073686565742773207573656420636F6C756D6E20636F756E742E0A
		Private Sub ParseColumns(doc As XmlDocument)
		  ' <cols><col min max width customWidth/></cols> — column widths in Excel
		  ' "character" units, converted to points. A <col> range is clamped to the
		  ' sheet's used column count so a default "all columns" entry (max=16384)
		  ' doesn't create thousands of width entries.
		  Var cols As XmlNodeList = doc.Xql("//*[local-name()='cols']/*[local-name()='col']")
		  For i As Integer = 0 To cols.Length - 1
		    Var node As XmlNode = cols.Item(i)
		    Var w As String = node.GetAttribute("width")
		    If w = "" Then Continue
		    Var widthPt As Double = XLSXHelpers.ColumnCharsToPoints(w.ToDouble)
		    If widthPt <= 0 Then Continue
		    Var minAttr As String = node.GetAttribute("min")
		    Var maxAttr As String = node.GetAttribute("max")
		    Var minC As Integer = If(minAttr <> "", minAttr.ToInteger, 0)
		    Var maxC As Integer = If(maxAttr <> "", maxAttr.ToInteger, minC)
		    If minC <= 0 Then Continue
		    Var hi As Integer = maxC
		    If hi <= 0 Or hi > mMaxCol Then hi = mMaxCol
		    For c As Integer = minC To hi
		      SetColumnWidth(c, widthPt)
		    Next
		  Next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 5061727365206F6E65203C726F773E20656C656D656E7420E28094206974657261746520697473203C633E206368696C6472656E2E0A
		Private Sub ParseRow(rowNode As XmlNode)
		  Var rAttr As String = rowNode.GetAttribute("r")
		  Var rowIndex As Integer = If(rAttr <> "", Integer.FromString(rAttr), 0)
		  ' Custom row height (in points) when present.
		  Var htAttr As String = rowNode.GetAttribute("ht")
		  If htAttr <> "" And rowIndex > 0 Then SetRowHeight(rowIndex, htAttr.ToDouble)
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

		  ' 1904 date system: shift date serials into the 1900 system the rest of the
		  ' model uses (+1462 days), so nothing downstream needs to know the epoch.
		  ' Guarded by mDate1904, so non-1904 files (the norm) pay nothing.
		  If mDate1904 And rawValue <> "" And mStyles <> Nil And styleIndex >= 0 Then
		    If cellType = XLSXEnums.eCellType.Number Or cellType = XLSXEnums.eCellType.FormulaCached Then
		      If XLSXFormatter.IsDateFormatCode(mStyles.NumberFormatCodeAt(styleIndex)) Then
		        rawValue = Str(rawValue.ToDouble + 1462.0)
		      End If
		    End If
		  End If

		  Var cell As New XLSXCell(cellType, rawValue, styleIndex)
		  ' Keep the formula text (A1 notation, as stored in the file) so it survives
		  ' a round-trip instead of collapsing to its cached value on save.
		  If cellType = XLSXEnums.eCellType.FormulaCached And fNode <> Nil And fNode.FirstChild <> Nil Then
		    cell.Formula = fNode.FirstChild.Value
		  End If
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

	#tag Property, Flags = &h21, Description = 4D617020636F6C2028496E74656765722C20312D626173656429202D3E20776964746820696E20706F696E74732E205370617273653B20616273656E74203D2064656661756C742E0A
		Private mColWidths As Dictionary
	#tag EndProperty

	#tag Property, Flags = &h21, Description = 4D617020726F772028496E74656765722C20312D626173656429202D3E2068656967687420696E20706F696E74732E205370617273653B20616273656E74203D2064656661756C742E0A
		Private mRowHeights As Dictionary
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mStyles As XLSXStyles
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mDate1904 As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h21, Description = 5368617265642073656E74696E656C2072657475726E65642062792043656C6C4174207768656E206E6F2063656C6C20657869737473206174207468617420706F736974696F6E2E0A
		Private mEmptyCell As XLSXCell
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mTables() As XLSXTable
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
