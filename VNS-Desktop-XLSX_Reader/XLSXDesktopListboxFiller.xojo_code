#tag Module
Protected Module XLSXDesktopListboxFiller
	#tag Method, Flags = &h0, Description = 46696C6C2061204465736B746F704C697374426F782066726F6D20616E20584C535853686565742E204120706C61696E20666972737420726F772069732070726F6D6F74656420746F2074686520636F6C756D6E2D686561646572206261723B2061207374796C656420666972737420726F7720286F722067726964206D6F646529207573657320412F422F43206865616465727320736F20746865207374796C656420726F772072656E6465727320617320646174612E20547261696C696E6720656D70747920726F7773207472696D6D65642028696E746572696F72206F6E6573206B6570742920756E6C65737320696E2067726964206D6F64652E0A
		Sub Fill(lb As DesktopListBox, sheet As XLSXSheet, styles As XLSXStyles, showAllCells As Boolean = False, showFormulas As Boolean = False)
		  lb.RemoveAllRows
		  Var cols As Integer = Max(1, sheet.ColCount)
		  lb.ColumnCount = cols

		  ' Header handling:
		  '  - promote (default for plain sheets): the first non-empty row becomes
		  '    the column-header bar; data starts on the next row.
		  '  - column-letters: A/B/C headers and every row shown as data. Used in
		  '    grid mode AND when the first row is styled (so its fill/font renders
		  '    as a normal styled row instead of being lost in the header bar).
		  ' Body trims only trailing empty rows (see below); interior ones are kept.
		  Var promoteHeader As Boolean = PromotesHeader(sheet, styles, showAllCells)
		  Var firstDataRow As Integer
		  Var colMaxLen() As Integer

		  If Not promoteHeader Then
		    For c As Integer = 1 To cols
		      Var h As String = XLSXCellRef.IndexToColLettersRaw(c)
		      lb.HeaderAt(c - 1) = h
		      colMaxLen.Add Max(h.Length, 4)
		    Next
		    firstDataRow = 1
		  Else
		    Var headerRow As Integer = FindFirstNonEmptyRow(sheet)
		    If headerRow = 0 Then Return
		    For c As Integer = 1 To cols
		      Var hc As XLSXCell = sheet.CellAtRaw(headerRow, c)
		      Var h As String = If(showFormulas And hc.HasFormula, hc.FormulaText, hc.DisplayText(styles))
		      lb.HeaderAt(c - 1) = h
		      colMaxLen.Add h.Length
		    Next
		    firstDataRow = headerRow + 1
		  End If

		  ' Body rows. Viewing mode trims only the *trailing* empty rows (Excel often
		  ' leaves styled-but-empty rows inflating RowCount); interior empty rows are
		  ' kept so the on-screen grid matches the file (and a saved copy of it).
		  Var lastRow As Integer = sheet.RowCount
		  If Not showAllCells Then lastRow = LastContentRow(sheet, cols, firstDataRow)
		  For r As Integer = firstDataRow To lastRow
		    Var rowTexts() As String
		    For c As Integer = 1 To cols
		      Var text As String
		      If sheet.IsCellMergedFollowerRaw(r, c) Then
		        text = ""
		      Else
		        Var dc As XLSXCell = sheet.CellAtRaw(r, c)
		        text = If(showFormulas And dc.HasFormula, dc.FormulaText, dc.DisplayText(styles))
		      End If
		      rowTexts.Add text
		    Next
		    lb.AddRow("")
		    Var lbRow As Integer = lb.LastAddedRowIndex
		    ' Remember which sheet row this listbox row shows, so cell edits can be
		    ' written back into the model (row indexes line up 1:1 in viewing mode).
		    lb.RowTagAt(lbRow) = r
		    For c As Integer = 0 To cols - 1
		      lb.CellTextAt(lbRow, c) = rowTexts(c)
		      Var n As Integer = rowTexts(c).Length
		      If n > colMaxLen(c) Then colMaxLen(c) = n
		      ' Stash a non-default visual style in the cell tag for the paint events
		      ' to render (merge followers stay blank/unstyled).
		      If Not sheet.IsCellMergedFollowerRaw(r, c + 1) Then
		        Var dataCell As XLSXCell = sheet.CellAtRaw(r, c + 1)
		        Var st As XLSXCellStyle = sheet.EffectiveStyleRaw(r, c + 1)
		        ' Excel right-aligns numbers/dates under General alignment; mirror that
		        ' (clone so the shared style isn't mutated).
		        Var numeric As Boolean = dataCell.eType = XLSXEnums.eCellType.Number _
		          Or dataCell.eType = XLSXEnums.eCellType.DateValue _
		          Or dataCell.eType = XLSXEnums.eCellType.FormulaCached
		        If numeric And st.AlignH = XLSXEnums.eAlignH.General Then
		          st = st.Clone
		          st.AlignH = XLSXEnums.eAlignH.Right
		        End If
		        If st <> Nil And Not st.IsDefault Then lb.CellTagAt(lbRow, c) = st
		      End If
		    Next
		  Next

		  AutosizeColumnWidths(lb, colMaxLen, sheet)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 536574206C622E436F6C756D6E5769647468732066726F6D207065722D636F6C756D6E206D617820636861726163746572206C656E677468732C207363616C656420627920616E206176657261676520676C79706820776964746820616E642063617070656420746F20612073656E7369626C652072616E676520736F2073696E676C65206C6F6E672063656C6C7320646F6E277420626C6F77206F757420746865206C61796F75742E0A
		Private Sub AutosizeColumnWidths(lb As DesktopListBox, colMaxLen() As Integer, sheet As XLSXSheet)
		  ' Build a ColumnWidths string in points. A column with an explicit width
		  ' from the workbook uses it verbatim; otherwise we fall back to a width
		  ' sized roughly to the widest cell content (caps prevent runaway widths).
		  ' DesktopListBox.HasHorizontalScrollbar is True in the static layout,
		  ' so totals exceeding the visible width will scroll horizontally.
		  Const kMinPx As Integer = 60
		  Const kMaxPx As Integer = 400
		  Const kCharPx As Integer = 7    ' approximate average glyph width at the default font
		  Const kPaddingPx As Integer = 16
		  Var widths() As Double
		  For j As Integer = 0 To colMaxLen.LastIndex
		    Var wpt As Double = sheet.ColumnWidthRaw(j + 1)   ' listbox col j -> sheet col j+1
		    If wpt > 0 Then
		      widths.Add XLSXHelpers.PointsToPixels(wpt)
		    Else
		      Var w As Integer = (colMaxLen(j) * kCharPx) + kPaddingPx
		      If w < kMinPx Then w = kMinPx
		      If w > kMaxPx Then w = kMaxPx
		      widths.Add w
		    End If
		  Next
		  ' The listbox can't scroll beyond the sum of its column widths, so the last
		  ' column's right edge ends up flush against the frame (partly under the
		  ' vertical scrollbar) with nothing left to scroll to. Give it slack.
		  ' MainWindow.CaptureColumnWidths subtracts the same amount, so this pad
		  ' never accumulates into the model or the saved file.
		  If widths.LastIndex >= 0 Then
		    widths(widths.LastIndex) = widths(widths.LastIndex) + LastColumnPadPixels
		  End If
		  Var parts() As String
		  For Each w As Double In widths
		    parts.Add Str(Round(w))
		  Next
		  lb.ColumnWidths = String.FromArray(parts, ",")
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 54686520736865657420726F77207573656420617320746865206C697374626F782068656164657220696E2076696577696E67206D6F646520286669727374206E6F6E2D656D70747920726F77292C20736F207468652077696E646F772063616E207265736F6C7665206865616465722D63656C6C207374796C657320666F72207061696E74696E672E0A
		Function HeaderRowIndex(sheet As XLSXSheet) As Integer
		  ' The sheet row used as the listbox header in viewing mode (the first
		  ' non-empty row). Lets the window resolve header-cell styles for painting.
		  Return FindFirstNonEmptyRow(sheet)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 54727565207768656E2074686520666972737420726F772073686F756C64206265636F6D652074686520636F6C756D6E2D686561646572206261723A206F6E6C7920666F72206120706C61696E2028756E7374796C65642920666972737420726F77206F7574736964652067726964206D6F64652E2041207374796C656420666972737420726F772073746179732061206461746120726F7720736F20697473207374796C696E672072656E646572732E0A
		Function PromotesHeader(sheet As XLSXSheet, styles As XLSXStyles, showAllCells As Boolean) As Boolean
		  ' True when the first row should become the column-header bar: only for a
		  ' plain (unstyled) first row outside grid mode. A styled first row stays a
		  ' data row so its styling renders (the header bar can't be styled on Web).
		  If showAllCells Then Return False
		  Return Not HeaderRowIsStyled(sheet, styles)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function HeaderRowIsStyled(sheet As XLSXSheet, styles As XLSXStyles) As Boolean
		  Var hr As Integer = FindFirstNonEmptyRow(sheet)
		  If hr <= 0 Then Return False
		  For c As Integer = 1 To sheet.ColCount
		    If Not sheet.EffectiveStyleRaw(hr, c).IsDefault Then Return True
		  Next
		  Return False
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 5363616E2074686520666972737420353020726F7773206F662074686520736865657420616E642072657475726E2074686520666972737420726F7720696E64657820746861742068617320616E79206E6F6E2D656D7074792063656C6C2E2052657475726E73203120696620616C6C2070726F62656420726F77732061726520656D7074792E0A
		Private Function FindFirstNonEmptyRow(sheet As XLSXSheet) As Integer
		  Var maxProbe As Integer = Min(sheet.RowCount, 50)
		  For r As Integer = 1 To maxProbe
		    For c As Integer = 1 To sheet.ColCount
		      If Not sheet.CellAtRaw(r, c).IsEmpty Then Return r
		    Next
		  Next
		  Return 1
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 4869676865737420726F772061742F616674657220666972737444617461526F7720686F6C64696E6720616E79206E6F6E2D656D7074792063656C6C3B206C657473207468652066696C6C6572207472696D20747261696C696E6720656D7074696573207768696C65206B656570696E6720696E746572696F72206F6E65732E2052657475726E7320666972737444617461526F772D31207768656E207468657265206973206E6F20626F647920636F6E74656E742E0A
		Private Function LastContentRow(sheet As XLSXSheet, cols As Integer, firstDataRow As Integer) As Integer
		  ' Highest row at/after firstDataRow that holds any non-empty cell. Lets the
		  ' filler trim trailing empty rows while keeping interior ones. Returns
		  ' firstDataRow - 1 when there is no body content at all.
		  For r As Integer = sheet.RowCount DownTo firstDataRow
		    For c As Integer = 1 To cols
		      If Not sheet.CellAtRaw(r, c).IsEmpty Then Return r
		    Next
		  Next
		  Return firstDataRow - 1
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function LastColumnPadPixels() As Integer
		  ' Slack added to the LAST column's listbox width so the grid can scroll
		  ' far enough to reveal it. Exposed so MainWindow.CaptureColumnWidths can
		  ' subtract it again and keep the model (and saved file) drift-free.
		  Return 24
		End Function
	#tag EndMethod

	#tag Note, Name = About
		Pours one XLSXSheet into a DesktopListBox.

		Public:
		  Fill(lb, sheet, styles)
		    - Resets the listbox (RemoveAllRows + sets ColumnCount).
		    - Uses the first non-empty row of the sheet as the header.
		    - For each subsequent row, writes one Listbox row using
		      cell.DisplayText(styles).
		    - Cells inside a merged range that are NOT the top-left anchor
		      render as empty so values don't appear duplicated.

		Lives in the Desktop project (not Common/) because DesktopListBox is
		a Desktop-only type. The Web project has its own filler at
		VNS-Web-XLSX-Reader/XLSXWebListboxFiller.xojo_code.
	#tag EndNote

	#tag ViewBehavior
	#tag EndViewBehavior
End Module
#tag EndModule
