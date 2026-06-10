#tag Module
Protected Module XLSXDesktopListboxFiller
	#tag Method, Flags = &h0, Description = 46696C6C2061204465736B746F704C697374426F782066726F6D20616E20584C535853686565742E2056696577696E67206D6F6465202864656661756C74293A206669727374206E6F6E2D656D70747920726F77206265636F6D657320746865206865616465722C20656D70747920726F777320736B69707065642E2047726964206D6F6465202873686F77416C6C43656C6C73293A20636F6C756D6E206C65747465727320617320686561646572732C20657665727920726F772073686F776E202D20666F722065646974696E672F6275696C64696E67207368656574732E0A
		Sub Fill(lb As DesktopListBox, sheet As XLSXSheet, styles As XLSXStyles, showAllCells As Boolean = False)
		  lb.RemoveAllRows
		  Var cols As Integer = Max(1, sheet.ColCount)
		  lb.ColumnCount = cols

		  ' Two display modes:
		  '  - viewing (default): first non-empty row becomes the header, rows
		  '    where every cell renders empty are skipped.
		  '  - grid (showAllCells): column letters as headers, every row shown —
		  '    used when editing / building a sheet from scratch.
		  Var firstDataRow As Integer
		  Var colMaxLen() As Integer

		  If showAllCells Then
		    For c As Integer = 1 To cols
		      Var h As String = XLSXCellRef.IndexToColLetters(c)
		      lb.HeaderAt(c - 1) = h
		      colMaxLen.Add Max(h.Length, 4)
		    Next
		    firstDataRow = 1
		  Else
		    Var headerRow As Integer = FindFirstNonEmptyRow(sheet)
		    If headerRow = 0 Then Return
		    For c As Integer = 1 To cols
		      Var h As String = sheet.CellAt(headerRow, c).DisplayText(styles)
		      lb.HeaderAt(c - 1) = h
		      colMaxLen.Add h.Length
		    Next
		    firstDataRow = headerRow + 1
		  End If

		  ' Body rows. Build texts first so viewing mode can skip rows where every
		  ' cell is empty (Excel often leaves styled-but-empty rows inflating RowCount).
		  For r As Integer = firstDataRow To sheet.RowCount
		    Var rowTexts() As String
		    Var anyNonEmpty As Boolean = False
		    For c As Integer = 1 To cols
		      Var text As String
		      If sheet.IsCellMergedFollower(r, c) Then
		        text = ""
		      Else
		        text = sheet.CellAt(r, c).DisplayText(styles)
		      End If
		      rowTexts.Add text
		      If text <> "" Then anyNonEmpty = True
		    Next
		    If Not showAllCells And Not anyNonEmpty Then Continue
		    lb.AddRow("")
		    Var lbRow As Integer = lb.LastAddedRowIndex
		    ' Remember which sheet row this listbox row shows, so cell edits can be
		    ' written back into the model (empty rows are skipped, so indexes differ).
		    lb.RowTagAt(lbRow) = r
		    For c As Integer = 0 To cols - 1
		      lb.CellTextAt(lbRow, c) = rowTexts(c)
		      Var n As Integer = rowTexts(c).Length
		      If n > colMaxLen(c) Then colMaxLen(c) = n
		    Next
		  Next

		  AutosizeColumnWidths(lb, colMaxLen)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 536574206C622E436F6C756D6E5769647468732066726F6D207065722D636F6C756D6E206D617820636861726163746572206C656E677468732C207363616C656420627920616E206176657261676520676C79706820776964746820616E642063617070656420746F20612073656E7369626C652072616E676520736F2073696E676C65206C6F6E672063656C6C7320646F6E277420626C6F77206F757420746865206C61796F75742E0A
		Private Sub AutosizeColumnWidths(lb As DesktopListBox, colMaxLen() As Integer)
		  ' Build a ColumnWidths string in points, sized roughly to the widest
		  ' cell content per column (header + body). Caps prevent runaway widths.
		  ' DesktopListBox.HasHorizontalScrollbar is True in the static layout,
		  ' so totals exceeding the visible width will scroll horizontally.
		  Const kMinPx As Integer = 60
		  Const kMaxPx As Integer = 400
		  Const kCharPx As Integer = 7    ' approximate average glyph width at the default font
		  Const kPaddingPx As Integer = 16
		  Var parts() As String
		  For Each n As Integer In colMaxLen
		    Var w As Integer = (n * kCharPx) + kPaddingPx
		    If w < kMinPx Then w = kMinPx
		    If w > kMaxPx Then w = kMaxPx
		    parts.Add Str(w)
		  Next
		  lb.ColumnWidths = String.FromArray(parts, ",")
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 5363616E2074686520666972737420353020726F7773206F662074686520736865657420616E642072657475726E2074686520666972737420726F7720696E64657820746861742068617320616E79206E6F6E2D656D7074792063656C6C2E2052657475726E73203120696620616C6C2070726F62656420726F77732061726520656D7074792E0A
		Private Function FindFirstNonEmptyRow(sheet As XLSXSheet) As Integer
		  Var maxProbe As Integer = Min(sheet.RowCount, 50)
		  For r As Integer = 1 To maxProbe
		    For c As Integer = 1 To sheet.ColCount
		      If Not sheet.CellAt(r, c).IsEmpty Then Return r
		    Next
		  Next
		  Return 1
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
