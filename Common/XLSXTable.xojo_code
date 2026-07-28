#tag Class
Protected Class XLSXTable
	#tag Method, Flags = &h0, Description = 4275696C642061207461626C65206F76657220616E20696E636C757369766520312D62617365642072616E67652077697468206120646973706C6179206E616D652E0A
		Sub Constructor(name As String, firstRow As Integer, firstCol As Integer, lastRow As Integer, lastCol As Integer)
		  Me.Name = name
		  Me.FirstRow = firstRow
		  Me.FirstCol = firstCol
		  Me.LastRow = lastRow
		  Me.LastCol = lastCol
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 54727565207768656E2028726F772C20636F6C292066616C6C7320696E7369646520746865207461626C652072616E67652C20696E636C75736976652E0A
		Function Contains(row As Integer, col As Integer) As Boolean
		  Return row >= FirstRow And row <= LastRow And col >= FirstCol And col <= LastCol
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 5472756520666F7220746865207461626C6527732068656164657220726F77287329202874686520666972737420486561646572526F77436F756E7420726F7773292E0A
		Function IsHeaderRow(row As Integer) As Boolean
		  ' The first HeaderRowCount rows of the table are header rows.
		  Return row >= FirstRow And row < FirstRow + HeaderRowCount
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 5472756520666F7220746865207461626C65277320746F74616C7320726F772873292028746865206C61737420546F74616C73526F77436F756E7420726F7773292E0A
		Function IsTotalsRow(row As Integer) As Boolean
		  ' The last TotalsRowCount rows of the table are the totals row.
		  If TotalsRowCount <= 0 Then Return False
		  Return row > LastRow - TotalsRowCount And row <= LastRow
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 466972737420626F647920286E6F6E2D6865616465722920726F77206F6620746865207461626C652E0A
		Function FirstDataRow() As Integer
		  ' First body (non-header) row of the table.
		  Return FirstRow + HeaderRowCount
		End Function
	#tag EndMethod

	#tag Property, Flags = &h0
		Name As String
	#tag EndProperty

	#tag Property, Flags = &h0
		FirstRow As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		FirstCol As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		LastRow As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		LastCol As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		HeaderRowCount As Integer = 1
	#tag EndProperty

	#tag Property, Flags = &h0
		TotalsRowCount As Integer = 0
	#tag EndProperty

	#tag Property, Flags = &h0
		StyleName As String
	#tag EndProperty

	#tag Property, Flags = &h0
		ShowRowStripes As Boolean = True
	#tag EndProperty

	#tag Property, Flags = &h0
		ShowColumnStripes As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h0
		ShowFirstColumn As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h0
		ShowLastColumn As Boolean = False
	#tag EndProperty

	#tag Note, Name = About
		An Excel Table (ListObject) overlaid on a worksheet range, parsed from
		xl/tables/tableN.xml. It carries the table's range, header/totals row
		counts, the built-in table style name (e.g. "TableStyleLight9"), and the
		banding flags from <tableStyleInfo>.

		The visual styling a table implies (coloured header, banded rows) is not
		stored per cell in styles.xml — it is generated from the style name and
		the workbook theme at render time. XLSXStyles.TableStyleCell turns a
		(style name, role) into an XLSXCellStyle so the existing paint path
		renders it; XLSXSheet.EffectiveStyle applies it to cells inside the range
		that have no explicit style of their own.
	#tag EndNote

	#tag ViewBehavior
	#tag EndViewBehavior
End Class
#tag EndClass
