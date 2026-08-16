#tag Class
Protected Class XLSXTable
	#tag Method, Flags = &h0, Description = 4275696C642061207461626C65206F76657220616E20696E636C757369766520312D62617365642072616E67652077697468206120646973706C6179206E616D652E0A
		Sub Constructor()
		  ' Empty table. Used by NewRaw; the range is filled in by the caller.
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 4275696C642061207461626C65206F76657220616E20696E636C757369766520312D62617365642072616E67652077697468206120646973706C6179206E616D652E0A
		Sub Constructor(name As String, firstRow As Integer, firstCol As Integer, lastRow As Integer, lastCol As Integer)
		  ' Corners are given in the CALLER's index base (0-based when
		  ' gZeroBasedSheetsRowsColumns is set) and stored internally as 1-based.
		  Me.Name = name
		  mFirstRow = XLSXHelpers.ToInternalIndex(firstRow)
		  mFirstCol = XLSXHelpers.ToInternalIndex(firstCol)
		  mLastRow = XLSXHelpers.ToInternalIndex(lastRow)
		  mLastCol = XLSXHelpers.ToInternalIndex(lastCol)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 4275696C642061207461626C652066726F6D20636F726E65727320746861742061726520616C726561647920696E7465726E616C20312D62617365642C206173207468652072656164657220737570706C696573207468656D2E0A
		Shared Function NewRaw(name As String, firstRow As Integer, firstCol As Integer, lastRow As Integer, lastCol As Integer) As XLSXTable
		  ' Build a table from corners that are ALREADY internal 1-based — the form
		  ' the reader parses out of xl/tables/tableN.xml. Never index-translated.
		  Var t As New XLSXTable
		  t.Name = name
		  t.mFirstRow = firstRow
		  t.mFirstCol = firstCol
		  t.mLastRow = lastRow
		  t.mLastCol = lastCol
		  Return t
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 54727565207768656E2028726F772C20636F6C292066616C6C7320696E7369646520746865207461626C652072616E67652C20696E636C75736976652E0A
		Function Contains(row As Integer, col As Integer) As Boolean
		  Return ContainsRaw(XLSXHelpers.ToInternalIndex(row), XLSXHelpers.ToInternalIndex(col))
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 5472756520696620616E20616C726561647920312D62617365642028726F772C20636F6C292066616C6C7320696E7369646520746865207461626C652072616E67652E0A
		Function ContainsRaw(row As Integer, col As Integer) As Boolean
		  ' (row, col) already internal 1-based.
		  Return row >= mFirstRow And row <= mLastRow And col >= mFirstCol And col <= mLastCol
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 5472756520666F7220746865207461626C6527732068656164657220726F77287329202874686520666972737420486561646572526F77436F756E7420726F7773292E0A
		Function IsHeaderRow(row As Integer) As Boolean
		  Return IsHeaderRowRaw(XLSXHelpers.ToInternalIndex(row))
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 5472756520666F7220746865207461626C6527732068656164657220726F772873292C20676976656E20616E20616C726561647920312D626173656420726F772E0A
		Function IsHeaderRowRaw(row As Integer) As Boolean
		  ' The first HeaderRowCount rows of the table are header rows.
		  Return row >= mFirstRow And row < mFirstRow + HeaderRowCount
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 5472756520666F7220746865207461626C65277320746F74616C7320726F772873292028746865206C61737420546F74616C73526F77436F756E7420726F7773292E0A
		Function IsTotalsRow(row As Integer) As Boolean
		  Return IsTotalsRowRaw(XLSXHelpers.ToInternalIndex(row))
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 5472756520666F7220746865207461626C65277320746F74616C7320726F772873292C20676976656E20616E20616C726561647920312D626173656420726F772E0A
		Function IsTotalsRowRaw(row As Integer) As Boolean
		  ' The last TotalsRowCount rows of the table are the totals row.
		  If TotalsRowCount <= 0 Then Return False
		  Return row > mLastRow - TotalsRowCount And row <= mLastRow
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 466972737420626F647920286E6F6E2D6865616465722920726F77206F6620746865207461626C652E0A
		Function FirstDataRow() As Integer
		  Return XLSXHelpers.ToPublicIndex(FirstDataRowRaw)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 466972737420626F647920726F77206F6620746865207461626C652C20616C7761797320312D62617365642E0A
		Function FirstDataRowRaw() As Integer
		  ' First body (non-header) row, internal 1-based.
		  Return mFirstRow + HeaderRowCount
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 546F7020726F77206F6620746865207461626C652072616E67652C20616C7761797320312D62617365642E0A
		Function FirstRowRaw() As Integer
		  ' Internal 1-based corners — what serializers must write to the file.
		  Return mFirstRow
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 4C65667420636F6C756D6E206F6620746865207461626C652072616E67652C20616C7761797320312D62617365642E0A
		Function FirstColRaw() As Integer
		  Return mFirstCol
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 426F74746F6D20726F77206F6620746865207461626C652072616E67652C20616C7761797320312D62617365642E0A
		Function LastRowRaw() As Integer
		  Return mLastRow
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 526967687420636F6C756D6E206F6620746865207461626C652072616E67652C20616C7761797320312D62617365642E0A
		Function LastColRaw() As Integer
		  Return mLastCol
		End Function
	#tag EndMethod

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  Return XLSXHelpers.ToPublicIndex(mFirstRow)
			End Get
		#tag EndGetter
		#tag Setter
			Set
			  mFirstRow = XLSXHelpers.ToInternalIndex(value)
			End Set
		#tag EndSetter
		FirstRow As Integer
	#tag EndComputedProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  Return XLSXHelpers.ToPublicIndex(mFirstCol)
			End Get
		#tag EndGetter
		#tag Setter
			Set
			  mFirstCol = XLSXHelpers.ToInternalIndex(value)
			End Set
		#tag EndSetter
		FirstCol As Integer
	#tag EndComputedProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  Return XLSXHelpers.ToPublicIndex(mLastRow)
			End Get
		#tag EndGetter
		#tag Setter
			Set
			  mLastRow = XLSXHelpers.ToInternalIndex(value)
			End Set
		#tag EndSetter
		LastRow As Integer
	#tag EndComputedProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  Return XLSXHelpers.ToPublicIndex(mLastCol)
			End Get
		#tag EndGetter
		#tag Setter
			Set
			  mLastCol = XLSXHelpers.ToInternalIndex(value)
			End Set
		#tag EndSetter
		LastCol As Integer
	#tag EndComputedProperty

	#tag Property, Flags = &h0
		Name As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mFirstRow As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mFirstCol As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mLastRow As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mLastCol As Integer
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

		Corners are stored internally as 1-based. FirstRow / FirstCol / LastRow /
		LastCol, Contains, IsHeaderRow, IsTotalsRow and FirstDataRow all speak the
		CALLER's index base, which is 1-based unless
		XLSXHelpers.gZeroBasedSheetsRowsColumns is True. The reader and the writer,
		which must always work in true 1-based file coordinates, use NewRaw plus
		the *Raw accessors.
	#tag EndNote

	#tag ViewBehavior
	#tag EndViewBehavior
End Class
#tag EndClass
