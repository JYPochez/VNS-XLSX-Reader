#tag Class
Protected Class XLSXCellRange
	#tag Method, Flags = &h0, Description = 4275696C6420612072616E67652066726F6D2069747320636F726E6572732028616C6C20312D62617365642C20696E636C7573697665292E0A
		Sub Constructor()
		  ' Empty range. Used by NewRaw; corners are filled in by the caller.
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 4275696C6420612072616E67652066726F6D2069747320636F726E6572732028616C6C20312D62617365642C20696E636C7573697665292E0A
		Sub Constructor(firstRow As Integer, firstCol As Integer, lastRow As Integer, lastCol As Integer)
		  ' Corners are given in the CALLER's index base (0-based when
		  ' gZeroBasedSheetsRowsColumns is set) and stored internally as 1-based.
		  mFirstRow = XLSXHelpers.ToInternalIndex(firstRow)
		  mFirstCol = XLSXHelpers.ToInternalIndex(firstCol)
		  mLastRow = XLSXHelpers.ToInternalIndex(lastRow)
		  mLastCol = XLSXHelpers.ToInternalIndex(lastCol)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 4275696C6420612072616E67652066726F6D20636F726E65727320746861742061726520616C726561647920696E7465726E616C20312D62617365642C206173207061727365727320616E642073657269616C697A65727320737570706C79207468656D2E0A
		Shared Function NewRaw(firstRow As Integer, firstCol As Integer, lastRow As Integer, lastCol As Integer) As XLSXCellRange
		  ' Build a range from corners that are ALREADY internal 1-based — the form
		  ' parsers and serializers work in. Never index-translated.
		  Var r As New XLSXCellRange
		  r.mFirstRow = firstRow
		  r.mFirstCol = firstCol
		  r.mLastRow = lastRow
		  r.mLastCol = lastCol
		  Return r
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 52657475726E7320547275652069662028726F772C20636F6C292066616C6C7320696E7369646520746869732072616E67652C20696E636C75736976652E0A
		Function Contains(row As Integer, col As Integer) As Boolean
		  ' (row, col) in the caller's index base.
		  Return ContainsRaw(XLSXHelpers.ToInternalIndex(row), XLSXHelpers.ToInternalIndex(col))
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 5472756520696620616E20616C726561647920312D62617365642028726F772C20636F6C292066616C6C7320696E7369646520746869732072616E67652C20696E636C75736976652E0A
		Function ContainsRaw(row As Integer, col As Integer) As Boolean
		  ' (row, col) already internal 1-based.
		  Return row >= mFirstRow And row <= mLastRow And col >= mFirstCol And col <= mLastCol
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 546F7020726F77206F66207468652072616E67652C20616C7761797320312D62617365642E0A
		Function FirstRowRaw() As Integer
		  ' Internal 1-based corner — what serializers must write to the file.
		  Return mFirstRow
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 4C65667420636F6C756D6E206F66207468652072616E67652C20616C7761797320312D62617365642E0A
		Function FirstColRaw() As Integer
		  Return mFirstCol
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 426F74746F6D20726F77206F66207468652072616E67652C20616C7761797320312D62617365642E0A
		Function LastRowRaw() As Integer
		  Return mLastRow
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 526967687420636F6C756D6E206F66207468652072616E67652C20616C7761797320312D62617365642E0A
		Function LastColRaw() As Integer
		  Return mLastCol
		End Function
	#tag EndMethod

	#tag ComputedProperty, Flags = &h0, Description = 546F7020726F77206F66207468652072616E67652028312D62617365642C20696E636C7573697665292E0A
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

	#tag ComputedProperty, Flags = &h0, Description = 4C65667420636F6C756D6E206F66207468652072616E67652028312D62617365642C20696E636C7573697665292E0A
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

	#tag ComputedProperty, Flags = &h0, Description = 426F74746F6D20726F77206F66207468652072616E67652028312D62617365642C20696E636C7573697665292E0A
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

	#tag ComputedProperty, Flags = &h0, Description = 526967687420636F6C756D6E206F66207468652072616E67652028312D62617365642C20696E636C7573697665292E0A
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

	#tag Note, Name = About
		An inclusive cell range (e.g. A1:C3).

		Used to represent merged-cell rectangles parsed from a worksheet's
		<mergeCells> element. Both corners are inclusive.

		Corners are stored internally as 1-based. FirstRow / FirstCol / LastRow /
		LastCol read and write in the CALLER's index base, which is 1-based unless
		XLSXHelpers.gZeroBasedSheetsRowsColumns is True. Parsers and serializers,
		which must always work in true 1-based file coordinates, use NewRaw plus
		the FirstRowRaw / FirstColRaw / LastRowRaw / LastColRaw accessors and
		ContainsRaw.
	#tag EndNote

	#tag ViewBehavior
	#tag EndViewBehavior
End Class
#tag EndClass
