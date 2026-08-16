#tag Class
Protected Class XLSXSheetIterator
Implements Iterator
	#tag Method, Flags = &h0, Description = 537461727420616E20697465726174696F6E206F766572206120776F726B626F6F6B2773207368656574732C20706F736974696F6E6564206265666F7265207468652066697273742073686565742E0A
		Sub Constructor(wb As XLSXWorkbook)
		  ' Positioned BEFORE the first sheet; the first MoveNext lands on it.
		  mWorkbook = wb
		  mIndex = 0
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 4974657261746F7220696E746572666163653A20616476616E636520746F20746865206E6578742073686565742E2046616C7365206F6E636520746865206C61737420736865657420686173206265656E207969656C6465642E0A
		Function MoveNext() As Boolean
		  ' Iterator: advance; False once the last sheet has been yielded.
		  If mWorkbook Is Nil Then Return False
		  mIndex = mIndex + 1
		  Return mIndex <= mWorkbook.SheetCount
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 4974657261746F7220696E746572666163653A20746865207368656574206174207468652063757272656E7420706F736974696F6E2E204F6E6C79206D65616E696E6766756C206166746572204D6F76654E6578742072657475726E656420547275652E0A
		Function Value() As Variant
		  ' Iterator: the sheet at the current position. Walks the workbook with the
		  ' always-1-based accessor, so iteration is unaffected by the public index
		  ' base (gZeroBasedSheetsRowsColumns).
		  If mWorkbook Is Nil Then Return Nil
		  Return mWorkbook.SheetAtRaw(mIndex)
		End Function
	#tag EndMethod

	#tag Property, Flags = &h21
		Private mWorkbook As XLSXWorkbook
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mIndex As Integer = 0
	#tag EndProperty

	#tag Note, Name = About
		Iterator over a workbook's sheets, in workbook order.

		You do not create this directly — XLSXWorkbook implements Iterable, so:

		  For Each sheet As XLSXSheet In workbook
		    System.DebugLog sheet.Name
		  Next

		Implements Xojo's Iterator interface (MoveNext / Value). As the framework
		requires, Value is only meaningful after a MoveNext that returned True.
		Adding or removing sheets while iterating is not supported.
	#tag EndNote

	#tag ViewBehavior
	#tag EndViewBehavior
End Class
#tag EndClass
