#tag Module
Protected Module XLSXCellRef
	#tag Method, Flags = &h0, Description = 436F6E7665727420636F6C756D6E206C65747465727320746F206120312D626173656420636F6C756D6E20696E6465782E20224122202D3E20312C20225A22202D3E2032362C2022414122202D3E2032372E2052657475726E732030206F6E20696E76616C696420696E7075742E0A
		Function ColLettersToIndex(letters As String) As Integer
		  ' Column letters -> index in the CALLER's base: "A" -> 1 normally, or 0
		  ' when gZeroBasedSheetsRowsColumns is set. Returns 0 on invalid input.
		  Var n As Integer = ColLettersToIndexRaw(letters)
		  If n = 0 Then Return 0
		  Return XLSXHelpers.ToPublicIndex(n)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 436F6C756D6E206C65747465727320746F20616E20616C7761797320312D626173656420696E6465783A204120697320312C2041412069732032372E2052657475726E732030206F6E20696E76616C696420696E7075742E0A
		Function ColLettersToIndexRaw(letters As String) As Integer
		  ' Always 1-based: "A" -> 1, "Z" -> 26, "AA" -> 27; 0 on invalid input.
		  ' Parsers and serializers use this, never the translating form.
		  Var n As Integer = 0
		  Var up As String = letters.Uppercase
		  For i As Integer = 0 To up.Length - 1
		    Var c As String = up.Middle(i, 1)
		    Var v As Integer = c.Asc - Asc("A") + 1
		    If v < 1 Or v > 26 Then Return 0
		    n = n * 26 + v
		  Next
		  Return n
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 436F6E76657274206120312D626173656420636F6C756D6E20696E64657820746F206C6574746572732E2031202D3E202241222C203237202D3E20224141222E2052657475726E7320656D70747920737472696E6720666F7220636F6C203C20312E0A
		Function IndexToColLetters(col As Integer) As String
		  ' Column index in the CALLER's base -> letters. 1 -> "A" normally; with
		  ' gZeroBasedSheetsRowsColumns set, 0 -> "A".
		  Return IndexToColLettersRaw(XLSXHelpers.ToInternalIndex(col))
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 416C7761797320312D626173656420636F6C756D6E20696E64657820746F206C6574746572733A203120697320412C2032372069732041412E20456D70747920737472696E672062656C6F7720312E0A
		Function IndexToColLettersRaw(col As Integer) As String
		  ' Always 1-based: 1 -> "A", 27 -> "AA". Empty string for col < 1.
		  If col < 1 Then Return ""
		  Var s As String = ""
		  Var n As Integer = col
		  While n > 0
		    Var r As Integer = (n - 1) Mod 26
		    s = Chr(Asc("A") + r) + s
		    n = (n - 1) \ 26
		  Wend
		  Return s
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 53706C697420616E204131207265666572656E636520282241423132222920696E746F20312D626173656420726F7720616E6420636F6C2E2052657475726E732046616C7365206F6E206D616C666F726D656420696E7075742E0A
		Function A1ToRowCol(a1 As String, ByRef row As Integer, ByRef col As Integer) As Boolean
		  ' Split "AB12" into row and col in the CALLER's index base.
		  If Not A1ToRowColRaw(a1, row, col) Then Return False
		  row = XLSXHelpers.ToPublicIndex(row)
		  col = XLSXHelpers.ToPublicIndex(col)
		  Return True
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 53706C697420616E204131207265666572656E636520696E746F20616E20616C7761797320312D626173656420726F7720616E6420636F6C756D6E2E2046616C7365206F6E206D616C666F726D656420696E7075742E0A
		Function A1ToRowColRaw(a1 As String, ByRef row As Integer, ByRef col As Integer) As Boolean
		  ' Split "AB12" into 1-based row and col; False on malformed input. A1
		  ' references inside a file are always 1-based, so parsers use this form.
		  Var letters As String = ""
		  Var digits As String = ""
		  For i As Integer = 0 To a1.Length - 1
		    Var c As String = a1.Middle(i, 1)
		    If c >= "0" And c <= "9" Then
		      digits = a1.Middle(i)
		      Exit For
		    Else
		      letters = letters + c
		    End If
		  Next
		  If letters = "" Or digits = "" Then Return False
		  col = ColLettersToIndexRaw(letters)
		  row = Integer.FromString(digits)
		  Return col > 0 And row > 0
		End Function
	#tag EndMethod

	#tag Note, Name = About
		Conversion helpers between Excel A1 cell references ("AB12") and
		(row, col) integers.

		Two families:
		  *Raw   - always 1-based, matching the file formats. Parsers and
		           serializers MUST use these:
		             ColLettersToIndexRaw("AA")      -> 27
		             IndexToColLettersRaw(27)        -> "AA"
		             A1ToRowColRaw("AB12", row, col) -> True; row=12, col=28
		  plain  - the caller's index base: 1-based unless
		           XLSXHelpers.gZeroBasedSheetsRowsColumns is True (then "A" -> 0).

		While the flag is False the two families behave identically, so existing
		behaviour is unchanged.

		ColLettersToIndexRaw returns 0 on invalid input rather than raising.
	#tag EndNote

	#tag ViewBehavior
	#tag EndViewBehavior
End Module
#tag EndModule
