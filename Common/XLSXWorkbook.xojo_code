#tag Class
Protected Class XLSXWorkbook
	#tag Method, Flags = &h0, Description = 4275696C6420616E20656D70747920776F726B626F6F6B2077697468206120736F75726365206E616D65202866696C656E616D65206F7220223C6D656D6F72793E22292E205368656574732061726520616464656420627920584C535852656164657220647572696E672070617273652E0A
		Sub Constructor(sourceName As String)
		  Me.SourceName = sourceName
		  mSheets = New Dictionary
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 417070656E6420612070617273656420736865657420617420746865206E65787420617661696C61626C6520696E6465782E205573656420696E7465726E616C6C7920627920584C53585265616465722E0A
		Sub AddSheet(sheet As XLSXSheet)
		  mSheets.Value(mSheets.KeyCount) = sheet
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function AddSheet(name As String) As XLSXSheet
		  ' Fluent authoring: create an empty sheet, append it at the next index,
		  ' and return it so the caller can start placing cells right away.
		  Var s As New XLSXSheet(name, mSheets.KeyCount + 1)
		  mSheets.Value(mSheets.KeyCount) = s
		  Return s
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 4E756D626572206F662073686565747320696E207468697320776F726B626F6F6B2E0A
		Function SheetCount() As Integer
		  Return mSheets.KeyCount
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 52657475726E732074686520692D74682073686565742028312D6261736564292E2052657475726E73204E696C20666F72206F75742D6F662D72616E676520696E64696365732E0A
		Function SheetAt(index As Integer) As XLSXSheet
		  ' 1-based for callers; internal storage 0-based.
		  If index < 1 Or index > mSheets.KeyCount Then Return Nil
		  Return mSheets.Value(index - 1)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 52657475726E73207468652066697273742073686565742077686F7365204E616D65206D6174636865732C206F72204E696C206966206E6F207368656574206973206E616D65642074686174207761792E0A
		Function SheetByName(name As String) As XLSXSheet
		  For i As Integer = 0 To mSheets.KeyCount - 1
		    Var s As XLSXSheet = mSheets.Value(i)
		    If s <> Nil And s.Name = name Then Return s
		  Next
		  Return Nil
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 52657475726E7320616C6C207368656574206E616D657320696E20776F726B626F6F6B206F726465722E0A
		Function SheetNames() As String()
		  Var arr() As String
		  For i As Integer = 0 To mSheets.KeyCount - 1
		    Var s As XLSXSheet = mSheets.Value(i)
		    If s <> Nil Then arr.Add(s.Name)
		  Next
		  Return arr
		End Function
	#tag EndMethod

	#tag Property, Flags = &h0, Description = 46696C656E616D6520286F7220223C6D656D6F72793E222920666F7220646961676E6F737469637320616E64205549207469746C6520646973706C61792E0A
		SourceName As String
	#tag EndProperty

	#tag Property, Flags = &h0
		SharedStrings() As String
	#tag EndProperty

	#tag Property, Flags = &h0, Description = 5468652070617273656420584C53585374796C657320666F722063656C6C5866202D3E20666F726D6174436F6465206C6F6F6B75702E0A
		Styles As XLSXStyles
	#tag EndProperty

	#tag Property, Flags = &h0, Description = 546865207265736F6C766564206261636B656E64207573656420627920584C53585265616465722E4F70656E20284D656D6F7279206F72204469736B3B206E65766572204175746F292E0A
		OpenMode As XLSXEnums.eOpenMode
	#tag EndProperty

	#tag Property, Flags = &h0, Description = 4D6963726F7365636F6E6473207370656E7420696E7369646520584C53585A69702E4F70656E20647572696E67207468697320776F726B626F6F6B27732070617273652E2055736566756C20746F20636F6D70617265204D656D6F7279207673204469736B206261636B656E64732E0A
		ZipMicroseconds As Double
	#tag EndProperty

	#tag Property, Flags = &h0, Description = 4D6963726F7365636F6E6473207370656E74206F6E20584D4C2070617273696E67202B20776F726B626F6F6B2F736865657420636F6E737472756374696F6E202865766572797468696E6720616674657220584C53585A69702E4F70656E292E204964656E746963616C20776F726B20696E20626F7468206261636B656E64732E0A
		XmlMicroseconds As Double
	#tag EndProperty

	#tag Property, Flags = &h21, Description = 496E7465726E616C2044696374696F6E617279206B6579656420627920302D626173656420696E646578202D3E20584C535853686565742E0A
		Private mSheets As Dictionary
	#tag EndProperty

	#tag Note, Name = About
		Top-level container for a parsed XLSX workbook.

		Built by XLSXReader.Open(...). Carries:
		  SourceName       - filename (or "<memory>") for diagnostics / UI title
		  SharedStrings()  - the workbook's resolved shared-string array
		  Styles           - the parsed XLSXStyles for cellXf -> formatCode lookup
		  SheetCount / SheetAt(i) / SheetByName(name) / SheetNames()

		Sheets are 1-based externally to match Excel's convention; the internal
		storage is 0-based. SheetAt returns Nil for out-of-range indices and
		SheetByName returns Nil if no sheet has that name.
	#tag EndNote

	#tag ViewBehavior
	#tag EndViewBehavior
End Class
#tag EndClass
