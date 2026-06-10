#tag Module
Protected Module XLSXHelpers
	#tag Method, Flags = &h0, Description = 457874656E73696F6E206D6574686F643A2072657475726E73207468652073796D626F6C6963206E616D65206F6620616E206543656C6C54797065202822456D707479222C2022537472222C20224E756D626572222C20224461746556616C7565222C2022426F6F6C222C2022466F726D756C61436163686564222C20224572726F7256616C22292E0A
		Function ToString(Extends t As XLSXEnums.eCellType) As String
		  Select Case t
		  Case XLSXEnums.eCellType.Empty
		    Return "Empty"
		  Case XLSXEnums.eCellType.Str
		    Return "Str"
		  Case XLSXEnums.eCellType.Number
		    Return "Number"
		  Case XLSXEnums.eCellType.DateValue
		    Return "DateValue"
		  Case XLSXEnums.eCellType.Bool
		    Return "Bool"
		  Case XLSXEnums.eCellType.FormulaCached
		    Return "FormulaCached"
		  Case XLSXEnums.eCellType.ErrorVal
		    Return "ErrorVal"
		  End Select
		  Return "Unknown(" + Integer(t).ToString + ")"
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 5061727365207468652073796D626F6C6963206E616D65206F6620616E206543656C6C547970652E2052657475726E73206543656C6C547970652E456D70747920666F7220756E6B6E6F776E20696E7075742E0A
		Function CellTypeFromString(s As String) As XLSXEnums.eCellType
		  Select Case s
		  Case "Empty"
		    Return XLSXEnums.eCellType.Empty
		  Case "Str"
		    Return XLSXEnums.eCellType.Str
		  Case "Number"
		    Return XLSXEnums.eCellType.Number
		  Case "DateValue"
		    Return XLSXEnums.eCellType.DateValue
		  Case "Bool"
		    Return XLSXEnums.eCellType.Bool
		  Case "FormulaCached"
		    Return XLSXEnums.eCellType.FormulaCached
		  Case "ErrorVal"
		    Return XLSXEnums.eCellType.ErrorVal
		  End Select
		  Return XLSXEnums.eCellType.Empty
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 457874656E73696F6E206D6574686F643A2072657475726E73207468652073796D626F6C6963206E616D65206F6620616E206550617273654572726F722028224E6F74415A6970222C20224D697373696E6750617274222C20224D616C666F726D6564584D4C222C2022456E63727970746564222C2022556E737570706F7274656422292E0A
		Function ToString(Extends e As XLSXEnums.eParseError) As String
		  Select Case e
		  Case XLSXEnums.eParseError.NotAZip
		    Return "NotAZip"
		  Case XLSXEnums.eParseError.MissingPart
		    Return "MissingPart"
		  Case XLSXEnums.eParseError.MalformedXML
		    Return "MalformedXML"
		  Case XLSXEnums.eParseError.Encrypted
		    Return "Encrypted"
		  Case XLSXEnums.eParseError.Unsupported
		    Return "Unsupported"
		  End Select
		  Return "Unknown(" + Integer(e).ToString + ")"
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 5061727365207468652073796D626F6C6963206E616D65206F6620616E206550617273654572726F722E2052657475726E73206550617273654572726F722E556E737570706F7274656420666F7220756E6B6E6F776E20696E7075742E0A
		Function ParseErrorFromString(s As String) As XLSXEnums.eParseError
		  Select Case s
		  Case "NotAZip"
		    Return XLSXEnums.eParseError.NotAZip
		  Case "MissingPart"
		    Return XLSXEnums.eParseError.MissingPart
		  Case "MalformedXML"
		    Return XLSXEnums.eParseError.MalformedXML
		  Case "Encrypted"
		    Return XLSXEnums.eParseError.Encrypted
		  Case "Unsupported"
		    Return XLSXEnums.eParseError.Unsupported
		  End Select
		  Return XLSXEnums.eParseError.Unsupported
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ToString(Extends m As XLSXEnums.eOpenMode) As String
		  Select Case m
		  Case XLSXEnums.eOpenMode.Auto
		    Return "Auto"
		  Case XLSXEnums.eOpenMode.Memory
		    Return "Memory"
		  Case XLSXEnums.eOpenMode.Disk
		    Return "Disk"
		  End Select
		  Return "Unknown(" + Integer(m).ToString + ")"
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 5061727365207468652073796D626F6C6963206E616D65206F6620616E20654F70656E4D6F64652E2052657475726E7320654F70656E4D6F64652E4175746F20666F7220756E6B6E6F776E20696E7075742E0A
		Function OpenModeFromString(s As String) As XLSXEnums.eOpenMode
		  Select Case s
		  Case "Auto"
		    Return XLSXEnums.eOpenMode.Auto
		  Case "Memory"
		    Return XLSXEnums.eOpenMode.Memory
		  Case "Disk"
		    Return XLSXEnums.eOpenMode.Disk
		  End Select
		  Return XLSXEnums.eOpenMode.Auto
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 4275696C6420616E20656D70747920696E2D6D656D6F727920776F726B626F6F6B2077697468206F6E6520736865657420616E64206120726F7773207820636F6C73206772696420657874656E742E205573656420627920746865204E657720627574746F6E733B207361766573207468726F75676820746865206E6F726D616C20777269746572732E0A
		Function NewWorkbook(sourceName As String, rows As Integer, cols As Integer) As XLSXWorkbook
		  ' Build an empty in-memory workbook with one sheet and a rows x cols grid
		  ' extent (no stored cells). Used by the "New" buttons to start a sheet
		  ' from scratch; the result saves through the normal writers.
		  Var wb As New XLSXWorkbook(sourceName)
		  wb.Styles = New XLSXStyles("")
		  Var sheet As New XLSXSheet(strings.kStrDefaultSheetName, 1)
		  For i As Integer = 1 To Max(rows, 1)
		    sheet.AppendRow
		  Next
		  For i As Integer = 1 To Max(cols, 1)
		    sheet.AppendColumn
		  Next
		  wb.AddSheet(sheet)
		  Return wb
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 45736361706520746865206669766520584D4C2D73656E736974697665206368617261637465727320666F7220656C656D656E74207465787420616E6420646F75626C652D71756F746564206174747269627574652076616C7565732E0A
		Function XmlEscape(s As String) As String
		  ' Escape the five XML-sensitive characters for use in element text and
		  ' double-quoted attribute values.
		  Var r As String = s.ReplaceAll("&", "&amp;")
		  r = r.ReplaceAll("<", "&lt;")
		  r = r.ReplaceAll(">", "&gt;")
		  r = r.ReplaceAll("""", "&quot;")
		  r = r.ReplaceAll("'", "&apos;")
		  Return r
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 54686520666F726D617420636F646520612063656C6C2072656E6465727320776974683A2061206469726563746C792D73657420466F726D6174436F64652077696E732C20656C7365206C6F6F6B207570205374796C65496E64657820696E20584C53585374796C65732E204D6972726F727320584C535843656C6C2E446973706C617954657874207265736F6C7574696F6E2E0A
		Function EffectiveFormatCode(cell As XLSXCell, styles As XLSXStyles) As String
		  ' The format code a cell renders with: a directly-set FormatCode wins
		  ' (ODS-loaded / edited cells), otherwise look up StyleIndex in XLSXStyles.
		  ' Mirrors the resolution order used by XLSXCell.DisplayText.
		  If cell Is Nil Then Return ""
		  If cell.FormatCode <> "" Then Return cell.FormatCode
		  If styles <> Nil And cell.StyleIndex >= 0 Then
		    Return styles.NumberFormatCodeAt(cell.StyleIndex)
		  End If
		  Return ""
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 54686520666F726D617420636F64652074686520777269746572732073686F756C6420706572736973743A20456666656374697665466F726D6174436F646520706C757320612064656661756C742049534F20646174652F74696D6520636F646520666F7220646174652063656C6C732077697468206E6F206578706C6963697420636F64652E0A
		Function WriteFormatCode(cell As XLSXCell, styles As XLSXStyles) As String
		  ' The format code the writers should persist for a cell. Same as
		  ' EffectiveFormatCode, plus a sensible default for date cells that carry
		  ' no explicit code (so a saved date doesn't degrade to a bare serial).
		  Var code As String = EffectiveFormatCode(cell, styles)
		  If code = "General" Then code = ""
		  If code = "" And cell <> Nil And cell.eType = XLSXEnums.eCellType.DateValue Then
		    Var d As Double = cell.NumberValue
		    If d < 1.0 Then
		      code = "hh:mm:ss"
		    ElseIf d - Floor(d) > 0.000001 Then
		      code = "yyyy-mm-dd hh:mm"
		    Else
		      code = "yyyy-mm-dd"
		    End If
		  End If
		  Return code
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 54727565207768656E2074686520737472696E6720706172736573206173206120706C61696E20646563696D616C206E756D6265722028646F7420736570617261746F722C206F7074696F6E616C207369676E2F6578706F6E656E74292E205573656420746F207479706520757365722D6564697465642063656C6C20746578742E0A
		Function IsNumericString(s As String) As Boolean
		  ' True when s parses as a plain decimal number ("42", "-3.14", "1.5e3").
		  ' Dot decimal separator only; used to type user-edited cell text.
		  Var t As String = s.Trim
		  Var n As Integer = t.Length
		  If n = 0 Then Return False
		  Var seenDigit As Boolean = False
		  Var seenDot As Boolean = False
		  Var seenExp As Boolean = False
		  Var signAllowed As Boolean = True
		  For i As Integer = 0 To n - 1
		    Var ch As String = t.Middle(i, 1)
		    Var c As Integer = Asc(ch)
		    If c >= 48 And c <= 57 Then
		      seenDigit = True
		      signAllowed = False
		    ElseIf (ch = "-" Or ch = "+") And signAllowed Then
		      signAllowed = False
		    ElseIf ch = "." And Not seenDot And Not seenExp Then
		      seenDot = True
		      signAllowed = False
		    ElseIf (ch = "e" Or ch = "E") And seenDigit And Not seenExp Then
		      seenExp = True
		      signAllowed = True
		    Else
		      Return False
		    End If
		  Next
		  ' Must end on a digit ("3.", "1e", "1e+" are rejected).
		  Var lastC As Integer = Asc(t.Middle(n - 1, 1))
		  Return seenDigit And lastC >= 48 And lastC <= 57
		End Function
	#tag EndMethod

	#tag Note, Name = About
		Project-wide helpers for the XLSX parser. Add new helper functions here rather
		than scattering them across the modules whose types they support.
		
		Conventions:
		  - enum -> string : extension method named ToString, declared
		                     `Function ToString(Extends e As <EnumType>) As String`.
		                     Callers write `someEnum.ToString` directly.
		  - string -> enum : regular function named <EnumName>FromString.
		
		Adding a new enum to XLSXEnums? Immediately add a matching ToString overload
		and FromString helper in this module.
	#tag EndNote

	#tag ViewBehavior
	#tag EndViewBehavior
End Module
#tag EndModule
