#tag Module
Protected Module XLSXFormatter
	#tag Method, Flags = &h0, Description = 436F6E7665727420616E20457863656C20646174652073657269616C206E756D62657220746F2061204461746554696D6520696E20746865206C6F63616C2074696D657A6F6E652E20457863656C2065706F636820697320313839392D31322D333020746F206D6174636820457863656C2773204C6F7475732D312D322D332D6C6561702D796561722D627567206265686176696F722E0A
		Function ExcelSerialToDateTime(serial As Double) As DateTime
		  ' Excel epoch is 1899-12-30 00:00:00 (compensates for Lotus 1-2-3 leap-year bug).
		  Var epoch As New DateTime(1899, 12, 30, 0, 0, 0, 0, TimeZone.Current)
		  Var seconds As Double = serial * 86400.0
		  Return New DateTime(epoch.SecondsFrom1970 + seconds, TimeZone.Current)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 496E7665727365206F6620457863656C53657269616C546F4461746554696D653A20636F6E766572742061204461746554696D6520746F20616E20457863656C2073657269616C2028646179732073696E636520313839392D31322D33302C206672616374696F6E616C2070617274203D2074696D65206F6620646179292E0A
		Function DateTimeToExcelSerial(dt As DateTime) As Double
		  ' Inverse of ExcelSerialToDateTime: a DateTime → Excel serial (days since
		  ' 1899-12-30, with the fractional part being the time of day). Used to feed
		  ' ODS ISO dates into the Excel-serial-based formatting path.
		  If dt Is Nil Then Return 0.0
		  Var epoch As New DateTime(1899, 12, 30, 0, 0, 0, 0, TimeZone.Current)
		  Return (dt.SecondsFrom1970 - epoch.SecondsFrom1970) / 86400.0
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 4465746563742077686574686572206120666F726D617420636F6465206361727269657320646174652F74696D6520746F6B656E732028792F6D2F642F682F7329206F7574736964652071756F746564206C69746572616C732E2052657475726E732046616C736520666F72202222202F202247656E6572616C22202F2070757265206E756D6572696320666F726D6174732E0A
		Function IsDateFormatCode(code As String) As Boolean
		  ' Detects y/m/d/h/s tokens outside quoted "..." literals.
		  ' Returns True for codes like "dd/mm/yyyy", "hh:mm:ss", "yyyy-mm-dd hh:mm".
		  ' Returns False for plain numeric codes ("0.00", "#,##0", "0%").
		  If code = "" Or code = "General" Then Return False
		  Var inQuote As Boolean = False
		  For i As Integer = 0 To code.Length - 1
		    Var c As String = code.Middle(i, 1)
		    If c = """" Then
		      inQuote = Not inQuote
		      Continue
		    End If
		    If inQuote Then Continue
		    Select Case c.Lowercase
		    Case "y", "d", "h", "s", "m"
		      ' "m" is ambiguous (month vs minute) but in either case this is a date/time format.
		      Return True
		    End Select
		  Next
		  Return False
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 5475726E2061207261772063656C6C2076616C7565202B2063656C6C2074797065202B20457863656C20666F726D617420636F646520696E746F2074686520646973706C61792D726561647920746578742E20526F75746573207468726F75676820466F726D61744E756D62657256616C7565202F20466F726D61744461746556616C7565206261736564206F6E207479706520616E6420497344617465466F726D6174436F64652E0A
		Function Format(rawValue As String, cellType As XLSXEnums.eCellType, formatCode As String) As String
		  Select Case cellType
		  Case XLSXEnums.eCellType.Empty
		    Return ""
		  Case XLSXEnums.eCellType.Str
		    Return rawValue
		  Case XLSXEnums.eCellType.Bool
		    If rawValue = "1" Then Return "TRUE"
		    Return "FALSE"
		  Case XLSXEnums.eCellType.ErrorVal
		    Return rawValue
		  Case XLSXEnums.eCellType.Number, XLSXEnums.eCellType.FormulaCached, XLSXEnums.eCellType.DateValue
		    If rawValue = "" Then Return ""
		    Var d As Double = rawValue.ToDouble
		    If IsDateFormatCode(formatCode) Or cellType = XLSXEnums.eCellType.DateValue Then
		      Var dt As DateTime = ExcelSerialToDateTime(d)
		      Return FormatDateValue(dt, formatCode)
		    Else
		      Return FormatNumberValue(d, formatCode)
		    End If
		  End Select
		  Return rawValue
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 4170706C79206120707261676D6174696320737562736574206F6620457863656C206E756D6572696320666F726D617420636F6465732E20556E6B6E6F776E20636F6465732066616C6C206261636B20746F20446F75626C652E546F537472696E672E0A
		Private Function FormatNumberValue(d As Double, formatCode As String) As String
		  ' Pragmatic subset of Excel number format codes.
		  '
		  ' Routing order:
		  '   1. Text-only (@) and empty/General → plain ToString.
		  '   2. Recognized "simple" codes (0, 0.00, #,##0, …, %).
		  '   3. Scientific (contains E+0 / E-0).
		  '   4. Accounting (contains _( and _)).
		  '   5. Currency tag [$X-Y] → extract symbol, format number, prefix it.
		  '   6. Fall back to Double.ToString.
		  '
		  ' Format codes in Excel can have up to 4 sections separated by ";"
		  ' (positive ; negative ; zero ; text). For pragmatism we look at the
		  ' first section for matching and handle the sign here in code.

		  If formatCode = "" Or formatCode = "General" Then Return d.ToString
		  If formatCode = "@" Then Return d.ToString

		  Var first As String = formatCode
		  Var semi As Integer = first.IndexOf(";")
		  If semi >= 0 Then first = first.Left(semi)
		  first = StripColorHints(first)

		  Select Case first
		  Case "0"
		    Return d.ToString("0")
		  Case "0.00"
		    Return d.ToString("0.00")
		  Case "#,##0"
		    Return d.ToString("#,##0")
		  Case "#,##0.00"
		    Return d.ToString("#,##0.00")
		  Case "0%"
		    Var pct As Double = d * 100.0
		    Return pct.ToString("0") + "%"
		  Case "0.00%"
		    Var pct As Double = d * 100.0
		    Return pct.ToString("0.00") + "%"
		  End Select

		  If IsScientificFormat(first) Then Return FormatScientific(d, first)
		  If IsAccountingFormat(formatCode) Then Return FormatAccounting(d, formatCode)
		  If HasCurrencyTag(first) Then Return FormatWithCurrencyTag(d, first)

		  Return d.ToString
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 52656D6F766520457863656C20636F6C6F722068696E7473206C696B65205B5265645D2C205B426C75655D2C205B426C61636B5D2E204C6561766573205B24582D595D2063757272656E6379207461677320696E746163742E0A
		Private Function StripColorHints(c As String) As String
		  ' Strip Excel color hints like [Red], [Blue], [Magenta], [Black] but
		  ' leave [$X-Y] currency tags alone — they're handled separately.
		  Var result As String = c
		  Var colors() As String = Array("[Red]", "[Blue]", "[Black]", "[Magenta]", "[Cyan]", "[Yellow]", "[Green]", "[White]")
		  For Each tag As String In colors
		    result = result.ReplaceAll(tag, "")
		  Next
		  Return result
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 44657465637420736369656E7469666963206E6F746174696F6E20746F6B656E732028452B30202F20452D302920696E206120666F726D617420636F64652E0A
		Private Function IsScientificFormat(c As String) As Boolean
		  Var lc As String = c.Lowercase
		  Return lc.IndexOf("e+0") >= 0 Or lc.IndexOf("e-0") >= 0
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 52656E646572206120446F75626C65207573696E67206120736369656E746966696320666F726D617420636F64652076696120586F6A6F2773206275696C742D696E20466F726D617428292066756E6374696F6E2E0A
		Private Function FormatScientific(d As Double, c As String) As String
		  ' Xojo's Format() supports "0.00E+00"-style codes natively.
		  Return Format(d, c)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 44657465637420457863656C206163636F756E74696E6720666F726D6174732028636F6E7461696E205F2820616E64205F2920737061636572207061697273292E0A
		Private Function IsAccountingFormat(c As String) As Boolean
		  ' Excel's accounting format uses _( and _) spacer pairs and parenthesized
		  ' negatives. A reasonable signature.
		  Return c.IndexOf("_(") >= 0 And c.IndexOf("_)") >= 0
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 52656E646572206120446F75626C6520696E206163636F756E74696E67207374796C653A206F7074696F6E616C2063757272656E6379207072656669782C2074686F7573616E647320736570617261746F722C20706172656E74686573657320666F72206E65676174697665732C206461736820666F72207A65726F2E0A
		Private Function FormatAccounting(d As Double, c As String) As String
		  ' Pragmatic accounting rendering:
		  '   - "$" prefix if the code contains "$" or [$X-Y] currency tag (using X).
		  '   - 2 decimals if the code contains ".00", else 0.
		  '   - Negative values in parentheses.
		  '   - Zero values shown as the currency symbol + "-".
		  Var twoDec As Boolean = c.IndexOf(".00") >= 0
		  Var core As String = If(twoDec, "#,##0.00", "#,##0")
		  Var sym As String = ExtractCurrencySymbol(c)
		  If d = 0.0 Then
		    If sym <> "" Then Return sym + " -"
		    Return "-"
		  End If
		  Var absD As Double = d
		  If absD < 0 Then absD = -absD
		  Var body As String = absD.ToString(core)
		  If sym <> "" Then body = sym + " " + body
		  If d < 0 Then Return "(" + body + ")"
		  Return body
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 54727565207768656E2074686520666F726D617420636F646520636F6E7461696E7320616E20457863656C2063757272656E637920746167205B24582D595D2E0A
		Private Function HasCurrencyTag(c As String) As Boolean
		  Return c.IndexOf("[$") >= 0
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 52656E646572206120446F75626C652077697468207468652063757272656E63792073796D626F6C2066726F6D2061205B24582D595D2074616720617320707265666978206F72207375666669782E0A
		Private Function FormatWithCurrencyTag(d As Double, c As String) As String
		  Var sym As String = ExtractCurrencySymbol(c)
		  Var twoDec As Boolean = c.IndexOf(".00") >= 0
		  Var core As String = If(twoDec, "#,##0.00", "#,##0")
		  Var body As String = d.ToString(core)
		  If sym = "" Then Return body
		  ' Symbol position: heuristic — if "[$" appears before "#" or "0", prefix; else suffix.
		  Var tagPos As Integer = c.IndexOf("[$")
		  Var numPos As Integer = c.IndexOf("#")
		  If numPos < 0 Then numPos = c.IndexOf("0")
		  If tagPos < numPos Then Return sym + body
		  Return body + " " + sym
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 457874726163742061206C69746572616C2063757272656E63792073796D626F6C2066726F6D206120666F726D617420636F646520636F6E7461696E696E67205B24582D595D2C205B24585D2C202224222C206F722022E282AC222E2052657475726E7320656D707479206966206E6F2073796D626F6C2E0A
		Private Function ExtractCurrencySymbol(c As String) As String
		  ' Detect a currency symbol from one of:
		  '   [$SYMBOL-LOCALEID]   — Excel's modern form
		  '   [$SYMBOL]            — older form
		  '   "$"                  — literal $ in quotes (accounting)
		  ' Returns "" if no symbol is detected.
		  Var i As Integer = c.IndexOf("[$")
		  If i >= 0 Then
		    Var j As Integer = c.IndexOf(i, "]")
		    If j > i Then
		      Var inside As String = c.Middle(i + 2, j - i - 2)
		      Var dash As Integer = inside.IndexOf("-")
		      If dash >= 0 Then inside = inside.Left(dash)
		      Return inside
		    End If
		  End If
		  If c.IndexOf("""$""") >= 0 Then Return "$"
		  Var euro As String = Chr(&h20AC) ' €
		  If c.IndexOf("""" + euro + """") >= 0 Then Return euro
		  Return ""
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 4170706C79206120707261676D6174696320737562736574206F6620457863656C20646174652F74696D6520666F726D617420636F6465732E20556E6B6E6F776E20636F6465732066616C6C206261636B20746F204461746554696D652E53514C4461746554696D652E0A
		Private Function FormatDateValue(dt As DateTime, formatCode As String) As String
		  ' Pragmatic subset of Excel date/time format codes. Unknown codes fall back to SQLDateTime.
		  ' Excel formats may include a ";@" trailing section (the text-section); strip it.
		  Var fc As String = formatCode
		  Var semi As Integer = fc.IndexOf(";")
		  If semi >= 0 Then fc = fc.Left(semi)
		  Var lc As String = fc.Lowercase
		  Select Case lc
		  Case "yyyy-mm-dd"
		    Return dt.SQLDate
		  Case "dd/mm/yyyy"
		    Return Pad2(dt.Day) + "/" + Pad2(dt.Month) + "/" + dt.Year.ToString
		  Case "mm/dd/yyyy"
		    Return Pad2(dt.Month) + "/" + Pad2(dt.Day) + "/" + dt.Year.ToString
		  Case "hh:mm"
		    Return Pad2(dt.Hour) + ":" + Pad2(dt.Minute)
		  Case "hh:mm:ss"
		    Return Pad2(dt.Hour) + ":" + Pad2(dt.Minute) + ":" + Pad2(dt.Second)
		  Case "yyyy-mm-dd hh:mm"
		    Return dt.SQLDate + " " + Pad2(dt.Hour) + ":" + Pad2(dt.Minute)
		  Case "m/d/yy h:mm", "m/d/yyyy h:mm"
		    Var yr As String = dt.Year.ToString
		    If lc = "m/d/yy h:mm" Then yr = yr.Right(2)
		    Return dt.Month.ToString + "/" + dt.Day.ToString + "/" + yr + " " _
		      + dt.Hour.ToString + ":" + Pad2(dt.Minute)
		  Else
		    Return dt.SQLDateTime
		  End Select
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 54776F2D6469676974207A65726F2D70616464656420646563696D616C20737472696E6720666F7220616E20696E746567657220302E2E39392E0A
		Private Function Pad2(n As Integer) As String
		  If n < 10 Then Return "0" + n.ToString
		  Return n.ToString
		End Function
	#tag EndMethod

	#tag Note, Name = About
		Pure-function helpers that turn raw cell values + Excel format codes into
		display-ready strings.

		Public surface:
		  Format(rawValue, cellType, formatCode)         -> String
		  ExcelSerialToDateTime(serial As Double)         -> DateTime
		  IsDateFormatCode(code As String)                -> Boolean

		Excel's epoch is 1899-12-30 (off by one day from the Lotus 1-2-3 leap-year
		bug — that is the value Excel actually uses, so we match it).

		Format-code subset supported in V1:
		  numbers : "" / "General" / "0" / "0.00" / "#,##0" / "#,##0.00"
		            / "0%" / "0.00%"
		  dates   : "yyyy-mm-dd" / "dd/mm/yyyy" / "hh:mm" / "hh:mm:ss"
		            / "yyyy-mm-dd hh:mm"

		Anything else falls through to a sensible default (Double.ToString /
		DateTime.SQLDateTime). To extend support, add a Case branch in
		FormatNumberValue / FormatDateValue.
	#tag EndNote

	#tag ViewBehavior
	#tag EndViewBehavior
End Module
#tag EndModule
