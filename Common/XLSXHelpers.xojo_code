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

	#tag Method, Flags = &h0, Description = 457874656E73696F6E206D6574686F643A2072657475726E73207468652073796D626F6C6963206E616D65206F6620616E20654F70656E4D6F64652028224175746F222C20224D656D6F7279222C20224469736B22292E0A
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

	#tag Method, Flags = &h0, Description = 436F6E7665727420616E20457863656C20636F6C756D6E20776964746820696E2063686172616374657220756E69747320746F20706F696E74732E206368617257696474685078206F7665727269646573207468652064656661756C74203720706978656C2070657220636861726163746572206D657472696320746F206D6174636820616E6F7468657220666F6E742E0A
		Function ColumnCharsToPoints(chars As Double, charWidthPx As Double = 7.0) As Double
		  ' Excel column width is measured in the default font's "max digit width"
		  ' (~7px for Calibri 11) + 5px cell padding; 1px = 0.75pt at 96 DPI.
		  ' Canonical column-width unit in the model is points.
		  ' charWidthPx lets a caller substitute the metric of the font they actually
		  ' render in — a wider face needs a bigger number (see AutoFitColumn).
		  If chars <= 0 Then Return 0.0
		  Var w As Double = If(charWidthPx > 0, charWidthPx, 7.0)
		  Return ((chars * w) + 5.0) * 0.75
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 496E7665727365206F6620436F6C756D6E4368617273546F506F696E74732C20666F722077726974696E6720584C5358203C636F6C2077696474683D222E2E2E223E2E0A
		Function ColumnPointsToChars(points As Double) As Double
		  ' Inverse of ColumnCharsToPoints, for writing XLSX <col width="…">.
		  If points <= 0 Then Return 0.0
		  Var px As Double = points / 0.75
		  Var chars As Double = (px - 5.0) / 7.0
		  If chars < 0 Then Return 0.0
		  Return chars
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 506172736520616E204F4453206C656E6774682028636D2F6D6D2F696E2F70742F70632F70782920746F20706F696E74733B20756E6B6E6F776E2F6D697373696E6720756E6974207472656174656420617320706F696E74732E0A
		Function OdsLengthToPoints(s As String) As Double
		  ' Parse an ODS length ("2.5cm", "1.2in", "30pt", "15mm", "1pc", "64px")
		  ' into points. Unknown / missing unit is treated as points.
		  Var t As String = s.Trim.Lowercase
		  If t = "" Then Return 0.0
		  Var num As String = ""
		  Var unit As String = ""
		  For i As Integer = 0 To t.Length - 1
		    Var ch As String = t.Middle(i, 1)
		    Var c As Integer = Asc(ch)
		    If (c >= 48 And c <= 57) Or ch = "." Or ch = "-" Or ch = "+" Then
		      num = num + ch
		    Else
		      unit = t.Middle(i)
		      Exit
		    End If
		  Next
		  Var v As Double = num.ToDouble
		  Select Case unit
		  Case "cm"
		    Return v * 28.3464567
		  Case "mm"
		    Return v * 2.83464567
		  Case "in"
		    Return v * 72.0
		  Case "pc"
		    Return v * 12.0
		  Case "px"
		    Return v * 0.75
		  Else
		    Return v   ' "pt" or unspecified
		  End Select
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 466F726D617420706F696E747320617320616E204F4453206C656E67746820737472696E6720696E2063656E74696D657472657320284C696272654F666669636527732070726566657272656420756E6974292E0A
		Function PointsToOdsLength(points As Double) As String
		  ' Emit an ODS length in centimetres (LibreOffice's preferred unit),
		  ' rounded to 4 decimals. Empty string for non-positive input.
		  If points <= 0 Then Return ""
		  Var cm As Double = points / 28.3464567
		  Var r As Double = Round(cm * 10000.0) / 10000.0
		  Return Str(r) + "cm"
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

	#tag Method, Flags = &h0, Description = 436F6E7665727420616E20457863656C20523143312072656C617469766520666F726D756C6120746F2041312C20616E63686F7265642061742028637572526F772C20637572436F6C292E204F6E6C79207265666572656E636520746F6B656E73206172652072657772697474656E2E0A
		Function FormulaToA1(formula As String, curRow As Integer, curCol As Integer) As String
		  ' Convert a formula written in Excel R1C1 relative notation into A1 notation,
		  ' anchored at (curRow, curCol). Only R1C1 reference tokens are rewritten; the
		  ' rest of the formula (functions, operators, literals) is copied verbatim.
		  ' Examples at (1,6): SUM(R[+3]C:R[+1000]C) -> SUM(F4:F1001); RC[-2] -> H1.
		  If formula = "" Then Return ""
		  Var re As New RegEx
		  re.SearchPattern = "(?<![A-Za-z0-9_.])R(\[[+-]?\d+\]|\d+)?C(\[[+-]?\d+\]|\d+)?(?![A-Za-z0-9_(])"
		  Var opts As New RegExOptions
		  opts.CaseSensitive = True   ' R1C1 tokens are uppercase R / C
		  re.Options = opts
		  Var result As String
		  Var cursor As Integer = 0   ' 0-based char index into formula
		  Var match As RegExMatch = re.Search(formula)
		  While match <> Nil
		    ' Whole-match start byte offset -> 0-based char index (for String.Middle).
		    Var startIdx As Integer = formula.LeftBytes(match.SubExpressionStartB(0)).Length
		    Var whole As String = match.SubExpressionString(0)
		    result = result + formula.Middle(cursor, startIdx - cursor)
		    result = result + R1C1RefToA1(SubExprOrEmpty(match, 1), SubExprOrEmpty(match, 2), curRow, curCol)
		    cursor = startIdx + whole.Length
		    match = re.Search   ' continue from the last match position
		  Wend
		  result = result + formula.Middle(cursor)
		  Return result
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Function R1C1RefToA1(rPart As String, cPart As String, curRow As Integer, curCol As Integer) As String
		  ' Build one A1 reference from the R (row) and C (column) capture groups.
		  ' An empty group is the current row/col; [n] is a relative offset; a bare
		  ' number is an absolute index (emitted with a $ prefix).
		  Var rowAbs As Boolean
		  Var colAbs As Boolean
		  Var row As Integer = ResolveR1C1Index(rPart, curRow, rowAbs)
		  Var col As Integer = ResolveR1C1Index(cPart, curCol, colAbs)
		  If row < 1 Then row = 1
		  If col < 1 Then col = 1
		  Var s As String
		  If colAbs Then s = s + "$"
		  s = s + XLSXCellRef.IndexToColLettersRaw(col)
		  If rowAbs Then s = s + "$"
		  s = s + Str(row)
		  Return s
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Function ResolveR1C1Index(part As String, current As Integer, ByRef isAbsolute As Boolean) As Integer
		  ' Resolve one R1C1 index component against the current row/col.
		  isAbsolute = False
		  If part = "" Then Return current              ' bare R or C -> current
		  If part.Left(1) = "[" Then                    ' [+n] / [-n] -> relative
		    Var inner As String = part.Middle(1, part.Length - 2)
		    If inner.Left(1) = "+" Then inner = inner.Middle(1)
		    Return current + inner.ToInteger
		  End If
		  isAbsolute = True                              ' digits only -> absolute
		  Return part.ToInteger
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 436F6E7665727420616E20413120666F726D756C6120696E746F204F4446204F70656E466F726D756C6120666F726D20666F722061202E6F6473207461626C653A666F726D756C612028726566732077726170706564206173205B2E41315D2C2070726566697865642077697468206F663A3D292E0A
		Function A1ToOdfFormula(a1 As String) As String
		  ' Convert an A1-notation formula into ODF (OpenFormula) form for a .ods
		  ' table:formula attribute: cell/range refs wrapped as [.A1] / [.A1:.B2]
		  ' (sheet-qualified Sheet!A1 -> [Sheet.A1]), the whole prefixed with of:=.
		  ' Inverse of OdfFormulaToA1. String literals and operators are copied as-is.
		  If a1 = "" Then Return ""
		  Var re As New RegEx
		  re.SearchPattern = "(?<![A-Za-z0-9_.$!])(?:('[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?[A-Za-z]{1,3}\$?[0-9]+)(?::(\$?[A-Za-z]{1,3}\$?[0-9]+))?(?![A-Za-z0-9_(])"
		  Var opts As New RegExOptions
		  opts.CaseSensitive = True
		  re.Options = opts
		  Var result As String
		  Var cursor As Integer = 0
		  Var match As RegExMatch = re.Search(a1)
		  While match <> Nil
		    Var startIdx As Integer = a1.LeftBytes(match.SubExpressionStartB(0)).Length
		    Var whole As String = match.SubExpressionString(0)
		    result = result + a1.Middle(cursor, startIdx - cursor)
		    Var sheet As String = SubExprOrEmpty(match, 1)
		    Var c1 As String = SubExprOrEmpty(match, 2)
		    Var c2 As String = SubExprOrEmpty(match, 3)
		    Var head As String = If(sheet <> "", "[" + sheet + ".", "[.")
		    If c2 <> "" Then
		      result = result + head + c1 + ":." + c2 + "]"
		    Else
		      result = result + head + c1 + "]"
		    End If
		    cursor = startIdx + whole.Length
		    match = re.Search
		  Wend
		  result = result + a1.Middle(cursor)
		  Return "of:=" + result
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 436F6E766572742061202E6F6473207461626C653A666F726D756C6120284F444629206261636B20746F2041313A20737472697020746865206F663A3D202F206F6F6F633A3D202F203D2070726566697820616E6420756E7772617020746865205B2E41315D207265666572656E6365732E0A
		Function OdfFormulaToA1(odf As String) As String
		  ' Convert a .ods table:formula (ODF/OpenFormula) back into A1 notation:
		  ' strip the of:= / oooc:= / = prefix and unwrap [.A1] / [Sheet.A1] refs.
		  ' Inverse of A1ToOdfFormula.
		  Var s As String = odf
		  If s.Left(4) = "of:=" Then
		    s = s.Middle(4)
		  ElseIf s.Left(6) = "oooc:=" Then
		    s = s.Middle(6)
		  ElseIf s.Left(1) = "=" Then
		    s = s.Middle(1)
		  End If
		  Var re As New RegEx
		  re.SearchPattern = "\[([^\]]+)\]"
		  Var result As String
		  Var cursor As Integer = 0
		  Var match As RegExMatch = re.Search(s)
		  While match <> Nil
		    Var startIdx As Integer = s.LeftBytes(match.SubExpressionStartB(0)).Length
		    Var whole As String = match.SubExpressionString(0)
		    result = result + s.Middle(cursor, startIdx - cursor)
		    result = result + OdfRefInnerToA1(SubExprOrEmpty(match, 1))
		    cursor = startIdx + whole.Length
		    match = re.Search
		  Wend
		  result = result + s.Middle(cursor)
		  Return result
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Function OdfRefInnerToA1(inner As String) As String
		  ' Convert the inside of one ODF [ ... ] reference to A1. Same-sheet parts
		  ' start with '.' (dropped); a Sheet.cell part becomes Sheet!cell.
		  Var parts() As String = inner.Split(":")
		  Var outs() As String
		  For Each pp As String In parts
		    If pp.Left(1) = "." Then
		      outs.Add pp.Middle(1)
		    ElseIf pp.IndexOf(".") >= 0 Then
		      Var dot As Integer = pp.IndexOf(".")
		      outs.Add pp.Left(dot) + "!" + pp.Middle(dot + 1)
		    Else
		      outs.Add pp
		    End If
		  Next
		  Return String.FromArray(outs, ":")
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function SubExprOrEmpty(match As RegExMatch, index As Integer) As String
		  ' RegExMatch.SubExpressionString raises OutOfBoundsException for an optional
		  ' capture group that didn't participate (index >= SubExpressionCount). This
		  ' returns "" instead, matching the behaviour the converters assume.
		  If match.SubExpressionCount > index Then Return match.SubExpressionString(index)
		  Return ""
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 506F696E747320746F2073637265656E20706978656C73206174203936204450492E20436F6C756D6E20776964746873206172652073746F72656420696E20706F696E74733B206C697374626F7820436F6C756D6E5769647468732061726520706978656C732E0A
		Function PointsToPixels(points As Double) As Double
		  ' Points -> screen pixels at 96 DPI (1pt = 4/3 px). The model stores column
		  ' widths in points; listbox ColumnWidths are pixels.
		  Return points * 4.0 / 3.0
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 53637265656E20706978656C7320746F20706F696E7473206174203936204450492E20496E7665727365206F6620506F696E7473546F506978656C732E0A
		Function PixelsToPoints(pixels As Double) As Double
		  ' Screen pixels -> points at 96 DPI (1px = 0.75pt). Inverse of PointsToPixels.
		  Return pixels * 0.75
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 5472616E736C61746520612063616C6C65722D737570706C6965642073686565742F726F772F636F6C756D6E20696E64657820696E746F20746865206D6F64656C277320696E7465726E616C20312D626173656420636F6F7264696E6174652E204964656E7469747920756E6C65737320675A65726F4261736564536865657473526F7773436F6C756D6E7320697320547275652E0A
		Function ToInternalIndex(publicIndex As Integer) As Integer
		  ' Translate a caller-supplied sheet/row/column index into the model's
		  ' internal 1-based coordinate. A no-op unless gZeroBasedSheetsRowsColumns
		  ' is True, so flag-off behaviour is byte-identical to before the flag existed.
		  If gZeroBasedSheetsRowsColumns Then Return publicIndex + 1
		  Return publicIndex
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 496E7665727365206F6620546F496E7465726E616C496E6465783A20616E20696E7465726E616C20312D626173656420636F6F7264696E6174652065787072657373656420696E207468652063616C6C6572277320696E64657820626173652E0A
		Function ToPublicIndex(internalIndex As Integer) As Integer
		  ' Inverse of ToInternalIndex: an internal 1-based coordinate as the caller
		  ' expects to see it. Also a no-op while the flag is False.
		  If gZeroBasedSheetsRowsColumns Then Return internalIndex - 1
		  Return internalIndex
		End Function
	#tag EndMethod

	#tag Property, Flags = &h0, Description = 476C6F62616C207377697463683A207768656E205472756520746865207075626C69632041504920616464726573736573207368656574732C20726F777320616E6420636F6C756D6E732066726F6D203020696E7374656164206F6620457863656C277320312E2046616C73652062792064656661756C742C20616E642061206E6F2D6F70207768656E2046616C73652E0A
		gZeroBasedSheetsRowsColumns As Boolean = False
	#tag EndProperty

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
