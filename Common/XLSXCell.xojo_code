#tag Class
Protected Class XLSXCell
	#tag Method, Flags = &h0, Description = 4275696C6420612063656C6C20776974682069747320747970652C207261772076616C75652028766572626174696D2066726F6D203C763E206F72203C69733E3C743E292C20616E64206F7074696F6E616C207374796C6520696E6465782E0A
		Sub Constructor(cellType As XLSXEnums.eCellType, rawValue As String, styleIndex As Integer = -1)
		  Me.eType = cellType
		  Me.RawString = rawValue
		  Me.StyleIndex = styleIndex
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 52657475726E73205472756520696620655479706520697320456D707479206F7220526177537472696E6720697320656D7074792E0A
		Function IsEmpty() As Boolean
		  ' A formula cell is never "empty" even before it carries a cached value,
		  ' so writers and listbox fillers keep it instead of skipping it.
		  If eType = XLSXEnums.eCellType.FormulaCached And Formula <> "" Then Return False
		  Return eType = XLSXEnums.eCellType.Empty Or RawString = ""
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function HasFormula() As Boolean
		  ' True if this cell carries a formula (whose text we preserved).
		  Return eType = XLSXEnums.eCellType.FormulaCached And Formula <> ""
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function FormulaText() As String
		  ' The formula as an "=" expression for display (empty if not a formula).
		  If Not HasFormula Then Return ""
		  Return "=" + Formula
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Shared Function TextCell(text As String) As XLSXCell
		  ' Fluent authoring: a string cell.
		  Return New XLSXCell(XLSXEnums.eCellType.Str, text)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Shared Function NumberCell(value As Double) As XLSXCell
		  ' Fluent authoring: a plain number cell. Str() keeps the invariant
		  ' period-decimal form the writers put verbatim into <v>.
		  Return New XLSXCell(XLSXEnums.eCellType.Number, Str(value))
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Shared Function MoneyCell(value As Double) As XLSXCell
		  ' Fluent authoring: a number cell with a two-decimal thousands format.
		  Var c As New XLSXCell(XLSXEnums.eCellType.Number, Str(value))
		  c.FormatCode = kMoneyFormat
		  Return c
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Shared Function DateCell(dt As DateTime) As XLSXCell
		  ' Fluent authoring: a date cell (Excel serial + ISO date format).
		  Var c As New XLSXCell(XLSXEnums.eCellType.DateValue, Str(XLSXFormatter.DateTimeToExcelSerial(dt)))
		  c.FormatCode = kDateFormat
		  Return c
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Shared Function DateTimeCell(dt As DateTime) As XLSXCell
		  ' Fluent authoring: a date+time cell (Excel serial + ISO date-time format).
		  Var c As New XLSXCell(XLSXEnums.eCellType.DateValue, Str(XLSXFormatter.DateTimeToExcelSerial(dt)))
		  c.FormatCode = kDateTimeFormat
		  Return c
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Shared Function BoolCell(value As Boolean) As XLSXCell
		  ' Fluent authoring: a boolean cell.
		  Return New XLSXCell(XLSXEnums.eCellType.Bool, If(value, "1", "0"))
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Shared Function FormulaCell(formula As String) As XLSXCell
		  ' Fluent authoring: a formula in Excel R1C1 relative notation
		  ' (e.g. "SUM(R[+3]C:R[+1000]C)"). Converted to A1 at write time using
		  ' the cell's anchor position. A leading "=" is optional and stripped.
		  Return MakeFormula(formula, True)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Shared Function FormulaCellA1(formula As String) As XLSXCell
		  ' Fluent authoring: a formula in plain A1 notation (e.g. "SUM(F4:F1001)").
		  ' Written verbatim. A leading "=" is optional and stripped.
		  Return MakeFormula(formula, False)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Shared Function MakeFormula(formula As String, isR1C1 As Boolean) As XLSXCell
		  Var f As String = formula
		  If f.Left(1) = "=" Then f = f.Middle(1)
		  Var c As New XLSXCell(XLSXEnums.eCellType.FormulaCached, "")
		  c.Formula = f
		  c.FormulaIsR1C1 = isR1C1
		  Return c
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Function EnsureStyle() As XLSXCellStyle
		  ' Lazily create the per-cell CellStyle so the fluent mutators can set it.
		  If CellStyle Is Nil Then CellStyle = New XLSXCellStyle
		  Return CellStyle
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Bold(on As Boolean = True) As XLSXCell
		  ' Fluent style: make the cell bold. Returns Me for chaining.
		  EnsureStyle.Bold = on
		  Return Me
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Italic(on As Boolean = True) As XLSXCell
		  EnsureStyle.Italic = on
		  Return Me
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Underline(on As Boolean = True) As XLSXCell
		  EnsureStyle.Underline = on
		  Return Me
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Format(code As String) As XLSXCell
		  ' Fluent style: set the number format code (e.g. "#,##0.00", "0%").
		  FormatCode = code
		  Return Me
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Money() As XLSXCell
		  ' Fluent style: two-decimal thousands number format.
		  FormatCode = kMoneyFormat
		  Return Me
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Align(h As XLSXEnums.eAlignH) As XLSXCell
		  ' Fluent style: horizontal alignment.
		  EnsureStyle.AlignH = h
		  Return Me
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function BackColor(c As Color) As XLSXCell
		  ' Fluent style: solid background fill.
		  Var st As XLSXCellStyle = EnsureStyle
		  st.BackgroundColor = c
		  st.HasBackground = True
		  Return Me
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function FontColor(c As Color) As XLSXCell
		  ' Fluent style: font colour.
		  Var st As XLSXCellStyle = EnsureStyle
		  st.FontColor = c
		  st.HasFontColor = True
		  Return Me
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function FontFace(name As String) As XLSXCell
		  ' Fluent style: font name. (FontFace avoids clashing with XLSXCellStyle.FontName.)
		  EnsureStyle.FontName = name
		  Return Me
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function FontSize(points As Double) As XLSXCell
		  ' Fluent style: font size in points.
		  EnsureStyle.FontSize = points
		  Return Me
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Border(style As XLSXEnums.eBorderStyle, borderColor As Color = &c000000) As XLSXCell
		  ' Fluent style: same border on all four edges.
		  Var st As XLSXCellStyle = EnsureStyle
		  st.BorderLeft = style
		  st.BorderRight = style
		  st.BorderTop = style
		  st.BorderBottom = style
		  st.BorderColor = borderColor
		  st.HasBorderColor = True
		  Return Me
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 506172736520526177537472696E67206173206120446F75626C652E2052657475726E7320302E3020696620526177537472696E6720697320656D7074792E0A
		Function NumberValue() As Double
		  If RawString = "" Then Return 0.0
		  Return RawString.ToDouble
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 52657475726E73207468652063656C6C206173204461746554696D65206966206554797065206973204461746556616C75652C206F7468657277697365204E696C2E0A
		Function DateValue() As DateTime
		  If eType <> XLSXEnums.eCellType.DateValue Then Return Nil
		  Return XLSXFormatter.ExcelSerialToDateTime(NumberValue)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 52657475726E73205472756520696620526177537472696E6720697320223122206F722028636173652D696E73656E73697469766529202274727565222E0A
		Function BooleanValue() As Boolean
		  If RawString = "1" Then Return True
		  If RawString.Lowercase = "true" Then Return True
		  Return False
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 5468652063656C6C27732076697375616C207374796C653A2061206469726563746C792D7365742043656C6C5374796C652077696E7320284F44532F656469746564292C20656C7365207265736F6C7665205374796C65496E6465782076696120584C53585374796C65732E204E65766572204E696C2E0A
		Function ResolvedStyle(styles As XLSXStyles) As XLSXCellStyle
		  ' The cell's visual style: a directly-set CellStyle wins (ODS-loaded /
		  ' edited cells), otherwise resolve StyleIndex via the workbook's XLSXStyles.
		  ' Never Nil — falls back to a shared default style.
		  If CellStyle <> Nil Then Return CellStyle
		  If styles <> Nil And StyleIndex >= 0 Then Return styles.CellStyleAt(StyleIndex)
		  If mDefaultStyle Is Nil Then mDefaultStyle = New XLSXCellStyle
		  Return mDefaultStyle
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 52657475726E732074686520757365722D726561647920646973706C617920737472696E6720666F7220746869732063656C6C2C206170706C79696E6720746865207374796C65277320666F726D617420636F64652076696120584C5358466F726D61747465722E204361636865642061667465722066697273742063616C6C2E0A
		Function DisplayText(styles As XLSXStyles) As String
		  If mDisplayCached Then Return mDisplayCache
		  ' A directly-set FormatCode (e.g. resolved by ODSReader) wins; otherwise
		  ' look the format code up via the workbook's XLSXStyles by StyleIndex.
		  Var fmt As String = FormatCode
		  If fmt = "" And styles <> Nil And StyleIndex >= 0 Then
		    fmt = styles.NumberFormatCodeAt(StyleIndex)
		  End If
		  ' If a numeric cell carries a date format, treat it as a date for formatting.
		  Var typeForFormat As XLSXEnums.eCellType = eType
		  If typeForFormat = XLSXEnums.eCellType.Number Or typeForFormat = XLSXEnums.eCellType.FormulaCached Then
		    If XLSXFormatter.IsDateFormatCode(fmt) Then
		      typeForFormat = XLSXEnums.eCellType.DateValue
		    End If
		  End If
		  mDisplayCache = XLSXFormatter.Format(RawString, typeForFormat, fmt)
		  mDisplayCached = True
		  Return mDisplayCache
		End Function
	#tag EndMethod

	#tag Property, Flags = &h0, Description = 5468652063656C6C27732076616C7565207479706520E280942064726976657320446973706C61795465787420726F7574696E672E0A
		eType As XLSXEnums.eCellType
	#tag EndProperty

	#tag Property, Flags = &h0, Description = 566572626174696D2063656C6C2076616C7565206173206974206170706561727320696E203C763E20286E756D6265727320617320646563696D616C20737472696E67732C2073686172656420737472696E677320616C7265616479207265736F6C76656420627920584C5358536865657429206F72203C69733E3C743E2E0A
		RawString As String
	#tag EndProperty

	#tag Property, Flags = &h0, Description = 5468652063656C6C277320407320617474726962757465202863656C6C586620696E646578292C206F72202D3120696620616273656E742E0A
		StyleIndex As Integer = -1
	#tag EndProperty

	#tag Property, Flags = &h0, Description = 41206469726563746C792D73657420457863656C20666F726D617420636F64652028652E672E207265736F6C766564206279204F4453526561646572292E205768656E206E6F6E2D656D7074792C20446973706C617954657874207573657320697420696E7374656164206F66206C6F6F6B696E672075702076696120584C53585374796C65732E0A
		FormatCode As String
	#tag EndProperty

	#tag Property, Flags = &h0, Description = 41206469726563746C792D7365742076697375616C207374796C652028652E672E20666F72204F44532D6C6F61646564206F72206564697465642063656C6C73293B207768656E2070726573656E742069742077696E73206F76657220746865205374796C65496E646578206C6F6F6B75702E0A
		CellStyle As XLSXCellStyle
	#tag EndProperty

	#tag Property, Flags = &h0
		Formula As String
	#tag EndProperty

	#tag Property, Flags = &h0
		FormulaIsR1C1 As Boolean = False
	#tag EndProperty


	#tag Property, Flags = &h21
		Private mDefaultStyle As XLSXCellStyle
	#tag EndProperty

	#tag Property, Flags = &h21, Description = 43616368656420446973706C61795465787420726573756C742E0A
		Private mDisplayCache As String
	#tag EndProperty

	#tag Property, Flags = &h21, Description = 54727565206F6E636520446973706C61795465787420686173206265656E20636F6D70757465642E0A
		Private mDisplayCached As Boolean = False
	#tag EndProperty

	#tag Constant, Name = kMoneyFormat, Type = String, Dynamic = False, Default = \"#,##0.00", Scope = Private
	#tag EndConstant

	#tag Constant, Name = kDateFormat, Type = String, Dynamic = False, Default = \"yyyy-mm-dd", Scope = Private
	#tag EndConstant

	#tag Constant, Name = kDateTimeFormat, Type = String, Dynamic = False, Default = \"yyyy-mm-dd hh:mm", Scope = Private
	#tag EndConstant

	#tag Note, Name = About
		One parsed cell from a worksheet.

		Constructor takes the cell type, the raw value (verbatim from <v> /
		<is><t>), and a style index (the cell's @s attribute, -1 if absent).

		Display: call DisplayText(styles) to get the user-ready text. This
		method:
		  - returns the raw string for Str / ErrorVal cells,
		  - converts Bool to TRUE/FALSE,
		  - applies XLSXFormatter.Format() to numbers, dates, formulas,
		  - upgrades a numeric cell to DateValue when its style's format code
		    contains date tokens (so 44621 with "dd/mm/yyyy" displays the
		    date, not the serial).
		The result is cached the first time it's computed.

		Type-typed accessors:
		  NumberValue   - parses RawString as Double
		  DateValue     - returns Nil unless eType is DateValue
		  BooleanValue  - True iff RawString is "1" or "true"
	#tag EndNote

	#tag ViewBehavior
	#tag EndViewBehavior
End Class
#tag EndClass
