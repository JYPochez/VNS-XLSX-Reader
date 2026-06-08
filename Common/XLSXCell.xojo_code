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
		  Return eType = XLSXEnums.eCellType.Empty Or RawString = ""
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

	#tag Property, Flags = &h21, Description = 43616368656420446973706C61795465787420726573756C742E0A
		Private mDisplayCache As String
	#tag EndProperty

	#tag Property, Flags = &h21, Description = 54727565206F6E636520446973706C61795465787420686173206265656E20636F6D70757465642E0A
		Private mDisplayCached As Boolean = False
	#tag EndProperty

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
