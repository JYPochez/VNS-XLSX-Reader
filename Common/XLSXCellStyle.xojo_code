#tag Class
Protected Class XLSXCellStyle
	#tag Method, Flags = &h0, Description = 54727565207768656E206E6F2076697375616C207374796C696E672069732073657420E28094206C657473207772697465727320736B6970207468652063656C6C20616E64207468652055492061766F696420637573746F6D207061696E74696E672E0A
		Function IsDefault() As Boolean
		  ' True when nothing visual is set — lets writers skip the cell and the UI
		  ' avoid custom painting.
		  If Bold Or Italic Or Underline Then Return False
		  If FontName <> "" Then Return False
		  If FontSize > 0 Then Return False
		  If HasFontColor Or HasBackground Then Return False
		  If AlignH <> XLSXEnums.eAlignH.General Then Return False
		  If AlignV <> XLSXEnums.eAlignV.Bottom Then Return False
		  If WrapText Then Return False
		  If BorderLeft <> XLSXEnums.eBorderStyle.None Then Return False
		  If BorderRight <> XLSXEnums.eBorderStyle.None Then Return False
		  If BorderTop <> XLSXEnums.eBorderStyle.None Then Return False
		  If BorderBottom <> XLSXEnums.eBorderStyle.None Then Return False
		  Return True
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Clone() As XLSXCellStyle
		  ' A shallow copy — used when a caller needs to tweak one attribute (e.g.
		  ' display alignment) without mutating a style shared across many cells.
		  Var c As New XLSXCellStyle
		  c.Bold = Bold
		  c.Italic = Italic
		  c.Underline = Underline
		  c.FontName = FontName
		  c.FontSize = FontSize
		  c.FontColor = FontColor
		  c.HasFontColor = HasFontColor
		  c.BackgroundColor = BackgroundColor
		  c.HasBackground = HasBackground
		  c.AlignH = AlignH
		  c.AlignV = AlignV
		  c.WrapText = WrapText
		  c.BorderLeft = BorderLeft
		  c.BorderRight = BorderRight
		  c.BorderTop = BorderTop
		  c.BorderBottom = BorderBottom
		  c.BorderColor = BorderColor
		  c.HasBorderColor = HasBorderColor
		  Return c
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 5472756520696620616E79206564676520686173206120626F72646572207374796C65206F74686572207468616E204E6F6E652E0A
		Function HasAnyBorder() As Boolean
		  Return BorderLeft <> XLSXEnums.eBorderStyle.None Or BorderRight <> XLSXEnums.eBorderStyle.None _
		    Or BorderTop <> XLSXEnums.eBorderStyle.None Or BorderBottom <> XLSXEnums.eBorderStyle.None
		End Function
	#tag EndMethod

	#tag Property, Flags = &h0
		Bold As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h0
		Italic As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h0
		Underline As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h0
		FontName As String
	#tag EndProperty

	#tag Property, Flags = &h0
		FontSize As Double = 0.0
	#tag EndProperty

	#tag Property, Flags = &h0
		FontColor As Color
	#tag EndProperty

	#tag Property, Flags = &h0
		HasFontColor As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h0
		BackgroundColor As Color
	#tag EndProperty

	#tag Property, Flags = &h0
		HasBackground As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h0
		AlignH As XLSXEnums.eAlignH = XLSXEnums.eAlignH.General
	#tag EndProperty

	#tag Property, Flags = &h0
		AlignV As XLSXEnums.eAlignV = XLSXEnums.eAlignV.Bottom
	#tag EndProperty

	#tag Property, Flags = &h0
		WrapText As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h0
		BorderLeft As XLSXEnums.eBorderStyle = XLSXEnums.eBorderStyle.None
	#tag EndProperty

	#tag Property, Flags = &h0
		BorderRight As XLSXEnums.eBorderStyle = XLSXEnums.eBorderStyle.None
	#tag EndProperty

	#tag Property, Flags = &h0
		BorderTop As XLSXEnums.eBorderStyle = XLSXEnums.eBorderStyle.None
	#tag EndProperty

	#tag Property, Flags = &h0
		BorderBottom As XLSXEnums.eBorderStyle = XLSXEnums.eBorderStyle.None
	#tag EndProperty

	#tag Property, Flags = &h0
		BorderColor As Color
	#tag EndProperty

	#tag Property, Flags = &h0
		HasBorderColor As Boolean = False
	#tag EndProperty

	#tag Note, Name = About
		Visual formatting for one cell (or one style slot): font, fill, alignment,
		and borders. Shared between the XLSX and ODS readers/writers and the UI.

		It carries only what the model needs to *preserve and render* — not the
		full Excel/ODF style model. Defaults match "no styling": General/Bottom
		alignment, no colors, no borders. IsDefault() is True for an unstyled cell
		so writers can skip it.

		Colors use the Xojo Color type. HasFontColor / HasBackground / HasBorderColor
		distinguish "explicitly set" from "left at the &c000000 default", since
		Color has no nil state.

		Borders are simplified to a 4-value style per edge (None/Thin/Medium/Thick)
		plus a single border color — enough to round-trip the common cases without
		modelling every Excel/ODF dash variant.

		Out of scope: rich text (per-run formatting within a cell), gradient fills,
		diagonal borders, number format (that lives on XLSXCell.FormatCode / styles).
	#tag EndNote

	#tag ViewBehavior
	#tag EndViewBehavior
End Class
#tag EndClass
