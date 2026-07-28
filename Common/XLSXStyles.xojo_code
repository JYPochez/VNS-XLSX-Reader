#tag Class
Protected Class XLSXStyles
	#tag Method, Flags = &h0, Description = 506172736520746865207374796C65732E786D6C20746578742E20456D70747920696E70757420697320616C6C6F77656420E28094207765207374696C6C2073656564206275696C742D696E206E756D466D74206964732E0A
		Sub Constructor(stylesXml As String, themeXml As String = "")
		  ' stylesXml may be empty when the workbook has no styles.xml part.
		  ' themeXml is xl/theme/theme1.xml (optional) — its colour scheme lets us
		  ' resolve <color theme="N" tint="…"/> references used by fonts/fills/borders.
		  mNumFmts = New Dictionary
		  mCellXfs = New Dictionary
		  mCellStyles = New Dictionary
		  SeedBuiltInNumFmts
		  ParseTheme(themeXml)   ' must run before colours are resolved below
		  If stylesXml = "" Then Return
		  Var doc As New XmlDocument
		  Try
		    doc.LoadXml(stylesXml)
		  Catch
		    Raise New XLSXException(XLSXEnums.eParseError.MalformedXML, "styles.xml")
		  End Try
		  ParseNumFmts(doc)
		  ParseFonts(doc)
		  ParseFills(doc)
		  ParseBorders(doc)
		  ParseCellXfs(doc)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 5265736F6C7665642076697375616C207374796C652028666F6E742F66696C6C2F616C69676E6D656E742F626F72646572732920666F7220612063656C6C586620696E6465782E204E65766572204E696C20E2809420616273656E7420696E64696365732072657475726E2061207368617265642064656661756C742E0A
		Function CellStyleAt(cellXfIndex As Integer) As XLSXCellStyle
		  ' Resolved visual style (font / fill / alignment / borders) for a cellXf
		  ' index. Never Nil — absent indices return a shared default style.
		  If mCellStyles.HasKey(cellXfIndex) Then Return mCellStyles.Value(cellXfIndex)
		  If mEmptyStyle Is Nil Then mEmptyStyle = New XLSXCellStyle
		  Return mEmptyStyle
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 52657475726E732074686520457863656C20666F726D617420636F646520666F722074686520676976656E2063656C6C586620696E6465782C206F7220656D70747920737472696E6720696620616273656E742E0A
		Function NumberFormatCodeAt(cellXfIndex As Integer) As String
		  ' Returns the format code (e.g. "0.00", "dd/mm/yyyy", "General") for the given
		  ' cellXf index, or "" if the index is out of range.
		  If Not mCellXfs.HasKey(cellXfIndex) Then Return ""
		  Var numFmtId As Integer = mCellXfs.Value(cellXfIndex)
		  If mNumFmts.HasKey(numFmtId) Then Return mNumFmts.Value(numFmtId)
		  Return ""
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 546F74616C206E756D626572206F66203C78663E20656E7472696573207061727365642066726F6D203C63656C6C5866733E2E0A
		Function CellXfCount() As Integer
		  Return mCellXfs.KeyCount
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 506F70756C61746520746865206275696C742D696E20457863656C206E756D466D74206964732028302E2E343920737562736574292E20437573746F6D20636F6465732066726F6D203C6E756D466D74733E206F766572726964652074686573652E0A
		Private Sub SeedBuiltInNumFmts()
		  ' Excel's built-in numFmtIds 0..49 (subset that we recognize).
		  ' Custom format codes (id >= 164) declared in <numFmts> override these.
		  mNumFmts.Value(0) = "General"
		  mNumFmts.Value(1) = "0"
		  mNumFmts.Value(2) = "0.00"
		  mNumFmts.Value(3) = "#,##0"
		  mNumFmts.Value(4) = "#,##0.00"
		  mNumFmts.Value(9) = "0%"
		  mNumFmts.Value(10) = "0.00%"
		  mNumFmts.Value(14) = "dd/mm/yyyy"
		  mNumFmts.Value(15) = "d-mmm-yy"
		  mNumFmts.Value(16) = "d-mmm"
		  mNumFmts.Value(17) = "mmm-yy"
		  mNumFmts.Value(18) = "h:mm AM/PM"
		  mNumFmts.Value(19) = "h:mm:ss AM/PM"
		  mNumFmts.Value(20) = "hh:mm"
		  mNumFmts.Value(21) = "hh:mm:ss"
		  mNumFmts.Value(22) = "yyyy-mm-dd hh:mm"
		  ' Built-in accounting / negative-in-parens forms (Excel ids 37..44).
		  mNumFmts.Value(37) = "#,##0 ;(#,##0)"
		  mNumFmts.Value(38) = "#,##0 ;[Red](#,##0)"
		  mNumFmts.Value(39) = "#,##0.00;(#,##0.00)"
		  mNumFmts.Value(40) = "#,##0.00;[Red](#,##0.00)"
		  mNumFmts.Value(41) = "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)"
		  mNumFmts.Value(42) = "_(""$""* #,##0_);_(""$""* (#,##0);_(""$""* ""-""_);_(@_)"
		  mNumFmts.Value(43) = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
		  mNumFmts.Value(44) = "_(""$""* #,##0.00_);_(""$""* (#,##0.00);_(""$""* ""-""??_);_(@_)"
		  mNumFmts.Value(48) = "##0.0E+0"
		  mNumFmts.Value(49) = "@"
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 5061727365203C6E756D466D74733E2F3C6E756D466D743E20656E74726965732028637573746F6D20666F726D617420636F6465732C20696473203E3D20313634292E0A
		Private Sub ParseNumFmts(doc As XmlDocument)
		  Var nodes As XmlNodeList = doc.Xql("//*[local-name()='numFmts']/*[local-name()='numFmt']")
		  For i As Integer = 0 To nodes.Length - 1
		    Var n As XmlNode = nodes.Item(i)
		    Var idAttr As String = n.GetAttribute("numFmtId")
		    Var codeAttr As String = n.GetAttribute("formatCode")
		    If idAttr <> "" Then
		      mNumFmts.Value(Integer.FromString(idAttr)) = codeAttr
		    End If
		  Next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 5061727365203C63656C6C5866733E2F3C78663E20656E74726965733B206275696C6420746865206D61702066726F6D207866496E64657820746F206E756D466D7449642E0A
		Private Sub ParseCellXfs(doc As XmlDocument)
		  Var nodes As XmlNodeList = doc.Xql("//*[local-name()='cellXfs']/*[local-name()='xf']")
		  For i As Integer = 0 To nodes.Length - 1
		    Var n As XmlNode = nodes.Item(i)
		    Var fmtAttr As String = n.GetAttribute("numFmtId")
		    Var numFmtId As Integer = If(fmtAttr <> "", Integer.FromString(fmtAttr), 0)
		    mCellXfs.Value(i) = numFmtId

		    ' Assemble the visual style: font + fill + border fragments by id,
		    ' plus the <alignment> child.
		    Var st As New XLSXCellStyle
		    Var fontId As Integer = IntAttr(n, "fontId", -1)
		    If fontId >= 0 And fontId <= mFonts.LastIndex Then CopyFont(mFonts(fontId), st)
		    Var fillId As Integer = IntAttr(n, "fillId", -1)
		    If fillId >= 0 And fillId <= mFills.LastIndex Then CopyFill(mFills(fillId), st)
		    Var borderId As Integer = IntAttr(n, "borderId", -1)
		    If borderId >= 0 And borderId <= mBorders.LastIndex Then CopyBorder(mBorders(borderId), st)
		    ApplyAlignment(n, st)
		    mCellStyles.Value(i) = st
		  Next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 5061727365203C666F6E74733E2F3C666F6E743E20696E746F207065722D666F6E7420584C535843656C6C5374796C6520667261676D656E74732028626F6C642F6974616C69632F756E6465726C696E652F6E616D652F73697A652F636F6C6F72292E0A
		Private Sub ParseFonts(doc As XmlDocument)
		  Var nodes As XmlNodeList = doc.Xql("//*[local-name()='fonts']/*[local-name()='font']")
		  For i As Integer = 0 To nodes.Length - 1
		    Var n As XmlNode = nodes.Item(i)
		    Var st As New XLSXCellStyle
		    st.Bold = FlagOn(n, "b")
		    st.Italic = FlagOn(n, "i")
		    ' Underline is on unless val="none" (val can be single/double/…).
		    Var uList As XmlNodeList = n.Xql("./*[local-name()='u']")
		    If uList.Length > 0 Then st.Underline = (uList.Item(0).GetAttribute("val") <> "none")
		    Var szList As XmlNodeList = n.Xql("./*[local-name()='sz']")
		    If szList.Length > 0 Then st.FontSize = szList.Item(0).GetAttribute("val").ToDouble
		    Var nameList As XmlNodeList = n.Xql("./*[local-name()='name' or local-name()='rFont']")
		    If nameList.Length > 0 Then st.FontName = nameList.Item(0).GetAttribute("val")
		    Var colorList As XmlNodeList = n.Xql("./*[local-name()='color']")
		    If colorList.Length > 0 Then
		      Var c As Color
		      If ResolveColor(colorList.Item(0), c) Then
		        st.FontColor = c
		        st.HasFontColor = True
		      End If
		    End If
		    mFonts.Add st
		  Next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 5061727365203C66696C6C733E2F3C66696C6C3E3B206120736F6C6964207061747465726E46696C6C2773206667436F6C6F72206265636F6D657320746865206261636B67726F756E6420636F6C6F722E0A
		Private Sub ParseFills(doc As XmlDocument)
		  Var nodes As XmlNodeList = doc.Xql("//*[local-name()='fills']/*[local-name()='fill']")
		  For i As Integer = 0 To nodes.Length - 1
		    Var st As New XLSXCellStyle
		    Var pf As XmlNodeList = nodes.Item(i).Xql("./*[local-name()='patternFill']")
		    If pf.Length > 0 Then
		      ' A solid fill uses fgColor as the visible fill color.
		      If pf.Item(0).GetAttribute("patternType") = "solid" Then
		        Var fg As XmlNodeList = pf.Item(0).Xql("./*[local-name()='fgColor']")
		        If fg.Length > 0 Then
		          Var c As Color
		          If ResolveColor(fg.Item(0), c) Then
		            st.BackgroundColor = c
		            st.HasBackground = True
		          End If
		        End If
		      End If
		    End If
		    mFills.Add st
		  Next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 5061727365203C626F72646572733E2F3C626F726465723E20696E746F207065722D6564676520626F72646572207374796C6573202B20612073696E676C6520626F7264657220636F6C6F722E0A
		Private Sub ParseBorders(doc As XmlDocument)
		  Var nodes As XmlNodeList = doc.Xql("//*[local-name()='borders']/*[local-name()='border']")
		  For i As Integer = 0 To nodes.Length - 1
		    Var n As XmlNode = nodes.Item(i)
		    Var st As New XLSXCellStyle
		    st.BorderLeft = EdgeStyle(n, "left", st)
		    st.BorderRight = EdgeStyle(n, "right", st)
		    st.BorderTop = EdgeStyle(n, "top", st)
		    st.BorderBottom = EdgeStyle(n, "bottom", st)
		    mBorders.Add st
		  Next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function EdgeStyle(borderNode As XmlNode, edge As String, st As XLSXCellStyle) As XLSXEnums.eBorderStyle
		  ' Read one <left|right|top|bottom style="…"><color rgb="…"/></…> edge.
		  ' Captures the first edge color into st.BorderColor (single-color model).
		  Var edgeList As XmlNodeList = borderNode.Xql("./*[local-name()='" + edge + "']")
		  If edgeList.Length = 0 Then Return XLSXEnums.eBorderStyle.None
		  Var styleAttr As String = edgeList.Item(0).GetAttribute("style")
		  Var bs As XLSXEnums.eBorderStyle = BorderStyleFromString(styleAttr)
		  If bs <> XLSXEnums.eBorderStyle.None And Not st.HasBorderColor Then
		    Var cl As XmlNodeList = edgeList.Item(0).Xql("./*[local-name()='color']")
		    If cl.Length > 0 Then
		      Var c As Color
		      If ResolveColor(cl.Item(0), c) Then
		        st.BorderColor = c
		        st.HasBorderColor = True
		      End If
		    End If
		  End If
		  Return bs
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub ApplyAlignment(xfNode As XmlNode, st As XLSXCellStyle)
		  Var al As XmlNodeList = xfNode.Xql("./*[local-name()='alignment']")
		  If al.Length = 0 Then Return
		  Var a As XmlNode = al.Item(0)
		  Select Case a.GetAttribute("horizontal")
		  Case "left"
		    st.AlignH = XLSXEnums.eAlignH.Left
		  Case "center", "centerContinuous"
		    st.AlignH = XLSXEnums.eAlignH.Center
		  Case "right"
		    st.AlignH = XLSXEnums.eAlignH.Right
		  End Select
		  Select Case a.GetAttribute("vertical")
		  Case "top"
		    st.AlignV = XLSXEnums.eAlignV.Top
		  Case "center"
		    st.AlignV = XLSXEnums.eAlignV.Middle
		  End Select
		  Var wrap As String = a.GetAttribute("wrapText")
		  If wrap = "1" Or wrap = "true" Then st.WrapText = True
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub CopyFont(src As XLSXCellStyle, dst As XLSXCellStyle)
		  dst.Bold = src.Bold
		  dst.Italic = src.Italic
		  dst.Underline = src.Underline
		  dst.FontName = src.FontName
		  dst.FontSize = src.FontSize
		  dst.FontColor = src.FontColor
		  dst.HasFontColor = src.HasFontColor
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub CopyFill(src As XLSXCellStyle, dst As XLSXCellStyle)
		  dst.BackgroundColor = src.BackgroundColor
		  dst.HasBackground = src.HasBackground
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub CopyBorder(src As XLSXCellStyle, dst As XLSXCellStyle)
		  dst.BorderLeft = src.BorderLeft
		  dst.BorderRight = src.BorderRight
		  dst.BorderTop = src.BorderTop
		  dst.BorderBottom = src.BorderBottom
		  dst.BorderColor = src.BorderColor
		  dst.HasBorderColor = src.HasBorderColor
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function FlagOn(parent As XmlNode, localName As String) As Boolean
		  ' A boolean font flag (<b>, <i>): on when the element is present and its
		  ' val attribute isn't an explicit off ("0" / "false").
		  Var list As XmlNodeList = parent.Xql("./*[local-name()='" + localName + "']")
		  If list.Length = 0 Then Return False
		  Var v As String = list.Item(0).GetAttribute("val")
		  Return v <> "0" And v <> "false"
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function IntAttr(node As XmlNode, name As String, defaultValue As Integer) As Integer
		  Var v As String = node.GetAttribute(name)
		  If v = "" Then Return defaultValue
		  Return v.ToInteger
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 506172736520616E204F4F584D4C2041415252474742422F5252474742422068657820636F6C6F722E2046616C736520666F72207468656D652F696E64657865642F656D70747920286E6F74207265736F6C76656420696E205631292E0A
		Private Function ParseHexColor(rgb As String, ByRef c As Color) As Boolean
		  ' OOXML hex colors are "AARRGGBB" (8 hex) or "RRGGBB" (6 hex). Returns False
		  ' for empty / non-hex input (theme & indexed colors are handled in ResolveColor).
		  Var s As String = rgb.Trim
		  If s.Length = 8 Then s = s.Middle(2)   ' drop alpha
		  If s.Length <> 6 Then Return False
		  Var r As Integer = Integer.FromHex(s.Middle(0, 2))
		  Var g As Integer = Integer.FromHex(s.Middle(2, 2))
		  Var b As Integer = Integer.FromHex(s.Middle(4, 2))
		  c = Color.RGB(r, g, b)
		  Return True
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 5265736F6C766520616E204F4F584D4C20636F6C6F757220656C656D656E742028726762202F207468656D652B74696E74202F20696E64657865642920746F206120436F6C6F722E2046616C7365207768656E20756E7265736F6C7661626C652028652E672E206175746F206F7220756E6B6E6F776E207468656D6520696E646578292E0A
		Private Function ResolveColor(node As XmlNode, ByRef c As Color) As Boolean
		  ' Resolve any OOXML colour element (<color>, <fgColor>, edge <color>) to a
		  ' Color, in priority order: explicit rgb, then theme index (+ optional tint),
		  ' then the legacy indexed palette. Returns False for auto / unresolved.
		  Var rgb As String = node.GetAttribute("rgb")
		  If rgb <> "" Then Return ParseHexColor(rgb, c)

		  Var themeAttr As String = node.GetAttribute("theme")
		  If themeAttr <> "" Then
		    Var tintAttr As String = node.GetAttribute("tint")
		    Var tint As Double = If(tintAttr <> "", tintAttr.ToDouble, 0.0)
		    Return ThemeColorAt(Integer.FromString(themeAttr), tint, c)
		  End If

		  Var idxAttr As String = node.GetAttribute("indexed")
		  If idxAttr <> "" Then Return IndexedColorAt(Integer.FromString(idxAttr), c)

		  Return False   ' auto="1" or nothing usable
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 5061727365207468656D65312E786D6C203C613A636C72536368656D653E20696E746F207468652031322D656E74727920696E64657865642070616C657474652C206170706C79696E6720457863656C277320646B2F6C7420737761702028302D3E6C74312C20312D3E646B312C20322D3E6C74322C20332D3E646B322C20342E2E392D3E616363656E74312E2E36292E0A
		Private Sub ParseTheme(themeXml As String)
		  ' Read the theme colour scheme (dk1/lt1/dk2/lt2/accent1..6/hlink/folHlink)
		  ' into mTheme, indexed the way <color theme="N"/> expects: Excel swaps the
		  ' first two pairs, so index 0->lt1, 1->dk1, 2->lt2, 3->dk2, 4..9->accent1..6.
		  If themeXml = "" Then Return
		  Var doc As New XmlDocument
		  Try
		    doc.LoadXml(themeXml)
		  Catch
		    Return   ' a broken theme just means we fall back to no theme colours
		  End Try

		  Var dk1, lt1, dk2, lt2 As Color
		  Call SchemeColor(doc, "dk1", dk1)
		  Call SchemeColor(doc, "lt1", lt1)
		  Call SchemeColor(doc, "dk2", dk2)
		  Call SchemeColor(doc, "lt2", lt2)
		  mTheme.Add lt1   ' 0
		  mTheme.Add dk1   ' 1
		  mTheme.Add lt2   ' 2
		  mTheme.Add dk2   ' 3
		  For Each name As String In Array("accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink")
		    Var c As Color
		    Call SchemeColor(doc, name, c)
		    mTheme.Add c   ' 4..11
		  Next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 526561642061206E616D656420636C72536368656D65206368696C642028652E672E20616363656E7431293A207072656665722073726762436C722076616C2C20656C73652066616C6C206261636B20746F20737973436C72206C617374436C722E0A
		Private Function SchemeColor(doc As XmlDocument, name As String, ByRef c As Color) As Boolean
		  ' One <a:clrScheme>/<a:NAME> child holds either <a:srgbClr val="RRGGBB"/> or
		  ' <a:sysClr val="window…" lastClr="RRGGBB"/>. Prefer srgbClr; fall back to
		  ' sysClr's cached lastClr (what Office wrote for the live system colour).
		  Var list As XmlNodeList = doc.Xql("//*[local-name()='clrScheme']/*[local-name()='" + name + "']")
		  If list.Length = 0 Then Return False
		  Var el As XmlNode = list.Item(0)
		  Var srgb As XmlNodeList = el.Xql("./*[local-name()='srgbClr']")
		  If srgb.Length > 0 Then Return ParseHexColor(srgb.Item(0).GetAttribute("val"), c)
		  Var sys As XmlNodeList = el.Xql("./*[local-name()='sysClr']")
		  If sys.Length > 0 Then Return ParseHexColor(sys.Item(0).GetAttribute("lastClr"), c)
		  Return False
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 5265736F6C76652061203C636F6C6F72207468656D653D224E222074696E743D222E2E2E222F3E20746F206120436F6C6F722C206170706C79696E67207468652074696E74207768656E206E6F6E2D7A65726F2E0A
		Private Function ThemeColorAt(index As Integer, tint As Double, ByRef c As Color) As Boolean
		  ' Map a theme index to its palette colour and apply the OOXML tint, if any.
		  If index < 0 Or index > mTheme.LastIndex Then Return False
		  Var base As Color = mTheme(index)
		  If tint = 0.0 Then
		    c = base
		  Else
		    c = ApplyTint(base, tint)
		  End If
		  Return True
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 4170706C7920746865204F4F584D4C2074696E7420286C69676874656E2F6461726B656E2920746F206120436F6C6F722076696120746865204C206368616E6E656C206F662048534C2E0A
		Private Function ApplyTint(base As Color, tint As Double) As Color
		  ' OOXML tint lightens (tint>0) or darkens (tint<0) by adjusting the HSL
		  ' lightness: L' = L*(1+tint) for tint<0, else L*(1-tint)+tint.
		  Var r As Double = base.Red / 255.0
		  Var g As Double = base.Green / 255.0
		  Var b As Double = base.Blue / 255.0

		  ' RGB -> HSL.
		  Var mx As Double = r
		  If g > mx Then mx = g
		  If b > mx Then mx = b
		  Var mn As Double = r
		  If g < mn Then mn = g
		  If b < mn Then mn = b
		  Var l As Double = (mx + mn) / 2.0
		  Var h As Double = 0.0
		  Var s As Double = 0.0
		  If mx <> mn Then
		    Var d As Double = mx - mn
		    If l > 0.5 Then
		      s = d / (2.0 - mx - mn)
		    Else
		      s = d / (mx + mn)
		    End If
		    If mx = r Then
		      h = (g - b) / d
		      If g < b Then h = h + 6.0
		    ElseIf mx = g Then
		      h = (b - r) / d + 2.0
		    Else
		      h = (r - g) / d + 4.0
		    End If
		    h = h / 6.0
		  End If

		  ' Apply tint to lightness, clamped to 0..1.
		  If tint < 0.0 Then
		    l = l * (1.0 + tint)
		  Else
		    l = l * (1.0 - tint) + tint
		  End If
		  If l < 0.0 Then l = 0.0
		  If l > 1.0 Then l = 1.0

		  ' HSL -> RGB.
		  If s = 0.0 Then
		    r = l
		    g = l
		    b = l
		  Else
		    Var q As Double
		    If l < 0.5 Then
		      q = l * (1.0 + s)
		    Else
		      q = l + s - l * s
		    End If
		    Var p As Double = 2.0 * l - q
		    r = Hue2Rgb(p, q, h + 1.0 / 3.0)
		    g = Hue2Rgb(p, q, h)
		    b = Hue2Rgb(p, q, h - 1.0 / 3.0)
		  End If

		  Return Color.RGB(ClampByte(r), ClampByte(g), ClampByte(b))
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 48534C2068656C7065723A20636F6E76657274206F6E6520687565207365676D656E7420746F2061206E6F726D616C697A656420524742206368616E6E656C2E0A
		Private Function Hue2Rgb(p As Double, q As Double, t As Double) As Double
		  If t < 0.0 Then t = t + 1.0
		  If t > 1.0 Then t = t - 1.0
		  If t < 1.0 / 6.0 Then Return p + (q - p) * 6.0 * t
		  If t < 1.0 / 2.0 Then Return q
		  If t < 2.0 / 3.0 Then Return p + (q - p) * (2.0 / 3.0 - t) * 6.0
		  Return p
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 5363616C652061206E6F726D616C697A656420302E2E31206368616E6E656C20746F206120726F756E6465642C20636C616D70656420302E2E32353520496E74656765722E0A
		Private Function ClampByte(v As Double) As Integer
		  ' Scale a normalized 0..1 channel to 0..255 with rounding + clamping.
		  Var n As Double = v * 255.0
		  If n < 0.0 Then n = 0.0
		  If n > 255.0 Then n = 255.0
		  Return Round(n)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 5265736F6C76652061206C6567616379203C636F6C6F7220696E64657865643D224E222F3E2066726F6D20746865207374616E646172642035362D636F6C6F75722070616C65747465202836343D626C61636B2C2036353D7768697465292E0A
		Private Function IndexedColorAt(index As Integer, ByRef c As Color) As Boolean
		  ' The legacy Excel indexed palette. 64 = system foreground (black),
		  ' 65 = system background (white); everything else maps to the classic table.
		  Var palette() As String = Array( _
		  "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", _
		  "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", _
		  "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080", _
		  "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF", _
		  "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF", _
		  "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99", _
		  "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696", _
		  "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333")
		  If index = 64 Or index = 65 Then
		    c = If(index = 64, Color.RGB(0, 0, 0), Color.RGB(255, 255, 255))
		    Return True
		  End If
		  If index < 0 Or index > palette.LastIndex Then Return False
		  Return ParseHexColor(palette(index), c)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 53796E74686573697A652074686520584C535843656C6C5374796C652061206275696C742D696E207461626C65207374796C6520696D706C69657320666F72206120726F6C653A206865616465722028736F6C696420616363656E742066696C6C202B20776869746520626F6C6429206F722061207374726970656420626F647920726F7720286C6967687420616363656E742074696E74292E20506C61696E2063656C6C732067657420612064656661756C74207374796C652E0A
		Function TableStyleCell(styleName As String, isHeader As Boolean, striped As Boolean) As XLSXCellStyle
		  ' Synthesize the cell style a built-in table style implies for one role:
		  ' header (solid accent fill + white bold text) or a striped body row (a
		  ' light accent tint). Plain body cells get a default (empty) style. The
		  ' accent comes from the style-name number resolved against the theme.
		  Var st As New XLSXCellStyle
		  If styleName = "" Then Return st

		  Var accent As Color
		  Var hasAccent As Boolean = TableStyleAccent(styleName, accent)

		  If isHeader Then
		    st.Bold = True
		    st.HasBackground = True
		    st.HasFontColor = True
		    st.FontColor = Color.RGB(255, 255, 255)
		    st.BackgroundColor = If(hasAccent, accent, Color.RGB(64, 64, 64))
		  ElseIf striped Then
		    st.HasBackground = True
		    st.BackgroundColor = If(hasAccent, ApplyTint(accent, 0.8), Color.RGB(217, 217, 217))
		  End If
		  Return st
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 4D61702061206275696C742D696E207461626C65207374796C65206E616D6520746F20697473207468656D6520616363656E7420636F6C6F75723B2074686520747261696C696E67206E756D626572206379636C657320696E2067726F757073206F6620736576656E2028303D6E6F20616363656E742C20312E2E363D616363656E74312E2E36292E0A
		Private Function TableStyleAccent(styleName As String, ByRef accent As Color) As Boolean
		  ' Map a built-in table style name to its theme accent colour. The trailing
		  ' number cycles in groups of seven (position 0 = no accent, 1..6 = accent
		  ' 1..6) — e.g. TableStyleLight9 -> (9-1) mod 7 = 1 -> accent1.
		  Var digits As String = ""
		  Var allDigits As String = "0123456789"
		  For i As Integer = styleName.Length - 1 DownTo 0
		    Var ch As String = styleName.Middle(i, 1)
		    If allDigits.IndexOf(ch) >= 0 Then
		      digits = ch + digits
		    Else
		      Exit
		    End If
		  Next
		  If digits = "" Then Return False
		  Var pos As Integer = (Integer.FromString(digits) - 1) Mod 7
		  If pos <= 0 Then Return False   ' position 0 = no accent (greyscale style)
		  Return ThemeColorAt(3 + pos, 0.0, accent)   ' accent1 -> theme index 4
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 4D617020616E204F4F584D4C20626F72646572207374796C65206E616D6520746F207468652065426F726465725374796C652073756273657420284E6F6E652F5468696E2F4D656469756D2F546869636B292E0A
		Private Function BorderStyleFromString(s As String) As XLSXEnums.eBorderStyle
		  Select Case s
		  Case "", "none"
		    ' No border. Excel omits the style attribute (""); LibreOffice writes
		    ' style="none" explicitly — both mean the edge has no border.
		    Return XLSXEnums.eBorderStyle.None
		  Case "medium", "mediumDashed", "mediumDashDot", "mediumDashDotDot"
		    Return XLSXEnums.eBorderStyle.Medium
		  Case "thick", "double"
		    Return XLSXEnums.eBorderStyle.Thick
		  Else
		    ' thin, hair, dotted, dashed, dashDot, … all collapse to Thin.
		    Return XLSXEnums.eBorderStyle.Thin
		  End Select
		End Function
	#tag EndMethod

	#tag Property, Flags = &h21, Description = 4D61702066726F6D206E756D466D7449642028496E74656765722920746F20666F726D617420636F64652028537472696E67292E0A
		Private mNumFmts As Dictionary
	#tag EndProperty

	#tag Property, Flags = &h21, Description = 4D61702066726F6D2063656C6C586620696E6465782028496E74656765722C20302D62617365642920746F206E756D466D7449642028496E7465676572292E0A
		Private mCellXfs As Dictionary
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mCellStyles As Dictionary
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mFonts() As XLSXCellStyle
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mFills() As XLSXCellStyle
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mBorders() As XLSXCellStyle
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mEmptyStyle As XLSXCellStyle
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mTheme() As Color
	#tag EndProperty

	#tag Note, Name = About
		Parses xl/styles.xml from an XLSX archive and resolves a cellXf index
		to a format code.

		Public surface:
		  Constructor(stylesXml As String)
		    Pass the UTF-8 text of styles.xml. Empty input is fine — we still
		    seed built-in numFmt ids.
		  NumberFormatCodeAt(cellXfIndex As Integer) As String
		    Returns "0.00" / "dd/mm/yyyy" / "General" / etc. — feed it to
		    XLSXFormatter.Format().
		  CellXfCount() As Integer
		    Total number of <xf> entries declared in <cellXfs>.

		Resolution chain:
		  cell @s attribute (an integer index into <cellXfs>)
		    -> <xf>.numFmtId
		      -> <numFmts>/<numFmt formatCode=...> if id >= 164
		      -> built-in code from SeedBuiltInNumFmts otherwise
		      -> "" (empty) on miss

		We only read numFmt codes — fonts, fills, borders, alignment are out
		of V1 scope.
	#tag EndNote

	#tag ViewBehavior
	#tag EndViewBehavior
End Class
#tag EndClass
