#tag DesktopWindow
Begin DesktopWindow MainWindow
   Backdrop        =   0
   BackgroundColor =   &cFFFFFF
   Composite       =   False
   DefaultLocation =   2
   FullScreen      =   False
   HasBackgroundColor=   False
   HasCloseButton  =   True
   HasFullScreenButton=   True
   HasMaximizeButton=   True
   HasMinimizeButton=   True
   HasTitleBar     =   True
   Height          =   600
   ImplicitInstance=   True
   MacProcID       =   0
   MaximumHeight   =   32000
   MaximumWidth    =   32000
   MenuBar         =   482459647
   MenuBarVisible  =   False
   MinimumHeight   =   200
   MinimumWidth    =   400
   Resizeable      =   True
   Title           =   "#strings.kStrAppTitle"
   Type            =   0
   Visible         =   True
   Width           =   900
   Begin DesktopButton ButtonOpen
      AllowAutoDeactivate=   True
      Bold            =   False
      Cancel          =   False
      Caption         =   "#strings.kStrMenuFileOpen"
      Default         =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   24
      Index           =   -2147483648
      Italic          =   False
      Left            =   8
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      MacButtonStyle  =   0
      Scope           =   0
      TabIndex        =   0
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   8
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   140
   End
   Begin DesktopCheckBox CheckboxInMemory
      AllowAutoDeactivate=   True
      Bold            =   False
      Caption         =   "#strings.kStrInMemory"
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   20
      Index           =   -2147483648
      Italic          =   False
      Left            =   160
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Scope           =   0
      State           =   1
      TabIndex        =   1
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   10
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Value           =   True
      VisualState     =   1
      Width           =   160
   End
   Begin DesktopLabel LabelParseTime
      AllowAutoDeactivate=   True
      Bold            =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   20
      Index           =   -2147483648
      Italic          =   False
      Left            =   324
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      Multiline       =   False
      Scope           =   0
      Selectable      =   False
      TabIndex        =   2
      TabPanelIndex   =   0
      Text            =   ""
      TextAlignment   =   0
      TextColor       =   &c777777
      Tooltip         =   ""
      Top             =   10
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   400
   End
   Begin DesktopTabPanel TabPanelSheets
      AllowAutoDeactivate=   True
      Bold            =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   560
      Index           =   -2147483648
      Italic          =   False
      Left            =   0
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      Panels          =   "(no workbook)"
      Scope           =   0
      SmallTabs       =   False
      TabDefinition   =   "Tab 0\rTab 1"
      TabIndex        =   3
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   40
      Transparent     =   False
      Underline       =   False
      Value           =   0
      Visible         =   True
      Width           =   900
   End
   Begin DesktopListBox ListboxData
      AllowAutoDeactivate=   True
      AllowAutoHideScrollbars=   True
      AllowExpandableRows=   False
      AllowFocusRing  =   True
      AllowResizableColumns=   True
      AllowRowDragging=   False
      AllowRowReordering=   False
      Bold            =   False
      ColumnCount     =   1
      ColumnWidths    =   ""
      DefaultRowHeight=   -1
      DropIndicatorVisible=   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      GridLineStyle   =   0
      HasBorder       =   True
      HasHeader       =   True
      HasHorizontalScrollbar=   True
      HasVerticalScrollbar=   True
      HeadingIndex    =   -1
      Height          =   516
      Index           =   -2147483648
      InitialValue    =   ""
      Italic          =   False
      Left            =   8
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      RequiresSelection=   False
      RowSelectionType=   0
      Scope           =   0
      TabIndex        =   4
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   72
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   884
      _ScrollOffset   =   0
      _ScrollWidth    =   -1
   End
End
#tag EndDesktopWindow

#tag WindowCode
	#tag MenuHandler
		Function FileOpen() As Boolean Handles FileOpen.Action
		  ShowOpenDialog
		  Return True
		End Function
	#tag EndMenuHandler


	#tag Method, Flags = &h21
		Private Sub ShowOpenDialog()
		  Var dlg As New OpenFileDialog
		  Var t As New FileType
		  t.Name = "Spreadsheet"
		  t.Extensions = "xlsx;ods"
		  dlg.Filter = t
		  Var f As FolderItem = dlg.ShowModal(Self)
		  If f = Nil Then Return
		  LoadWorkbook(f)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub FillCurrentSheet()
		  If mWorkbook Is Nil Then Return
		  Var idx As Integer = TabPanelSheets.SelectedPanelIndex
		  If idx < 0 Or idx >= mWorkbook.SheetCount Then Return
		  Var sheet As XLSXSheet = mWorkbook.SheetAt(idx + 1)
		  XLSXDesktopListboxFiller.Fill(ListboxData, sheet, mWorkbook.Styles)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LoadWorkbook(file As FolderItem)
		  Var mode As XLSXEnums.eOpenMode
		  If CheckboxInMemory.Value Then
		    mode = XLSXEnums.eOpenMode.Memory
		  Else
		    mode = XLSXEnums.eOpenMode.Disk
		  End If
		  Try
		    Var wb As XLSXWorkbook = SpreadsheetReader.Open(file, mode)
		    mWorkbook = wb
		    Var zipMs As Integer = Floor(wb.ZipMicroseconds / 1000.0)
		    Var xmlMs As Integer = Floor(wb.XmlMicroseconds / 1000.0)
		    Var totalMs As Integer = zipMs + xmlMs
		    Self.Title = strings.kStrAppTitle + " — " + wb.SourceName + " [" + Str(wb.SheetCount) + "]"
		    LabelParseTime.Text = strings.kStrParseTime + Str(totalMs) + strings.kStrParseTimeUnit _
		      + " (zip " + Str(zipMs) + " + xml " + Str(xmlMs) + ", " + wb.OpenMode.ToString + ")"
		    RebuildTabs(wb)
		  Catch ex As XLSXException
		    LabelParseTime.Text = ""
		    ShowErrorFor(ex)
		  End Try
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub RebuildTabs(wb As XLSXWorkbook)
		  ' Clear existing panels.
		  While TabPanelSheets.PanelCount > 0
		    TabPanelSheets.RemovePanelAt(TabPanelSheets.PanelCount - 1)
		  Wend

		  ' Add one panel per sheet.
		  For Each sn As String In wb.SheetNames
		    TabPanelSheets.AddPanel(sn)
		  Next

		  ' Show the first sheet (PanelIndex on child controls is 1-based; 0 = all panels).
		  If wb.SheetCount > 0 Then
		    TabPanelSheets.SelectedPanelIndex = 0
		    ListboxData.PanelIndex = 1
		    FillCurrentSheet
		  End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub ShowErrorFor(ex As XLSXException)
		  Var msg As String
		  Select Case ex.Code
		  Case XLSXEnums.eParseError.NotAZip
		    msg = strings.kStrErrorNotXLSX
		  Case XLSXEnums.eParseError.MalformedXML
		    msg = strings.kStrErrorMalformed
		  Case XLSXEnums.eParseError.Encrypted
		    msg = strings.kStrErrorEncrypted
		  Case XLSXEnums.eParseError.MissingPart
		    msg = strings.kStrErrorNotXLSX
		  Else
		    msg = strings.kStrErrorGeneric
		  End Select
		  If ex.Detail <> "" Then msg = msg + " (" + ex.Detail + ")"
		  Var d As New MessageDialog
		  d.Title = strings.kStrErrorTitle
		  d.Message = strings.kStrErrorTitle
		  d.Explanation = msg
		  Call d.ShowModal
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h21
		Private mWorkbook As XLSXWorkbook
	#tag EndProperty


#tag EndWindowCode

#tag Events ButtonOpen
	#tag Event
		Sub Pressed()
		  ShowOpenDialog
		End Sub
	#tag EndEvent
#tag EndEvents

#tag Events TabPanelSheets
	#tag Event
		Sub PanelChanged()
		  ListboxData.PanelIndex = TabPanelSheets.SelectedPanelIndex + 1
		  FillCurrentSheet
		End Sub
	#tag EndEvent
#tag EndEvents
