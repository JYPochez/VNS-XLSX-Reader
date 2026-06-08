#tag WebPage
Begin WebPage MainPage
   AllowTabOrderWrap=   True
   Compatibility   =   ""
   ControlCount    =   0
   ControlID       =   ""
   CSSClasses      =   ""
   Enabled         =   False
   Height          =   600
   ImplicitInstance=   True
   Index           =   -2147483648
   Indicator       =   0
   IsImplicitInstance=   False
   LayoutDirection =   0
   LayoutType      =   0
   Left            =   0
   LockBottom      =   False
   LockHorizontal  =   False
   LockLeft        =   True
   LockRight       =   False
   LockTop         =   True
   LockVertical    =   False
   MinimumHeight   =   400
   MinimumWidth    =   600
   PanelIndex      =   0
   ScaleFactor     =   0.0
   TabIndex        =   0
   Title           =   "VNS XLSX Reader"
   Top             =   0
   Visible         =   True
   Width           =   900
   _ImplicitInstance=   False
   _mDesignHeight  =   0
   _mDesignWidth   =   0
   _mName          =   ""
   _mPanelIndex    =   -1
   Begin WebFileUploader UploaderXLSX
      AllowedFileTypes=   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,.xlsx,application/vnd.oasis.opendocument.spreadsheet,.ods"
      Caption         =   "Select"
      ControlID       =   ""
      CSSClasses      =   ""
      Enabled         =   True
      Filter          =   ""
      HasFileNameField=   True
      Height          =   32
      Hint            =   ""
      Index           =   -2147483648
      Indicator       =   0
      Left            =   16
      LockBottom      =   False
      LockedInPosition=   False
      LockHorizontal  =   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      LockVertical    =   False
      MaximumBytes    =   0
      MaximumFileCount=   0
      MultipleFiles   =   "False"
      PanelIndex      =   0
      Scope           =   0
      TabIndex        =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   16
      UploadTimeout   =   0
      Visible         =   True
      Width           =   868
      _mPanelIndex    =   -1
   End
   Begin WebCheckBox CheckboxInMemory
      Bold            =   "False"
      Caption         =   "#strings.kStrInMemory"
      ControlID       =   ""
      CSSClasses      =   ""
      Enabled         =   True
      FontName        =   ""
      FontSize        =   "0.0"
      Height          =   24
      Index           =   -2147483648
      Indicator       =   0
      Italic          =   "False"
      Left            =   16
      LockBottom      =   False
      LockedInPosition=   False
      LockHorizontal  =   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      LockVertical    =   False
      PanelIndex      =   0
      Scope           =   0
      State           =   1
      TabIndex        =   1
      TabStop         =   True
      Tooltip         =   ""
      Top             =   56
      Underline       =   "False"
      Value           =   True
      Visible         =   True
      Width           =   180
      _mPanelIndex    =   -1
   End
   Begin WebLabel LabelParseTime
      Bold            =   "False"
      ControlID       =   ""
      CSSClasses      =   ""
      Enabled         =   True
      FontName        =   ""
      FontSize        =   "0.0"
      Height          =   24
      Index           =   -2147483648
      Indicator       =   0
      Italic          =   "False"
      Left            =   200
      LockBottom      =   False
      LockedInPosition=   False
      LockHorizontal  =   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      LockVertical    =   False
      Multiline       =   False
      PanelIndex      =   0
      Scope           =   0
      TabIndex        =   2
      Text            =   ""
      TextAlignment   =   0
      TextColor       =   &c777777
      Tooltip         =   ""
      Top             =   56
      Underline       =   "False"
      Visible         =   True
      Width           =   680
      _mPanelIndex    =   -1
   End
   Begin WebTabPanel TabPanelSheets
      ControlCount    =   0
      ControlID       =   ""
      CSSClasses      =   ""
      Enabled         =   True
      HasBorder       =   True
      Height          =   32
      Index           =   -2147483648
      Indicator       =   0
      LayoutDirection =   "LayoutDirections.LeftToRight"
      LayoutType      =   "LayoutTypes.Fixed"
      Left            =   16
      LockBottom      =   False
      LockedInPosition=   False
      LockHorizontal  =   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      LockVertical    =   False
      PanelCount      =   2
      PanelIndex      =   0
      Panels          =   "(no workbook)"
      Scope           =   0
      SelectedPanelIndex=   0
      TabDefinition   =   "Tab 0\rTab 1"
      TabIndex        =   3
      TabStop         =   True
      Tooltip         =   ""
      Top             =   92
      Visible         =   True
      Width           =   868
      _mDesignHeight  =   0
      _mDesignWidth   =   0
      _mPanelIndex    =   -1
   End
   Begin WebListBox ListboxData
      AllowAutoHideScrollbars=   "True"
      AllowResizableColumns=   "True"
      AllowRowReordering=   False
      Bold            =   "False"
      ColumnCount     =   1
      ColumnsResizable=   "True"
      ColumnWidths    =   ""
      ControlID       =   ""
      CSSClasses      =   ""
      DefaultRowHeight=   49
      Enabled         =   True
      FontName        =   ""
      FontSize        =   "0.0"
      GridLineStyle   =   3
      HasBorder       =   True
      HasHeader       =   True
      Header          =   ""
      HeaderHeight    =   0
      Height          =   500
      HighlightSortedColumn=   True
      Index           =   -2147483648
      Indicator       =   0
      InitialValue    =   ""
      Italic          =   "False"
      LastAddedRowIndex=   0
      LastColumnIndex =   0
      LastRowIndex    =   0
      Left            =   16
      LockBottom      =   True
      LockedInPosition=   False
      LockHorizontal  =   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      LockVertical    =   False
      NoRowsMessage   =   ""
      PanelIndex      =   0
      ProcessingMessage=   ""
      RowCount        =   0
      RowSelectionType=   1
      Scope           =   0
      SearchCriteria  =   ""
      SelectedRowColor=   &c0d6efd
      SelectedRowIndex=   0
      SelectionType   =   "0"
      TabIndex        =   4
      TabStop         =   True
      Tooltip         =   ""
      Top             =   132
      Underline       =   "False"
      Visible         =   True
      Width           =   868
      _mPanelIndex    =   -1
   End
End
#tag EndWebPage

#tag WindowCode
	#tag Method, Flags = &h21
		Private Sub FillCurrentSheet()
		  If mWorkbook Is Nil Then Return
		  Var idx As Integer = TabPanelSheets.SelectedPanelIndex
		  If idx < 0 Or idx >= mWorkbook.SheetCount Then Return
		  Var sheet As XLSXSheet = mWorkbook.SheetAt(idx + 1)
		  XLSXWebListboxFiller.Fill(ListboxData, sheet, mWorkbook.Styles)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub LoadFromUploadedFile(file As WebUploadedFile)
		  Var mode As XLSXEnums.eOpenMode
		  If CheckboxInMemory.Value Then
		    mode = XLSXEnums.eOpenMode.Memory
		  Else
		    mode = XLSXEnums.eOpenMode.Disk
		  End If
		  Try
		    Var wb As XLSXWorkbook = SpreadsheetReader.Open(file.File, file.Name, mode)
		    mWorkbook = wb
		    Var zipMs As Integer = Floor(wb.ZipMicroseconds / 1000.0)
		    Var xmlMs As Integer = Floor(wb.XmlMicroseconds / 1000.0)
		    Var totalMs As Integer = zipMs + xmlMs
		    Self.Title = strings.kStrAppTitle + " — " + file.Name + " [" + Str(wb.SheetCount) + "]"
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
		  While TabPanelSheets.PanelCount > 0
		    TabPanelSheets.RemovePanelAt(TabPanelSheets.PanelCount - 1)
		  Wend
		  For Each sn As String In wb.SheetNames
		    TabPanelSheets.AddPanel(sn)
		  Next
		  If wb.SheetCount > 0 Then
		    TabPanelSheets.SelectedPanelIndex = 0
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
		  Var d As New WebMessageDialog
		  d.Title = strings.kStrErrorTitle
		  d.Message = strings.kStrErrorTitle
		  d.Explanation = msg
		  d.Show
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h21
		Private mWorkbook As XLSXWorkbook
	#tag EndProperty


#tag EndWindowCode

#tag Events UploaderXLSX
	#tag Event
		Sub UploadFinished(files() As WebUploadedFile)
		  If files.Count > 0 Then LoadFromUploadedFile(files(0))
		End Sub
	#tag EndEvent
	#tag Event
		Sub FileAdded(filename As String, bytes As UInt64, mimeType As String)
		  #Pragma Unused filename
		  #Pragma Unused bytes
		  #Pragma Unused mimeType
		  ' Single-file UX: start the upload immediately when a file is selected.
		  UploaderXLSX.StartUpload
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events TabPanelSheets
	#tag Event
		Sub PanelChanged()
		  FillCurrentSheet
		End Sub
	#tag EndEvent
#tag EndEvents
