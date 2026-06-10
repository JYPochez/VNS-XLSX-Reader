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
      Width           =   380
      _mPanelIndex    =   -1
   End
   Begin WebPopupMenu PopupFormat
      ControlID       =   ""
      CSSClasses      =   ""
      Enabled         =   False
      FontName        =   ""
      FontSize        =   "0.0"
      Height          =   32
      Index           =   -2147483648
      Indicator       =   0
      InitialValue    =   ""
      Left            =   600
      LockBottom      =   False
      LockedInPosition=   False
      LockHorizontal  =   False
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   True
      LockVertical    =   False
      PanelIndex      =   0
      Scope           =   0
      TabIndex        =   5
      TabStop         =   True
      Tooltip         =   ""
      Top             =   52
      Visible         =   True
      Width           =   148
      _mPanelIndex    =   -1
   End
   Begin WebButton ButtonSave
      Bold            =   "False"
      Cancel          =   False
      Caption         =   "#strings.kStrSaveButton"
      ControlID       =   ""
      CSSClasses      =   ""
      Default         =   False
      Enabled         =   False
      FontName        =   ""
      FontSize        =   "0.0"
      Height          =   32
      Index           =   -2147483648
      Indicator       =   0
      Italic          =   "False"
      Left            =   760
      LockBottom      =   False
      LockedInPosition=   False
      LockHorizontal  =   False
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   True
      LockVertical    =   False
      PanelIndex      =   0
      Scope           =   0
      TabIndex        =   6
      TabStop         =   True
      Tooltip         =   ""
      Top             =   52
      Underline       =   "False"
      Visible         =   True
      Width           =   124
      _mPanelIndex    =   -1
   End
   Begin WebButton ButtonNew
      Bold            =   "False"
      Cancel          =   False
      Caption         =   "#strings.kStrNewButton"
      ControlID       =   ""
      CSSClasses      =   ""
      Default         =   False
      Enabled         =   True
      FontName        =   ""
      FontSize        =   "0.0"
      Height          =   32
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
      TabIndex        =   7
      TabStop         =   True
      Tooltip         =   ""
      Top             =   92
      Underline       =   "False"
      Visible         =   True
      Width           =   100
      _mPanelIndex    =   -1
   End
   Begin WebButton ButtonAddRow
      Bold            =   "False"
      Cancel          =   False
      Caption         =   "#strings.kStrAddRow"
      ControlID       =   ""
      CSSClasses      =   ""
      Default         =   False
      Enabled         =   False
      FontName        =   ""
      FontSize        =   "0.0"
      Height          =   32
      Index           =   -2147483648
      Indicator       =   0
      Italic          =   "False"
      Left            =   132
      LockBottom      =   False
      LockedInPosition=   False
      LockHorizontal  =   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      LockVertical    =   False
      PanelIndex      =   0
      Scope           =   0
      TabIndex        =   8
      TabStop         =   True
      Tooltip         =   ""
      Top             =   92
      Underline       =   "False"
      Visible         =   True
      Width           =   80
      _mPanelIndex    =   -1
   End
   Begin WebButton ButtonDelRow
      Bold            =   "False"
      Cancel          =   False
      Caption         =   "#strings.kStrDelRow"
      ControlID       =   ""
      CSSClasses      =   ""
      Default         =   False
      Enabled         =   False
      FontName        =   ""
      FontSize        =   "0.0"
      Height          =   32
      Index           =   -2147483648
      Indicator       =   0
      Italic          =   "False"
      Left            =   220
      LockBottom      =   False
      LockedInPosition=   False
      LockHorizontal  =   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      LockVertical    =   False
      PanelIndex      =   0
      Scope           =   0
      TabIndex        =   9
      TabStop         =   True
      Tooltip         =   ""
      Top             =   92
      Underline       =   "False"
      Visible         =   True
      Width           =   80
      _mPanelIndex    =   -1
   End
   Begin WebButton ButtonAddCol
      Bold            =   "False"
      Cancel          =   False
      Caption         =   "#strings.kStrAddCol"
      ControlID       =   ""
      CSSClasses      =   ""
      Default         =   False
      Enabled         =   False
      FontName        =   ""
      FontSize        =   "0.0"
      Height          =   32
      Index           =   -2147483648
      Indicator       =   0
      Italic          =   "False"
      Left            =   308
      LockBottom      =   False
      LockedInPosition=   False
      LockHorizontal  =   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      LockVertical    =   False
      PanelIndex      =   0
      Scope           =   0
      TabIndex        =   10
      TabStop         =   True
      Tooltip         =   ""
      Top             =   92
      Underline       =   "False"
      Visible         =   True
      Width           =   80
      _mPanelIndex    =   -1
   End
   Begin WebButton ButtonDelCol
      Bold            =   "False"
      Cancel          =   False
      Caption         =   "#strings.kStrDelCol"
      ControlID       =   ""
      CSSClasses      =   ""
      Default         =   False
      Enabled         =   False
      FontName        =   ""
      FontSize        =   "0.0"
      Height          =   32
      Index           =   -2147483648
      Indicator       =   0
      Italic          =   "False"
      Left            =   396
      LockBottom      =   False
      LockedInPosition=   False
      LockHorizontal  =   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      LockVertical    =   False
      PanelIndex      =   0
      Scope           =   0
      TabIndex        =   11
      TabStop         =   True
      Tooltip         =   ""
      Top             =   92
      Underline       =   "False"
      Visible         =   True
      Width           =   80
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
      Top             =   132
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
      Height          =   460
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
      Top             =   172
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
		  XLSXWebListboxFiller.Fill(ListboxData, sheet, mWorkbook.Styles, mShowAllCells)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function CurrentSheet() As XLSXSheet
		  If mWorkbook Is Nil Then Return Nil
		  Var idx As Integer = TabPanelSheets.SelectedPanelIndex
		  If idx < 0 Then Return Nil
		  Return mWorkbook.SheetAt(idx + 1)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub DoNewWorkbook()
		  ' Start a sheet from scratch: empty workbook, one sheet, 8x4 grid,
		  ' shown in grid mode (all cells visible and editable).
		  mWorkbook = XLSXHelpers.NewWorkbook(strings.kStrUntitledName, 8, 4)
		  mShowAllCells = True
		  Self.Title = strings.kStrAppTitle + " — " + mWorkbook.SourceName
		  LabelParseTime.Text = ""
		  EnableEditControls(False)
		  RebuildTabs(mWorkbook)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub EnableEditControls(isOds As Boolean)
		  If PopupFormat.RowCount = 0 Then
		    PopupFormat.AddRow(strings.kStrFileTypeXlsx)
		    PopupFormat.AddRow(strings.kStrFileTypeOds)
		  End If
		  PopupFormat.SelectedRowIndex = If(isOds, 1, 0)
		  PopupFormat.Enabled = True
		  ButtonSave.Enabled = True
		  ButtonAddRow.Enabled = True
		  ButtonDelRow.Enabled = True
		  ButtonAddCol.Enabled = True
		  ButtonDelCol.Enabled = True
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub RefreshAfterStructureChange()
		  ' Structural edits switch to grid mode so the added/removed row or
		  ' column is actually visible (viewing mode hides empty rows).
		  mShowAllCells = True
		  FillCurrentSheet
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
		    mShowAllCells = False
		    EnableEditControls(file.Name.Lowercase.EndsWith(".ods"))
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


	#tag Method, Flags = &h21
		Private Sub SaveWorkbook()
		  ' Serialize to the format picked in PopupFormat and push it to the
		  ' browser as a download.
		  If mWorkbook Is Nil Then Return
		  Var wantOds As Boolean = (PopupFormat.SelectedRowIndex = 1)
		  Try
		    Var data As MemoryBlock
		    If wantOds Then
		      data = ODSWriter.ToMemoryBlock(mWorkbook)
		    Else
		      data = XLSXWriter.ToMemoryBlock(mWorkbook)
		    End If
		    mDownloadFile = New WebFile
		    If wantOds Then
		      mDownloadFile.MimeType = "application/vnd.oasis.opendocument.spreadsheet"
		      mDownloadFile.FileName = BaseName(mWorkbook.SourceName) + ".ods"
		    Else
		      mDownloadFile.MimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
		      mDownloadFile.FileName = BaseName(mWorkbook.SourceName) + ".xlsx"
		    End If
		    mDownloadFile.ForceDownload = True
		    mDownloadFile.Data = data.StringValue(0, data.Size)
		    LabelParseTime.Text = strings.kStrSavedPrefix + mDownloadFile.FileName
		    Self.GoToURL(mDownloadFile.URL)
		  Catch ex As XLSXException
		    ShowSaveError(ex.Detail)
		  End Try
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub ApplyCellEdit(lbRow As Integer, lbCol As Integer, newText As String)
		  ' Write a committed listbox edit back into the workbook model.
		  If mWorkbook Is Nil Then Return
		  Var idx As Integer = TabPanelSheets.SelectedPanelIndex
		  If idx < 0 Then Return
		  Var sheet As XLSXSheet = mWorkbook.SheetAt(idx + 1)
		  If sheet Is Nil Then Return

		  Var sheetRow As Integer = ListboxData.CellTagAt(lbRow, 0).IntegerValue
		  Var sheetCol As Integer = lbCol + 1
		  If sheetRow <= 0 Or sheetCol <= 0 Then Return

		  Var oldCell As XLSXCell = sheet.CellAt(sheetRow, sheetCol)
		  Var newCell As XLSXCell

		  If newText.Trim = "" Then
		    newCell = New XLSXCell(XLSXEnums.eCellType.Empty, "", oldCell.StyleIndex)
		  ElseIf XLSXHelpers.IsNumericString(newText) Then
		    ' Numeric input keeps the old cell's type family and format, so editing
		    ' a date's serial or a formatted number preserves its rendering.
		    Var newType As XLSXEnums.eCellType = XLSXEnums.eCellType.Number
		    If oldCell.eType = XLSXEnums.eCellType.DateValue Then newType = XLSXEnums.eCellType.DateValue
		    newCell = New XLSXCell(newType, newText.Trim, oldCell.StyleIndex)
		    newCell.FormatCode = oldCell.FormatCode
		  Else
		    newCell = New XLSXCell(XLSXEnums.eCellType.Str, newText, -1)
		  End If

		  sheet.PutCell(sheetRow, sheetCol, newCell)
		  ' Re-render through the formatter so the cell shows its formatted value.
		  ListboxData.CellTextAt(lbRow, lbCol) = newCell.DisplayText(mWorkbook.Styles)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub ShowSaveError(detail As String)
		  Var msg As String = strings.kStrErrorSaveGeneric
		  If detail <> "" Then msg = msg + " (" + detail + ")"
		  Var d As New WebMessageDialog
		  d.Title = strings.kStrErrorSaveTitle
		  d.Message = strings.kStrErrorSaveTitle
		  d.Explanation = msg
		  d.Show
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function BaseName(name As String) As String
		  ' Filename without its last extension ("Book1.xlsx" -> "Book1").
		  Var parts() As String = name.Split(".")
		  If parts.Count > 1 Then parts.RemoveAt(parts.LastIndex)
		  Return String.FromArray(parts, ".")
		End Function
	#tag EndMethod


	#tag Property, Flags = &h21
		Private mWorkbook As XLSXWorkbook
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mDownloadFile As WebFile
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mShowAllCells As Boolean = False
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
#tag Events ButtonSave
	#tag Event
		Sub Pressed()
		  SaveWorkbook
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonNew
	#tag Event
		Sub Pressed()
		  DoNewWorkbook
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonAddRow
	#tag Event
		Sub Pressed()
		  Var sheet As XLSXSheet = CurrentSheet
		  If sheet Is Nil Then Return
		  sheet.AppendRow
		  RefreshAfterStructureChange
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonDelRow
	#tag Event
		Sub Pressed()
		  Var sheet As XLSXSheet = CurrentSheet
		  If sheet Is Nil Or sheet.RowCount <= 1 Then Return
		  sheet.RemoveLastRow
		  RefreshAfterStructureChange
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonAddCol
	#tag Event
		Sub Pressed()
		  Var sheet As XLSXSheet = CurrentSheet
		  If sheet Is Nil Then Return
		  sheet.AppendColumn
		  RefreshAfterStructureChange
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ButtonDelCol
	#tag Event
		Sub Pressed()
		  Var sheet As XLSXSheet = CurrentSheet
		  If sheet Is Nil Or sheet.ColCount <= 1 Then Return
		  sheet.RemoveLastColumn
		  RefreshAfterStructureChange
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ListboxData
	#tag Event
		Sub Pressed(row As Integer, column As Integer)
		  ' Click a cell to edit it in place; CellAction commits the edit.
		  ' Cells default to CellTypes.Normal (not editable), so the type must be
		  ' switched to TextField before EditCellAt has any effect.
		  If mWorkbook Is Nil Then Return
		  If row < 0 Or column < 0 Then Return
		  Me.CellTypeAt(row, column) = WebListBox.CellTypes.TextField
		  Me.EditCellAt(row, column)
		End Sub
	#tag EndEvent
	#tag Event
		Sub CellAction(row As Integer, column As Integer, value As Variant)
		  Var newText As String
		  If value <> Nil Then
		    newText = value.StringValue
		  Else
		    newText = Me.CellTextAt(row, column)
		  End If
		  ApplyCellEdit(row, column, newText)
		End Sub
	#tag EndEvent
#tag EndEvents
