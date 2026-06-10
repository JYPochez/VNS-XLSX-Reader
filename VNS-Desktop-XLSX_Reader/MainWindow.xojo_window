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
   Begin DesktopButton ButtonNew
      AllowAutoDeactivate=   True
      Bold            =   False
      Cancel          =   False
      Caption         =   "#strings.kStrNewButton"
      Default         =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   24
      Index           =   -2147483648
      Italic          =   False
      Left            =   156
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      MacButtonStyle  =   0
      Scope           =   0
      TabIndex        =   6
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   8
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   100
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
      Left            =   8
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
      Top             =   44
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
      Left            =   176
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
      Top             =   44
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   290
   End
   Begin DesktopButton ButtonAddRow
      AllowAutoDeactivate=   True
      Bold            =   False
      Cancel          =   False
      Caption         =   "#strings.kStrAddRow"
      Default         =   False
      Enabled         =   False
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   24
      Index           =   -2147483648
      Italic          =   False
      Left            =   596
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   True
      MacButtonStyle  =   0
      Scope           =   0
      TabIndex        =   7
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   40
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   70
   End
   Begin DesktopButton ButtonDelRow
      AllowAutoDeactivate=   True
      Bold            =   False
      Cancel          =   False
      Caption         =   "#strings.kStrDelRow"
      Default         =   False
      Enabled         =   False
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   24
      Index           =   -2147483648
      Italic          =   False
      Left            =   670
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   True
      MacButtonStyle  =   0
      Scope           =   0
      TabIndex        =   8
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   40
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   70
   End
   Begin DesktopButton ButtonAddCol
      AllowAutoDeactivate=   True
      Bold            =   False
      Cancel          =   False
      Caption         =   "#strings.kStrAddCol"
      Default         =   False
      Enabled         =   False
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   24
      Index           =   -2147483648
      Italic          =   False
      Left            =   744
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   True
      MacButtonStyle  =   0
      Scope           =   0
      TabIndex        =   9
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   40
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   70
   End
   Begin DesktopButton ButtonDelCol
      AllowAutoDeactivate=   True
      Bold            =   False
      Cancel          =   False
      Caption         =   "#strings.kStrDelCol"
      Default         =   False
      Enabled         =   False
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   24
      Index           =   -2147483648
      Italic          =   False
      Left            =   818
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   True
      MacButtonStyle  =   0
      Scope           =   0
      TabIndex        =   10
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   40
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   70
   End
   Begin DesktopButton ButtonSave
      AllowAutoDeactivate=   True
      Bold            =   False
      Cancel          =   False
      Caption         =   "#strings.kStrSaveButton"
      Default         =   False
      Enabled         =   False
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   24
      Index           =   -2147483648
      Italic          =   False
      Left            =   752
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   True
      MacButtonStyle  =   0
      Scope           =   0
      TabIndex        =   5
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   8
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   140
   End
   Begin DesktopTabPanel TabPanelSheets
      AllowAutoDeactivate=   True
      Bold            =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   528
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
      Top             =   72
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
      GridLineStyle   =   3
      HasBorder       =   True
      HasHeader       =   True
      HasHorizontalScrollbar=   True
      HasVerticalScrollbar=   True
      HeadingIndex    =   -1
      Height          =   484
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
      Top             =   104
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
		  XLSXDesktopListboxFiller.Fill(ListboxData, sheet, mWorkbook.Styles, mShowAllCells)
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
		  EnableEditButtons
		  RebuildTabs(mWorkbook)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub EnableEditButtons()
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
		    mShowAllCells = False
		    EnableEditButtons
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


	#tag Method, Flags = &h21
		Private Sub SaveWorkbook()
		  If mWorkbook Is Nil Then Return

		  Var dlg As New SaveFileDialog
		  dlg.Title = strings.kStrSaveDialogTitle
		  Var tXlsx As New FileType
		  tXlsx.Name = strings.kStrFileTypeXlsx
		  tXlsx.Extensions = "xlsx"
		  Var tOds As New FileType
		  tOds.Name = strings.kStrFileTypeOds
		  tOds.Extensions = "ods"
		  dlg.Filter = tXlsx + tOds

		  ' Default to the format the workbook was loaded from; the user picks the
		  ' other format by changing the extension in the dialog.
		  Var defaultExt As String = ".xlsx"
		  If mWorkbook.SourceName.Lowercase.EndsWith(".ods") Then defaultExt = ".ods"
		  dlg.SuggestedFileName = BaseName(mWorkbook.SourceName) + defaultExt

		  Var f As FolderItem = dlg.ShowModal(Self)
		  If f = Nil Then Return

		  Var lower As String = f.Name.Lowercase
		  If Not lower.EndsWith(".xlsx") And Not lower.EndsWith(".ods") Then
		    f = f.Parent.Child(f.Name + defaultExt)
		  End If

		  Try
		    SpreadsheetWriter.Save(mWorkbook, f)
		    LabelParseTime.Text = strings.kStrSavedPrefix + f.Name
		  Catch ex As XLSXException
		    ShowSaveError(ex.Detail)
		  Catch ex As IOException
		    ShowSaveError(ex.Message)
		  End Try
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub ApplyCellEdit(lbRow As Integer, lbCol As Integer)
		  ' Write a committed listbox edit back into the workbook model.
		  If mWorkbook Is Nil Then Return
		  Var idx As Integer = TabPanelSheets.SelectedPanelIndex
		  If idx < 0 Then Return
		  Var sheet As XLSXSheet = mWorkbook.SheetAt(idx + 1)
		  If sheet Is Nil Then Return

		  Var sheetRow As Integer = ListboxData.RowTagAt(lbRow).IntegerValue
		  Var sheetCol As Integer = lbCol + 1
		  If sheetRow <= 0 Or sheetCol <= 0 Then Return

		  Var newText As String = ListboxData.CellTextAt(lbRow, lbCol)
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
		  Var d As New MessageDialog
		  d.Title = strings.kStrErrorSaveTitle
		  d.Message = strings.kStrErrorSaveTitle
		  d.Explanation = msg
		  Call d.ShowModal
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
		Private mShowAllCells As Boolean = False
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
		Function CellPressed(row As Integer, column As Integer, x As Integer, y As Integer) As Boolean
		  ' Click a cell to edit it in place; CellAction commits the edit.
		  If mWorkbook Is Nil Then Return False
		  If row < 0 Or column < 0 Then Return False
		  Me.CellTypeAt(row, column) = DesktopListBox.CellTypes.TextField
		  Me.EditCellAt(row, column)
		  Return True
		End Function
	#tag EndEvent
	#tag Event
		Sub CellAction(row As Integer, column As Integer)
		  ApplyCellEdit(row, column)
		End Sub
	#tag EndEvent
#tag EndEvents
