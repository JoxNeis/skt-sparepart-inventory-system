Attribute VB_Name = "managerForm"
Option Explicit

' =========================================================
' FORM MANAGER
'
' Shared helpers for worksheets and user forms:
'   - Application state (events, screen updating, cursor,
'     status bar) while a button is running
'   - User messages (info, warning, confirmation)
'   - Reading / clearing single input cells
'   - Date picker (formTanggal), opened by double-click
'   - Text prompt (AskText)
'
' Excel tables (ListObject) are handled by managerTable.
' Errors are shown with managerError.ShowError.
'
' Typical button procedure:
'
'   Public Sub SaveSomething_Click()
'       On Error GoTo ErrorHandler
'       Dim state As FormState
'       Dim errInfo As ErrorInfo
'
'       state = managerForm.BeginBusy("Menyimpan data...")
'       ...
'   CleanUp:
'       managerForm.EndBusy state
'       Exit Sub
'   ErrorHandler:
'       errInfo = managerError.CaptureError()
'       managerForm.EndBusy state
'       managerError.ShowError errInfo
'   End Sub
' =========================================================

Private Const MODULE_NAME As String = "managerForm"

' =========================================================
' ERROR CODES (10300 - 10399)
' Usage: managerError.ThrowError managerForm.CellContainsError, ...
' =========================================================
Public Const CellContainsError As Long = 10303      ' input cell holds #N/A, #VALUE!, ...
Public Const EmptyWorksheet As Long = 10306         ' ws parameter is Nothing

' =========================================================
' DATE PICKER
'
' Cell that opened formTanggal. formTanggal writes the
' selected date back to this cell.
' =========================================================
Public CallerCell As Range

' =========================================================
' APPLICATION STATE SNAPSHOT
' =========================================================
Public Type FormState
    EnableEvents As Boolean
    ScreenUpdating As Boolean
    Cursor As Long
    IsCaptured As Boolean
End Type


' =========================================================
' APPLICATION STATE
' =========================================================

' Save the current application state without changing it.
Public Function CaptureState() As FormState

    Dim state As FormState

    state.EnableEvents = Application.EnableEvents
    state.ScreenUpdating = Application.ScreenUpdating
    state.Cursor = Application.Cursor
    state.IsCaptured = True

    CaptureState = state

End Function


' Save the current state, then switch off events and
' (optionally) screen updating, show a status bar text and
' (optionally) the wait cursor.
Public Function BeginBusy( _
    Optional ByVal statusText As String = vbNullString, _
    Optional ByVal showWaitCursor As Boolean = False, _
    Optional ByVal disableScreenUpdating As Boolean = True) As FormState

    Dim state As FormState

    state = CaptureState()

    Application.EnableEvents = False

    If disableScreenUpdating Then
        Application.ScreenUpdating = False
    End If

    If showWaitCursor Then
        Application.Cursor = xlWait
    End If

    If Len(statusText) > 0 Then
        Application.StatusBar = statusText
    End If

    BeginBusy = state

End Function


' Restore the state saved by CaptureState / BeginBusy and
' clear the status bar.
'
' Safe to call from an error handler (does not use
' On Error, so it does not clear the Err object). When the
' state was never captured, Excel defaults are restored.
Public Sub EndBusy( _
    ByRef state As FormState)

    Application.StatusBar = False

    If state.IsCaptured Then
        Application.Cursor = state.Cursor
        Application.EnableEvents = state.EnableEvents
        Application.ScreenUpdating = state.ScreenUpdating
    Else
        Application.Cursor = xlDefault
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End If

End Sub


' Show a status bar text, optionally with the wait cursor.
Public Sub SetStatus( _
    ByVal statusText As String, _
    Optional ByVal showWaitCursor As Boolean = False)

    Application.StatusBar = statusText

    If showWaitCursor Then
        Application.Cursor = xlWait
    End If

End Sub


' Clear the status bar and reset the cursor.
Public Sub ClearStatus()

    Application.StatusBar = False
    Application.Cursor = xlDefault

End Sub


' =========================================================
' MESSAGES
' =========================================================

Public Sub ShowInfo( _
    ByVal message As String, _
    Optional ByVal title As String = vbNullString)

    MsgBox _
        Prompt:=message, _
        Buttons:=vbInformation + vbOKOnly, _
        title:=DefaultTitle(title)

End Sub


Public Sub ShowWarning( _
    ByVal message As String, _
    Optional ByVal title As String = vbNullString)

    MsgBox _
        Prompt:=message, _
        Buttons:=vbExclamation + vbOKOnly, _
        title:=DefaultTitle(title)

End Sub


' Yes / No question. "No" is the default button.
' Returns True only when the user clicks Yes.
Public Function Confirm( _
    ByVal message As String, _
    Optional ByVal title As String = vbNullString) As Boolean

    Confirm = (MsgBox( _
        Prompt:=message, _
        Buttons:=vbQuestion + vbYesNo + vbDefaultButton2, _
        title:=DefaultTitle(title)) = vbYes)

End Function


Private Function DefaultTitle( _
    ByVal title As String) As String

    If Len(title) = 0 Then
        DefaultTitle = coreConfig.ApplicationName()
    Else
        DefaultTitle = title
    End If

End Function


' =========================================================
' CELLS (single input cells, not tables)
' =========================================================

' True when target overlaps area (both on the same sheet).
Public Function IsInRange( _
    ByVal target As Range, _
    ByVal area As Range) As Boolean

    IsInRange = False

    If target Is Nothing Then Exit Function
    If area Is Nothing Then Exit Function
    If Not target.Worksheet Is area.Worksheet Then Exit Function

    IsInRange = Not (Intersect(target, area) Is Nothing)

End Function


' Raw value of a cell.
Public Function GetCellValue( _
    ByVal ws As Worksheet, _
    ByVal cellAddress As String) As Variant

    RequireWorksheet ws, MODULE_NAME & ".GetCellValue"

    GetCellValue = ws.Range(cellAddress).value

End Function


' Read a cell as trimmed text.
' Empty / Null -> "". Error value -> managerForm.CellContainsError.
Public Function GetCellText( _
    ByVal ws As Worksheet, _
    ByVal cellAddress As String) As String

    Dim value As Variant

    value = GetCellValue(ws, cellAddress)

    If IsError(value) Then
        managerError.ThrowError _
            managerForm.CellContainsError, _
            MODULE_NAME & ".GetCellText", _
            "Cell " & cellAddress & " berisi nilai error."
    End If

    If IsNull(value) Or IsEmpty(value) Then
        GetCellText = vbNullString
    Else
        GetCellText = Trim$(CStr(value))
    End If

End Function


' True when the cell holds a valid date. Does not raise:
' the sheet throws its own error code.
'
'   If Not managerForm.IsCellDate(Me, CELL_ORDER_DATE) Then
'       managerError.ThrowError sheetInputOrder.InvalidOrderDate, ...
'   End If
'   orderDate = DateValue(managerForm.GetCellValue(Me, CELL_ORDER_DATE))
Public Function IsCellDate( _
    ByVal ws As Worksheet, _
    ByVal cellAddress As String) As Boolean

    Dim value As Variant

    value = GetCellValue(ws, cellAddress)

    If IsError(value) Then
        IsCellDate = False
    ElseIf managerValidationText.IsBlank(value) Then
        IsCellDate = False
    Else
        IsCellDate = IsDate(value)
    End If

End Function


' Clear the contents of one or more cells. A single cell
' that is part of a merged area clears the whole area
' (ClearContents on part of a merged area raises 1004).
'   managerForm.ClearCells Me, "B4", "B5", "B6"
Public Sub ClearCells( _
    ByVal ws As Worksheet, _
    ParamArray cellAddresses() As Variant)

    Dim cellAddress As Variant
    Dim target As Range

    RequireWorksheet ws, MODULE_NAME & ".ClearCells"

    For Each cellAddress In cellAddresses

        Set target = ws.Range(CStr(cellAddress))

        If target.Cells.Count = 1 Then
            target.MergeArea.ClearContents
        Else
            target.ClearContents
        End If

    Next cellAddress

End Sub


' =========================================================
' DATE PICKER (formTanggal)
' =========================================================

' Open formTanggal for the target cell. For a merged area
' the date is written to its top-left cell.
Public Sub ShowDatePicker( _
    ByVal target As Range)

    Set CallerCell = target.Cells(1, 1).MergeArea.Cells(1, 1)
    formTanggal.Show
    Set CallerCell = Nothing

End Sub


' For Worksheet_BeforeDoubleClick: open the date picker when
' the double-clicked cell touches dateCells. Returns True
' (and sets cancel, so Excel does not enter edit mode) when
' the picker was shown.
'
'   If managerForm.HandleDatePickerDoubleClick( _
'       Target, Me.Range("K7"), Cancel) Then Exit Sub
Public Function HandleDatePickerDoubleClick( _
    ByVal target As Range, _
    ByVal dateCells As Range, _
    ByRef cancel As Boolean) As Boolean

    HandleDatePickerDoubleClick = False

    If Not IsInRange(target.Cells(1, 1), dateCells) Then Exit Function

    cancel = True
    ShowDatePicker target.Cells(1, 1)

    HandleDatePickerDoubleClick = True

End Function


' =========================================================
' TEXT PROMPT
' =========================================================

' Ask for a line of text. cancelled is True when the user
' pressed Cancel (an empty answer is NOT a cancel).
Public Function AskText( _
    ByVal message As String, _
    ByRef cancelled As Boolean, _
    Optional ByVal title As String = vbNullString, _
    Optional ByVal defaultText As String = vbNullString) As String

    Dim answer As Variant

    answer = Application.InputBox( _
        Prompt:=message, _
        title:=DefaultTitle(title), _
        Default:=defaultText, _
        Type:=2)

    If VarType(answer) = vbBoolean Then
        cancelled = True
        AskText = vbNullString
    Else
        cancelled = False
        AskText = Trim$(CStr(answer))
    End If

End Function


' =========================================================
' INTERNAL
' =========================================================

Private Sub RequireWorksheet( _
    ByVal ws As Worksheet, _
    ByVal errorSource As String)

    If ws Is Nothing Then
        managerError.ThrowError _
            managerForm.EmptyWorksheet, _
            errorSource, _
            "Worksheet tidak boleh kosong."
    End If

End Sub
