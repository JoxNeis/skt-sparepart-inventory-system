Attribute VB_Name = "controllerMDPosMobile"
Option Explicit

' =========================================================
' CONTROLLER - MASTER DATA POS MOBILE (sheet MD_POS_MOBILE)
'
'   Form: LOKASI / PLANT K6 (dropdown), NAMA POS MOBILE K7,
'         KODE POS MOBILE K8
'   BATAL I9:J10, SIMPAN K9:L10, REFRESH I12:L13
'
'   TableMDPosMobile (ID, NAMA POS MOBILE, KODE, EDIT, HAPUS)
'     shows the mobile maintenances of the plant in K6; every
'     refresh also rewrites the POS MOBILE dropdown list
'     ("code - name", LISTS sheet).
'     EDIT  : loads the row into the form (SIMPAN updates it)
'     HAPUS : archives the row
'
' Request sheets call EnsurePlant so the dropdown always
' lists the mobile maintenances of their own plant.
' =========================================================

Private Const MODULE_NAME As String = "controllerMDPosMobile"

' =========================================================
' ERROR CODES (14700 - 14799)
' =========================================================
' (validation errors come from managerSheet / the model)

Private Const HEADER_CELL As String = "I5"
Private Const PLANT_CELL As String = "K6"
Private Const NAME_CELL As String = "K7"
Private Const CODE_CELL As String = "K8"

Private Const BUTTON_CANCEL As String = "I9:J10"
Private Const BUTTON_SAVE As String = "K9:L10"
Private Const BUTTON_REFRESH As String = "I12:L13"

Private Const TABLE_NAME As String = "TableMDPosMobile"
Private Const COL_ID As String = "ID"
Private Const COL_NAME As String = "NAMA POS MOBILE"
Private Const COL_CODE As String = "KODE"
Private Const COL_EDIT As String = "EDIT"
Private Const COL_DELETE As String = "HAPUS"

Private Const HEADER_NEW As String = "FORM PENGISIAN POS MOBILE BARU"
Private Const HEADER_EDIT As String = "FORM EDIT POS MOBILE (ID "

' Plant whose list is in the table (this session).
Private mLoadedPlant As String


' =========================================================
' EVENTS
' =========================================================

Public Sub OnActivate( _
    ByVal ws As Worksheet)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    managerSheet.PreparePlantCell ws, PLANT_CELL
    ShowHeader ws

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerError.ShowError errInfo

End Sub


Public Sub OnDoubleClick( _
    ByVal ws As Worksheet, _
    ByVal target As Range, _
    ByRef cancel As Boolean)

    Dim tbl As ListObject
    Dim rowIndex As Long

    If managerSheet.IsHit(target, ws, BUTTON_SAVE) Then
        cancel = True
        Save ws
        Exit Sub
    ElseIf managerSheet.IsHit(target, ws, BUTTON_CANCEL) Then
        cancel = True
        ClearForm ws
        Exit Sub
    ElseIf managerSheet.IsHit(target, ws, BUTTON_REFRESH) Then
        cancel = True
        Refresh ws
        Exit Sub
    End If

    Set tbl = managerTable.GetTable(ws, TABLE_NAME)
    rowIndex = managerTable.GetRowIndex(tbl, target.Cells(1, 1))
    If rowIndex = 0 Then Exit Sub
    If Not managerTable.HasValue(tbl, rowIndex, COL_ID) Then Exit Sub

    If managerTable.IsInColumn(target.Cells(1, 1), tbl, COL_EDIT) Then
        cancel = True
        LoadForEdit ws, rowIndex
    ElseIf managerTable.IsInColumn(target.Cells(1, 1), tbl, COL_DELETE) Then
        cancel = True
        ArchiveRow ws, rowIndex
    End If

End Sub


Public Sub OnChange( _
    ByVal ws As Worksheet, _
    ByVal target As Range)

    If managerSheet.IsHit(target, ws, PLANT_CELL) Then
        managerSheet.SetEditId ws, 0
        Refresh ws
    End If

End Sub


' =========================================================
' ACTIONS
' =========================================================

Public Sub Save( _
    ByVal ws As Worksheet)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo
    Dim item As MobileMaintenance
    Dim plantName As String
    Dim editId As Long

    plantName = managerSheet.ReadText(ws, PLANT_CELL, "LOKASI / PLANT")
    editId = managerSheet.GetEditId(ws)

    Set item = New MobileMaintenance
    If editId > 0 Then item.Id = editId
    item.Name = managerSheet.ReadText(ws, NAME_CELL, "NAMA POS MOBILE")
    item.Code = managerSheet.ReadText(ws, CODE_CELL, "KODE POS MOBILE")

    state = managerForm.BeginBusy("Menyimpan pos mobile...")

    If editId > 0 Then
        serviceMobileMaintenance.UpdateMobileMaintenance item, plantName
    Else
        serviceMobileMaintenance.CreateMobileMaintenance item, plantName
    End If

    RefreshTable ws
    ClearFormCells ws

    managerForm.EndBusy state

    managerForm.ShowInfo "Pos mobile " & item.Name & " tersimpan."

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


Public Sub ClearForm( _
    ByVal ws As Worksheet)

    ClearFormCells ws

End Sub


Public Sub Refresh( _
    ByVal ws As Worksheet)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo

    state = managerForm.BeginBusy("Memuat data pos mobile...")
    RefreshTable ws
    managerForm.EndBusy state

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


' Switch the sheet to plantName (and reload) when it shows
' another plant. Used by the request sheets; errors are shown,
' never raised.
Public Sub EnsurePlant( _
    ByVal plantName As String)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo
    Dim ws As Worksheet
    Dim eventsWereOn As Boolean

    plantName = Trim$(plantName)
    If Len(plantName) = 0 Then Exit Sub

    Set ws = managerSheet.GetSheet(managerSheet.SHEET_MD_POS_MOBILE)

    If managerValidationText.IsSameText(managerForm.GetCellText(ws, PLANT_CELL), plantName) _
        And managerValidationText.IsSameText(mLoadedPlant, plantName) Then Exit Sub

    eventsWereOn = Application.EnableEvents
    Application.EnableEvents = False

    ws.Range(PLANT_CELL).value = plantName
    managerSheet.SetEditId ws, 0
    ClearFormCells ws
    RefreshTable ws

    Application.EnableEvents = eventsWereOn

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    Application.EnableEvents = True
    managerError.ShowError errInfo

End Sub


' Fill the table with the plant in K6 (no error handling:
' the caller decides).
Public Sub RefreshTable( _
    ByVal ws As Worksheet)

    Dim items As Collection
    Dim item As MobileMaintenance
    Dim data() As Variant
    Dim r As Long
    Dim tbl As ListObject
    Dim plantName As String

    Set tbl = managerTable.GetTable(ws, TABLE_NAME)
    plantName = managerForm.GetCellText(ws, PLANT_CELL)

    If Len(plantName) = 0 Then
        managerTable.SetRowCount tbl, 0
        managerList.WriteMobilePosts Nothing
        mLoadedPlant = vbNullString
        Exit Sub
    End If

    Set items = serviceMobileMaintenance.ListByPlant(plantName)
    mLoadedPlant = plantName

    ' Pos mobile dropdowns ("code - name") follow this table.
    managerList.WriteMobilePosts items

    If items.Count = 0 Then
        managerTable.SetRowCount tbl, 0
        Exit Sub
    End If

    ReDim data(1 To items.Count, 1 To 5)

    For Each item In items
        r = r + 1
        data(r, 1) = item.Id
        data(r, 2) = item.Name
        data(r, 3) = item.Code
        data(r, 4) = managerSheet.LABEL_EDIT
        data(r, 5) = managerSheet.LABEL_DELETE
    Next item

    managerTable.WriteData tbl, data, items.Count, _
        Array("0", managerSheet.TEXT_FORMAT, managerSheet.TEXT_FORMAT, _
              managerSheet.TEXT_FORMAT, managerSheet.TEXT_FORMAT)

End Sub


' =========================================================
' INTERNAL
' =========================================================

Private Sub LoadForEdit( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long)

    Dim tbl As ListObject
    Dim editId As Long

    Set tbl = managerTable.GetTable(ws, TABLE_NAME)
    editId = managerTable.GetLong(tbl, rowIndex, COL_ID)

    ws.Range(NAME_CELL).value = managerTable.GetText(tbl, rowIndex, COL_NAME)
    ws.Range(CODE_CELL).NumberFormat = managerSheet.TEXT_FORMAT
    ws.Range(CODE_CELL).value = managerTable.GetText(tbl, rowIndex, COL_CODE)

    managerSheet.SetEditId ws, editId
    ShowHeader ws

    managerSheet.SelectCell ws.Range(NAME_CELL)

End Sub


Private Sub ArchiveRow( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo
    Dim tbl As ListObject
    Dim itemId As Long
    Dim itemName As String

    Set tbl = managerTable.GetTable(ws, TABLE_NAME)
    itemId = managerTable.GetLong(tbl, rowIndex, COL_ID)
    itemName = managerTable.GetText(tbl, rowIndex, COL_NAME)

    If Not managerForm.Confirm("Hapus pos mobile " & itemName & "?" & vbCrLf & _
        "Request lama tetap tersimpan.") Then Exit Sub

    state = managerForm.BeginBusy("Menghapus pos mobile...")

    serviceMobileMaintenance.ArchiveMobileMaintenance itemId, itemName

    If managerSheet.GetEditId(ws) = itemId Then ClearFormCells ws
    RefreshTable ws

    managerForm.EndBusy state

    Application.StatusBar = "Pos mobile " & itemName & " dihapus."

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


Private Sub ClearFormCells( _
    ByVal ws As Worksheet)

    managerForm.ClearCells ws, NAME_CELL, CODE_CELL
    managerSheet.SetEditId ws, 0
    ShowHeader ws

End Sub


Private Sub ShowHeader( _
    ByVal ws As Worksheet)

    Dim editId As Long

    editId = managerSheet.GetEditId(ws)

    If editId > 0 Then
        ws.Range(HEADER_CELL).value = HEADER_EDIT & editId & ")"
    Else
        ws.Range(HEADER_CELL).value = HEADER_NEW
    End If

End Sub
