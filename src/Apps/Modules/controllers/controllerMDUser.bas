Attribute VB_Name = "controllerMDUser"
Option Explicit

' =========================================================
' CONTROLLER - MASTER DATA USER (sheet MD_USER)
'
'   Form: NAMA DEPAN K6, NAMA TENGAH K7, NAMA BELAKANG K8,
'         USERNAME PMI K9 (Windows login), JOBDESK K10, TIER K11
'   BATAL I12:J13, SIMPAN K12:L13, REFRESH I15:L16
'
'   A new user belongs to the application location
'   (SETTINGS > LOKASI). The username cannot be changed once
'   saved: it is stored in every document the user makes.
'   Tier drives approvals: the reviewer of an update request
'   needs a higher tier than the requester.
'
'   TableMDUser (ID, NAMA DEPAN, NAMA TENGAH, NAMA BELAKANG,
'   NAMA LENGKAP, USERNAME, JOBDESK, TIER, EDIT, HAPUS)
' =========================================================

Private Const MODULE_NAME As String = "controllerMDUser"

' =========================================================
' ERROR CODES (14800 - 14899)
' =========================================================
Public Const EmptyLocation As Long = 14801

Private Const HEADER_CELL As String = "I5"
Private Const FIRST_NAME_CELL As String = "K6"
Private Const MIDDLE_NAME_CELL As String = "K7"
Private Const LAST_NAME_CELL As String = "K8"
Private Const USERNAME_CELL As String = "K9"
Private Const JOBDESK_CELL As String = "K10"
Private Const TIER_CELL As String = "K11"

Private Const BUTTON_CANCEL As String = "I12:J13"
Private Const BUTTON_SAVE As String = "K12:L13"
Private Const BUTTON_REFRESH As String = "I15:L16"

Private Const TABLE_NAME As String = "TableMDUser"
Private Const COL_ID As String = "ID"
Private Const COL_FIRST_NAME As String = "NAMA DEPAN"
Private Const COL_MIDDLE_NAME As String = "NAMA TENGAH"
Private Const COL_LAST_NAME As String = "NAMA BELAKANG"
Private Const COL_USERNAME As String = "USERNAME"
Private Const COL_JOBDESK As String = "JOBDESK"
Private Const COL_TIER As String = "TIER"
Private Const COL_EDIT As String = "EDIT"
Private Const COL_DELETE As String = "HAPUS"

Private Const HEADER_NEW As String = "FORM PENGISIAN USER BARU"
Private Const HEADER_EDIT As String = "FORM EDIT USER (ID "


' =========================================================
' EVENTS
' =========================================================

Public Sub OnActivate( _
    ByVal ws As Worksheet)

    ShowHeader ws
    managerList.ApplyWholeNumberRule ws.Range(TIER_CELL), 0

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
        Refresh ws, True
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


' =========================================================
' ACTIONS
' =========================================================

Public Sub Save( _
    ByVal ws As Worksheet)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo
    Dim userItem As User
    Dim editId As Long
    Dim plantName As String

    editId = managerSheet.GetEditId(ws)

    Set userItem = New User

    With userItem
        If editId > 0 Then .Id = editId
        .FirstName = managerSheet.ReadText(ws, FIRST_NAME_CELL, "NAMA DEPAN")
        .MiddleName = managerSheet.ReadText(ws, MIDDLE_NAME_CELL, "NAMA TENGAH", False)
        .LastName = managerSheet.ReadText(ws, LAST_NAME_CELL, "NAMA BELAKANG", False)
        .Username = managerSheet.ReadText(ws, USERNAME_CELL, "USERNAME PMI")
        .Jobdesk = managerSheet.ReadText(ws, JOBDESK_CELL, "JOBDESK", False)
        .Tier = managerSheet.ReadWholeNumber(ws, TIER_CELL, "TIER", 0)
    End With

    If editId = 0 Then
        plantName = managerSheet.SettingsPlantName()
        If Len(plantName) = 0 Then
            managerError.ThrowError controllerMDUser.EmptyLocation, _
                MODULE_NAME & ".Save", _
                "Isi LOKASI di sheet SETTINGS terlebih dahulu: user baru " & _
                "terdaftar di lokasi tersebut."
        End If
    End If

    state = managerForm.BeginBusy("Menyimpan user...")

    If editId > 0 Then
        serviceUser.UpdateUser userItem
    Else
        serviceUser.CreateUser userItem, plantName
    End If

    RefreshTable ws
    ClearFormCells ws

    managerForm.EndBusy state

    managerForm.ShowInfo "User " & userItem.Username & " (" & userItem.FullName & ") tersimpan."

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
    ByVal ws As Worksheet, _
    Optional ByVal showResult As Boolean = False)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo
    Dim count As Long

    state = managerForm.BeginBusy("Memuat data user...")
    count = RefreshTable(ws)
    managerForm.EndBusy state

    If showResult Then Application.StatusBar = count & " user dimuat."

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


' Fill the table from the database (no error handling: the
' caller decides). Returns the number of users.
Public Function RefreshTable( _
    ByVal ws As Worksheet) As Long

    Dim users As Collection
    Dim userItem As User
    Dim data() As Variant
    Dim r As Long
    Dim tbl As ListObject

    Set tbl = managerTable.GetTable(ws, TABLE_NAME)
    Set users = serviceUser.ListUsers()

    If users.Count = 0 Then
        managerTable.SetRowCount tbl, 0
        Exit Function
    End If

    ReDim data(1 To users.Count, 1 To 10)

    For Each userItem In users
        r = r + 1
        data(r, 1) = userItem.Id
        data(r, 2) = userItem.FirstName
        data(r, 3) = userItem.MiddleName
        data(r, 4) = userItem.LastName
        data(r, 5) = userItem.FullName
        data(r, 6) = userItem.Username
        data(r, 7) = userItem.Jobdesk
        data(r, 8) = userItem.Tier
        data(r, 9) = managerSheet.LABEL_EDIT
        data(r, 10) = managerSheet.LABEL_DELETE
    Next userItem

    managerTable.WriteData tbl, data, users.Count, _
        Array("0", managerSheet.TEXT_FORMAT, managerSheet.TEXT_FORMAT, _
              managerSheet.TEXT_FORMAT, managerSheet.TEXT_FORMAT, managerSheet.TEXT_FORMAT, _
              managerSheet.TEXT_FORMAT, "0", managerSheet.TEXT_FORMAT, managerSheet.TEXT_FORMAT)

    RefreshTable = users.Count

End Function


' =========================================================
' INTERNAL
' =========================================================

Private Sub LoadForEdit( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long)

    Dim tbl As ListObject

    Set tbl = managerTable.GetTable(ws, TABLE_NAME)

    ws.Range(FIRST_NAME_CELL).value = managerTable.GetText(tbl, rowIndex, COL_FIRST_NAME)
    ws.Range(MIDDLE_NAME_CELL).value = managerTable.GetText(tbl, rowIndex, COL_MIDDLE_NAME)
    ws.Range(LAST_NAME_CELL).value = managerTable.GetText(tbl, rowIndex, COL_LAST_NAME)
    ws.Range(USERNAME_CELL).NumberFormat = managerSheet.TEXT_FORMAT
    ws.Range(USERNAME_CELL).value = managerTable.GetText(tbl, rowIndex, COL_USERNAME)
    ws.Range(JOBDESK_CELL).value = managerTable.GetText(tbl, rowIndex, COL_JOBDESK)
    ws.Range(TIER_CELL).value = managerTable.GetLong(tbl, rowIndex, COL_TIER)

    managerSheet.SetEditId ws, managerTable.GetLong(tbl, rowIndex, COL_ID)
    ShowHeader ws

    managerSheet.SelectCell ws.Range(FIRST_NAME_CELL)

End Sub


Private Sub ArchiveRow( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo
    Dim tbl As ListObject
    Dim userId As Long
    Dim username As String

    Set tbl = managerTable.GetTable(ws, TABLE_NAME)
    userId = managerTable.GetLong(tbl, rowIndex, COL_ID)
    username = managerTable.GetText(tbl, rowIndex, COL_USERNAME)

    If Not managerForm.Confirm("Hapus user " & username & "?" & vbCrLf & _
        "Dokumen yang dibuat user ini tetap tersimpan.") Then Exit Sub

    state = managerForm.BeginBusy("Menghapus user...")

    serviceUser.ArchiveUser userId

    If managerSheet.GetEditId(ws) = userId Then ClearFormCells ws
    RefreshTable ws

    managerForm.EndBusy state

    Application.StatusBar = "User " & username & " dihapus."

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


Private Sub ClearFormCells( _
    ByVal ws As Worksheet)

    managerForm.ClearCells ws, FIRST_NAME_CELL, MIDDLE_NAME_CELL, LAST_NAME_CELL, _
        USERNAME_CELL, JOBDESK_CELL, TIER_CELL
    managerSheet.SetEditId ws, 0
    ShowHeader ws

End Sub


Private Sub ShowHeader( _
    ByVal ws As Worksheet)

    Dim editId As Long

    editId = managerSheet.GetEditId(ws)

    If editId > 0 Then
        ws.Range(HEADER_CELL).value = HEADER_EDIT & editId & _
            ") - USERNAME TIDAK DAPAT DIUBAH"
    Else
        ws.Range(HEADER_CELL).value = HEADER_NEW
    End If

End Sub
