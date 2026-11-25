Attribute VB_Name = "controllerMDVendor"
Option Explicit

' =========================================================
' CONTROLLER - MASTER DATA VENDOR (sheet MD_VENDOR)
'
'   Form: NAMA VENDOR K6, NAMA PERUSAHAAN K7, EMAIL VENDOR K8
'   BATAL I9:J10, SIMPAN K9:L10, REFRESH I12:L13
'
'   TableMDVendor (ID, NAMA VENDOR, NAMA PERUSAHAAN,
'   EMAIL VENDOR, EDIT, HAPUS). Every refresh also rewrites
'   the vendor dropdown list ("code - name", LISTS sheet).
'   The vendor code is generated (V0001, V0002, ...).
'     EDIT  : loads the row into the form (SIMPAN updates it)
'     HAPUS : archives the vendor
' =========================================================

Private Const MODULE_NAME As String = "controllerMDVendor"

' =========================================================
' ERROR CODES (14900 - 14999)
' =========================================================
' (validation errors come from managerSheet / the Vendor model)

Private Const HEADER_CELL As String = "I5"
Private Const NAME_CELL As String = "K6"
Private Const COMPANY_CELL As String = "K7"
Private Const EMAIL_CELL As String = "K8"

Private Const BUTTON_CANCEL As String = "I9:J10"
Private Const BUTTON_SAVE As String = "K9:L10"
Private Const BUTTON_REFRESH As String = "I12:L13"

Private Const TABLE_NAME As String = "TableMDVendor"
Private Const COL_ID As String = "ID"
Private Const COL_NAME As String = "NAMA VENDOR"
Private Const COL_COMPANY As String = "NAMA PERUSAHAAN"
Private Const COL_EMAIL As String = "EMAIL VENDOR"
Private Const COL_EDIT As String = "EDIT"
Private Const COL_DELETE As String = "HAPUS"

Private Const HEADER_NEW As String = "FORM PENGISIAN VENDOR BARU"
Private Const HEADER_EDIT As String = "FORM EDIT VENDOR (ID "


' =========================================================
' EVENTS
' =========================================================

Public Sub OnActivate( _
    ByVal ws As Worksheet)

    ShowHeader ws

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
    Dim vendorItem As Vendor
    Dim editId As Long

    editId = managerSheet.GetEditId(ws)

    Set vendorItem = New Vendor

    With vendorItem
        If editId > 0 Then .Id = editId
        .Name = managerSheet.ReadText(ws, NAME_CELL, "NAMA VENDOR")
        .CompanyName = managerSheet.ReadText(ws, COMPANY_CELL, "NAMA PERUSAHAAN", False)
        .Email = managerSheet.ReadText(ws, EMAIL_CELL, "EMAIL VENDOR", False)
    End With

    state = managerForm.BeginBusy("Menyimpan vendor...")

    If editId > 0 Then
        serviceVendor.UpdateVendor vendorItem
    Else
        serviceVendor.CreateVendor vendorItem
    End If

    RefreshTable ws
    ClearFormCells ws

    managerForm.EndBusy state

    managerForm.ShowInfo "Vendor " & vendorItem.Name & " tersimpan."

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

    state = managerForm.BeginBusy("Memuat data vendor...")
    count = RefreshTable(ws)
    managerForm.EndBusy state

    If showResult Then Application.StatusBar = count & " vendor dimuat."

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


' Fill the table from the database (no error handling: the
' caller decides). Returns the number of vendors.
Public Function RefreshTable( _
    ByVal ws As Worksheet) As Long

    Dim vendors As Collection
    Dim vendorItem As Vendor
    Dim data() As Variant
    Dim r As Long
    Dim tbl As ListObject

    Set tbl = managerTable.GetTable(ws, TABLE_NAME)
    Set vendors = serviceVendor.ListVendors()

    ' Vendor dropdowns ("code - name") follow this table.
    managerList.WriteVendors vendors

    If vendors.Count = 0 Then
        managerTable.SetRowCount tbl, 0
        Exit Function
    End If

    ReDim data(1 To vendors.Count, 1 To 6)

    For Each vendorItem In vendors
        r = r + 1
        data(r, 1) = vendorItem.Id
        data(r, 2) = vendorItem.Name
        data(r, 3) = vendorItem.CompanyName
        data(r, 4) = vendorItem.Email
        data(r, 5) = managerSheet.LABEL_EDIT
        data(r, 6) = managerSheet.LABEL_DELETE
    Next vendorItem

    managerTable.WriteData tbl, data, vendors.Count, _
        Array("0", managerSheet.TEXT_FORMAT, managerSheet.TEXT_FORMAT, _
              managerSheet.TEXT_FORMAT, managerSheet.TEXT_FORMAT, managerSheet.TEXT_FORMAT)

    RefreshTable = vendors.Count

End Function


' =========================================================
' INTERNAL
' =========================================================

Private Sub LoadForEdit( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long)

    Dim tbl As ListObject

    Set tbl = managerTable.GetTable(ws, TABLE_NAME)

    ws.Range(NAME_CELL).value = managerTable.GetText(tbl, rowIndex, COL_NAME)
    ws.Range(COMPANY_CELL).value = managerTable.GetText(tbl, rowIndex, COL_COMPANY)
    ws.Range(EMAIL_CELL).value = managerTable.GetText(tbl, rowIndex, COL_EMAIL)

    managerSheet.SetEditId ws, managerTable.GetLong(tbl, rowIndex, COL_ID)
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
    Dim vendorId As Long
    Dim vendorName As String

    Set tbl = managerTable.GetTable(ws, TABLE_NAME)
    vendorId = managerTable.GetLong(tbl, rowIndex, COL_ID)
    vendorName = managerTable.GetText(tbl, rowIndex, COL_NAME)

    If Not managerForm.Confirm("Hapus vendor " & vendorName & "?" & vbCrLf & _
        "Order lama tetap tersimpan.") Then Exit Sub

    state = managerForm.BeginBusy("Menghapus vendor...")

    serviceVendor.ArchiveVendor vendorId

    If managerSheet.GetEditId(ws) = vendorId Then ClearFormCells ws
    RefreshTable ws

    managerForm.EndBusy state

    Application.StatusBar = "Vendor " & vendorName & " dihapus."

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


Private Sub ClearFormCells( _
    ByVal ws As Worksheet)

    managerForm.ClearCells ws, NAME_CELL, COMPANY_CELL, EMAIL_CELL
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
