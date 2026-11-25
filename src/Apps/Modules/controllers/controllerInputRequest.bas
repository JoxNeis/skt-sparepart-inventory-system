Attribute VB_Name = "controllerInputRequest"
Option Explicit

' =========================================================
' CONTROLLER - INPUT REQUEST / PENGELUARAN (sheet INPUT_REQUEST)
'
'   LOKASI           K5   dropdown (plant)
'   TANGGAL          K7   double-click = date picker
'   BERSIHKAN TABEL  N5:O5   double-click
'   TAMBAH BARIS     N7:O7   double-click
'   SIMPAN           N9:O9   double-click
'   CEK SPAREPART    R6 (dropdown, free text) + CARI T6:U6
'
'   TableInputRequest: NAMA SPAREPART (dropdown), JUMLAH ORDER
'   (quantity issued), SATUAN (automatic), POS MOBILE MTC
'   (dropdown), REMARKS, PIC, HAPUS (double-click)
'
' The POS MOBILE MTC dropdown lists MD_POS_MOBILE, so that
' sheet is switched to the plant of LOKASI.
' =========================================================

Private Const MODULE_NAME As String = "controllerInputRequest"

' =========================================================
' ERROR CODES (14300 - 14399)
' =========================================================
Public Const NoRequestRow As Long = 14301

Private Const PLANT_CELL As String = "K5"
Private Const DATE_CELL As String = "K7"

Private Const BUTTON_CLEAR As String = "N5:O5"
Private Const BUTTON_ADD_ROW As String = "N7:O7"
Private Const BUTTON_SAVE As String = "N9:O9"

Private Const SEARCH_CELL As String = "R6"
Private Const BUTTON_SEARCH As String = "T6:U6"
Private Const RESULT_FIRST_CELL As String = "Q9"

Private Const TABLE_NAME As String = "TableInputRequest"
Private Const COL_SPAREPART As String = "NAMA SPAREPART"
Private Const COL_QUANTITY As String = "JUMLAH ORDER"
Private Const COL_UNIT As String = "SATUAN"
Private Const COL_MOBILE As String = "POS MOBILE MTC"
Private Const COL_REMARKS As String = "REMARKS"
Private Const COL_PIC As String = "PIC"
Private Const COL_DELETE As String = "HAPUS"


' =========================================================
' EVENTS
' =========================================================

Public Sub OnActivate( _
    ByVal ws As Worksheet)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo

    state = managerForm.BeginBusy()

    managerSheet.PreparePlantCell ws, PLANT_CELL
    managerSheet.PrepareDateCell ws, DATE_CELL, Date
    managerList.ApplyDropdown ws.Range(SEARCH_CELL), managerList.LIST_SPAREPART, False
    PrepareTable GetTable(ws)

    controllerMDPosMobile.EnsurePlant managerForm.GetCellText(ws, PLANT_CELL)

    managerForm.EndBusy state
    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


Public Sub OnDoubleClick( _
    ByVal ws As Worksheet, _
    ByVal target As Range, _
    ByRef cancel As Boolean)

    Dim tbl As ListObject

    If managerForm.HandleDatePickerDoubleClick(target, ws.Range(DATE_CELL), cancel) Then
        Exit Sub
    End If

    If managerSheet.IsHit(target, ws, BUTTON_SAVE) Then
        cancel = True
        SaveRequests ws
    ElseIf managerSheet.IsHit(target, ws, BUTTON_ADD_ROW) Then
        cancel = True
        AddRow ws
    ElseIf managerSheet.IsHit(target, ws, BUTTON_CLEAR) Then
        cancel = True
        ClearTable ws, True
    ElseIf managerSheet.IsHit(target, ws, BUTTON_SEARCH) Then
        cancel = True
        controllerShared.SearchSparepart ws, PLANT_CELL, SEARCH_CELL, RESULT_FIRST_CELL
    Else
        Set tbl = GetTable(ws)
        If managerTable.IsInColumn(target.Cells(1, 1), tbl, COL_DELETE) Then
            cancel = True
            DeleteRow ws, managerTable.GetRowIndex(tbl, target.Cells(1, 1))
        End If
    End If

End Sub


Public Sub OnChange( _
    ByVal ws As Worksheet, _
    ByVal target As Range)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    If managerSheet.IsHit(target, ws, PLANT_CELL) Then
        controllerMDPosMobile.EnsurePlant managerForm.GetCellText(ws, PLANT_CELL)
        Exit Sub
    End If

    controllerShared.FillUnits GetTable(ws), target, COL_SPAREPART, COL_UNIT

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerError.ShowError errInfo

End Sub


' =========================================================
' ACTIONS
' =========================================================

Public Sub AddRow( _
    ByVal ws As Worksheet)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo
    Dim tbl As ListObject
    Dim newRow As Long
    Dim previousPic As String
    Dim previousMobile As String

    state = managerForm.BeginBusy()

    Set tbl = GetTable(ws)

    previousPic = managerTable.GetText(tbl, managerTable.RowCount(tbl), COL_PIC)
    previousMobile = managerTable.GetText(tbl, managerTable.RowCount(tbl), COL_MOBILE)

    newRow = managerTable.AddStyledRow(tbl)

    managerTable.SetValue tbl, newRow, COL_DELETE, managerSheet.LABEL_DELETE
    managerTable.SetValue tbl, newRow, COL_PIC, previousPic
    managerTable.SetValue tbl, newRow, COL_MOBILE, previousMobile

    PrepareTable tbl

    managerForm.EndBusy state

    managerTable.SelectCell tbl, newRow, COL_SPAREPART

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


Public Sub ClearTable( _
    ByVal ws As Worksheet, _
    ByVal askFirst As Boolean)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo

    If askFirst Then
        If Not managerForm.Confirm("Hapus semua baris pada tabel request?") Then Exit Sub
    End If

    state = managerForm.BeginBusy()

    managerTable.SetRowCount GetTable(ws), 1
    ws.Range(RESULT_FIRST_CELL).Resize(1, 5).ClearContents
    PrepareTable GetTable(ws)

    managerForm.EndBusy state
    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


Public Sub DeleteRow( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo
    Dim tbl As ListObject

    Set tbl = GetTable(ws)
    If rowIndex = 0 Then Exit Sub

    If controllerShared.RowHasInput(tbl, rowIndex, InputColumns()) Then
        If Not managerForm.Confirm("Hapus baris ke-" & rowIndex & "?") Then Exit Sub
    End If

    state = managerForm.BeginBusy()
    controllerShared.DeleteInputRow tbl, rowIndex, COL_DELETE, managerSheet.LABEL_DELETE
    managerForm.EndBusy state

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


Public Sub SaveRequests( _
    ByVal ws As Worksheet)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo
    Dim tbl As ListObject
    Dim plantName As String
    Dim requestDate As Date
    Dim spareparts As Collection
    Dim mobiles As Collection
    Dim requests As Collection
    Dim requestItem As Request
    Dim uuids As Collection
    Dim rowIndex As Long
    Dim message As String
    Dim uuid As Variant

    Set tbl = GetTable(ws)

    plantName = managerSheet.ReadText(ws, PLANT_CELL, "LOKASI")
    requestDate = managerSheet.ReadDate(ws, DATE_CELL, "TANGGAL REQUEST")

    state = managerForm.BeginBusy("Memeriksa request...")

    Set spareparts = serviceSparepart.ListSpareparts()
    Set mobiles = serviceMobileMaintenance.ListByPlant(plantName)
    Set requests = New Collection

    For rowIndex = 1 To managerTable.RowCount(tbl)

        If controllerShared.RowHasInput(tbl, rowIndex, InputColumns()) Then

            Set requestItem = New Request

            Set requestItem.Sparepart = controllerShared.RequireSparepart( _
                spareparts, _
                managerSheet.TableRequiredText(tbl, rowIndex, COL_SPAREPART), _
                managerTable.GetCell(tbl, rowIndex, COL_SPAREPART))

            Set requestItem.MobileMaintenance = controllerShared.RequireMobileMaintenance( _
                mobiles, _
                managerSheet.TableRequiredText(tbl, rowIndex, COL_MOBILE), _
                plantName, _
                managerTable.GetCell(tbl, rowIndex, COL_MOBILE))

            requestItem.RequestDate = requestDate
            requestItem.Quantity = managerSheet.TableWholeNumber(tbl, rowIndex, COL_QUANTITY, 1)
            requestItem.Remarks = managerTable.GetText(tbl, rowIndex, COL_REMARKS)
            requestItem.PersonInCharge = managerSheet.TableRequiredText(tbl, rowIndex, COL_PIC)

            requests.Add requestItem

        End If

    Next rowIndex

    managerForm.EndBusy state

    If requests.Count = 0 Then
        managerError.ThrowError controllerInputRequest.NoRequestRow, _
            MODULE_NAME & ".SaveRequests", _
            "Isi minimal satu baris request."
    End If

    If Not managerForm.Confirm( _
        "Simpan " & requests.Count & " request (pengeluaran) untuk " & plantName & _
        " tanggal " & Format$(requestDate, managerSheet.DATE_FORMAT) & "?") Then Exit Sub

    state = managerForm.BeginBusy("Menyimpan request...", True)

    Set uuids = serviceRequest.CreateRequests(plantName, requests)

    managerTable.SetRowCount tbl, 1
    PrepareTable tbl
    ws.Range(RESULT_FIRST_CELL).Resize(1, 5).ClearContents

    managerForm.EndBusy state

    message = requests.Count & " request tersimpan:" & vbCrLf
    For Each uuid In uuids
        message = message & "  " & CStr(uuid) & vbCrLf
    Next uuid

    managerForm.ShowInfo message

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


' =========================================================
' INTERNAL
' =========================================================

Private Function GetTable( _
    ByVal ws As Worksheet) As ListObject

    Set GetTable = managerTable.GetTable(ws, TABLE_NAME)

End Function


' Columns that make a row count as a request. PIC and POS
' MOBILE are left out: TAMBAH BARIS copies them from the row
' above.
Private Function InputColumns() As Variant

    InputColumns = Array(COL_SPAREPART, COL_QUANTITY, COL_REMARKS)

End Function


Private Sub PrepareTable( _
    ByVal tbl As ListObject)

    managerSheet.FillColumn tbl, COL_DELETE, managerSheet.LABEL_DELETE

    managerList.ApplyDropdown managerTable.GetColumnRange(tbl, COL_SPAREPART), _
        managerList.LIST_SPAREPART
    managerList.ApplyDropdown managerTable.GetColumnRange(tbl, COL_MOBILE), _
        managerList.LIST_MOBILE_POST
    managerList.ApplyWholeNumberRule managerTable.GetColumnRange(tbl, COL_QUANTITY), 1

End Sub
