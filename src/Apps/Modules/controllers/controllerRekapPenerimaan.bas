Attribute VB_Name = "controllerRekapPenerimaan"
Option Explicit

' =========================================================
' CONTROLLER - REKAP PENERIMAAN (sheet REKAP_PENERIMAAN)
'
'   LOKASI K5, DARI K9 / HINGGA K10 (arrival date, double-click
'   = date picker), CARI W5:X6, BARIS YANG DITAMPILKAN X7,
'   SEBELUMNYA S7, SELANJUTNYA U7, FILTER KOLOM P7 / NILAI Q7
'
'   TableRekapPenerimaan
'     EDIT  : row becomes editable -> AJUKAN = update proposal
'             (dates, quantities, remarks, PIC; the order stays)
'     HAPUS : delete proposal (all quantities 0)
'     JUMLAH YANG HARUS DIPENUHI = remaining quantity of the
'             order before this receive
' =========================================================

Private Const MODULE_NAME As String = "controllerRekapPenerimaan"

' =========================================================
' ERROR CODES (15200 - 15299)
' =========================================================
Public Const ReceiveNotLoaded As Long = 15201

Private Const LAYOUT_KEY As String = "REKAP_PENERIMAAN"
Private Const TABLE_NAME As String = "TableRekapPenerimaan"

Private Const COL_ORDER_UUID As String = "ORDER UUID"
Private Const COL_UUID As String = "PENERIMAAN UUID"
Private Const COL_PO As String = "PO ID"
Private Const COL_SPAREPART As String = "NAMA SPAREPART"
Private Const COL_ORDER_DATE As String = "TANGGAL ORDER"
Private Const COL_ARRIVAL As String = "TANGGAL DATANG"
Private Const COL_VERIFICATION As String = "TANGGAL VERIFIKASI"
Private Const COL_REMAINING As String = "JUMLAH YANG HARUS DIPENUHI"
Private Const COL_RECEIVED As String = "JUMLAH PENERIMAAN"
Private Const COL_RECEIVED_TYPO As String = "JUMLAH PENERIMAAN2"
Private Const COL_OK As String = "JUMLAH OK"
Private Const COL_NOT_OK As String = "JUMLAH NOT OK"
Private Const COL_UNIT As String = "SATUAN"
Private Const COL_REMARKS As String = "REMARKS"
Private Const COL_PIC As String = "PIC"
Private Const COL_EDIT As String = "EDIT"
Private Const COL_DELETE As String = "HAPUS"

Private mReceives As Collection
Private mPlantName As String
Private mFromDate As Date
Private mToDate As Date


' =========================================================
' EVENTS
' =========================================================

Public Sub OnActivate( _
    ByVal ws As Worksheet)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo

    state = managerForm.BeginBusy()

    FixHeaders ws
    managerPaging.PrepareSheet ws, GetLayout(), FilterHeaders()
    PrepareTable ws

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

    Dim layout As PagingLayout
    Dim tbl As ListObject
    Dim rowIndex As Long
    Dim delta As Long
    Dim labelText As String

    Set layout = GetLayout()

    If managerPaging.IsDateCell(ws, layout, target) Then
        managerForm.HandleDatePickerDoubleClick target, target.Cells(1, 1), cancel
        Exit Sub
    End If

    delta = managerPaging.PageButtonDelta(ws, layout, target)

    If delta <> 0 Then
        cancel = True
        MovePage ws, delta
        Exit Sub
    End If

    If managerSheet.IsHit(target, ws, layout.SearchButton) Then
        cancel = True
        Search ws
        Exit Sub
    End If

    Set tbl = managerTable.GetTable(ws, TABLE_NAME)
    rowIndex = managerTable.GetRowIndex(tbl, target.Cells(1, 1))
    If rowIndex = 0 Then Exit Sub

    If managerTable.IsInColumn(target.Cells(1, 1), tbl, COL_ARRIVAL) _
        Or managerTable.IsInColumn(target.Cells(1, 1), tbl, COL_VERIFICATION) Then

        If controllerShared.IsEditRow(tbl, rowIndex, COL_EDIT) Then
            managerForm.HandleDatePickerDoubleClick target, target.Cells(1, 1), cancel
        End If
        Exit Sub
    End If

    If managerTable.IsInColumn(target.Cells(1, 1), tbl, COL_EDIT) Then

        cancel = True
        labelText = managerTable.GetText(tbl, rowIndex, COL_EDIT)

        If labelText = managerSheet.LABEL_EDIT Then
            controllerShared.EnterEditMode tbl, rowIndex, COL_EDIT, COL_DELETE, COL_RECEIVED
        ElseIf labelText = managerSheet.LABEL_SUBMIT Then
            SubmitUpdate ws, rowIndex
        End If

    ElseIf managerTable.IsInColumn(target.Cells(1, 1), tbl, COL_DELETE) Then

        cancel = True
        labelText = managerTable.GetText(tbl, rowIndex, COL_DELETE)

        If labelText = managerSheet.LABEL_DELETE Then
            ProposeDelete ws, rowIndex
        ElseIf labelText = managerSheet.LABEL_CANCEL Then
            RenderPage ws
        End If

    End If

End Sub


Public Sub OnChange( _
    ByVal ws As Worksheet, _
    ByVal target As Range)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo
    Dim layout As PagingLayout
    Dim tbl As ListObject

    Set layout = GetLayout()

    If managerPaging.IsViewCell(ws, layout, target) Then
        If managerPaging.HasState(LAYOUT_KEY) Then Search ws
        Exit Sub
    End If

    Set tbl = managerTable.GetTable(ws, TABLE_NAME)

    If Not managerTable.IsInTable(target, tbl) Then Exit Sub

    controllerShared.GuardRecapChange tbl, target, COL_EDIT, EditableColumns()

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerError.ShowError errInfo

End Sub


' =========================================================
' ACTIONS
' =========================================================

' CARI: remember the criteria and load page 1 (count + rows,
' paged by the query).
Public Sub Search( _
    ByVal ws As Worksheet)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo
    Dim layout As PagingLayout

    Set layout = GetLayout()

    managerPaging.ReadCriteria ws, layout, mPlantName, mFromDate, mToDate

    LoadPage ws, managerPaging.FirstPageRequest(ws, layout, FilterHeaders()), 1

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerError.ShowError errInfo

End Sub

' SEBELUMNYA / SELANJUTNYA: load the neighbouring page.
Private Sub MovePage( _
    ByVal ws As Worksheet, _
    ByVal delta As Long)

    Dim request As PageRequest
    Dim newPage As Long

    Set request = managerPaging.MoveRequest(GetLayout(), delta, newPage)

    If Not request Is Nothing Then LoadPage ws, request, newPage

End Sub

' Load one page of the search and show it.
Private Sub LoadPage( _
    ByVal ws As Worksheet, _
    ByVal request As PageRequest, _
    ByVal pageNumber As Long)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo
    Dim totalRows As Long
    Dim items As Collection
    Dim receiveItem As Receive
    Dim data() As Variant
    Dim r As Long
    Dim keyHolder As New PageRequest
    Dim firstKey As PageRequest
    Dim lastKey As PageRequest

    state = managerForm.BeginBusy("Mengambil data penerimaan...", True)

    totalRows = serviceReceive.CountReceives(mPlantName, mFromDate, mToDate, request)
    Set items = serviceReceive.ListReceives(mPlantName, mFromDate, mToDate, request)

    FixHeaders ws
    Set mReceives = New Collection

    If items.Count > 0 Then

        ReDim data(1 To items.Count, 1 To 16)

        For Each receiveItem In items

            r = r + 1
            mReceives.Add receiveItem, receiveItem.UUID

            data(r, 1) = receiveItem.Order.UUID
            data(r, 2) = receiveItem.UUID
            data(r, 3) = receiveItem.Order.PoNumber
            data(r, 4) = receiveItem.Order.Sparepart.DisplayName
            data(r, 5) = receiveItem.Order.OrderDate
            data(r, 6) = receiveItem.ArrivalDate
            If receiveItem.IsVerified Then data(r, 7) = receiveItem.VerificationDate
            data(r, 8) = receiveItem.RemainingQuantityBefore
            data(r, 9) = receiveItem.ReceivedQuantity
            data(r, 10) = receiveItem.OkQuantity
            data(r, 11) = receiveItem.NotOkQuantity
            data(r, 12) = receiveItem.UnitOfMeasurement
            data(r, 13) = receiveItem.Remarks
            data(r, 14) = receiveItem.PersonInCharge
            data(r, 15) = managerSheet.LABEL_EDIT
            data(r, 16) = managerSheet.LABEL_DELETE

            Set lastKey = keyHolder.KeyOf(receiveItem.ArrivalDate, receiveItem.Id)
            If r = 1 Then Set firstKey = lastKey

        Next receiveItem

        managerPaging.StorePage GetLayout(), data, items.Count, totalRows, _
            pageNumber, request, firstKey, lastKey
    Else
        managerPaging.StorePage GetLayout(), Empty, 0, totalRows, _
            pageNumber, request, Nothing, Nothing
    End If

    RenderPage ws

    managerForm.EndBusy state

    If totalRows = 0 Then
        Application.StatusBar = "Tidak ada penerimaan pada periode ini."
    End If

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


' AJUKAN: save the edited row as an update proposal.
Public Sub SubmitUpdate( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo
    Dim tbl As ListObject
    Dim original As Receive
    Dim proposed As Receive
    Dim reason As String
    Dim proposedUUID As String

    Set tbl = managerTable.GetTable(ws, TABLE_NAME)
    Set original = CachedReceive(managerTable.GetText(tbl, rowIndex, COL_UUID))

    Set proposed = New Receive

    With proposed
        .ArrivalDate = managerSheet.TableDate(tbl, rowIndex, COL_ARRIVAL)
        .VerificationDate = managerSheet.TableDate(tbl, rowIndex, COL_VERIFICATION, False)
        .ReceivedQuantity = managerSheet.TableWholeNumber(tbl, rowIndex, COL_RECEIVED, 0)
        .OkQuantity = managerSheet.TableWholeNumber(tbl, rowIndex, COL_OK, 0)
        .NotOkQuantity = managerSheet.TableWholeNumber(tbl, rowIndex, COL_NOT_OK, 0)
        .Remarks = managerTable.GetText(tbl, rowIndex, COL_REMARKS)
        .PersonInCharge = managerSheet.TableRequiredText(tbl, rowIndex, COL_PIC)
    End With

    If Not controllerShared.AskReason( _
        "Alasan perubahan penerimaan " & original.UUID & ":", reason) Then Exit Sub

    state = managerForm.BeginBusy("Menyimpan usulan update...", True)

    proposedUUID = serviceReceive.ProposeReceiveUpdate(mPlantName, original, proposed, reason)

    RenderPage ws

    managerForm.EndBusy state

    managerForm.ShowInfo "Usulan update penerimaan " & original.UUID & _
        " tersimpan sebagai " & proposedUUID & "." & vbCrLf & vbCrLf & _
        "Perubahan berlaku (dan stok dihitung ulang) setelah disetujui di REQUEST UPDATE."

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


' HAPUS: propose to delete the receive.
Public Sub ProposeDelete( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo
    Dim tbl As ListObject
    Dim original As Receive
    Dim reason As String
    Dim proposedUUID As String

    Set tbl = managerTable.GetTable(ws, TABLE_NAME)
    Set original = CachedReceive(managerTable.GetText(tbl, rowIndex, COL_UUID))

    If Not managerForm.Confirm("Ajukan penghapusan penerimaan " & original.UUID & vbCrLf & _
        "(order " & original.Order.UUID & ", OK " & original.OkQuantity & " " & _
        original.UnitOfMeasurement & ")?") Then Exit Sub

    If Not controllerShared.AskReason( _
        "Alasan penghapusan penerimaan " & original.UUID & ":", reason) Then Exit Sub

    state = managerForm.BeginBusy("Menyimpan usulan hapus...", True)

    proposedUUID = serviceReceive.ProposeReceiveDelete(mPlantName, original, reason)

    managerForm.EndBusy state

    managerForm.ShowInfo "Usulan hapus penerimaan " & original.UUID & _
        " tersimpan (" & proposedUUID & ")." & vbCrLf & vbCrLf & _
        "Penerimaan terhapus setelah disetujui di REQUEST UPDATE."

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


' =========================================================
' INTERNAL
' =========================================================

Private Function GetLayout() As PagingLayout

    Dim result As PagingLayout

    Set result = New PagingLayout
    result.UseStandardRecap LAYOUT_KEY, TABLE_NAME

    ' This sheet has CARI and BARIS YANG DITAMPILKAN on the right.
    result.SearchButton = "W5:X6"
    result.PageSizeCell = "X7"

    Set GetLayout = result

End Function


' The design has the header "JUMLAH PENERIMAAN2" (typo).
Private Sub FixHeaders( _
    ByVal ws As Worksheet)

    Dim tbl As ListObject

    Set tbl = managerTable.GetTable(ws, TABLE_NAME)

    If managerTable.HasColumn(tbl, COL_RECEIVED_TYPO) _
        And Not managerTable.HasColumn(tbl, COL_RECEIVED) Then

        tbl.ListColumns(COL_RECEIVED_TYPO).Name = COL_RECEIVED
    End If

End Sub


Private Function FilterHeaders() As Variant

    FilterHeaders = Array(COL_ORDER_UUID, COL_UUID, COL_PO, COL_SPAREPART, _
        COL_ORDER_DATE, COL_ARRIVAL, COL_VERIFICATION, COL_REMAINING, COL_RECEIVED, _
        COL_OK, COL_NOT_OK, COL_UNIT, COL_REMARKS, COL_PIC)

End Function


Private Function EditableColumns() As Variant

    EditableColumns = Array(COL_ARRIVAL, COL_VERIFICATION, COL_RECEIVED, _
        COL_OK, COL_NOT_OK, COL_REMARKS, COL_PIC)

End Function


Private Function ColumnFormats() As Variant

    ColumnFormats = Array( _
        managerSheet.TEXT_FORMAT, managerSheet.TEXT_FORMAT, managerSheet.TEXT_FORMAT, _
        managerSheet.TEXT_FORMAT, managerSheet.DATE_FORMAT, managerSheet.DATE_FORMAT, _
        managerSheet.DATE_FORMAT, managerSheet.NUMBER_FORMAT, managerSheet.NUMBER_FORMAT, _
        managerSheet.NUMBER_FORMAT, managerSheet.NUMBER_FORMAT, managerSheet.TEXT_FORMAT, _
        managerSheet.TEXT_FORMAT, managerSheet.TEXT_FORMAT, managerSheet.TEXT_FORMAT, _
        managerSheet.TEXT_FORMAT)

End Function


Private Sub RenderPage( _
    ByVal ws As Worksheet)

    Dim eventsWereOn As Boolean

    eventsWereOn = Application.EnableEvents
    Application.EnableEvents = False

    managerPaging.Render ws, GetLayout(), ColumnFormats()
    PrepareTable ws

    Application.EnableEvents = eventsWereOn

End Sub


Private Sub PrepareTable( _
    ByVal ws As Worksheet)

    Dim tbl As ListObject

    Set tbl = managerTable.GetTable(ws, TABLE_NAME)

    managerList.ApplyWholeNumberRule managerTable.GetColumnRange(tbl, COL_RECEIVED), 0
    managerList.ApplyWholeNumberRule managerTable.GetColumnRange(tbl, COL_OK), 0
    managerList.ApplyWholeNumberRule managerTable.GetColumnRange(tbl, COL_NOT_OK), 0

    controllerShared.EnsureEditHighlight tbl, COL_EDIT, EditableColumns()

End Sub


Private Function CachedReceive( _
    ByVal receiveUUID As String) As Receive

    If mReceives Is Nothing Then GoTo NotLoaded

    On Error GoTo NotLoaded
    Set CachedReceive = mReceives(receiveUUID)
    Exit Function

NotLoaded:
    managerError.ThrowError controllerRekapPenerimaan.ReceiveNotLoaded, _
        MODULE_NAME & ".CachedReceive", _
        "Data penerimaan belum dimuat atau sudah kedaluwarsa. Klik CARI lalu ulangi."

End Function
