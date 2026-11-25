Attribute VB_Name = "controllerRekapOrder"
Option Explicit

' =========================================================
' CONTROLLER - REKAP ORDER (sheet REKAP_ORDER)
'
'   LOKASI K5, DARI K9 / HINGGA K10 (double-click = date)
'   CARI S9:U9, SEBELUMNYA S7, SELANJUTNYA U7,
'   FILTER KOLOM P7 / NILAI Q7, BARIS YANG DITAMPILKAN U10
'   Paging and filter are done by order_list / order_list_count
'   (one page is loaded at a time).
'
'   TableRekapOrder (NAMA SPAREPART and VENDOR show "code - name",
'   the same text as their dropdowns)
'     PO ID    : type a new value -> saved directly (no approval)
'     EDIT     : row becomes editable -> AJUKAN = update proposal
'     HAPUS    : delete proposal (quantity 0), needs approval
'     ESTIMASI KEDATANGAN = order date + lead time of the
'                           sparepart (REKAP INVENTORY)
' =========================================================

Private Const MODULE_NAME As String = "controllerRekapOrder"

' =========================================================
' ERROR CODES (15000 - 15099)
' =========================================================
Public Const OrderNotLoaded As Long = 15001

Private Const LAYOUT_KEY As String = "REKAP_ORDER"
Private Const TABLE_NAME As String = "TableRekapOrder"

Private Const COL_UUID As String = "UUID"
Private Const COL_PO As String = "PO ID"
Private Const COL_SPAREPART As String = "NAMA SPAREPART"
Private Const COL_QUANTITY As String = "JUMLAH ORDER"
Private Const COL_UNIT As String = "SATUAN"
Private Const COL_VENDOR As String = "VENDOR"
Private Const COL_DATE As String = "TANGGAL ORDER"
Private Const COL_ESTIMATE As String = "ESTIMASI KEDATANGAN"
Private Const COL_STATUS As String = "STATUS KEDATANGAN"
Private Const COL_REMARKS As String = "REMARKS"
Private Const COL_PIC As String = "PIC"
Private Const COL_EDIT As String = "EDIT"
Private Const COL_DELETE As String = "HAPUS"

Private Const PO_COLUMN_INDEX As Long = 2

Private mOrders As Collection
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

    If managerTable.IsInColumn(target.Cells(1, 1), tbl, COL_DATE) Then
        If controllerShared.IsEditRow(tbl, rowIndex, COL_EDIT) Then
            managerForm.HandleDatePickerDoubleClick target, target.Cells(1, 1), cancel
        End If
        Exit Sub
    End If

    If managerTable.IsInColumn(target.Cells(1, 1), tbl, COL_EDIT) Then

        cancel = True
        labelText = managerTable.GetText(tbl, rowIndex, COL_EDIT)

        If labelText = managerSheet.LABEL_EDIT Then
            controllerShared.EnterEditMode tbl, rowIndex, COL_EDIT, COL_DELETE, COL_SPAREPART
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

    If controllerShared.GuardRecapChange(tbl, target, COL_EDIT, _
        EditableColumns(), Array(COL_PO)) Then Exit Sub

    controllerShared.FillUnits tbl, target, COL_SPAREPART, COL_UNIT

    If target.Cells.Count = 1 Then
        If managerTable.IsInColumn(target, tbl, COL_PO) Then
            SavePoNumber ws, managerTable.GetRowIndex(tbl, target)
        End If
    End If

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
    Dim orderItem As Order
    Dim data() As Variant
    Dim r As Long
    Dim keyHolder As New PageRequest
    Dim firstKey As PageRequest
    Dim lastKey As PageRequest

    state = managerForm.BeginBusy("Mengambil data order...", True)

    totalRows = serviceOrder.CountOrders(mPlantName, mFromDate, mToDate, request)
    Set items = serviceOrder.ListOrders(mPlantName, mFromDate, mToDate, request)

    Set mOrders = New Collection

    If items.Count > 0 Then

        ReDim data(1 To items.Count, 1 To 13)

        For Each orderItem In items

            r = r + 1
            mOrders.Add orderItem, orderItem.UUID

            data(r, 1) = orderItem.UUID
            data(r, 2) = orderItem.PoNumber
            data(r, 3) = orderItem.Sparepart.DisplayName
            data(r, 4) = orderItem.Quantity
            data(r, 5) = orderItem.UnitOfMeasurement
            If orderItem.HasVendor Then data(r, 6) = orderItem.Vendor.DisplayName
            data(r, 7) = orderItem.OrderDate
            If orderItem.EstimatedArrival <> 0 Then data(r, 8) = orderItem.EstimatedArrival
            data(r, 9) = orderItem.Status
            data(r, 10) = orderItem.Remarks
            data(r, 11) = orderItem.PersonInCharge
            data(r, 12) = managerSheet.LABEL_EDIT
            data(r, 13) = managerSheet.LABEL_DELETE

            Set lastKey = keyHolder.KeyOf(orderItem.OrderDate, orderItem.Id)
            If r = 1 Then Set firstKey = lastKey

        Next orderItem

        managerPaging.StorePage GetLayout(), data, items.Count, totalRows, _
            pageNumber, request, firstKey, lastKey
    Else
        managerPaging.StorePage GetLayout(), Empty, 0, totalRows, _
            pageNumber, request, Nothing, Nothing
    End If

    RenderPage ws

    managerForm.EndBusy state

    If totalRows = 0 Then
        Application.StatusBar = "Tidak ada order pada periode ini."
    End If

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


' Save the PO ID typed in a row (asks first).
Public Sub SavePoNumber( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo
    Dim tbl As ListObject
    Dim orderItem As Order
    Dim poNumber As String

    Set tbl = managerTable.GetTable(ws, TABLE_NAME)
    Set orderItem = CachedOrder(managerTable.GetText(tbl, rowIndex, COL_UUID))

    poNumber = managerTable.GetText(tbl, rowIndex, COL_PO)

    If poNumber = orderItem.PoNumber Then Exit Sub

    If Not managerForm.Confirm("Simpan PO ID '" & poNumber & "' untuk order " & _
        orderItem.UUID & "?") Then

        RestorePo ws, tbl, rowIndex, orderItem.PoNumber
        Exit Sub
    End If

    state = managerForm.BeginBusy("Menyimpan PO ID...")

    serviceOrder.SetPoNumber orderItem.UUID, poNumber

    orderItem.PoNumber = poNumber
    UpdateCachedPo tbl, rowIndex, poNumber

    managerForm.EndBusy state

    Application.StatusBar = "PO ID order " & orderItem.UUID & " tersimpan."

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
    Dim original As Order
    Dim proposed As Order
    Dim reason As String
    Dim proposedUUID As String

    Set tbl = managerTable.GetTable(ws, TABLE_NAME)
    Set original = CachedOrder(managerTable.GetText(tbl, rowIndex, COL_UUID))

    state = managerForm.BeginBusy("Memeriksa perubahan...")

    Set proposed = New Order

    Set proposed.Sparepart = controllerShared.RequireSparepart( _
        serviceSparepart.ListSpareparts(), _
        managerSheet.TableRequiredText(tbl, rowIndex, COL_SPAREPART), _
        managerTable.GetCell(tbl, rowIndex, COL_SPAREPART))

    Set proposed.Vendor = controllerShared.FindVendor( _
        serviceVendor.ListVendors(), _
        managerTable.GetText(tbl, rowIndex, COL_VENDOR), _
        managerTable.GetCell(tbl, rowIndex, COL_VENDOR))

    proposed.OrderDate = managerSheet.TableDate(tbl, rowIndex, COL_DATE)
    proposed.Quantity = managerSheet.TableWholeNumber(tbl, rowIndex, COL_QUANTITY, 0)
    proposed.Remarks = managerTable.GetText(tbl, rowIndex, COL_REMARKS)
    proposed.PersonInCharge = managerSheet.TableRequiredText(tbl, rowIndex, COL_PIC)

    managerForm.EndBusy state

    If Not controllerShared.AskReason( _
        "Alasan perubahan order " & original.UUID & ":", reason) Then Exit Sub

    state = managerForm.BeginBusy("Menyimpan usulan update...", True)

    proposedUUID = serviceOrder.ProposeOrderUpdate(mPlantName, original, proposed, reason)

    RenderPage ws

    managerForm.EndBusy state

    managerForm.ShowInfo "Usulan update order " & original.UUID & _
        " tersimpan sebagai " & proposedUUID & "." & vbCrLf & vbCrLf & _
        "Order baru berlaku setelah disetujui di REQUEST UPDATE."

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


' HAPUS: propose to delete the order (quantity 0).
Public Sub ProposeDelete( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo
    Dim tbl As ListObject
    Dim original As Order
    Dim reason As String
    Dim message As String
    Dim proposedUUID As String

    Set tbl = managerTable.GetTable(ws, TABLE_NAME)
    Set original = CachedOrder(managerTable.GetText(tbl, rowIndex, COL_UUID))

    message = "Ajukan penghapusan order " & original.UUID & vbCrLf & _
        original.Sparepart.DisplayName & " x " & original.Quantity & " " & _
        original.UnitOfMeasurement & "?"

    If original.ReceivedQuantityTotal > 0 Then
        message = message & vbCrLf & vbCrLf & _
            "Order ini sudah memiliki penerimaan (" & original.ReceivedQuantityTotal & _
            "). Penerimaan tetap tercatat pada order yang dihapus."
    End If

    If Not managerForm.Confirm(message) Then Exit Sub

    If Not controllerShared.AskReason( _
        "Alasan penghapusan order " & original.UUID & ":", reason) Then Exit Sub

    state = managerForm.BeginBusy("Menyimpan usulan hapus...", True)

    proposedUUID = serviceOrder.ProposeOrderDelete(mPlantName, original, reason)

    managerForm.EndBusy state

    managerForm.ShowInfo "Usulan hapus order " & original.UUID & _
        " tersimpan (" & proposedUUID & ")." & vbCrLf & vbCrLf & _
        "Order terhapus setelah disetujui di REQUEST UPDATE."

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

    Set GetLayout = result

End Function


Private Function FilterHeaders() As Variant

    FilterHeaders = Array(COL_UUID, COL_PO, COL_SPAREPART, COL_QUANTITY, COL_UNIT, _
        COL_VENDOR, COL_DATE, COL_ESTIMATE, COL_STATUS, COL_REMARKS, COL_PIC)

End Function


Private Function EditableColumns() As Variant

    EditableColumns = Array(COL_SPAREPART, COL_QUANTITY, COL_VENDOR, _
        COL_DATE, COL_REMARKS, COL_PIC)

End Function


Private Function ColumnFormats() As Variant

    ColumnFormats = Array( _
        managerSheet.TEXT_FORMAT, managerSheet.TEXT_FORMAT, managerSheet.TEXT_FORMAT, _
        managerSheet.NUMBER_FORMAT, managerSheet.TEXT_FORMAT, managerSheet.TEXT_FORMAT, _
        managerSheet.DATE_FORMAT, managerSheet.DATE_FORMAT, managerSheet.TEXT_FORMAT, _
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


' Dropdowns for edit rows and the edit-row highlight.
Private Sub PrepareTable( _
    ByVal ws As Worksheet)

    Dim tbl As ListObject

    Set tbl = managerTable.GetTable(ws, TABLE_NAME)

    managerList.ApplyDropdown managerTable.GetColumnRange(tbl, COL_SPAREPART), _
        managerList.LIST_SPAREPART
    managerList.ApplyDropdown managerTable.GetColumnRange(tbl, COL_VENDOR), _
        managerList.LIST_VENDOR
    managerList.ApplyWholeNumberRule managerTable.GetColumnRange(tbl, COL_QUANTITY), 0

    controllerShared.EnsureEditHighlight tbl, COL_EDIT, EditableColumns()

End Sub





Private Function CachedOrder( _
    ByVal orderUUID As String) As Order

    If mOrders Is Nothing Then GoTo NotLoaded

    On Error GoTo NotLoaded
    Set CachedOrder = mOrders(orderUUID)
    Exit Function

NotLoaded:
    managerError.ThrowError controllerRekapOrder.OrderNotLoaded, _
        MODULE_NAME & ".CachedOrder", _
        "Data order belum dimuat atau sudah kedaluwarsa. Klik CARI lalu ulangi."

End Function


Private Sub RestorePo( _
    ByVal ws As Worksheet, _
    ByVal tbl As ListObject, _
    ByVal rowIndex As Long, _
    ByVal poNumber As String)

    Dim eventsWereOn As Boolean

    eventsWereOn = Application.EnableEvents
    Application.EnableEvents = False

    managerTable.SetValue tbl, rowIndex, COL_PO, poNumber

    Application.EnableEvents = eventsWereOn

End Sub


Private Sub UpdateCachedPo( _
    ByVal tbl As ListObject, _
    ByVal rowIndex As Long, _
    ByVal poNumber As String)

    Dim dataRow As Long
    Dim state As PageState

    dataRow = managerPaging.DataRowOf(LAYOUT_KEY, rowIndex)

    If dataRow > 0 Then
        Set state = managerPaging.GetState(LAYOUT_KEY)
        state.CellValue(dataRow, PO_COLUMN_INDEX) = poNumber
    End If

End Sub
