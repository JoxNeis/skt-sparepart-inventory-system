Attribute VB_Name = "controllerRekapPendingOrder"
Option Explicit

' =========================================================
' CONTROLLER - REKAP PENDING ORDER (sheet REKAP_PENDING_ORDER)
'
' Orders with a remaining quantity (not everything OK yet).
'
'   LOKASI K5, DARI K9 / HINGGA K10 (double-click = date)
'   CARI S9:U9, SEBELUMNYA S7, SELANJUTNYA U7,
'   FILTER KOLOM P7 / NILAI Q7, BARIS YANG DITAMPILKAN U10
'
'   TableRekapPendingOrder
'     PO ID        : the only editable column
'     SIMPAN PO ID : double-click saves the PO ID of the row
'     JUMLAH SUDAH DATANG = OK quantity received so far
'     JUMLAH BELUM DATANG = ordered - OK
' =========================================================

Private Const MODULE_NAME As String = "controllerRekapPendingOrder"

' =========================================================
' ERROR CODES (15100 - 15199)
' =========================================================
Public Const OrderNotLoaded As Long = 15101

Private Const LAYOUT_KEY As String = "REKAP_PENDING_ORDER"
Private Const TABLE_NAME As String = "TableRekapPendingOrder"

Private Const COL_UUID As String = "UUID"
Private Const COL_PO As String = "PO ID"
Private Const COL_SPAREPART As String = "NAMA SPAREPART"
Private Const COL_DATE As String = "TANGGAL ORDER"
Private Const COL_QUANTITY As String = "JUMLAH ORDER"
Private Const COL_RECEIVED As String = "JUMLAH SUDAH DATANG"
Private Const COL_REMAINING As String = "JUMLAH BELUM DATANG"
Private Const COL_UNIT As String = "SATUAN"
Private Const COL_VENDOR As String = "VENDOR"
Private Const COL_REMARKS As String = "REMARKS"
Private Const COL_PIC As String = "PIC"
Private Const COL_SAVE_PO As String = "SIMPAN PO ID"

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

    If managerTable.IsInColumn(target.Cells(1, 1), tbl, COL_SAVE_PO) Then
        cancel = True
        rowIndex = managerTable.GetRowIndex(tbl, target.Cells(1, 1))
        If rowIndex > 0 Then SavePoNumber ws, rowIndex
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

    managerSheet.UndoIfReadOnly target, tbl.DataBodyRange, _
        managerTable.GetColumnRange(tbl, COL_PO), _
        "Hanya PO ID yang bisa diubah di sini (lalu double-click SIMPAN PO)."

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

    LoadPage ws, managerPaging.FirstPageRequest(ws, layout, FilterHeaders(), FilterIndexes()), 1

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

    state = managerForm.BeginBusy("Mengambil data pending order...", True)

    totalRows = serviceOrder.CountPendingOrders(mPlantName, mFromDate, mToDate, request)
    Set items = serviceOrder.ListPendingOrders(mPlantName, mFromDate, mToDate, request)

    Set mOrders = New Collection

    If items.Count > 0 Then

        ReDim data(1 To items.Count, 1 To 12)

        For Each orderItem In items

            r = r + 1
            mOrders.Add orderItem, orderItem.UUID

            data(r, 1) = orderItem.UUID
            data(r, 2) = orderItem.PoNumber
            data(r, 3) = orderItem.Sparepart.DisplayName
            data(r, 4) = orderItem.OrderDate
            data(r, 5) = orderItem.Quantity
            data(r, 6) = orderItem.OkQuantityTotal
            data(r, 7) = orderItem.RemainingQuantity
            data(r, 8) = orderItem.UnitOfMeasurement
            If orderItem.HasVendor Then data(r, 9) = orderItem.Vendor.DisplayName
            data(r, 10) = orderItem.Remarks
            data(r, 11) = orderItem.PersonInCharge
            data(r, 12) = managerSheet.LABEL_SAVE_PO

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
        Application.StatusBar = "Tidak ada pending order pada periode ini."
    End If

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


Public Sub SavePoNumber( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo
    Dim tbl As ListObject
    Dim orderItem As Order
    Dim poNumber As String
    Dim dataRow As Long
    Dim pagingState As PageState

    Set tbl = managerTable.GetTable(ws, TABLE_NAME)
    Set orderItem = CachedOrder(managerTable.GetText(tbl, rowIndex, COL_UUID))

    poNumber = managerTable.GetText(tbl, rowIndex, COL_PO)

    If poNumber = orderItem.PoNumber Then
        Application.StatusBar = "PO ID tidak berubah."
        Exit Sub
    End If

    If Not managerForm.Confirm("Simpan PO ID '" & poNumber & "' untuk order " & _
        orderItem.UUID & "?") Then Exit Sub

    state = managerForm.BeginBusy("Menyimpan PO ID...")

    serviceOrder.SetPoNumber orderItem.UUID, poNumber

    orderItem.PoNumber = poNumber

    dataRow = managerPaging.DataRowOf(LAYOUT_KEY, rowIndex)
    If dataRow > 0 Then
        Set pagingState = managerPaging.GetState(LAYOUT_KEY)
        pagingState.CellValue(dataRow, PO_COLUMN_INDEX) = poNumber
    End If

    managerForm.EndBusy state

    Application.StatusBar = "PO ID order " & orderItem.UUID & " tersimpan."

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

    FilterHeaders = Array(COL_UUID, COL_PO, COL_SPAREPART, COL_DATE, COL_QUANTITY, _
        COL_RECEIVED, COL_REMAINING, COL_UNIT, COL_VENDOR, COL_REMARKS, COL_PIC)

End Function


' order_list_pending column number of each KOLOM choice.
Private Function FilterIndexes() As Variant

    FilterIndexes = Array(1, 2, 3, 7, 4, 12, 13, 5, 6, 10, 11)

End Function


Private Function ColumnFormats() As Variant

    ColumnFormats = Array( _
        managerSheet.TEXT_FORMAT, managerSheet.TEXT_FORMAT, managerSheet.TEXT_FORMAT, _
        managerSheet.DATE_FORMAT, managerSheet.NUMBER_FORMAT, managerSheet.NUMBER_FORMAT, _
        managerSheet.NUMBER_FORMAT, managerSheet.TEXT_FORMAT, managerSheet.TEXT_FORMAT, _
        managerSheet.TEXT_FORMAT, managerSheet.TEXT_FORMAT, managerSheet.TEXT_FORMAT)

End Function


Private Sub RenderPage( _
    ByVal ws As Worksheet)

    Dim eventsWereOn As Boolean

    eventsWereOn = Application.EnableEvents
    Application.EnableEvents = False

    managerPaging.Render ws, GetLayout(), ColumnFormats()

    Application.EnableEvents = eventsWereOn

End Sub


Private Function CachedOrder( _
    ByVal orderUUID As String) As Order

    If mOrders Is Nothing Then GoTo NotLoaded

    On Error GoTo NotLoaded
    Set CachedOrder = mOrders(orderUUID)
    Exit Function

NotLoaded:
    managerError.ThrowError controllerRekapPendingOrder.OrderNotLoaded, _
        MODULE_NAME & ".CachedOrder", _
        "Data order belum dimuat atau sudah kedaluwarsa. Klik CARI lalu ulangi."

End Function
