Attribute VB_Name = "controllerInputPenerimaan"
Option Explicit

' =========================================================
' CONTROLLER - INPUT PENERIMAAN (sheet INPUT_PENERIMAAN)
'
' CARI lists the pending orders of the plant (order date in
' DARI - HINGGA). The user fills the receive columns of the
' orders that arrived and double-clicks SIMPAN.
'
'   LOKASI K5, DARI K9 / HINGGA K10 (double-click = date)
'   CARI S9:U9, SIMPAN P9:Q10, SEBELUMNYA S7, SELANJUTNYA U7,
'   FILTER KOLOM P7 / NILAI Q7, BARIS YANG DITAMPILKAN U10
'
'   TableInputPenerimaan, editable columns:
'     PO ID, TANGGAL DATANG / TANGGAL VERIFIKASI (double-click
'     = date picker), JUMLAH PENERIMAAN, JUMLAH OK,
'     JUMLAH NOT OK, REMARKS, PIC
'
'   JUMLAH PENERIMAAN fills OK (= all) and NOT OK (= 0); OK
'   and NOT OK keep each other in step. Pages come from
'   order_list_pending (paged by the query); what is typed is
'   remembered per order, so it survives paging and filtering,
'   and SIMPAN saves the typed rows of every page.
' =========================================================

Private Const MODULE_NAME As String = "controllerInputPenerimaan"

' =========================================================
' ERROR CODES (14200 - 14299)
' =========================================================
Public Const OrderNotLoaded As Long = 14201
Public Const NothingToSave As Long = 14202
Public Const EmptyPersonInCharge As Long = 14203

Private Const LAYOUT_KEY As String = "INPUT_PENERIMAAN"
Private Const TABLE_NAME As String = "TableInputPenerimaan"
Private Const BUTTON_SAVE As String = "P9:Q10"

Private Const COL_ORDER_UUID As String = "ORDER UUID"
Private Const COL_PO As String = "PO ID"
Private Const COL_SPAREPART As String = "NAMA SPAREPART"
Private Const COL_ORDER_DATE As String = "TANGGAL ORDER"
Private Const COL_ARRIVAL As String = "TANGGAL DATANG"
Private Const COL_VERIFICATION As String = "TANGGAL VERIFIKASI"
Private Const COL_REMAINING As String = "JUMLAH YANG HARUS DIPENUHI"
Private Const COL_RECEIVED As String = "JUMLAH PENERIMAAN"
Private Const COL_OK As String = "JUMLAH OK"
Private Const COL_NOT_OK As String = "JUMLAH NOT OK"
Private Const COL_UNIT As String = "SATUAN"
Private Const COL_REMARKS As String = "REMARKS"
Private Const COL_PIC As String = "PIC"

' Positions in the data array (= table column positions).
Private Const IX_UUID As Long = 1
Private Const IX_PO As Long = 2
Private Const IX_ARRIVAL As Long = 5
Private Const IX_VERIFICATION As Long = 6
Private Const IX_RECEIVED As Long = 8
Private Const IX_OK As Long = 9
Private Const IX_NOT_OK As Long = 10
Private Const IX_REMARKS As Long = 12
Private Const IX_PIC As Long = 13
Private Const IX_ORIGINAL_PO As Long = 14       ' hidden
Private Const DATA_COLUMNS As Long = 14

Private mOrders As Object             ' order uuid -> Order (every page seen)
Private mTyped As Object              ' order uuid -> typed values (Array, IX_ positions)
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
    Dim delta As Long

    Set layout = GetLayout()
    Set tbl = managerTable.GetTable(ws, TABLE_NAME)

    If managerPaging.IsDateCell(ws, layout, target) _
        Or managerTable.IsInColumn(target.Cells(1, 1), tbl, COL_ARRIVAL) _
        Or managerTable.IsInColumn(target.Cells(1, 1), tbl, COL_VERIFICATION) Then

        managerForm.HandleDatePickerDoubleClick target, target.Cells(1, 1), cancel
        Exit Sub
    End If

    delta = managerPaging.PageButtonDelta(ws, layout, target)

    If delta <> 0 Then
        cancel = True
        CaptureTyped ws
        MovePage ws, delta
        Exit Sub
    End If

    If managerSheet.IsHit(target, ws, layout.SearchButton) Then
        cancel = True
        Search ws, True
    ElseIf managerSheet.IsHit(target, ws, BUTTON_SAVE) Then
        cancel = True
        SaveReceives ws
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
        If managerPaging.HasState(LAYOUT_KEY) Then
            CaptureTyped ws
            LoadPage ws, managerPaging.FirstPageRequest(ws, layout, FilterHeaders(), FilterIndexes()), 1
        End If
        Exit Sub
    End If

    Set tbl = managerTable.GetTable(ws, TABLE_NAME)

    If Not managerTable.IsInTable(target, tbl) Then Exit Sub

    If managerSheet.UndoIfReadOnly(target, tbl.DataBodyRange, EditableRange(tbl), _
        "Kolom ini hanya bisa dibaca. Isi kolom berwarna kuning.") Then Exit Sub

    If target.Cells.Count = 1 Then
        BalanceQuantities tbl, target
    End If

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerError.ShowError errInfo

End Sub


' =========================================================
' ACTIONS
' =========================================================

' CARI. askWhenEdited: warn before throwing away typed rows.
Public Sub Search( _
    ByVal ws As Worksheet, _
    Optional ByVal askWhenEdited As Boolean = False)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo
    Dim layout As PagingLayout

    Set layout = GetLayout()

    If askWhenEdited Then
        CaptureTyped ws
        If HasTypedRows() Then
            If Not managerForm.Confirm( _
                "Data penerimaan yang sudah diisi belum disimpan dan akan hilang. Lanjutkan?") Then
                Exit Sub
            End If
        End If
    End If

    managerPaging.ReadCriteria ws, layout, mPlantName, mFromDate, mToDate

    Set mOrders = CreateObject("Scripting.Dictionary")
    Set mTyped = CreateObject("Scripting.Dictionary")

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

    If mOrders Is Nothing Then Set mOrders = CreateObject("Scripting.Dictionary")
    If mTyped Is Nothing Then Set mTyped = CreateObject("Scripting.Dictionary")

    If items.Count > 0 Then

        ReDim data(1 To items.Count, 1 To DATA_COLUMNS)

        For Each orderItem In items

            r = r + 1
            Set mOrders(orderItem.UUID) = orderItem

            data(r, IX_UUID) = orderItem.UUID
            data(r, IX_PO) = orderItem.PoNumber
            data(r, 3) = orderItem.Sparepart.DisplayName
            data(r, 4) = orderItem.OrderDate
            data(r, 7) = orderItem.RemainingQuantity
            data(r, 11) = orderItem.UnitOfMeasurement
            data(r, IX_ORIGINAL_PO) = orderItem.PoNumber
            ApplyTyped data, r

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
        Application.StatusBar = "Tidak ada order yang menunggu penerimaan pada periode ini."
    End If

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


' SIMPAN: every typed row of every page.
Public Sub SaveReceives( _
    ByVal ws As Worksheet)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo
    Dim receives As Collection
    Dim poChanges As Collection
    Dim change As Variant
    Dim uuids As Collection
    Dim key As Variant
    Dim values As Variant
    Dim message As String
    Dim uuid As Variant
    Dim pageNumber As Long
    Dim request As PageRequest

    If Not managerPaging.HasState(LAYOUT_KEY) Then
        managerError.ThrowError controllerInputPenerimaan.NothingToSave, _
            MODULE_NAME & ".SaveReceives", "Klik CARI terlebih dahulu."
    End If

    CaptureTyped ws

    Set receives = New Collection
    Set poChanges = New Collection

    For Each key In mTyped.Keys

        values = mTyped(key)

        If CStr(values(IX_PO)) <> CStr(values(IX_ORIGINAL_PO)) Then
            poChanges.Add Array(CStr(key), Trim$(CStr(values(IX_PO))))
        End If

        If ValuesAreTyped(values) Then
            receives.Add BuildReceive(CStr(key), values)
        End If

    Next key

    If receives.Count = 0 And poChanges.Count = 0 Then
        managerError.ThrowError controllerInputPenerimaan.NothingToSave, _
            MODULE_NAME & ".SaveReceives", _
            "Belum ada penerimaan yang diisi (JUMLAH PENERIMAAN)."
    End If

    message = "Simpan " & receives.Count & " penerimaan"
    If poChanges.Count > 0 Then message = message & " dan " & poChanges.Count & " PO ID"
    message = message & " untuk " & mPlantName & "?"

    If Not managerForm.Confirm(message) Then Exit Sub

    state = managerForm.BeginBusy("Menyimpan penerimaan...", True)

    For Each change In poChanges
        serviceOrder.SetPoNumber CStr(change(0)), CStr(change(1))
    Next change

    message = vbNullString

    If receives.Count > 0 Then
        Set uuids = serviceReceive.CreateReceives(mPlantName, receives)

        message = receives.Count & " penerimaan tersimpan:" & vbCrLf
        For Each uuid In uuids
            message = message & "  " & CStr(uuid) & vbCrLf
        Next uuid
    End If

    If poChanges.Count > 0 Then
        message = message & poChanges.Count & " PO ID tersimpan." & vbCrLf
    End If

    managerForm.EndBusy state

    managerForm.ShowInfo message

    ' Start again from page 1 with the same filter.
    mTyped.RemoveAll
    Set request = managerPaging.CurrentRequest(GetLayout(), pageNumber)
    LoadPage ws, request.FirstPage(), 1

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

    FilterHeaders = Array(COL_ORDER_UUID, COL_PO, COL_SPAREPART, COL_ORDER_DATE, _
        COL_REMAINING, COL_UNIT)

End Function


' order_list_pending column number of each KOLOM choice.
Private Function FilterIndexes() As Variant

    FilterIndexes = Array(1, 2, 3, 7, 13, 5)

End Function


Private Function EditableIndexes() As Variant

    EditableIndexes = Array(IX_PO, IX_ARRIVAL, IX_VERIFICATION, IX_RECEIVED, _
        IX_OK, IX_NOT_OK, IX_REMARKS, IX_PIC)

End Function


Private Function EditableRange( _
    ByVal tbl As ListObject) As Range

    Set EditableRange = Union( _
        managerTable.GetColumnRange(tbl, COL_PO), _
        managerTable.GetColumnRange(tbl, COL_ARRIVAL), _
        managerTable.GetColumnRange(tbl, COL_VERIFICATION), _
        managerTable.GetColumnRange(tbl, COL_RECEIVED), _
        managerTable.GetColumnRange(tbl, COL_OK), _
        managerTable.GetColumnRange(tbl, COL_NOT_OK), _
        managerTable.GetColumnRange(tbl, COL_REMARKS), _
        managerTable.GetColumnRange(tbl, COL_PIC))

End Function


Private Function ColumnFormats() As Variant

    ColumnFormats = Array( _
        managerSheet.TEXT_FORMAT, managerSheet.TEXT_FORMAT, managerSheet.TEXT_FORMAT, _
        managerSheet.DATE_FORMAT, managerSheet.DATE_FORMAT, managerSheet.DATE_FORMAT, _
        managerSheet.NUMBER_FORMAT, managerSheet.NUMBER_FORMAT, managerSheet.NUMBER_FORMAT, _
        managerSheet.NUMBER_FORMAT, managerSheet.TEXT_FORMAT, managerSheet.TEXT_FORMAT, _
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

End Sub


' A row counts as a receive when a date or quantity is typed.
Private Function ValuesAreTyped( _
    ByVal values As Variant) As Boolean

    Dim columnIndex As Variant

    For Each columnIndex In Array(IX_ARRIVAL, IX_VERIFICATION, IX_RECEIVED, IX_OK, IX_NOT_OK)
        If managerValidationText.IsNotBlank(values(CLng(columnIndex))) Then
            ValuesAreTyped = True
            Exit Function
        End If
    Next columnIndex

End Function


' Remember what is typed on the page shown (per order), so it
' survives paging, filtering and is saved by SIMPAN.
Private Sub CaptureTyped( _
    ByVal ws As Worksheet)

    Dim pagingState As PageState
    Dim tbl As ListObject
    Dim tableRow As Long
    Dim values() As Variant
    Dim columnIndex As Variant
    Dim orderUUID As String

    If mTyped Is Nothing Then Exit Sub
    If Not managerPaging.HasState(LAYOUT_KEY) Then Exit Sub

    Set pagingState = managerPaging.GetState(LAYOUT_KEY)
    Set tbl = managerTable.GetTable(ws, TABLE_NAME)

    For tableRow = 1 To pagingState.PageRowCount

        If tableRow > tbl.ListRows.Count Then Exit For

        orderUUID = CStr(pagingState.CellValue(tableRow, IX_UUID))

        ReDim values(1 To DATA_COLUMNS)
        For Each columnIndex In EditableIndexes()
            values(CLng(columnIndex)) = tbl.DataBodyRange.Cells(tableRow, CLng(columnIndex)).value
        Next columnIndex
        values(IX_ORIGINAL_PO) = pagingState.CellValue(tableRow, IX_ORIGINAL_PO)

        If ValuesAreTyped(values) _
            Or CStr(values(IX_PO)) <> CStr(values(IX_ORIGINAL_PO)) Then
            mTyped(orderUUID) = values
        ElseIf mTyped.Exists(orderUUID) Then
            mTyped.Remove orderUUID
        End If

    Next tableRow

End Sub


' Put the remembered values of an order back on its row.
Private Sub ApplyTyped( _
    ByRef data() As Variant, _
    ByVal dataRow As Long)

    Dim values As Variant
    Dim columnIndex As Variant

    If mTyped Is Nothing Then Exit Sub
    If Not mTyped.Exists(CStr(data(dataRow, IX_UUID))) Then Exit Sub

    values = mTyped(CStr(data(dataRow, IX_UUID)))

    For Each columnIndex In EditableIndexes()
        data(dataRow, CLng(columnIndex)) = values(CLng(columnIndex))
    Next columnIndex

End Sub


Private Function HasTypedRows() As Boolean

    Dim key As Variant

    If mTyped Is Nothing Then Exit Function

    For Each key In mTyped.Keys
        If ValuesAreTyped(mTyped(key)) Then
            HasTypedRows = True
            Exit Function
        End If
    Next key

End Function


' One typed row -> Receive. Errors name the order.
Private Function BuildReceive( _
    ByVal orderUUID As String, _
    ByVal values As Variant) As Receive

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo
    Dim label As String
    Dim receiveItem As Receive
    Dim received As Long
    Dim okQuantity As Long
    Dim notOkQuantity As Long
    Dim okValue As Variant
    Dim notOkValue As Variant
    Dim picText As String

    label = "Order " & orderUUID & ": "

    received = managerSheet.ParseWholeNumber( _
        values(IX_RECEIVED), label & "JUMLAH PENERIMAAN", 1)

    okValue = values(IX_OK)
    notOkValue = values(IX_NOT_OK)

    If managerValidationText.IsBlank(okValue) And managerValidationText.IsBlank(notOkValue) Then
        okQuantity = received
        notOkQuantity = 0
    ElseIf managerValidationText.IsBlank(okValue) Then
        notOkQuantity = managerSheet.ParseWholeNumber(notOkValue, label & "JUMLAH NOT OK", 0)
        okQuantity = received - notOkQuantity
    ElseIf managerValidationText.IsBlank(notOkValue) Then
        okQuantity = managerSheet.ParseWholeNumber(okValue, label & "JUMLAH OK", 0)
        notOkQuantity = received - okQuantity
    Else
        okQuantity = managerSheet.ParseWholeNumber(okValue, label & "JUMLAH OK", 0)
        notOkQuantity = managerSheet.ParseWholeNumber(notOkValue, label & "JUMLAH NOT OK", 0)
    End If

    picText = Trim$(coreHelper.TextOrEmpty(values(IX_PIC)))

    If Len(picText) = 0 Then
        managerError.ThrowError controllerInputPenerimaan.EmptyPersonInCharge, _
            MODULE_NAME & ".BuildReceive", label & "PIC wajib diisi."
    End If

    Set receiveItem = New Receive

    With receiveItem
        Set .Order = CachedOrder(orderUUID)
        .ArrivalDate = managerSheet.ParseDate( _
            values(IX_ARRIVAL), label & "TANGGAL DATANG")
        .VerificationDate = managerSheet.ParseDate( _
            values(IX_VERIFICATION), label & "TANGGAL VERIFIKASI", False)
        .ReceivedQuantity = received
        .OkQuantity = okQuantity
        .NotOkQuantity = notOkQuantity
        .Remarks = coreHelper.TextOrEmpty(values(IX_REMARKS))
        .PersonInCharge = picText
        .Validate
    End With

    Set BuildReceive = receiveItem

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, MODULE_NAME & ".BuildReceive", label

End Function


' Keep JUMLAH PENERIMAAN = OK + NOT OK while the user types.
Private Sub BalanceQuantities( _
    ByVal tbl As ListObject, _
    ByVal target As Range)

    Dim rowIndex As Long
    Dim received As Variant
    Dim okValue As Variant
    Dim notOkValue As Variant
    Dim eventsWereOn As Boolean

    rowIndex = managerTable.GetRowIndex(tbl, target)
    If rowIndex = 0 Then Exit Sub

    received = managerTable.GetValue(tbl, rowIndex, COL_RECEIVED)
    okValue = managerTable.GetValue(tbl, rowIndex, COL_OK)
    notOkValue = managerTable.GetValue(tbl, rowIndex, COL_NOT_OK)

    If managerValidationNumeric.IsNotNumber(received) Then Exit Sub

    eventsWereOn = Application.EnableEvents
    Application.EnableEvents = False

    If managerTable.IsInColumn(target, tbl, COL_RECEIVED) Then

        If managerValidationNumeric.IsNumber(okValue) Then
            If CDbl(received) - CDbl(okValue) >= 0 Then
                managerTable.SetValue tbl, rowIndex, COL_NOT_OK, CDbl(received) - CDbl(okValue)
            Else
                managerTable.SetValue tbl, rowIndex, COL_OK, received
                managerTable.SetValue tbl, rowIndex, COL_NOT_OK, 0
            End If
        Else
            managerTable.SetValue tbl, rowIndex, COL_OK, received
            managerTable.SetValue tbl, rowIndex, COL_NOT_OK, 0
        End If

        If Not managerTable.HasValue(tbl, rowIndex, COL_ARRIVAL) Then
            managerTable.SetValue tbl, rowIndex, COL_ARRIVAL, Date
        End If

    ElseIf managerTable.IsInColumn(target, tbl, COL_OK) Then

        If managerValidationNumeric.IsNumber(okValue) Then
            managerTable.SetValue tbl, rowIndex, COL_NOT_OK, CDbl(received) - CDbl(okValue)
        End If

    ElseIf managerTable.IsInColumn(target, tbl, COL_NOT_OK) Then

        If managerValidationNumeric.IsNumber(notOkValue) Then
            managerTable.SetValue tbl, rowIndex, COL_OK, CDbl(received) - CDbl(notOkValue)
        End If

    End If

    Application.EnableEvents = eventsWereOn

End Sub


Private Function CachedOrder( _
    ByVal orderUUID As String) As Order

    If Not mOrders Is Nothing Then
        If mOrders.Exists(orderUUID) Then
            Set CachedOrder = mOrders(orderUUID)
            Exit Function
        End If
    End If

    managerError.ThrowError controllerInputPenerimaan.OrderNotLoaded, _
        MODULE_NAME & ".CachedOrder", _
        "Data order belum dimuat atau sudah kedaluwarsa. Klik CARI lalu ulangi."

End Function
