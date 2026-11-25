Attribute VB_Name = "controllerKartuStock"
Option Explicit

' =========================================================
' CONTROLLER - KARTU STOCK (sheet KARTU_STOCK)
'
'   PLANT          K5  dropdown (plant)
'   NAMA SPAREPART K6  dropdown (sparepart, "code - name")
'   PERIODE        K7  month (dropdown)   L7 year (dropdown)
'   LOKASI         K8  plant type, filled automatically
'   CARI           T6:V7   double-click
'   BARIS YANG DITAMPILKAN  V8  rows per page
'
'   HASIL PENCARIAN  M5:N7  BARIS DATA M7, TOTAL HALAMAN N7
'   HALAMAN          P5:R7  SEBELUMNYA P7, SEKARANG Q7,
'                           SELANJUTNYA R7 (double-click)
'   (the design has no paging cells on this sheet: both blocks
'   are copied once from REKAP_ORDER, see EnsurePagingBlocks)
'
'   TableKartuStock: one line per movement of the month. The
'   paging is done by stock_ledger_list_by_month /
'   stock_ledger_count_by_month; STOK AWAL of the first line
'   of any page comes from stock_ledger_get_balance_before, so
'   STOK AWAL / STOK AKHIR are right on every page.
' =========================================================

Private Const MODULE_NAME As String = "controllerKartuStock"

' =========================================================
' ERROR CODES (14500 - 14599)
' =========================================================
Public Const InvalidYear As Long = 14501

Private Const LAYOUT_KEY As String = "KARTU_STOCK"

Private Const PLANT_CELL As String = "K5"
Private Const SPAREPART_CELL As String = "K6"
Private Const MONTH_CELL As String = "K7"
Private Const YEAR_CELL As String = "L7"
Private Const LOCATION_CELL As String = "K8"

Private Const TABLE_NAME As String = "TableKartuStock"

Private Const TYPE_IN As String = "PENERIMAAN"
Private Const TYPE_OUT As String = "PENGELUARAN"
Private Const TYPE_OPENING As String = "SALDO AWAL"

' Paging blocks copied from REKAP_ORDER (HASIL PENCARIAN and
' HALAMAN) and where they go on this sheet.
Private Const SOURCE_SHEET As String = "REKAP_ORDER"
Private Const SOURCE_COUNT_BLOCK As String = "M5:N7"
Private Const SOURCE_PAGE_BLOCK As String = "S5:U7"
Private Const TARGET_COUNT_BLOCK As String = "M5"
Private Const TARGET_PAGE_BLOCK As String = "P5"
Private Const PAGE_BLOCK_HEADER As String = "HALAMAN"

' PERIODE was set to the current month in this session.
Private mPeriodSet As Boolean

' Criteria of the last CARI.
Private mPlantName As String
Private mSparepart As Sparepart
Private mYear As Integer
Private mMonth As Integer


' =========================================================
' EVENTS
' =========================================================

Public Sub OnActivate( _
    ByVal ws As Worksheet)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo

    state = managerForm.BeginBusy()

    EnsurePagingBlocks ws

    managerPaging.PrepareSheet ws, GetLayout()
    managerList.ApplyDropdown ws.Range(SPAREPART_CELL), managerList.LIST_SPAREPART

    managerList.ApplyItemsDropdown ws.Range(MONTH_CELL), managerSheet.IndonesianMonths()
    managerList.ApplyItemsDropdown ws.Range(YEAR_CELL), managerList.YearItems(), False

    ' The design stores JANUARI: start every session on the
    ' current month instead.
    If Not mPeriodSet _
        Or managerValidationText.IsBlank(ws.Range(MONTH_CELL).value) _
        Or managerValidationNumeric.IsNotNumber(ws.Range(YEAR_CELL).value) Then

        ws.Range(MONTH_CELL).value = managerSheet.IndonesianMonth(Month(Date))
        ws.Range(YEAR_CELL).value = Year(Date)
        mPeriodSet = True
    End If

    ShowLocation ws

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
    Dim delta As Long

    Set layout = GetLayout()

    delta = managerPaging.PageButtonDelta(ws, layout, target)

    If delta <> 0 Then
        cancel = True
        MovePage ws, delta
        Exit Sub
    End If

    If managerSheet.IsHit(target, ws, layout.SearchButton) Then
        cancel = True
        Search ws
    End If

End Sub


Public Sub OnChange( _
    ByVal ws As Worksheet, _
    ByVal target As Range)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    If managerSheet.IsHit(target, ws, PLANT_CELL) Then
        ShowLocation ws
    ElseIf managerPaging.IsViewCell(ws, GetLayout(), target) Then
        If managerPaging.HasState(LAYOUT_KEY) Then Search ws
    End If

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    Application.EnableEvents = True
    managerError.ShowError errInfo

End Sub


' =========================================================
' ACTIONS
' =========================================================

' CARI: remember plant, sparepart and period, load page 1.
Public Sub Search( _
    ByVal ws As Worksheet)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo
    Dim plantName As String
    Dim sparepartItem As Sparepart
    Dim monthNumber As Long
    Dim yearNumber As Long

    plantName = managerSheet.ReadText(ws, PLANT_CELL, "PLANT")
    monthNumber = managerSheet.ParseMonth(ws.Range(MONTH_CELL).value, ws.Range(MONTH_CELL))
    yearNumber = managerSheet.ReadWholeNumber(ws, YEAR_CELL, "TAHUN", 1900)

    state = managerForm.BeginBusy("Mencari sparepart...")

    Set sparepartItem = controllerShared.RequireSparepart( _
        serviceSparepart.ListSpareparts(), _
        managerSheet.ReadText(ws, SPAREPART_CELL, "NAMA SPAREPART"), _
        ws.Range(SPAREPART_CELL))

    managerForm.EndBusy state

    mPlantName = plantName
    Set mSparepart = sparepartItem
    mYear = CInt(yearNumber)
    mMonth = CInt(monthNumber)

    LoadPage ws, managerPaging.FirstPageRequest(ws, GetLayout()), 1

    ShowLocation ws

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

    With result
        .Key = LAYOUT_KEY
        .TableName = TABLE_NAME
        .PlantCell = PLANT_CELL
        .CountCell = "M7"
        .PagesCell = "N7"
        .PrevButton = "P7"
        .PageCell = "Q7"
        .NextButton = "R7"
        .SearchButton = "T6:V7"
        .PageSizeCell = "V8"
    End With

    Set GetLayout = result

End Function


' SEBELUMNYA / SELANJUTNYA: load the neighbouring page.
Private Sub MovePage( _
    ByVal ws As Worksheet, _
    ByVal delta As Long)

    Dim request As PageRequest
    Dim newPage As Long

    Set request = managerPaging.MoveRequest(GetLayout(), delta, newPage)

    If Not request Is Nothing Then LoadPage ws, request, newPage

End Sub


' Load one page of the stock card and show it. NO continues
' over the pages. A month without movement shows one SALDO
' AWAL line.
Private Sub LoadPage( _
    ByVal ws As Worksheet, _
    ByVal request As PageRequest, _
    ByVal pageNumber As Long)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo
    Dim entries As Collection
    Dim entry As StockLedgerEntry
    Dim totalRows As Long
    Dim beginningStock As Long
    Dim data() As Variant
    Dim rowCount As Long
    Dim r As Long
    Dim firstKey As PageRequest
    Dim lastKey As PageRequest

    If mSparepart Is Nothing Then
        managerError.ThrowError managerPaging.NoSearchResult, _
            MODULE_NAME & ".LoadPage", "Klik CARI terlebih dahulu."
    End If

    state = managerForm.BeginBusy("Mengambil kartu stok...", True)

    Set entries = serviceStockLedger.GetStockCardPage( _
        mPlantName, mSparepart.Id, mYear, mMonth, request, totalRows, beginningStock)

    rowCount = entries.Count
    If rowCount = 0 Then rowCount = 1        ' opening balance line

    ReDim data(1 To rowCount, 1 To 14)

    If entries.Count = 0 Then

        data(1, 1) = 1
        data(1, 2) = DateSerial(mYear, mMonth, 1)
        data(1, 3) = beginningStock
        data(1, 10) = beginningStock
        data(1, 11) = mSparepart.UnitOfMeasurement
        data(1, 12) = TYPE_OPENING
        data(1, 13) = "Tidak ada transaksi pada periode ini."

    Else

        For Each entry In entries

            r = r + 1

            data(r, 1) = (pageNumber - 1) * request.PageSize + r
            data(r, 2) = entry.MovementDate
            data(r, 3) = entry.BeginningStock

            If entry.Direction = "IN" Then
                data(r, 4) = entry.IncomingReference
                data(r, 5) = entry.IncomingQuantity
                data(r, 6) = entry.IncomingFrom
                data(r, 12) = TYPE_IN
            Else
                data(r, 7) = entry.OutgoingReference
                data(r, 8) = entry.OutgoingQuantity
                data(r, 9) = entry.OutgoingTo
                data(r, 12) = TYPE_OUT
            End If

            data(r, 10) = entry.EndStock
            data(r, 11) = mSparepart.UnitOfMeasurement
            data(r, 13) = entry.Remarks
            data(r, 14) = entry.PersonInCharge

            Set lastKey = serviceStockLedger.StockCardKey(entry)
            If r = 1 Then Set firstKey = lastKey

        Next entry

    End If

    managerPaging.StorePage GetLayout(), data, rowCount, totalRows, _
        pageNumber, request, firstKey, lastKey

    RenderPage ws

    managerForm.EndBusy state

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


Private Sub RenderPage( _
    ByVal ws As Worksheet)

    Dim eventsWereOn As Boolean

    eventsWereOn = Application.EnableEvents
    Application.EnableEvents = False

    managerPaging.Render ws, GetLayout(), _
        Array("0", managerSheet.DATE_FORMAT, managerSheet.NUMBER_FORMAT, _
              managerSheet.TEXT_FORMAT, managerSheet.NUMBER_FORMAT, managerSheet.TEXT_FORMAT, _
              managerSheet.TEXT_FORMAT, managerSheet.NUMBER_FORMAT, managerSheet.TEXT_FORMAT, _
              managerSheet.NUMBER_FORMAT, managerSheet.TEXT_FORMAT, managerSheet.TEXT_FORMAT, _
              managerSheet.TEXT_FORMAT, managerSheet.TEXT_FORMAT)

    Application.EnableEvents = eventsWereOn

End Sub


' The design has no paging cells on this sheet: copy the
' HASIL PENCARIAN and HALAMAN blocks of REKAP_ORDER (same look,
' arrows included) once into the empty area M5:R7.
Public Sub EnsurePagingBlocks( _
    ByVal ws As Worksheet)

    Dim source As Worksheet

    If UCase$(Trim$(coreHelper.TextOrEmpty( _
        ws.Range(TARGET_PAGE_BLOCK).value))) = PAGE_BLOCK_HEADER Then Exit Sub

    Set source = managerSheet.GetSheet(SOURCE_SHEET)

    source.Range(SOURCE_COUNT_BLOCK).Copy Destination:=ws.Range(TARGET_COUNT_BLOCK)
    source.Range(SOURCE_PAGE_BLOCK).Copy Destination:=ws.Range(TARGET_PAGE_BLOCK)

    ws.Range("M7:N7").ClearContents
    ws.Range("Q7").ClearContents

End Sub


' LOKASI = type of the selected plant (e.g. WAREHOUSE).
Private Sub ShowLocation( _
    ByVal ws As Worksheet)

    Dim eventsWereOn As Boolean

    eventsWereOn = Application.EnableEvents
    Application.EnableEvents = False

    ws.Range(LOCATION_CELL).value = _
        managerList.PlantType(managerForm.GetCellText(ws, PLANT_CELL))

    Application.EnableEvents = eventsWereOn

End Sub
