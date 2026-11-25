Attribute VB_Name = "controllerRekapInventory"
Option Explicit

' =========================================================
' CONTROLLER - REKAP INVENTORY (sheet REKAP_INVENTORY)
'
'   LOKASI   K5   dropdown (plant)
'   PERIODE  K6   filled on REFRESH (current month)
'   REFRESH  Q8:R8   double-click
'   DATA YANG DITAMPILKAN  R9 (rows per page)
'   DATA PER TANGGAL       R10 (time of the last REFRESH)
'   SEBELUMNYA L10, SELANJUTNYA N10, SEKARANG M10,
'   BARIS DATA I10, TOTAL HALAMAN J10
'
'   TableRekapInventory: stock and reorder point per sparepart
'     AVERAGE DAILY USAGE, SD DEMAND, LEAD TIME : editable
'     SIMPAN : double-click recalculates safety stock and
'              reorder point of the row
'   JUMLAH STOK at or below the reorder point is shown in red.
'   Paging is done by manufacturing_plant_stock_list /
'   manufacturing_plant_stock_count (one page at a time).
' =========================================================

Private Const MODULE_NAME As String = "controllerRekapInventory"

' =========================================================
' ERROR CODES (14400 - 14499)
' =========================================================
Public Const StockNotLoaded As Long = 14401

Private Const LAYOUT_KEY As String = "REKAP_INVENTORY"
Private Const TABLE_NAME As String = "TableRekapInventory"
Private Const PERIOD_CELL As String = "K6"
Private Const AS_OF_CELL As String = "R10"

Private Const COL_NO As String = "NO"
Private Const COL_SPAREPART As String = "NAMA SPAREPART"
Private Const COL_STOCK As String = "JUMLAH STOK"
Private Const COL_UNIT As String = "SATUAN"
Private Const COL_ROP As String = "REORDER POINT"
Private Const COL_SAFETY As String = "SAFETY STOCK"
Private Const COL_USAGE As String = "AVERAGE DAILY USAGE"
Private Const COL_SD As String = "SD DEMAND"
Private Const COL_LEAD As String = "LEAD TIME"
Private Const COL_SAVE As String = "SIMPAN"

Private Const IX_ROP As Long = 5
Private Const IX_SAFETY As Long = 6
Private Const IX_USAGE As Long = 7
Private Const IX_SD As Long = 8
Private Const IX_LEAD As Long = 9
Private Const IX_SPAREPART_ID As Long = 11     ' hidden
Private Const DATA_COLUMNS As Long = 11

Private Const RULE_EDITABLE As String = "SIS_INVENTORY_EDITABLE"
Private Const RULE_URGENT As String = "SIS_INVENTORY_URGENT"

Private mPlantName As String


' =========================================================
' EVENTS
' =========================================================

Public Sub OnActivate( _
    ByVal ws As Worksheet)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo

    state = managerForm.BeginBusy()

    managerPaging.PrepareSheet ws, GetLayout()
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
    Dim rowIndex As Long

    Set layout = GetLayout()

    delta = managerPaging.PageButtonDelta(ws, layout, target)

    If delta <> 0 Then
        cancel = True
        MovePage ws, delta
        Exit Sub
    End If

    If managerSheet.IsHit(target, ws, layout.SearchButton) Then
        cancel = True
        Refresh ws
        Exit Sub
    End If

    Set tbl = managerTable.GetTable(ws, TABLE_NAME)

    If managerTable.IsInColumn(target.Cells(1, 1), tbl, COL_SAVE) Then
        cancel = True
        rowIndex = managerTable.GetRowIndex(tbl, target.Cells(1, 1))
        If rowIndex > 0 Then SaveReorderPoint ws, rowIndex
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
        If managerPaging.HasState(LAYOUT_KEY) Then Refresh ws
        Exit Sub
    End If

    Set tbl = managerTable.GetTable(ws, TABLE_NAME)

    If Not managerTable.IsInTable(target, tbl) Then Exit Sub

    managerSheet.UndoIfReadOnly target, tbl.DataBodyRange, EditableRange(tbl), _
        "Hanya AVERAGE DAILY USAGE, SD DEMAND dan LEAD TIME yang bisa diubah."

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerError.ShowError errInfo

End Sub


' =========================================================
' ACTIONS
' =========================================================

' REFRESH: load page 1 of the stock list of the plant (count
' + rows, paged by manufacturing_plant_stock_list).
Public Sub Refresh( _
    ByVal ws As Worksheet)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo
    Dim layout As PagingLayout

    Set layout = GetLayout()
    mPlantName = managerSheet.ReadText(ws, layout.PlantCell, "LOKASI")

    LoadPage ws, managerPaging.FirstPageRequest(ws, layout), 1

    ' As text: Excel would turn "NOVEMBER 2026" into a date.
    ws.Range(PERIOD_CELL).NumberFormat = managerSheet.TEXT_FORMAT
    ws.Range(PERIOD_CELL).value = managerSheet.PeriodText(Date)
    ws.Range(AS_OF_CELL).NumberFormat = "dd-mmm-yyyy hh:mm"
    ws.Range(AS_OF_CELL).value = Now

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


' Load one page of the stock list and show it. NO continues
' over the pages.
Private Sub LoadPage( _
    ByVal ws As Worksheet, _
    ByVal request As PageRequest, _
    ByVal pageNumber As Long)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo
    Dim totalRows As Long
    Dim stocks As Collection
    Dim stockItem As ManufacturingPlantStock
    Dim data() As Variant
    Dim r As Long
    Dim keyHolder As New PageRequest
    Dim firstKey As PageRequest
    Dim lastKey As PageRequest

    state = managerForm.BeginBusy("Mengambil data inventory...", True)

    Set stocks = servicePlantStock.ListPlantStocks(mPlantName, request)
    totalRows = servicePlantStock.CountPlantStocks(mPlantName, request)

    If stocks.Count > 0 Then

        ReDim data(1 To stocks.Count, 1 To DATA_COLUMNS)

        For Each stockItem In stocks

            r = r + 1

            data(r, 1) = (pageNumber - 1) * request.PageSize + r
            data(r, 2) = stockItem.Sparepart.DisplayName
            data(r, 3) = stockItem.CurrentStock
            data(r, 4) = stockItem.Sparepart.UnitOfMeasurement
            data(r, IX_ROP) = stockItem.ReorderPoint
            data(r, IX_SAFETY) = stockItem.SafetyStock
            data(r, IX_USAGE) = stockItem.AverageDailyUsage
            data(r, IX_SD) = stockItem.DailyDemandStdDev
            data(r, IX_LEAD) = stockItem.LeadTimeDays
            data(r, 10) = managerSheet.LABEL_SAVE
            data(r, IX_SPAREPART_ID) = stockItem.Sparepart.Id

            Set lastKey = keyHolder.KeyOf(#1/1/1900#, stockItem.Id, stockItem.Sparepart.Name)
            If r = 1 Then Set firstKey = lastKey

        Next stockItem

        managerPaging.StorePage GetLayout(), data, stocks.Count, totalRows, _
            pageNumber, request, firstKey, lastKey
    Else
        managerPaging.StorePage GetLayout(), Empty, 0, totalRows, _
            pageNumber, request, Nothing, Nothing
    End If

    RenderPage ws

    managerForm.EndBusy state

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


' SIMPAN of one row: store the planning inputs and show the
' recalculated safety stock and reorder point.
Public Sub SaveReorderPoint( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo
    Dim tbl As ListObject
    Dim pagingState As PageState
    Dim dataRow As Long
    Dim sparepartId As Long
    Dim leadTime As Long
    Dim usage As Double
    Dim deviation As Double
    Dim stockItem As ManufacturingPlantStock
    Dim eventsWereOn As Boolean

    Set tbl = managerTable.GetTable(ws, TABLE_NAME)

    dataRow = managerPaging.DataRowOf(LAYOUT_KEY, rowIndex)

    If dataRow = 0 Or Len(mPlantName) = 0 Then
        managerError.ThrowError controllerRekapInventory.StockNotLoaded, _
            MODULE_NAME & ".SaveReorderPoint", _
            "Data inventory belum dimuat. Klik REFRESH lalu ulangi."
    End If

    Set pagingState = managerPaging.GetState(LAYOUT_KEY)
    sparepartId = CLng(pagingState.CellValue(dataRow, IX_SPAREPART_ID))

    usage = managerSheet.TableDecimal(tbl, rowIndex, COL_USAGE)
    deviation = managerSheet.TableDecimal(tbl, rowIndex, COL_SD)
    leadTime = managerSheet.TableWholeNumber(tbl, rowIndex, COL_LEAD, 0, False, 0)

    state = managerForm.BeginBusy("Menghitung reorder point...")

    Set stockItem = servicePlantStock.UpdateReorderPoint( _
        mPlantName, sparepartId, leadTime, usage, deviation)

    pagingState.CellValue(dataRow, IX_ROP) = stockItem.ReorderPoint
    pagingState.CellValue(dataRow, IX_SAFETY) = stockItem.SafetyStock
    pagingState.CellValue(dataRow, IX_USAGE) = stockItem.AverageDailyUsage
    pagingState.CellValue(dataRow, IX_SD) = stockItem.DailyDemandStdDev
    pagingState.CellValue(dataRow, IX_LEAD) = stockItem.LeadTimeDays

    eventsWereOn = Application.EnableEvents
    Application.EnableEvents = False
    managerTable.SetValue tbl, rowIndex, COL_ROP, stockItem.ReorderPoint
    managerTable.SetValue tbl, rowIndex, COL_SAFETY, stockItem.SafetyStock
    Application.EnableEvents = eventsWereOn

    managerForm.EndBusy state

    Application.StatusBar = stockItem.Sparepart.DisplayName & ": reorder point " & _
        stockItem.ReorderPoint & ", safety stock " & stockItem.SafetyStock & "."

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
        .PlantCell = "K5"
        .CountCell = "I10"
        .PagesCell = "J10"
        .PrevButton = "L10"
        .PageCell = "M10"
        .NextButton = "N10"
        .SearchButton = "Q8:R8"
        .PageSizeCell = "R9"
    End With

    Set GetLayout = result

End Function


Private Function EditableRange( _
    ByVal tbl As ListObject) As Range

    Set EditableRange = Union( _
        managerTable.GetColumnRange(tbl, COL_USAGE), _
        managerTable.GetColumnRange(tbl, COL_SD), _
        managerTable.GetColumnRange(tbl, COL_LEAD))

End Function


Private Function ColumnFormats() As Variant

    ColumnFormats = Array("0", managerSheet.TEXT_FORMAT, managerSheet.NUMBER_FORMAT, _
        managerSheet.TEXT_FORMAT, managerSheet.NUMBER_FORMAT, managerSheet.NUMBER_FORMAT, _
        managerSheet.DECIMAL_FORMAT, managerSheet.DECIMAL_FORMAT, managerSheet.NUMBER_FORMAT, _
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


' Yellow planning inputs, red stock at / below the ROP.
Private Sub PrepareTable( _
    ByVal ws As Worksheet)

    Dim tbl As ListObject
    Dim columnName As Variant

    Set tbl = managerTable.GetTable(ws, TABLE_NAME)

    ' One rule per input column (each needs its own tag).
    For Each columnName In Array(COL_USAGE, COL_SD, COL_LEAD)
        managerTable.EnsureFormulaRule tbl, _
            RULE_EDITABLE & "_" & managerTable.ColumnLetter(tbl, CStr(columnName)), _
            "TRUE", CStr(columnName), managerSheet.EDIT_FILL_COLOR
    Next columnName

    managerTable.EnsureFormulaRule tbl, RULE_URGENT, _
        "AND(" & managerTable.RowCellFormula(tbl, COL_ROP) & ">0," & _
        managerTable.RowCellFormula(tbl, COL_STOCK) & "<=" & _
        managerTable.RowCellFormula(tbl, COL_ROP) & ")", _
        COL_STOCK, -1, RGB(192, 0, 0), True

    managerList.ApplyWholeNumberRule managerTable.GetColumnRange(tbl, COL_LEAD), 0

End Sub
