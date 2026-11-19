Attribute VB_Name = "sKartuStock"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Const MODULE_NAME As String = "UI_STOCK_MOVEMENT"

Private Const TABLE_NAME As String = "TableKartuStock"

' =========================================================
' Table column headers (TableKartuStock, A11:L)
'
' These names must exactly match the Excel table headers.
' =========================================================
Private Const COLUMN_NO As String = "No"
Private Const COLUMN_TRANSACTION_DATE As String = "Tanggal Transaksi"
Private Const COLUMN_BEGINNING_STOCK As String = "Stock Awal"
Private Const COLUMN_INCOMING_REFERENCE As String = "No Bukti Penerimaan"
Private Const COLUMN_INCOMING_FROM As String = "Terima Dari"
Private Const COLUMN_INCOMING_AMOUNT As String = "Jumlah Terima"
Private Const COLUMN_OUTGOING_REFERENCE As String = "No Bukti Pengeluaran"
Private Const COLUMN_OUTGOING_TO As String = "Keluar Ke"
Private Const COLUMN_OUTGOING_AMOUNT As String = "Jumlah Keluar"
Private Const COLUMN_END_STOCK As String = "Stock Akhir"
Private Const COLUMN_REMARKS As String = "Remarks"
Private Const COLUMN_PIC As String = "PIC"

' =========================================================
' Filter cells
' =========================================================
Private Const CELL_PLANT As String = "B4"
Private Const CELL_MONTH As String = "B6"
Private Const CELL_YEAR As String = "C6"
Private Const CELL_CATALOGUE As String = "B8"

Public Sub LoadStockMovement()

    On Error GoTo ErrorHandler

    Dim tbl As ListObject
    Dim stockMovements As Collection
    Dim stockMovementItem As StockMovement

    Dim manufacturingPlantName As String
    Dim catalogueName As String
    Dim periodMonth As Integer
    Dim periodYear As Long
    Dim yearValue As Variant
    Dim beginningStock As Double

    Dim rowIndex As Long
    Dim itemNumber As Long

    Dim allColumns As Variant
    Dim columnName As Variant
    Dim columnRange As Range

    Dim state As FormState
    Dim errInfo As ErrorInfo

    ' =====================================================
    ' Read filter values from worksheet
    '
    ' B4 = Lokasi
    ' B6 = Bulan
    ' C6 = Tahun
    ' B8 = Nama Barang
    ' =====================================================
    manufacturingPlantName = FormManager.GetCellText( _
        Me, CELL_PLANT)

    catalogueName = FormManager.GetCellText( _
        Me, CELL_CATALOGUE)

    periodMonth = GetMonthNumber( _
        FormManager.GetCellValue(Me, CELL_MONTH))

    yearValue = FormManager.GetCellValue(Me, CELL_YEAR)

    ' Not IsNumeric (instead of ValidationNumeric.IsNotNumber)
    ' keeps accepting a blank year as 0 here, as before; it is
    ' rejected below as "Periode tahun tidak valid."
    If Not IsNumeric(yearValue) Then
        ErrorManager.ThrowError _
            sheetKartuStock.InvalidYearNumber, _
            MODULE_NAME & ".LoadStockMovement", _
            "Tahun harus berupa angka."
    End If

    periodYear = CLng(yearValue)

    ' =====================================================
    ' Validate worksheet inputs
    ' =====================================================
    If ValidationText.IsBlank(manufacturingPlantName) Then
        ErrorManager.ThrowError _
            sheetKartuStock.EmptyPlantName, _
            MODULE_NAME & ".LoadStockMovement", _
            "Lokasi tidak boleh kosong."
    End If

    If ValidationText.IsBlank(catalogueName) Then
        ErrorManager.ThrowError _
            sheetKartuStock.EmptyCatalogueName, _
            MODULE_NAME & ".LoadStockMovement", _
            "Nama barang tidak boleh kosong."
    End If

    If ValidationNumeric.IsOutOfRange(periodMonth, 1, 12) Then
        ErrorManager.ThrowError _
            sheetKartuStock.InvalidPeriodMonth, _
            MODULE_NAME & ".LoadStockMovement", _
            "Periode bulan tidak valid."
    End If

    If ValidationNumeric.IsZeroOrNegative(periodYear) Then
        ErrorManager.ThrowError _
            sheetKartuStock.InvalidPeriodYear, _
            MODULE_NAME & ".LoadStockMovement", _
            "Periode tahun tidak valid."
    End If

    state = FormManager.BeginBusy()

    ' =====================================================
    ' Call service functions
    ' (after the input checks, so the sheet's own messages
    ' are shown for a blank location / catalogue)
    ' =====================================================
    beginningStock = serviceStockMovement.SelectBeginningStock( _
        stockDate:=DateSerial(periodYear, periodMonth, 1), _
        manufacturingPlantName:=manufacturingPlantName, _
        catalogueName:=catalogueName)

    Set stockMovements = serviceStockMovement.SelectStockMovement( _
        manufacturingPlantName:=manufacturingPlantName, _
        catalogueName:=catalogueName, _
        movementMonth:=periodMonth, _
        movementYear:=periodYear, _
        beginningStock:=beginningStock)

    ' =====================================================
    ' Get and validate TableKartuStock
    ' =====================================================
    Set tbl = TableManager.GetTable(Me, TABLE_NAME)

    allColumns = Array( _
        COLUMN_NO, _
        COLUMN_TRANSACTION_DATE, _
        COLUMN_BEGINNING_STOCK, _
        COLUMN_INCOMING_REFERENCE, _
        COLUMN_INCOMING_FROM, _
        COLUMN_INCOMING_AMOUNT, _
        COLUMN_OUTGOING_REFERENCE, _
        COLUMN_OUTGOING_TO, _
        COLUMN_OUTGOING_AMOUNT, _
        COLUMN_END_STOCK, _
        COLUMN_REMARKS, _
        COLUMN_PIC)

    TableManager.ValidateColumns tbl, allColumns

    ' =====================================================
    ' Clear previous output
    '
    ' The table rows stay; only their values are cleared.
    ' =====================================================
    TableManager.ClearContents tbl

    ' =====================================================
    ' Write stock movements to TableKartuStock
    '
    ' Existing (cleared) rows are reused; a row is added
    ' only when the table has no free row left.
    ' =====================================================
    rowIndex = 0
    itemNumber = 1

    For Each stockMovementItem In stockMovements

        rowIndex = rowIndex + 1

        If Not TableManager.IsValidRow(tbl, rowIndex) Then
            rowIndex = TableManager.AddRow(tbl)
        End If

        TableManager.SetValue _
            tbl, rowIndex, COLUMN_NO, itemNumber

        TableManager.SetValue _
            tbl, rowIndex, COLUMN_TRANSACTION_DATE, _
            CDate(stockMovementItem.MovementDate)

        TableManager.SetValue _
            tbl, rowIndex, COLUMN_BEGINNING_STOCK, _
            stockMovementItem.beginningStock

        TableManager.SetValue _
            tbl, rowIndex, COLUMN_INCOMING_REFERENCE, _
            stockMovementItem.IncomingReferenceNumber

        TableManager.SetValue _
            tbl, rowIndex, COLUMN_INCOMING_FROM, _
            stockMovementItem.IncomingFrom

        If stockMovementItem.incomingAmount > 0 Then
            TableManager.SetValue _
                tbl, rowIndex, COLUMN_INCOMING_AMOUNT, _
                stockMovementItem.incomingAmount
        Else
            TableManager.SetValue _
                tbl, rowIndex, COLUMN_INCOMING_AMOUNT, Empty
        End If

        TableManager.SetValue _
            tbl, rowIndex, COLUMN_OUTGOING_REFERENCE, _
            stockMovementItem.OutgoingReferenceNumber

        TableManager.SetValue _
            tbl, rowIndex, COLUMN_OUTGOING_TO, _
            stockMovementItem.OutgoingTo

        If stockMovementItem.outgoingAmount > 0 Then
            TableManager.SetValue _
                tbl, rowIndex, COLUMN_OUTGOING_AMOUNT, _
                stockMovementItem.outgoingAmount
        Else
            TableManager.SetValue _
                tbl, rowIndex, COLUMN_OUTGOING_AMOUNT, Empty
        End If

        TableManager.SetValue _
            tbl, rowIndex, COLUMN_END_STOCK, _
            stockMovementItem.EndStock

        TableManager.SetValue _
            tbl, rowIndex, COLUMN_REMARKS, _
            stockMovementItem.remarks

        TableManager.SetValue _
            tbl, rowIndex, COLUMN_PIC, _
            stockMovementItem.personInCharge

        itemNumber = itemNumber + 1

    Next stockMovementItem

    ' =====================================================
    ' Apply output formatting
    ' =====================================================
    If stockMovements.Count > 0 Then

        TableManager.FormatColumn _
            tbl, COLUMN_TRANSACTION_DATE, "dd-mmm-yyyy"

        TableManager.FormatColumn _
            tbl, COLUMN_BEGINNING_STOCK, "#,##0.00"

        TableManager.FormatColumn _
            tbl, COLUMN_INCOMING_AMOUNT, "#,##0.00"

        TableManager.FormatColumn _
            tbl, COLUMN_OUTGOING_AMOUNT, "#,##0.00"

        TableManager.FormatColumn _
            tbl, COLUMN_END_STOCK, "#,##0.00"

        For Each columnName In allColumns

            Set columnRange = TableManager.GetColumnRange( _
                tbl, CStr(columnName))

            If Not columnRange Is Nothing Then
                columnRange.VerticalAlignment = xlCenter
            End If

        Next columnName

    End If

    If stockMovements.Count = 0 Then
        FormManager.ShowInfo _
            "Tidak ada data pergerakan stok yang ditemukan " & _
            "untuk filter tersebut.", _
            "Kartu Stock"
    Else
        FormManager.ShowInfo _
            CStr(stockMovements.Count) & _
            " data pergerakan stok berhasil dimuat.", _
            "Kartu Stock"
    End If

CleanUp:
    FormManager.EndBusy state

    Set columnRange = Nothing
    Set stockMovementItem = Nothing
    Set stockMovements = Nothing
    Set tbl = Nothing

    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    FormManager.EndBusy state

    Set columnRange = Nothing
    Set stockMovementItem = Nothing
    Set stockMovements = Nothing
    Set tbl = Nothing

    ErrorManager.ShowError errInfo, _
        "Gagal menampilkan kartu stock."

End Sub

Private Function GetMonthNumber( _
    ByVal monthValue As Variant) As Integer

    Dim monthText As String

    If IsNumeric(monthValue) Then

        GetMonthNumber = CInt(monthValue)
        Exit Function

    End If

    monthText = LCase$(TrimText(CStr(monthValue)))

    Select Case monthText
        Case "januari", "january"
            GetMonthNumber = 1

        Case "februari", "february"
            GetMonthNumber = 2

        Case "maret", "march"
            GetMonthNumber = 3

        Case "april"
            GetMonthNumber = 4

        Case "mei", "may"
            GetMonthNumber = 5

        Case "juni", "june"
            GetMonthNumber = 6

        Case "juli", "july"
            GetMonthNumber = 7

        Case "agustus", "august"
            GetMonthNumber = 8

        Case "september"
            GetMonthNumber = 9

        Case "oktober", "october"
            GetMonthNumber = 10

        Case "november"
            GetMonthNumber = 11

        Case "desember", "december"
            GetMonthNumber = 12

        Case Else
            ErrorManager.ThrowError _
                sheetKartuStock.InvalidMonthName, _
                MODULE_NAME & ".GetMonthNumber", _
                "Nama bulan tidak valid: " & _
                CStr(monthValue)
    End Select

End Function
