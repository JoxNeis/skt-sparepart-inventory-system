Attribute VB_Name = "serviceStockMovement"
Option Explicit

' =========================================================
' SERVICE - STOCK MOVEMENT (KARTU STOCK)
'
' Public entry points (called by sheets). Read only, so no
' transaction is started.
' =========================================================

Private Const MODULE_NAME As String = "SERVICE_STOCK_MOVEMENT"

' =========================================================
' ERROR CODES (17800 - 17899)
' Usage: ErrorManager.ThrowError serviceStockMovement.InvalidMovementMonth, ...
' =========================================================
Public Const EmptyManufacturingPlantName As Long = 17801
Public Const EmptyCatalogueName As Long = 17802
Public Const InvalidMovementMonth As Long = 17803       ' month outside 1 - 12
Public Const InvalidMovementYear As Long = 17804        ' year <= 0
Public Const NegativeBeginningStock As Long = 17805
Public Const InvalidStockDate As Long = 17806           ' stock date <= 0


' Stock movements of one catalogue in one month, with the
' running beginning / end stock filled in, starting from
' beginningStock.
Public Function SelectStockMovement( _
    ByVal manufacturingPlantName As String, _
    ByVal catalogueName As String, _
    ByVal movementMonth As Integer, _
    ByVal movementYear As Long, _
    Optional ByVal beginningStock As Double = 0) As Collection

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim stockMovements As Collection
    Dim stockMovementItem As StockMovement
    Dim currentStock As Double
    Dim errInfo As ErrorInfo

    manufacturingPlantName = Trim$(manufacturingPlantName)
    catalogueName = Trim$(catalogueName)

    ' =====================================================
    ' Step 1: Validate
    ' =====================================================
    RequirePlantName manufacturingPlantName, MODULE_NAME & ".SelectStockMovement"
    RequireCatalogueName catalogueName, MODULE_NAME & ".SelectStockMovement"

    If ValidationNumeric.IsOutOfRange(movementMonth, 1, 12) Then
        ErrorManager.ThrowError _
            serviceStockMovement.InvalidMovementMonth, _
            MODULE_NAME & ".SelectStockMovement", _
            "Bulan harus berada di antara 1 dan 12."
    End If

    If ValidationNumeric.IsZeroOrNegative(movementYear) Then
        ErrorManager.ThrowError _
            serviceStockMovement.InvalidMovementYear, _
            MODULE_NAME & ".SelectStockMovement", _
            "Tahun harus lebih besar dari 0."
    End If

    If ValidationNumeric.IsNegative(beginningStock) Then
        ErrorManager.ThrowError _
            serviceStockMovement.NegativeBeginningStock, _
            MODULE_NAME & ".SelectStockMovement", _
            "Stok awal tidak boleh kurang dari 0."
    End If

    ' =====================================================
    ' Step 2: Query
    ' =====================================================
    Set conn = DatabaseManager.GetConnection()

    Set stockMovements = daoStockMovement.SelectStockMovement( _
        conn, manufacturingPlantName, catalogueName, _
        movementMonth, movementYear)

    ' =====================================================
    ' Step 3: Running stock
    ' =====================================================
    currentStock = beginningStock

    For Each stockMovementItem In stockMovements

        stockMovementItem.beginningStock = currentStock

        currentStock = currentStock + _
            stockMovementItem.incomingAmount - _
            stockMovementItem.outgoingAmount

        stockMovementItem.EndStock = currentStock

    Next stockMovementItem

    Set SelectStockMovement = stockMovements

    Exit Function

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".SelectStockMovement", _
        "Gagal mengambil data pergerakan stok."

End Function


' Total stock of one catalogue before stockDate.
Public Function SelectBeginningStock( _
    ByVal stockDate As Date, _
    ByVal manufacturingPlantName As String, _
    ByVal catalogueName As String) As Double

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim errInfo As ErrorInfo

    manufacturingPlantName = Trim$(manufacturingPlantName)
    catalogueName = Trim$(catalogueName)

    If ValidationNumeric.IsZeroOrNegative(stockDate) Then
        ErrorManager.ThrowError _
            serviceStockMovement.InvalidStockDate, _
            MODULE_NAME & ".SelectBeginningStock", _
            "Tanggal stok tidak valid."
    End If

    RequirePlantName manufacturingPlantName, MODULE_NAME & ".SelectBeginningStock"
    RequireCatalogueName catalogueName, MODULE_NAME & ".SelectBeginningStock"

    Set conn = DatabaseManager.GetConnection()

    SelectBeginningStock = daoStockMovement.SelectBeginningStock( _
        conn, stockDate, manufacturingPlantName, catalogueName)

    Exit Function

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".SelectBeginningStock", _
        "Gagal mengambil data stok awal."

End Function


' =========================================================
' INTERNAL
' =========================================================

Private Sub RequirePlantName( _
    ByVal manufacturingPlantName As String, _
    ByVal errorSource As String)

    If ValidationText.IsBlank(manufacturingPlantName) Then
        ErrorManager.ThrowError _
            serviceStockMovement.EmptyManufacturingPlantName, _
            errorSource, _
            "Nama manufacturing plant tidak boleh kosong."
    End If

End Sub


Private Sub RequireCatalogueName( _
    ByVal catalogueName As String, _
    ByVal errorSource As String)

    If ValidationText.IsBlank(catalogueName) Then
        ErrorManager.ThrowError _
            serviceStockMovement.EmptyCatalogueName, _
            errorSource, _
            "Nama catalogue tidak boleh kosong."
    End If

End Sub
