Attribute VB_Name = "servicePlantStock"
Option Explicit

' =========================================================
' SERVICE - PLANT STOCK
'
' Public entry points (called by sheets) get the
' connection and own the transaction:
'   SelectManufacturingPlantStocks, UpdateROP
'
' Helpers for other services (run INSIDE the caller's
' transaction, receive conn, never commit / roll back):
'   ApplyStockChange, RecalculateStock
' =========================================================

Private Const MODULE_NAME As String = "SERVICE_PLANT_STOCK"

' Z-score used for the safety stock (95% service level).
Private Const SAFETY_STOCK_Z_SCORE As Double = 1.65

' =========================================================
' ERROR CODES (17500 - 17599)
' Usage: managerError.ThrowError servicePlantStock.ZeroStockChanges, ...
' =========================================================
Public Const ZeroStockChanges As Long = 17501
Public Const EmptyCatalogueName As Long = 17502
Public Const EmptyManufacturingPlantName As Long = 17503
Public Const NoRowAffectedUpdateStock As Long = 17504        ' stock row not found
Public Const MultipleRowsAffectedUpdateStock As Long = 17505
Public Const NegativeAverageDailyUsage As Long = 17506
Public Const NegativeSdDemand As Long = 17507
Public Const NegativeLeadTime As Long = 17508
Public Const NoRowAffectedUpdateROP As Long = 17509          ' stock row not found
Public Const MultipleRowsAffectedUpdateROP As Long = 17510


' =========================================================
' PUBLIC ENTRY POINTS
' =========================================================

' Stock of every catalogue of one manufacturing plant.
' Returns a disconnected recordset (close it after use).
Public Function SelectManufacturingPlantStocks( _
    ByVal manufacturingPlantName As String) As ADODB.recordset

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim errInfo As ErrorInfo

    manufacturingPlantName = Trim$(manufacturingPlantName)

    RequirePlantName manufacturingPlantName, _
        MODULE_NAME & ".SelectManufacturingPlantStocks"

    Set conn = managerDatabase.GetConnection()

    Set SelectManufacturingPlantStocks = _
        daoPlantStock.SelectManufacturingPlantStocks( _
            conn, manufacturingPlantName)

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".SelectManufacturingPlantStocks", _
        "Gagal mengambil data stok manufacturing plant."

End Function


' Update the ROP inputs of one stock row and write the
' activity log, in one transaction.
Public Sub UpdateROP( _
    ByVal averageDailyUsage As Double, _
    ByVal sdDemand As Double, _
    ByVal leadTime As Double, _
    ByVal catalogueName As String, _
    ByVal manufacturingPlantName As String)

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim affected As Long
    Dim inTransaction As Boolean
    Dim errInfo As ErrorInfo

    catalogueName = Trim$(catalogueName)
    manufacturingPlantName = Trim$(manufacturingPlantName)

    ' =====================================================
    ' Step 1: Validate
    ' =====================================================
    If managerValidationNumeric.IsNegative(averageDailyUsage) Then
        managerError.ThrowError _
            servicePlantStock.NegativeAverageDailyUsage, _
            MODULE_NAME & ".UpdateROP", _
            "Average daily usage tidak boleh kurang dari nol."
    End If

    If managerValidationNumeric.IsNegative(sdDemand) Then
        managerError.ThrowError _
            servicePlantStock.NegativeSdDemand, _
            MODULE_NAME & ".UpdateROP", _
            "SD demand tidak boleh kurang dari nol."
    End If

    If managerValidationNumeric.IsNegative(leadTime) Then
        managerError.ThrowError _
            servicePlantStock.NegativeLeadTime, _
            MODULE_NAME & ".UpdateROP", _
            "Lead time tidak boleh kurang dari nol."
    End If

    RequireCatalogueName catalogueName, MODULE_NAME & ".UpdateROP"
    RequirePlantName manufacturingPlantName, MODULE_NAME & ".UpdateROP"

    ' =====================================================
    ' Step 2: Begin transaction
    ' =====================================================
    Set conn = managerDatabase.GetConnection()

    managerDatabase.BeginTransaction conn
    inTransaction = True

    ' =====================================================
    ' Step 3: Update ROP
    ' =====================================================
    affected = daoPlantStock.UpdateROP( _
        conn, averageDailyUsage, sdDemand, leadTime, _
        catalogueName, manufacturingPlantName)

    If managerValidationNumeric.IsLessThan(affected, 1) Then
        managerError.ThrowError _
            servicePlantStock.NoRowAffectedUpdateROP, _
            MODULE_NAME & ".UpdateROP", _
            "ROP tidak diperbarui. Data stock untuk " & _
            "catalogue '" & catalogueName & _
            "' pada manufacturing plant '" & _
            manufacturingPlantName & "' tidak ditemukan."
    End If

    If managerValidationNumeric.IsGreaterThan(affected, 1) Then
        managerError.ThrowError _
            servicePlantStock.MultipleRowsAffectedUpdateROP, _
            MODULE_NAME & ".UpdateROP", _
            "Pembaruan ROP memengaruhi lebih dari " & _
            "satu baris. Transaksi dibatalkan."
    End If

    ' =====================================================
    ' Step 4: Activity log
    ' =====================================================
    serviceActivityLog.WriteActivityLog _
        conn, _
        "PlantStock.UpdateROP", _
        BuildROPUpdateDetail( _
            catalogueName, manufacturingPlantName, _
            averageDailyUsage, sdDemand, leadTime)

    ' =====================================================
    ' Step 5: Commit
    ' =====================================================
    managerDatabase.CommitTransaction conn
    inTransaction = False

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    If inTransaction Then
        managerDatabase.RollbackTransaction conn
    End If

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".UpdateROP", _
        "Gagal memperbarui reorder point dan safety stock."

End Sub


' =========================================================
' HELPERS FOR OTHER SERVICES (inside caller's transaction)
' =========================================================

' Add stockChanges (negative = stock out) to the current
' stock. Exactly one stock row must change.
Public Sub ApplyStockChange( _
    ByVal conn As ADODB.connection, _
    ByVal stockChanges As Double, _
    ByVal catalogueName As String, _
    ByVal manufacturingPlantName As String)

    On Error GoTo ErrorHandler

    Dim affected As Long
    Dim errInfo As ErrorInfo

    catalogueName = Trim$(catalogueName)
    manufacturingPlantName = Trim$(manufacturingPlantName)

    If managerValidationNumeric.IsZero(stockChanges) Then
        managerError.ThrowError _
            servicePlantStock.ZeroStockChanges, _
            MODULE_NAME & ".ApplyStockChange", _
            "Perubahan stok tidak boleh nol."
    End If

    RequireCatalogueName catalogueName, MODULE_NAME & ".ApplyStockChange"
    RequirePlantName manufacturingPlantName, MODULE_NAME & ".ApplyStockChange"

    affected = daoPlantStock.UpdateCurrentStock( _
        conn, stockChanges, catalogueName, manufacturingPlantName)

    RequireSingleStockRow affected, MODULE_NAME & ".ApplyStockChange"

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ApplyStockChange", _
        "Gagal memperbarui stok saat ini."

End Sub


' Recalculate the current stock from all movements until
' now + 1 day and overwrite the stored value.
Public Sub RecalculateStock( _
    ByVal conn As ADODB.connection, _
    ByVal catalogueName As String, _
    ByVal manufacturingPlantName As String)

    On Error GoTo ErrorHandler

    Dim calculationDate As Date
    Dim currentStock As Double
    Dim affected As Long
    Dim errInfo As ErrorInfo

    catalogueName = Trim$(catalogueName)
    manufacturingPlantName = Trim$(manufacturingPlantName)

    RequireCatalogueName catalogueName, MODULE_NAME & ".RecalculateStock"
    RequirePlantName manufacturingPlantName, MODULE_NAME & ".RecalculateStock"

    ' Example: 21-Sep-2026 10:00 becomes 22-Sep-2026 10:00.
    calculationDate = DateAdd("d", 1, Now())

    currentStock = daoStockMovement.SelectBeginningStock( _
        conn, calculationDate, manufacturingPlantName, catalogueName)

    affected = daoPlantStock.UpdateCurrentStockExact( _
        conn, currentStock, catalogueName, manufacturingPlantName)

    RequireSingleStockRow affected, MODULE_NAME & ".RecalculateStock"

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".RecalculateStock", _
        "Gagal memperbarui stok secara tepat."

End Sub


' =========================================================
' INTERNAL
' =========================================================

Private Sub RequireCatalogueName( _
    ByVal catalogueName As String, _
    ByVal errorSource As String)

    If managerValidationText.IsBlank(catalogueName) Then
        managerError.ThrowError _
            servicePlantStock.EmptyCatalogueName, _
            errorSource, _
            "Nama catalogue tidak boleh kosong."
    End If

End Sub


Private Sub RequirePlantName( _
    ByVal manufacturingPlantName As String, _
    ByVal errorSource As String)

    If managerValidationText.IsBlank(manufacturingPlantName) Then
        managerError.ThrowError _
            servicePlantStock.EmptyManufacturingPlantName, _
            errorSource, _
            "Nama manufacturing plant tidak boleh kosong."
    End If

End Sub


Private Sub RequireSingleStockRow( _
    ByVal affected As Long, _
    ByVal errorSource As String)

    If managerValidationNumeric.IsLessThan(affected, 1) Then
        managerError.ThrowError _
            servicePlantStock.NoRowAffectedUpdateStock, _
            errorSource, _
            "Stok tidak diperbarui. Data stok untuk " & _
            "manufacturing plant dan catalogue tersebut " & _
            "tidak ditemukan."
    End If

    If managerValidationNumeric.IsGreaterThan(affected, 1) Then
        managerError.ThrowError _
            servicePlantStock.MultipleRowsAffectedUpdateStock, _
            errorSource, _
            "Pembaruan stok memengaruhi lebih dari satu baris."
    End If

End Sub


Private Function BuildROPUpdateDetail( _
    ByVal catalogueName As String, _
    ByVal manufacturingPlantName As String, _
    ByVal averageDailyUsage As Double, _
    ByVal sdDemand As Double, _
    ByVal leadTime As Double) As String

    Dim safetyStock As Double
    Dim reorderPoint As Double
    Dim detail As String

    safetyStock = sdDemand * Sqr(leadTime) * SAFETY_STOCK_Z_SCORE

    reorderPoint = safetyStock + (averageDailyUsage * leadTime)

    detail = "catalogue_name=" & Trim$(catalogueName)
    detail = detail & "; manufacturing_plant_name=" & Trim$(manufacturingPlantName)
    detail = detail & "; average_daily_usage=" & Format$(averageDailyUsage, "0.0000")
    detail = detail & "; sd_demand=" & Format$(sdDemand, "0.0000")
    detail = detail & "; lead_time=" & Format$(leadTime, "0.0000")
    detail = detail & "; z_score=1.65"
    detail = detail & "; safety_stock=" & Format$(safetyStock, "0.0000")
    detail = detail & "; reorder_point=" & Format$(reorderPoint, "0.0000")

    BuildROPUpdateDetail = detail

End Function
