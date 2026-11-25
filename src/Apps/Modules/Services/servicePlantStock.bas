Attribute VB_Name = "servicePlantStock"
Option Explicit

' =========================================================
' SERVICE - PLANT STOCK
'
' manufacturing_plant_stocks.current_stock is a cache of the
' stock ledger (vw_stock_ledger). Every service that changes
' the ledger calls RecalculateCurrentStock for the touched
' spareparts INSIDE its transaction, after checking with
' RequireNonNegativeLedger that the stock card never goes
' below 0.
'
' Reorder point (rounded up to whole units):
'   safety_stock  = service_level_z * daily_demand_std_dev
'                   * Sqr(lead_time_days)
'   reorder_point = safety_stock
'                   + average_daily_usage * lead_time_days
' =========================================================

Private Const MODULE_NAME As String = "servicePlantStock"

' Any date after every possible movement.
Private Const END_OF_TIME As Date = #12/31/9999#

' Any date before every possible movement.
Private Const START_OF_TIME As Date = #1/1/1900#

' =========================================================
' ERROR CODES (17500 - 17599)
' Usage: managerError.ThrowError servicePlantStock.NegativeLedger, ...
' =========================================================
Public Const StockRowNotFound As Long = 17501
Public Const NegativeLedger As Long = 17502         ' stock card would go below 0
Public Const InvalidLeadTime As Long = 17503
Public Const InvalidAverageDailyUsage As Long = 17504
Public Const InvalidDailyDemandStdDev As Long = 17505
Public Const UpdateReorderPointFailed As Long = 17506


' Stock rows of the active spareparts of the plant, by
' sparepart name. Missing rows are created first (zero stock).
' page: filter and page (see PageRequest); Nothing = every row.
Public Function ListPlantStocks( _
    ByVal plantName As String, _
    Optional ByVal page As PageRequest = Nothing) As Collection

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim plantItem As ManufacturingPlant
    Dim inTransaction As Boolean
    Dim errInfo As ErrorInfo

    Set conn = managerDatabase.GetConnection()
    Set plantItem = serviceManufacturingPlant.FindPlant(conn, plantName)

    managerDatabase.BeginTransaction conn
    inTransaction = True

    daoPlantStock.CreateMissing conn, plantItem.Id

    managerDatabase.CommitTransaction conn
    inTransaction = False

    Set ListPlantStocks = daoPlantStock.ListByPlant(conn, plantItem, page)

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()
    If inTransaction Then managerDatabase.RollbackTransaction conn

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ListPlantStocks", _
        "Gagal mengambil data stok."

End Function


' Number of rows ListPlantStocks gives over all pages.
Public Function CountPlantStocks( _
    ByVal plantName As String, _
    Optional ByVal page As PageRequest = Nothing) As Long

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim errInfo As ErrorInfo

    Set conn = managerDatabase.GetConnection()

    CountPlantStocks = daoPlantStock.CountByPlant( _
        conn, serviceManufacturingPlant.FindPlant(conn, plantName), page)

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".CountPlantStocks", _
        "Gagal menghitung data stok."

End Function


' The stock row of one sparepart in stocks, or Nothing.
Public Function FindInList( _
    ByVal stocks As Collection, _
    ByVal sparepartId As Long) As ManufacturingPlantStock

    Dim stockItem As ManufacturingPlantStock

    For Each stockItem In stocks
        If stockItem.Sparepart.Id = sparepartId Then
            Set FindInList = stockItem
            Exit Function
        End If
    Next stockItem

End Function


' Store the planning inputs of one sparepart and recalculate
' its safety stock and reorder point. Returns the updated row.
Public Function UpdateReorderPoint( _
    ByVal plantName As String, _
    ByVal sparepartId As Long, _
    ByVal leadTimeDays As Long, _
    ByVal averageDailyUsage As Double, _
    ByVal dailyDemandStdDev As Double) As ManufacturingPlantStock

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim plantItem As ManufacturingPlant
    Dim stockItem As ManufacturingPlantStock
    Dim inTransaction As Boolean
    Dim affected As Long
    Dim errInfo As ErrorInfo

    If leadTimeDays < 0 Then
        managerError.ThrowError servicePlantStock.InvalidLeadTime, _
            MODULE_NAME & ".UpdateReorderPoint", _
            "Lead time tidak boleh negatif."
    End If

    If averageDailyUsage < 0 Then
        managerError.ThrowError servicePlantStock.InvalidAverageDailyUsage, _
            MODULE_NAME & ".UpdateReorderPoint", _
            "Average daily usage tidak boleh negatif."
    End If

    If dailyDemandStdDev < 0 Then
        managerError.ThrowError servicePlantStock.InvalidDailyDemandStdDev, _
            MODULE_NAME & ".UpdateReorderPoint", _
            "SD demand tidak boleh negatif."
    End If

    Set conn = managerDatabase.GetConnection()
    Set plantItem = serviceManufacturingPlant.FindPlant(conn, plantName)

    managerDatabase.BeginTransaction conn
    inTransaction = True

    daoPlantStock.CreateMissing conn, plantItem.Id

    affected = daoPlantStock.RecalculateReorderPoint( _
        conn, plantItem.Id, sparepartId, _
        leadTimeDays, averageDailyUsage, dailyDemandStdDev)

    managerDatabase.RequireAffectedRows affected, 1, _
        servicePlantStock.UpdateReorderPointFailed, _
        MODULE_NAME & ".UpdateReorderPoint", _
        "Reorder point tidak tersimpan (baris stok tidak ditemukan)."

    Set stockItem = FindInList(daoPlantStock.ListByPlant(conn, plantItem), sparepartId)

    If stockItem Is Nothing Then
        managerError.ThrowError servicePlantStock.StockRowNotFound, _
            MODULE_NAME & ".UpdateReorderPoint", _
            "Baris stok sparepart tidak ditemukan."
    End If

    serviceActivityLog.WriteActivityLog conn, _
        "PlantStock.UpdateROP", _
        "ROP " & stockItem.Sparepart.DisplayName & " di " & plantItem.Name & _
        " = " & stockItem.ReorderPoint & " (safety stock " & _
        stockItem.SafetyStock & ").", _
        coreJson.JsonObject( _
            "manufacturing_plant", coreJson.JsonObject( _
                "code", coreJson.JsonText(plantItem.Code), _
                "name", coreJson.JsonText(plantItem.Name)), _
            "sparepart", serviceSparepart.BuildSparepartReferenceJson(stockItem.Sparepart), _
            "lead_time_days", coreJson.JsonNumber(leadTimeDays), _
            "average_daily_usage", coreJson.JsonNumber(averageDailyUsage), _
            "daily_demand_std_dev", coreJson.JsonNumber(dailyDemandStdDev), _
            "service_level_z", coreJson.JsonNumber(stockItem.ServiceLevelZ), _
            "safety_stock", coreJson.JsonNumber(stockItem.SafetyStock), _
            "reorder_point", coreJson.JsonNumber(stockItem.ReorderPoint))

    managerDatabase.CommitTransaction conn
    inTransaction = False

    Set UpdateReorderPoint = stockItem

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()
    If inTransaction Then managerDatabase.RollbackTransaction conn

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".UpdateReorderPoint", _
        "Gagal menyimpan reorder point."

End Function


' Set current_stock of one sparepart to the full ledger
' balance. Call INSIDE the caller's transaction, after
' RequireNonNegativeLedger.
Public Sub RecalculateCurrentStock( _
    ByVal conn As ADODB.connection, _
    ByVal plantId As Long, _
    ByVal sparepartId As Long)

    Dim balance As Long
    Dim affected As Long

    balance = daoStockLedger.GetBalance(conn, plantId, sparepartId, END_OF_TIME)

    If balance < 0 Then balance = 0

    affected = daoPlantStock.SetCurrentStock(conn, plantId, sparepartId, balance)

    If affected = 0 Then
        daoPlantStock.CreateMissing conn, plantId
        affected = daoPlantStock.SetCurrentStock(conn, plantId, sparepartId, balance)
    End If

    managerDatabase.RequireAffectedRows affected, 1, _
        servicePlantStock.StockRowNotFound, _
        MODULE_NAME & ".RecalculateCurrentStock", _
        "Baris stok sparepart (id " & sparepartId & ") tidak ditemukan."

End Sub


' Raise NegativeLedger when the stock card of the sparepart
' goes below 0 on any day. Call INSIDE the caller's
' transaction, after the change, so the change is counted.
Public Sub RequireNonNegativeLedger( _
    ByVal conn As ADODB.connection, _
    ByVal plantId As Long, _
    ByVal sparepartId As Long, _
    ByVal sparepartLabel As String)

    Dim lowest As Long

    lowest = daoStockLedger.GetAvailableQuantity( _
        conn, plantId, sparepartId, START_OF_TIME)

    If lowest < 0 Then
        managerError.ThrowError servicePlantStock.NegativeLedger, _
            MODULE_NAME & ".RequireNonNegativeLedger", _
            "Stok " & sparepartLabel & " akan menjadi minus (" & lowest & _
            ") pada kartu stok. Perubahan dibatalkan."
    End If

End Sub
