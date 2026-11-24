Attribute VB_Name = "daoPlantStock"
Option Explicit

' =========================================================
' DAO - PLANT STOCK
'
' One function = one saved Access query.
' Logic lives in servicePlantStock.
' =========================================================

Private Const MODULE_NAME As String = "DAO_PLANT_STOCK"


' update_current_stock
' Adds stockChanges (can be negative) to the current stock.
' Returns the affected row count.
Public Function UpdateCurrentStock( _
    ByVal conn As ADODB.connection, _
    ByVal stockChanges As Double, _
    ByVal catalogueName As String, _
    ByVal manufacturingPlantName As String) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "update_current_stock")

    managerDatabase.AddDoubleParameter cmd, "p_stock_changes", stockChanges
    managerDatabase.AddTextParameter cmd, "p_catalogue_name", catalogueName
    managerDatabase.AddTextParameter cmd, "p_manufacturing_plant_name", manufacturingPlantName

    UpdateCurrentStock = managerDatabase.ExecuteNonQuery(cmd)

End Function


' update_current_stock_exact
' Overwrites the current stock with newStock.
' Returns the affected row count.
Public Function UpdateCurrentStockExact( _
    ByVal conn As ADODB.connection, _
    ByVal newStock As Double, _
    ByVal catalogueName As String, _
    ByVal manufacturingPlantName As String) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "update_current_stock_exact")

    managerDatabase.AddDoubleParameter cmd, "p_new_stock", newStock
    managerDatabase.AddTextParameter cmd, "p_catalogue_name", catalogueName
    managerDatabase.AddTextParameter cmd, "p_manufacturing_plant_name", manufacturingPlantName

    UpdateCurrentStockExact = managerDatabase.ExecuteNonQuery(cmd)

End Function


' select_plant_stocks
' Returns a disconnected recordset.
Public Function SelectManufacturingPlantStocks( _
    ByVal conn As ADODB.connection, _
    ByVal manufacturingPlantName As String) As ADODB.recordset

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "select_plant_stocks")

    managerDatabase.AddTextParameter cmd, "p_manufacturing_plant_name", manufacturingPlantName

    Set SelectManufacturingPlantStocks = managerDatabase.ExecuteRecordset(cmd)

End Function


' update_rop
' Returns the affected row count.
Public Function UpdateROP( _
    ByVal conn As ADODB.connection, _
    ByVal averageDailyUsage As Double, _
    ByVal sdDemand As Double, _
    ByVal leadTime As Double, _
    ByVal catalogueName As String, _
    ByVal manufacturingPlantName As String) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "update_rop")

    managerDatabase.AddDoubleParameter cmd, "p_average_daily_usage", averageDailyUsage
    managerDatabase.AddDoubleParameter cmd, "p_sd_demand", sdDemand
    managerDatabase.AddDoubleParameter cmd, "p_lead_time", leadTime
    managerDatabase.AddTextParameter cmd, "p_catalogue_name", catalogueName
    managerDatabase.AddTextParameter cmd, "p_manufacturing_plant_name", manufacturingPlantName

    UpdateROP = managerDatabase.ExecuteNonQuery(cmd)

End Function
