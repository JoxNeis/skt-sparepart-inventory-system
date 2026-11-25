Attribute VB_Name = "daoPlantStock"
Option Explicit

' =========================================================
' DAO - PLANT STOCK
' Table: manufacturing_plant_stocks
'
' One function = one saved Access query.
' current_stock is a cache of the stock ledger; services
' keep it in step (servicePlantStock.RecalculateCurrentStock).
' =========================================================

Private Const MODULE_NAME As String = "daoPlantStock"


' manufacturing_plant_stock_list
' Stock rows of one plant (active spareparts), by sparepart
' name. page: filter (FilterValue on "code - name" / alias)
' and page, keyed on (KeyText = name, KeyId = row id)
' (Nothing = every row).
Public Function ListByPlant( _
    ByVal conn As ADODB.connection, _
    ByVal plantItem As ManufacturingPlant, _
    Optional ByVal page As PageRequest = Nothing) As Collection

    Dim cmd As ADODB.Command
    Dim rs As ADODB.recordset
    Dim stocks As Collection

    If page Is Nothing Then Set page = New PageRequest

    Set cmd = managerDatabase.CreateCommand(conn, "manufacturing_plant_stock_list")

    managerDatabase.AddLongParameter cmd, "p_manufacturing_plant_id", plantItem.Id
    managerDatabase.AddPageFilterParameters cmd, page, False
    managerDatabase.AddTextParameter cmd, "p_direction", page.Direction, 4
    managerDatabase.AddTextParameter cmd, "p_key_text", page.KeyText, 255
    managerDatabase.AddLongParameter cmd, "p_key_id", page.KeyId

    Set rs = managerDatabase.ExecuteRecordset(cmd)
    managerDatabase.MoveToPageStart rs, page

    Set stocks = New Collection

    Do While Not rs.EOF
        If managerDatabase.IsPageFull(stocks.Count, page) Then Exit Do
        stocks.Add MapStock(rs, plantItem)
        rs.MoveNext
    Loop

    managerDatabase.CloseRecordset rs

    Set ListByPlant = stocks

End Function


' manufacturing_plant_stock_count
' Number of rows ListByPlant returns over all pages.
Public Function CountByPlant( _
    ByVal conn As ADODB.connection, _
    ByVal plantItem As ManufacturingPlant, _
    Optional ByVal page As PageRequest = Nothing) As Long

    Dim cmd As ADODB.Command

    If page Is Nothing Then Set page = New PageRequest

    Set cmd = managerDatabase.CreateCommand(conn, "manufacturing_plant_stock_count")

    managerDatabase.AddLongParameter cmd, "p_manufacturing_plant_id", plantItem.Id
    managerDatabase.AddPageFilterParameters cmd, page, False

    CountByPlant = managerDatabase.ExecuteCount(cmd)

End Function


' manufacturing_plant_stock_create_missing
' Adds a zero stock row for every active sparepart the plant
' does not have yet. Returns the number of rows added.
Public Function CreateMissing( _
    ByVal conn As ADODB.connection, _
    ByVal plantId As Long) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "manufacturing_plant_stock_create_missing")

    managerDatabase.AddLongParameter cmd, "p_manufacturing_plant_id", plantId

    CreateMissing = managerDatabase.ExecuteNonQuery(cmd)

End Function


' manufacturing_plant_stock_recalculate_reorder_point
' Stores the planning inputs and recalculates safety stock
' and reorder point in the query.
Public Function RecalculateReorderPoint( _
    ByVal conn As ADODB.connection, _
    ByVal plantId As Long, _
    ByVal sparepartId As Long, _
    ByVal leadTimeDays As Long, _
    ByVal averageDailyUsage As Double, _
    ByVal dailyDemandStdDev As Double) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, _
        "manufacturing_plant_stock_recalculate_reorder_point")

    managerDatabase.AddLongParameter cmd, "p_manufacturing_plant_id", plantId
    managerDatabase.AddLongParameter cmd, "p_sparepart_id", sparepartId
    managerDatabase.AddLongParameter cmd, "p_lead_time_days", leadTimeDays
    managerDatabase.AddDoubleParameter cmd, "p_average_daily_usage", averageDailyUsage
    managerDatabase.AddDoubleParameter cmd, "p_daily_demand_std_dev", dailyDemandStdDev

    RecalculateReorderPoint = managerDatabase.ExecuteNonQuery(cmd)

End Function


' manufacturing_plant_stock_set_current_stock
' 0 rows when the row does not exist or the stock < 0.
Public Function SetCurrentStock( _
    ByVal conn As ADODB.connection, _
    ByVal plantId As Long, _
    ByVal sparepartId As Long, _
    ByVal currentStock As Long) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, _
        "manufacturing_plant_stock_set_current_stock")

    managerDatabase.AddLongParameter cmd, "p_manufacturing_plant_id", plantId
    managerDatabase.AddLongParameter cmd, "p_sparepart_id", sparepartId
    managerDatabase.AddLongParameter cmd, "p_current_stock", currentStock

    SetCurrentStock = managerDatabase.ExecuteNonQuery(cmd)

End Function


' manufacturing_plant_stock_adjust_current_stock
' 0 rows when the stock would go below 0.
Public Function AdjustCurrentStock( _
    ByVal conn As ADODB.connection, _
    ByVal plantId As Long, _
    ByVal sparepartId As Long, _
    ByVal quantityChange As Long) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, _
        "manufacturing_plant_stock_adjust_current_stock")

    managerDatabase.AddLongParameter cmd, "p_manufacturing_plant_id", plantId
    managerDatabase.AddLongParameter cmd, "p_sparepart_id", sparepartId
    managerDatabase.AddLongParameter cmd, "p_quantity_change", quantityChange

    AdjustCurrentStock = managerDatabase.ExecuteNonQuery(cmd)

End Function


' manufacturing_plant_stock_update
' Stores every planning column as given (no recalculation).
Public Function UpdatePlantStock( _
    ByVal conn As ADODB.connection, _
    ByVal stockItem As ManufacturingPlantStock) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "manufacturing_plant_stock_update")

    managerDatabase.AddLongParameter cmd, "p_manufacturing_plant_id", stockItem.ManufacturingPlant.Id
    managerDatabase.AddLongParameter cmd, "p_sparepart_id", stockItem.Sparepart.Id
    managerDatabase.AddLongParameter cmd, "p_reorder_point", stockItem.ReorderPoint
    managerDatabase.AddLongParameter cmd, "p_safety_stock", stockItem.SafetyStock
    managerDatabase.AddLongParameter cmd, "p_lead_time_days", stockItem.LeadTimeDays
    managerDatabase.AddDoubleParameter cmd, "p_average_daily_usage", stockItem.AverageDailyUsage
    managerDatabase.AddDoubleParameter cmd, "p_daily_demand_std_dev", stockItem.DailyDemandStdDev
    managerDatabase.AddDoubleParameter cmd, "p_service_level_z", stockItem.ServiceLevelZ

    UpdatePlantStock = managerDatabase.ExecuteNonQuery(cmd)

End Function


' =========================================================
' INTERNAL
' =========================================================

Private Function MapStock( _
    ByVal rs As ADODB.recordset, _
    ByVal plantItem As ManufacturingPlant) As ManufacturingPlantStock

    Dim stockItem As ManufacturingPlantStock
    Dim sparepartItem As Sparepart

    Set sparepartItem = New Sparepart

    With sparepartItem
        .Id = managerDatabase.FieldLong(rs, "sparepart_id")
        .Code = managerDatabase.FieldText(rs, "sparepart_code")
        .Name = managerDatabase.FieldText(rs, "sparepart_name")
        .UnitOfMeasurement = managerDatabase.FieldText(rs, "sparepart_unit_of_measurement")
    End With

    Set stockItem = New ManufacturingPlantStock

    With stockItem
        .Id = managerDatabase.FieldLong(rs, "id")
        Set .ManufacturingPlant = plantItem
        Set .Sparepart = sparepartItem
        .CurrentStock = managerDatabase.FieldLong(rs, "current_stock")
        .ReorderPoint = managerDatabase.FieldLong(rs, "reorder_point")
        .SafetyStock = managerDatabase.FieldLong(rs, "safety_stock")
        .LeadTimeDays = managerDatabase.FieldLong(rs, "lead_time_days")
        .AverageDailyUsage = managerDatabase.FieldDouble(rs, "average_daily_usage")
        .DailyDemandStdDev = managerDatabase.FieldDouble(rs, "daily_demand_std_dev")
        .ServiceLevelZ = managerDatabase.FieldDouble(rs, "service_level_z")
    End With

    Set MapStock = stockItem

End Function
