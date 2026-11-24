Attribute VB_Name = "daoUpdateOrder"
Option Explicit

' =========================================================
' DAO - ORDER UPDATE REQUESTS
'
' One function = one saved Access query.
' Logic lives in serviceTransactionRequestUpdate / serviceOrder.
' =========================================================

Private Const MODULE_NAME As String = "DAO_UPDATE_ORDER"


' select_order_request_update
' Returns a disconnected recordset.
Public Function SelectOrderRequestUpdate( _
    ByVal conn As ADODB.connection, _
    ByVal manufacturingPlantName As String, _
    ByVal fromDate As Date, _
    ByVal untilDate As Date) As ADODB.recordset

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "select_order_request_update")

    managerDatabase.AddTextParameter cmd, "p_manufacturing_plant_name", manufacturingPlantName
    managerDatabase.AddDateParameter cmd, "p_from_date", fromDate
    managerDatabase.AddDateParameter cmd, "p_until_date", untilDate

    Set SelectOrderRequestUpdate = managerDatabase.ExecuteRecordset(cmd)

End Function


' update_receive_order_id
' Moves every receive from the old order to the new order.
' Returns the affected row count (0 is valid: no receives).
Public Function UpdateOrderReceives( _
    ByVal conn As ADODB.connection, _
    ByVal oldOrderUUID As String, _
    ByVal newOrderUUID As String) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "update_receive_order_id")

    managerDatabase.AddTextParameter cmd, "p_old_order_uuid", oldOrderUUID
    managerDatabase.AddTextParameter cmd, "p_new_order_uuid", newOrderUUID

    UpdateOrderReceives = managerDatabase.ExecuteNonQuery(cmd)

End Function
