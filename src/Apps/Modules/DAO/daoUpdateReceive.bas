Attribute VB_Name = "daoUpdateReceive"
Option Explicit

' =========================================================
' DAO - RECEIVE UPDATE REQUESTS
'
' One function = one saved Access query.
' Logic lives in serviceTransactionRequestUpdate.
' =========================================================

Private Const MODULE_NAME As String = "DAO_UPDATE_RECEIVE"


' select_receive_request_update
' Returns a disconnected recordset.
Public Function SelectReceiveRequestUpdate( _
    ByVal conn As ADODB.connection, _
    ByVal manufacturingPlantName As String, _
    ByVal fromDate As Date, _
    ByVal untilDate As Date) As ADODB.recordset

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "select_receive_request_update")

    managerDatabase.AddTextParameter cmd, "p_manufacturing_plant_name", manufacturingPlantName
    managerDatabase.AddDateParameter cmd, "p_from_date", fromDate
    managerDatabase.AddDateParameter cmd, "p_until_date", untilDate

    Set SelectReceiveRequestUpdate = managerDatabase.ExecuteRecordset(cmd)

End Function
