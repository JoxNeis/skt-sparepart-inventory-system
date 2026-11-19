Attribute VB_Name = "daoUpdateRequest"
Option Explicit

' =========================================================
' DAO - REQUEST UPDATE REQUESTS
'
' One function = one saved Access query.
' Logic lives in serviceTransactionRequestUpdate.
' =========================================================

Private Const MODULE_NAME As String = "DAO_UPDATE_REQUEST"


' select_request_request_update
' Returns a disconnected recordset.
Public Function SelectRequestRequestUpdate( _
    ByVal conn As ADODB.connection, _
    ByVal manufacturingPlantName As String, _
    ByVal fromDate As Date, _
    ByVal untilDate As Date) As ADODB.recordset

    Dim cmd As ADODB.Command

    Set cmd = DatabaseManager.CreateCommand(conn, "select_request_request_update")

    DatabaseManager.AddTextParameter cmd, "p_manufacturing_plant_name", manufacturingPlantName
    DatabaseManager.AddDateParameter cmd, "p_from_date", fromDate
    DatabaseManager.AddDateParameter cmd, "p_until_date", untilDate

    Set SelectRequestRequestUpdate = DatabaseManager.ExecuteRecordset(cmd)

End Function
