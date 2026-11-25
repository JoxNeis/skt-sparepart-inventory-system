Attribute VB_Name = "daoActivityLog"
Option Explicit

' =========================================================
' DAO - ACTIVITY LOG
' Table: activity_logs
'
' One function = one saved Access query.
' Use serviceActivityLog.WriteActivityLog from services; it
' validates the input and checks the affected row count.
' =========================================================

Private Const MODULE_NAME As String = "daoActivityLog"


' activity_log_create
' Returns the affected row count.
Public Function CreateActivityLog( _
    ByVal conn As ADODB.connection, _
    ByVal activityCode As String, _
    ByVal detail As String, _
    ByVal json As String, _
    ByVal username As String) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "activity_log_create")

    managerDatabase.AddTextParameter cmd, "p_activity_code", activityCode, 50
    managerDatabase.AddNullableLongTextParameter cmd, "p_detail", detail
    managerDatabase.AddNullableLongTextParameter cmd, "p_json", json
    managerDatabase.AddTextParameter cmd, "p_username", username, 64

    CreateActivityLog = managerDatabase.ExecuteNonQuery(cmd)

End Function
