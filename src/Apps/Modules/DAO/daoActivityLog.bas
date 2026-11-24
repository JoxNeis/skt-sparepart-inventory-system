Attribute VB_Name = "daoActivityLog"
Option Explicit

' =========================================================
' DAO - ACTIVITY LOG
'
' One function = one saved Access query.
' Use serviceActivityLog.WriteActivityLog from services; it
' validates the input and checks the affected row count.
' =========================================================

Private Const MODULE_NAME As String = "DAO_ACTIVITY_LOG"


' create_activity_log
' Returns the affected row count.
Public Function CreateActivityLog( _
    ByVal conn As ADODB.connection, _
    ByVal activityCode As String, _
    ByVal detail As String, _
    ByVal user As String) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "create_activity_log")

    managerDatabase.AddTextParameter cmd, "p_activity_code", activityCode
    managerDatabase.AddLongTextParameter cmd, "p_detail", detail
    managerDatabase.AddTextParameter cmd, "p_user", user

    CreateActivityLog = managerDatabase.ExecuteNonQuery(cmd)

End Function
