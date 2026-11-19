Attribute VB_Name = "daoUser"
Option Explicit

' =========================================================
' DAO - USER
'
' One function = one saved Access query.
' Validation, transaction, row-count checks, activity log
' and the can_approve conversion live in serviceUser.
' The SIMPAN button logic lives in sMasterDataUser.
' =========================================================

Private Const MODULE_NAME As String = "DAO_USER"

Private Const QUERY_SELECT_USERS As String = _
    "select_users"


' select_users
' Returns a disconnected recordset (the caller closes it).
Public Function SelectUsers( _
    ByVal conn As ADODB.connection) As ADODB.recordset

    Dim cmd As ADODB.Command

    Set cmd = DatabaseManager.CreateCommand(conn, QUERY_SELECT_USERS)

    Set SelectUsers = DatabaseManager.ExecuteRecordset(cmd)

End Function


' create_user
' Parameter order: p_manufacturing_plant_name, p_username,
' p_first_name, p_middle_name, p_last_name, p_position,
' p_tier.
' Returns the affected row count.
Public Function CreateUser( _
    ByVal conn As ADODB.connection, _
    ByVal manufacturingPlantName As String, _
    ByVal username As String, _
    ByVal firstName As String, _
    ByVal middleName As String, _
    ByVal lastName As String, _
    ByVal position As String, _
    ByVal tier As Long) As Long

    Dim cmd As ADODB.Command

    Set cmd = DatabaseManager.CreateCommand(conn, "create_user")

    DatabaseManager.AddTextParameter cmd, "p_manufacturing_plant_name", manufacturingPlantName
    DatabaseManager.AddTextParameter cmd, "p_username", username
    DatabaseManager.AddTextParameter cmd, "p_first_name", firstName
    DatabaseManager.AddTextParameter cmd, "p_middle_name", middleName
    DatabaseManager.AddTextParameter cmd, "p_last_name", lastName
    DatabaseManager.AddTextParameter cmd, "p_position", position
    DatabaseManager.AddLongParameter cmd, "p_tier", tier

    CreateUser = DatabaseManager.ExecuteNonQuery(cmd)

End Function


' compare_user
' Parameter order: p_requestor_username, p_approver_username.
' Returns a disconnected recordset with the can_approve
' field. No row means requestor or approver was not found.
' The caller closes it.
Public Function CompareUser( _
    ByVal conn As ADODB.connection, _
    ByVal requestorUsername As String, _
    ByVal approverUsername As String) As ADODB.recordset

    Dim cmd As ADODB.Command

    Set cmd = DatabaseManager.CreateCommand(conn, "compare_user")

    DatabaseManager.AddTextParameter cmd, "p_requestor_username", requestorUsername
    DatabaseManager.AddTextParameter cmd, "p_approver_username", approverUsername

    Set CompareUser = DatabaseManager.ExecuteRecordset(cmd)

End Function
