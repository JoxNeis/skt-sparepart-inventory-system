Attribute VB_Name = "daoUser"
Option Explicit

' =========================================================
' DAO - USER
' Table: users
'
' One function = one saved Access query.
' Validation, transaction, row-count checks and activity
' log live in serviceUser.
' =========================================================

Private Const MODULE_NAME As String = "daoUser"


' user_list
' Active users of every plant, ordered by username.
Public Function ListUsers( _
    ByVal conn As ADODB.connection) As Collection

    Dim cmd As ADODB.Command
    Dim rs As ADODB.recordset
    Dim users As Collection

    Set cmd = managerDatabase.CreateCommand(conn, "user_list")
    Set rs = managerDatabase.ExecuteRecordset(cmd)

    Set users = New Collection

    Do While Not rs.EOF
        users.Add MapUser(rs)
        rs.MoveNext
    Loop

    managerDatabase.CloseRecordset rs

    Set ListUsers = users

End Function


' user_create
Public Function CreateUser( _
    ByVal conn As ADODB.connection, _
    ByVal userItem As User) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "user_create")

    managerDatabase.AddLongParameter cmd, "p_manufacturing_plant_id", userItem.ManufacturingPlant.Id
    managerDatabase.AddTextParameter cmd, "p_username", userItem.Username, 64
    managerDatabase.AddTextParameter cmd, "p_first_name", userItem.FirstName, 100
    managerDatabase.AddNullableTextParameter cmd, "p_middle_name", userItem.MiddleName, 100
    managerDatabase.AddNullableTextParameter cmd, "p_last_name", userItem.LastName, 100
    managerDatabase.AddNullableTextParameter cmd, "p_jobdesk", userItem.Jobdesk, 255
    managerDatabase.AddLongParameter cmd, "p_tier", userItem.Tier

    CreateUser = managerDatabase.ExecuteNonQuery(cmd)

End Function


' user_update (the username cannot change)
Public Function UpdateUser( _
    ByVal conn As ADODB.connection, _
    ByVal userItem As User) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "user_update")

    managerDatabase.AddLongParameter cmd, "p_user_id", userItem.Id
    managerDatabase.AddLongParameter cmd, "p_manufacturing_plant_id", userItem.ManufacturingPlant.Id
    managerDatabase.AddTextParameter cmd, "p_first_name", userItem.FirstName, 100
    managerDatabase.AddNullableTextParameter cmd, "p_middle_name", userItem.MiddleName, 100
    managerDatabase.AddNullableTextParameter cmd, "p_last_name", userItem.LastName, 100
    managerDatabase.AddNullableTextParameter cmd, "p_jobdesk", userItem.Jobdesk, 255
    managerDatabase.AddLongParameter cmd, "p_tier", userItem.Tier

    UpdateUser = managerDatabase.ExecuteNonQuery(cmd)

End Function


' user_archive
Public Function ArchiveUser( _
    ByVal conn As ADODB.connection, _
    ByVal userId As Long) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "user_archive")

    managerDatabase.AddLongParameter cmd, "p_user_id", userId

    ArchiveUser = managerDatabase.ExecuteNonQuery(cmd)

End Function


' user_check_can_approve
' False when either user is unknown or the approver is
' archived or does not have a higher tier.
Public Function CheckCanApprove( _
    ByVal conn As ADODB.connection, _
    ByVal requestedBy As String, _
    ByVal approvedBy As String) As Boolean

    Dim cmd As ADODB.Command
    Dim rs As ADODB.recordset

    Set cmd = managerDatabase.CreateCommand(conn, "user_check_can_approve")

    managerDatabase.AddTextParameter cmd, "p_requested_by", requestedBy, 64
    managerDatabase.AddTextParameter cmd, "p_approved_by", approvedBy, 64

    Set rs = managerDatabase.ExecuteRecordset(cmd)

    If rs.EOF Then
        CheckCanApprove = False
    Else
        CheckCanApprove = managerDatabase.FieldBoolean(rs, "can_approve")
    End If

    managerDatabase.CloseRecordset rs

End Function


' =========================================================
' INTERNAL
' =========================================================

Private Function MapUser( _
    ByVal rs As ADODB.recordset) As User

    Dim userItem As User
    Dim plantItem As ManufacturingPlant

    Set plantItem = New ManufacturingPlant
    plantItem.Id = managerDatabase.FieldLong(rs, "manufacturing_plant_id")
    plantItem.Name = managerDatabase.FieldText(rs, "manufacturing_plant_name")

    Set userItem = New User

    With userItem
        .Id = managerDatabase.FieldLong(rs, "id")
        Set .ManufacturingPlant = plantItem
        .Username = managerDatabase.FieldText(rs, "username")
        .FirstName = managerDatabase.FieldText(rs, "first_name")
        .MiddleName = managerDatabase.FieldText(rs, "middle_name")
        .LastName = managerDatabase.FieldText(rs, "last_name")
        .Jobdesk = managerDatabase.FieldText(rs, "jobdesk")
        .Tier = managerDatabase.FieldLong(rs, "tier")
    End With

    Set MapUser = userItem

End Function
