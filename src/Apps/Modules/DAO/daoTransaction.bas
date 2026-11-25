Attribute VB_Name = "daoTransaction"
Option Explicit

' =========================================================
' DAO - TRANSACTION (document header)
' Table: transactions
'
' One function = one saved Access query.
' The connection (and its running transaction) is owned by
' the calling service. Row-count checks are done by
' serviceTransaction.
' =========================================================

Private Const MODULE_NAME As String = "daoTransaction"


' transaction_get_next_uuid
' <PLANT>-<PREFIX>-<yyyymmdd>-<nnn>. "" when the day already
' has 999 documents (the query returns Null).
Public Function GetNextUUID( _
    ByVal conn As ADODB.connection, _
    ByVal plantCode As String, _
    ByVal uuidPrefix As String, _
    ByVal transactionDate As Date) As String

    Dim cmd As ADODB.Command
    Dim rs As ADODB.recordset

    Set cmd = managerDatabase.CreateCommand(conn, "transaction_get_next_uuid")

    managerDatabase.AddTextParameter cmd, "p_manufacturing_plant_code", plantCode, 15
    managerDatabase.AddTextParameter cmd, "p_uuid_prefix", uuidPrefix, 3
    managerDatabase.AddDateParameter cmd, "p_transaction_date", DateValue(transactionDate)

    Set rs = managerDatabase.ExecuteRecordset(cmd)

    If rs.EOF Then
        GetNextUUID = vbNullString
    Else
        GetNextUUID = managerDatabase.FieldText(rs, "next_uuid")
    End If

    managerDatabase.CloseRecordset rs

End Function


' transaction_create
' activatedAt: a Date for an active document, Null for a
' draft (proposed update).
Public Function CreateTransaction( _
    ByVal conn As ADODB.connection, _
    ByVal transactionUUID As String, _
    ByVal transactionType As String, _
    ByVal plantId As Long, _
    ByVal createdBy As String, _
    ByVal activatedAt As Variant) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "transaction_create")

    managerDatabase.AddTextParameter cmd, "p_uuid", transactionUUID, 36
    managerDatabase.AddTextParameter cmd, "p_transaction_type", transactionType, 15
    managerDatabase.AddLongParameter cmd, "p_manufacturing_plant_id", plantId
    managerDatabase.AddTextParameter cmd, "p_created_by", createdBy, 64
    managerDatabase.AddDateParameter cmd, "p_activated_at", activatedAt

    CreateTransaction = managerDatabase.ExecuteNonQuery(cmd)

End Function


' transaction_activate (draft -> active)
Public Function ActivateTransaction( _
    ByVal conn As ADODB.connection, _
    ByVal transactionUUID As String) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "transaction_activate")

    managerDatabase.AddTextParameter cmd, "p_uuid", transactionUUID, 36

    ActivateTransaction = managerDatabase.ExecuteNonQuery(cmd)

End Function


' transaction_reverse
Public Function ReverseTransaction( _
    ByVal conn As ADODB.connection, _
    ByVal transactionUUID As String, _
    ByVal reversedBy As String) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "transaction_reverse")

    managerDatabase.AddTextParameter cmd, "p_uuid", transactionUUID, 36
    managerDatabase.AddTextParameter cmd, "p_reversed_by", reversedBy, 64

    ReverseTransaction = managerDatabase.ExecuteNonQuery(cmd)

End Function
