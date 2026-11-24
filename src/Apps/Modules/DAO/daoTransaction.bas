Attribute VB_Name = "daoTransaction"
Option Explicit

' =========================================================
' DAO - TRANSACTION
'
' One function = one saved Access query.
' The connection (and its running transaction) is owned by
' the calling service. Validation and row-count checks are
' done by serviceTransaction.
' =========================================================

Private Const MODULE_NAME As String = "DAO_TRANSACTION"


' select_uuid
' Returns the generated UUID, or "" when the query returns
' no row / Null.
Public Function SelectUUID( _
    ByVal conn As ADODB.connection, _
    ByVal transactionCode As String, _
    ByVal transactionDate As Date) As String

    Dim cmd As ADODB.Command
    Dim rs As ADODB.recordset

    Set cmd = managerDatabase.CreateCommand(conn, "select_uuid")

    managerDatabase.AddTextParameter cmd, "p_transaction_type", transactionCode
    managerDatabase.AddDateParameter cmd, "p_transaction_date", transactionDate

    Set rs = managerDatabase.ExecuteRecordset(cmd)

    If rs.EOF Then
        SelectUUID = vbNullString
    Else
        SelectUUID = managerDatabase.FieldText(rs, "generated_uuid")
    End If

    managerDatabase.CloseRecordset rs

End Function


' create_transaction
' activeAt: Date, or Null for an inactive (proposed) transaction.
' Returns the affected row count.
Public Function CreateTransaction( _
    ByVal conn As ADODB.connection, _
    ByVal transactionUUID As String, _
    ByVal transactionType As String, _
    ByVal createdBy As String, _
    ByVal activeAt As Variant) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "create_transaction")

    managerDatabase.AddTextParameter cmd, "p_uuid", transactionUUID
    managerDatabase.AddTextParameter cmd, "p_transaction_type", transactionType
    managerDatabase.AddTextParameter cmd, "p_created_by", createdBy
    managerDatabase.AddDateParameter cmd, "p_active_at", activeAt

    CreateTransaction = managerDatabase.ExecuteNonQuery(cmd)

End Function


' update_transaction_reversal
' Returns the affected row count.
Public Function ReverseTransaction( _
    ByVal conn As ADODB.connection, _
    ByVal originalTransactionUUID As String) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "update_transaction_reversal")

    managerDatabase.AddTextParameter cmd, "p_original_transaction_uuid", originalTransactionUUID

    ReverseTransaction = managerDatabase.ExecuteNonQuery(cmd)

End Function


' update_transaction_activation
' Returns the affected row count.
Public Function ActivateTransaction( _
    ByVal conn As ADODB.connection, _
    ByVal transactionUUID As String) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "update_transaction_activation")

    managerDatabase.AddTextParameter cmd, "p_transaction_uuid", transactionUUID

    ActivateTransaction = managerDatabase.ExecuteNonQuery(cmd)

End Function
