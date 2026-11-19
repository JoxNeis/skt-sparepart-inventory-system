Attribute VB_Name = "daoTransactionRequestUpdate"
Option Explicit

' =========================================================
' DAO - TRANSACTION REQUEST UPDATE
'
' One function = one saved Access query.
' Logic lives in serviceTransactionRequestUpdate.
' =========================================================

Private Const MODULE_NAME As String = "DAO_TRANSACTION_REQUEST_UPDATE"


' create_transaction_update_request
' Returns the affected row count.
Public Function CreateTransactionRequestUpdate( _
    ByVal conn As ADODB.connection, _
    ByVal originalTransactionUUID As String, _
    ByVal proposedTransactionUUID As String, _
    ByVal requestedBy As String) As Long

    Dim cmd As ADODB.Command

    Set cmd = DatabaseManager.CreateCommand(conn, "create_transaction_update_request")

    DatabaseManager.AddTextParameter cmd, "p_original_transaction_uuid", originalTransactionUUID
    DatabaseManager.AddTextParameter cmd, "p_proposed_transaction_uuid", proposedTransactionUUID
    DatabaseManager.AddTextParameter cmd, "p_requested_by", requestedBy

    CreateTransactionRequestUpdate = DatabaseManager.ExecuteNonQuery(cmd)

End Function


' update_transaction_request_update_approved
' Returns the affected row count (0 when the request is not
' found, already processed, or the user tier is too low).
Public Function ApproveRequestUpdate( _
    ByVal conn As ADODB.connection, _
    ByVal originalTransactionUUID As String, _
    ByVal proposedTransactionUUID As String, _
    ByVal user As String) As Long

    Dim cmd As ADODB.Command

    Set cmd = DatabaseManager.CreateCommand(conn, "update_transaction_request_update_approved")

    DatabaseManager.AddTextParameter cmd, "p_original_transaction_uuid", originalTransactionUUID
    DatabaseManager.AddTextParameter cmd, "p_proposed_transaction_uuid", proposedTransactionUUID
    DatabaseManager.AddTextParameter cmd, "p_user", user

    ApproveRequestUpdate = DatabaseManager.ExecuteNonQuery(cmd)

End Function
