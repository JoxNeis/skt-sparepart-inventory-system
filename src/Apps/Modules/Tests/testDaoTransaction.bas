Attribute VB_Name = "testDaoTransaction"
Option Explicit

Private Const MODULE_NAME As String = "TRANSACTION_TEST"

' =========================================================
' ERROR CODES (16500 - 16599)
' Usage: ErrorManager.ThrowError testDaoTransaction.TransactionRollbackFailed, ...
' =========================================================
Public Const EmptyCreatedTransactionUUID As Long = 16501        ' CreateTransaction returned no UUID
Public Const NotFoundTransactionBeforeRollback As Long = 16502  ' inserted row not visible before rollback
Public Const TransactionRollbackFailed As Long = 16503          ' row still exists after rollback

Public Sub TestCreateTransactionRollback()

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim UUID As String
    Dim existsBeforeRollback As Boolean
    Dim existsAfterRollback As Boolean
    Dim transactionStarted As Boolean
    Dim errInfo As ErrorInfo

    Set conn = DatabaseManager.GetConnection()

    If conn.State <> adStateOpen Then
        ErrorManager.ThrowError _
            DatabaseManager.ConnectionNotOpen, _
            MODULE_NAME & ".TestCreateTransactionRollback", _
            "Database connection is not open."
    End If

    Debug.Print String$(60, "=")
    Debug.Print "TEST: CreateTransaction rollback"
    Debug.Print "Started at: " & Format$(Now, "yyyy-mm-dd hh:nn:ss")

    '==========================================================
    ' Begin transaction
    '==========================================================
    DatabaseManager.BeginTransaction conn
    transactionStarted = True

    '==========================================================
    ' Execute CreateTransaction
    '==========================================================
    UUID = serviceTransaction.CreateTransaction( _
        conn:=conn, _
        transactionType:="ORDER", _
        transactionDate:=Now, _
        isActive:=True)

    Debug.Print "Generated UUID: " & UUID

    If ValidationText.IsBlank(UUID) Then
        ErrorManager.ThrowError _
            testDaoTransaction.EmptyCreatedTransactionUUID, _
            MODULE_NAME & ".TestCreateTransactionRollback", _
            "FAILED: CreateTransaction returned an empty UUID."
    End If

    '==========================================================
    ' Confirm the row exists before rollback
    '==========================================================
    existsBeforeRollback = TransactionExists( _
        UUID:=UUID, _
        conn:=conn)

    Debug.Print "Exists before rollback: " & existsBeforeRollback

    If Not existsBeforeRollback Then
        ErrorManager.ThrowError _
            testDaoTransaction.NotFoundTransactionBeforeRollback, _
            MODULE_NAME & ".TestCreateTransactionRollback", _
            "FAILED: The transaction was not found before rollback."
    End If

    '==========================================================
    ' Roll back deliberately
    '==========================================================
    DatabaseManager.RollbackTransaction conn
    transactionStarted = False

    '==========================================================
    ' Confirm rollback removed the inserted row
    '==========================================================
    existsAfterRollback = TransactionExists( _
        UUID:=UUID, _
        conn:=conn)

    Debug.Print "Exists after rollback: " & existsAfterRollback

    If existsAfterRollback Then
        ErrorManager.ThrowError _
            testDaoTransaction.TransactionRollbackFailed, _
            MODULE_NAME & ".TestCreateTransactionRollback", _
            "FAILED: The transaction still exists after rollback."
    End If

    Debug.Print "RESULT: PASSED"
    Debug.Print String$(60, "=")

CleanUp:
    ' The shared connection stays open (DatabaseManager owns it).
    If transactionStarted Then
        DatabaseManager.RollbackTransaction conn
    End If

    Set conn = Nothing
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    Debug.Print "RESULT: FAILED"
    Debug.Print "Error number: " & errInfo.Number
    Debug.Print "Error source: " & errInfo.Source
    Debug.Print "Description: " & errInfo.Description
    Debug.Print String$(60, "=")

    If transactionStarted Then
        DatabaseManager.RollbackTransaction conn
    End If

    Set conn = Nothing

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".TestCreateTransactionRollback"

End Sub

Private Function TransactionExists( _
    ByVal UUID As String, _
    ByVal conn As ADODB.connection) As Boolean

    On Error GoTo ErrorHandler

    Dim cmd As ADODB.Command
    Dim rs As ADODB.recordset
    Dim errInfo As ErrorInfo

    TransactionExists = False

    ' Plain SQL text (not a saved query), so switch the
    ' command type after DatabaseManager created it.
    Set cmd = DatabaseManager.CreateCommand(conn, _
        "SELECT COUNT(*) AS row_count " & _
        "FROM transactions " & _
        "WHERE uuid = ?;")

    cmd.CommandType = adCmdText

    DatabaseManager.AddTextParameter cmd, "p_uuid", UUID

    Set rs = DatabaseManager.ExecuteRecordset(cmd)

    If Not rs.EOF Then
        If Not DatabaseManager.FieldIsNull(rs, "row_count") Then
            TransactionExists = _
                (DatabaseManager.FieldLong(rs, "row_count") > 0)
        End If
    End If

CleanUp:
    DatabaseManager.CloseRecordset rs
    Set cmd = Nothing
    Exit Function

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    DatabaseManager.CloseRecordset rs
    Set cmd = Nothing

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".TransactionExists"

End Function
