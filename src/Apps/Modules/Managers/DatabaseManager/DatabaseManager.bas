Attribute VB_Name = "DatabaseManager"
Option Explicit

' =========================================================
' DATABASE MANAGER
'
' Connection, transaction and query helpers.
'
' LAYERING
'   Sheet    -> Service -> DAO -> DatabaseManager
'
'   Service : gets the connection, owns the transaction,
'             validates, runs the business logic and checks
'             the affected row counts.
'   DAO     : one Public function = one saved Access query.
'             Always receives the connection (and with it
'             the running transaction) as first parameter.
'             Never opens, commits or rolls back.
'
' SERVICE TEMPLATE (write)
'
'   Dim conn As ADODB.connection
'   Dim inTransaction As Boolean
'   Dim errInfo As ErrorInfo
'
'   Set conn = DatabaseManager.GetConnection()
'   DatabaseManager.BeginTransaction conn
'   inTransaction = True
'
'   affected = daoOrder.CreateOrder(conn, ...)
'   ...
'
'   DatabaseManager.CommitTransaction conn
'   inTransaction = False
'   Exit Sub
'
' ErrorHandler:
'   errInfo = ErrorManager.CaptureError()
'   If inTransaction Then DatabaseManager.RollbackTransaction conn
'   ErrorManager.Rethrow errInfo, MODULE_NAME & ".CreateOrder", "..."
'
' DAO TEMPLATE
'
'   Public Function CreateOrder( _
'       ByVal conn As ADODB.connection, ...) As Long
'
'       Dim cmd As ADODB.Command
'       Set cmd = DatabaseManager.CreateCommand(conn, "create_order")
'       DatabaseManager.AddTextParameter cmd, "p_uuid", uuid
'       CreateOrder = DatabaseManager.ExecuteNonQuery(cmd)
'   End Function
' =========================================================

Private Const MODULE_NAME As String = "DatabaseManager"

' =========================================================
' ERROR CODES (10100 - 10199)
' Usage: ErrorManager.ThrowError DatabaseManager.ConnectionFailed, ...
' =========================================================
Public Const ConnectionFailed As Long = 10101       ' connection could not be created
Public Const ConnectionNotOpen As Long = 10102      ' conn.State <> adStateOpen
Public Const RecordsetNotOpen As Long = 10103       ' rs.State <> adStateOpen
Public Const DatabaseFileNotFound As Long = 10104   ' .accdb file does not exist
Public Const EmptyConnection As Long = 10105        ' conn parameter is Nothing
Public Const EmptyQueryName As Long = 10106         ' saved query name is blank
Public Const EmptyCommand As Long = 10107           ' cmd parameter is Nothing
Public Const EmptyRecordset As Long = 10108         ' rs parameter is Nothing
Public Const FieldNotFound As Long = 10109          ' recordset has no such field
Public Const InvalidFieldNumber As Long = 10110     ' field value is not numeric
Public Const InvalidFieldDate As Long = 10111       ' field value is not a date

' Default size of text parameters (Access short text).
Public Const TEXT_SIZE As Long = 255

Private mConnection As ADODB.connection


' =========================================================
' CONNECTION
' =========================================================

' Shared connection to the Access database. Opened on first
' use and opened again when it was closed.
Public Function GetConnection() As ADODB.connection

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    If mConnection Is Nothing Then
        Set mConnection = OpenAccessConnection(pConfig.databasePath())
    ElseIf mConnection.State <> adStateOpen Then
        Set mConnection = Nothing
        Set mConnection = OpenAccessConnection(pConfig.databasePath())
    End If

    If mConnection Is Nothing Then
        ErrorManager.ThrowError _
            DatabaseManager.ConnectionFailed, _
            MODULE_NAME & ".GetConnection", _
            "Koneksi dengan database gagal dibuat."
    End If

    Set GetConnection = mConnection

    Exit Function

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".GetConnection", _
        "Gagal mendapatkan koneksi database.", _
        DatabaseManager.ConnectionFailed

End Function


' Raise when the connection is Nothing or not open.
Public Sub RequireConnection( _
    ByVal conn As ADODB.connection, _
    ByVal errorSource As String)

    If conn Is Nothing Then
        ErrorManager.ThrowError _
            DatabaseManager.EmptyConnection, _
            errorSource, _
            "Koneksi database tidak boleh kosong."
    End If

    If conn.State <> adStateOpen Then
        ErrorManager.ThrowError _
            DatabaseManager.ConnectionNotOpen, _
            errorSource, _
            "Koneksi database belum terbuka."
    End If

End Sub


' Close the shared connection (for example on workbook close).
Public Sub DisconnectFromDatabase()

    On Error Resume Next
    If Not mConnection Is Nothing Then
        If mConnection.State = adStateOpen Then
            mConnection.Close
        End If
        Set mConnection = Nothing
    End If
    On Error GoTo 0

    Debug.Print "Database ditutup."

End Sub


Private Function OpenAccessConnection( _
    ByVal databasePath As String) As ADODB.connection

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim errInfo As ErrorInfo

    If Dir(databasePath) = "" Then
        ErrorManager.ThrowError _
            DatabaseManager.DatabaseFileNotFound, _
            MODULE_NAME & ".OpenAccessConnection", _
            "File database (Microsoft Access) tidak ditemukan: " & databasePath
    End If

    Set conn = New ADODB.connection
    conn.Open _
        "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & databasePath & ";"

    Debug.Print "Terhubung ke database."

    Set OpenAccessConnection = conn

    Exit Function

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".OpenAccessConnection", _
        "Gagal membuat koneksi database.", _
        DatabaseManager.ConnectionFailed

End Function


' =========================================================
' TRANSACTION (services only)
' =========================================================

Public Sub BeginTransaction( _
    ByVal conn As ADODB.connection)

    RequireConnection conn, MODULE_NAME & ".BeginTransaction"

    conn.BeginTrans

End Sub


Public Sub CommitTransaction( _
    ByVal conn As ADODB.connection)

    RequireConnection conn, MODULE_NAME & ".CommitTransaction"

    conn.CommitTrans

End Sub


' Safe to call from an error handler: never raises.
' Capture the error with ErrorManager.CaptureError() BEFORE
' calling this, because On Error clears the Err object.
Public Sub RollbackTransaction( _
    ByVal conn As ADODB.connection)

    On Error Resume Next

    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then
            conn.RollbackTrans
        End If
    End If

    On Error GoTo 0

End Sub


' =========================================================
' COMMAND (DAO only)
' =========================================================

' New command for a saved Access query on the given
' connection (the running transaction applies).
Public Function CreateCommand( _
    ByVal conn As ADODB.connection, _
    ByVal queryName As String) As ADODB.Command

    Dim cmd As ADODB.Command

    RequireConnection conn, MODULE_NAME & ".CreateCommand"

    If ValidationText.IsBlank(queryName) Then
        ErrorManager.ThrowError _
            DatabaseManager.EmptyQueryName, _
            MODULE_NAME & ".CreateCommand", _
            "Nama query tidak boleh kosong."
    End If

    Set cmd = New ADODB.Command

    With cmd
        Set .ActiveConnection = conn
        .CommandType = adCmdStoredProc
        .CommandText = Trim$(queryName)
    End With

    Set CreateCommand = cmd

End Function


' Parameters must be added in the same order as the
' PARAMETERS declaration of the saved Access query.

Public Sub AddTextParameter( _
    ByVal cmd As ADODB.Command, _
    ByVal parameterName As String, _
    ByVal value As String, _
    Optional ByVal size As Long = TEXT_SIZE)

    RequireCommand cmd, MODULE_NAME & ".AddTextParameter"

    cmd.Parameters.Append cmd.CreateParameter( _
        parameterName, adVarWChar, adParamInput, size, value)

End Sub


' Long Text (memo) parameter.
Public Sub AddLongTextParameter( _
    ByVal cmd As ADODB.Command, _
    ByVal parameterName As String, _
    ByVal value As String)

    Dim size As Long

    RequireCommand cmd, MODULE_NAME & ".AddLongTextParameter"

    size = Len(value)
    If size = 0 Then size = 1

    cmd.Parameters.Append cmd.CreateParameter( _
        parameterName, adLongVarWChar, adParamInput, size, value)

End Sub


' value may be Null (for example an inactive transaction).
Public Sub AddDateParameter( _
    ByVal cmd As ADODB.Command, _
    ByVal parameterName As String, _
    ByVal value As Variant)

    Dim prm As ADODB.Parameter

    RequireCommand cmd, MODULE_NAME & ".AddDateParameter"

    Set prm = cmd.CreateParameter( _
        parameterName, adDate, adParamInput)

    If IsNull(value) Or IsEmpty(value) Then
        prm.value = Null
    Else
        prm.value = CDate(value)
    End If

    cmd.Parameters.Append prm

End Sub


Public Sub AddDoubleParameter( _
    ByVal cmd As ADODB.Command, _
    ByVal parameterName As String, _
    ByVal value As Double)

    RequireCommand cmd, MODULE_NAME & ".AddDoubleParameter"

    cmd.Parameters.Append cmd.CreateParameter( _
        parameterName, adDouble, adParamInput, , value)

End Sub


' Long Integer (adInteger).
Public Sub AddLongParameter( _
    ByVal cmd As ADODB.Command, _
    ByVal parameterName As String, _
    ByVal value As Long)

    RequireCommand cmd, MODULE_NAME & ".AddLongParameter"

    cmd.Parameters.Append cmd.CreateParameter( _
        parameterName, adInteger, adParamInput, , value)

End Sub


' Integer (adSmallInt).
Public Sub AddIntegerParameter( _
    ByVal cmd As ADODB.Command, _
    ByVal parameterName As String, _
    ByVal value As Integer)

    RequireCommand cmd, MODULE_NAME & ".AddIntegerParameter"

    cmd.Parameters.Append cmd.CreateParameter( _
        parameterName, adSmallInt, adParamInput, , value)

End Sub


Public Sub AddBooleanParameter( _
    ByVal cmd As ADODB.Command, _
    ByVal parameterName As String, _
    ByVal value As Boolean)

    RequireCommand cmd, MODULE_NAME & ".AddBooleanParameter"

    cmd.Parameters.Append cmd.CreateParameter( _
        parameterName, adBoolean, adParamInput, , value)

End Sub


' INSERT / UPDATE / DELETE. Returns the affected row count.
' The caller (service) decides whether the count is valid.
Public Function ExecuteNonQuery( _
    ByVal cmd As ADODB.Command) As Long

    Dim affected As Long

    RequireCommand cmd, MODULE_NAME & ".ExecuteNonQuery"

    affected = 0

    cmd.Execute _
        recordsAffected:=affected, _
        Options:=adExecuteNoRecords

    ExecuteNonQuery = affected

End Function


' SELECT. Returns an open, read-only, client-side and
' disconnected recordset: all rows are already loaded, so
' the caller (service or sheet) can keep reading it after
' the transaction ends. Close it with CloseRecordset.
Public Function ExecuteRecordset( _
    ByVal cmd As ADODB.Command) As ADODB.recordset

    Dim rs As ADODB.recordset

    RequireCommand cmd, MODULE_NAME & ".ExecuteRecordset"

    Set rs = New ADODB.recordset

    With rs
        .CursorLocation = adUseClient
        .Open _
            source:=cmd, _
            CursorType:=adOpenStatic, _
            LockType:=adLockReadOnly
    End With

    If rs.State <> adStateOpen Then
        ErrorManager.ThrowError _
            DatabaseManager.RecordsetNotOpen, _
            MODULE_NAME & ".ExecuteRecordset", _
            "Recordset query '" & cmd.CommandText & "' tidak terbuka."
    End If

    Set rs.ActiveConnection = Nothing

    Set ExecuteRecordset = rs

End Function


Public Sub CloseRecordset( _
    ByRef rs As ADODB.recordset)

    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then
            rs.Close
        End If
        Set rs = Nothing
    End If
    On Error GoTo 0

End Sub


' =========================================================
' RECORDSET FIELDS
' Used by DAOs, and by sheets that read a disconnected
' recordset returned by a service.
' =========================================================

Public Function HasField( _
    ByVal rs As ADODB.recordset, _
    ByVal fieldName As String) As Boolean

    Dim fld As ADODB.Field

    HasField = False

    If rs Is Nothing Then Exit Function

    For Each fld In rs.Fields
        If StrComp(fld.Name, fieldName, vbTextCompare) = 0 Then
            HasField = True
            Exit Function
        End If
    Next fld

End Function


' Raise DatabaseManager.FieldNotFound for the first missing
' field in requiredFields (an Array).
Public Sub ValidateFields( _
    ByVal rs As ADODB.recordset, _
    ByVal requiredFields As Variant)

    Dim fieldName As Variant

    RequireRecordset rs, MODULE_NAME & ".ValidateFields"

    For Each fieldName In requiredFields
        If Not HasField(rs, CStr(fieldName)) Then
            ErrorManager.ThrowError _
                DatabaseManager.FieldNotFound, _
                MODULE_NAME & ".ValidateFields", _
                "Field '" & CStr(fieldName) & _
                "' tidak ditemukan pada hasil query."
        End If
    Next fieldName

End Sub


Public Function FieldValue( _
    ByVal rs As ADODB.recordset, _
    ByVal fieldName As String) As Variant

    RequireRecordset rs, MODULE_NAME & ".FieldValue"

    If Not HasField(rs, fieldName) Then
        ErrorManager.ThrowError _
            DatabaseManager.FieldNotFound, _
            MODULE_NAME & ".FieldValue", _
            "Field '" & fieldName & "' tidak ditemukan pada hasil query."
    End If

    FieldValue = rs.Fields(fieldName).value

End Function


Public Function FieldIsNull( _
    ByVal rs As ADODB.recordset, _
    ByVal fieldName As String) As Boolean

    Dim value As Variant

    value = FieldValue(rs, fieldName)

    FieldIsNull = (IsNull(value) Or IsEmpty(value))

End Function


' Trimmed text. Null -> "".
Public Function FieldText( _
    ByVal rs As ADODB.recordset, _
    ByVal fieldName As String) As String

    Dim value As Variant

    value = FieldValue(rs, fieldName)

    If IsNull(value) Or IsEmpty(value) Then
        FieldText = vbNullString
    Else
        FieldText = Trim$(CStr(value))
    End If

End Function


' Null -> 0. Non-numeric -> DatabaseManager.InvalidFieldNumber.
Public Function FieldDouble( _
    ByVal rs As ADODB.recordset, _
    ByVal fieldName As String) As Double

    Dim value As Variant

    value = FieldValue(rs, fieldName)

    If IsNull(value) Or IsEmpty(value) Then
        FieldDouble = 0
        Exit Function
    End If

    If Not IsNumeric(value) Then
        ErrorManager.ThrowError _
            DatabaseManager.InvalidFieldNumber, _
            MODULE_NAME & ".FieldDouble", _
            "Field '" & fieldName & "' bukan angka yang valid."
    End If

    FieldDouble = CDbl(value)

End Function


' Null -> 0. Non-numeric -> DatabaseManager.InvalidFieldNumber.
Public Function FieldLong( _
    ByVal rs As ADODB.recordset, _
    ByVal fieldName As String) As Long

    Dim value As Variant

    value = FieldValue(rs, fieldName)

    If IsNull(value) Or IsEmpty(value) Then
        FieldLong = 0
        Exit Function
    End If

    If Not IsNumeric(value) Then
        ErrorManager.ThrowError _
            DatabaseManager.InvalidFieldNumber, _
            MODULE_NAME & ".FieldLong", _
            "Field '" & fieldName & "' bukan angka yang valid."
    End If

    FieldLong = CLng(value)

End Function


' Null -> 0 (empty date). Not a date -> InvalidFieldDate.
Public Function FieldDate( _
    ByVal rs As ADODB.recordset, _
    ByVal fieldName As String) As Date

    Dim value As Variant

    value = FieldValue(rs, fieldName)

    If IsNull(value) Or IsEmpty(value) Then
        FieldDate = 0
        Exit Function
    End If

    If Not IsDate(value) Then
        ErrorManager.ThrowError _
            DatabaseManager.InvalidFieldDate, _
            MODULE_NAME & ".FieldDate", _
            "Field '" & fieldName & "' bukan tanggal yang valid."
    End If

    FieldDate = CDate(value)

End Function


' Null -> False.
Public Function FieldBoolean( _
    ByVal rs As ADODB.recordset, _
    ByVal fieldName As String) As Boolean

    Dim value As Variant

    value = FieldValue(rs, fieldName)

    If IsNull(value) Or IsEmpty(value) Then
        FieldBoolean = False
    Else
        FieldBoolean = CBool(value)
    End If

End Function


' =========================================================
' INTERNAL
' =========================================================

Private Sub RequireCommand( _
    ByVal cmd As ADODB.Command, _
    ByVal errorSource As String)

    If cmd Is Nothing Then
        ErrorManager.ThrowError _
            DatabaseManager.EmptyCommand, _
            errorSource, _
            "Command database tidak boleh kosong."
    End If

End Sub


Private Sub RequireRecordset( _
    ByVal rs As ADODB.recordset, _
    ByVal errorSource As String)

    If rs Is Nothing Then
        ErrorManager.ThrowError _
            DatabaseManager.EmptyRecordset, _
            errorSource, _
            "Recordset tidak boleh kosong."
    End If

    If rs.State <> adStateOpen Then
        ErrorManager.ThrowError _
            DatabaseManager.RecordsetNotOpen, _
            errorSource, _
            "Recordset tidak terbuka."
    End If

End Sub
