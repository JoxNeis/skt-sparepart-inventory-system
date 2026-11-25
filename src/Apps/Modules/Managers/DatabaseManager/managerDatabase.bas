Attribute VB_Name = "managerDatabase"
Option Explicit

' =========================================================
' DATABASE MANAGER
'
' Connection, transaction and query helpers.
'
' LAYERING
'   Sheet    -> Service -> DAO -> managerDatabase
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
'   Set conn = managerDatabase.GetConnection()
'   managerDatabase.BeginTransaction conn
'   inTransaction = True
'
'   affected = daoOrder.CreateOrder(conn, ...)
'   ...
'
'   managerDatabase.CommitTransaction conn
'   inTransaction = False
'   Exit Sub
'
' ErrorHandler:
'   errInfo = managerError.CaptureError()
'   If inTransaction Then managerDatabase.RollbackTransaction conn
'   managerError.Rethrow errInfo, MODULE_NAME & ".CreateOrder", "..."
'
' DAO TEMPLATE
'
'   Public Function CreateOrder( _
'       ByVal conn As ADODB.connection, ...) As Long
'
'       Dim cmd As ADODB.Command
'       Set cmd = managerDatabase.CreateCommand(conn, "create_order")
'       managerDatabase.AddTextParameter cmd, "p_uuid", uuid
'       CreateOrder = managerDatabase.ExecuteNonQuery(cmd)
'   End Function
' =========================================================

Private Const MODULE_NAME As String = "managerDatabase"

' =========================================================
' ERROR CODES (10100 - 10199)
' Usage: managerError.ThrowError managerDatabase.ConnectionFailed, ...
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
        Set mConnection = OpenAccessConnection(coreConfig.databasePath())
    ElseIf mConnection.State <> adStateOpen Then
        Set mConnection = Nothing
        Set mConnection = OpenAccessConnection(coreConfig.databasePath())
    End If

    If mConnection Is Nothing Then
        managerError.ThrowError _
            managerDatabase.ConnectionFailed, _
            MODULE_NAME & ".GetConnection", _
            "Koneksi dengan database gagal dibuat."
    End If

    Set GetConnection = mConnection

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".GetConnection", _
        "Gagal mendapatkan koneksi database.", _
        managerDatabase.ConnectionFailed

End Function


' Raise when the connection is Nothing or not open.
Public Sub RequireConnection( _
    ByVal conn As ADODB.connection, _
    ByVal errorSource As String)

    If conn Is Nothing Then
        managerError.ThrowError _
            managerDatabase.EmptyConnection, _
            errorSource, _
            "Koneksi database tidak boleh kosong."
    End If

    If conn.State <> adStateOpen Then
        managerError.ThrowError _
            managerDatabase.ConnectionNotOpen, _
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
        managerError.ThrowError _
            managerDatabase.DatabaseFileNotFound, _
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
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".OpenAccessConnection", _
        "Gagal membuat koneksi database.", _
        managerDatabase.ConnectionFailed

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
' Capture the error with managerError.CaptureError() BEFORE
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

    If managerValidationText.IsBlank(queryName) Then
        managerError.ThrowError _
            managerDatabase.EmptyQueryName, _
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


' Optional Short Text column: blank text is sent as NULL.
Public Sub AddNullableTextParameter( _
    ByVal cmd As ADODB.Command, _
    ByVal parameterName As String, _
    ByVal value As String, _
    Optional ByVal size As Long = TEXT_SIZE)

    Dim prm As ADODB.Parameter

    RequireCommand cmd, MODULE_NAME & ".AddNullableTextParameter"

    value = Trim$(value)

    Set prm = cmd.CreateParameter( _
        parameterName, adVarWChar, adParamInput, size)

    If Len(value) = 0 Then
        prm.value = Null
    Else
        prm.value = value
    End If

    cmd.Parameters.Append prm

End Sub


' Optional Long Text (memo) column: blank text is sent as NULL.
Public Sub AddNullableLongTextParameter( _
    ByVal cmd As ADODB.Command, _
    ByVal parameterName As String, _
    ByVal value As String)

    Dim prm As ADODB.Parameter
    Dim size As Long

    RequireCommand cmd, MODULE_NAME & ".AddNullableLongTextParameter"

    value = Trim$(value)

    size = Len(value)
    If size = 0 Then size = 1

    Set prm = cmd.CreateParameter( _
        parameterName, adLongVarWChar, adParamInput, size)

    If Len(value) = 0 Then
        prm.value = Null
    Else
        prm.value = value
    End If

    cmd.Parameters.Append prm

End Sub


' Optional Long Integer column (for example a nullable
' foreign key). Null, Empty or 0 is sent as NULL.
Public Sub AddNullableLongParameter( _
    ByVal cmd As ADODB.Command, _
    ByVal parameterName As String, _
    ByVal value As Variant)

    Dim prm As ADODB.Parameter

    RequireCommand cmd, MODULE_NAME & ".AddNullableLongParameter"

    Set prm = cmd.CreateParameter( _
        parameterName, adInteger, adParamInput)

    If IsNull(value) Or IsEmpty(value) Then
        prm.value = Null
    ElseIf CLng(value) = 0 Then
        prm.value = Null
    Else
        prm.value = CLng(value)
    End If

    cmd.Parameters.Append prm

End Sub


' =========================================================
' PAGED LIST QUERIES (see PageRequest)
'
'   Set cmd = managerDatabase.CreateCommand(conn, "order_list")
'   ...criteria parameters...
'   managerDatabase.AddPageFilterParameters cmd, page
'   managerDatabase.AddPageDateKeyParameters cmd, page
'   Set rs = managerDatabase.ExecuteRecordset(cmd)
'   managerDatabase.MoveToPageStart rs, page
'   Do While Not rs.EOF
'       If managerDatabase.IsPageFull(items.Count, page) Then Exit Do
'       ...
' =========================================================

' p_filter_index (when withIndex) and p_filter_value.
Public Sub AddPageFilterParameters( _
    ByVal cmd As ADODB.Command, _
    ByVal page As PageRequest, _
    Optional ByVal withIndex As Boolean = True)

    If withIndex Then
        AddLongParameter cmd, "p_filter_index", page.FilterIndex
    End If

    AddNullableTextParameter cmd, "p_filter_value", page.FilterValue

End Sub


' p_direction, p_key_date, p_key_id.
Public Sub AddPageDateKeyParameters( _
    ByVal cmd As ADODB.Command, _
    ByVal page As PageRequest)

    AddTextParameter cmd, "p_direction", page.Direction, 4
    AddDateParameter cmd, "p_key_date", page.KeyDate
    AddLongParameter cmd, "p_key_id", page.KeyId

End Sub


' Put the recordset on the first row of the page: the first
' row for NEXT, the first of the last PageSize rows for PREV
' (the query returns every row before the key, ascending).
Public Sub MoveToPageStart( _
    ByVal rs As ADODB.recordset, _
    ByVal page As PageRequest)

    If rs.EOF And rs.BOF Then Exit Sub

    rs.MoveFirst

    If page.PageSize <= 0 Then Exit Sub
    If Not page.IsPrevious Then Exit Sub

    If rs.RecordCount > page.PageSize Then
        rs.Move rs.RecordCount - page.PageSize
    End If

End Sub


' True when rowsRead rows already fill the page.
Public Function IsPageFull( _
    ByVal rowsRead As Long, _
    ByVal page As PageRequest) As Boolean

    If page.PageSize <= 0 Then
        IsPageFull = False
    Else
        IsPageFull = (rowsRead >= page.PageSize)
    End If

End Function


' total_rows of a *_count query.
Public Function ExecuteCount( _
    ByVal cmd As ADODB.Command) As Long

    Dim rs As ADODB.recordset

    Set rs = ExecuteRecordset(cmd)

    If rs.EOF Then
        ExecuteCount = 0
    Else
        ExecuteCount = FieldLong(rs, "total_rows")
    End If

    CloseRecordset rs

End Function


' Raise errorCode when an INSERT / UPDATE affected fewer
' rows than expected. Used by services after every write.
Public Sub RequireAffectedRows( _
    ByVal affected As Long, _
    ByVal expected As Long, _
    ByVal errorCode As Long, _
    ByVal errorSource As String, _
    ByVal errorMessage As String)

    If affected < expected Then
        managerError.ThrowError errorCode, errorSource, errorMessage
    End If

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
        managerError.ThrowError _
            managerDatabase.RecordsetNotOpen, _
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


' Raise managerDatabase.FieldNotFound for the first missing
' field in requiredFields (an Array).
Public Sub ValidateFields( _
    ByVal rs As ADODB.recordset, _
    ByVal requiredFields As Variant)

    Dim fieldName As Variant

    RequireRecordset rs, MODULE_NAME & ".ValidateFields"

    For Each fieldName In requiredFields
        If Not HasField(rs, CStr(fieldName)) Then
            managerError.ThrowError _
                managerDatabase.FieldNotFound, _
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
        managerError.ThrowError _
            managerDatabase.FieldNotFound, _
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


' Null -> 0. Non-numeric -> managerDatabase.InvalidFieldNumber.
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
        managerError.ThrowError _
            managerDatabase.InvalidFieldNumber, _
            MODULE_NAME & ".FieldDouble", _
            "Field '" & fieldName & "' bukan angka yang valid."
    End If

    FieldDouble = CDbl(value)

End Function


' Null -> 0. Non-numeric -> managerDatabase.InvalidFieldNumber.
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
        managerError.ThrowError _
            managerDatabase.InvalidFieldNumber, _
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
        managerError.ThrowError _
            managerDatabase.InvalidFieldDate, _
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
        managerError.ThrowError _
            managerDatabase.EmptyCommand, _
            errorSource, _
            "Command database tidak boleh kosong."
    End If

End Sub


Private Sub RequireRecordset( _
    ByVal rs As ADODB.recordset, _
    ByVal errorSource As String)

    If rs Is Nothing Then
        managerError.ThrowError _
            managerDatabase.EmptyRecordset, _
            errorSource, _
            "Recordset tidak boleh kosong."
    End If

    If rs.State <> adStateOpen Then
        managerError.ThrowError _
            managerDatabase.RecordsetNotOpen, _
            errorSource, _
            "Recordset tidak terbuka."
    End If

End Sub
