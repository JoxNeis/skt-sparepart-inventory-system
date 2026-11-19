Attribute VB_Name = "serviceUser"
Option Explicit

' =========================================================
' SERVICE - USER
'
' Public entry points (called by sheets) get the
' connection and own the transaction:
'   SelectUsers, CreateUser, CompareUser
' =========================================================

Private Const MODULE_NAME As String = "SERVICE_USER"

' =========================================================
' ERROR CODES (18100 - 18199)
' Usage: ErrorManager.ThrowError serviceUser.EmptyUsername, ...
' =========================================================
Public Const EmptyPlantName As Long = 18101                 ' manufacturing plant name is blank
Public Const EmptyUsername As Long = 18102
Public Const EmptyFirstName As Long = 18103
Public Const EmptyLastName As Long = 18104
Public Const EmptyPosition As Long = 18105
Public Const NegativeTier As Long = 18106                   ' tier <= 0
Public Const EmptyCurrentUser As Long = 18107               ' creating user cannot be determined
Public Const NoRowAffectedCreateUser As Long = 18108        ' 0 rows inserted
Public Const MultipleRowsAffectedCreateUser As Long = 18109 ' > 1 row inserted
Public Const EmptyRequestorUsername As Long = 18110
Public Const EmptyApproverUsername As Long = 18111
Public Const SameRequestorApprover As Long = 18112
Public Const NotFoundRequestorOrApprover As Long = 18113    ' compare_user returned no row
Public Const InvalidCanApproveErrorValue As Long = 18114    ' can_approve contains an error value
Public Const EmptyCanApproveResult As Long = 18115          ' can_approve is Null / Empty
Public Const InvalidCanApproveConversion As Long = 18116    ' can_approve cannot be converted to Boolean


' =========================================================
' PUBLIC ENTRY POINTS
' =========================================================

' Every user. Read only, no transaction.
' Returns a disconnected recordset (close it after use).
Public Function SelectUsers() As ADODB.recordset

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim errInfo As ErrorInfo

    Set conn = DatabaseManager.GetConnection()

    Set SelectUsers = daoUser.SelectUsers(conn)

    Exit Function

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".SelectUsers", _
        "Gagal mengambil daftar user."

End Function


' Insert one user and write the activity log, in one
' transaction. Middle name is optional.
Public Sub CreateUser( _
    ByVal manufacturingPlantName As String, _
    ByVal username As String, _
    ByVal firstName As String, _
    ByVal middleName As String, _
    ByVal lastName As String, _
    ByVal position As String, _
    ByVal tier As Long)

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim affected As Long
    Dim inTransaction As Boolean

    Dim currentUser As String
    Dim activityDetail As String

    Dim errInfo As ErrorInfo

    inTransaction = False

    ' =====================================================
    ' Step 1: Normalize input parameters
    ' =====================================================
    manufacturingPlantName = Trim$(manufacturingPlantName)
    username = Trim$(username)
    firstName = Trim$(firstName)
    middleName = Trim$(middleName)
    lastName = Trim$(lastName)
    position = Trim$(position)

    ' =====================================================
    ' Step 2: Validate input parameters
    ' =====================================================
    If ValidationText.IsBlank(manufacturingPlantName) Then
        ErrorManager.ThrowError _
            serviceUser.EmptyPlantName, _
            MODULE_NAME & ".CreateUser", _
            "Nama manufacturing plant " & _
            "tidak boleh kosong."
    End If

    If ValidationText.IsBlank(username) Then
        ErrorManager.ThrowError _
            serviceUser.EmptyUsername, _
            MODULE_NAME & ".CreateUser", _
            "Username PMI tidak boleh kosong."
    End If

    If ValidationText.IsBlank(firstName) Then
        ErrorManager.ThrowError _
            serviceUser.EmptyFirstName, _
            MODULE_NAME & ".CreateUser", _
            "Nama depan tidak boleh kosong."
    End If

    ' Middle name is optional, so an empty value is valid.

    If ValidationText.IsBlank(lastName) Then
        ErrorManager.ThrowError _
            serviceUser.EmptyLastName, _
            MODULE_NAME & ".CreateUser", _
            "Nama belakang tidak boleh kosong."
    End If

    If ValidationText.IsBlank(position) Then
        ErrorManager.ThrowError _
            serviceUser.EmptyPosition, _
            MODULE_NAME & ".CreateUser", _
            "Posisi user tidak boleh kosong."
    End If

    If ValidationNumeric.IsZeroOrNegative(tier) Then
        ErrorManager.ThrowError _
            serviceUser.NegativeTier, _
            MODULE_NAME & ".CreateUser", _
            "Tier user harus lebih besar dari nol."
    End If

    ' =====================================================
    ' Step 3: Determine the current application user
    ' =====================================================
    currentUser = Trim$(pConfig.currentUser())

    If ValidationText.IsBlank(currentUser) Then
        ErrorManager.ThrowError _
            serviceUser.EmptyCurrentUser, _
            MODULE_NAME & ".CreateUser", _
            "User yang membuat data user baru " & _
            "tidak dapat ditentukan."
    End If

    ' =====================================================
    ' Step 4: Begin transaction
    ' =====================================================
    Set conn = DatabaseManager.GetConnection()

    DatabaseManager.BeginTransaction conn
    inTransaction = True

    ' =====================================================
    ' Step 5: Insert user
    ' =====================================================
    affected = daoUser.CreateUser( _
        conn, manufacturingPlantName, username, firstName, _
        middleName, lastName, position, tier)

    ' =====================================================
    ' Step 6: Validate affected rows
    ' =====================================================
    If ValidationNumeric.IsLessThan(affected, 1) Then
        ErrorManager.ThrowError _
            serviceUser.NoRowAffectedCreateUser, _
            MODULE_NAME & ".CreateUser", _
            "User tidak berhasil dibuat." & _
            vbCrLf & _
            "Manufacturing plant mungkin tidak " & _
            "ditemukan atau username sudah digunakan."
    End If

    If ValidationNumeric.IsGreaterThan(affected, 1) Then
        ErrorManager.ThrowError _
            serviceUser.MultipleRowsAffectedCreateUser, _
            MODULE_NAME & ".CreateUser", _
            "Pembuatan user memengaruhi lebih dari " & _
            "satu record. Transaksi dibatalkan."
    End If

    ' =====================================================
    ' Step 7: Build activity-log detail
    ' =====================================================
    activityDetail = BuildUserCreateDetail( _
        manufacturingPlantName:=manufacturingPlantName, _
        username:=username, _
        firstName:=firstName, _
        middleName:=middleName, _
        lastName:=lastName, _
        position:=position, _
        tier:=tier)

    ' =====================================================
    ' Step 8: Write activity log
    ' =====================================================
    serviceActivityLog.WriteActivityLog _
        conn, _
        "User.Create", _
        activityDetail

    ' =====================================================
    ' Step 9: Commit
    ' =====================================================
    DatabaseManager.CommitTransaction conn
    inTransaction = False

    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    If inTransaction Then
        DatabaseManager.RollbackTransaction conn
    End If

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".CreateUser", _
        "Gagal membuat data user."

End Sub


' True when the approver may approve the requestor's
' request (compare_user.can_approve). Read only, no
' transaction.
Public Function CompareUser( _
    ByVal requestorUsername As String, _
    ByVal approverUsername As String) As Boolean

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim rs As ADODB.recordset
    Dim canApproveValue As Variant
    Dim errInfo As ErrorInfo

    requestorUsername = Trim$(requestorUsername)
    approverUsername = Trim$(approverUsername)

    ' Default result.
    CompareUser = False

    ' =====================================================
    ' Step 1: Validate parameters
    ' =====================================================
    If ValidationText.IsBlank(requestorUsername) Then
        ErrorManager.ThrowError _
            serviceUser.EmptyRequestorUsername, _
            MODULE_NAME & ".CompareUser", _
            "Username requestor tidak boleh kosong."
    End If

    If ValidationText.IsBlank(approverUsername) Then
        ErrorManager.ThrowError _
            serviceUser.EmptyApproverUsername, _
            MODULE_NAME & ".CompareUser", _
            "Username approver tidak boleh kosong."
    End If

    If ValidationText.IsSameText(requestorUsername, approverUsername) Then
        ErrorManager.ThrowError _
            serviceUser.SameRequestorApprover, _
            MODULE_NAME & ".CompareUser", _
            "Requestor dan approver tidak boleh " & _
            "menggunakan username yang sama."
    End If

    ' =====================================================
    ' Step 2: Query
    ' =====================================================
    Set conn = DatabaseManager.GetConnection()

    Set rs = daoUser.CompareUser(conn, requestorUsername, approverUsername)

    ' =====================================================
    ' Step 3: Validate the query result
    '
    ' No rows normally means one or both usernames were
    ' not found in the users table.
    ' =====================================================
    If rs.EOF Then
        ErrorManager.ThrowError _
            serviceUser.NotFoundRequestorOrApprover, _
            MODULE_NAME & ".CompareUser", _
            "Requestor atau approver tidak ditemukan." & _
            vbCrLf & _
            "Requestor: " & requestorUsername & vbCrLf & _
            "Approver: " & approverUsername
    End If

    ' =====================================================
    ' Step 4: Read and convert can_approve
    ' =====================================================
    canApproveValue = DatabaseManager.FieldValue(rs, "can_approve")

    If IsError(canApproveValue) Then
        ErrorManager.ThrowError _
            serviceUser.InvalidCanApproveErrorValue, _
            MODULE_NAME & ".CompareUser", _
            "Hasil can_approve berisi nilai error."
    End If

    If IsNull(canApproveValue) Or _
       IsEmpty(canApproveValue) Then

        ErrorManager.ThrowError _
            serviceUser.EmptyCanApproveResult, _
            MODULE_NAME & ".CompareUser", _
            "Hasil can_approve tidak boleh kosong."
    End If

    If VarType(canApproveValue) = vbBoolean Then

        CompareUser = CBool(canApproveValue)

    ElseIf IsNumeric(canApproveValue) Then

        CompareUser = (CLng(canApproveValue) <> 0)

    Else

        Select Case UCase$(Trim$(CStr(canApproveValue)))

            Case "TRUE", "YES", "YA"
                CompareUser = True

            Case "FALSE", "NO", "TIDAK"
                CompareUser = False

            Case Else
                ErrorManager.ThrowError _
                    serviceUser.InvalidCanApproveConversion, _
                    MODULE_NAME & ".CompareUser", _
                    "Hasil can_approve tidak dapat " & _
                    "dikonversi menjadi Boolean."

        End Select

    End If

    DatabaseManager.CloseRecordset rs

    Exit Function

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    DatabaseManager.CloseRecordset rs

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".CompareUser", _
        "Gagal membandingkan tier requestor " & _
        "dan approver."

End Function


' =========================================================
' INTERNAL
' =========================================================

Private Function BuildUserCreateDetail( _
    ByVal manufacturingPlantName As String, _
    ByVal username As String, _
    ByVal firstName As String, _
    ByVal middleName As String, _
    ByVal lastName As String, _
    ByVal position As String, _
    ByVal tier As Long) As String

    Dim detail As String
    Dim fullName As String

    fullName = BuildUserFullName( _
        firstName:=firstName, _
        middleName:=middleName, _
        lastName:=lastName)

    detail = "username=" & Trim$(username)
    detail = detail & "; full_name=" & fullName
    detail = detail & "; first_name=" & Trim$(firstName)
    detail = detail & "; middle_name=" & Trim$(middleName)
    detail = detail & "; last_name=" & Trim$(lastName)
    detail = detail & "; position=" & Trim$(position)
    detail = detail & "; tier=" & CStr(tier)
    detail = detail & "; manufacturing_plant_name=" & Trim$(manufacturingPlantName)

    BuildUserCreateDetail = detail

End Function


Private Function BuildUserFullName( _
    ByVal firstName As String, _
    ByVal middleName As String, _
    ByVal lastName As String) As String

    Dim fullName As String

    firstName = Trim$(firstName)
    middleName = Trim$(middleName)
    lastName = Trim$(lastName)

    If Len(firstName) > 0 Then
        fullName = firstName
    End If

    If Len(middleName) > 0 Then

        If Len(fullName) > 0 Then
            fullName = fullName & " "
        End If

        fullName = fullName & middleName

    End If

    If Len(lastName) > 0 Then

        If Len(fullName) > 0 Then
            fullName = fullName & " "
        End If

        fullName = fullName & lastName

    End If

    BuildUserFullName = Trim$(fullName)

End Function
