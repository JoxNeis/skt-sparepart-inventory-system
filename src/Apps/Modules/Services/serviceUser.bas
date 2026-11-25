Attribute VB_Name = "serviceUser"
Option Explicit

' =========================================================
' SERVICE - USER
'
' Master data users. users.username is the Windows login and
' is the value stored in every *_by column, so every document
' needs the current Windows user to be registered here
' (RequireRegisteredUser).
' =========================================================

Private Const MODULE_NAME As String = "serviceUser"

' =========================================================
' ERROR CODES (18100 - 18199)
' Usage: managerError.ThrowError serviceUser.UnregisteredUser, ...
' =========================================================
Public Const EmptyUser As Long = 18101
Public Const DuplicateUsername As Long = 18102
Public Const CreateFailed As Long = 18103
Public Const UpdateFailed As Long = 18104
Public Const ArchiveFailed As Long = 18105
Public Const UnregisteredUser As Long = 18106       ' Windows user not in users
Public Const EmptyPlant As Long = 18107
Public Const UserNotFound As Long = 18108
Public Const ArchiveSelf As Long = 18109            ' cannot archive yourself


' Active users of every plant, ordered by username.
Public Function ListUsers() As Collection

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    Set ListUsers = daoUser.ListUsers(managerDatabase.GetConnection())

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ListUsers", _
        "Gagal mengambil data user."

End Function


' Create a user of the plant named plantName.
Public Sub CreateUser( _
    ByVal userItem As User, _
    ByVal plantName As String)

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim inTransaction As Boolean
    Dim affected As Long
    Dim errInfo As ErrorInfo

    RequireUser userItem, MODULE_NAME & ".CreateUser"

    Set conn = managerDatabase.GetConnection()

    Set userItem.ManufacturingPlant = _
        serviceManufacturingPlant.FindPlant(conn, plantName)

    If Not FindInList(daoUser.ListUsers(conn), userItem.Username) Is Nothing Then
        managerError.ThrowError _
            serviceUser.DuplicateUsername, _
            MODULE_NAME & ".CreateUser", _
            "Username '" & userItem.Username & "' sudah terdaftar."
    End If

    managerDatabase.BeginTransaction conn
    inTransaction = True

    affected = daoUser.CreateUser(conn, userItem)

    managerDatabase.RequireAffectedRows affected, 1, _
        serviceUser.CreateFailed, _
        MODULE_NAME & ".CreateUser", _
        "User tidak tersimpan (0 baris terpengaruh)."

    serviceActivityLog.WriteActivityLog conn, _
        "User.Create", _
        "User " & userItem.Username & " (" & userItem.FullName & ") dibuat.", _
        BuildUserJson(userItem)

    managerDatabase.CommitTransaction conn
    inTransaction = False

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    If inTransaction Then managerDatabase.RollbackTransaction conn

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".CreateUser", _
        "Gagal menyimpan user."

End Sub


' Update names, jobdesk and tier of an existing user (by
' Id). The username and plant of the stored user are kept.
Public Sub UpdateUser( _
    ByVal userItem As User)

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim inTransaction As Boolean
    Dim affected As Long
    Dim storedUser As User
    Dim errInfo As ErrorInfo

    RequireUser userItem, MODULE_NAME & ".UpdateUser"

    Set conn = managerDatabase.GetConnection()

    Set storedUser = FindById(daoUser.ListUsers(conn), userItem.Id)

    If storedUser Is Nothing Then
        managerError.ThrowError _
            serviceUser.UserNotFound, _
            MODULE_NAME & ".UpdateUser", _
            "User tidak ditemukan atau sudah diarsipkan."
    End If

    userItem.Username = storedUser.Username
    Set userItem.ManufacturingPlant = storedUser.ManufacturingPlant

    managerDatabase.BeginTransaction conn
    inTransaction = True

    affected = daoUser.UpdateUser(conn, userItem)

    managerDatabase.RequireAffectedRows affected, 1, _
        serviceUser.UpdateFailed, _
        MODULE_NAME & ".UpdateUser", _
        "User tidak diperbarui (0 baris terpengaruh)."

    serviceActivityLog.WriteActivityLog conn, _
        "User.Update", _
        "User " & userItem.Username & " diperbarui.", _
        BuildUserJson(userItem)

    managerDatabase.CommitTransaction conn
    inTransaction = False

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    If inTransaction Then managerDatabase.RollbackTransaction conn

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".UpdateUser", _
        "Gagal memperbarui user."

End Sub


' Archive (soft delete) a user.
Public Sub ArchiveUser( _
    ByVal userId As Long)

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim inTransaction As Boolean
    Dim affected As Long
    Dim storedUser As User
    Dim errInfo As ErrorInfo

    Set conn = managerDatabase.GetConnection()

    Set storedUser = FindById(daoUser.ListUsers(conn), userId)

    If storedUser Is Nothing Then
        managerError.ThrowError _
            serviceUser.UserNotFound, _
            MODULE_NAME & ".ArchiveUser", _
            "User tidak ditemukan atau sudah diarsipkan."
    End If

    If managerValidationText.IsSameText(storedUser.Username, coreConfig.currentUser()) Then
        managerError.ThrowError _
            serviceUser.ArchiveSelf, _
            MODULE_NAME & ".ArchiveUser", _
            "User yang sedang dipakai tidak dapat dihapus."
    End If

    managerDatabase.BeginTransaction conn
    inTransaction = True

    affected = daoUser.ArchiveUser(conn, userId)

    managerDatabase.RequireAffectedRows affected, 1, _
        serviceUser.ArchiveFailed, _
        MODULE_NAME & ".ArchiveUser", _
        "User tidak dihapus (0 baris terpengaruh)."

    serviceActivityLog.WriteActivityLog conn, _
        "User.Archive", _
        "User " & storedUser.Username & " diarsipkan.", _
        BuildUserJson(storedUser)

    managerDatabase.CommitTransaction conn
    inTransaction = False

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    If inTransaction Then managerDatabase.RollbackTransaction conn

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ArchiveUser", _
        "Gagal menghapus user."

End Sub


' The registered user of the current Windows login, or
' Nothing when it is not registered.
Public Function GetCurrentUser() As User

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    Set GetCurrentUser = FindInList( _
        daoUser.ListUsers(managerDatabase.GetConnection()), _
        coreConfig.currentUser())

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".GetCurrentUser", _
        "Gagal mengambil data user."

End Function


' The registered user of the current Windows login, on the
' caller's connection. Raises UnregisteredUser.
Public Function RequireRegisteredUser( _
    ByVal conn As ADODB.connection) As User

    Dim userItem As User

    Set userItem = FindInList(daoUser.ListUsers(conn), coreConfig.currentUser())

    If userItem Is Nothing Then
        managerError.ThrowError _
            serviceUser.UnregisteredUser, _
            MODULE_NAME & ".RequireRegisteredUser", _
            "User Windows '" & coreConfig.currentUser() & _
            "' belum terdaftar. Tambahkan user ini di MASTER DATA USER " & _
            "terlebih dahulu."
    End If

    Set RequireRegisteredUser = userItem

End Function


' The user in users with this username (ignoring case), or
' Nothing.
Public Function FindInList( _
    ByVal users As Collection, _
    ByVal username As String) As User

    Dim userItem As User

    username = Trim$(username)
    If Len(username) = 0 Then Exit Function

    For Each userItem In users
        If managerValidationText.IsSameText(userItem.Username, username) Then
            Set FindInList = userItem
            Exit Function
        End If
    Next userItem

End Function


' =========================================================
' INTERNAL
' =========================================================

Private Function FindById( _
    ByVal users As Collection, _
    ByVal userId As Long) As User

    Dim userItem As User

    For Each userItem In users
        If userItem.Id = userId Then
            Set FindById = userItem
            Exit Function
        End If
    Next userItem

End Function


Private Sub RequireUser( _
    ByVal userItem As User, _
    ByVal errorSource As String)

    If userItem Is Nothing Then
        managerError.ThrowError _
            serviceUser.EmptyUser, _
            errorSource, _
            "Data user tidak boleh kosong."
    End If

End Sub


Private Function BuildUserJson( _
    ByVal userItem As User) As String

    Dim plantJson As String

    If userItem.ManufacturingPlant Is Nothing Then
        plantJson = "null"
    Else
        plantJson = coreJson.JsonObject( _
            "code", coreJson.JsonText(userItem.ManufacturingPlant.Code), _
            "name", coreJson.JsonText(userItem.ManufacturingPlant.Name))
    End If

    BuildUserJson = coreJson.JsonObject( _
        "id", coreJson.JsonNumber(userItem.Id), _
        "manufacturing_plant", plantJson, _
        "username", coreJson.JsonText(userItem.Username), _
        "first_name", coreJson.JsonText(userItem.FirstName), _
        "middle_name", coreJson.JsonText(userItem.MiddleName), _
        "last_name", coreJson.JsonText(userItem.LastName), _
        "jobdesk", coreJson.JsonText(userItem.Jobdesk), _
        "tier", coreJson.JsonNumber(userItem.Tier))

End Function
