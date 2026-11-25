Attribute VB_Name = "serviceSparepart"
Option Explicit

' =========================================================
' SERVICE - SPAREPART (catalogue)
'
' Sheets pick spareparts by their display name
' ("code - name", spareparts.display). FindInList also
' accepts the plain name or code.
' =========================================================

Private Const MODULE_NAME As String = "serviceSparepart"

' =========================================================
' ERROR CODES (17200 - 17299)
' Usage: managerError.ThrowError serviceSparepart.DuplicateCode, ...
' =========================================================
Public Const EmptySparepart As Long = 17201
Public Const DuplicateCode As Long = 17202
Public Const CreateFailed As Long = 17203
Public Const UpdateFailed As Long = 17204
Public Const ArchiveFailed As Long = 17205
Public Const SparepartNotFound As Long = 17206
Public Const AmbiguousSparepart As Long = 17207     ' same name, several codes


' Active spareparts, ordered by name.
Public Function ListSpareparts() As Collection

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    Set ListSpareparts = daoSparepart.ListSpareparts(managerDatabase.GetConnection())

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ListSpareparts", _
        "Gagal mengambil data sparepart."

End Function


' Create a sparepart. A blank alias becomes the name.
Public Sub CreateSparepart( _
    ByVal sparepartItem As Sparepart)

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim inTransaction As Boolean
    Dim affected As Long
    Dim errInfo As ErrorInfo

    RequireSparepart sparepartItem, MODULE_NAME & ".CreateSparepart"

    If Len(sparepartItem.Alias) = 0 Then
        sparepartItem.Alias = sparepartItem.Name
    End If

    Set conn = managerDatabase.GetConnection()

    RequireUniqueCode daoSparepart.ListSpareparts(conn), sparepartItem, _
        MODULE_NAME & ".CreateSparepart"

    managerDatabase.BeginTransaction conn
    inTransaction = True

    affected = daoSparepart.CreateSparepart(conn, sparepartItem)

    managerDatabase.RequireAffectedRows affected, 1, _
        serviceSparepart.CreateFailed, _
        MODULE_NAME & ".CreateSparepart", _
        "Sparepart tidak tersimpan (0 baris terpengaruh). " & _
        "Kode mungkin sudah dipakai sparepart yang diarsipkan."

    serviceActivityLog.WriteActivityLog conn, _
        "Sparepart.Create", _
        "Sparepart " & sparepartItem.DisplayName & " dibuat.", _
        BuildSparepartJson(sparepartItem)

    managerDatabase.CommitTransaction conn
    inTransaction = False

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    If inTransaction Then managerDatabase.RollbackTransaction conn

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".CreateSparepart", _
        "Gagal menyimpan sparepart."

End Sub


' Update a sparepart (by Id).
Public Sub UpdateSparepart( _
    ByVal sparepartItem As Sparepart)

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim inTransaction As Boolean
    Dim affected As Long
    Dim errInfo As ErrorInfo

    RequireSparepart sparepartItem, MODULE_NAME & ".UpdateSparepart"

    If Len(sparepartItem.Alias) = 0 Then
        sparepartItem.Alias = sparepartItem.Name
    End If

    Set conn = managerDatabase.GetConnection()

    RequireUniqueCode daoSparepart.ListSpareparts(conn), sparepartItem, _
        MODULE_NAME & ".UpdateSparepart"

    managerDatabase.BeginTransaction conn
    inTransaction = True

    affected = daoSparepart.UpdateSparepart(conn, sparepartItem)

    managerDatabase.RequireAffectedRows affected, 1, _
        serviceSparepart.UpdateFailed, _
        MODULE_NAME & ".UpdateSparepart", _
        "Sparepart tidak diperbarui (tidak ditemukan atau sudah diarsipkan)."

    serviceActivityLog.WriteActivityLog conn, _
        "Sparepart.Update", _
        "Sparepart " & sparepartItem.DisplayName & " diperbarui.", _
        BuildSparepartJson(sparepartItem)

    managerDatabase.CommitTransaction conn
    inTransaction = False

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    If inTransaction Then managerDatabase.RollbackTransaction conn

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".UpdateSparepart", _
        "Gagal memperbarui sparepart."

End Sub


' Archive (soft delete) a sparepart.
Public Sub ArchiveSparepart( _
    ByVal sparepartId As Long)

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim inTransaction As Boolean
    Dim affected As Long
    Dim errInfo As ErrorInfo

    Set conn = managerDatabase.GetConnection()

    managerDatabase.BeginTransaction conn
    inTransaction = True

    affected = daoSparepart.ArchiveSparepart(conn, sparepartId)

    managerDatabase.RequireAffectedRows affected, 1, _
        serviceSparepart.ArchiveFailed, _
        MODULE_NAME & ".ArchiveSparepart", _
        "Sparepart tidak dihapus (tidak ditemukan atau sudah diarsipkan)."

    serviceActivityLog.WriteActivityLog conn, _
        "Sparepart.Archive", _
        "Sparepart id " & sparepartId & " diarsipkan.", _
        coreJson.JsonObject("id", coreJson.JsonNumber(sparepartId))

    managerDatabase.CommitTransaction conn
    inTransaction = False

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    If inTransaction Then managerDatabase.RollbackTransaction conn

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ArchiveSparepart", _
        "Gagal menghapus sparepart."

End Sub


' The sparepart in spareparts that text points to:
'   1. display name "code - name"
'   2. code
'   3. name (only when exactly one sparepart has it)
' Ignoring case. Nothing when there is none. Raises
' AmbiguousSparepart when the name belongs to several codes.
Public Function FindInList( _
    ByVal spareparts As Collection, _
    ByVal text As String) As Sparepart

    Dim sparepartItem As Sparepart
    Dim nameMatch As Sparepart
    Dim nameMatchCount As Long

    text = Trim$(text)
    If Len(text) = 0 Then Exit Function

    For Each sparepartItem In spareparts

        If managerValidationText.IsSameText(sparepartItem.DisplayName, text) _
            Or managerValidationText.IsSameText(sparepartItem.Code, text) Then

            Set FindInList = sparepartItem
            Exit Function
        End If

        If managerValidationText.IsSameText(sparepartItem.Name, text) Then
            nameMatchCount = nameMatchCount + 1
            Set nameMatch = sparepartItem
        End If

    Next sparepartItem

    If nameMatchCount = 1 Then
        Set FindInList = nameMatch
    ElseIf nameMatchCount > 1 Then
        managerError.ThrowError _
            serviceSparepart.AmbiguousSparepart, _
            MODULE_NAME & ".FindInList", _
            "Ada " & nameMatchCount & " sparepart bernama '" & text & _
            "'. Pilih dari daftar (kode - nama)."
    End If

End Function


' Spareparts whose display name, alias or applicable contains
' text (ignoring case), exact matches first.
Public Function SearchInList( _
    ByVal spareparts As Collection, _
    ByVal text As String) As Collection

    Dim result As Collection
    Dim sparepartItem As Sparepart
    Dim exact As Sparepart

    Set result = New Collection

    text = Trim$(text)

    If Len(text) = 0 Then
        Set SearchInList = result
        Exit Function
    End If

    On Error Resume Next
    Set exact = FindInList(spareparts, text)
    On Error GoTo 0

    If Not exact Is Nothing Then result.Add exact

    For Each sparepartItem In spareparts

        If Not sparepartItem Is exact Then
            If InStr(1, sparepartItem.DisplayName, text, vbTextCompare) > 0 _
                Or InStr(1, sparepartItem.Alias, text, vbTextCompare) > 0 _
                Or InStr(1, sparepartItem.Applicable, text, vbTextCompare) > 0 Then

                result.Add sparepartItem
            End If
        End If

    Next sparepartItem

    Set SearchInList = result

End Function


' JSON for activity logs:
' {"code": "", "name": "", "unit_of_measurement": ""}.
Public Function BuildSparepartReferenceJson( _
    ByVal sparepartItem As Sparepart) As String

    If sparepartItem Is Nothing Then
        BuildSparepartReferenceJson = "null"
    Else
        BuildSparepartReferenceJson = coreJson.JsonObject( _
            "code", coreJson.JsonText(sparepartItem.Code), _
            "name", coreJson.JsonText(sparepartItem.Name), _
            "unit_of_measurement", coreJson.JsonText(sparepartItem.UnitOfMeasurement))
    End If

End Function


' =========================================================
' INTERNAL
' =========================================================

Private Sub RequireSparepart( _
    ByVal sparepartItem As Sparepart, _
    ByVal errorSource As String)

    If sparepartItem Is Nothing Then
        managerError.ThrowError _
            serviceSparepart.EmptySparepart, _
            errorSource, _
            "Data sparepart tidak boleh kosong."
    End If

End Sub


' Another active sparepart may not have the same code.
Private Sub RequireUniqueCode( _
    ByVal spareparts As Collection, _
    ByVal sparepartItem As Sparepart, _
    ByVal errorSource As String)

    Dim other As Sparepart

    For Each other In spareparts
        If other.Id <> sparepartItem.Id Then
            If managerValidationText.IsSameText(other.Code, sparepartItem.Code) Then
                managerError.ThrowError _
                    serviceSparepart.DuplicateCode, _
                    errorSource, _
                    "Kode sparepart '" & sparepartItem.Code & _
                    "' sudah dipakai oleh " & other.Name & "."
            End If
        End If
    Next other

End Sub


Private Function BuildSparepartJson( _
    ByVal sparepartItem As Sparepart) As String

    BuildSparepartJson = coreJson.JsonObject( _
        "id", coreJson.JsonNumber(sparepartItem.Id), _
        "code", coreJson.JsonText(sparepartItem.Code), _
        "name", coreJson.JsonText(sparepartItem.Name), _
        "category", coreJson.JsonText(sparepartItem.Category), _
        "applicable", coreJson.JsonText(sparepartItem.Applicable), _
        "unit_of_measurement", coreJson.JsonText(sparepartItem.UnitOfMeasurement), _
        "alias", coreJson.JsonText(sparepartItem.Alias), _
        "status", coreJson.JsonText(sparepartItem.Status))

End Function
