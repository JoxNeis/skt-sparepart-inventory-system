Attribute VB_Name = "serviceMobileMaintenance"
Option Explicit

' =========================================================
' SERVICE - MOBILE MAINTENANCE (pos mobile)
'
' Code and name are unique PER PLANT. Sheets pick a mobile
' maintenance by name within the plant of the sheet.
' =========================================================

Private Const MODULE_NAME As String = "serviceMobileMaintenance"

' =========================================================
' ERROR CODES (17300 - 17399)
' Usage: managerError.ThrowError serviceMobileMaintenance.DuplicateName, ...
' =========================================================
Public Const EmptyMobileMaintenance As Long = 17301
Public Const DuplicateName As Long = 17302
Public Const DuplicateCode As Long = 17303
Public Const CreateFailed As Long = 17304
Public Const UpdateFailed As Long = 17305
Public Const ArchiveFailed As Long = 17306
Public Const MobileMaintenanceNotFound As Long = 17307


' Active mobile maintenances of one plant, ordered by name.
Public Function ListByPlant( _
    ByVal plantName As String) As Collection

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim errInfo As ErrorInfo

    Set conn = managerDatabase.GetConnection()

    Set ListByPlant = daoMobileMaintenance.ListByPlant( _
        conn, serviceManufacturingPlant.FindPlant(conn, plantName))

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ListByPlant", _
        "Gagal mengambil data pos mobile."

End Function


' Create a mobile maintenance in the plant named plantName.
Public Sub CreateMobileMaintenance( _
    ByVal item As MobileMaintenance, _
    ByVal plantName As String)

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim inTransaction As Boolean
    Dim affected As Long
    Dim errInfo As ErrorInfo

    RequireItem item, MODULE_NAME & ".CreateMobileMaintenance"

    Set conn = managerDatabase.GetConnection()

    Set item.ManufacturingPlant = serviceManufacturingPlant.FindPlant(conn, plantName)

    RequireUnique daoMobileMaintenance.ListByPlant(conn, item.ManufacturingPlant), _
        item, MODULE_NAME & ".CreateMobileMaintenance"

    managerDatabase.BeginTransaction conn
    inTransaction = True

    affected = daoMobileMaintenance.CreateMobileMaintenance(conn, item)

    managerDatabase.RequireAffectedRows affected, 1, _
        serviceMobileMaintenance.CreateFailed, _
        MODULE_NAME & ".CreateMobileMaintenance", _
        "Pos mobile tidak tersimpan (0 baris terpengaruh). " & _
        "Kode atau nama mungkin sudah dipakai pos mobile yang diarsipkan."

    serviceActivityLog.WriteActivityLog conn, _
        "MobileMaintenance.Create", _
        "Pos mobile " & item.DisplayName & " dibuat di " & _
        item.ManufacturingPlant.Name & ".", _
        BuildJson(item)

    managerDatabase.CommitTransaction conn
    inTransaction = False

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    If inTransaction Then managerDatabase.RollbackTransaction conn

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".CreateMobileMaintenance", _
        "Gagal menyimpan pos mobile."

End Sub


' Update code and name of a mobile maintenance (by Id) of
' the plant named plantName.
Public Sub UpdateMobileMaintenance( _
    ByVal item As MobileMaintenance, _
    ByVal plantName As String)

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim inTransaction As Boolean
    Dim affected As Long
    Dim errInfo As ErrorInfo

    RequireItem item, MODULE_NAME & ".UpdateMobileMaintenance"

    Set conn = managerDatabase.GetConnection()

    Set item.ManufacturingPlant = serviceManufacturingPlant.FindPlant(conn, plantName)

    RequireUnique daoMobileMaintenance.ListByPlant(conn, item.ManufacturingPlant), _
        item, MODULE_NAME & ".UpdateMobileMaintenance"

    managerDatabase.BeginTransaction conn
    inTransaction = True

    affected = daoMobileMaintenance.UpdateMobileMaintenance(conn, item)

    managerDatabase.RequireAffectedRows affected, 1, _
        serviceMobileMaintenance.UpdateFailed, _
        MODULE_NAME & ".UpdateMobileMaintenance", _
        "Pos mobile tidak diperbarui (tidak ditemukan atau sudah diarsipkan)."

    serviceActivityLog.WriteActivityLog conn, _
        "MobileMaintenance.Update", _
        "Pos mobile " & item.DisplayName & " diperbarui.", _
        BuildJson(item)

    managerDatabase.CommitTransaction conn
    inTransaction = False

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    If inTransaction Then managerDatabase.RollbackTransaction conn

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".UpdateMobileMaintenance", _
        "Gagal memperbarui pos mobile."

End Sub


' Archive (soft delete) a mobile maintenance.
Public Sub ArchiveMobileMaintenance( _
    ByVal mobileMaintenanceId As Long, _
    ByVal displayName As String)

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim inTransaction As Boolean
    Dim affected As Long
    Dim errInfo As ErrorInfo

    Set conn = managerDatabase.GetConnection()

    managerDatabase.BeginTransaction conn
    inTransaction = True

    affected = daoMobileMaintenance.ArchiveMobileMaintenance(conn, mobileMaintenanceId)

    managerDatabase.RequireAffectedRows affected, 1, _
        serviceMobileMaintenance.ArchiveFailed, _
        MODULE_NAME & ".ArchiveMobileMaintenance", _
        "Pos mobile tidak dihapus (tidak ditemukan atau sudah diarsipkan)."

    serviceActivityLog.WriteActivityLog conn, _
        "MobileMaintenance.Archive", _
        "Pos mobile " & displayName & " diarsipkan.", _
        coreJson.JsonObject( _
            "id", coreJson.JsonNumber(mobileMaintenanceId), _
            "name", coreJson.JsonText(displayName))

    managerDatabase.CommitTransaction conn
    inTransaction = False

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    If inTransaction Then managerDatabase.RollbackTransaction conn

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ArchiveMobileMaintenance", _
        "Gagal menghapus pos mobile."

End Sub


' The mobile maintenance in items whose name, code or display
' name equals text (ignoring case). Nothing when there is none.
Public Function FindInList( _
    ByVal items As Collection, _
    ByVal text As String) As MobileMaintenance

    Dim item As MobileMaintenance

    text = Trim$(text)
    If Len(text) = 0 Then Exit Function

    For Each item In items
        If managerValidationText.IsSameText(item.Name, text) _
            Or managerValidationText.IsSameText(item.Code, text) _
            Or managerValidationText.IsSameText(item.DisplayName, text) Then

            Set FindInList = item
            Exit Function
        End If
    Next item

End Function


' JSON for activity logs: {"code": "", "name": ""}.
Public Function BuildReferenceJson( _
    ByVal item As MobileMaintenance) As String

    If item Is Nothing Then
        BuildReferenceJson = "null"
    Else
        BuildReferenceJson = coreJson.JsonObject( _
            "code", coreJson.JsonText(item.Code), _
            "name", coreJson.JsonText(item.Name))
    End If

End Function


' =========================================================
' INTERNAL
' =========================================================

Private Sub RequireItem( _
    ByVal item As MobileMaintenance, _
    ByVal errorSource As String)

    If item Is Nothing Then
        managerError.ThrowError _
            serviceMobileMaintenance.EmptyMobileMaintenance, _
            errorSource, _
            "Data pos mobile tidak boleh kosong."
    End If

End Sub


' No other mobile maintenance of the plant may have the same
' code or name.
Private Sub RequireUnique( _
    ByVal items As Collection, _
    ByVal item As MobileMaintenance, _
    ByVal errorSource As String)

    Dim other As MobileMaintenance

    For Each other In items

        If other.Id <> item.Id Then

            If managerValidationText.IsSameText(other.Name, item.Name) Then
                managerError.ThrowError _
                    serviceMobileMaintenance.DuplicateName, _
                    errorSource, _
                    "Pos mobile dengan nama '" & item.Name & _
                    "' sudah ada di plant ini."
            End If

            If managerValidationText.IsSameText(other.Code, item.Code) Then
                managerError.ThrowError _
                    serviceMobileMaintenance.DuplicateCode, _
                    errorSource, _
                    "Pos mobile dengan kode '" & item.Code & _
                    "' sudah ada di plant ini."
            End If

        End If

    Next other

End Sub


Private Function BuildJson( _
    ByVal item As MobileMaintenance) As String

    BuildJson = coreJson.JsonObject( _
        "id", coreJson.JsonNumber(item.Id), _
        "manufacturing_plant", coreJson.JsonObject( _
            "code", coreJson.JsonText(item.ManufacturingPlant.Code), _
            "name", coreJson.JsonText(item.ManufacturingPlant.Name)), _
        "code", coreJson.JsonText(item.Code), _
        "name", coreJson.JsonText(item.Name))

End Function
