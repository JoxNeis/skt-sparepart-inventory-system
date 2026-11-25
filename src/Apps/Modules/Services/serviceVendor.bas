Attribute VB_Name = "serviceVendor"
Option Explicit

' =========================================================
' SERVICE - VENDOR
'
' The vendor form has no code field, so a new vendor gets a
' generated code: V0001, V0002, ... (the next free number).
' Sheets pick vendors by "code - name"; names are also kept
' unique (ignoring case).
' =========================================================

Private Const MODULE_NAME As String = "serviceVendor"

' =========================================================
' ERROR CODES (18200 - 18299)
' Usage: managerError.ThrowError serviceVendor.DuplicateVendorName, ...
' =========================================================
Public Const EmptyVendor As Long = 18201
Public Const DuplicateVendorName As Long = 18202
Public Const CreateFailed As Long = 18203
Public Const UpdateFailed As Long = 18204
Public Const ArchiveFailed As Long = 18205
Public Const VendorNotFound As Long = 18206
Public Const NoFreeVendorCode As Long = 18207

Private Const CODE_PREFIX As String = "V"
Private Const CODE_DIGITS As String = "0000"
Private Const CODE_ATTEMPTS As Long = 50


' Active vendors, ordered by name.
Public Function ListVendors() As Collection

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    Set ListVendors = daoVendor.ListVendors(managerDatabase.GetConnection())

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ListVendors", _
        "Gagal mengambil data vendor."

End Function


' Create a vendor. vendorItem.Code is generated when blank.
Public Sub CreateVendor( _
    ByVal vendorItem As Vendor)

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim inTransaction As Boolean
    Dim affected As Long
    Dim vendors As Collection
    Dim errInfo As ErrorInfo

    RequireVendor vendorItem, MODULE_NAME & ".CreateVendor"

    Set conn = managerDatabase.GetConnection()
    Set vendors = daoVendor.ListVendors(conn)

    RequireUniqueName vendors, vendorItem, MODULE_NAME & ".CreateVendor"

    managerDatabase.BeginTransaction conn
    inTransaction = True

    If Len(vendorItem.Code) = 0 Then
        affected = InsertWithGeneratedCode(conn, vendorItem, NextCodeNumber(vendors))
    Else
        affected = daoVendor.CreateVendor(conn, vendorItem)
    End If

    managerDatabase.RequireAffectedRows affected, 1, _
        serviceVendor.CreateFailed, _
        MODULE_NAME & ".CreateVendor", _
        "Vendor tidak tersimpan (0 baris terpengaruh)."

    serviceActivityLog.WriteActivityLog conn, _
        "Vendor.Create", _
        "Vendor " & vendorItem.Name & " dibuat.", _
        BuildVendorJson(vendorItem)

    managerDatabase.CommitTransaction conn
    inTransaction = False

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    If inTransaction Then managerDatabase.RollbackTransaction conn

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".CreateVendor", _
        "Gagal menyimpan vendor."

End Sub


' Update a vendor (by Id). The stored code is kept.
Public Sub UpdateVendor( _
    ByVal vendorItem As Vendor)

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim inTransaction As Boolean
    Dim affected As Long
    Dim vendors As Collection
    Dim storedVendor As Vendor
    Dim errInfo As ErrorInfo

    RequireVendor vendorItem, MODULE_NAME & ".UpdateVendor"

    Set conn = managerDatabase.GetConnection()
    Set vendors = daoVendor.ListVendors(conn)

    Set storedVendor = FindById(vendors, vendorItem.Id)

    If storedVendor Is Nothing Then
        managerError.ThrowError _
            serviceVendor.VendorNotFound, _
            MODULE_NAME & ".UpdateVendor", _
            "Vendor tidak ditemukan atau sudah diarsipkan."
    End If

    RequireUniqueName vendors, vendorItem, MODULE_NAME & ".UpdateVendor"

    vendorItem.Code = storedVendor.Code

    managerDatabase.BeginTransaction conn
    inTransaction = True

    affected = daoVendor.UpdateVendor(conn, vendorItem)

    managerDatabase.RequireAffectedRows affected, 1, _
        serviceVendor.UpdateFailed, _
        MODULE_NAME & ".UpdateVendor", _
        "Vendor tidak diperbarui (0 baris terpengaruh)."

    serviceActivityLog.WriteActivityLog conn, _
        "Vendor.Update", _
        "Vendor " & vendorItem.Name & " diperbarui.", _
        BuildVendorJson(vendorItem)

    managerDatabase.CommitTransaction conn
    inTransaction = False

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    If inTransaction Then managerDatabase.RollbackTransaction conn

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".UpdateVendor", _
        "Gagal memperbarui vendor."

End Sub


' Archive (soft delete) a vendor.
Public Sub ArchiveVendor( _
    ByVal vendorId As Long)

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim inTransaction As Boolean
    Dim affected As Long
    Dim storedVendor As Vendor
    Dim errInfo As ErrorInfo

    Set conn = managerDatabase.GetConnection()

    Set storedVendor = FindById(daoVendor.ListVendors(conn), vendorId)

    If storedVendor Is Nothing Then
        managerError.ThrowError _
            serviceVendor.VendorNotFound, _
            MODULE_NAME & ".ArchiveVendor", _
            "Vendor tidak ditemukan atau sudah diarsipkan."
    End If

    managerDatabase.BeginTransaction conn
    inTransaction = True

    affected = daoVendor.ArchiveVendor(conn, vendorId)

    managerDatabase.RequireAffectedRows affected, 1, _
        serviceVendor.ArchiveFailed, _
        MODULE_NAME & ".ArchiveVendor", _
        "Vendor tidak dihapus (0 baris terpengaruh)."

    serviceActivityLog.WriteActivityLog conn, _
        "Vendor.Archive", _
        "Vendor " & storedVendor.Name & " diarsipkan.", _
        BuildVendorJson(storedVendor)

    managerDatabase.CommitTransaction conn
    inTransaction = False

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    If inTransaction Then managerDatabase.RollbackTransaction conn

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ArchiveVendor", _
        "Gagal menghapus vendor."

End Sub


' The vendor in vendors whose name, code or display name
' equals text (ignoring case). Nothing when there is none.
Public Function FindInList( _
    ByVal vendors As Collection, _
    ByVal text As String) As Vendor

    Dim vendorItem As Vendor

    text = Trim$(text)
    If Len(text) = 0 Then Exit Function

    For Each vendorItem In vendors
        If managerValidationText.IsSameText(vendorItem.Name, text) _
            Or managerValidationText.IsSameText(vendorItem.Code, text) _
            Or managerValidationText.IsSameText(vendorItem.DisplayName, text) Then

            Set FindInList = vendorItem
            Exit Function
        End If
    Next vendorItem

End Function


' JSON for activity logs: {"code": "", "name": ""}.
Public Function BuildVendorReferenceJson( _
    ByVal vendorItem As Vendor) As String

    If vendorItem Is Nothing Then
        BuildVendorReferenceJson = "null"
    Else
        BuildVendorReferenceJson = coreJson.JsonObject( _
            "code", coreJson.JsonText(vendorItem.Code), _
            "name", coreJson.JsonText(vendorItem.Name))
    End If

End Function


' =========================================================
' INTERNAL
' =========================================================

Private Function FindById( _
    ByVal vendors As Collection, _
    ByVal vendorId As Long) As Vendor

    Dim vendorItem As Vendor

    For Each vendorItem In vendors
        If vendorItem.Id = vendorId Then
            Set FindById = vendorItem
            Exit Function
        End If
    Next vendorItem

End Function


' 1 + the highest number of the active V#### codes.
Private Function NextCodeNumber( _
    ByVal vendors As Collection) As Long

    Dim vendorItem As Vendor
    Dim digits As String
    Dim result As Long

    result = 1

    For Each vendorItem In vendors
        If UCase$(Left$(vendorItem.Code, Len(CODE_PREFIX))) = CODE_PREFIX Then
            digits = Mid$(vendorItem.Code, Len(CODE_PREFIX) + 1)
            If Len(digits) > 0 And Not (digits Like "*[!0-9]*") Then
                If CLng(Val(digits)) >= result Then result = CLng(Val(digits)) + 1
            End If
        End If
    Next vendorItem

    NextCodeNumber = result

End Function


' Insert with V<firstNumber>; an archived vendor may still own
' that code (vendor_list does not show archived rows), so on a
' duplicate the next number is tried. Returns the affected rows.
Private Function InsertWithGeneratedCode( _
    ByVal conn As ADODB.connection, _
    ByVal vendorItem As Vendor, _
    ByVal firstNumber As Long) As Long

    Dim attempt As Long
    Dim errNumber As Long
    Dim errSource As String
    Dim errText As String

    For attempt = 0 To CODE_ATTEMPTS - 1

        vendorItem.Code = CODE_PREFIX & Format$(firstNumber + attempt, CODE_DIGITS)

        On Error Resume Next
        InsertWithGeneratedCode = daoVendor.CreateVendor(conn, vendorItem)
        errNumber = Err.Number
        errSource = Err.Source
        errText = Err.Description
        On Error GoTo 0

        If errNumber = 0 Then Exit Function

        If InStr(1, errText, "duplicate", vbTextCompare) = 0 Then
            Err.Raise errNumber, errSource, errText
        End If

    Next attempt

    managerError.ThrowError serviceVendor.NoFreeVendorCode, _
        MODULE_NAME & ".InsertWithGeneratedCode", _
        "Tidak ada kode vendor yang bebas mulai " & CODE_PREFIX & _
        Format$(firstNumber, CODE_DIGITS) & "."

End Function


Private Sub RequireVendor( _
    ByVal vendorItem As Vendor, _
    ByVal errorSource As String)

    If vendorItem Is Nothing Then
        managerError.ThrowError _
            serviceVendor.EmptyVendor, _
            errorSource, _
            "Data vendor tidak boleh kosong."
    End If

End Sub


' Another active vendor may not have the same name.
Private Sub RequireUniqueName( _
    ByVal vendors As Collection, _
    ByVal vendorItem As Vendor, _
    ByVal errorSource As String)

    Dim other As Vendor

    For Each other In vendors
        If other.Id <> vendorItem.Id Then
            If managerValidationText.IsSameText(other.Name, vendorItem.Name) Then
                managerError.ThrowError _
                    serviceVendor.DuplicateVendorName, _
                    errorSource, _
                    "Vendor dengan nama '" & vendorItem.Name & "' sudah ada."
            End If
        End If
    Next other

End Sub


Private Function BuildVendorJson( _
    ByVal vendorItem As Vendor) As String

    BuildVendorJson = coreJson.JsonObject( _
        "id", coreJson.JsonNumber(vendorItem.Id), _
        "code", coreJson.JsonText(vendorItem.Code), _
        "name", coreJson.JsonText(vendorItem.Name), _
        "company_name", coreJson.JsonText(vendorItem.CompanyName), _
        "email", coreJson.JsonText(vendorItem.Email))

End Function
