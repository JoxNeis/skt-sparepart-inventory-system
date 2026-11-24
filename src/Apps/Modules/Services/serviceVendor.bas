Attribute VB_Name = "serviceVendor"
Option Explicit

' =========================================================
' SERVICE - VENDOR
'
' Public entry points (called by sheets) get the
' connection and own the transaction:
'   SelectVendors, CreateVendor
' =========================================================

Private Const MODULE_NAME As String = "SERVICE_VENDOR"

' =========================================================
' ERROR CODES (18200 - 18299)
' Usage: managerError.ThrowError serviceVendor.EmptyVendorObject, ...
' =========================================================
Public Const EmptyRecordsetVendorName As Long = 18201          ' query returned a blank vendor name
Public Const EmptyRecordsetVendorCode As Long = 18202          ' query returned a blank vendor code
Public Const EmptyRecordsetVendorCompanyName As Long = 18203   ' query returned a blank company name
Public Const EmptyVendorObject As Long = 18204                 ' Vendor object is Nothing
Public Const EmptyVendorName As Long = 18205
Public Const EmptyVendorCompanyName As Long = 18206
Public Const EmptyVendorCode As Long = 18207
Public Const EmptyCurrentUser As Long = 18208                  ' creating user cannot be determined
Public Const NoRowAffectedCreateVendor As Long = 18209         ' 0 rows inserted
Public Const MultipleRowsAffectedCreateVendor As Long = 18210  ' > 1 row inserted


' =========================================================
' PUBLIC ENTRY POINTS
' =========================================================

' Every vendor. Read only, no transaction.
Public Function SelectVendors() As Collection

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim vendors As Collection
    Dim vendorItem As Vendor
    Dim recordNumber As Long
    Dim errInfo As ErrorInfo

    ' =====================================================
    ' Step 1: Query
    ' =====================================================
    Set conn = managerDatabase.GetConnection()

    Set vendors = daoVendor.SelectVendors(conn)

    ' =====================================================
    ' Step 2: Validate values required by Vendor properties
    ' =====================================================
    recordNumber = 0

    For Each vendorItem In vendors

        recordNumber = recordNumber + 1

        If managerValidationText.IsBlank(vendorItem.Name) Then
            managerError.ThrowError _
                serviceVendor.EmptyRecordsetVendorName, _
                MODULE_NAME & ".SelectVendors", _
                "Nama vendor pada record ke-" & _
                CStr(recordNumber) & _
                " tidak boleh kosong."
        End If

        If managerValidationText.IsBlank(vendorItem.Code) Then
            managerError.ThrowError _
                serviceVendor.EmptyRecordsetVendorCode, _
                MODULE_NAME & ".SelectVendors", _
                "Kode vendor '" & vendorItem.Name & _
                "' tidak boleh kosong."
        End If

        If managerValidationText.IsBlank(vendorItem.CompanyName) Then
            managerError.ThrowError _
                serviceVendor.EmptyRecordsetVendorCompanyName, _
                MODULE_NAME & ".SelectVendors", _
                "Nama perusahaan vendor '" & _
                vendorItem.Name & _
                "' tidak boleh kosong."
        End If

    Next vendorItem

    Set SelectVendors = vendors

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".SelectVendors", _
        "Gagal mengambil daftar vendor."

End Function


' Insert one vendor and write the activity log, in one
' transaction.
Public Sub CreateVendor( _
    ByVal vendorItem As Vendor)

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim affected As Long
    Dim inTransaction As Boolean

    Dim vendorName As String
    Dim vendorCompanyName As String
    Dim vendorCode As String
    Dim currentUser As String
    Dim activityDetail As String

    Dim errInfo As ErrorInfo

    inTransaction = False

    ' =====================================================
    ' Step 1: Validate Vendor object
    ' =====================================================
    If vendorItem Is Nothing Then
        managerError.ThrowError _
            serviceVendor.EmptyVendorObject, _
            MODULE_NAME & ".CreateVendor", _
            "Objek Vendor tidak boleh kosong."
    End If

    ' =====================================================
    ' Step 2: Read and normalize Vendor properties
    ' =====================================================
    vendorName = Trim$(vendorItem.Name)
    vendorCompanyName = Trim$(vendorItem.CompanyName)
    vendorCode = Trim$(vendorItem.Code)

    ' =====================================================
    ' Step 3: Validate Vendor properties
    ' =====================================================
    If managerValidationText.IsBlank(vendorName) Then
        managerError.ThrowError _
            serviceVendor.EmptyVendorName, _
            MODULE_NAME & ".CreateVendor", _
            "Nama vendor tidak boleh kosong."
    End If

    If managerValidationText.IsBlank(vendorCompanyName) Then
        managerError.ThrowError _
            serviceVendor.EmptyVendorCompanyName, _
            MODULE_NAME & ".CreateVendor", _
            "Nama perusahaan vendor " & _
            "tidak boleh kosong."
    End If

    If managerValidationText.IsBlank(vendorCode) Then
        managerError.ThrowError _
            serviceVendor.EmptyVendorCode, _
            MODULE_NAME & ".CreateVendor", _
            "Kode vendor tidak boleh kosong."
    End If

    ' =====================================================
    ' Step 4: Determine current user
    ' =====================================================
    currentUser = Trim$(coreConfig.currentUser())

    If managerValidationText.IsBlank(currentUser) Then
        managerError.ThrowError _
            serviceVendor.EmptyCurrentUser, _
            MODULE_NAME & ".CreateVendor", _
            "User yang membuat vendor " & _
            "tidak dapat ditentukan."
    End If

    ' =====================================================
    ' Step 5: Begin transaction
    ' =====================================================
    Set conn = managerDatabase.GetConnection()

    managerDatabase.BeginTransaction conn
    inTransaction = True

    ' =====================================================
    ' Step 6: Insert vendor
    ' =====================================================
    affected = daoVendor.CreateVendor(conn, vendorItem)

    ' =====================================================
    ' Step 7: Validate affected rows
    ' =====================================================
    If managerValidationNumeric.IsLessThan(affected, 1) Then
        managerError.ThrowError _
            serviceVendor.NoRowAffectedCreateVendor, _
            MODULE_NAME & ".CreateVendor", _
            "Vendor tidak berhasil dibuat. " & _
            "Nama atau kode vendor mungkin sudah ada."
    End If

    If managerValidationNumeric.IsGreaterThan(affected, 1) Then
        managerError.ThrowError _
            serviceVendor.MultipleRowsAffectedCreateVendor, _
            MODULE_NAME & ".CreateVendor", _
            "Pembuatan vendor memengaruhi lebih dari " & _
            "satu record. Transaksi dibatalkan."
    End If

    ' =====================================================
    ' Step 8: Build activity detail
    ' =====================================================
    activityDetail = BuildVendorCreateDetail(vendorItem)

    ' =====================================================
    ' Step 9: Write activity log
    ' =====================================================
    serviceActivityLog.WriteActivityLog _
        conn, _
        "Vendor.Create", _
        activityDetail

    ' =====================================================
    ' Step 10: Commit
    ' =====================================================
    managerDatabase.CommitTransaction conn
    inTransaction = False

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    If inTransaction Then
        managerDatabase.RollbackTransaction conn
    End If

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".CreateVendor", _
        "Gagal membuat data vendor."

End Sub


' =========================================================
' INTERNAL
' =========================================================

Private Function BuildVendorCreateDetail( _
    ByVal vendorItem As Vendor) As String

    On Error GoTo ErrorHandler

    Dim detail As String
    Dim errInfo As ErrorInfo

    If vendorItem Is Nothing Then
        managerError.ThrowError _
            serviceVendor.EmptyVendorObject, _
            MODULE_NAME & ".BuildVendorCreateDetail", _
            "Objek Vendor tidak boleh kosong."
    End If

    detail = "name=" & Trim$(vendorItem.Name)
    detail = detail & "; company_name=" & Trim$(vendorItem.CompanyName)
    detail = detail & "; code=" & Trim$(vendorItem.Code)

    BuildVendorCreateDetail = detail

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".BuildVendorCreateDetail"

End Function
