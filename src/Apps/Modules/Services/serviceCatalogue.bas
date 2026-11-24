Attribute VB_Name = "serviceCatalogue"
Option Explicit

' =========================================================
' SERVICE - CATALOGUE
'
' Public entry points (called by sheets) get the
' connection and own the transaction:
'   GetAllCatalogues, CreateCatalogue
' =========================================================

Private Const MODULE_NAME As String = "SERVICE_CATALOGUE"

' =========================================================
' ERROR CODES (17200 - 17299)
' Usage: managerError.ThrowError serviceCatalogue.EmptyCatalogueObject, ...
' =========================================================
Public Const EmptyCatalogueObject As Long = 17201                 ' Catalogue object is Nothing
Public Const EmptyCatalogueName As Long = 17202
Public Const EmptyCatalogueAlias As Long = 17203
Public Const EmptyCatalogueCode As Long = 17204
Public Const EmptyCatalogueCategory As Long = 17205
Public Const EmptyCatalogueUnitOfMeasurement As Long = 17206
Public Const EmptyCurrentUser As Long = 17207                     ' creating user cannot be determined
Public Const NoRowAffectedCreateCatalogue As Long = 17208         ' 0 rows inserted
Public Const MultipleRowsAffectedCreateCatalogue As Long = 17209  ' > 1 row inserted


' =========================================================
' PUBLIC ENTRY POINTS
' =========================================================

' Every catalogue. Read only, no transaction.
Public Function GetAllCatalogues() As Collection

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim errInfo As ErrorInfo

    Set conn = managerDatabase.GetConnection()

    Set GetAllCatalogues = daoCatalogue.GetAllCatalogues(conn)

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".GetAllCatalogues", _
        "Gagal mencari katalog."

End Function


' Insert one catalogue and write the activity log, in one
' transaction.
Public Sub CreateCatalogue( _
    ByVal catalogueItem As Catalogue)

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim affected As Long
    Dim inTransaction As Boolean

    Dim catalogueName As String
    Dim catalogueAlias As String
    Dim catalogueCode As String
    Dim catalogueCategory As String
    Dim unitOfMeasurement As String

    Dim activityDetail As String
    Dim currentUser As String

    Dim errInfo As ErrorInfo

    inTransaction = False

    ' =====================================================
    ' Step 1: Validate Catalogue object
    ' =====================================================
    If catalogueItem Is Nothing Then
        managerError.ThrowError _
            serviceCatalogue.EmptyCatalogueObject, _
            MODULE_NAME & ".CreateCatalogue", _
            "Objek Catalogue tidak boleh kosong."
    End If

    ' =====================================================
    ' Step 2: Read and normalize Catalogue properties
    ' =====================================================
    catalogueName = Trim$(catalogueItem.Name)
    catalogueAlias = Trim$(catalogueItem.Alias)
    catalogueCode = Trim$(catalogueItem.Code)
    catalogueCategory = Trim$(catalogueItem.Category)
    unitOfMeasurement = Trim$(catalogueItem.unitOfMeasurement)

    ' =====================================================
    ' Step 3: Validate Catalogue property values
    ' =====================================================
    If managerValidationText.IsBlank(catalogueName) Then
        managerError.ThrowError _
            serviceCatalogue.EmptyCatalogueName, _
            MODULE_NAME & ".CreateCatalogue", _
            "Nama catalogue tidak boleh kosong."
    End If

    If managerValidationText.IsBlank(catalogueAlias) Then
        managerError.ThrowError _
            serviceCatalogue.EmptyCatalogueAlias, _
            MODULE_NAME & ".CreateCatalogue", _
            "Alias catalogue tidak boleh kosong."
    End If

    If managerValidationText.IsBlank(catalogueCode) Then
        managerError.ThrowError _
            serviceCatalogue.EmptyCatalogueCode, _
            MODULE_NAME & ".CreateCatalogue", _
            "Kode catalogue tidak boleh kosong."
    End If

    If managerValidationText.IsBlank(catalogueCategory) Then
        managerError.ThrowError _
            serviceCatalogue.EmptyCatalogueCategory, _
            MODULE_NAME & ".CreateCatalogue", _
            "Kategori catalogue tidak boleh kosong."
    End If

    If managerValidationText.IsBlank(unitOfMeasurement) Then
        managerError.ThrowError _
            serviceCatalogue.EmptyCatalogueUnitOfMeasurement, _
            MODULE_NAME & ".CreateCatalogue", _
            "Satuan catalogue tidak boleh kosong."
    End If

    ' =====================================================
    ' Step 4: Determine current user
    ' =====================================================
    currentUser = Trim$(coreConfig.currentUser())

    If managerValidationText.IsBlank(currentUser) Then
        managerError.ThrowError _
            serviceCatalogue.EmptyCurrentUser, _
            MODULE_NAME & ".CreateCatalogue", _
            "User yang membuat catalogue " & _
            "tidak dapat ditentukan."
    End If

    ' =====================================================
    ' Step 5: Begin transaction
    ' =====================================================
    Set conn = managerDatabase.GetConnection()

    managerDatabase.BeginTransaction conn
    inTransaction = True

    ' =====================================================
    ' Step 6: Insert catalogue
    ' =====================================================
    affected = daoCatalogue.CreateCatalogue(conn, catalogueItem)

    ' =====================================================
    ' Step 7: Validate affected rows
    ' =====================================================
    If managerValidationNumeric.IsLessThan(affected, 1) Then
        managerError.ThrowError _
            serviceCatalogue.NoRowAffectedCreateCatalogue, _
            MODULE_NAME & ".CreateCatalogue", _
            "Catalogue tidak berhasil dibuat. " & _
            "Nama, alias, atau kode mungkin sudah " & _
            "digunakan oleh catalogue lain."
    End If

    If managerValidationNumeric.IsGreaterThan(affected, 1) Then
        managerError.ThrowError _
            serviceCatalogue.MultipleRowsAffectedCreateCatalogue, _
            MODULE_NAME & ".CreateCatalogue", _
            "Pembuatan catalogue memengaruhi lebih " & _
            "dari satu record. Transaksi dibatalkan."
    End If

    ' =====================================================
    ' Step 8: Build activity-log detail
    ' =====================================================
    activityDetail = BuildCatalogueCreateDetail(catalogueItem)

    ' =====================================================
    ' Step 9: Write activity log
    ' =====================================================
    serviceActivityLog.WriteActivityLog _
        conn, _
        "Catalogue.Create", _
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
        MODULE_NAME & ".CreateCatalogue", _
        "Gagal membuat data catalogue."

End Sub


' =========================================================
' INTERNAL
' =========================================================

Private Function BuildCatalogueCreateDetail( _
    ByVal catalogueItem As Catalogue) As String

    On Error GoTo ErrorHandler

    Dim detail As String
    Dim errInfo As ErrorInfo

    If catalogueItem Is Nothing Then
        managerError.ThrowError _
            serviceCatalogue.EmptyCatalogueObject, _
            MODULE_NAME & ".BuildCatalogueCreateDetail", _
            "Objek Catalogue tidak boleh kosong."
    End If

    detail = "name=" & Trim$(catalogueItem.Name)
    detail = detail & "; alias=" & Trim$(catalogueItem.Alias)
    detail = detail & "; code=" & Trim$(catalogueItem.Code)
    detail = detail & "; category=" & Trim$(catalogueItem.Category)
    detail = detail & "; unit_of_measurement=" & Trim$(catalogueItem.unitOfMeasurement)

    BuildCatalogueCreateDetail = detail

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".BuildCatalogueCreateDetail"

End Function
