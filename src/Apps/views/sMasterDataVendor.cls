Attribute VB_Name = "sMasterDataVendor"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Const MODULE_NAME As String = _
    "SHEET_MASTER_VENDOR"

Private Const TABLE_NAME As String = _
    "TableVendors"

' =========================================================
' Table column headers
'
' These names must exactly match the Excel table headers.
' =========================================================
Private Const COLUMN_ID As String = "Id"
Private Const COLUMN_NAME As String = "Nama"

Private Const COLUMN_COMPANY_NAME As String = _
    "Nama Perusahaan"

Private Const COLUMN_CODE As String = "Kode"


' =========================================================
' Vendor input cells
' =========================================================
Private Const CELL_VENDOR_NAME As String = "B3"
Private Const CELL_VENDOR_CODE As String = "B4"

Private Const CELL_VENDOR_COMPANY_NAME As String = _
    "B5"

' =========================================================
' REFRESH button
' =========================================================
Public Sub RefreshVendors_Click()

    On Error GoTo ErrorHandler

    Dim tbl As ListObject
    Dim rowIndex As Long

    Dim vendors As Collection
    Dim vendorItem As Vendor

    Dim vendorCount As Long
    Dim rowNumber As Long

    Dim state As FormState
    Dim errInfo As ErrorInfo

    state = managerForm.BeginBusy( _
        "Mengambil data vendor...")

    ' =====================================================
    ' Step 1: Get and validate TableVendors
    ' =====================================================
    Set tbl = managerTable.GetTable(Me, TABLE_NAME)

    ValidateVendorTableColumns tbl

    ' =====================================================
    ' Step 2: Call service function
    ' =====================================================
    Set vendors = serviceVendor.SelectVendors()

    If vendors Is Nothing Then
        managerError.ThrowError _
            sheetMasterDataVendor.NullVendorCollection, _
            MODULE_NAME & ".RefreshVendors_Click", _
            "DAO tidak mengembalikan collection vendor."
    End If

    vendorCount = vendors.Count

    ' =====================================================
    ' Step 3: Clear existing vendor rows
    ' =====================================================
    ClearVendorTable tbl

    ' =====================================================
    ' Step 4: Fill TableVendors
    ' =====================================================
    rowNumber = 0

    For Each vendorItem In vendors

        rowNumber = rowNumber + 1

        If vendorItem Is Nothing Then
            managerError.ThrowError _
                sheetMasterDataVendor.NullVendorInCollection, _
                MODULE_NAME & ".RefreshVendors_Click", _
                "Objek Vendor pada indeks " & _
                CStr(rowNumber) & _
                " tidak berhasil dibaca."
        End If

        ValidateVendorItem _
            vendorItem:=vendorItem, _
            rowNumber:=rowNumber

        rowIndex = managerTable.AddRow(tbl)

        WriteVendorTableRow _
            tbl:=tbl, _
            rowIndex:=rowIndex, _
            vendorItem:=vendorItem

    Next vendorItem

    ' =====================================================
    ' Step 5: Apply formatting
    ' =====================================================
    ApplyVendorTableFormatting tbl

    managerForm.ClearStatus

    ' =====================================================
    ' Step 6: Display result
    ' =====================================================
    If vendorCount = 0 Then

        managerForm.ShowInfo _
            "Tidak ada data vendor yang ditemukan.", _
            "Master Data Vendor"

    Else

        managerForm.ShowInfo _
            CStr(vendorCount) & _
            " data vendor berhasil ditampilkan.", _
            "Master Data Vendor"

    End If

CleanUp:
    Set vendorItem = Nothing
    Set vendors = Nothing
    Set tbl = Nothing

    managerForm.EndBusy state

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    Set vendorItem = Nothing
    Set vendors = Nothing
    Set tbl = Nothing

    managerForm.EndBusy state

    managerError.ShowError errInfo, _
        "Gagal menampilkan master data vendor."

End Sub


' =========================================================
' Validate one Vendor object
' =========================================================
Private Sub ValidateVendorItem( _
    ByVal vendorItem As Vendor, _
    ByVal rowNumber As Long)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    If vendorItem Is Nothing Then
        managerError.ThrowError _
            sheetMasterDataVendor.EmptyVendorItem, _
            MODULE_NAME & ".ValidateVendorItem", _
            "Objek Vendor tidak boleh kosong."
    End If

    If managerValidationNumeric.IsZeroOrNegative(vendorItem.Id) Then
        managerError.ThrowError _
            sheetMasterDataVendor.InvalidVendorId, _
            MODULE_NAME & ".ValidateVendorItem", _
            "ID vendor pada record ke-" & _
            CStr(rowNumber) & _
            " tidak valid."
    End If

    If managerValidationText.IsBlank(vendorItem.Name) Then
        managerError.ThrowError _
            sheetMasterDataVendor.EmptyRecordVendorName, _
            MODULE_NAME & ".ValidateVendorItem", _
            "Nama vendor pada record ke-" & _
            CStr(rowNumber) & _
            " tidak boleh kosong."
    End If

    If managerValidationText.IsBlank(vendorItem.CompanyName) Then
        managerError.ThrowError _
            sheetMasterDataVendor.EmptyRecordVendorCompanyName, _
            MODULE_NAME & ".ValidateVendorItem", _
            "Nama perusahaan vendor '" & _
            vendorItem.Name & _
            "' tidak boleh kosong."
    End If

    If managerValidationText.IsBlank(vendorItem.Code) Then
        managerError.ThrowError _
            sheetMasterDataVendor.EmptyRecordVendorCode, _
            MODULE_NAME & ".ValidateVendorItem", _
            "Kode vendor '" & _
            vendorItem.Name & _
            "' tidak boleh kosong."
    End If

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ValidateVendorItem"

End Sub


' =========================================================
' Write one Vendor object into one table row
' =========================================================
Private Sub WriteVendorTableRow( _
    ByVal tbl As ListObject, _
    ByVal rowIndex As Long, _
    ByVal vendorItem As Vendor)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    If tbl Is Nothing Then
        managerError.ThrowError _
            managerTable.TableNotFound, _
            MODULE_NAME & ".WriteVendorTableRow", _
            "TableVendors belum diinisialisasi."
    End If

    If Not managerTable.IsValidRow(tbl, rowIndex) Then
        managerError.ThrowError _
            managerTable.InvalidRow, _
            MODULE_NAME & ".WriteVendorTableRow", _
            "Target row vendor belum diinisialisasi."
    End If

    If vendorItem Is Nothing Then
        managerError.ThrowError _
            sheetMasterDataVendor.EmptyVendorItem, _
            MODULE_NAME & ".WriteVendorTableRow", _
            "Objek Vendor tidak boleh kosong."
    End If

    ' =====================================================
    ' D: Id
    ' =====================================================
    managerTable.SetValue _
        tbl, rowIndex, COLUMN_ID, vendorItem.Id

    ' =====================================================
    ' E: Nama
    ' =====================================================
    managerTable.SetValue _
        tbl, rowIndex, COLUMN_NAME, vendorItem.Name

    ' =====================================================
    ' F: Nama Perusahaan
    ' =====================================================
    managerTable.SetValue _
        tbl, rowIndex, COLUMN_COMPANY_NAME, _
        vendorItem.CompanyName

    ' =====================================================
    ' G: Kode
    ' =====================================================
    managerTable.SetValue _
        tbl, rowIndex, COLUMN_CODE, vendorItem.Code

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".WriteVendorTableRow", _
        "Gagal menulis satu data vendor ke tabel."

End Sub


' =========================================================
' Clear existing vendor rows
' =========================================================
Private Sub ClearVendorTable( _
    ByVal tbl As ListObject)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    If tbl Is Nothing Then
        managerError.ThrowError _
            managerTable.TableNotFound, _
            MODULE_NAME & ".ClearVendorTable", _
            "TableVendors belum diinisialisasi."
    End If

    managerTable.ClearRows tbl

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ClearVendorTable", _
        "Gagal membersihkan data vendor sebelumnya."

End Sub


' =========================================================
' Validate required table columns
' =========================================================
Private Sub ValidateVendorTableColumns( _
    ByVal tbl As ListObject)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    If tbl Is Nothing Then
        managerError.ThrowError _
            managerTable.TableNotFound, _
            MODULE_NAME & ".ValidateVendorTableColumns", _
            "TableVendors tidak ditemukan " & _
            "pada worksheet ini."
    End If

    managerTable.ValidateColumns tbl, Array( _
        COLUMN_ID, _
        COLUMN_NAME, _
        COLUMN_COMPANY_NAME, _
        COLUMN_CODE)

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & _
        ".ValidateVendorTableColumns", _
        "Struktur TableVendors tidak sesuai."

End Sub


' =========================================================
' Apply vendor table formatting
' =========================================================
Private Sub ApplyVendorTableFormatting( _
    ByVal tbl As ListObject)

    On Error GoTo ErrorHandler

    Dim allColumns As Variant
    Dim columnName As Variant
    Dim columnRange As Range

    Dim errInfo As ErrorInfo

    If tbl Is Nothing Then
        managerError.ThrowError _
            managerTable.TableNotFound, _
            MODULE_NAME & ".ApplyVendorTableFormatting", _
            "TableVendors belum diinisialisasi."
    End If

    If Not managerTable.HasRows(tbl) Then
        GoTo CleanUp
    End If

    ' =====================================================
    ' General formatting (every table column)
    ' =====================================================
    allColumns = Array( _
        COLUMN_ID, _
        COLUMN_NAME, _
        COLUMN_COMPANY_NAME, _
        COLUMN_CODE)

    For Each columnName In allColumns

        Set columnRange = managerTable.GetColumnRange( _
            tbl, CStr(columnName))

        If Not columnRange Is Nothing Then
            columnRange.VerticalAlignment = xlCenter
        End If

    Next columnName

    ' =====================================================
    ' ID formatting
    ' =====================================================
    managerTable.FormatColumn _
        tbl, COLUMN_ID, "0", xlCenter

    ' =====================================================
    ' Text formatting
    ' =====================================================
    managerTable.FormatColumn _
        tbl, COLUMN_NAME, "@"

    managerTable.FormatColumn _
        tbl, COLUMN_COMPANY_NAME, "@"

    managerTable.FormatColumn _
        tbl, COLUMN_CODE, "@", xlCenter

CleanUp:
    Set columnRange = Nothing
    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    Set columnRange = Nothing

    managerError.Rethrow errInfo, _
        MODULE_NAME & _
        ".ApplyVendorTableFormatting", _
        "Gagal menerapkan format TableVendors."

End Sub


' =========================================================
' SIMPAN button
' =========================================================
Public Sub SaveVendor_Click()

    On Error GoTo ErrorHandler

    Dim vendorItem As Vendor

    Dim vendorName As String
    Dim vendorCode As String
    Dim vendorCompanyName As String

    Dim confirmationMessage As String
    Dim successMessage As String

    Dim state As FormState
    Dim errInfo As ErrorInfo

    state = managerForm.BeginBusy()

    ' =====================================================
    ' Step 1: Read input values
    ' =====================================================
    vendorName = managerForm.GetCellText( _
        Me, CELL_VENDOR_NAME)

    vendorCode = managerForm.GetCellText( _
        Me, CELL_VENDOR_CODE)

    vendorCompanyName = managerForm.GetCellText( _
        Me, CELL_VENDOR_COMPANY_NAME)

    ' =====================================================
    ' Step 2: Validate input values
    ' =====================================================
    If managerValidationText.IsBlank(vendorName) Then
        managerError.ThrowError _
            sheetMasterDataVendor.EmptyVendorName, _
            MODULE_NAME & ".SaveVendor_Click", _
            "Nama vendor wajib diisi."
    End If

    If managerValidationText.IsBlank(vendorCode) Then
        managerError.ThrowError _
            sheetMasterDataVendor.EmptyVendorCode, _
            MODULE_NAME & ".SaveVendor_Click", _
            "Kode vendor wajib diisi."
    End If

    If managerValidationText.IsBlank(vendorCompanyName) Then
        managerError.ThrowError _
            sheetMasterDataVendor.EmptyVendorCompanyName, _
            MODULE_NAME & ".SaveVendor_Click", _
            "Nama perusahaan vendor wajib diisi."
    End If

    ' =====================================================
    ' Step 3: Build Vendor object
    ' =====================================================
    Set vendorItem = New Vendor

    With vendorItem
        .Name = vendorName
        .Code = vendorCode
        .CompanyName = vendorCompanyName
    End With

    ' =====================================================
    ' Step 4: Build confirmation message
    ' =====================================================
    confirmationMessage = _
        "Apakah Anda yakin ingin menyimpan " & _
        "vendor baru berikut?"

    confirmationMessage = confirmationMessage & _
        vbCrLf & vbCrLf & _
        "Nama: " & vendorItem.Name

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "Kode: " & vendorItem.Code

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "Nama perusahaan: " & _
        vendorItem.CompanyName

    If Not managerForm.Confirm( _
        confirmationMessage, _
        "Konfirmasi Simpan Vendor") Then

        GoTo CleanUp
    End If

    ' =====================================================
    ' Step 5: Set processing status
    ' =====================================================
    managerForm.SetStatus _
        "Menyimpan vendor baru...", True

    ' =====================================================
    ' Step 6: Call service procedure
    ' =====================================================
    serviceVendor.CreateVendor _
        vendorItem:=vendorItem

    managerForm.ClearStatus

    ' =====================================================
    ' Step 7: Build success message
    ' =====================================================
    successMessage = _
        "Vendor baru berhasil disimpan."

    successMessage = successMessage & _
        vbCrLf & vbCrLf & _
        "Nama: " & vendorItem.Name

    successMessage = successMessage & _
        vbCrLf & _
        "Kode: " & vendorItem.Code

    successMessage = successMessage & _
        vbCrLf & _
        "Nama perusahaan: " & _
        vendorItem.CompanyName

    managerForm.ShowInfo _
        successMessage, _
        "Penyimpanan Vendor Berhasil"

    ' =====================================================
    ' Step 8: Clear input form
    ' =====================================================
    ClearVendorInput

    ' =====================================================
    ' Step 9: Refresh TableVendors
    '
    ' RefreshVendors_Click calls SelectVendors and fills
    ' TableVendors with the latest database records.
    ' =====================================================
    RefreshVendors_Click

CleanUp:
    managerForm.EndBusy state
    managerForm.ClearStatus

    Set vendorItem = Nothing

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerForm.EndBusy state
    managerForm.ClearStatus

    Set vendorItem = Nothing

    managerError.ShowError errInfo, _
        "Gagal menyimpan vendor baru."

End Sub


' =========================================================
' Clear vendor input cells
' =========================================================
Private Sub ClearVendorInput()

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    managerForm.ClearCells Me, _
        CELL_VENDOR_NAME, _
        CELL_VENDOR_CODE, _
        CELL_VENDOR_COMPANY_NAME

    Me.Range(CELL_VENDOR_NAME).Select

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ClearVendorInput"

End Sub
