Attribute VB_Name = "sMasterDataCatalogue"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

' Error codes: sheetMasterDataCatalogue (managerErrorCodesSheets.bas)

Private Const MODULE_NAME As String = _
    "SHEET_MASTER_CATALOGUE"

Private Const TABLE_NAME As String = _
    "TableCatalogues"

' =========================================================
' Table column headers
'
' The text must exactly match the Excel table headers.
' =========================================================
Private Const COLUMN_ID As String = "Id"
Private Const COLUMN_NAME As String = "Nama"
Private Const COLUMN_ALIAS As String = "Alias"
Private Const COLUMN_CODE As String = "Kode"
Private Const COLUMN_CATEGORY As String = "Kategori"
Private Const COLUMN_UOM As String = "Satuan"

' =========================================================
' Input cells
' =========================================================
Private Const CELL_NAME As String = "B4"
Private Const CELL_ALIAS As String = "B5"
Private Const CELL_CODE As String = "B6"
Private Const CELL_CATEGORY As String = "B7"
Private Const CELL_UOM As String = "B8"


' =========================================================
' REFRESH button procedure
' =========================================================
Public Sub RefreshCatalogues_Click()

    On Error GoTo ErrorHandler

    Dim tbl As ListObject
    Dim rowIndex As Long

    Dim catalogues As Collection
    Dim catalogueItem As Catalogue

    Dim catalogueCount As Long
    Dim rowNumber As Long

    Dim state As FormState
    Dim errInfo As ErrorInfo

    state = managerForm.BeginBusy( _
        "Mengambil data katalog...")

    ' =====================================================
    ' Step 1: Get and validate TableCatalogues
    ' =====================================================
    Set tbl = managerTable.GetTable(Me, TABLE_NAME)

    ValidateCatalogueTableColumns tbl

    ' =====================================================
    ' Step 2: Call service function
    ' =====================================================
    Set catalogues = serviceCatalogue.GetAllCatalogues()

    If catalogues Is Nothing Then
        managerError.ThrowError _
            sheetMasterDataCatalogue.NullCatalogueCollection, _
            MODULE_NAME & ".RefreshCatalogues_Click", _
            "DAO tidak mengembalikan collection katalog."
    End If

    catalogueCount = catalogues.Count

    ' =====================================================
    ' Step 3: Remove existing table rows
    ' =====================================================
    ClearCatalogueTable tbl

    ' =====================================================
    ' Step 4: Fill TableCatalogues
    ' =====================================================
    rowNumber = 0

    For Each catalogueItem In catalogues

        rowNumber = rowNumber + 1

        If catalogueItem Is Nothing Then
            managerError.ThrowError _
                sheetMasterDataCatalogue.NullCatalogueInCollection, _
                MODULE_NAME & ".RefreshCatalogues_Click", _
                "Objek Catalogue pada indeks " & _
                CStr(rowNumber) & _
                " tidak berhasil dibaca."
        End If

        rowIndex = managerTable.AddRow(tbl)

        WriteCatalogueTableRow _
            tbl:=tbl, _
            rowIndex:=rowIndex, _
            catalogueItem:=catalogueItem

    Next catalogueItem

    ' =====================================================
    ' Step 5: Apply table formatting
    ' =====================================================
    ApplyCatalogueTableFormatting tbl

    managerForm.ClearStatus

    ' =====================================================
    ' Step 6: Show result
    ' =====================================================
    If catalogueCount = 0 Then

        managerForm.ShowInfo _
            "Tidak ada data katalog yang ditemukan.", _
            "Master Data Katalog"

    Else

        managerForm.ShowInfo _
            CStr(catalogueCount) & _
            " data katalog berhasil ditampilkan.", _
            "Master Data Katalog"

    End If

CleanUp:
    Set catalogueItem = Nothing
    Set catalogues = Nothing
    Set tbl = Nothing

    managerForm.EndBusy state

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    Set catalogueItem = Nothing
    Set catalogues = Nothing
    Set tbl = Nothing

    managerForm.EndBusy state

    managerError.ShowError errInfo, _
        "Gagal menampilkan master data katalog."

End Sub


' =========================================================
' Write one Catalogue object into one table row
' =========================================================
Private Sub WriteCatalogueTableRow( _
    ByVal tbl As ListObject, _
    ByVal rowIndex As Long, _
    ByVal catalogueItem As Catalogue)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    If tbl Is Nothing Then
        managerError.ThrowError _
            managerTable.TableNotFound, _
            MODULE_NAME & ".WriteCatalogueTableRow", _
            "TableCatalogues belum diinisialisasi."
    End If

    If Not managerTable.IsValidRow(tbl, rowIndex) Then
        managerError.ThrowError _
            managerTable.InvalidRow, _
            MODULE_NAME & ".WriteCatalogueTableRow", _
            "Target row katalog belum diinisialisasi."
    End If

    If catalogueItem Is Nothing Then
        managerError.ThrowError _
            sheetMasterDataCatalogue.EmptyCatalogueItem, _
            MODULE_NAME & ".WriteCatalogueTableRow", _
            "Objek Catalogue tidak boleh kosong."
    End If

    ' =====================================================
    ' A: Id
    ' =====================================================
    managerTable.SetValue _
        tbl, rowIndex, COLUMN_ID, catalogueItem.Id

    ' =====================================================
    ' B: Nama
    ' =====================================================
    managerTable.SetValue _
        tbl, rowIndex, COLUMN_NAME, catalogueItem.Name

    ' =====================================================
    ' C: Alias
    ' =====================================================
    managerTable.SetValue _
        tbl, rowIndex, COLUMN_ALIAS, catalogueItem.Alias

    ' =====================================================
    ' D: Kode
    ' =====================================================
    managerTable.SetValue _
        tbl, rowIndex, COLUMN_CODE, catalogueItem.Code

    ' =====================================================
    ' E: Kategori
    ' =====================================================
    managerTable.SetValue _
        tbl, rowIndex, COLUMN_CATEGORY, catalogueItem.Category

    ' =====================================================
    ' F: Satuan
    ' =====================================================
    managerTable.SetValue _
        tbl, rowIndex, COLUMN_UOM, catalogueItem.unitOfMeasurement

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".WriteCatalogueTableRow", _
        "Gagal menulis satu data katalog ke tabel."

End Sub


' =========================================================
' Clear existing rows while preserving the table
' =========================================================
Private Sub ClearCatalogueTable( _
    ByVal tbl As ListObject)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    If tbl Is Nothing Then
        managerError.ThrowError _
            managerTable.TableNotFound, _
            MODULE_NAME & ".ClearCatalogueTable", _
            "TableCatalogues belum diinisialisasi."
    End If

    managerTable.ClearRows tbl

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ClearCatalogueTable", _
        "Gagal membersihkan data katalog sebelumnya."

End Sub


' =========================================================
' Validate required table columns
' =========================================================
Private Sub ValidateCatalogueTableColumns( _
    ByVal tbl As ListObject)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    If tbl Is Nothing Then
        managerError.ThrowError _
            managerTable.TableNotFound, _
            MODULE_NAME & ".ValidateCatalogueTableColumns", _
            "TableCatalogues tidak ditemukan " & _
            "pada worksheet ini."
    End If

    managerTable.ValidateColumns tbl, Array( _
        COLUMN_ID, _
        COLUMN_NAME, _
        COLUMN_ALIAS, _
        COLUMN_CODE, _
        COLUMN_CATEGORY, _
        COLUMN_UOM)

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ValidateCatalogueTableColumns", _
        "Struktur TableCatalogues tidak sesuai."

End Sub


' =========================================================
' Apply table formatting
' =========================================================
Private Sub ApplyCatalogueTableFormatting( _
    ByVal tbl As ListObject)

    On Error GoTo ErrorHandler

    Dim allColumns As Variant
    Dim columnName As Variant
    Dim columnRange As Range

    Dim errInfo As ErrorInfo

    If tbl Is Nothing Then
        managerError.ThrowError _
            managerTable.TableNotFound, _
            MODULE_NAME & ".ApplyCatalogueTableFormatting", _
            "TableCatalogues belum diinisialisasi."
    End If

    If Not managerTable.HasRows(tbl) Then
        GoTo CleanUp
    End If

    ' =====================================================
    ' General vertical alignment (every table column)
    ' =====================================================
    allColumns = Array( _
        COLUMN_ID, _
        COLUMN_NAME, _
        COLUMN_ALIAS, _
        COLUMN_CODE, _
        COLUMN_CATEGORY, _
        COLUMN_UOM)

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
    ' Text columns
    ' =====================================================
    managerTable.FormatColumn _
        tbl, COLUMN_NAME, "@"

    managerTable.FormatColumn _
        tbl, COLUMN_ALIAS, "@"

    managerTable.FormatColumn _
        tbl, COLUMN_CODE, "@"

    managerTable.FormatColumn _
        tbl, COLUMN_CATEGORY, "@"

    managerTable.FormatColumn _
        tbl, COLUMN_UOM, "@", xlCenter

CleanUp:
    Set columnRange = Nothing
    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    Set columnRange = Nothing

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ApplyCatalogueTableFormatting", _
        "Gagal menerapkan format TableCatalogues."

End Sub



' =========================================================
' SIMPAN button
' =========================================================
Public Sub SaveCatalogue_Click()

    On Error GoTo ErrorHandler

    Dim catalogueItem As Catalogue

    Dim catalogueName As String
    Dim catalogueAlias As String
    Dim catalogueCode As String
    Dim catalogueCategory As String
    Dim unitOfMeasurement As String

    Dim confirmationMessage As String

    Dim state As FormState
    Dim errInfo As ErrorInfo

    state = managerForm.BeginBusy()

    ' =====================================================
    ' Step 1: Read input values
    ' =====================================================
    catalogueName = managerForm.GetCellText( _
        Me, CELL_NAME)

    catalogueAlias = managerForm.GetCellText( _
        Me, CELL_ALIAS)

    catalogueCode = managerForm.GetCellText( _
        Me, CELL_CODE)

    catalogueCategory = managerForm.GetCellText( _
        Me, CELL_CATEGORY)

    unitOfMeasurement = managerForm.GetCellText( _
        Me, CELL_UOM)

    ' =====================================================
    ' Step 2: Validate input values
    ' =====================================================
    If managerValidationText.IsBlank(catalogueName) Then
        managerError.ThrowError _
            sheetMasterDataCatalogue.EmptyCatalogueName, _
            MODULE_NAME & ".SaveCatalogue_Click", _
            "Nama katalog wajib diisi."
    End If

    If managerValidationText.IsBlank(catalogueAlias) Then
        managerError.ThrowError _
            sheetMasterDataCatalogue.EmptyCatalogueAlias, _
            MODULE_NAME & ".SaveCatalogue_Click", _
            "Alias katalog wajib diisi."
    End If

    If managerValidationText.IsBlank(catalogueCode) Then
        managerError.ThrowError _
            sheetMasterDataCatalogue.EmptyCatalogueCode, _
            MODULE_NAME & ".SaveCatalogue_Click", _
            "Kode katalog wajib diisi."
    End If

    If managerValidationText.IsBlank(catalogueCategory) Then
        managerError.ThrowError _
            sheetMasterDataCatalogue.EmptyCatalogueCategory, _
            MODULE_NAME & ".SaveCatalogue_Click", _
            "Kategori katalog wajib diisi."
    End If

    If managerValidationText.IsBlank(unitOfMeasurement) Then
        managerError.ThrowError _
            sheetMasterDataCatalogue.EmptyCatalogueUnitOfMeasurement, _
            MODULE_NAME & ".SaveCatalogue_Click", _
            "Satuan katalog wajib diisi."
    End If

    ' =====================================================
    ' Step 3: Build Catalogue object
    ' =====================================================
    Set catalogueItem = New Catalogue

    With catalogueItem
        .Name = catalogueName
        .Alias = catalogueAlias
        .Code = catalogueCode
        .Category = catalogueCategory
        .unitOfMeasurement = unitOfMeasurement
    End With

    ' =====================================================
    ' Step 4: Build confirmation message
    ' =====================================================
    confirmationMessage = _
        "Apakah Anda yakin ingin menyimpan " & _
        "katalog baru berikut?"

    confirmationMessage = confirmationMessage & _
        vbCrLf & vbCrLf & _
        "Nama: " & catalogueItem.Name

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "Alias: " & catalogueItem.Alias

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "Kode: " & catalogueItem.Code

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "Kategori: " & catalogueItem.Category

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "Satuan: " & _
        catalogueItem.unitOfMeasurement

    If Not managerForm.Confirm( _
        confirmationMessage, _
        "Konfirmasi Simpan Katalog") Then

        GoTo CleanUp
    End If

    ' =====================================================
    ' Step 5: Call service procedure
    ' =====================================================
    managerForm.SetStatus _
        "Menyimpan katalog baru...", True

    serviceCatalogue.CreateCatalogue _
        catalogueItem:=catalogueItem

    managerForm.ClearStatus

    ' =====================================================
    ' Step 6: Show success message
    ' =====================================================
    managerForm.ShowInfo _
        "Katalog baru berhasil disimpan." & _
        vbCrLf & vbCrLf & _
        "Nama: " & catalogueItem.Name & vbCrLf & _
        "Kode: " & catalogueItem.Code, _
        "Penyimpanan Berhasil"

    ' =====================================================
    ' Step 7: Clear inputs
    ' =====================================================
    ClearCatalogueInput

    ' =====================================================
    ' Step 8: Refresh catalogue named range
    '
    ' This makes the new catalogue available immediately
    ' in dropdowns that use listCatalogues.
    ' Replace modListCatalogue with your actual module name.
    ' =====================================================
    On Error Resume Next

    sMasterDataCatalogue.RefreshCatalogues_Click

    On Error GoTo ErrorHandler

CleanUp:
    managerForm.EndBusy state

    Set catalogueItem = Nothing

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerForm.EndBusy state

    Set catalogueItem = Nothing

    managerError.ShowError errInfo, _
        "Gagal menyimpan katalog baru."

End Sub


' =========================================================
' REFRESH button
'
' Clears all input fields so the form is ready for a new
' catalogue entry.
' =========================================================
Public Sub RefreshCatalogueInput_Click()

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo

    state = managerForm.CaptureState()

    If Not managerForm.Confirm( _
        "Apakah Anda yakin ingin membersihkan " & _
        "seluruh input katalog?", _
        "Konfirmasi Refresh Form") Then

        GoTo CleanUp
    End If

    ' Events off while clearing; screen updating unchanged.
    state = managerForm.BeginBusy( _
        disableScreenUpdating:=False)

    ClearCatalogueInput

CleanUp:
    managerForm.EndBusy state

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerForm.EndBusy state

    managerError.ShowError errInfo, _
        "Gagal membersihkan form input katalog."

End Sub


' =========================================================
' Clear catalogue input cells
' =========================================================
Private Sub ClearCatalogueInput()

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    managerForm.ClearCells Me, _
        CELL_NAME, _
        CELL_ALIAS, _
        CELL_CODE, _
        CELL_CATEGORY, _
        CELL_UOM

    Me.Range(CELL_NAME).Select

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ClearCatalogueInput"

End Sub
