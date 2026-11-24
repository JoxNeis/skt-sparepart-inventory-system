Attribute VB_Name = "sMasterDataCatalogue"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

' Error codes: sheetMasterDataCatalogue (ErrorCodesSheets.bas)

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

    state = FormManager.BeginBusy( _
        "Mengambil data katalog...")

    ' =====================================================
    ' Step 1: Get and validate TableCatalogues
    ' =====================================================
    Set tbl = TableManager.GetTable(Me, TABLE_NAME)

    ValidateCatalogueTableColumns tbl

    ' =====================================================
    ' Step 2: Call service function
    ' =====================================================
    Set catalogues = serviceCatalogue.GetAllCatalogues()

    If catalogues Is Nothing Then
        ErrorManager.ThrowError _
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
            ErrorManager.ThrowError _
                sheetMasterDataCatalogue.NullCatalogueInCollection, _
                MODULE_NAME & ".RefreshCatalogues_Click", _
                "Objek Catalogue pada indeks " & _
                CStr(rowNumber) & _
                " tidak berhasil dibaca."
        End If

        rowIndex = TableManager.AddRow(tbl)

        WriteCatalogueTableRow _
            tbl:=tbl, _
            rowIndex:=rowIndex, _
            catalogueItem:=catalogueItem

    Next catalogueItem

    ' =====================================================
    ' Step 5: Apply table formatting
    ' =====================================================
    ApplyCatalogueTableFormatting tbl

    FormManager.ClearStatus

    ' =====================================================
    ' Step 6: Show result
    ' =====================================================
    If catalogueCount = 0 Then

        FormManager.ShowInfo _
            "Tidak ada data katalog yang ditemukan.", _
            "Master Data Katalog"

    Else

        FormManager.ShowInfo _
            CStr(catalogueCount) & _
            " data katalog berhasil ditampilkan.", _
            "Master Data Katalog"

    End If

CleanUp:
    Set catalogueItem = Nothing
    Set catalogues = Nothing
    Set tbl = Nothing

    FormManager.EndBusy state

    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    Set catalogueItem = Nothing
    Set catalogues = Nothing
    Set tbl = Nothing

    FormManager.EndBusy state

    ErrorManager.ShowError errInfo, _
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
        ErrorManager.ThrowError _
            TableManager.TableNotFound, _
            MODULE_NAME & ".WriteCatalogueTableRow", _
            "TableCatalogues belum diinisialisasi."
    End If

    If Not TableManager.IsValidRow(tbl, rowIndex) Then
        ErrorManager.ThrowError _
            TableManager.InvalidRow, _
            MODULE_NAME & ".WriteCatalogueTableRow", _
            "Target row katalog belum diinisialisasi."
    End If

    If catalogueItem Is Nothing Then
        ErrorManager.ThrowError _
            sheetMasterDataCatalogue.EmptyCatalogueItem, _
            MODULE_NAME & ".WriteCatalogueTableRow", _
            "Objek Catalogue tidak boleh kosong."
    End If

    ' =====================================================
    ' A: Id
    ' =====================================================
    TableManager.SetValue _
        tbl, rowIndex, COLUMN_ID, catalogueItem.Id

    ' =====================================================
    ' B: Nama
    ' =====================================================
    TableManager.SetValue _
        tbl, rowIndex, COLUMN_NAME, catalogueItem.Name

    ' =====================================================
    ' C: Alias
    ' =====================================================
    TableManager.SetValue _
        tbl, rowIndex, COLUMN_ALIAS, catalogueItem.Alias

    ' =====================================================
    ' D: Kode
    ' =====================================================
    TableManager.SetValue _
        tbl, rowIndex, COLUMN_CODE, catalogueItem.Code

    ' =====================================================
    ' E: Kategori
    ' =====================================================
    TableManager.SetValue _
        tbl, rowIndex, COLUMN_CATEGORY, catalogueItem.Category

    ' =====================================================
    ' F: Satuan
    ' =====================================================
    TableManager.SetValue _
        tbl, rowIndex, COLUMN_UOM, catalogueItem.unitOfMeasurement

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.Rethrow errInfo, _
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
        ErrorManager.ThrowError _
            TableManager.TableNotFound, _
            MODULE_NAME & ".ClearCatalogueTable", _
            "TableCatalogues belum diinisialisasi."
    End If

    TableManager.ClearRows tbl

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.Rethrow errInfo, _
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
        ErrorManager.ThrowError _
            TableManager.TableNotFound, _
            MODULE_NAME & ".ValidateCatalogueTableColumns", _
            "TableCatalogues tidak ditemukan " & _
            "pada worksheet ini."
    End If

    TableManager.ValidateColumns tbl, Array( _
        COLUMN_ID, _
        COLUMN_NAME, _
        COLUMN_ALIAS, _
        COLUMN_CODE, _
        COLUMN_CATEGORY, _
        COLUMN_UOM)

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.Rethrow errInfo, _
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
        ErrorManager.ThrowError _
            TableManager.TableNotFound, _
            MODULE_NAME & ".ApplyCatalogueTableFormatting", _
            "TableCatalogues belum diinisialisasi."
    End If

    If Not TableManager.HasRows(tbl) Then
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

        Set columnRange = TableManager.GetColumnRange( _
            tbl, CStr(columnName))

        If Not columnRange Is Nothing Then
            columnRange.VerticalAlignment = xlCenter
        End If

    Next columnName

    ' =====================================================
    ' ID formatting
    ' =====================================================
    TableManager.FormatColumn _
        tbl, COLUMN_ID, "0", xlCenter

    ' =====================================================
    ' Text columns
    ' =====================================================
    TableManager.FormatColumn _
        tbl, COLUMN_NAME, "@"

    TableManager.FormatColumn _
        tbl, COLUMN_ALIAS, "@"

    TableManager.FormatColumn _
        tbl, COLUMN_CODE, "@"

    TableManager.FormatColumn _
        tbl, COLUMN_CATEGORY, "@"

    TableManager.FormatColumn _
        tbl, COLUMN_UOM, "@", xlCenter

CleanUp:
    Set columnRange = Nothing
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    Set columnRange = Nothing

    ErrorManager.Rethrow errInfo, _
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

    state = FormManager.BeginBusy()

    ' =====================================================
    ' Step 1: Read input values
    ' =====================================================
    catalogueName = FormManager.GetCellText( _
        Me, CELL_NAME)

    catalogueAlias = FormManager.GetCellText( _
        Me, CELL_ALIAS)

    catalogueCode = FormManager.GetCellText( _
        Me, CELL_CODE)

    catalogueCategory = FormManager.GetCellText( _
        Me, CELL_CATEGORY)

    unitOfMeasurement = FormManager.GetCellText( _
        Me, CELL_UOM)

    ' =====================================================
    ' Step 2: Validate input values
    ' =====================================================
    If ValidationText.IsBlank(catalogueName) Then
        ErrorManager.ThrowError _
            sheetMasterDataCatalogue.EmptyCatalogueName, _
            MODULE_NAME & ".SaveCatalogue_Click", _
            "Nama katalog wajib diisi."
    End If

    If ValidationText.IsBlank(catalogueAlias) Then
        ErrorManager.ThrowError _
            sheetMasterDataCatalogue.EmptyCatalogueAlias, _
            MODULE_NAME & ".SaveCatalogue_Click", _
            "Alias katalog wajib diisi."
    End If

    If ValidationText.IsBlank(catalogueCode) Then
        ErrorManager.ThrowError _
            sheetMasterDataCatalogue.EmptyCatalogueCode, _
            MODULE_NAME & ".SaveCatalogue_Click", _
            "Kode katalog wajib diisi."
    End If

    If ValidationText.IsBlank(catalogueCategory) Then
        ErrorManager.ThrowError _
            sheetMasterDataCatalogue.EmptyCatalogueCategory, _
            MODULE_NAME & ".SaveCatalogue_Click", _
            "Kategori katalog wajib diisi."
    End If

    If ValidationText.IsBlank(unitOfMeasurement) Then
        ErrorManager.ThrowError _
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

    If Not FormManager.Confirm( _
        confirmationMessage, _
        "Konfirmasi Simpan Katalog") Then

        GoTo CleanUp
    End If

    ' =====================================================
    ' Step 5: Call service procedure
    ' =====================================================
    FormManager.SetStatus _
        "Menyimpan katalog baru...", True

    serviceCatalogue.CreateCatalogue _
        catalogueItem:=catalogueItem

    FormManager.ClearStatus

    ' =====================================================
    ' Step 6: Show success message
    ' =====================================================
    FormManager.ShowInfo _
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
    FormManager.EndBusy state

    Set catalogueItem = Nothing

    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    FormManager.EndBusy state

    Set catalogueItem = Nothing

    ErrorManager.ShowError errInfo, _
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

    state = FormManager.CaptureState()

    If Not FormManager.Confirm( _
        "Apakah Anda yakin ingin membersihkan " & _
        "seluruh input katalog?", _
        "Konfirmasi Refresh Form") Then

        GoTo CleanUp
    End If

    ' Events off while clearing; screen updating unchanged.
    state = FormManager.BeginBusy( _
        disableScreenUpdating:=False)

    ClearCatalogueInput

CleanUp:
    FormManager.EndBusy state

    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    FormManager.EndBusy state

    ErrorManager.ShowError errInfo, _
        "Gagal membersihkan form input katalog."

End Sub


' =========================================================
' Clear catalogue input cells
' =========================================================
Private Sub ClearCatalogueInput()

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    FormManager.ClearCells Me, _
        CELL_NAME, _
        CELL_ALIAS, _
        CELL_CODE, _
        CELL_CATEGORY, _
        CELL_UOM

    Me.Range(CELL_NAME).Select

    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".ClearCatalogueInput"

End Sub
