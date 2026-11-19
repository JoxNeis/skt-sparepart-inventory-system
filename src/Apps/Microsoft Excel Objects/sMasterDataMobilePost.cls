Attribute VB_Name = "sMasterDataMobilePost"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

' Error codes: sheetMasterDataMobilePost (ErrorCodesSheets.bas)

Private Const MODULE_NAME As String = _
    "SHEET_MASTER_MOBILE_POST"

Private Const TABLE_NAME As String = _
    "TableMobilePosts"

Private Const CELL_MANUFACTURING_PLANT As String = _
    "B3"

Private Const CELL_MOBILE_POST_NAME As String = "B4"
Private Const CELL_MOBILE_POST_CODE As String = "B5"
Private Const COLUMN_NAME As String = _
    "Nama"


' =========================================================
' SIMPAN button
' =========================================================
Public Sub SaveMobilePost_Click()

    On Error GoTo ErrorHandler

    Dim plantItem As Plant
    Dim mobilePostItem As MobilePost

    Dim manufacturingPlantName As String
    Dim mobilePostName As String
    Dim mobilePostCode As String

    Dim confirmationMessage As String
    Dim successMessage As String

    Dim state As FormState
    Dim errInfo As ErrorInfo

    state = FormManager.BeginBusy()

    ' =====================================================
    ' Step 1: Read input values
    ' =====================================================
    manufacturingPlantName = FormManager.GetCellText( _
        Me, CELL_MANUFACTURING_PLANT)

    mobilePostName = FormManager.GetCellText( _
        Me, CELL_MOBILE_POST_NAME)

    mobilePostCode = FormManager.GetCellText( _
        Me, CELL_MOBILE_POST_CODE)

    ' =====================================================
    ' Step 3: Build Plant object
    ' =====================================================
    Set plantItem = New Plant

    plantItem.Name = manufacturingPlantName

    ' =====================================================
    ' Step 4: Build MobilePost object
    ' =====================================================
    Set mobilePostItem = New MobilePost

    With mobilePostItem
        .Name = mobilePostName
        .Code = mobilePostCode

        Set .Plant = plantItem
    End With

    ' =====================================================
    ' Step 5: Build confirmation message
    ' =====================================================
    confirmationMessage = _
        "Apakah Anda yakin ingin menyimpan " & _
        "mobile post baru berikut?"

    confirmationMessage = confirmationMessage & _
        vbCrLf & vbCrLf & _
        "Lokasi: " & _
        mobilePostItem.Plant.Name

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "Nama mobile post: " & _
        mobilePostItem.Name

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "Kode mobile post: " & _
        mobilePostItem.Code

    If Not FormManager.Confirm( _
        confirmationMessage, _
        "Konfirmasi Simpan Mobile Post") Then

        GoTo CleanUp
    End If

    ' =====================================================
    ' Step 6: Set processing status
    ' =====================================================
    FormManager.SetStatus _
        "Menyimpan mobile post baru...", True

    ' =====================================================
    ' Step 7: Call service procedure
    ' =====================================================
    serviceMobilePost.CreateMobilePost _
        mobilePostItem:=mobilePostItem

    FormManager.ClearStatus

    ' =====================================================
    ' Step 8: Build success message
    ' =====================================================
    successMessage = _
        "Mobile post baru berhasil disimpan."

    successMessage = successMessage & _
        vbCrLf & vbCrLf & _
        "Lokasi: " & _
        mobilePostItem.Plant.Name

    successMessage = successMessage & _
        vbCrLf & _
        "Nama mobile post: " & _
        mobilePostItem.Name

    successMessage = successMessage & _
        vbCrLf & _
        "Kode mobile post: " & _
        mobilePostItem.Code

    FormManager.ShowInfo _
        successMessage, _
        "Penyimpanan Mobile Post Berhasil"

    ' =====================================================
    ' Step 9: Clear mobile-post inputs
    '
    ' The manufacturing plant is retained because the user
    ' may want to create another mobile post for the same
    ' plant.
    ' =====================================================
    ClearMobilePostInput

    ' =====================================================
    ' Step 10: Refresh TableMobilePosts
    '
    ' This calls:
    ' serviceMobilePost.SelectMobilePostByManufacturingPlantName
    ' and fills TableMobilePosts.
    ' =====================================================
    RefreshMobilePosts_Click

CleanUp:
    FormManager.EndBusy state

    Set mobilePostItem = Nothing
    Set plantItem = Nothing

    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    FormManager.EndBusy state

    Set mobilePostItem = Nothing
    Set plantItem = Nothing

    ErrorManager.ShowError errInfo, _
        "Gagal menyimpan mobile post baru."

End Sub


' =========================================================
' Clear mobile-post input cells
'
' B3 is not cleared because it contains the selected
' manufacturing plant.
' =========================================================
Private Sub ClearMobilePostInput()

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    FormManager.ClearCells Me, _
        CELL_MOBILE_POST_NAME, _
        CELL_MOBILE_POST_CODE

    Me.Range( _
        CELL_MOBILE_POST_NAME _
    ).Select

    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".ClearMobilePostInput"

End Sub


' =========================================================
' REFRESH button
' =========================================================
Public Sub RefreshMobilePosts_Click()

    On Error GoTo ErrorHandler

    Dim tbl As ListObject
    Dim rowIndex As Long

    Dim mobilePosts As Collection
    Dim mobilePostItem As MobilePost

    Dim manufacturingPlantName As String
    Dim mobilePostName As String

    Dim mobilePostCount As Long
    Dim rowNumber As Long

    Dim state As FormState
    Dim errInfo As ErrorInfo

    state = FormManager.BeginBusy( _
        "Mengambil data mobile post...")

    ' =====================================================
    ' Step 1: Read manufacturing plant from B3
    ' =====================================================
    manufacturingPlantName = FormManager.GetCellText( _
        Me, CELL_MANUFACTURING_PLANT)

    If ValidationText.IsBlank(manufacturingPlantName) Then
        ErrorManager.ThrowError _
            sheetMasterDataMobilePost.EmptyManufacturingPlantName, _
            MODULE_NAME & ".RefreshMobilePosts_Click", _
            "Lokasi manufacturing plant " & _
            "tidak boleh kosong."
    End If

    ' =====================================================
    ' Step 2: Get and validate TableMobilePosts
    ' =====================================================
    Set tbl = TableManager.GetTable(Me, TABLE_NAME)

    ValidateMobilePostTableColumns tbl

    ' =====================================================
    ' Step 3: Call service function
    ' =====================================================
    Set mobilePosts = _
        serviceMobilePost. _
        SelectMobilePostByManufacturingPlantName( _
            manufacturingPlantName:= _
                manufacturingPlantName)

    If mobilePosts Is Nothing Then
        ErrorManager.ThrowError _
            sheetMasterDataMobilePost.NullMobilePostCollection, _
            MODULE_NAME & ".RefreshMobilePosts_Click", _
            "DAO tidak mengembalikan collection " & _
            "mobile post."
    End If

    mobilePostCount = mobilePosts.Count

    ' =====================================================
    ' Step 4: Clear existing rows
    ' =====================================================
    ClearMobilePostTable tbl

    ' =====================================================
    ' Step 5: Fill TableMobilePosts
    ' =====================================================
    rowNumber = 0

    For Each mobilePostItem In mobilePosts

        rowNumber = rowNumber + 1

        If mobilePostItem Is Nothing Then
            ErrorManager.ThrowError _
                sheetMasterDataMobilePost.NullMobilePostInCollection, _
                MODULE_NAME & ".RefreshMobilePosts_Click", _
                "Objek MobilePost pada indeks " & _
                CStr(rowNumber) & _
                " tidak berhasil dibaca."
        End If

        mobilePostName = TrimText( _
            mobilePostItem.Name)

        If ValidationText.IsBlank(mobilePostName) Then
            ErrorManager.ThrowError _
                sheetMasterDataMobilePost.EmptyMobilePostNameInCollection, _
                MODULE_NAME & ".RefreshMobilePosts_Click", _
                "Nama mobile post pada indeks " & _
                CStr(rowNumber) & _
                " tidak boleh kosong."
        End If

        rowIndex = TableManager.AddRow(tbl)

        TableManager.SetValue _
            tbl, rowIndex, COLUMN_NAME, mobilePostName

    Next mobilePostItem

    ' =====================================================
    ' Step 6: Apply formatting
    ' =====================================================
    ApplyMobilePostTableFormatting tbl

    FormManager.ClearStatus

    ' =====================================================
    ' Step 7: Show result
    ' =====================================================
    If mobilePostCount = 0 Then

        FormManager.ShowInfo _
            "Tidak ada mobile post untuk " & _
            "manufacturing plant '" & _
            manufacturingPlantName & "'.", _
            "Master Data Mobile Post"

    Else

        FormManager.ShowInfo _
            CStr(mobilePostCount) & _
            " data mobile post berhasil ditampilkan.", _
            "Master Data Mobile Post"

    End If

CleanUp:
    Set mobilePostItem = Nothing
    Set mobilePosts = Nothing
    Set tbl = Nothing

    FormManager.EndBusy state

    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    Set mobilePostItem = Nothing
    Set mobilePosts = Nothing
    Set tbl = Nothing

    FormManager.EndBusy state

    ErrorManager.ShowError errInfo, _
        "Gagal menampilkan master data mobile post."

End Sub


' =========================================================
' Clear existing mobile-post table rows
' =========================================================
Private Sub ClearMobilePostTable( _
    ByVal tbl As ListObject)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    If tbl Is Nothing Then
        ErrorManager.ThrowError _
            TableManager.TableNotFound, _
            MODULE_NAME & ".ClearMobilePostTable", _
            "TableMobilePosts belum diinisialisasi."
    End If

    TableManager.ClearRows tbl

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".ClearMobilePostTable", _
        "Gagal membersihkan data mobile post lama."

End Sub


' =========================================================
' Validate TableMobilePosts structure
' =========================================================
Private Sub ValidateMobilePostTableColumns( _
    ByVal tbl As ListObject)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    If tbl Is Nothing Then
        ErrorManager.ThrowError _
            TableManager.TableNotFound, _
            MODULE_NAME & ".ValidateMobilePostTableColumns", _
            "TableMobilePosts tidak ditemukan " & _
            "pada worksheet ini."
    End If

    TableManager.ValidateColumns tbl, Array( _
        COLUMN_NAME)

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".ValidateMobilePostTableColumns", _
        "Struktur TableMobilePosts tidak sesuai."

End Sub


' =========================================================
' Apply mobile-post table formatting
' =========================================================
Private Sub ApplyMobilePostTableFormatting( _
    ByVal tbl As ListObject)

    On Error GoTo ErrorHandler

    Dim columnRange As Range

    Dim errInfo As ErrorInfo

    If tbl Is Nothing Then
        ErrorManager.ThrowError _
            TableManager.TableNotFound, _
            MODULE_NAME & ".ApplyMobilePostTableFormatting", _
            "TableMobilePosts belum diinisialisasi."
    End If

    If Not TableManager.HasRows(tbl) Then
        GoTo CleanUp
    End If

    ' TableMobilePosts has a single column (Nama), so the
    ' column range is the whole table body.
    Set columnRange = TableManager.GetColumnRange( _
        tbl, COLUMN_NAME)

    If Not columnRange Is Nothing Then
        columnRange.VerticalAlignment = xlCenter
    End If

    TableManager.FormatColumn _
        tbl, COLUMN_NAME, "@", xlLeft

CleanUp:
    Set columnRange = Nothing
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    Set columnRange = Nothing

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".ApplyMobilePostTableFormatting", _
        "Gagal menerapkan format " & _
        "TableMobilePosts."

End Sub
