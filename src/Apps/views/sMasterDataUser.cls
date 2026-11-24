Attribute VB_Name = "sMasterDataUser"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Const MODULE_NAME As String = _
    "SHEET_MASTER_USER"

Private Const TABLE_NAME As String = _
    "TableUser"

' =========================================================
' Table column headers
'
' These names must exactly match the Excel table headers.
' =========================================================
Private Const COLUMN_ID As String = "Id"
Private Const COLUMN_FULL_NAME As String = "Nama"

Private Const COLUMN_USERNAME As String = _
    "Username PMI"

Private Const COLUMN_PLANT As String = "Lokasi"
Private Const COLUMN_POSITION As String = "Posisi"
Private Const COLUMN_TIER As String = "Tier"

' =========================================================
' User input cells
' =========================================================
Private Const CELL_USER_PLANT As String = "B2"
Private Const CELL_USER_USERNAME As String = "B3"
Private Const CELL_USER_FIRST_NAME As String = "B4"
Private Const CELL_USER_MIDDLE_NAME As String = "B5"
Private Const CELL_USER_LAST_NAME As String = "B6"
Private Const CELL_USER_POSITION As String = "B7"
Private Const CELL_USER_TIER As String = "B8"


' =========================================================
' SIMPAN button
' =========================================================
Public Sub SaveUser_Click()

    On Error GoTo ErrorHandler

    Dim manufacturingPlantName As String
    Dim username As String
    Dim firstName As String
    Dim middleName As String
    Dim lastName As String
    Dim userPosition As String

    Dim tierValue As Variant
    Dim tier As Long

    Dim fullName As String
    Dim confirmationMessage As String
    Dim successMessage As String

    Dim state As FormState
    Dim errInfo As ErrorInfo

    state = FormManager.BeginBusy()

    ' =====================================================
    ' Step 1: Read user input
    ' =====================================================
    manufacturingPlantName = FormManager.GetCellText( _
        Me, CELL_USER_PLANT)

    username = FormManager.GetCellText( _
        Me, CELL_USER_USERNAME)

    firstName = FormManager.GetCellText( _
        Me, CELL_USER_FIRST_NAME)

    middleName = FormManager.GetCellText( _
        Me, CELL_USER_MIDDLE_NAME)

    lastName = FormManager.GetCellText( _
        Me, CELL_USER_LAST_NAME)

    userPosition = FormManager.GetCellText( _
        Me, CELL_USER_POSITION)

    tierValue = FormManager.GetCellValue( _
        Me, CELL_USER_TIER)

    ' =====================================================
    ' Step 2: Validate required text inputs
    ' =====================================================
    If ValidationText.IsBlank(manufacturingPlantName) Then
        ErrorManager.ThrowError _
            sheetMasterDataUser.EmptyPlantName, _
            MODULE_NAME & ".SaveUser_Click", _
            "Lokasi manufacturing plant wajib diisi."
    End If

    If ValidationText.IsBlank(username) Then
        ErrorManager.ThrowError _
            sheetMasterDataUser.EmptyUsername, _
            MODULE_NAME & ".SaveUser_Click", _
            "Username PMI wajib diisi."
    End If

    If ValidationText.IsBlank(firstName) Then
        ErrorManager.ThrowError _
            sheetMasterDataUser.EmptyFirstName, _
            MODULE_NAME & ".SaveUser_Click", _
            "Nama depan wajib diisi."
    End If

    ' Middle name is optional.

    If ValidationText.IsBlank(lastName) Then
        ErrorManager.ThrowError _
            sheetMasterDataUser.EmptyLastName, _
            MODULE_NAME & ".SaveUser_Click", _
            "Nama belakang wajib diisi."
    End If

    If ValidationText.IsBlank(userPosition) Then
        ErrorManager.ThrowError _
            sheetMasterDataUser.EmptyPosition, _
            MODULE_NAME & ".SaveUser_Click", _
            "Posisi user wajib diisi."
    End If

    ' =====================================================
    ' Step 3: Validate tier
    ' =====================================================
    If IsError(tierValue) Then
        ErrorManager.ThrowError _
            sheetMasterDataUser.InvalidTierErrorValue, _
            MODULE_NAME & ".SaveUser_Click", _
            "Tier berisi nilai error."
    End If

    ' Null, Empty and text with only spaces.
    If ValidationText.IsBlank(tierValue) Then
        ErrorManager.ThrowError _
            sheetMasterDataUser.EmptyTier, _
            MODULE_NAME & ".SaveUser_Click", _
            "Tier user wajib diisi."
    End If

    If ValidationNumeric.IsNotNumber(tierValue) Then
        ErrorManager.ThrowError _
            sheetMasterDataUser.InvalidTierNumber, _
            MODULE_NAME & ".SaveUser_Click", _
            "Tier user harus berupa angka."
    End If

    If Not ValidationNumeric.IsWholeNumber(CDbl(tierValue)) Then
        ErrorManager.ThrowError _
            sheetMasterDataUser.InvalidTierWholeNumber, _
            MODULE_NAME & ".SaveUser_Click", _
            "Tier user harus berupa bilangan bulat."
    End If

    tier = CLng(tierValue)

    If ValidationNumeric.IsZeroOrNegative(tier) Then
        ErrorManager.ThrowError _
            sheetMasterDataUser.ZeroOrNegativeTier, _
            MODULE_NAME & ".SaveUser_Click", _
            "Tier user harus lebih besar dari nol."
    End If

    ' =====================================================
    ' Step 4: Build full name for confirmation
    ' =====================================================
    fullName = BuildInputUserFullName( _
        firstName:=firstName, _
        middleName:=middleName, _
        lastName:=lastName)

    ' =====================================================
    ' Step 5: Build confirmation message
    ' =====================================================
    confirmationMessage = _
        "Apakah Anda yakin ingin membuat user baru berikut?"

    confirmationMessage = confirmationMessage & _
        vbCrLf & vbCrLf & _
        "Lokasi: " & manufacturingPlantName

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "Username PMI: " & username

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "Nama: " & fullName

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "Posisi: " & userPosition

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "Tier: " & CStr(tier)

    If Not FormManager.Confirm( _
        confirmationMessage, _
        "Konfirmasi Simpan User") Then

        GoTo CleanUp
    End If

    ' =====================================================
    ' Step 6: Call service procedure
    ' =====================================================
    FormManager.SetStatus _
        "Menyimpan user baru...", True

    serviceUser.CreateUser _
        manufacturingPlantName:=manufacturingPlantName, _
        username:=username, _
        firstName:=firstName, _
        middleName:=middleName, _
        lastName:=lastName, _
        position:=userPosition, _
        tier:=tier

    FormManager.ClearStatus

    ' =====================================================
    ' Step 7: Show success message
    ' =====================================================
    successMessage = _
        "User baru berhasil disimpan."

    successMessage = successMessage & _
        vbCrLf & vbCrLf & _
        "Username PMI: " & username

    successMessage = successMessage & _
        vbCrLf & _
        "Nama: " & fullName

    successMessage = successMessage & _
        vbCrLf & _
        "Lokasi: " & manufacturingPlantName

    successMessage = successMessage & _
        vbCrLf & _
        "Posisi: " & userPosition

    successMessage = successMessage & _
        vbCrLf & _
        "Tier: " & CStr(tier)

    FormManager.ShowInfo _
        successMessage, _
        "Penyimpanan User Berhasil"

    ' =====================================================
    ' Step 8: Clear form
    '
    ' Location in B2 is preserved.
    ' =====================================================
    ClearUserInput

    ' =====================================================
    ' Step 9: Refresh TableUsers
    ' =====================================================
    RefreshUsers_Click

CleanUp:
    FormManager.EndBusy state
    FormManager.ClearStatus

    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    FormManager.EndBusy state
    FormManager.ClearStatus

    ErrorManager.ShowError errInfo, _
        "Gagal menyimpan user baru."

End Sub


' =========================================================
' Clear user input
'
' B2 is preserved because it contains the selected
' manufacturing plant.
' =========================================================
Private Sub ClearUserInput()

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    FormManager.ClearCells Me, _
        CELL_USER_USERNAME, _
        CELL_USER_FIRST_NAME, _
        CELL_USER_MIDDLE_NAME, _
        CELL_USER_LAST_NAME, _
        CELL_USER_POSITION, _
        CELL_USER_TIER

    Me.Range( _
        CELL_USER_USERNAME _
    ).Select

    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".ClearUserInput"

End Sub


' =========================================================
' Build full name from form input
' =========================================================
Private Function BuildInputUserFullName( _
    ByVal firstName As String, _
    ByVal middleName As String, _
    ByVal lastName As String) As String

    Dim fullName As String

    firstName = TrimText(firstName)
    middleName = TrimText(middleName)
    lastName = TrimText(lastName)

    If Len(firstName) > 0 Then
        fullName = firstName
    End If

    If Len(middleName) > 0 Then

        If Len(fullName) > 0 Then
            fullName = fullName & " "
        End If

        fullName = fullName & middleName

    End If

    If Len(lastName) > 0 Then

        If Len(fullName) > 0 Then
            fullName = fullName & " "
        End If

        fullName = fullName & lastName

    End If

    BuildInputUserFullName = _
        TrimText(fullName)

End Function
' =========================================================
' REFRESH button
' =========================================================
Public Sub RefreshUsers_Click()

    On Error GoTo ErrorHandler

    Dim tbl As ListObject
    Dim rowIndex As Long

    Dim rs As ADODB.recordset

    Dim fullName As String
    Dim rowNumber As Long

    Dim state As FormState
    Dim errInfo As ErrorInfo

    state = FormManager.BeginBusy( _
        "Mengambil data user...")

    ' =====================================================
    ' Step 1: Get and validate TableUsers
    ' =====================================================
    Set tbl = TableManager.GetTable(Me, TABLE_NAME)

    ValidateUserTableColumns tbl

    ' =====================================================
    ' Step 2: Call service function
    ' =====================================================
    Set rs = serviceUser.SelectUsers()

    If rs Is Nothing Then
        ErrorManager.ThrowError _
            sheetMasterDataUser.NullUserRecordset, _
            MODULE_NAME & ".RefreshUsers_Click", _
            "DAO tidak mengembalikan recordset user."
    End If

    If rs.State <> adStateOpen Then
        ErrorManager.ThrowError _
            DatabaseManager.RecordsetNotOpen, _
            MODULE_NAME & ".RefreshUsers_Click", _
            "Recordset user tidak terbuka."
    End If

    ' =====================================================
    ' Step 3: Validate required recordset fields
    ' =====================================================
    DatabaseManager.ValidateFields rs, Array( _
        "id", _
        "first_name", _
        "middle_name", _
        "last_name", _
        "username", _
        "manufacturing_plant_name", _
        "position", _
        "tier")

    ' =====================================================
    ' Step 4: Clear previous table rows
    ' =====================================================
    ClearUserTable tbl

    ' =====================================================
    ' Step 5: Fill TableUsers
    ' =====================================================
    rowNumber = 0

    Do While Not rs.EOF

        rowNumber = rowNumber + 1

        fullName = BuildUserFullName( _
            firstName:=DatabaseManager.FieldText( _
                rs, "first_name"), _
            middleName:=DatabaseManager.FieldText( _
                rs, "middle_name"), _
            lastName:=DatabaseManager.FieldText( _
                rs, "last_name"))

        If ValidationText.IsBlank(fullName) Then
            ErrorManager.ThrowError _
                sheetMasterDataUser.EmptyUserFullName, _
                MODULE_NAME & ".RefreshUsers_Click", _
                "Nama user pada record ke-" & _
                CStr(rowNumber) & _
                " tidak boleh kosong."
        End If

        rowIndex = TableManager.AddRow(tbl)

        WriteUserTableRow _
            tbl:=tbl, _
            rowIndex:=rowIndex, _
            rs:=rs, _
            fullName:=fullName

        rs.MoveNext
    Loop

    ' =====================================================
    ' Step 6: Apply formatting
    ' =====================================================
    ApplyUserTableFormatting tbl

    FormManager.ClearStatus

    ' =====================================================
    ' Step 7: Display result
    ' =====================================================
    If rowNumber = 0 Then

        FormManager.ShowInfo _
            "Tidak ada data user yang ditemukan.", _
            "Master Data User"

    Else

        FormManager.ShowInfo _
            CStr(rowNumber) & _
            " data user berhasil ditampilkan.", _
            "Master Data User"

    End If

CleanUp:
    DatabaseManager.CloseRecordset rs

    Set tbl = Nothing

    FormManager.EndBusy state

    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    DatabaseManager.CloseRecordset rs

    Set tbl = Nothing

    FormManager.EndBusy state

    ErrorManager.ShowError errInfo, _
        "Gagal menampilkan master data user."

End Sub


' =========================================================
' Write one user record into TableUsers
' =========================================================
Private Sub WriteUserTableRow( _
    ByVal tbl As ListObject, _
    ByVal rowIndex As Long, _
    ByVal rs As ADODB.recordset, _
    ByVal fullName As String)

    On Error GoTo ErrorHandler

    Dim userId As Long

    Dim errInfo As ErrorInfo

    If tbl Is Nothing Then
        ErrorManager.ThrowError _
            TableManager.TableNotFound, _
            MODULE_NAME & ".WriteUserTableRow", _
            "TableUsers belum diinisialisasi."
    End If

    If Not TableManager.IsValidRow(tbl, rowIndex) Then
        ErrorManager.ThrowError _
            TableManager.InvalidRow, _
            MODULE_NAME & ".WriteUserTableRow", _
            "Target row user belum diinisialisasi."
    End If

    If rs Is Nothing Then
        ErrorManager.ThrowError _
            DatabaseManager.EmptyRecordset, _
            MODULE_NAME & ".WriteUserTableRow", _
            "Recordset user belum diinisialisasi."
    End If

    userId = DatabaseManager.FieldLong(rs, "id")

    If ValidationNumeric.IsZeroOrNegative(userId) Then
        ErrorManager.ThrowError _
            sheetMasterDataUser.InvalidUserId, _
            MODULE_NAME & ".WriteUserTableRow", _
            "ID user tidak valid."
    End If

    ' =====================================================
    ' Id
    ' =====================================================
    TableManager.SetValue _
        tbl, rowIndex, COLUMN_ID, userId

    ' =====================================================
    ' Nama lengkap
    ' =====================================================
    TableManager.SetValue _
        tbl, rowIndex, COLUMN_FULL_NAME, fullName

    ' =====================================================
    ' Username PMI
    ' =====================================================
    TableManager.SetValue _
        tbl, rowIndex, COLUMN_USERNAME, _
        DatabaseManager.FieldText(rs, "username")

    ' =====================================================
    ' Lokasi manufacturing plant
    ' =====================================================
    TableManager.SetValue _
        tbl, rowIndex, COLUMN_PLANT, _
        DatabaseManager.FieldText(rs, "manufacturing_plant_name")

    ' =====================================================
    ' Posisi
    ' =====================================================
    TableManager.SetValue _
        tbl, rowIndex, COLUMN_POSITION, _
        DatabaseManager.FieldText(rs, "position")

    ' =====================================================
    ' Tier
    ' =====================================================
    TableManager.SetValue _
        tbl, rowIndex, COLUMN_TIER, _
        DatabaseManager.FieldText(rs, "tier")

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".WriteUserTableRow", _
        "Gagal menulis satu data user ke tabel."

End Sub


' =========================================================
' Build full name
'
' Empty middle name is ignored, so this will not produce
' duplicate spaces.
' =========================================================
Private Function BuildUserFullName( _
    ByVal firstName As String, _
    ByVal middleName As String, _
    ByVal lastName As String) As String

    Dim fullName As String

    firstName = TrimText(firstName)
    middleName = TrimText(middleName)
    lastName = TrimText(lastName)

    If Len(firstName) > 0 Then
        fullName = firstName
    End If

    If Len(middleName) > 0 Then

        If Len(fullName) > 0 Then
            fullName = fullName & " "
        End If

        fullName = fullName & middleName

    End If

    If Len(lastName) > 0 Then

        If Len(fullName) > 0 Then
            fullName = fullName & " "
        End If

        fullName = fullName & lastName

    End If

    BuildUserFullName = TrimText(fullName)

End Function


' =========================================================
' Clear existing user table rows
' =========================================================
Private Sub ClearUserTable( _
    ByVal tbl As ListObject)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    If tbl Is Nothing Then
        ErrorManager.ThrowError _
            TableManager.TableNotFound, _
            MODULE_NAME & ".ClearUserTable", _
            "TableUsers belum diinisialisasi."
    End If

    TableManager.ClearRows tbl

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".ClearUserTable", _
        "Gagal membersihkan data user sebelumnya."

End Sub


' =========================================================
' Validate TableUsers columns
' =========================================================
Private Sub ValidateUserTableColumns( _
    ByVal tbl As ListObject)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    If tbl Is Nothing Then
        ErrorManager.ThrowError _
            TableManager.TableNotFound, _
            MODULE_NAME & ".ValidateUserTableColumns", _
            "TableUsers tidak ditemukan " & _
            "pada worksheet ini."
    End If

    TableManager.ValidateColumns tbl, Array( _
        COLUMN_ID, _
        COLUMN_FULL_NAME, _
        COLUMN_USERNAME, _
        COLUMN_PLANT, _
        COLUMN_POSITION, _
        COLUMN_TIER)

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & _
        ".ValidateUserTableColumns", _
        "Struktur TableUsers tidak sesuai."

End Sub


' =========================================================
' Apply TableUsers formatting
' =========================================================
Private Sub ApplyUserTableFormatting( _
    ByVal tbl As ListObject)

    On Error GoTo ErrorHandler

    Dim allColumns As Variant
    Dim columnName As Variant
    Dim columnRange As Range

    Dim errInfo As ErrorInfo

    If tbl Is Nothing Then
        ErrorManager.ThrowError _
            TableManager.TableNotFound, _
            MODULE_NAME & ".ApplyUserTableFormatting", _
            "TableUsers belum diinisialisasi."
    End If

    If Not TableManager.HasRows(tbl) Then
        GoTo CleanUp
    End If

    ' =====================================================
    ' Vertical alignment (every table column)
    ' =====================================================
    allColumns = Array( _
        COLUMN_ID, _
        COLUMN_FULL_NAME, _
        COLUMN_USERNAME, _
        COLUMN_PLANT, _
        COLUMN_POSITION, _
        COLUMN_TIER)

    For Each columnName In allColumns

        Set columnRange = TableManager.GetColumnRange( _
            tbl, CStr(columnName))

        If Not columnRange Is Nothing Then
            columnRange.VerticalAlignment = xlCenter
        End If

    Next columnName

    ' =====================================================
    ' ID
    ' =====================================================
    TableManager.FormatColumn _
        tbl, COLUMN_ID, "0", xlCenter

    ' =====================================================
    ' Text fields
    ' =====================================================
    TableManager.FormatColumn _
        tbl, COLUMN_FULL_NAME, "@"

    TableManager.FormatColumn _
        tbl, COLUMN_USERNAME, "@"

    TableManager.FormatColumn _
        tbl, COLUMN_PLANT, "@"

    TableManager.FormatColumn _
        tbl, COLUMN_POSITION, "@"

    TableManager.FormatColumn _
        tbl, COLUMN_TIER, "@", xlCenter

CleanUp:
    Set columnRange = Nothing
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    Set columnRange = Nothing

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & _
        ".ApplyUserTableFormatting", _
        "Gagal menerapkan format TableUsers."

End Sub
