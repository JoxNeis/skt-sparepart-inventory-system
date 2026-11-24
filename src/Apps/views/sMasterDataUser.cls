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

    state = managerForm.BeginBusy()

    ' =====================================================
    ' Step 1: Read user input
    ' =====================================================
    manufacturingPlantName = managerForm.GetCellText( _
        Me, CELL_USER_PLANT)

    username = managerForm.GetCellText( _
        Me, CELL_USER_USERNAME)

    firstName = managerForm.GetCellText( _
        Me, CELL_USER_FIRST_NAME)

    middleName = managerForm.GetCellText( _
        Me, CELL_USER_MIDDLE_NAME)

    lastName = managerForm.GetCellText( _
        Me, CELL_USER_LAST_NAME)

    userPosition = managerForm.GetCellText( _
        Me, CELL_USER_POSITION)

    tierValue = managerForm.GetCellValue( _
        Me, CELL_USER_TIER)

    ' =====================================================
    ' Step 2: Validate required text inputs
    ' =====================================================
    If managerValidationText.IsBlank(manufacturingPlantName) Then
        managerError.ThrowError _
            sheetMasterDataUser.EmptyPlantName, _
            MODULE_NAME & ".SaveUser_Click", _
            "Lokasi manufacturing plant wajib diisi."
    End If

    If managerValidationText.IsBlank(username) Then
        managerError.ThrowError _
            sheetMasterDataUser.EmptyUsername, _
            MODULE_NAME & ".SaveUser_Click", _
            "Username PMI wajib diisi."
    End If

    If managerValidationText.IsBlank(firstName) Then
        managerError.ThrowError _
            sheetMasterDataUser.EmptyFirstName, _
            MODULE_NAME & ".SaveUser_Click", _
            "Nama depan wajib diisi."
    End If

    ' Middle name is optional.

    If managerValidationText.IsBlank(lastName) Then
        managerError.ThrowError _
            sheetMasterDataUser.EmptyLastName, _
            MODULE_NAME & ".SaveUser_Click", _
            "Nama belakang wajib diisi."
    End If

    If managerValidationText.IsBlank(userPosition) Then
        managerError.ThrowError _
            sheetMasterDataUser.EmptyPosition, _
            MODULE_NAME & ".SaveUser_Click", _
            "Posisi user wajib diisi."
    End If

    ' =====================================================
    ' Step 3: Validate tier
    ' =====================================================
    If IsError(tierValue) Then
        managerError.ThrowError _
            sheetMasterDataUser.InvalidTierErrorValue, _
            MODULE_NAME & ".SaveUser_Click", _
            "Tier berisi nilai error."
    End If

    ' Null, Empty and text with only spaces.
    If managerValidationText.IsBlank(tierValue) Then
        managerError.ThrowError _
            sheetMasterDataUser.EmptyTier, _
            MODULE_NAME & ".SaveUser_Click", _
            "Tier user wajib diisi."
    End If

    If managerValidationNumeric.IsNotNumber(tierValue) Then
        managerError.ThrowError _
            sheetMasterDataUser.InvalidTierNumber, _
            MODULE_NAME & ".SaveUser_Click", _
            "Tier user harus berupa angka."
    End If

    If Not managerValidationNumeric.IsWholeNumber(CDbl(tierValue)) Then
        managerError.ThrowError _
            sheetMasterDataUser.InvalidTierWholeNumber, _
            MODULE_NAME & ".SaveUser_Click", _
            "Tier user harus berupa bilangan bulat."
    End If

    tier = CLng(tierValue)

    If managerValidationNumeric.IsZeroOrNegative(tier) Then
        managerError.ThrowError _
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

    If Not managerForm.Confirm( _
        confirmationMessage, _
        "Konfirmasi Simpan User") Then

        GoTo CleanUp
    End If

    ' =====================================================
    ' Step 6: Call service procedure
    ' =====================================================
    managerForm.SetStatus _
        "Menyimpan user baru...", True

    serviceUser.CreateUser _
        manufacturingPlantName:=manufacturingPlantName, _
        username:=username, _
        firstName:=firstName, _
        middleName:=middleName, _
        lastName:=lastName, _
        position:=userPosition, _
        tier:=tier

    managerForm.ClearStatus

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

    managerForm.ShowInfo _
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
    managerForm.EndBusy state
    managerForm.ClearStatus

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerForm.EndBusy state
    managerForm.ClearStatus

    managerError.ShowError errInfo, _
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

    managerForm.ClearCells Me, _
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
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
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

    state = managerForm.BeginBusy( _
        "Mengambil data user...")

    ' =====================================================
    ' Step 1: Get and validate TableUsers
    ' =====================================================
    Set tbl = managerTable.GetTable(Me, TABLE_NAME)

    ValidateUserTableColumns tbl

    ' =====================================================
    ' Step 2: Call service function
    ' =====================================================
    Set rs = serviceUser.SelectUsers()

    If rs Is Nothing Then
        managerError.ThrowError _
            sheetMasterDataUser.NullUserRecordset, _
            MODULE_NAME & ".RefreshUsers_Click", _
            "DAO tidak mengembalikan recordset user."
    End If

    If rs.State <> adStateOpen Then
        managerError.ThrowError _
            managerDatabase.RecordsetNotOpen, _
            MODULE_NAME & ".RefreshUsers_Click", _
            "Recordset user tidak terbuka."
    End If

    ' =====================================================
    ' Step 3: Validate required recordset fields
    ' =====================================================
    managerDatabase.ValidateFields rs, Array( _
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
            firstName:=managerDatabase.FieldText( _
                rs, "first_name"), _
            middleName:=managerDatabase.FieldText( _
                rs, "middle_name"), _
            lastName:=managerDatabase.FieldText( _
                rs, "last_name"))

        If managerValidationText.IsBlank(fullName) Then
            managerError.ThrowError _
                sheetMasterDataUser.EmptyUserFullName, _
                MODULE_NAME & ".RefreshUsers_Click", _
                "Nama user pada record ke-" & _
                CStr(rowNumber) & _
                " tidak boleh kosong."
        End If

        rowIndex = managerTable.AddRow(tbl)

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

    managerForm.ClearStatus

    ' =====================================================
    ' Step 7: Display result
    ' =====================================================
    If rowNumber = 0 Then

        managerForm.ShowInfo _
            "Tidak ada data user yang ditemukan.", _
            "Master Data User"

    Else

        managerForm.ShowInfo _
            CStr(rowNumber) & _
            " data user berhasil ditampilkan.", _
            "Master Data User"

    End If

CleanUp:
    managerDatabase.CloseRecordset rs

    Set tbl = Nothing

    managerForm.EndBusy state

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerDatabase.CloseRecordset rs

    Set tbl = Nothing

    managerForm.EndBusy state

    managerError.ShowError errInfo, _
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
        managerError.ThrowError _
            managerTable.TableNotFound, _
            MODULE_NAME & ".WriteUserTableRow", _
            "TableUsers belum diinisialisasi."
    End If

    If Not managerTable.IsValidRow(tbl, rowIndex) Then
        managerError.ThrowError _
            managerTable.InvalidRow, _
            MODULE_NAME & ".WriteUserTableRow", _
            "Target row user belum diinisialisasi."
    End If

    If rs Is Nothing Then
        managerError.ThrowError _
            managerDatabase.EmptyRecordset, _
            MODULE_NAME & ".WriteUserTableRow", _
            "Recordset user belum diinisialisasi."
    End If

    userId = managerDatabase.FieldLong(rs, "id")

    If managerValidationNumeric.IsZeroOrNegative(userId) Then
        managerError.ThrowError _
            sheetMasterDataUser.InvalidUserId, _
            MODULE_NAME & ".WriteUserTableRow", _
            "ID user tidak valid."
    End If

    ' =====================================================
    ' Id
    ' =====================================================
    managerTable.SetValue _
        tbl, rowIndex, COLUMN_ID, userId

    ' =====================================================
    ' Nama lengkap
    ' =====================================================
    managerTable.SetValue _
        tbl, rowIndex, COLUMN_FULL_NAME, fullName

    ' =====================================================
    ' Username PMI
    ' =====================================================
    managerTable.SetValue _
        tbl, rowIndex, COLUMN_USERNAME, _
        managerDatabase.FieldText(rs, "username")

    ' =====================================================
    ' Lokasi manufacturing plant
    ' =====================================================
    managerTable.SetValue _
        tbl, rowIndex, COLUMN_PLANT, _
        managerDatabase.FieldText(rs, "manufacturing_plant_name")

    ' =====================================================
    ' Posisi
    ' =====================================================
    managerTable.SetValue _
        tbl, rowIndex, COLUMN_POSITION, _
        managerDatabase.FieldText(rs, "position")

    ' =====================================================
    ' Tier
    ' =====================================================
    managerTable.SetValue _
        tbl, rowIndex, COLUMN_TIER, _
        managerDatabase.FieldText(rs, "tier")

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
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
        managerError.ThrowError _
            managerTable.TableNotFound, _
            MODULE_NAME & ".ClearUserTable", _
            "TableUsers belum diinisialisasi."
    End If

    managerTable.ClearRows tbl

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
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
        managerError.ThrowError _
            managerTable.TableNotFound, _
            MODULE_NAME & ".ValidateUserTableColumns", _
            "TableUsers tidak ditemukan " & _
            "pada worksheet ini."
    End If

    managerTable.ValidateColumns tbl, Array( _
        COLUMN_ID, _
        COLUMN_FULL_NAME, _
        COLUMN_USERNAME, _
        COLUMN_PLANT, _
        COLUMN_POSITION, _
        COLUMN_TIER)

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
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
        managerError.ThrowError _
            managerTable.TableNotFound, _
            MODULE_NAME & ".ApplyUserTableFormatting", _
            "TableUsers belum diinisialisasi."
    End If

    If Not managerTable.HasRows(tbl) Then
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

        Set columnRange = managerTable.GetColumnRange( _
            tbl, CStr(columnName))

        If Not columnRange Is Nothing Then
            columnRange.VerticalAlignment = xlCenter
        End If

    Next columnName

    ' =====================================================
    ' ID
    ' =====================================================
    managerTable.FormatColumn _
        tbl, COLUMN_ID, "0", xlCenter

    ' =====================================================
    ' Text fields
    ' =====================================================
    managerTable.FormatColumn _
        tbl, COLUMN_FULL_NAME, "@"

    managerTable.FormatColumn _
        tbl, COLUMN_USERNAME, "@"

    managerTable.FormatColumn _
        tbl, COLUMN_PLANT, "@"

    managerTable.FormatColumn _
        tbl, COLUMN_POSITION, "@"

    managerTable.FormatColumn _
        tbl, COLUMN_TIER, "@", xlCenter

CleanUp:
    Set columnRange = Nothing
    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    Set columnRange = Nothing

    managerError.Rethrow errInfo, _
        MODULE_NAME & _
        ".ApplyUserTableFormatting", _
        "Gagal menerapkan format TableUsers."

End Sub
