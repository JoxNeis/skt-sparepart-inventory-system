Attribute VB_Name = "sInputRequest"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

' Error codes: sheetInputRequest (managerErrorCodesSheets.bas)

Private Const TABLE_NAME As String = "TableInputRequest"

Private Const CELL_LOCATION As String = "B4"
Private Const CELL_REQUEST_DATE As String = "B6"

Private Const COLUMN_CATALOGUE As String = "Nama Barang"
Private Const COLUMN_AMOUNT As String = "Jumlah Barang"
Private Const COLUMN_MOBILE_POST As String = "Mobile Post"
Private Const COLUMN_REMARKS As String = "Remarks"
Private Const COLUMN_PIC As String = "PIC"

Private Const MODULE_NAME As String = "Worksheet_INPUT_REQUEST"

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    managerForm.HandleDatePickerSelection Target, Me.Range("B6")
End Sub

Private Sub Worksheet_Activate()

    Me.Range(CELL_LOCATION).Select

End Sub

Public Sub SaveRequest_Click()

    On Error GoTo ErrorHandler

    Dim plantName As String
    Dim requestDateValue As Date

    Dim tbl As ListObject
    Dim rowIndex As Long

    Dim mobilePostItem As MobilePost
    Dim catalogueItem As Catalogue
    Dim requestItem As Request

    Dim catalogueName As String
    Dim amountCellValue As Variant
    Dim amountValue As Double
    Dim mobilePostName As String
    Dim remarksValue As String
    Dim picValue As String

    Dim transactionUUID As String

    Dim savedCount As Long
    Dim currentRowIndex As Long
    Dim errInfo As ErrorInfo

    ' =====================================================
    ' Step 1: Read header values
    ' =====================================================
    plantName = managerForm.GetCellText(Me, CELL_LOCATION)

    If managerValidationText.IsBlank(plantName) Then
        managerError.ThrowError _
            sheetInputRequest.EmptyPlantName, _
            MODULE_NAME & ".SaveRequest_Click", _
            "Nama manufacturing plant tidak boleh kosong."
    End If

    If Not managerForm.IsCellDate(Me, CELL_REQUEST_DATE) Then

        managerError.ThrowError _
            sheetInputRequest.InvalidRequestDate, _
            MODULE_NAME & ".SaveRequest_Click", _
            "Tanggal request tidak valid."
    End If

    requestDateValue = DateValue( _
        CDate(managerForm.GetCellValue(Me, CELL_REQUEST_DATE)))

    Set tbl = managerTable.GetTable(Me, TABLE_NAME)

    If Not managerTable.HasRows(tbl) Then

        managerForm.ShowWarning "Tidak ada baris data untuk disimpan."

        GoTo CleanUp

    End If

    ' =====================================================
    ' Step 2: Loop every row and call CreateRequest
    ' =====================================================
    savedCount = 0

    For rowIndex = 1 To managerTable.RowCount(tbl)

        currentRowIndex = rowIndex

        ' -------------------------------------------------
        ' Skip a completely empty table row
        ' -------------------------------------------------
        If Not managerTable.IsRowEmpty(tbl, rowIndex) Then

            ' -------------------------------------------------
            ' Read the current row values
            ' -------------------------------------------------
            catalogueName = managerTable.GetText( _
                tbl, rowIndex, COLUMN_CATALOGUE)

            mobilePostName = managerTable.GetText( _
                tbl, rowIndex, COLUMN_MOBILE_POST)

            remarksValue = managerTable.GetText( _
                tbl, rowIndex, COLUMN_REMARKS)

            picValue = managerTable.GetText( _
                tbl, rowIndex, COLUMN_PIC)

            ' -------------------------------------------------
            ' Validate the amount before converting to Double
            ' (blank passes here as 0 and is rejected by the
            ' "lebih besar dari nol" check below, as before)
            ' -------------------------------------------------
            amountCellValue = managerTable.GetValue( _
                tbl, rowIndex, COLUMN_AMOUNT)

            If Not IsNumeric(amountCellValue) Then

                managerError.ThrowError _
                    sheetInputRequest.InvalidAmount, _
                    MODULE_NAME & ".SaveRequest_Click", _
                    "Jumlah barang harus berupa angka."

            End If

            amountValue = managerTable.GetDouble( _
                tbl, rowIndex, COLUMN_AMOUNT)

            ' -------------------------------------------------
            ' Validate the current row
            ' -------------------------------------------------
            If managerValidationText.IsBlank(catalogueName) Then
                managerError.ThrowError _
                    sheetInputRequest.EmptyCatalogueName, _
                    MODULE_NAME & ".SaveRequest_Click", _
                    "Nama barang tidak boleh kosong."
            End If

            If managerValidationNumeric.IsZeroOrNegative(amountValue) Then
                managerError.ThrowError _
                    sheetInputRequest.ZeroOrNegativeAmount, _
                    MODULE_NAME & ".SaveRequest_Click", _
                    "Jumlah barang harus lebih besar dari nol."
            End If

            If managerValidationText.IsBlank(mobilePostName) Then
                managerError.ThrowError _
                    sheetInputRequest.EmptyMobilePostName, _
                    MODULE_NAME & ".SaveRequest_Click", _
                    "Nama mobile post tidak boleh kosong."
            End If

            If managerValidationText.IsBlank(picValue) Then
                managerError.ThrowError _
                    sheetInputRequest.EmptyPersonInCharge, _
                    MODULE_NAME & ".SaveRequest_Click", _
                    "PIC tidak boleh kosong."
            End If

            ' -------------------------------------------------
            ' Create Catalogue object
            ' -------------------------------------------------
            Set catalogueItem = New Catalogue

            catalogueItem.Name = catalogueName

            ' -------------------------------------------------
            ' Create MobilePost object
            ' -------------------------------------------------
            Set mobilePostItem = New MobilePost

            mobilePostItem.Name = mobilePostName
            Set mobilePostItem.Plant = New Plant
            mobilePostItem.Plant.Name = plantName

            ' -------------------------------------------------
            ' Create Request object
            ' -------------------------------------------------
            Set requestItem = New Request

            requestItem.orderDate = requestDateValue

            Set requestItem.Catalogue = catalogueItem
            Set requestItem.MobilePost = mobilePostItem

            requestItem.Amount = amountValue
            requestItem.remarks = remarksValue
            requestItem.SystemRemarks = vbNullString
            requestItem.personInCharge = picValue

            ' -------------------------------------------------
            ' Call the service function
            ' -------------------------------------------------
            transactionUUID = serviceRequest.CreateRequest( _
                requestItem:=requestItem)

            savedCount = savedCount + 1

            Debug.Print _
                "REQUEST SAVED", _
                transactionUUID, _
                catalogueName, _
                amountValue, _
                mobilePostName

            Set requestItem = Nothing
            Set mobilePostItem = Nothing
            Set catalogueItem = Nothing

        End If

    Next rowIndex

    ' =====================================================
    ' Step 3: Display the result
    ' =====================================================
    If savedCount = 0 Then

        managerForm.ShowWarning _
            "Tidak ada data request yang dapat disimpan."

        GoTo CleanUp

    End If

    managerForm.ShowInfo _
        savedCount & _
        " request berhasil disimpan."

    ' =====================================================
    ' Step 4: Clear the table after saving
    ' =====================================================
    managerTable.ClearContents tbl

    ClearEmptyRequestRows

    managerTable.AddRow tbl

    Me.Range(CELL_LOCATION).Select

CleanUp:
    Set tbl = Nothing

    Set requestItem = Nothing
    Set mobilePostItem = Nothing
    Set catalogueItem = Nothing

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.ShowError errInfo, _
        "Baris ke-" & currentRowIndex & " pada tabel."

    Resume CleanUp

End Sub

Public Sub AddRequestRow_Click()

    On Error GoTo ErrorHandler

    Dim tbl As ListObject
    Dim rowIndex As Long
    Dim errInfo As ErrorInfo

    Set tbl = managerTable.GetTable(Me, TABLE_NAME)

    rowIndex = managerTable.AddRow(tbl)

    managerTable.SelectCell tbl, rowIndex, COLUMN_CATALOGUE

CleanUp:
    Set tbl = Nothing

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.ShowError errInfo

    Resume CleanUp

End Sub

Public Sub ClearRequestTable_Click()

    On Error GoTo ErrorHandler

    Dim tbl As ListObject
    Dim errInfo As ErrorInfo

    Set tbl = managerTable.GetTable(Me, TABLE_NAME)

    If Not managerTable.HasRows(tbl) Then

        managerForm.ShowInfo "Tabel input request sudah kosong."

        GoTo CleanUp

    End If

    If Not managerForm.Confirm( _
        "Apakah Anda yakin ingin membersihkan " & _
        "seluruh data pada tabel?") Then

        GoTo CleanUp
    End If

    managerTable.ClearContents tbl

    ClearEmptyRequestRows

    managerTable.AddRow tbl

    Me.Range(CELL_LOCATION).Select

CleanUp:
    Set tbl = Nothing
    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.ShowError errInfo

    Resume CleanUp

End Sub

Public Sub ClearEmptyRequestRows()

    On Error GoTo ErrorHandler

    Dim tbl As ListObject
    Dim errInfo As ErrorInfo

    Set tbl = managerTable.GetTable(Me, TABLE_NAME)

    managerTable.DeleteEmptyRows tbl

CleanUp:
    Set tbl = Nothing

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.ShowError errInfo

    Resume CleanUp

End Sub
