Attribute VB_Name = "sInputPenerimaan"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

' Error codes: sheetInputPenerimaan (managerErrorCodesSheets.bas)

Private Const MODULE_NAME As String = "SHEET_INPUT_RECEIVE"

Private Const TABLE_NAME As String = "TableInputPenerimaan"

Private Const CELL_LOCATION As String = "B4"
Private Const CELL_FROM_DATE As String = "B7"
Private Const CELL_UNTIL_DATE As String = "B8"

Private Const COLUMN_NO As String = "No"
Private Const COLUMN_UUID As String = "ORDER UUID"
Private Const COLUMN_PO_ID As String = "PO ID"
Private Const COLUMN_CATALOGUE As String = "Nama Barang"
Private Const COLUMN_AMOUNT As String = "Jumlah Order"
Private Const COLUMN_UOM As String = "Satuan"
Private Const COLUMN_ORDER_DATE As String = "Tanggal Order"
Private Const COLUMN_REMARKS As String = "Remarks"
Private Const COLUMN_PIC As String = "PIC"
Private Const COLUMN_ARRIVAL_DATE As String = "Tanggal Datang"
Private Const COLUMN_VERIFICATION_DATE As String = "Tanggal Verifikasi"
Private Const COLUMN_TOTAL_RECEIVE As String = "Jumlah Diterima"
Private Const COLUMN_OK_AMOUNT As String = "Jumlah OK"
Private Const COLUMN_NOT_OK_AMOUNT As String = "Jumlah NOT OK"

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    managerForm.HandleDatePickerSelection Target, Me.Range("B7:B8")
End Sub

Public Sub SearchReceive_Click()

    On Error GoTo ErrorHandler

    Dim plantName As String
    Dim fromDateValue As Date
    Dim untilDateValue As Date

    Dim receives As Collection
    Dim receiveItem As Receive

    Dim tbl As ListObject
    Dim rowIndex As Long
    Dim rowNumber As Long
    Dim errInfo As ErrorInfo

    ' =====================================================
    ' Step 1: Read filter values
    ' =====================================================
    plantName = managerForm.GetCellText(Me, CELL_LOCATION)
    fromDateValue = CDate(managerForm.GetCellValue(Me, CELL_FROM_DATE))
    untilDateValue = CDate(managerForm.GetCellValue(Me, CELL_UNTIL_DATE))

    ' =====================================================
    ' Step 2: Call the select function
    ' =====================================================
    Set receives = serviceReceive.SelectPendingOrderReceive( _
        manufacturingPlantName:=plantName, _
        fromDate:=fromDateValue, _
        untilDate:=untilDateValue)

    ' =====================================================
    ' Step 3: Clear the existing table rows
    ' =====================================================
    Set tbl = managerTable.GetTable(Me, TABLE_NAME)

    managerTable.ClearRows tbl

    ' =====================================================
    ' Step 4: Fill the table with the result
    ' =====================================================
    rowNumber = 0

    For Each receiveItem In receives

        rowNumber = rowNumber + 1

        rowIndex = managerTable.AddRow(tbl)

        managerTable.SetValue tbl, rowIndex, COLUMN_NO, rowNumber
        managerTable.SetValue tbl, rowIndex, COLUMN_UUID, receiveItem.Order.UUID
        managerTable.SetValue tbl, rowIndex, COLUMN_PO_ID, receiveItem.Order.poID
        managerTable.SetValue tbl, rowIndex, COLUMN_CATALOGUE, receiveItem.Order.Catalogue.Name
        managerTable.SetValue tbl, rowIndex, COLUMN_AMOUNT, receiveItem.RemainingOrder
        managerTable.SetValue tbl, rowIndex, COLUMN_UOM, receiveItem.Order.Catalogue.unitOfMeasurement
        managerTable.SetValue tbl, rowIndex, COLUMN_ORDER_DATE, receiveItem.Order.orderDate

    Next receiveItem

CleanUp:
    Set tbl = Nothing
    Set receiveItem = Nothing
    Set receives = Nothing
    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.ShowError errInfo

    Resume CleanUp

End Sub

Public Sub SaveReceive_Click()

    On Error GoTo ErrorHandler

    Dim tbl As ListObject
    Dim rowIndex As Long

    Dim receiveItem As Receive
    Dim orderItem As Order
    Dim catalogueItem As Catalogue
    Dim plantItem As Plant

    Dim arrivalDateValue As Variant
    Dim verificationDateValue As Variant

    Dim plantName As String
    Dim catalogueName As String
    Dim orderUuidValue As String
    Dim orderPoIdValue As String

    Dim savedCount As Long
    Dim skippedCount As Long
    Dim currentRowIndex As Long
    Dim errInfo As ErrorInfo

    ' =====================================================
    ' Step 1: Read manufacturing plant name
    ' =====================================================
    plantName = managerForm.GetCellText(Me, CELL_LOCATION)

    If managerValidationText.IsBlank(plantName) Then
        managerError.ThrowError _
            sheetInputPenerimaan.EmptyPlantName, _
            MODULE_NAME & ".SaveReceive_Click", _
            "Nama manufacturing plant tidak boleh kosong."
    End If

    ' =====================================================
    ' Step 2: Get input table
    ' =====================================================
    Set tbl = managerTable.GetTable(Me, TABLE_NAME)

    If Not managerTable.HasRows(tbl) Then
        managerForm.ShowWarning "Tidak ada baris data untuk disimpan."

        GoTo CleanUp
    End If

    savedCount = 0
    skippedCount = 0
    currentRowIndex = 0

    ' =====================================================
    ' Step 3: Process each table row
    ' =====================================================
    For rowIndex = 1 To managerTable.RowCount(tbl)

        currentRowIndex = rowIndex

        orderUuidValue = managerTable.GetText( _
            tbl, rowIndex, COLUMN_UUID)

        orderPoIdValue = managerTable.GetText( _
            tbl, rowIndex, COLUMN_PO_ID)

        catalogueName = managerTable.GetText( _
            tbl, rowIndex, COLUMN_CATALOGUE)

        arrivalDateValue = managerTable.GetValue( _
            tbl, rowIndex, COLUMN_ARRIVAL_DATE)

        verificationDateValue = managerTable.GetValue( _
            tbl, rowIndex, COLUMN_VERIFICATION_DATE)

        ' =================================================
        ' Skip rows without a PO ID
        ' =================================================
        If managerValidationText.IsBlank(orderPoIdValue) Then
            skippedCount = skippedCount + 1
            GoTo ContinueLoop
        End If

        ' =================================================
        ' Validate order UUID
        ' =================================================
        If managerValidationText.IsBlank(orderUuidValue) Then
            managerError.ThrowError _
                sheetInputPenerimaan.EmptyOrderUUID, _
                MODULE_NAME & ".SaveReceive_Click", _
                "ORDER UUID tidak boleh kosong."
        End If

        ' =================================================
        ' Validate catalogue name
        ' =================================================
        If managerValidationText.IsBlank(catalogueName) Then
            managerError.ThrowError _
                sheetInputPenerimaan.EmptyCatalogueName, _
                MODULE_NAME & ".SaveReceive_Click", _
                "Nama barang atau catalogue tidak boleh kosong."
        End If

        ' =================================================
        ' Skip rows without an arrival date
        ' =================================================
        If IsEmpty(arrivalDateValue) Then
            skippedCount = skippedCount + 1
            GoTo ContinueLoop
        End If

        If Not IsDate(arrivalDateValue) Then
            managerError.ThrowError _
                sheetInputPenerimaan.InvalidArrivalDate, _
                MODULE_NAME & ".SaveReceive_Click", _
                "Tanggal datang tidak valid."
        End If

        If Not IsEmpty(verificationDateValue) Then
            If Not IsDate(verificationDateValue) Then
                managerError.ThrowError _
                    sheetInputPenerimaan.InvalidVerificationDate, _
                    MODULE_NAME & ".SaveReceive_Click", _
                    "Tanggal verifikasi tidak valid."
            End If
        End If

        ' =================================================
        ' Step 4: Update order PO ID
        ' =================================================
        serviceOrder.UpdateOrderPOID _
            orderUUID:=orderUuidValue, _
            poID:=orderPoIdValue

        ' =================================================
        ' Step 5: Create related objects
        ' =================================================
        Set plantItem = New Plant
        plantItem.Name = plantName

        Set catalogueItem = New Catalogue
        catalogueItem.Name = catalogueName

        Set orderItem = New Order
        orderItem.UUID = orderUuidValue
        orderItem.Amount = managerTable.GetDouble( _
            tbl, rowIndex, COLUMN_TOTAL_RECEIVE)

        Set orderItem.Plant = plantItem
        Set orderItem.Catalogue = catalogueItem

        ' =================================================
        ' Step 6: Create Receive object
        ' =================================================
        Set receiveItem = New Receive

        Set receiveItem.Order = orderItem

        receiveItem.ArrivalDate = managerTable.GetDate( _
            tbl, rowIndex, COLUMN_ARRIVAL_DATE)

        If Not IsEmpty(verificationDateValue) Then
            receiveItem.VerificationDate = managerTable.GetDate( _
                tbl, rowIndex, COLUMN_VERIFICATION_DATE)
        End If

        receiveItem.totalReceiveAmount = managerTable.GetDouble( _
            tbl, rowIndex, COLUMN_TOTAL_RECEIVE)

        receiveItem.okAmount = managerTable.GetDouble( _
            tbl, rowIndex, COLUMN_OK_AMOUNT)

        receiveItem.notOKAmount = managerTable.GetDouble( _
            tbl, rowIndex, COLUMN_NOT_OK_AMOUNT)

        receiveItem.remarks = managerTable.GetText( _
            tbl, rowIndex, COLUMN_REMARKS)

        receiveItem.personInCharge = managerTable.GetText( _
            tbl, rowIndex, COLUMN_PIC)

        ' =================================================
        ' Step 7: Validate receive amounts
        ' =================================================
        If managerValidationNumeric.IsZeroOrNegative( _
            receiveItem.totalReceiveAmount) Then

            managerError.ThrowError _
                sheetInputPenerimaan.InvalidTotalReceiveAmount, _
                MODULE_NAME & ".SaveReceive_Click", _
                "Jumlah diterima harus lebih besar dari nol."
        End If

        If managerValidationNumeric.IsNegative(receiveItem.okAmount) Then
            managerError.ThrowError _
                sheetInputPenerimaan.NegativeOkAmount, _
                MODULE_NAME & ".SaveReceive_Click", _
                "Jumlah OK tidak boleh kurang dari nol."
        End If

        If managerValidationNumeric.IsNegative(receiveItem.notOKAmount) Then
            managerError.ThrowError _
                sheetInputPenerimaan.NegativeNotOkAmount, _
                MODULE_NAME & ".SaveReceive_Click", _
                "Jumlah NOT OK tidak boleh kurang dari nol."
        End If

        If receiveItem.okAmount + _
            receiveItem.notOKAmount <> _
            receiveItem.totalReceiveAmount Then

            managerError.ThrowError _
                sheetInputPenerimaan.InvalidOkNotOkSum, _
                MODULE_NAME & ".SaveReceive_Click", _
                "Jumlah OK ditambah jumlah NOT OK harus " & _
                "sama dengan jumlah diterima." & vbCrLf & _
                "Jumlah Diterima: " & _
                Format$( _
                    receiveItem.totalReceiveAmount, _
                    "#,##0.00") & _
                vbCrLf & _
                "Jumlah OK: " & _
                Format$( _
                    receiveItem.okAmount, _
                    "#,##0.00") & _
                vbCrLf & _
                "Jumlah NOT OK: " & _
                Format$( _
                    receiveItem.notOKAmount, _
                    "#,##0.00")
        End If

        If managerValidationText.IsBlank(receiveItem.personInCharge) Then
            managerError.ThrowError _
                sheetInputPenerimaan.EmptyPersonInCharge, _
                MODULE_NAME & ".SaveReceive_Click", _
                "Person in charge tidak boleh kosong."
        End If

        ' =================================================
        ' Step 8: Save receive
        ' =================================================
        serviceReceive.CreateReceive _
            receiveItem:=receiveItem, _
            orderTransactionUUID:=orderUuidValue

        savedCount = savedCount + 1

ContinueLoop:
        Set receiveItem = Nothing
        Set orderItem = Nothing
        Set catalogueItem = Nothing
        Set plantItem = Nothing

        arrivalDateValue = Empty
        verificationDateValue = Empty

    Next rowIndex

    ' =====================================================
    ' Step 9: Show result
    ' =====================================================
    managerForm.ShowInfo _
        savedCount & _
        " penerimaan berhasil disimpan." & _
        vbCrLf & _
        skippedCount & _
        " baris dilewati karena belum lengkap." & _
        vbCrLf & _
        "Catatan: PO ID dan tanggal datang wajib diisi."

    ' =====================================================
    ' Step 10: Clear table after successful process
    ' =====================================================
    managerTable.ClearContents tbl

CleanUp:
    Set receiveItem = Nothing
    Set orderItem = Nothing
    Set catalogueItem = Nothing
    Set plantItem = Nothing

    Set tbl = Nothing

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.ShowError errInfo, _
        "Baris ke-" & currentRowIndex & " pada tabel."

    Resume CleanUp

End Sub
