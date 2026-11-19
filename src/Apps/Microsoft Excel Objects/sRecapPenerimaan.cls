Attribute VB_Name = "sRecapPenerimaan"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Const MODULE_NAME As String = _
    "SHEET_RECAP_RECEIVE"

Private Const TABLE_NAME As String = _
    "TableRecapReceive"

Private Const CELL_LOCATION As String = "B4"
Private Const CELL_FROM_DATE As String = "B7"
Private Const CELL_UNTIL_DATE As String = "B8"

Private Const COLUMN_NO As String = "No"
Private Const COLUMN_ORDER_UUID As String = "ORDER UUID"
Private Const COLUMN_RECEIVE_UUID As String = "RECEIVE UUID"
Private Const COLUMN_PO_ID As String = "PO ID"
Private Const COLUMN_CATALOGUE As String = "Nama Barang"
Private Const COLUMN_ARRIVAL_DATE As String = "Tanggal Datang"

Private Const COLUMN_VERIFICATION_DATE As String = _
    "Tanggal Verifikasi"

Private Const COLUMN_UOM As String = "Satuan"
Private Const COLUMN_ORDER_AMOUNT As String = "Jumlah Order"

Private Const COLUMN_RECEIVE_AMOUNT As String = _
    "Jumlah Diterima"

Private Const COLUMN_OK_AMOUNT As String = "Jumlah OK"

Private Const COLUMN_NOT_OK_AMOUNT As String = _
    "Jumlah NOT OK"

Private Const COLUMN_REMARKS As String = "Remarks"
Private Const COLUMN_PIC As String = "PIC"
Private Const COLUMN_UPDATE As String = "UPDATE"


' =========================================================
' WORKSHEET EVENTS
' =========================================================

Private Sub Worksheet_SelectionChange( _
    ByVal Target As Range)

    On Error GoTo ErrorHandler

    Dim tbl As ListObject
    Dim selectedRowIndex As Long

    Dim state As FormState
    Dim errInfo As ErrorInfo

    state = FormManager.CaptureState()

    ' =====================================================
    ' Date input cells
    ' =====================================================
    If FormManager.IsInRange( _
        Target, _
        Me.Range( _
            CELL_FROM_DATE & ":" & CELL_UNTIL_DATE _
        )) Then

        Application.EnableEvents = False

        FormManager.ShowDatePicker Target

        GoTo CleanUp
    End If

    ' =====================================================
    ' Get the receive recap table
    ' =====================================================
    Set tbl = TableManager.GetTable(Me, TABLE_NAME)

    If Not TableManager.HasRows(tbl) Then
        GoTo CleanUp
    End If

    ' =====================================================
    ' UPDATE column clicked
    ' =====================================================
    If TableManager.IsInColumn( _
        Target, _
        tbl, _
        COLUMN_UPDATE) Then

        If Not ValidationText.IsSameText( _
            TrimText(CStr(Target.value)), _
            "UPDATE") Then

            GoTo CleanUp
        End If

        Application.EnableEvents = False

        selectedRowIndex = _
            TableManager.GetRowIndex(tbl, Target)

        CreateReceiveUpdateRequest _
            tbl:=tbl, _
            selectedRowIndex:=selectedRowIndex

        GoTo CleanUp
    End If

CleanUp:
    Set tbl = Nothing

    FormManager.EndBusy state
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    Set tbl = Nothing

    FormManager.EndBusy state
    FormManager.ClearStatus

    ErrorManager.ShowError errInfo

End Sub


' =========================================================
' SEARCH RECEIVE
' =========================================================

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

    Dim state As FormState
    Dim errInfo As ErrorInfo

    ' =====================================================
    ' Step 1: Read and validate filter values
    ' =====================================================
    plantName = FormManager.GetCellText(Me, CELL_LOCATION)

    If ValidationText.IsBlank(plantName) Then
        ErrorManager.ThrowError _
            sheetRecapPenerimaan.EmptyPlantName, _
            MODULE_NAME & ".SearchReceive_Click", _
            "Lokasi belum dipilih."
    End If

    If Not FormManager.IsCellDate(Me, CELL_FROM_DATE) Then
        ErrorManager.ThrowError _
            sheetRecapPenerimaan.InvalidFromDate, _
            MODULE_NAME & ".SearchReceive_Click", _
            "Tanggal awal tidak valid."
    End If

    If Not FormManager.IsCellDate(Me, CELL_UNTIL_DATE) Then
        ErrorManager.ThrowError _
            sheetRecapPenerimaan.InvalidUntilDate, _
            MODULE_NAME & ".SearchReceive_Click", _
            "Tanggal akhir tidak valid."
    End If

    fromDateValue = DateValue( _
        FormManager.GetCellValue(Me, CELL_FROM_DATE))

    untilDateValue = DateValue( _
        FormManager.GetCellValue(Me, CELL_UNTIL_DATE))

    If fromDateValue > untilDateValue Then
        ErrorManager.ThrowError _
            sheetRecapPenerimaan.InvalidDateRange, _
            MODULE_NAME & ".SearchReceive_Click", _
            "Tanggal awal tidak boleh lebih besar " & _
            "dari tanggal akhir."
    End If

    state = FormManager.BeginBusy( _
        "Mengambil data penerimaan...")

    ' =====================================================
    ' Step 2: Get and validate receive table
    ' =====================================================
    Set tbl = TableManager.GetTable(Me, TABLE_NAME)

    ValidateReceiveTableColumns tbl

    ' =====================================================
    ' Step 3: Call the service function
    ' =====================================================
    Set receives = serviceReceive.SelectReceives( _
        manufacturingPlantName:=plantName, _
        fromDate:=fromDateValue, _
        untilDate:=untilDateValue)

    If receives Is Nothing Then
        ErrorManager.ThrowError _
            sheetRecapPenerimaan.EmptyReceiveCollection, _
            MODULE_NAME & ".SearchReceive_Click", _
            "Hasil pencarian penerimaan tidak " & _
            "dapat dibaca."
    End If

    ' =====================================================
    ' Step 4: Clear existing table rows
    ' =====================================================
    TableManager.ClearRows tbl

    ' =====================================================
    ' Step 5: Fill table with search results
    ' =====================================================
    rowNumber = 0

    For Each receiveItem In receives

        rowNumber = rowNumber + 1

        rowIndex = TableManager.AddRow(tbl)

        ' No
        TableManager.SetValue _
            tbl, rowIndex, COLUMN_NO, rowNumber

        ' Order UUID
        TableManager.SetValue _
            tbl, rowIndex, COLUMN_ORDER_UUID, _
            receiveItem.Order.UUID

        ' Receive UUID
        TableManager.SetValue _
            tbl, rowIndex, COLUMN_RECEIVE_UUID, _
            receiveItem.UUID

        ' PO ID
        TableManager.SetValue _
            tbl, rowIndex, COLUMN_PO_ID, _
            receiveItem.Order.poID

        ' Catalogue
        TableManager.SetValue _
            tbl, rowIndex, COLUMN_CATALOGUE, _
            receiveItem.Order.Catalogue.Name

        ' Arrival date
        If receiveItem.ArrivalDate > 0 Then
            TableManager.SetValue _
                tbl, rowIndex, COLUMN_ARRIVAL_DATE, _
                receiveItem.ArrivalDate
        End If

        ' Verification date
        If receiveItem.VerificationDate > 0 Then
            TableManager.SetValue _
                tbl, rowIndex, COLUMN_VERIFICATION_DATE, _
                receiveItem.VerificationDate
        End If

        ' Unit of measurement
        TableManager.SetValue _
            tbl, rowIndex, COLUMN_UOM, _
            receiveItem.unitOfMeasurement

        ' Remaining order
        TableManager.SetValue _
            tbl, rowIndex, COLUMN_ORDER_AMOUNT, _
            receiveItem.RemainingOrder

        ' Total receive
        TableManager.SetValue _
            tbl, rowIndex, COLUMN_RECEIVE_AMOUNT, _
            receiveItem.totalReceiveAmount

        ' OK amount
        TableManager.SetValue _
            tbl, rowIndex, COLUMN_OK_AMOUNT, _
            receiveItem.okAmount

        ' NOT OK amount
        TableManager.SetValue _
            tbl, rowIndex, COLUMN_NOT_OK_AMOUNT, _
            receiveItem.notOKAmount

        ' Remarks
        TableManager.SetValue _
            tbl, rowIndex, COLUMN_REMARKS, _
            receiveItem.remarks

        ' PIC
        TableManager.SetValue _
            tbl, rowIndex, COLUMN_PIC, _
            receiveItem.personInCharge

        ' Update action
        TableManager.SetValue _
            tbl, rowIndex, COLUMN_UPDATE, "UPDATE"

    Next receiveItem

    ' =====================================================
    ' Step 6: Apply table formatting
    ' =====================================================
    ApplyReceiveTableFormat tbl

    ' =====================================================
    ' Step 7: Show result
    ' =====================================================
    If rowNumber = 0 Then
        FormManager.ShowInfo _
            "Tidak ada data penerimaan yang ditemukan."
    Else
        FormManager.ShowInfo _
            CStr(rowNumber) & _
            " data penerimaan ditemukan."
    End If

CleanUp:
    Set tbl = Nothing
    Set receiveItem = Nothing
    Set receives = Nothing

    FormManager.EndBusy state

    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    Set tbl = Nothing
    Set receiveItem = Nothing
    Set receives = Nothing

    FormManager.EndBusy state

    ErrorManager.ShowError errInfo

End Sub


' =========================================================
' CREATE RECEIVE UPDATE REQUEST
' =========================================================

Private Sub CreateReceiveUpdateRequest( _
    ByVal tbl As ListObject, _
    ByVal selectedRowIndex As Long)

    On Error GoTo ErrorHandler

    Dim originalUUID As String
    Dim orderTransactionUUID As String
    Dim plantName As String

    Dim catalogueName As String
    Dim unitOfMeasurement As String
    Dim poID As String
    Dim remarks As String
    Dim personInCharge As String

    Dim arrivalDateValue As Variant
    Dim verificationDateValue As Variant

    Dim parsedArrivalDate As Date
    Dim parsedVerificationDate As Date

    Dim totalReceiveAmount As Double
    Dim okAmount As Double
    Dim notOKAmount As Double

    Dim confirmationMessage As String
    Dim successMessage As String
    Dim mismatchMessage As String

    Dim proposedReceive As Receive
    Dim proposedOrder As Order
    Dim proposedCatalogue As Catalogue
    Dim proposedPlant As Plant

    Dim errInfo As ErrorInfo

    ' =====================================================
    ' Step 1: Validate selected row
    ' =====================================================
    If tbl Is Nothing Then
        ErrorManager.ThrowError _
            sheetRecapPenerimaan.EmptyReceiveTable, _
            MODULE_NAME & ".CreateReceiveUpdateRequest", _
            "Tabel penerimaan belum diinisialisasi."
    End If

    If Not TableManager.HasRows(tbl) Then
        ErrorManager.ThrowError _
            sheetRecapPenerimaan.EmptyReceiveTableRows, _
            MODULE_NAME & ".CreateReceiveUpdateRequest", _
            "Tabel penerimaan tidak memiliki data."
    End If

    If Not TableManager.IsValidRow(tbl, selectedRowIndex) Then
        ErrorManager.ThrowError _
            sheetRecapPenerimaan.InvalidSelectedRow, _
            MODULE_NAME & ".CreateReceiveUpdateRequest", _
            "Baris penerimaan yang dipilih tidak valid."
    End If

    ' =====================================================
    ' Step 2: Read UUID values
    ' =====================================================
    originalUUID = TableManager.GetText( _
        tbl, selectedRowIndex, COLUMN_RECEIVE_UUID)

    orderTransactionUUID = TableManager.GetText( _
        tbl, selectedRowIndex, COLUMN_ORDER_UUID)

    If ValidationText.IsBlank(originalUUID) Then
        ErrorManager.ThrowError _
            sheetRecapPenerimaan.EmptyOriginalReceiveUUID, _
            MODULE_NAME & ".CreateReceiveUpdateRequest", _
            "UUID transaksi receive original " & _
            "tidak ditemukan."
    End If

    If ValidationText.IsBlank(orderTransactionUUID) Then
        ErrorManager.ThrowError _
            sheetRecapPenerimaan.EmptyOrderTransactionUUID, _
            MODULE_NAME & ".CreateReceiveUpdateRequest", _
            "UUID transaksi order tidak ditemukan."
    End If

    ' =====================================================
    ' Step 3: Read manufacturing plant
    ' =====================================================
    plantName = FormManager.GetCellText(Me, CELL_LOCATION)

    If ValidationText.IsBlank(plantName) Then
        ErrorManager.ThrowError _
            sheetRecapPenerimaan.EmptyUpdatePlantName, _
            MODULE_NAME & ".CreateReceiveUpdateRequest", _
            "Lokasi manufacturing plant " & _
            "tidak ditemukan."
    End If

    ' =====================================================
    ' Step 4: Read text values
    ' =====================================================
    catalogueName = TableManager.GetText( _
        tbl, selectedRowIndex, COLUMN_CATALOGUE)

    unitOfMeasurement = TableManager.GetText( _
        tbl, selectedRowIndex, COLUMN_UOM)

    poID = TableManager.GetText( _
        tbl, selectedRowIndex, COLUMN_PO_ID)

    remarks = TableManager.GetText( _
        tbl, selectedRowIndex, COLUMN_REMARKS)

    personInCharge = TableManager.GetText( _
        tbl, selectedRowIndex, COLUMN_PIC)

    If ValidationText.IsBlank(catalogueName) Then
        ErrorManager.ThrowError _
            sheetRecapPenerimaan.EmptyCatalogueName, _
            MODULE_NAME & ".CreateReceiveUpdateRequest", _
            "Nama barang tidak boleh kosong."
    End If

    If ValidationText.IsBlank(unitOfMeasurement) Then
        ErrorManager.ThrowError _
            sheetRecapPenerimaan.EmptyUnitOfMeasurement, _
            MODULE_NAME & ".CreateReceiveUpdateRequest", _
            "Satuan barang tidak boleh kosong."
    End If

    If ValidationText.IsBlank(personInCharge) Then
        ErrorManager.ThrowError _
            sheetRecapPenerimaan.EmptyPersonInCharge, _
            MODULE_NAME & ".CreateReceiveUpdateRequest", _
            "PIC penerimaan tidak boleh kosong."
    End If

    ' =====================================================
    ' Step 5: Validate arrival date
    ' =====================================================
    arrivalDateValue = TableManager.GetValue( _
        tbl, selectedRowIndex, COLUMN_ARRIVAL_DATE)

    If ValidationText.IsBlank(arrivalDateValue) Then
        ErrorManager.ThrowError _
            sheetRecapPenerimaan.EmptyArrivalDate, _
            MODULE_NAME & ".CreateReceiveUpdateRequest", _
            "Tanggal datang tidak boleh kosong."
    End If

    If Not IsDate(arrivalDateValue) Then
        ErrorManager.ThrowError _
            sheetRecapPenerimaan.InvalidArrivalDate, _
            MODULE_NAME & ".CreateReceiveUpdateRequest", _
            "Tanggal datang tidak valid."
    End If

    parsedArrivalDate = TableManager.GetDate( _
        tbl, selectedRowIndex, COLUMN_ARRIVAL_DATE)

    ' =====================================================
    ' Step 6: Validate verification date
    ' =====================================================
    verificationDateValue = TableManager.GetValue( _
        tbl, selectedRowIndex, COLUMN_VERIFICATION_DATE)

    If ValidationText.IsNotBlank(verificationDateValue) Then

        If Not IsDate(verificationDateValue) Then
            ErrorManager.ThrowError _
                sheetRecapPenerimaan.InvalidVerificationDate, _
                MODULE_NAME & ".CreateReceiveUpdateRequest", _
                "Tanggal verifikasi tidak valid."
        End If

        parsedVerificationDate = TableManager.GetDate( _
            tbl, selectedRowIndex, COLUMN_VERIFICATION_DATE)

        If DateValue(parsedVerificationDate) < _
           DateValue(parsedArrivalDate) Then

            ErrorManager.ThrowError _
                sheetRecapPenerimaan.VerificationBeforeArrival, _
                MODULE_NAME & ".CreateReceiveUpdateRequest", _
                "Tanggal verifikasi tidak boleh " & _
                "lebih kecil dari tanggal datang."
        End If
    End If

    ' =====================================================
    ' Step 7: Validate amounts
    ' =====================================================
    ValidateNumericTableValue _
        tbl:=tbl, _
        columnName:=COLUMN_RECEIVE_AMOUNT, _
        selectedRowIndex:=selectedRowIndex, _
        fieldDescription:="Jumlah diterima", _
        emptyErrorCode:=sheetRecapPenerimaan.EmptyReceiveAmount, _
        invalidErrorCode:=sheetRecapPenerimaan.InvalidReceiveAmount

    ValidateNumericTableValue _
        tbl:=tbl, _
        columnName:=COLUMN_OK_AMOUNT, _
        selectedRowIndex:=selectedRowIndex, _
        fieldDescription:="Jumlah OK", _
        emptyErrorCode:=sheetRecapPenerimaan.EmptyOkAmount, _
        invalidErrorCode:=sheetRecapPenerimaan.InvalidOkAmount

    ValidateNumericTableValue _
        tbl:=tbl, _
        columnName:=COLUMN_NOT_OK_AMOUNT, _
        selectedRowIndex:=selectedRowIndex, _
        fieldDescription:="Jumlah NOT OK", _
        emptyErrorCode:=sheetRecapPenerimaan.EmptyNotOkAmount, _
        invalidErrorCode:=sheetRecapPenerimaan.InvalidNotOkAmount

    totalReceiveAmount = TableManager.GetDouble( _
        tbl, selectedRowIndex, COLUMN_RECEIVE_AMOUNT)

    okAmount = TableManager.GetDouble( _
        tbl, selectedRowIndex, COLUMN_OK_AMOUNT)

    notOKAmount = TableManager.GetDouble( _
        tbl, selectedRowIndex, COLUMN_NOT_OK_AMOUNT)

    If ValidationNumeric.IsZeroOrNegative(totalReceiveAmount) Then
        ErrorManager.ThrowError _
            sheetRecapPenerimaan.ZeroOrNegativeReceiveAmount, _
            MODULE_NAME & ".CreateReceiveUpdateRequest", _
            "Jumlah diterima harus lebih besar dari nol."
    End If

    If ValidationNumeric.IsNegative(okAmount) Then
        ErrorManager.ThrowError _
            sheetRecapPenerimaan.NegativeOkAmount, _
            MODULE_NAME & ".CreateReceiveUpdateRequest", _
            "Jumlah OK tidak boleh kurang dari nol."
    End If

    If ValidationNumeric.IsNegative(notOKAmount) Then
        ErrorManager.ThrowError _
            sheetRecapPenerimaan.NegativeNotOkAmount, _
            MODULE_NAME & ".CreateReceiveUpdateRequest", _
            "Jumlah NOT OK tidak boleh kurang dari nol."
    End If

    If ValidationNumeric.IsGreaterThan( _
        Abs(totalReceiveAmount - (okAmount + notOKAmount)), _
        0.000001) Then

        mismatchMessage = _
            "Jumlah OK ditambah jumlah NOT OK harus " & _
            "sama dengan jumlah diterima."

        mismatchMessage = mismatchMessage & _
            vbCrLf & vbCrLf & _
            "Jumlah diterima: " & _
            Format$(totalReceiveAmount, "#,##0.00")

        mismatchMessage = mismatchMessage & _
            vbCrLf & _
            "Jumlah OK: " & _
            Format$(okAmount, "#,##0.00")

        mismatchMessage = mismatchMessage & _
            vbCrLf & _
            "Jumlah NOT OK: " & _
            Format$(notOKAmount, "#,##0.00")

        ErrorManager.ThrowError _
            sheetRecapPenerimaan.InvalidOkNotOkTotal, _
            MODULE_NAME & ".CreateReceiveUpdateRequest", _
            mismatchMessage
    End If

    ' =====================================================
    ' Step 8: Build catalogue object
    ' =====================================================
    Set proposedCatalogue = New Catalogue

    With proposedCatalogue
        .Name = catalogueName
        .unitOfMeasurement = unitOfMeasurement
    End With

    ' =====================================================
    ' Step 9: Build plant object
    ' =====================================================
    Set proposedPlant = New Plant

    proposedPlant.Name = plantName

    ' =====================================================
    ' Step 10: Build related order object
    ' =====================================================
    Set proposedOrder = New Order

    With proposedOrder
        .UUID = orderTransactionUUID
        .poID = poID

        Set .Catalogue = proposedCatalogue
        Set .Plant = proposedPlant
        .Amount = TableManager.GetDouble( _
            tbl, selectedRowIndex, COLUMN_ORDER_AMOUNT)
    End With

    ' =====================================================
    ' Step 11: Build proposed receive object
    ' =====================================================
    Set proposedReceive = New Receive

    With proposedReceive
        Set .Order = proposedOrder

        .ArrivalDate = parsedArrivalDate

        If ValidationText.IsNotBlank(verificationDateValue) Then
            .VerificationDate = parsedVerificationDate
        End If

        .totalReceiveAmount = totalReceiveAmount
        .okAmount = okAmount
        .notOKAmount = notOKAmount
        .remarks = remarks
        .SystemRemarks = vbNullString
        .personInCharge = personInCharge
    End With

    ' =====================================================
    ' Step 12: Build confirmation message
    ' =====================================================
    confirmationMessage = _
        "Apakah Anda yakin ingin membuat permintaan " & _
        "perubahan penerimaan berikut?"

    confirmationMessage = confirmationMessage & _
        vbCrLf & vbCrLf & _
        "UUID receive original:" & vbCrLf & _
        originalUUID

    confirmationMessage = confirmationMessage & _
        vbCrLf & vbCrLf & _
        "UUID order:" & vbCrLf & _
        orderTransactionUUID

    confirmationMessage = confirmationMessage & _
        vbCrLf & vbCrLf & _
        "PO ID: " & poID

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "Nama barang: " & catalogueName

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "Tanggal datang: " & _
        Format$( _
            proposedReceive.ArrivalDate, _
            "dd-mmm-yyyy")

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        GetVerificationDateMessage( _
            verificationDateValue)

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "Jumlah diterima: " & _
        Format$(totalReceiveAmount, "#,##0.00") & _
        " " & unitOfMeasurement

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "Jumlah OK: " & _
        Format$(okAmount, "#,##0.00") & _
        " " & unitOfMeasurement

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "Jumlah NOT OK: " & _
        Format$(notOKAmount, "#,##0.00") & _
        " " & unitOfMeasurement

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "PIC: " & personInCharge

    confirmationMessage = confirmationMessage & _
        vbCrLf & vbCrLf & _
        "Transaksi receive proposed akan dibuat dalam " & _
        "kondisi tidak aktif dan menunggu persetujuan."

    If Not FormManager.Confirm( _
        confirmationMessage, _
        "Konfirmasi Permintaan Update Penerimaan") Then

        GoTo CleanUp
    End If

    ' =====================================================
    ' Step 13: Create proposed receive update
    ' =====================================================
    FormManager.SetStatus _
        "Membuat permintaan update penerimaan...", True

    serviceReceive.CreateReceive _
        receiveItem:=proposedReceive, _
        orderTransactionUUID:=orderTransactionUUID, _
        originalUUID:=originalUUID

    ' =====================================================
    ' Step 14: Show success message
    ' =====================================================
    FormManager.ClearStatus

    successMessage = _
        "Permintaan perubahan penerimaan berhasil dibuat."

    successMessage = successMessage & _
        vbCrLf & vbCrLf & _
        "UUID receive original:" & vbCrLf & _
        originalUUID

    successMessage = successMessage & _
        vbCrLf & vbCrLf & _
        "Status: Menunggu persetujuan."

    FormManager.ShowInfo _
        successMessage, _
        "Permintaan Update Berhasil"

    ' =====================================================
    ' Step 15: Refresh receive data
    ' =====================================================
    SearchReceive_Click

CleanUp:
    FormManager.ClearStatus

    Set proposedReceive = Nothing
    Set proposedOrder = Nothing
    Set proposedCatalogue = Nothing
    Set proposedPlant = Nothing

    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    FormManager.ClearStatus

    Set proposedReceive = Nothing
    Set proposedOrder = Nothing
    Set proposedCatalogue = Nothing
    Set proposedPlant = Nothing

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".CreateReceiveUpdateRequest"

End Sub


' =========================================================
' TABLE FORMATTING
' =========================================================

Private Sub ApplyReceiveTableFormat( _
    ByVal tbl As ListObject)

    On Error GoTo ErrorHandler

    Dim dateColumns As Variant
    Dim amountColumns As Variant
    Dim allColumns As Variant
    Dim columnName As Variant
    Dim columnRange As Range

    Dim errInfo As ErrorInfo

    If tbl Is Nothing Then
        ErrorManager.ThrowError _
            sheetRecapPenerimaan.EmptyFormatReceiveTable, _
            MODULE_NAME & ".ApplyReceiveTableFormat", _
            "Objek tabel penerimaan belum diinisialisasi."
    End If

    If Not TableManager.HasRows(tbl) Then
        GoTo CleanUp
    End If

    dateColumns = Array( _
        COLUMN_ARRIVAL_DATE, _
        COLUMN_VERIFICATION_DATE)

    amountColumns = Array( _
        COLUMN_ORDER_AMOUNT, _
        COLUMN_RECEIVE_AMOUNT, _
        COLUMN_OK_AMOUNT, _
        COLUMN_NOT_OK_AMOUNT)

    allColumns = Array( _
        COLUMN_NO, _
        COLUMN_ORDER_UUID, _
        COLUMN_RECEIVE_UUID, _
        COLUMN_PO_ID, _
        COLUMN_CATALOGUE, _
        COLUMN_ARRIVAL_DATE, _
        COLUMN_VERIFICATION_DATE, _
        COLUMN_UOM, _
        COLUMN_ORDER_AMOUNT, _
        COLUMN_RECEIVE_AMOUNT, _
        COLUMN_OK_AMOUNT, _
        COLUMN_NOT_OK_AMOUNT, _
        COLUMN_REMARKS, _
        COLUMN_PIC, _
        COLUMN_UPDATE)

    ' =====================================================
    ' Date formats
    ' =====================================================
    For Each columnName In dateColumns

        TableManager.FormatColumn _
            tbl, CStr(columnName), "dd-mmm-yyyy"

    Next columnName

    ' =====================================================
    ' Amount formats
    ' =====================================================
    For Each columnName In amountColumns

        TableManager.FormatColumn _
            tbl, CStr(columnName), "#,##0.00"

    Next columnName

    ' =====================================================
    ' General alignment
    ' =====================================================
    For Each columnName In allColumns

        Set columnRange = TableManager.GetColumnRange( _
            tbl, CStr(columnName))

        If Not columnRange Is Nothing Then
            columnRange.VerticalAlignment = xlCenter
        End If

    Next columnName

    ' =====================================================
    ' UPDATE column format
    ' =====================================================
    Set columnRange = TableManager.GetColumnRange( _
        tbl, COLUMN_UPDATE)

    If Not columnRange Is Nothing Then

        With columnRange

            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter

            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)

            .Interior.Pattern = xlSolid
            .Interior.Color = RGB(0, 102, 204)
        End With

    End If

CleanUp:
    Set columnRange = Nothing
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    Set columnRange = Nothing

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".ApplyReceiveTableFormat"

End Sub


' =========================================================
' TABLE COLUMN VALIDATION
' =========================================================

Private Sub ValidateReceiveTableColumns( _
    ByVal tbl As ListObject)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    If tbl Is Nothing Then
        ErrorManager.ThrowError _
            sheetRecapPenerimaan.NotFoundReceiveTable, _
            MODULE_NAME & ".ValidateReceiveTableColumns", _
            "TableRecapReceive tidak ditemukan."
    End If

    TableManager.ValidateColumns tbl, Array( _
        COLUMN_NO, _
        COLUMN_ORDER_UUID, _
        COLUMN_RECEIVE_UUID, _
        COLUMN_PO_ID, _
        COLUMN_CATALOGUE, _
        COLUMN_ARRIVAL_DATE, _
        COLUMN_VERIFICATION_DATE, _
        COLUMN_UOM, _
        COLUMN_ORDER_AMOUNT, _
        COLUMN_RECEIVE_AMOUNT, _
        COLUMN_OK_AMOUNT, _
        COLUMN_NOT_OK_AMOUNT, _
        COLUMN_REMARKS, _
        COLUMN_PIC, _
        COLUMN_UPDATE)

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".ValidateReceiveTableColumns"

End Sub


' =========================================================
' VALUE VALIDATION HELPERS
' =========================================================

Private Sub ValidateNumericTableValue( _
    ByVal tbl As ListObject, _
    ByVal columnName As String, _
    ByVal selectedRowIndex As Long, _
    ByVal fieldDescription As String, _
    ByVal emptyErrorCode As Long, _
    ByVal invalidErrorCode As Long)

    On Error GoTo ErrorHandler

    Dim fieldValue As Variant

    Dim errInfo As ErrorInfo

    fieldValue = TableManager.GetValue( _
        tbl, selectedRowIndex, columnName)

    If ValidationText.IsBlank(fieldValue) Then
        ErrorManager.ThrowError _
            emptyErrorCode, _
            MODULE_NAME & ".ValidateNumericTableValue", _
            fieldDescription & _
            " tidak boleh kosong."
    End If

    If ValidationNumeric.IsNotNumber(fieldValue) Then
        ErrorManager.ThrowError _
            invalidErrorCode, _
            MODULE_NAME & ".ValidateNumericTableValue", _
            fieldDescription & _
            " harus berupa angka."
    End If

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".ValidateNumericTableValue"

End Sub


' =========================================================
' MESSAGE HELPERS
' =========================================================

Private Function GetVerificationDateMessage( _
    ByVal verificationDateValue As Variant) As String

    If ValidationText.IsNotBlank(verificationDateValue) Then

        If IsDate(verificationDateValue) Then
            GetVerificationDateMessage = _
                "Tanggal verifikasi: " & _
                Format$( _
                    CDate(verificationDateValue), _
                    "dd-mmm-yyyy")
        Else
            GetVerificationDateMessage = _
                "Tanggal verifikasi: Tidak valid"
        End If

    Else
        GetVerificationDateMessage = _
            "Tanggal verifikasi: Belum diisi"
    End If

End Function
