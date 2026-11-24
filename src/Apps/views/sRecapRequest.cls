Attribute VB_Name = "sRecapRequest"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Const MODULE_NAME As String = _
    "SHEET_RECAP_REQUEST"

Private Const TABLE_NAME As String = "TableRecapRequest"

Private Const CELL_LOCATION As String = "B4"
Private Const CELL_FROM_DATE As String = "B7"
Private Const CELL_UNTIL_DATE As String = "B8"

Private Const COLUMN_NO As String = "No"
Private Const COLUMN_UUID As String = "UUID"
Private Const COLUMN_MOBILE_POST As String = "Mobile Post"
Private Const COLUMN_CATALOGUE As String = "Nama Barang"
Private Const COLUMN_AMOUNT As String = "Jumlah Barang"
Private Const COLUMN_UOM As String = "Satuan"
Private Const COLUMN_ORDER_DATE As String = "Tanggal Order"
Private Const COLUMN_REMARKS As String = "Remarks"
Private Const COLUMN_PIC As String = "PIC"
Private Const COLUMN_UPDATE As String = "UPDATE"

Private Sub Worksheet_SelectionChange( _
    ByVal Target As Range)

    On Error GoTo ErrorHandler

    Dim tbl As ListObject
    Dim selectedRowIndex As Long

    Dim state As FormState
    Dim errInfo As ErrorInfo

    ' Only handle one selected cell.
    If Target.CountLarge <> 1 Then
        Exit Sub
    End If

    state = managerForm.CaptureState()

    ' =====================================================
    ' Date input cells
    ' =====================================================
    If managerForm.IsInRange( _
        Target, _
        Me.Range( _
            CELL_FROM_DATE & ":" & CELL_UNTIL_DATE _
        )) Then

        Application.EnableEvents = False

        managerForm.ShowDatePicker Target

        GoTo CleanUp
    End If

    ' =====================================================
    ' Get request recap table
    ' =====================================================
    Set tbl = managerTable.GetTable(Me, TABLE_NAME)

    If Not managerTable.HasRows(tbl) Then
        GoTo CleanUp
    End If

    ' =====================================================
    ' UPDATE column clicked
    ' =====================================================
    If managerTable.IsInColumn( _
        Target, _
        tbl, _
        COLUMN_UPDATE) Then

        If Not managerValidationText.IsSameText( _
            TrimText(CStr(Target.value)), _
            "UPDATE") Then

            GoTo CleanUp
        End If

        Application.EnableEvents = False

        selectedRowIndex = _
            managerTable.GetRowIndex(tbl, Target)

        CreateRequestUpdateProposal _
            tbl:=tbl, _
            selectedRowIndex:=selectedRowIndex, _
            actionCell:=Target

        GoTo CleanUp
    End If

CleanUp:
    Set tbl = Nothing

    managerForm.EndBusy state
    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    Set tbl = Nothing

    managerForm.EndBusy state
    managerForm.ClearStatus

    managerError.ShowError errInfo

End Sub


Public Sub SearchRequest_Click()

    On Error GoTo ErrorHandler

    Dim plantName As String
    Dim fromDateValue As Date
    Dim untilDateValue As Date

    Dim requests As Collection
    Dim requestItem As Request

    Dim tbl As ListObject
    Dim rowIndex As Long
    Dim rowNumber As Long

    Dim errInfo As ErrorInfo

    plantName = managerForm.GetCellText(Me, CELL_LOCATION)

    If managerValidationText.IsBlank(plantName) Then
        managerError.ThrowError _
            sheetRecapRequest.EmptyPlantName, _
            MODULE_NAME & ".SearchRequest_Click", _
            "Nama manufacturing plant tidak boleh kosong."
    End If

    If Not managerForm.IsCellDate(Me, CELL_FROM_DATE) Then
        managerError.ThrowError _
            sheetRecapRequest.InvalidFromDate, _
            MODULE_NAME & ".SearchRequest_Click", _
            "Tanggal awal tidak valid."
    End If

    If Not managerForm.IsCellDate(Me, CELL_UNTIL_DATE) Then
        managerError.ThrowError _
            sheetRecapRequest.InvalidUntilDate, _
            MODULE_NAME & ".SearchRequest_Click", _
            "Tanggal akhir tidak valid."
    End If

    fromDateValue = DateValue( _
        CDate(managerForm.GetCellValue(Me, CELL_FROM_DATE)))

    untilDateValue = DateValue( _
        CDate(managerForm.GetCellValue(Me, CELL_UNTIL_DATE)))

    If fromDateValue > untilDateValue Then
        managerError.ThrowError _
            sheetRecapRequest.InvalidDateRange, _
            MODULE_NAME & ".SearchRequest_Click", _
            "Tanggal awal tidak boleh lebih besar " & _
            "dari tanggal akhir."
    End If

    Set requests = _
        serviceRequest. _
        GetRequestsByManufacturingPlantAndOrderDate( _
            manufacturingPlantName:=plantName, _
            fromDate:=fromDateValue, _
            untilDate:=untilDateValue)

    Set tbl = managerTable.GetTable(Me, TABLE_NAME)

    ' =====================================================
    ' Step 4: Clear the existing table rows
    ' =====================================================
    managerTable.ClearRows tbl

    ' =====================================================
    ' Step 5: Fill the table with the result
    ' =====================================================
    rowNumber = 0

    For Each requestItem In requests

        rowNumber = rowNumber + 1

        rowIndex = managerTable.AddRow(tbl)

        managerTable.SetValue _
            tbl, rowIndex, COLUMN_NO, rowNumber

        managerTable.SetValue _
            tbl, rowIndex, COLUMN_UUID, _
            requestItem.UUID

        managerTable.SetValue _
            tbl, rowIndex, COLUMN_MOBILE_POST, _
            requestItem.MobilePost.Name

        managerTable.SetValue _
            tbl, rowIndex, COLUMN_CATALOGUE, _
            requestItem.Catalogue.Name

        managerTable.SetValue _
            tbl, rowIndex, COLUMN_AMOUNT, _
            requestItem.Amount

        managerTable.SetValue _
            tbl, rowIndex, COLUMN_UOM, _
            requestItem.unitOfMeasurement

        If requestItem.orderDate > 0 Then

            managerTable.SetValue _
                tbl, rowIndex, COLUMN_ORDER_DATE, _
                requestItem.orderDate

        End If

        managerTable.SetValue _
            tbl, rowIndex, COLUMN_REMARKS, _
            requestItem.remarks

        managerTable.SetValue _
            tbl, rowIndex, COLUMN_PIC, _
            requestItem.personInCharge

        managerTable.SetValue _
            tbl, rowIndex, COLUMN_UPDATE, "UPDATE"

    Next requestItem

    ApplyRequestTableFormat tbl

CleanUp:
    Set tbl = Nothing
    Set requestItem = Nothing
    Set requests = Nothing

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    Set tbl = Nothing
    Set requestItem = Nothing
    Set requests = Nothing

    managerError.ShowError errInfo

End Sub

Private Sub ApplyRequestTableFormat( _
    ByVal tbl As ListObject)

    On Error GoTo ErrorHandler

    Dim columnRange As Range

    Dim errInfo As ErrorInfo

    If Not managerTable.HasRows(tbl) Then
        GoTo CleanUp
    End If

    managerTable.FormatColumn _
        tbl, COLUMN_ORDER_DATE, "dd-mmm-yyyy"

    managerTable.FormatColumn _
        tbl, COLUMN_AMOUNT, "#,##0.00"

    Set columnRange = managerTable.GetColumnRange( _
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
    Set tbl = Nothing
    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    Set columnRange = Nothing
    Set tbl = Nothing

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ApplyRequestTableFormat"

End Sub

Private Sub CreateRequestUpdateProposal( _
    ByVal tbl As ListObject, _
    ByVal selectedRowIndex As Long, _
    ByVal actionCell As Range)

    On Error GoTo ErrorHandler

    Dim originalUUID As String
    Dim proposedUUID As String

    Dim manufacturingPlantName As String
    Dim mobilePostName As String
    Dim catalogueName As String
    Dim unitOfMeasurement As String
    Dim remarks As String
    Dim personInCharge As String

    Dim requestDateValue As Variant
    Dim requestAmountValue As Variant

    Dim requestDate As Date
    Dim requestAmount As Double

    Dim proposedRequest As Request
    Dim proposedMobilePost As MobilePost
    Dim proposedPlant As Plant
    Dim proposedCatalogue As Catalogue

    Dim confirmationMessage As String
    Dim successMessage As String

    Dim actionCellChanged As Boolean

    Dim errInfo As ErrorInfo

    actionCellChanged = False

    ' =====================================================
    ' Step 1: Validate table
    ' =====================================================
    If tbl Is Nothing Then
        managerError.ThrowError _
            sheetRecapRequest.EmptyRequestTable, _
            MODULE_NAME & ".CreateRequestUpdateProposal", _
            "Tabel request belum diinisialisasi."
    End If

    If Not managerTable.HasRows(tbl) Then
        managerError.ThrowError _
            sheetRecapRequest.EmptyRequestTableRows, _
            MODULE_NAME & ".CreateRequestUpdateProposal", _
            "Tabel request tidak memiliki data."
    End If

    If Not managerTable.IsValidRow(tbl, selectedRowIndex) Then
        managerError.ThrowError _
            sheetRecapRequest.InvalidSelectedRow, _
            MODULE_NAME & ".CreateRequestUpdateProposal", _
            "Baris request yang dipilih tidak valid."
    End If

    ' =====================================================
    ' Step 2: Validate action cell
    ' =====================================================
    If actionCell Is Nothing Then
        managerError.ThrowError _
            sheetRecapRequest.EmptyActionCell, _
            MODULE_NAME & ".CreateRequestUpdateProposal", _
            "Cell UPDATE tidak ditemukan."
    End If

    If Not managerValidationText.IsSameText( _
        TrimText(CStr(actionCell.value)), _
        "UPDATE") Then

        GoTo CleanUp
    End If

    ' =====================================================
    ' Step 3: Read original transaction UUID
    ' =====================================================
    originalUUID = managerTable.GetText( _
        tbl, selectedRowIndex, COLUMN_UUID)

    If managerValidationText.IsBlank(originalUUID) Then
        managerError.ThrowError _
            sheetRecapRequest.EmptyOriginalRequestUUID, _
            MODULE_NAME & ".CreateRequestUpdateProposal", _
            "UUID transaksi request original " & _
            "tidak ditemukan."
    End If

    ' =====================================================
    ' Step 4: Read manufacturing plant
    ' =====================================================
    manufacturingPlantName = managerForm.GetCellText( _
        Me, CELL_LOCATION)

    If managerValidationText.IsBlank(manufacturingPlantName) Then
        managerError.ThrowError _
            sheetRecapRequest.EmptyUpdatePlantName, _
            MODULE_NAME & ".CreateRequestUpdateProposal", _
            "Nama manufacturing plant " & _
            "tidak boleh kosong."
    End If

    ' =====================================================
    ' Step 5: Read text values from selected row
    ' =====================================================
    mobilePostName = managerTable.GetText( _
        tbl, selectedRowIndex, COLUMN_MOBILE_POST)

    catalogueName = managerTable.GetText( _
        tbl, selectedRowIndex, COLUMN_CATALOGUE)

    unitOfMeasurement = managerTable.GetText( _
        tbl, selectedRowIndex, COLUMN_UOM)

    remarks = managerTable.GetText( _
        tbl, selectedRowIndex, COLUMN_REMARKS)

    personInCharge = managerTable.GetText( _
        tbl, selectedRowIndex, COLUMN_PIC)

    If managerValidationText.IsBlank(mobilePostName) Then
        managerError.ThrowError _
            sheetRecapRequest.EmptyMobilePostName, _
            MODULE_NAME & ".CreateRequestUpdateProposal", _
            "Nama mobile post tidak boleh kosong."
    End If

    If managerValidationText.IsBlank(catalogueName) Then
        managerError.ThrowError _
            sheetRecapRequest.EmptyCatalogueName, _
            MODULE_NAME & ".CreateRequestUpdateProposal", _
            "Nama barang tidak boleh kosong."
    End If

    If managerValidationText.IsBlank(unitOfMeasurement) Then
        managerError.ThrowError _
            sheetRecapRequest.EmptyUnitOfMeasurement, _
            MODULE_NAME & ".CreateRequestUpdateProposal", _
            "Satuan barang tidak boleh kosong."
    End If

    If managerValidationText.IsBlank(personInCharge) Then
        managerError.ThrowError _
            sheetRecapRequest.EmptyPersonInCharge, _
            MODULE_NAME & ".CreateRequestUpdateProposal", _
            "PIC request tidak boleh kosong."
    End If

    ' =====================================================
    ' Step 6: Read and validate request date
    ' =====================================================
    requestDateValue = managerTable.GetValue( _
        tbl, selectedRowIndex, COLUMN_ORDER_DATE)

    If managerValidationText.IsBlank(requestDateValue) Then
        managerError.ThrowError _
            sheetRecapRequest.EmptyRequestDate, _
            MODULE_NAME & ".CreateRequestUpdateProposal", _
            "Tanggal request tidak boleh kosong."
    End If

    If Not IsDate(requestDateValue) Then
        managerError.ThrowError _
            sheetRecapRequest.InvalidRequestDate, _
            MODULE_NAME & ".CreateRequestUpdateProposal", _
            "Tanggal request tidak valid."
    End If

    requestDate = managerTable.GetDate( _
        tbl, selectedRowIndex, COLUMN_ORDER_DATE)

    ' =====================================================
    ' Step 7: Read and validate request amount
    ' =====================================================
    requestAmountValue = managerTable.GetValue( _
        tbl, selectedRowIndex, COLUMN_AMOUNT)

    If managerValidationText.IsBlank(requestAmountValue) Then
        managerError.ThrowError _
            sheetRecapRequest.EmptyRequestAmount, _
            MODULE_NAME & ".CreateRequestUpdateProposal", _
            "Jumlah request tidak boleh kosong."
    End If

    If managerValidationNumeric.IsNotNumber(requestAmountValue) Then
        managerError.ThrowError _
            sheetRecapRequest.InvalidRequestAmount, _
            MODULE_NAME & ".CreateRequestUpdateProposal", _
            "Jumlah request harus berupa angka."
    End If

    requestAmount = managerTable.GetDouble( _
        tbl, selectedRowIndex, COLUMN_AMOUNT)

    If managerValidationNumeric.IsZeroOrNegative(requestAmount) Then
        managerError.ThrowError _
            sheetRecapRequest.ZeroOrNegativeRequestAmount, _
            MODULE_NAME & ".CreateRequestUpdateProposal", _
            "Jumlah request harus lebih besar dari nol."
    End If

    ' =====================================================
    ' Step 8: Build manufacturing plant object
    ' =====================================================
    Set proposedPlant = New Plant

    proposedPlant.Name = _
        manufacturingPlantName

    ' =====================================================
    ' Step 9: Build mobile post object
    ' =====================================================
    Set proposedMobilePost = New MobilePost

    With proposedMobilePost
        .Name = mobilePostName

        Set .Plant = proposedPlant
    End With

    ' =====================================================
    ' Step 10: Build catalogue object
    ' =====================================================
    Set proposedCatalogue = New Catalogue

    With proposedCatalogue
        .Name = catalogueName
        .unitOfMeasurement = unitOfMeasurement
    End With

    ' =====================================================
    ' Step 11: Build proposed request object
    ' =====================================================
    Set proposedRequest = New Request

    With proposedRequest
        Set .MobilePost = proposedMobilePost
        Set .Catalogue = proposedCatalogue

        .orderDate = requestDate
        .Amount = requestAmount
        .unitOfMeasurement = unitOfMeasurement
        .remarks = remarks
        .SystemRemarks = vbNullString
        .personInCharge = personInCharge
    End With

    ' =====================================================
    ' Step 12: Build confirmation message
    ' =====================================================
    confirmationMessage = _
        "Apakah Anda yakin ingin membuat " & _
        "permintaan perubahan request berikut?"

    confirmationMessage = confirmationMessage & _
        vbCrLf & vbCrLf & _
        "UUID request original:" & vbCrLf & _
        originalUUID

    confirmationMessage = confirmationMessage & _
        vbCrLf & vbCrLf & _
        "Manufacturing plant: " & _
        manufacturingPlantName

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "Mobile post: " & mobilePostName

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "Nama barang: " & catalogueName

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "Jumlah request: " & _
        Format$(requestAmount, "#,##0.00") & _
        " " & unitOfMeasurement

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "Tanggal request: " & _
        Format$( _
            requestDate, _
            "dd-mmm-yyyy")

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "PIC: " & personInCharge

    If Len(remarks) > 0 Then
        confirmationMessage = confirmationMessage & _
            vbCrLf & _
            "Remarks: " & remarks
    End If

    confirmationMessage = confirmationMessage & _
        vbCrLf & vbCrLf & _
        "Request proposed akan dibuat dalam kondisi " & _
        "tidak aktif dan menunggu persetujuan."

    If Not managerForm.Confirm( _
        confirmationMessage, _
        "Konfirmasi Permintaan Update Request") Then

        GoTo CleanUp
    End If

    ' =====================================================
    ' Step 13: Set temporary processing status
    ' =====================================================
    actionCell.value = "MEMPROSES..."
    actionCellChanged = True

    With actionCell
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)

        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(128, 128, 128)
    End With

    managerForm.SetStatus _
        "Membuat permintaan update request...", True

    ' =====================================================
    ' Step 14: Create proposed request
    '
    ' Supplying originalUUID causes CreateRequest to:
    ' 1. Create an inactive REQUEST transaction.
    ' 2. Create the proposed request record.
    ' 3. Create a transaction update request.
    ' 4. Avoid changing plant stock before approval.
    ' =====================================================
    proposedUUID = serviceRequest.CreateRequest( _
        requestItem:=proposedRequest, _
        originalUUID:=originalUUID)

    proposedUUID = TrimText(proposedUUID)

    If managerValidationText.IsBlank(proposedUUID) Then
        managerError.ThrowError _
            sheetRecapRequest.EmptyProposedRequestUUID, _
            MODULE_NAME & ".CreateRequestUpdateProposal", _
            "UUID transaksi request proposed " & _
            "tidak berhasil dibuat."
    End If

    If managerValidationText.IsSameText( _
        originalUUID, _
        proposedUUID) Then

        managerError.ThrowError _
            sheetRecapRequest.SameOriginalProposedUUID, _
            MODULE_NAME & ".CreateRequestUpdateProposal", _
            "UUID transaksi request original dan " & _
            "proposed tidak boleh sama."
    End If

    ' =====================================================
    ' Step 15: Mark action as submitted
    ' =====================================================
    actionCell.value = "DIAJUKAN"

    With actionCell
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)

        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(0, 128, 0)
    End With

    managerForm.ClearStatus

    ' =====================================================
    ' Step 16: Show success message
    ' =====================================================
    successMessage = _
        "Permintaan perubahan request berhasil dibuat."

    successMessage = successMessage & _
        vbCrLf & vbCrLf & _
        "UUID request original:" & vbCrLf & _
        originalUUID

    successMessage = successMessage & _
        vbCrLf & vbCrLf & _
        "UUID request proposed:" & vbCrLf & _
        proposedUUID

    successMessage = successMessage & _
        vbCrLf & vbCrLf & _
        "Status: Menunggu persetujuan."

    managerForm.ShowInfo _
        successMessage, _
        "Permintaan Update Berhasil"

    ' =====================================================
    ' Step 17: Refresh request recap
    ' =====================================================
    SearchRequest_Click

CleanUp:
    managerForm.ClearStatus

    Set proposedRequest = Nothing
    Set proposedMobilePost = Nothing
    Set proposedPlant = Nothing
    Set proposedCatalogue = Nothing

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    ' Restore the action cell when proposal creation fails.
    If Not actionCell Is Nothing Then

        If actionCellChanged Then
            actionCell.value = "UPDATE"

            With actionCell
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter

                .Font.Bold = True
                .Font.Color = RGB(255, 255, 255)

                .Interior.Pattern = xlSolid
                .Interior.Color = RGB(0, 102, 204)
            End With
        End If

    End If

    managerForm.ClearStatus

    Set proposedRequest = Nothing
    Set proposedMobilePost = Nothing
    Set proposedPlant = Nothing
    Set proposedCatalogue = Nothing

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".CreateRequestUpdateProposal"

End Sub
