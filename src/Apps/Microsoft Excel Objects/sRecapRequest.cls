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
    ' Get request recap table
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

        CreateRequestUpdateProposal _
            tbl:=tbl, _
            selectedRowIndex:=selectedRowIndex, _
            actionCell:=Target

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

    plantName = FormManager.GetCellText(Me, CELL_LOCATION)

    If ValidationText.IsBlank(plantName) Then
        ErrorManager.ThrowError _
            sheetRecapRequest.EmptyPlantName, _
            MODULE_NAME & ".SearchRequest_Click", _
            "Nama manufacturing plant tidak boleh kosong."
    End If

    If Not FormManager.IsCellDate(Me, CELL_FROM_DATE) Then
        ErrorManager.ThrowError _
            sheetRecapRequest.InvalidFromDate, _
            MODULE_NAME & ".SearchRequest_Click", _
            "Tanggal awal tidak valid."
    End If

    If Not FormManager.IsCellDate(Me, CELL_UNTIL_DATE) Then
        ErrorManager.ThrowError _
            sheetRecapRequest.InvalidUntilDate, _
            MODULE_NAME & ".SearchRequest_Click", _
            "Tanggal akhir tidak valid."
    End If

    fromDateValue = DateValue( _
        CDate(FormManager.GetCellValue(Me, CELL_FROM_DATE)))

    untilDateValue = DateValue( _
        CDate(FormManager.GetCellValue(Me, CELL_UNTIL_DATE)))

    If fromDateValue > untilDateValue Then
        ErrorManager.ThrowError _
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

    Set tbl = TableManager.GetTable(Me, TABLE_NAME)

    ' =====================================================
    ' Step 4: Clear the existing table rows
    ' =====================================================
    TableManager.ClearRows tbl

    ' =====================================================
    ' Step 5: Fill the table with the result
    ' =====================================================
    rowNumber = 0

    For Each requestItem In requests

        rowNumber = rowNumber + 1

        rowIndex = TableManager.AddRow(tbl)

        TableManager.SetValue _
            tbl, rowIndex, COLUMN_NO, rowNumber

        TableManager.SetValue _
            tbl, rowIndex, COLUMN_UUID, _
            requestItem.UUID

        TableManager.SetValue _
            tbl, rowIndex, COLUMN_MOBILE_POST, _
            requestItem.MobilePost.Name

        TableManager.SetValue _
            tbl, rowIndex, COLUMN_CATALOGUE, _
            requestItem.Catalogue.Name

        TableManager.SetValue _
            tbl, rowIndex, COLUMN_AMOUNT, _
            requestItem.Amount

        TableManager.SetValue _
            tbl, rowIndex, COLUMN_UOM, _
            requestItem.unitOfMeasurement

        If requestItem.orderDate > 0 Then

            TableManager.SetValue _
                tbl, rowIndex, COLUMN_ORDER_DATE, _
                requestItem.orderDate

        End If

        TableManager.SetValue _
            tbl, rowIndex, COLUMN_REMARKS, _
            requestItem.remarks

        TableManager.SetValue _
            tbl, rowIndex, COLUMN_PIC, _
            requestItem.personInCharge

        TableManager.SetValue _
            tbl, rowIndex, COLUMN_UPDATE, "UPDATE"

    Next requestItem

    ApplyRequestTableFormat tbl

CleanUp:
    Set tbl = Nothing
    Set requestItem = Nothing
    Set requests = Nothing

    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    Set tbl = Nothing
    Set requestItem = Nothing
    Set requests = Nothing

    ErrorManager.ShowError errInfo

End Sub

Private Sub ApplyRequestTableFormat( _
    ByVal tbl As ListObject)

    On Error GoTo ErrorHandler

    Dim columnRange As Range

    Dim errInfo As ErrorInfo

    If Not TableManager.HasRows(tbl) Then
        GoTo CleanUp
    End If

    TableManager.FormatColumn _
        tbl, COLUMN_ORDER_DATE, "dd-mmm-yyyy"

    TableManager.FormatColumn _
        tbl, COLUMN_AMOUNT, "#,##0.00"

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
    Set tbl = Nothing
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    Set columnRange = Nothing
    Set tbl = Nothing

    ErrorManager.Rethrow errInfo, _
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
        ErrorManager.ThrowError _
            sheetRecapRequest.EmptyRequestTable, _
            MODULE_NAME & ".CreateRequestUpdateProposal", _
            "Tabel request belum diinisialisasi."
    End If

    If Not TableManager.HasRows(tbl) Then
        ErrorManager.ThrowError _
            sheetRecapRequest.EmptyRequestTableRows, _
            MODULE_NAME & ".CreateRequestUpdateProposal", _
            "Tabel request tidak memiliki data."
    End If

    If Not TableManager.IsValidRow(tbl, selectedRowIndex) Then
        ErrorManager.ThrowError _
            sheetRecapRequest.InvalidSelectedRow, _
            MODULE_NAME & ".CreateRequestUpdateProposal", _
            "Baris request yang dipilih tidak valid."
    End If

    ' =====================================================
    ' Step 2: Validate action cell
    ' =====================================================
    If actionCell Is Nothing Then
        ErrorManager.ThrowError _
            sheetRecapRequest.EmptyActionCell, _
            MODULE_NAME & ".CreateRequestUpdateProposal", _
            "Cell UPDATE tidak ditemukan."
    End If

    If Not ValidationText.IsSameText( _
        TrimText(CStr(actionCell.value)), _
        "UPDATE") Then

        GoTo CleanUp
    End If

    ' =====================================================
    ' Step 3: Read original transaction UUID
    ' =====================================================
    originalUUID = TableManager.GetText( _
        tbl, selectedRowIndex, COLUMN_UUID)

    If ValidationText.IsBlank(originalUUID) Then
        ErrorManager.ThrowError _
            sheetRecapRequest.EmptyOriginalRequestUUID, _
            MODULE_NAME & ".CreateRequestUpdateProposal", _
            "UUID transaksi request original " & _
            "tidak ditemukan."
    End If

    ' =====================================================
    ' Step 4: Read manufacturing plant
    ' =====================================================
    manufacturingPlantName = FormManager.GetCellText( _
        Me, CELL_LOCATION)

    If ValidationText.IsBlank(manufacturingPlantName) Then
        ErrorManager.ThrowError _
            sheetRecapRequest.EmptyUpdatePlantName, _
            MODULE_NAME & ".CreateRequestUpdateProposal", _
            "Nama manufacturing plant " & _
            "tidak boleh kosong."
    End If

    ' =====================================================
    ' Step 5: Read text values from selected row
    ' =====================================================
    mobilePostName = TableManager.GetText( _
        tbl, selectedRowIndex, COLUMN_MOBILE_POST)

    catalogueName = TableManager.GetText( _
        tbl, selectedRowIndex, COLUMN_CATALOGUE)

    unitOfMeasurement = TableManager.GetText( _
        tbl, selectedRowIndex, COLUMN_UOM)

    remarks = TableManager.GetText( _
        tbl, selectedRowIndex, COLUMN_REMARKS)

    personInCharge = TableManager.GetText( _
        tbl, selectedRowIndex, COLUMN_PIC)

    If ValidationText.IsBlank(mobilePostName) Then
        ErrorManager.ThrowError _
            sheetRecapRequest.EmptyMobilePostName, _
            MODULE_NAME & ".CreateRequestUpdateProposal", _
            "Nama mobile post tidak boleh kosong."
    End If

    If ValidationText.IsBlank(catalogueName) Then
        ErrorManager.ThrowError _
            sheetRecapRequest.EmptyCatalogueName, _
            MODULE_NAME & ".CreateRequestUpdateProposal", _
            "Nama barang tidak boleh kosong."
    End If

    If ValidationText.IsBlank(unitOfMeasurement) Then
        ErrorManager.ThrowError _
            sheetRecapRequest.EmptyUnitOfMeasurement, _
            MODULE_NAME & ".CreateRequestUpdateProposal", _
            "Satuan barang tidak boleh kosong."
    End If

    If ValidationText.IsBlank(personInCharge) Then
        ErrorManager.ThrowError _
            sheetRecapRequest.EmptyPersonInCharge, _
            MODULE_NAME & ".CreateRequestUpdateProposal", _
            "PIC request tidak boleh kosong."
    End If

    ' =====================================================
    ' Step 6: Read and validate request date
    ' =====================================================
    requestDateValue = TableManager.GetValue( _
        tbl, selectedRowIndex, COLUMN_ORDER_DATE)

    If ValidationText.IsBlank(requestDateValue) Then
        ErrorManager.ThrowError _
            sheetRecapRequest.EmptyRequestDate, _
            MODULE_NAME & ".CreateRequestUpdateProposal", _
            "Tanggal request tidak boleh kosong."
    End If

    If Not IsDate(requestDateValue) Then
        ErrorManager.ThrowError _
            sheetRecapRequest.InvalidRequestDate, _
            MODULE_NAME & ".CreateRequestUpdateProposal", _
            "Tanggal request tidak valid."
    End If

    requestDate = TableManager.GetDate( _
        tbl, selectedRowIndex, COLUMN_ORDER_DATE)

    ' =====================================================
    ' Step 7: Read and validate request amount
    ' =====================================================
    requestAmountValue = TableManager.GetValue( _
        tbl, selectedRowIndex, COLUMN_AMOUNT)

    If ValidationText.IsBlank(requestAmountValue) Then
        ErrorManager.ThrowError _
            sheetRecapRequest.EmptyRequestAmount, _
            MODULE_NAME & ".CreateRequestUpdateProposal", _
            "Jumlah request tidak boleh kosong."
    End If

    If ValidationNumeric.IsNotNumber(requestAmountValue) Then
        ErrorManager.ThrowError _
            sheetRecapRequest.InvalidRequestAmount, _
            MODULE_NAME & ".CreateRequestUpdateProposal", _
            "Jumlah request harus berupa angka."
    End If

    requestAmount = TableManager.GetDouble( _
        tbl, selectedRowIndex, COLUMN_AMOUNT)

    If ValidationNumeric.IsZeroOrNegative(requestAmount) Then
        ErrorManager.ThrowError _
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

    If Not FormManager.Confirm( _
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

    FormManager.SetStatus _
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

    If ValidationText.IsBlank(proposedUUID) Then
        ErrorManager.ThrowError _
            sheetRecapRequest.EmptyProposedRequestUUID, _
            MODULE_NAME & ".CreateRequestUpdateProposal", _
            "UUID transaksi request proposed " & _
            "tidak berhasil dibuat."
    End If

    If ValidationText.IsSameText( _
        originalUUID, _
        proposedUUID) Then

        ErrorManager.ThrowError _
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

    FormManager.ClearStatus

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

    FormManager.ShowInfo _
        successMessage, _
        "Permintaan Update Berhasil"

    ' =====================================================
    ' Step 17: Refresh request recap
    ' =====================================================
    SearchRequest_Click

CleanUp:
    FormManager.ClearStatus

    Set proposedRequest = Nothing
    Set proposedMobilePost = Nothing
    Set proposedPlant = Nothing
    Set proposedCatalogue = Nothing

    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

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

    FormManager.ClearStatus

    Set proposedRequest = Nothing
    Set proposedMobilePost = Nothing
    Set proposedPlant = Nothing
    Set proposedCatalogue = Nothing

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".CreateRequestUpdateProposal"

End Sub
