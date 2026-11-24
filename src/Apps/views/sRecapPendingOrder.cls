Attribute VB_Name = "sRecapPendingOrder"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

' Error codes: sheetRecapPendingOrder (managerErrorCodesSheets.bas)

Private Const MODULE_NAME As String = "SHEET_RECAP_PENDING_ORDER"

Private Const TABLE_NAME As String = "TableRecapPendingOrder"

Private Const CELL_LOCATION As String = "B4"
Private Const CELL_FROM_DATE As String = "B7"
Private Const CELL_UNTIL_DATE As String = "B8"

Private Const COLUMN_NO As String = "No"
Private Const COLUMN_UUID As String = "UUID"
Private Const COLUMN_PO_ID As String = "PO ID"
Private Const COLUMN_CATALOGUE As String = "Nama Barang"
Private Const COLUMN_AMOUNT As String = "Jumlah Order"
Private Const COLUMN_UOM As String = "Satuan"
Private Const COLUMN_VENDOR As String = "Vendor"
Private Const COLUMN_ORDER_DATE As String = "Tanggal Order"
Private Const COLUMN_STATUS As String = "Status Kedatangan"
Private Const COLUMN_REMARKS As String = "Remarks"
Private Const COLUMN_PIC As String = "PIC"
Private Const COLUMN_UPDATE As String = "UPDATE"
Private Const COLUMN_DELETE As String = "DELETE"


Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    On Error GoTo ErrorHandler

    Dim tbl As ListObject
    Dim selectedRowIndex As Long

    Dim state As FormState
    Dim errInfo As ErrorInfo

    state = managerForm.CaptureState()

    '=====================================================
    'Date input cells
    '=====================================================
    If managerForm.IsInRange( _
        Target, _
        Me.Range(CELL_FROM_DATE & ":" & CELL_UNTIL_DATE) _
    ) Then

        Application.EnableEvents = False

        managerForm.ShowDatePicker Target

        GoTo CleanUp

    End If

    '=====================================================
    'Get the order recap table
    '=====================================================
    Set tbl = managerTable.GetTable(Me, TABLE_NAME)

    If Not managerTable.HasRows(tbl) Then
        GoTo CleanUp
    End If

    '=====================================================
    'UPDATE column clicked
    '=====================================================
    If managerTable.IsInColumn( _
        Target, _
        tbl, _
        COLUMN_UPDATE) Then

        Application.EnableEvents = False

        selectedRowIndex = _
            managerTable.GetRowIndex(tbl, Target)

        CreateOrderUpdateRequest _
            tbl:=tbl, _
            selectedRowIndex:=selectedRowIndex

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

    managerError.ShowError errInfo

End Sub


Private Sub CreateOrderUpdateRequest( _
    ByVal tbl As ListObject, _
    ByVal selectedRowIndex As Long)

    On Error GoTo ErrorHandler

    Dim originalUUID As String
    Dim plantName As String

    Dim proposedOrder As Order
    Dim proposedCatalogue As Catalogue
    Dim proposedPlant As Plant
    Dim proposedVendor As Vendor

    Dim errInfo As ErrorInfo

    '=====================================================
    'Validate selected row
    '=====================================================
    If Not managerTable.IsValidRow(tbl, selectedRowIndex) Then

        managerError.ThrowError _
            sheetRecapPendingOrder.InvalidSelectedOrderRow, _
            MODULE_NAME & ".CreateOrderUpdateRequest", _
            "Baris order yang dipilih tidak valid."

    End If

    '=====================================================
    'Get the original transaction UUID
    '=====================================================
    originalUUID = managerTable.GetText( _
        tbl, selectedRowIndex, COLUMN_UUID)

    If managerValidationText.IsBlank(originalUUID) Then

        managerError.ThrowError _
            sheetRecapPendingOrder.EmptyOriginalOrderUUID, _
            MODULE_NAME & ".CreateOrderUpdateRequest", _
            "UUID order original tidak ditemukan."

    End If

    '=====================================================
    'Validate PO ID
    '=====================================================
    If managerValidationText.IsBlank(managerTable.GetText( _
        tbl, selectedRowIndex, COLUMN_PO_ID)) Then

        managerError.ThrowError _
            sheetRecapPendingOrder.EmptyOrderPoId, _
            MODULE_NAME & ".CreateOrderUpdateRequest", _
            "PO ID order tidak boleh kosong."

    End If

    '=====================================================
    'Validate catalogue name
    '=====================================================
    If managerValidationText.IsBlank(managerTable.GetText( _
        tbl, selectedRowIndex, COLUMN_CATALOGUE)) Then

        managerError.ThrowError _
            sheetRecapPendingOrder.EmptyOrderCatalogueName, _
            MODULE_NAME & ".CreateOrderUpdateRequest", _
            "Nama barang tidak boleh kosong."

    End If

    '=====================================================
    'Validate order amount
    '=====================================================
    If managerValidationNumeric.IsNotNumber(managerTable.GetValue( _
        tbl, selectedRowIndex, COLUMN_AMOUNT)) Then

        managerError.ThrowError _
            sheetRecapPendingOrder.InvalidOrderAmount, _
            MODULE_NAME & ".CreateOrderUpdateRequest", _
            "Jumlah order harus berupa angka."

    End If

    If managerValidationNumeric.IsZeroOrNegative(managerTable.GetDouble( _
        tbl, selectedRowIndex, COLUMN_AMOUNT)) Then

        managerError.ThrowError _
            sheetRecapPendingOrder.ZeroOrNegativeOrderAmount, _
            MODULE_NAME & ".CreateOrderUpdateRequest", _
            "Jumlah order harus lebih besar dari nol."

    End If

    '=====================================================
    'Validate vendor
    '=====================================================
    If managerValidationText.IsBlank(managerTable.GetText( _
        tbl, selectedRowIndex, COLUMN_VENDOR)) Then

        managerError.ThrowError _
            sheetRecapPendingOrder.EmptyOrderVendorName, _
            MODULE_NAME & ".CreateOrderUpdateRequest", _
            "Vendor tidak boleh kosong."

    End If

    '=====================================================
    'Validate order date
    '=====================================================
    If Not IsDate(managerTable.GetValue( _
        tbl, selectedRowIndex, COLUMN_ORDER_DATE)) Then

        managerError.ThrowError _
            sheetRecapPendingOrder.InvalidOrderDate, _
            MODULE_NAME & ".CreateOrderUpdateRequest", _
            "Tanggal order tidak valid."

    End If

    '=====================================================
    'Validate manufacturing plant
    '=====================================================
    plantName = managerForm.GetCellText(Me, CELL_LOCATION)

    If managerValidationText.IsBlank(plantName) Then

        managerError.ThrowError _
            sheetRecapPendingOrder.EmptyUpdatePlantLocation, _
            MODULE_NAME & ".CreateOrderUpdateRequest", _
            "Lokasi manufacturing plant tidak ditemukan."

    End If

    '=====================================================
    'Validate person in charge
    '=====================================================
    If managerValidationText.IsBlank(managerTable.GetText( _
        tbl, selectedRowIndex, COLUMN_PIC)) Then

        managerError.ThrowError _
            sheetRecapPendingOrder.EmptyOrderPersonInCharge, _
            MODULE_NAME & ".CreateOrderUpdateRequest", _
            "PIC order tidak boleh kosong."

    End If

    '=====================================================
    'Build the proposed catalogue object
    '=====================================================
    Set proposedCatalogue = New Catalogue

    proposedCatalogue.Name = managerTable.GetText( _
        tbl, selectedRowIndex, COLUMN_CATALOGUE)

    proposedCatalogue.unitOfMeasurement = managerTable.GetText( _
        tbl, selectedRowIndex, COLUMN_UOM)

    '=====================================================
    'Build the proposed manufacturing plant object
    '=====================================================
    Set proposedPlant = New Plant

    proposedPlant.Name = plantName

    '=====================================================
    'Build the proposed vendor object
    '=====================================================
    Set proposedVendor = New Vendor

    proposedVendor.Name = managerTable.GetText( _
        tbl, selectedRowIndex, COLUMN_VENDOR)

    '=====================================================
    'Build the proposed order object
    '=====================================================
    Set proposedOrder = New Order

    With proposedOrder

        .poID = managerTable.GetText( _
            tbl, selectedRowIndex, COLUMN_PO_ID)

        .orderDate = managerTable.GetDate( _
            tbl, selectedRowIndex, COLUMN_ORDER_DATE)

        .Amount = managerTable.GetDouble( _
            tbl, selectedRowIndex, COLUMN_AMOUNT)

        .remarks = managerTable.GetText( _
            tbl, selectedRowIndex, COLUMN_REMARKS)

        .personInCharge = managerTable.GetText( _
            tbl, selectedRowIndex, COLUMN_PIC)

        .Status = managerTable.GetText( _
            tbl, selectedRowIndex, COLUMN_STATUS)

        Set .Catalogue = proposedCatalogue
        Set .Plant = proposedPlant
        Set .Vendor = proposedVendor

    End With

    '=====================================================
    'Show confirmation dialog
    '=====================================================
    If Not managerForm.Confirm( _
        "Apakah Anda yakin ingin membuat " & _
        "permintaan perubahan order berikut?" & _
        vbCrLf & vbCrLf & _
        "UUID original: " & originalUUID & vbCrLf & _
        "PO ID: " & proposedOrder.poID & vbCrLf & _
        "Nama barang: " & proposedOrder.Catalogue.Name & vbCrLf & _
        "Jumlah: " & Format$( _
            proposedOrder.Amount, _
            "#,##0.00" _
        ) & " " & _
        proposedOrder.Catalogue.unitOfMeasurement & vbCrLf & _
        "Vendor: " & proposedOrder.Vendor.Name & vbCrLf & _
        "Tanggal order: " & Format$( _
            proposedOrder.orderDate, _
            "dd-mmm-yyyy" _
        ), _
        "Konfirmasi Permintaan Update Order") Then

        GoTo CleanUp
    End If

    '=====================================================
    'Create proposed order
    '
    'Passing originalUUID makes CreateOrder:
    ' 1. Create an inactive proposed transaction
    ' 2. Create the proposed order
    ' 3. Create the transaction update request
    ' 4. Move related receive records
    '=====================================================
    serviceOrder.CreateOrder _
        orderItem:=proposedOrder, _
        originalUUID:=originalUUID

    '=====================================================
    'Show success message
    '=====================================================
    managerForm.ShowInfo _
        "Permintaan perubahan order berhasil dibuat." & _
        vbCrLf & vbCrLf & _
        "UUID original: " & originalUUID & vbCrLf & _
        "PO ID proposed: " & proposedOrder.poID & vbCrLf & _
        "Status: Menunggu persetujuan", _
        "Permintaan Update Berhasil"

CleanUp:
    Set proposedOrder = Nothing
    Set proposedCatalogue = Nothing
    Set proposedPlant = Nothing
    Set proposedVendor = Nothing

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    Set proposedOrder = Nothing
    Set proposedCatalogue = Nothing
    Set proposedPlant = Nothing
    Set proposedVendor = Nothing

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".CreateOrderUpdateRequest"

End Sub


Public Sub SearchOrder_Click()

    On Error GoTo ErrorHandler

    Dim plantName As String
    Dim fromDateValue As Date
    Dim untilDateValue As Date

    Dim orders As Collection
    Dim orderItem As Order

    Dim tbl As ListObject
    Dim rowIndex As Long
    Dim rowNumber As Long

    Dim errInfo As ErrorInfo

    '=====================================================
    'Step 1: Read and validate filter values
    '=====================================================
    plantName = managerForm.GetCellText(Me, CELL_LOCATION)

    If managerValidationText.IsBlank(plantName) Then

        managerError.ThrowError _
            sheetRecapPendingOrder.EmptyPlantLocation, _
            MODULE_NAME & ".SearchOrder_Click", _
            "Lokasi belum dipilih."

    End If

    If Not managerForm.IsCellDate(Me, CELL_FROM_DATE) Then

        managerError.ThrowError _
            sheetRecapPendingOrder.InvalidFromDate, _
            MODULE_NAME & ".SearchOrder_Click", _
            "Tanggal Dari tidak valid."

    End If

    If Not managerForm.IsCellDate(Me, CELL_UNTIL_DATE) Then

        managerError.ThrowError _
            sheetRecapPendingOrder.InvalidUntilDate, _
            MODULE_NAME & ".SearchOrder_Click", _
            "Tanggal Hingga tidak valid."

    End If

    fromDateValue = CDate( _
        managerForm.GetCellValue(Me, CELL_FROM_DATE) _
    )

    untilDateValue = CDate( _
        managerForm.GetCellValue(Me, CELL_UNTIL_DATE) _
    )

    If fromDateValue > untilDateValue Then

        managerError.ThrowError _
            sheetRecapPendingOrder.FromDateExceedsUntilDate, _
            MODULE_NAME & ".SearchOrder_Click", _
            "Tanggal Dari tidak boleh lebih besar " & _
            "dari Tanggal Hingga."

    End If

    '=====================================================
    'Step 2: Get the order recap table
    '=====================================================
    Set tbl = managerTable.GetTable(Me, TABLE_NAME)

    '=====================================================
    'Step 3: Select orders
    '=====================================================
    Set orders = serviceOrder.SelectPendingOrder( _
        manufacturingPlantName:=plantName, _
        fromDate:=fromDateValue, _
        untilDate:=untilDateValue)

    If orders Is Nothing Then

        managerError.ThrowError _
            sheetRecapPendingOrder.EmptyPendingOrderCollection, _
            MODULE_NAME & ".SearchOrder_Click", _
            "Hasil pencarian order tidak dapat dibaca."

    End If

    '=====================================================
    'Step 4: Clear existing table rows
    '=====================================================
    managerTable.ClearRows tbl

    '=====================================================
    'Step 5: Fill the table
    '=====================================================
    rowNumber = 0

    For Each orderItem In orders

        rowNumber = rowNumber + 1

        rowIndex = managerTable.AddRow(tbl)

        managerTable.SetValue _
            tbl, rowIndex, COLUMN_NO, rowNumber

        managerTable.SetValue _
            tbl, rowIndex, COLUMN_UUID, orderItem.UUID

        managerTable.SetValue _
            tbl, rowIndex, COLUMN_PO_ID, orderItem.poID

        managerTable.SetValue _
            tbl, rowIndex, COLUMN_CATALOGUE, orderItem.Catalogue.Name

        managerTable.SetValue _
            tbl, rowIndex, COLUMN_AMOUNT, orderItem.Amount

        managerTable.SetValue _
            tbl, rowIndex, COLUMN_UOM, orderItem.Catalogue.unitOfMeasurement

        managerTable.SetValue _
            tbl, rowIndex, COLUMN_VENDOR, orderItem.Vendor.Name

        managerTable.SetValue _
            tbl, rowIndex, COLUMN_ORDER_DATE, orderItem.orderDate

        managerTable.SetValue _
            tbl, rowIndex, COLUMN_STATUS, orderItem.Status

        managerTable.SetValue _
            tbl, rowIndex, COLUMN_REMARKS, orderItem.remarks

        managerTable.SetValue _
            tbl, rowIndex, COLUMN_PIC, orderItem.personInCharge

        managerTable.SetValue _
            tbl, rowIndex, COLUMN_UPDATE, "UPDATE"

    Next orderItem

    '=====================================================
    'Step 6: Apply table formatting
    '=====================================================
    ApplyOrderTableFormat tbl

CleanUp:
    Set tbl = Nothing
    Set orderItem = Nothing
    Set orders = Nothing

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    Set tbl = Nothing
    Set orderItem = Nothing
    Set orders = Nothing

    managerError.ShowError errInfo

End Sub


Private Sub ApplyOrderTableFormat( _
    ByVal tbl As ListObject)

    On Error GoTo ErrorHandler

    Dim columnRange As Range
    Dim errInfo As ErrorInfo

    If Not managerTable.HasRows(tbl) Then
        GoTo CleanUp
    End If

    '=====================================================
    'Order date format
    '=====================================================
    managerTable.FormatColumn _
        tbl, COLUMN_ORDER_DATE, "dd-mmm-yyyy"

    '=====================================================
    'Order amount format
    '=====================================================
    managerTable.FormatColumn _
        tbl, COLUMN_AMOUNT, "#,##0.00"

    '=====================================================
    'UPDATE column format
    '=====================================================
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

    '=====================================================
    'DELETE column format
    '=====================================================

CleanUp:
    Set columnRange = Nothing
    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    Set columnRange = Nothing

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ApplyOrderTableFormat"

End Sub
