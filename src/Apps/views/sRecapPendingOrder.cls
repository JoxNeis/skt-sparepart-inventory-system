Attribute VB_Name = "sRecapPendingOrder"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

' Error codes: sheetRecapPendingOrder (ErrorCodesSheets.bas)

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

    state = FormManager.CaptureState()

    '=====================================================
    'Date input cells
    '=====================================================
    If FormManager.IsInRange( _
        Target, _
        Me.Range(CELL_FROM_DATE & ":" & CELL_UNTIL_DATE) _
    ) Then

        Application.EnableEvents = False

        FormManager.ShowDatePicker Target

        GoTo CleanUp

    End If

    '=====================================================
    'Get the order recap table
    '=====================================================
    Set tbl = TableManager.GetTable(Me, TABLE_NAME)

    If Not TableManager.HasRows(tbl) Then
        GoTo CleanUp
    End If

    '=====================================================
    'UPDATE column clicked
    '=====================================================
    If TableManager.IsInColumn( _
        Target, _
        tbl, _
        COLUMN_UPDATE) Then

        Application.EnableEvents = False

        selectedRowIndex = _
            TableManager.GetRowIndex(tbl, Target)

        CreateOrderUpdateRequest _
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

    ErrorManager.ShowError errInfo

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
    If Not TableManager.IsValidRow(tbl, selectedRowIndex) Then

        ErrorManager.ThrowError _
            sheetRecapPendingOrder.InvalidSelectedOrderRow, _
            MODULE_NAME & ".CreateOrderUpdateRequest", _
            "Baris order yang dipilih tidak valid."

    End If

    '=====================================================
    'Get the original transaction UUID
    '=====================================================
    originalUUID = TableManager.GetText( _
        tbl, selectedRowIndex, COLUMN_UUID)

    If ValidationText.IsBlank(originalUUID) Then

        ErrorManager.ThrowError _
            sheetRecapPendingOrder.EmptyOriginalOrderUUID, _
            MODULE_NAME & ".CreateOrderUpdateRequest", _
            "UUID order original tidak ditemukan."

    End If

    '=====================================================
    'Validate PO ID
    '=====================================================
    If ValidationText.IsBlank(TableManager.GetText( _
        tbl, selectedRowIndex, COLUMN_PO_ID)) Then

        ErrorManager.ThrowError _
            sheetRecapPendingOrder.EmptyOrderPoId, _
            MODULE_NAME & ".CreateOrderUpdateRequest", _
            "PO ID order tidak boleh kosong."

    End If

    '=====================================================
    'Validate catalogue name
    '=====================================================
    If ValidationText.IsBlank(TableManager.GetText( _
        tbl, selectedRowIndex, COLUMN_CATALOGUE)) Then

        ErrorManager.ThrowError _
            sheetRecapPendingOrder.EmptyOrderCatalogueName, _
            MODULE_NAME & ".CreateOrderUpdateRequest", _
            "Nama barang tidak boleh kosong."

    End If

    '=====================================================
    'Validate order amount
    '=====================================================
    If ValidationNumeric.IsNotNumber(TableManager.GetValue( _
        tbl, selectedRowIndex, COLUMN_AMOUNT)) Then

        ErrorManager.ThrowError _
            sheetRecapPendingOrder.InvalidOrderAmount, _
            MODULE_NAME & ".CreateOrderUpdateRequest", _
            "Jumlah order harus berupa angka."

    End If

    If ValidationNumeric.IsZeroOrNegative(TableManager.GetDouble( _
        tbl, selectedRowIndex, COLUMN_AMOUNT)) Then

        ErrorManager.ThrowError _
            sheetRecapPendingOrder.ZeroOrNegativeOrderAmount, _
            MODULE_NAME & ".CreateOrderUpdateRequest", _
            "Jumlah order harus lebih besar dari nol."

    End If

    '=====================================================
    'Validate vendor
    '=====================================================
    If ValidationText.IsBlank(TableManager.GetText( _
        tbl, selectedRowIndex, COLUMN_VENDOR)) Then

        ErrorManager.ThrowError _
            sheetRecapPendingOrder.EmptyOrderVendorName, _
            MODULE_NAME & ".CreateOrderUpdateRequest", _
            "Vendor tidak boleh kosong."

    End If

    '=====================================================
    'Validate order date
    '=====================================================
    If Not IsDate(TableManager.GetValue( _
        tbl, selectedRowIndex, COLUMN_ORDER_DATE)) Then

        ErrorManager.ThrowError _
            sheetRecapPendingOrder.InvalidOrderDate, _
            MODULE_NAME & ".CreateOrderUpdateRequest", _
            "Tanggal order tidak valid."

    End If

    '=====================================================
    'Validate manufacturing plant
    '=====================================================
    plantName = FormManager.GetCellText(Me, CELL_LOCATION)

    If ValidationText.IsBlank(plantName) Then

        ErrorManager.ThrowError _
            sheetRecapPendingOrder.EmptyUpdatePlantLocation, _
            MODULE_NAME & ".CreateOrderUpdateRequest", _
            "Lokasi manufacturing plant tidak ditemukan."

    End If

    '=====================================================
    'Validate person in charge
    '=====================================================
    If ValidationText.IsBlank(TableManager.GetText( _
        tbl, selectedRowIndex, COLUMN_PIC)) Then

        ErrorManager.ThrowError _
            sheetRecapPendingOrder.EmptyOrderPersonInCharge, _
            MODULE_NAME & ".CreateOrderUpdateRequest", _
            "PIC order tidak boleh kosong."

    End If

    '=====================================================
    'Build the proposed catalogue object
    '=====================================================
    Set proposedCatalogue = New Catalogue

    proposedCatalogue.Name = TableManager.GetText( _
        tbl, selectedRowIndex, COLUMN_CATALOGUE)

    proposedCatalogue.unitOfMeasurement = TableManager.GetText( _
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

    proposedVendor.Name = TableManager.GetText( _
        tbl, selectedRowIndex, COLUMN_VENDOR)

    '=====================================================
    'Build the proposed order object
    '=====================================================
    Set proposedOrder = New Order

    With proposedOrder

        .poID = TableManager.GetText( _
            tbl, selectedRowIndex, COLUMN_PO_ID)

        .orderDate = TableManager.GetDate( _
            tbl, selectedRowIndex, COLUMN_ORDER_DATE)

        .Amount = TableManager.GetDouble( _
            tbl, selectedRowIndex, COLUMN_AMOUNT)

        .remarks = TableManager.GetText( _
            tbl, selectedRowIndex, COLUMN_REMARKS)

        .personInCharge = TableManager.GetText( _
            tbl, selectedRowIndex, COLUMN_PIC)

        .Status = TableManager.GetText( _
            tbl, selectedRowIndex, COLUMN_STATUS)

        Set .Catalogue = proposedCatalogue
        Set .Plant = proposedPlant
        Set .Vendor = proposedVendor

    End With

    '=====================================================
    'Show confirmation dialog
    '=====================================================
    If Not FormManager.Confirm( _
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
    FormManager.ShowInfo _
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
    errInfo = ErrorManager.CaptureError()

    Set proposedOrder = Nothing
    Set proposedCatalogue = Nothing
    Set proposedPlant = Nothing
    Set proposedVendor = Nothing

    ErrorManager.Rethrow errInfo, _
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
    plantName = FormManager.GetCellText(Me, CELL_LOCATION)

    If ValidationText.IsBlank(plantName) Then

        ErrorManager.ThrowError _
            sheetRecapPendingOrder.EmptyPlantLocation, _
            MODULE_NAME & ".SearchOrder_Click", _
            "Lokasi belum dipilih."

    End If

    If Not FormManager.IsCellDate(Me, CELL_FROM_DATE) Then

        ErrorManager.ThrowError _
            sheetRecapPendingOrder.InvalidFromDate, _
            MODULE_NAME & ".SearchOrder_Click", _
            "Tanggal Dari tidak valid."

    End If

    If Not FormManager.IsCellDate(Me, CELL_UNTIL_DATE) Then

        ErrorManager.ThrowError _
            sheetRecapPendingOrder.InvalidUntilDate, _
            MODULE_NAME & ".SearchOrder_Click", _
            "Tanggal Hingga tidak valid."

    End If

    fromDateValue = CDate( _
        FormManager.GetCellValue(Me, CELL_FROM_DATE) _
    )

    untilDateValue = CDate( _
        FormManager.GetCellValue(Me, CELL_UNTIL_DATE) _
    )

    If fromDateValue > untilDateValue Then

        ErrorManager.ThrowError _
            sheetRecapPendingOrder.FromDateExceedsUntilDate, _
            MODULE_NAME & ".SearchOrder_Click", _
            "Tanggal Dari tidak boleh lebih besar " & _
            "dari Tanggal Hingga."

    End If

    '=====================================================
    'Step 2: Get the order recap table
    '=====================================================
    Set tbl = TableManager.GetTable(Me, TABLE_NAME)

    '=====================================================
    'Step 3: Select orders
    '=====================================================
    Set orders = serviceOrder.SelectPendingOrder( _
        manufacturingPlantName:=plantName, _
        fromDate:=fromDateValue, _
        untilDate:=untilDateValue)

    If orders Is Nothing Then

        ErrorManager.ThrowError _
            sheetRecapPendingOrder.EmptyPendingOrderCollection, _
            MODULE_NAME & ".SearchOrder_Click", _
            "Hasil pencarian order tidak dapat dibaca."

    End If

    '=====================================================
    'Step 4: Clear existing table rows
    '=====================================================
    TableManager.ClearRows tbl

    '=====================================================
    'Step 5: Fill the table
    '=====================================================
    rowNumber = 0

    For Each orderItem In orders

        rowNumber = rowNumber + 1

        rowIndex = TableManager.AddRow(tbl)

        TableManager.SetValue _
            tbl, rowIndex, COLUMN_NO, rowNumber

        TableManager.SetValue _
            tbl, rowIndex, COLUMN_UUID, orderItem.UUID

        TableManager.SetValue _
            tbl, rowIndex, COLUMN_PO_ID, orderItem.poID

        TableManager.SetValue _
            tbl, rowIndex, COLUMN_CATALOGUE, orderItem.Catalogue.Name

        TableManager.SetValue _
            tbl, rowIndex, COLUMN_AMOUNT, orderItem.Amount

        TableManager.SetValue _
            tbl, rowIndex, COLUMN_UOM, orderItem.Catalogue.unitOfMeasurement

        TableManager.SetValue _
            tbl, rowIndex, COLUMN_VENDOR, orderItem.Vendor.Name

        TableManager.SetValue _
            tbl, rowIndex, COLUMN_ORDER_DATE, orderItem.orderDate

        TableManager.SetValue _
            tbl, rowIndex, COLUMN_STATUS, orderItem.Status

        TableManager.SetValue _
            tbl, rowIndex, COLUMN_REMARKS, orderItem.remarks

        TableManager.SetValue _
            tbl, rowIndex, COLUMN_PIC, orderItem.personInCharge

        TableManager.SetValue _
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
    errInfo = ErrorManager.CaptureError()

    Set tbl = Nothing
    Set orderItem = Nothing
    Set orders = Nothing

    ErrorManager.ShowError errInfo

End Sub


Private Sub ApplyOrderTableFormat( _
    ByVal tbl As ListObject)

    On Error GoTo ErrorHandler

    Dim columnRange As Range
    Dim errInfo As ErrorInfo

    If Not TableManager.HasRows(tbl) Then
        GoTo CleanUp
    End If

    '=====================================================
    'Order date format
    '=====================================================
    TableManager.FormatColumn _
        tbl, COLUMN_ORDER_DATE, "dd-mmm-yyyy"

    '=====================================================
    'Order amount format
    '=====================================================
    TableManager.FormatColumn _
        tbl, COLUMN_AMOUNT, "#,##0.00"

    '=====================================================
    'UPDATE column format
    '=====================================================
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

    '=====================================================
    'DELETE column format
    '=====================================================

CleanUp:
    Set columnRange = Nothing
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    Set columnRange = Nothing

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".ApplyOrderTableFormat"

End Sub
