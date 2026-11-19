Attribute VB_Name = "sRecapOrder"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Const MODULE_NAME As String = "SHEET_RECAP_ORDER"

Private Const TABLE_NAME As String = "TableRecapOrder"

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

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    On Error GoTo ErrorHandler

    Dim tbl As ListObject
    Dim selectedRowIndex As Long

    Dim errInfo As ErrorInfo

    '=====================================================
    'Date input cells
    '=====================================================
    If FormManager.IsInRange( _
        Target, _
        Me.Range(CELL_FROM_DATE & ":" & CELL_UNTIL_DATE)) Then

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

        UpdateOrderPOID _
            tbl:=tbl, _
            selectedRowIndex:=selectedRowIndex

        GoTo CleanUp

    End If

CleanUp:
    Application.EnableEvents = True

    Set tbl = Nothing

    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.ShowError errInfo

    Resume CleanUp

End Sub


Private Sub UpdateOrderPOID( _
    ByVal tbl As ListObject, _
    ByVal selectedRowIndex As Long)

    On Error GoTo ErrorHandler

    Dim orderUUID As String
    Dim orderPOID As String

    Dim errInfo As ErrorInfo

    '=====================================================
    'Validate selected row
    '=====================================================
    If Not TableManager.IsValidRow(tbl, selectedRowIndex) Then

        ErrorManager.ThrowError _
            sheetRecapOrder.InvalidSelectedRow, _
            MODULE_NAME & ".UpdateOrderPOID", _
            "Baris order yang dipilih tidak valid."

    End If

    '=====================================================
    'Get order UUID
    '=====================================================
    orderUUID = TableManager.GetText( _
        tbl, selectedRowIndex, COLUMN_UUID)

    If ValidationText.IsBlank(orderUUID) Then

        ErrorManager.ThrowError _
            sheetRecapOrder.EmptyOrderUUID, _
            MODULE_NAME & ".UpdateOrderPOID", _
            "UUID order tidak ditemukan."

    End If

    '=====================================================
    'Get the new PO ID
    '=====================================================
    orderPOID = TableManager.GetText( _
        tbl, selectedRowIndex, COLUMN_PO_ID)

    If ValidationText.IsBlank(orderPOID) Then

        ErrorManager.ThrowError _
            sheetRecapOrder.EmptyPoId, _
            MODULE_NAME & ".UpdateOrderPOID", _
            "PO ID order tidak boleh kosong."

    End If

    '=====================================================
    'Show confirmation dialog
    '=====================================================
    If Not FormManager.Confirm( _
        "Apakah Anda yakin ingin memperbarui " & _
        "PO ID order berikut?" & _
        vbCrLf & vbCrLf & _
        "UUID: " & orderUUID & _
        vbCrLf & _
        "PO ID baru: " & orderPOID, _
        "Konfirmasi Update PO ID") Then

        Exit Sub
    End If

    '=====================================================
    'Update PO ID in the database
    '=====================================================
    serviceOrder.UpdateOrderPOID _
        orderUUID, _
        orderPOID

    '=====================================================
    'Show success message
    '=====================================================
    FormManager.ShowInfo _
        "PO ID berhasil diperbarui." & _
        vbCrLf & vbCrLf & _
        "UUID: " & orderUUID & _
        vbCrLf & _
        "PO ID: " & orderPOID, _
        "Update Berhasil"

    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".UpdateOrderPOID"

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
            sheetRecapOrder.EmptyPlantName, _
            MODULE_NAME & ".SearchOrder_Click", _
            "Lokasi belum dipilih."

    End If

    If Not FormManager.IsCellDate(Me, CELL_FROM_DATE) Then

        ErrorManager.ThrowError _
            sheetRecapOrder.InvalidFromDate, _
            MODULE_NAME & ".SearchOrder_Click", _
            "Tanggal Dari tidak valid."

    End If

    If Not FormManager.IsCellDate(Me, CELL_UNTIL_DATE) Then

        ErrorManager.ThrowError _
            sheetRecapOrder.InvalidUntilDate, _
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
            sheetRecapOrder.InvalidDateRange, _
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
    Set orders = serviceOrder.SelectOrder( _
        manufacturingPlantName:=plantName, _
        fromDate:=fromDateValue, _
        untilDate:=untilDateValue)

    If orders Is Nothing Then

        ErrorManager.ThrowError _
            sheetRecapOrder.EmptyOrderCollection, _
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

    ErrorManager.ShowError errInfo

    Resume CleanUp

End Sub

Private Sub ApplyOrderTableFormat( _
    ByVal tbl As ListObject)

    On Error GoTo ErrorHandler

    Dim updateRange As Range

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
    Set updateRange = TableManager.GetColumnRange( _
        tbl, COLUMN_UPDATE)

    If Not updateRange Is Nothing Then

        With updateRange

            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter

            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)

            .Interior.Pattern = xlSolid
            .Interior.Color = RGB(0, 102, 204)
        End With

    End If

CleanUp:
    Set updateRange = Nothing
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    Set updateRange = Nothing

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".ApplyOrderTableFormat"

End Sub
