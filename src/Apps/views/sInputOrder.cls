Attribute VB_Name = "sInputOrder"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Const TABLE_NAME As String = "TableInputOrder"

Private Const CELL_LOCATION As String = "B4"
Private Const CELL_ORDER_DATE As String = "B6"

Private Const COLUMN_CATALOGUE As String = "Nama Barang"
Private Const COLUMN_AMOUNT As String = "Jumlah Barang"
Private Const COLUMN_VENDOR As String = "Vendor"
Private Const COLUMN_REMARKS As String = "Remarks"
Private Const COLUMN_PIC As String = "PIC"

Private Const CLASS_NAME As String = "Worksheet_INPUT_ORDER"

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    managerForm.HandleDatePickerSelection Target, Me.Range("B6")
End Sub

Private Sub Worksheet_Activate()

    Me.Range(CELL_LOCATION).Select

End Sub

Public Sub SaveOrder_Click()

    On Error GoTo ErrorHandler

    Dim plantName As String
    Dim orderDateValue As Date

    Dim tbl As ListObject
    Dim rowIndex As Long

    Dim plantItem As Plant
    Dim catalogueItem As Catalogue
    Dim vendorItem As Vendor
    Dim orderItem As Order

    Dim catalogueName As String
    Dim amountValue As Double
    Dim vendorName As String
    Dim remarksValue As String
    Dim picValue As String

    Dim savedCount As Long
    Dim currentRowIndex As Long
    Dim errInfo As ErrorInfo

    ' =====================================================
    ' Step 1: Read header values (Plant + Order Date)
    ' =====================================================
    plantName = managerForm.GetCellText(Me, CELL_LOCATION)
    orderDateValue = CDate(managerForm.GetCellValue(Me, CELL_ORDER_DATE))

    Set tbl = managerTable.GetTable(Me, TABLE_NAME)

    If Not managerTable.HasRows(tbl) Then
        managerForm.ShowWarning "Tidak ada baris data untuk disimpan."
        Exit Sub
    End If

    ' =====================================================
    ' Step 2: Loop every row and call CreateOrder
    ' =====================================================
    savedCount = 0

    For rowIndex = 1 To managerTable.RowCount(tbl)

        currentRowIndex = rowIndex

        catalogueName = managerTable.GetText( _
            tbl, rowIndex, COLUMN_CATALOGUE)

        amountValue = managerTable.GetDouble( _
            tbl, rowIndex, COLUMN_AMOUNT)

        vendorName = managerTable.GetText( _
            tbl, rowIndex, COLUMN_VENDOR)

        remarksValue = managerTable.GetText( _
            tbl, rowIndex, COLUMN_REMARKS)

        picValue = managerTable.GetText( _
            tbl, rowIndex, COLUMN_PIC)

        Set plantItem = New Plant
        plantItem.Name = plantName

        Set catalogueItem = New Catalogue
        catalogueItem.Name = catalogueName

        Set vendorItem = New Vendor
        vendorItem.Name = vendorName

        Set orderItem = New Order
        Set orderItem.Plant = plantItem
        orderItem.orderDate = orderDateValue
        Set orderItem.Catalogue = catalogueItem
        orderItem.Amount = amountValue
        Set orderItem.Vendor = vendorItem
        orderItem.remarks = remarksValue
        orderItem.personInCharge = picValue

        serviceOrder.CreateOrder orderItem:=orderItem

        savedCount = savedCount + 1

        Set orderItem = Nothing
        Set plantItem = Nothing
        Set catalogueItem = Nothing
        Set vendorItem = Nothing

    Next rowIndex

    managerForm.ShowInfo savedCount & " order berhasil disimpan."
    
    managerTable.ClearContents tbl

CleanUp:
    Set tbl = Nothing
    Set orderItem = Nothing
    Set plantItem = Nothing
    Set catalogueItem = Nothing
    Set vendorItem = Nothing
    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    Set tbl = Nothing
    Set orderItem = Nothing
    Set plantItem = Nothing
    Set catalogueItem = Nothing
    Set vendorItem = Nothing

    managerError.ShowError errInfo, _
        "Baris ke-" & currentRowIndex & " pada tabel."

End Sub

Public Sub AddOrderRow_Click()

    On Error GoTo ErrorHandler

    Dim tbl As ListObject
    Dim errInfo As ErrorInfo

    Set tbl = managerTable.GetTable(Me, TABLE_NAME)
    managerTable.AddRow tbl

    Set tbl = Nothing

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    Set tbl = Nothing

    managerError.ShowError errInfo

End Sub

Public Sub ClearEmptyOrderRows()

    On Error GoTo ErrorHandler

    Dim tbl As ListObject
    Dim errInfo As ErrorInfo

    Set tbl = managerTable.GetTable(Me, TABLE_NAME)

    managerTable.DeleteEmptyRows tbl

    Set tbl = Nothing

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    Set tbl = Nothing

    managerError.ShowError errInfo

End Sub
