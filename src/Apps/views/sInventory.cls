Attribute VB_Name = "sInventory"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

' Error codes: sheetInventory (managerErrorCodesSheets.bas)

Private Const MODULE_NAME As String = _
    "SHEET_INVENTORY"

Private Const TABLE_NAME As String = _
    "TableInventory"

Private Const CELL_MANUFACTURING_PLANT As String = _
    "B4"

' =========================================================
' Table column names
'
' These names must exactly match the Excel table headers.
' =========================================================
Private Const COLUMN_NO As String = _
    "No"

Private Const COLUMN_CATALOGUE As String = _
    "Nama Barang"

Private Const COLUMN_UOM As String = _
    "Satuan"

Private Const COLUMN_CURRENT_STOCK As String = _
    "Stock Sekarang"

Private Const COLUMN_REORDER_POINT As String = _
    "Reorder Point"

Private Const COLUMN_SAFETY_STOCK As String = _
    "Safety Stock"

Private Const COLUMN_LEADTIME As String = _
    "Lead Time"

Private Const COLUMN_AVERAGE_DAILY_USAGE As String = _
    "Average Daily Usage"

Private Const COLUMN_SD_DEMAND As String = _
    "SD Demand"

Private Const COLUMN_STATUS As String = _
    "Status"

Private Const COLUMN_UPDATE As String = _
    "UPDATE"
Private Sub Worksheet_SelectionChange( _
    ByVal Target As Range)

    On Error GoTo ErrorHandler

    Dim tbl As ListObject
    Dim selectedRowIndex As Long

    Dim state As FormState
    Dim errInfo As ErrorInfo

    state = managerForm.CaptureState()

    ' Only process one selected cell.
    If Target.CountLarge <> 1 Then
        Exit Sub
    End If

    ' =====================================================
    ' Get TableInventory
    ' =====================================================
    Set tbl = managerTable.GetTable(Me, TABLE_NAME)

    If Not managerTable.HasRows(tbl) Then
        GoTo CleanUp
    End If

    ' =====================================================
    ' Detect UPDATE column click
    ' =====================================================
    If Not managerTable.IsInColumn( _
        Target, _
        tbl, _
        COLUMN_UPDATE) Then

        GoTo CleanUp
    End If

    If Not managerValidationText.IsSameText( _
        TrimText(CStr(Target.value)), _
        "UPDATE") Then

        GoTo CleanUp
    End If

    Application.EnableEvents = False

    selectedRowIndex = _
        managerTable.GetRowIndex(tbl, Target)

    HandleUpdateROP _
        tbl:=tbl, _
        selectedRowIndex:=selectedRowIndex, _
        actionCell:=Target

CleanUp:
    Set tbl = Nothing

    managerForm.EndBusy state

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    Set tbl = Nothing

    managerForm.EndBusy state
    managerForm.ClearStatus

    managerError.ShowError errInfo, _
        "Gagal memproses update reorder point."

End Sub


' =========================================================
' CARI button
' =========================================================
Public Sub SearchInventory_Click()

    On Error GoTo ErrorHandler

    Dim tbl As ListObject
    Dim rowIndex As Long

    Dim rs As ADODB.recordset

    Dim manufacturingPlantName As String
    Dim inventoryStatus As String

    Dim currentStock As Double
    Dim reorderPoint As Double
    Dim safetyStock As Double
    Dim averageDailyUsage As Double
    Dim sdDemand As Double

    Dim rowNumber As Long

    Dim state As FormState
    Dim errInfo As ErrorInfo

    state = managerForm.BeginBusy( _
        "Mengambil data stock inventory...")

    ' =====================================================
    ' Step 1: Read and validate manufacturing plant
    ' =====================================================
    manufacturingPlantName = managerForm.GetCellText( _
        Me, CELL_MANUFACTURING_PLANT)

    If managerValidationText.IsBlank(manufacturingPlantName) Then
        managerError.ThrowError _
            sheetInventory.EmptyPlantName, _
            MODULE_NAME & ".SearchInventory_Click", _
            "Nama manufacturing plant " & _
            "tidak boleh kosong."
    End If

    ' =====================================================
    ' Step 2: Get and validate TableInventory
    ' =====================================================
    Set tbl = managerTable.GetTable(Me, TABLE_NAME)

    ValidateInventoryTableColumns tbl

    ' =====================================================
    ' Step 3: Call service function
    ' =====================================================
    Set rs = _
        servicePlantStock.SelectManufacturingPlantStocks( _
            manufacturingPlantName:= _
                manufacturingPlantName)

    If rs Is Nothing Then
        managerError.ThrowError _
            sheetInventory.EmptyInventoryRecordset, _
            MODULE_NAME & ".SearchInventory_Click", _
            "DAO tidak mengembalikan recordset."
    End If

    If rs.State <> adStateOpen Then
        managerError.ThrowError _
            managerDatabase.RecordsetNotOpen, _
            MODULE_NAME & ".SearchInventory_Click", _
            "Recordset stock inventory " & _
            "tidak terbuka."
    End If

    ' =====================================================
    ' Step 4: Clear previous table records
    ' =====================================================
    ClearInventoryTable tbl

    ' =====================================================
    ' Step 5: Fill TableInventory
    ' =====================================================
    rowNumber = 0

    Do While Not rs.EOF

        rowNumber = rowNumber + 1

        currentStock = managerDatabase.FieldDouble( _
            rs, "current_stock")

        reorderPoint = managerDatabase.FieldDouble( _
            rs, "reorder_point")

        safetyStock = managerDatabase.FieldDouble( _
            rs, "safety_stock")

        averageDailyUsage = managerDatabase.FieldDouble( _
            rs, "average_daily_usage")

        sdDemand = managerDatabase.FieldDouble( _
            rs, "sd_demand")

        inventoryStatus = DetermineInventoryStatus( _
            currentStock:=currentStock, _
            reorderPoint:=reorderPoint, _
            safetyStock:=safetyStock)

        rowIndex = managerTable.AddRow(tbl)

        ' No
        managerTable.SetValue _
            tbl, rowIndex, COLUMN_NO, rowNumber

        ' Nama Barang
        managerTable.SetValue _
            tbl, rowIndex, COLUMN_CATALOGUE, _
            managerDatabase.FieldText(rs, "catalogue_name")

        ' Satuan
        managerTable.SetValue _
            tbl, rowIndex, COLUMN_UOM, _
            managerDatabase.FieldText( _
                rs, "catalogue_unit_of_measurement")

        ' Stock Sekarang
        managerTable.SetValue _
            tbl, rowIndex, COLUMN_CURRENT_STOCK, _
            currentStock

        ' Reorder Point
        managerTable.SetValue _
            tbl, rowIndex, COLUMN_REORDER_POINT, _
            reorderPoint

        ' Safety Stock
        managerTable.SetValue _
            tbl, rowIndex, COLUMN_SAFETY_STOCK, _
            safetyStock

        ' Lead Time
        managerTable.SetValue _
            tbl, rowIndex, COLUMN_LEADTIME, _
            managerDatabase.FieldText(rs, "lead_time")

        ' Average Daily Usage
        managerTable.SetValue _
            tbl, rowIndex, COLUMN_AVERAGE_DAILY_USAGE, _
            averageDailyUsage

        ' SD Demand
        managerTable.SetValue _
            tbl, rowIndex, COLUMN_SD_DEMAND, _
            sdDemand

        ' Status
        managerTable.SetValue _
            tbl, rowIndex, COLUMN_STATUS, _
            inventoryStatus

        ' Action
        managerTable.SetValue _
            tbl, rowIndex, COLUMN_UPDATE, "UPDATE"

        rs.MoveNext
    Loop

    ' =====================================================
    ' Step 6: Apply formatting
    ' =====================================================
    ApplyInventoryTableFormatting tbl

    ' =====================================================
    ' Step 7: Show search result
    ' =====================================================
    If rowNumber = 0 Then

        managerForm.ShowInfo _
            "Tidak ada data stock inventory untuk " & _
            "manufacturing plant '" & _
            manufacturingPlantName & "'.", _
            "Stock Inventory"

    Else

        managerForm.ShowInfo _
            CStr(rowNumber) & _
            " data stock inventory ditemukan.", _
            "Stock Inventory"

    End If

CleanUp:
    managerDatabase.CloseRecordset rs

    Set tbl = Nothing
    Set rs = Nothing

    managerForm.EndBusy state

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerDatabase.CloseRecordset rs

    Set tbl = Nothing
    Set rs = Nothing

    managerForm.EndBusy state

    managerError.ShowError errInfo, _
        "Gagal menampilkan data stock inventory."

End Sub


' =========================================================
' Clear previous inventory records
' =========================================================
Private Sub ClearInventoryTable( _
    ByVal tbl As ListObject)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    If tbl Is Nothing Then
        managerError.ThrowError _
            managerTable.TableNotFound, _
            MODULE_NAME & ".ClearInventoryTable", _
            "TableInventory belum diinisialisasi."
    End If

    managerTable.ClearRows tbl

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ClearInventoryTable", _
        "Gagal membersihkan data stock inventory."

End Sub


' =========================================================
' Determine inventory status
'
' KRITIS:
'   Current stock is at or below safety stock.
'
' PERLU REORDER:
'   Current stock is above safety stock but at or below
'   reorder point.
'
' AMAN:
'   Current stock is above reorder point.
' =========================================================
Private Function DetermineInventoryStatus( _
    ByVal currentStock As Double, _
    ByVal reorderPoint As Double, _
    ByVal safetyStock As Double) As String

    If currentStock <= safetyStock Then

        DetermineInventoryStatus = _
            "KRITIS"

    ElseIf currentStock <= reorderPoint Then

        DetermineInventoryStatus = _
            "PERLU REORDER"

    Else

        DetermineInventoryStatus = _
            "AMAN"

    End If

End Function


' =========================================================
' Apply table formatting
' =========================================================
Private Sub ApplyInventoryTableFormatting( _
    ByVal tbl As ListObject)

    On Error GoTo ErrorHandler

    Dim rowIndex As Long
    Dim statusCell As Range
    Dim columnRange As Range

    Dim statusValue As String

    Dim amountColumns As Variant
    Dim allColumns As Variant
    Dim columnName As Variant

    Dim errInfo As ErrorInfo

    If tbl Is Nothing Then
        managerError.ThrowError _
            managerTable.TableNotFound, _
            MODULE_NAME & ".ApplyInventoryTableFormatting", _
            "TableInventory belum diinisialisasi."
    End If

    If Not managerTable.HasRows(tbl) Then
        GoTo CleanUp
    End If

    amountColumns = Array( _
        COLUMN_CURRENT_STOCK, _
        COLUMN_REORDER_POINT, _
        COLUMN_SAFETY_STOCK, _
        COLUMN_AVERAGE_DAILY_USAGE, _
        COLUMN_SD_DEMAND)

    allColumns = Array( _
        COLUMN_NO, _
        COLUMN_CATALOGUE, _
        COLUMN_UOM, _
        COLUMN_CURRENT_STOCK, _
        COLUMN_REORDER_POINT, _
        COLUMN_SAFETY_STOCK, _
        COLUMN_LEADTIME, _
        COLUMN_AVERAGE_DAILY_USAGE, _
        COLUMN_SD_DEMAND, _
        COLUMN_STATUS, _
        COLUMN_UPDATE)

    ' =====================================================
    ' Numeric formats
    ' =====================================================
    For Each columnName In amountColumns

        managerTable.FormatColumn _
            tbl, CStr(columnName), "#,##0.00"

    Next columnName

    ' =====================================================
    ' General alignment
    ' =====================================================
    For Each columnName In allColumns

        Set columnRange = managerTable.GetColumnRange( _
            tbl, CStr(columnName))

        If Not columnRange Is Nothing Then
            columnRange.VerticalAlignment = xlCenter
        End If

    Next columnName

    managerTable.FormatColumn _
        tbl, COLUMN_NO, vbNullString, xlCenter

    Set columnRange = managerTable.GetColumnRange( _
        tbl, COLUMN_STATUS)

    If Not columnRange Is Nothing Then

        With columnRange

            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
        End With

    End If

    ' =====================================================
    ' UPDATE action formatting
    ' =====================================================
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

    ' =====================================================
    ' Apply status colors
    ' =====================================================
    For rowIndex = 1 To managerTable.RowCount(tbl)

        Set statusCell = managerTable.GetCell( _
            tbl, rowIndex, COLUMN_STATUS)

        statusValue = managerTable.GetText( _
            tbl, rowIndex, COLUMN_STATUS)

        With statusCell
            .Interior.Pattern = xlSolid

            Select Case UCase$(statusValue)

                Case "AMAN"
                    .Font.Color = RGB(0, 97, 0)
                    .Interior.Color = _
                        RGB(198, 239, 206)

                Case "PERLU REORDER"
                    .Font.Color = RGB(156, 101, 0)
                    .Interior.Color = _
                        RGB(255, 235, 156)

                Case "KRITIS"
                    .Font.Color = RGB(156, 0, 6)
                    .Interior.Color = _
                        RGB(255, 199, 206)

                Case Else
                    .Font.ColorIndex = _
                        xlAutomatic
                    .Interior.Pattern = _
                        xlNone

            End Select
        End With

        Set statusCell = Nothing

    Next rowIndex

CleanUp:
    Set statusCell = Nothing
    Set columnRange = Nothing

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    Set statusCell = Nothing
    Set columnRange = Nothing

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ApplyInventoryTableFormatting", _
        "Gagal menerapkan format tabel inventory."

End Sub


' =========================================================
' Validate TableInventory columns
' =========================================================
Private Sub ValidateInventoryTableColumns( _
    ByVal tbl As ListObject)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    If tbl Is Nothing Then
        managerError.ThrowError _
            managerTable.TableNotFound, _
            MODULE_NAME & ".ValidateInventoryTableColumns", _
            "TableInventory tidak ditemukan."
    End If

    managerTable.ValidateColumns tbl, Array( _
        COLUMN_NO, _
        COLUMN_CATALOGUE, _
        COLUMN_UOM, _
        COLUMN_CURRENT_STOCK, _
        COLUMN_REORDER_POINT, _
        COLUMN_SAFETY_STOCK, _
        COLUMN_AVERAGE_DAILY_USAGE, _
        COLUMN_SD_DEMAND, _
        COLUMN_STATUS, _
        COLUMN_UPDATE)

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ValidateInventoryTableColumns", _
        "Struktur TableInventory tidak sesuai."

End Sub

Private Sub HandleUpdateROP( _
    ByVal tbl As ListObject, _
    ByVal selectedRowIndex As Long, _
    ByVal actionCell As Range)

    On Error GoTo ErrorHandler

    Dim manufacturingPlantName As String
    Dim catalogueName As String
    Dim unitOfMeasurement As String

    Dim averageDailyUsageValue As Variant
    Dim sdDemandValue As Variant
    Dim leadTimeValue As Variant

    Dim averageDailyUsage As Double
    Dim sdDemand As Double
    Dim leadTime As Double

    Dim calculatedSafetyStock As Double
    Dim calculatedReorderPoint As Double

    Dim confirmationMessage As String
    Dim successMessage As String

    Dim actionCellChanged As Boolean

    Dim errInfo As ErrorInfo

    actionCellChanged = False

    ' =====================================================
    ' Step 1: Validate table and selected row
    ' =====================================================
    If tbl Is Nothing Then
        managerError.ThrowError _
            managerTable.TableNotFound, _
            MODULE_NAME & ".HandleUpdateROP", _
            "TableInventory belum diinisialisasi."
    End If

    If Not managerTable.HasRows(tbl) Then
        managerError.ThrowError _
            sheetInventory.EmptyInventoryTableData, _
            MODULE_NAME & ".HandleUpdateROP", _
            "TableInventory tidak memiliki data."
    End If

    If Not managerTable.IsValidRow(tbl, selectedRowIndex) Then

        managerError.ThrowError _
            managerTable.InvalidRow, _
            MODULE_NAME & ".HandleUpdateROP", _
            "Baris inventory yang dipilih tidak valid."
    End If

    If actionCell Is Nothing Then
        managerError.ThrowError _
            sheetInventory.EmptyUpdateActionCell, _
            MODULE_NAME & ".HandleUpdateROP", _
            "Cell UPDATE tidak ditemukan."
    End If

    If Not managerValidationText.IsSameText( _
        TrimText(CStr(actionCell.value)), _
        "UPDATE") Then

        GoTo CleanUp
    End If

    ' =====================================================
    ' Step 2: Read manufacturing plant
    ' =====================================================
    manufacturingPlantName = managerForm.GetCellText( _
        Me, CELL_MANUFACTURING_PLANT)

    If managerValidationText.IsBlank(manufacturingPlantName) Then
        managerError.ThrowError _
            sheetInventory.EmptyPlantName, _
            MODULE_NAME & ".HandleUpdateROP", _
            "Nama manufacturing plant " & _
            "tidak boleh kosong."
    End If

    ' =====================================================
    ' Step 3: Read catalogue information
    ' =====================================================
    catalogueName = managerTable.GetText( _
        tbl, selectedRowIndex, COLUMN_CATALOGUE)

    unitOfMeasurement = managerTable.GetText( _
        tbl, selectedRowIndex, COLUMN_UOM)

    If managerValidationText.IsBlank(catalogueName) Then
        managerError.ThrowError _
            sheetInventory.EmptyCatalogueName, _
            MODULE_NAME & ".HandleUpdateROP", _
            "Nama barang tidak boleh kosong."
    End If

    If managerValidationText.IsBlank(unitOfMeasurement) Then
        managerError.ThrowError _
            sheetInventory.EmptyUnitOfMeasurement, _
            MODULE_NAME & ".HandleUpdateROP", _
            "Satuan barang tidak boleh kosong."
    End If

    ' =====================================================
    ' Step 4: Read calculation values
    ' =====================================================
    averageDailyUsageValue = managerTable.GetValue( _
        tbl, selectedRowIndex, COLUMN_AVERAGE_DAILY_USAGE)

    sdDemandValue = managerTable.GetValue( _
        tbl, selectedRowIndex, COLUMN_SD_DEMAND)

    leadTimeValue = managerTable.GetValue( _
        tbl, selectedRowIndex, COLUMN_LEADTIME)

    ' =====================================================
    ' Step 5: Validate numeric values
    ' =====================================================
    ValidateInventoryNumericValue _
        fieldValue:=averageDailyUsageValue, _
        fieldDescription:="Average Daily Usage", _
        emptyErrorCode:=sheetInventory.EmptyAverageDailyUsage, _
        invalidErrorCode:=sheetInventory.InvalidAverageDailyUsage

    ValidateInventoryNumericValue _
        fieldValue:=sdDemandValue, _
        fieldDescription:="SD Demand", _
        emptyErrorCode:=sheetInventory.EmptySdDemand, _
        invalidErrorCode:=sheetInventory.InvalidSdDemand

    ValidateInventoryNumericValue _
        fieldValue:=leadTimeValue, _
        fieldDescription:="Lead Time", _
        emptyErrorCode:=sheetInventory.EmptyLeadTime, _
        invalidErrorCode:=sheetInventory.InvalidLeadTime

    averageDailyUsage = managerTable.GetDouble( _
        tbl, selectedRowIndex, COLUMN_AVERAGE_DAILY_USAGE)

    sdDemand = managerTable.GetDouble( _
        tbl, selectedRowIndex, COLUMN_SD_DEMAND)

    leadTime = managerTable.GetDouble( _
        tbl, selectedRowIndex, COLUMN_LEADTIME)

    If managerValidationNumeric.IsNegative(averageDailyUsage) Then
        managerError.ThrowError _
            sheetInventory.NegativeAverageDailyUsage, _
            MODULE_NAME & ".HandleUpdateROP", _
            "Average Daily Usage tidak boleh " & _
            "kurang dari nol."
    End If

    If managerValidationNumeric.IsNegative(sdDemand) Then
        managerError.ThrowError _
            sheetInventory.NegativeSdDemand, _
            MODULE_NAME & ".HandleUpdateROP", _
            "SD Demand tidak boleh kurang dari nol."
    End If

    If managerValidationNumeric.IsNegative(leadTime) Then
        managerError.ThrowError _
            sheetInventory.NegativeLeadTime, _
            MODULE_NAME & ".HandleUpdateROP", _
            "Lead Time tidak boleh kurang dari nol."
    End If

    ' =====================================================
    ' Step 6: Calculate preview values
    '
    ' Safety Stock:
    '   SD Demand * Sqr(Lead Time) * 1.65
    '
    ' Reorder Point:
    '   Safety Stock + ADU * Lead Time
    ' =====================================================
    calculatedSafetyStock = _
        sdDemand * Sqr(leadTime) * 1.65

    calculatedReorderPoint = _
        calculatedSafetyStock + _
        (averageDailyUsage * leadTime)

    ' =====================================================
    ' Step 7: Build confirmation message
    ' =====================================================
    confirmationMessage = _
        "Apakah Anda yakin ingin memperbarui " & _
        "parameter inventory berikut?"

    confirmationMessage = confirmationMessage & _
        vbCrLf & vbCrLf & _
        "Manufacturing plant: " & _
        manufacturingPlantName

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "Nama barang: " & catalogueName

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "Satuan: " & unitOfMeasurement

    confirmationMessage = confirmationMessage & _
        vbCrLf & vbCrLf & _
        "Average Daily Usage: " & _
        Format$(averageDailyUsage, "#,##0.00")

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "SD Demand: " & _
        Format$(sdDemand, "#,##0.00")

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "Lead Time: " & _
        Format$(leadTime, "#,##0.00") & _
        " hari"

    confirmationMessage = confirmationMessage & _
        vbCrLf & vbCrLf & _
        "Hasil perhitungan:"

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "Safety Stock: " & _
        Format$(calculatedSafetyStock, "#,##0.00")

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "Reorder Point: " & _
        Format$(calculatedReorderPoint, "#,##0.00")

    confirmationMessage = confirmationMessage & _
        vbCrLf & vbCrLf & _
        "Z Score yang digunakan: 1.65"

    If Not managerForm.Confirm( _
        confirmationMessage, _
        "Konfirmasi Update Inventory") Then

        GoTo CleanUp
    End If

    ' =====================================================
    ' Step 8: Set temporary processing status
    ' =====================================================
    actionCell.value = "MEMPROSES..."
    actionCellChanged = True

    With actionCell
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

        .Font.Bold = True
        .Font.Color = RGB(128, 128, 128)

        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(230, 230, 230)
    End With

    managerForm.SetStatus _
        "Memperbarui reorder point dan safety stock...", _
        True

    ' =====================================================
    ' Step 9: Call service update procedure
    ' =====================================================
    servicePlantStock.UpdateROP _
        averageDailyUsage:=averageDailyUsage, _
        sdDemand:=sdDemand, _
        leadTime:=leadTime, _
        catalogueName:=catalogueName, _
        manufacturingPlantName:= _
            manufacturingPlantName

    ' =====================================================
    ' Step 10: Update displayed calculated values
    ' =====================================================
    managerTable.SetValue _
        tbl, selectedRowIndex, COLUMN_SAFETY_STOCK, _
        calculatedSafetyStock

    managerTable.SetValue _
        tbl, selectedRowIndex, COLUMN_REORDER_POINT, _
        calculatedReorderPoint

    actionCell.value = "BERHASIL"

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
    ' Step 11: Show success message
    ' =====================================================
    successMessage = _
        "Reorder point dan safety stock " & _
        "berhasil diperbarui."

    successMessage = successMessage & _
        vbCrLf & vbCrLf & _
        "Nama barang: " & catalogueName

    successMessage = successMessage & _
        vbCrLf & _
        "Safety Stock: " & _
        Format$(calculatedSafetyStock, "#,##0.00")

    successMessage = successMessage & _
        vbCrLf & _
        "Reorder Point: " & _
        Format$(calculatedReorderPoint, "#,##0.00")

    managerForm.ShowInfo _
        successMessage, _
        "Update Inventory Berhasil"

    ' =====================================================
    ' Step 12: Refresh data from database
    ' =====================================================
    SearchInventory_Click

CleanUp:
    managerForm.ClearStatus

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

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

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".HandleUpdateROP", _
        "Gagal memperbarui reorder point " & _
        "dan safety stock."

End Sub


Private Sub ValidateInventoryNumericValue( _
    ByVal fieldValue As Variant, _
    ByVal fieldDescription As String, _
    ByVal emptyErrorCode As Long, _
    ByVal invalidErrorCode As Long)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    If IsError(fieldValue) Then
        managerError.ThrowError _
            managerTable.CellContainsError, _
            MODULE_NAME & ".ValidateInventoryNumericValue", _
            fieldDescription & _
            " berisi nilai error."
    End If

    If managerValidationText.IsBlank(fieldValue) Then
        managerError.ThrowError _
            emptyErrorCode, _
            MODULE_NAME & ".ValidateInventoryNumericValue", _
            fieldDescription & _
            " tidak boleh kosong."
    End If

    If managerValidationNumeric.IsNotNumber(fieldValue) Then
        managerError.ThrowError _
            invalidErrorCode, _
            MODULE_NAME & ".ValidateInventoryNumericValue", _
            fieldDescription & _
            " harus berupa angka."
    End If

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ValidateInventoryNumericValue"

End Sub
