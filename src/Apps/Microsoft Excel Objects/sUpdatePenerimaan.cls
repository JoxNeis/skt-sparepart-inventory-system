Attribute VB_Name = "sUpdatePenerimaan"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Const MODULE_NAME As String = _
    "SHEET_UPDATE_RECEIVE"

Private Const TABLE_REQUEST_UPDATE_RECEIVE As String = _
    "TableRequestUpdateReceive"

Private Const CELL_MANUFACTURING_PLANT As String = "B4"
Private Const CELL_FROM_DATE As String = "B7"
Private Const CELL_UNTIL_DATE As String = "B8"

' =========================================================
' Table columns (TableRequestUpdateReceive)
'
' TableManager addresses columns by header text. These
' names must exactly match the Excel table headers. The
' comment shows the column position in the table.
' =========================================================
Private Const COL_NUMBER As String = "No"                                   ' 1  A
Private Const COL_REQUEST_DATE As String = "Tanggal Pengajuan"              ' 2  B
Private Const COL_ORDER_UUID As String = "ORDER UUID"                       ' 3  C
Private Const COL_ORDER_PO_ID As String = "ORDER PO ID"                     ' 4  D
Private Const COL_ORDER_AMOUNT As String = "Jumlah Order"                   ' 5  E

' Old receive columns
Private Const COL_OLD_RECEIVE_UUID As String = "UUID LAMA"                  ' 6  F
Private Const COL_OLD_CATALOGUE As String = "NAMA BARANG LAMA"              ' 7  G
Private Const COL_OLD_ARRIVAL_DATE As String = "TANGGAL DATANG LAMA"        ' 8  H
Private Const COL_OLD_VERIFICATION_DATE As String = "TANGGAL VERIFIKASI LAMA" ' 9  I
Private Const COL_OLD_UNIT As String = "SATUAN LAMA"                        ' 10 J
Private Const COL_OLD_TOTAL_AMOUNT As String = "JUMLAH DITERIMA LAMA"       ' 11 K
Private Const COL_OLD_OK_AMOUNT As String = "JUMLAH OK LAMA"                ' 12 L
Private Const COL_OLD_NOT_OK_AMOUNT As String = "JUMLAH NOT OK LAMA"        ' 13 M
Private Const COL_OLD_REMARKS As String = "REMARKS LAMA"                    ' 14 N
Private Const COL_OLD_PIC As String = "PIC LAMA"                            ' 15 O

' New receive columns
Private Const COL_NEW_RECEIVE_UUID As String = "UUID BARU"                  ' 16 P
Private Const COL_NEW_CATALOGUE As String = "NAMA BARANG BARU"              ' 17 Q
Private Const COL_NEW_ARRIVAL_DATE As String = "TANGGAL DATANG BARU"        ' 18 R
Private Const COL_NEW_VERIFICATION_DATE As String = "TANGGAL VERIFIKASI BARU" ' 19 S
Private Const COL_NEW_UNIT As String = "SATUAN BARU"                        ' 20 T
Private Const COL_NEW_TOTAL_AMOUNT As String = "JUMLAH DITERIMA BARU"       ' 21 U
Private Const COL_NEW_OK_AMOUNT As String = "JUMLAH OK BARU"                ' 22 V
Private Const COL_NEW_NOT_OK_AMOUNT As String = "JUMLAH NOT OK BARU"        ' 23 W
Private Const COL_NEW_REMARKS As String = "REMARKS BARU"                    ' 24 X
Private Const COL_NEW_PIC As String = "PIC BARU"                            ' 25 Y

Private Const COL_REQUESTED_BY As String = "PENGAJU"                        ' 26 Z
Private Const COL_ACTION As String = "SETUJUI"                              ' 27 AA

Private Sub Worksheet_SelectionChange( _
    ByVal Target As Range)

    On Error GoTo ErrorHandler

    Dim tbl As ListObject

    Dim errInfo As ErrorInfo

    ' =====================================================
    ' Date input cells
    ' =====================================================
    If FormManager.IsInRange( _
        Target, _
        Me.Range(CELL_FROM_DATE & ":" & CELL_UNTIL_DATE)) Then

        Application.EnableEvents = False

        FormManager.ShowDatePicker Target

        GoTo CleanUp

    End If

    ' =====================================================
    ' Get update request table
    ' =====================================================
    Set tbl = TableManager.GetTable( _
        Me, TABLE_REQUEST_UPDATE_RECEIVE)

    If Not TableManager.HasRows(tbl) Then
        GoTo CleanUp
    End If

    ' =====================================================
    ' Detect SETUJUI column click
    ' =====================================================
    If Not TableManager.IsInColumn( _
        Target, tbl, COL_ACTION) Then

        GoTo CleanUp
    End If

    If Not ValidationText.IsSameText( _
        pHelper.TrimText(CStr(Target.value)), _
        "SETUJUI") Then

        GoTo CleanUp
    End If

    Application.EnableEvents = False

    HandleApproveReceiveRequestUpdate _
        tbl:=tbl, _
        targetCell:=Target

CleanUp:
    Set tbl = Nothing

    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    Set tbl = Nothing

    FormManager.ClearStatus
    Application.EnableEvents = True

    ErrorManager.ShowError errInfo, _
        "Gagal memproses persetujuan " & _
        "update penerimaan."

End Sub
Public Sub CariPengajuanUpdateReceive()

    On Error GoTo ErrorHandler

    Dim tbl As ListObject
    Dim rs As ADODB.recordset

    Dim manufacturingPlantName As String
    Dim fromDate As Date
    Dim untilDate As Date

    Dim resultCount As Long

    Dim state As FormState
    Dim errInfo As ErrorInfo

    state = FormManager.BeginBusy( _
        "Mengambil data pengajuan update penerimaan...")

    ' =====================================================
    ' Step 1: Read and validate filter values
    ' =====================================================
    manufacturingPlantName = FormManager.GetCellText( _
        Me, CELL_MANUFACTURING_PLANT)

    ValidateReceiveSearchParameters _
        manufacturingPlantName:=manufacturingPlantName, _
        fromDateValue:=FormManager.GetCellValue(Me, CELL_FROM_DATE), _
        untilDateValue:=FormManager.GetCellValue(Me, CELL_UNTIL_DATE)

    fromDate = DateValue( _
        FormManager.GetCellValue(Me, CELL_FROM_DATE))

    untilDate = DateValue( _
        FormManager.GetCellValue(Me, CELL_UNTIL_DATE))

    ' =====================================================
    ' Step 2: Get the target table
    ' =====================================================
    Set tbl = TableManager.GetTable( _
        Me, TABLE_REQUEST_UPDATE_RECEIVE)

    ' =====================================================
    ' Step 3: Clear previous records
    ' =====================================================
    ClearReceiveRequestUpdateTable tbl

    ' =====================================================
    ' Step 4: Retrieve records from the service
    ' =====================================================
    Set rs = serviceTransactionRequestUpdate.SelectReceiveRequestUpdate( _
        manufacturingPlantName:=manufacturingPlantName, _
        fromDate:=fromDate, _
        untilDate:=untilDate)

    If rs Is Nothing Then
        ErrorManager.ThrowError _
            DatabaseManager.EmptyRecordset, _
            MODULE_NAME & ".CariPengajuanUpdateReceive", _
            "DAO tidak mengembalikan recordset."
    End If

    If rs.State <> adStateOpen Then
        ErrorManager.ThrowError _
            DatabaseManager.RecordsetNotOpen, _
            MODULE_NAME & ".CariPengajuanUpdateReceive", _
            "Recordset pengajuan update penerimaan " & _
            "tidak terbuka."
    End If

    ' =====================================================
    ' Step 5: Fill the table
    ' =====================================================
    resultCount = FillReceiveRequestUpdateTable( _
        tbl:=tbl, _
        rs:=rs)

    ' =====================================================
    ' Step 6: Apply formatting
    ' =====================================================
    ApplyReceiveRequestUpdateFormatting tbl

    ' =====================================================
    ' Step 7: Show result
    ' =====================================================
    If resultCount = 0 Then
        FormManager.ShowInfo _
            "Tidak ada pengajuan update penerimaan " & _
            "yang ditemukan.", _
            "Pengajuan Update Penerimaan"
    Else
        FormManager.ShowInfo _
            CStr(resultCount) & _
            " pengajuan update penerimaan ditemukan.", _
            "Pengajuan Update Penerimaan"
    End If

CleanUp:
    DatabaseManager.CloseRecordset rs
    Set tbl = Nothing

    FormManager.EndBusy state
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    DatabaseManager.CloseRecordset rs
    Set tbl = Nothing

    FormManager.EndBusy state

    ErrorManager.ShowError errInfo, _
        "Gagal menampilkan pengajuan " & _
        "update penerimaan."

End Sub


Private Function FillReceiveRequestUpdateTable( _
    ByVal tbl As ListObject, _
    ByVal rs As ADODB.recordset) As Long

    On Error GoTo ErrorHandler

    Dim rowIndex As Long
    Dim rowNumber As Long

    Dim errInfo As ErrorInfo

    If tbl Is Nothing Then
        ErrorManager.ThrowError _
            TableManager.TableNotFound, _
            MODULE_NAME & ".FillReceiveRequestUpdateTable", _
            "Table pengajuan update penerimaan " & _
            "tidak ditemukan."
    End If

    If rs Is Nothing Then
        ErrorManager.ThrowError _
            DatabaseManager.EmptyRecordset, _
            MODULE_NAME & ".FillReceiveRequestUpdateTable", _
            "Recordset pengajuan update penerimaan kosong."
    End If

    rowNumber = 0

    Do While Not rs.EOF

        rowNumber = rowNumber + 1

        rowIndex = TableManager.AddRow(tbl)

        WriteReceiveRequestUpdateRow _
            tbl:=tbl, _
            rowIndex:=rowIndex, _
            rs:=rs, _
            rowNumber:=rowNumber

        rs.MoveNext
    Loop

    FillReceiveRequestUpdateTable = rowNumber

CleanUp:
    Exit Function

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".FillReceiveRequestUpdateTable", _
        "Gagal mengisi tabel pengajuan " & _
        "update penerimaan."

End Function


Private Sub WriteReceiveRequestUpdateRow( _
    ByVal tbl As ListObject, _
    ByVal rowIndex As Long, _
    ByVal rs As ADODB.recordset, _
    ByVal rowNumber As Long)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    ' =====================================================
    ' General request and order information
    ' =====================================================

    ' A: No
    TableManager.SetValue tbl, rowIndex, COL_NUMBER, _
        rowNumber

    ' B: Tanggal Pengajuan
    TableManager.SetValue tbl, rowIndex, COL_REQUEST_DATE, _
        RecordsetDate(rs, "update_request_date")

    ' C: Order UUID
    TableManager.SetValue tbl, rowIndex, COL_ORDER_UUID, _
        DatabaseManager.FieldText(rs, "old_order_transaction_uuid")

    ' D: Order PO ID
    TableManager.SetValue tbl, rowIndex, COL_ORDER_PO_ID, _
        DatabaseManager.FieldText(rs, "old_order_poid")

    ' E: Jumlah Order
    TableManager.SetValue tbl, rowIndex, COL_ORDER_AMOUNT, _
        DatabaseManager.FieldDouble(rs, "old_order_amount")

    ' =====================================================
    ' Old receive information
    ' =====================================================

    ' F: UUID Lama
    TableManager.SetValue tbl, rowIndex, COL_OLD_RECEIVE_UUID, _
        DatabaseManager.FieldText(rs, "old_transaction_uuid")

    ' G: Nama Barang Lama
    TableManager.SetValue tbl, rowIndex, COL_OLD_CATALOGUE, _
        DatabaseManager.FieldText(rs, "old_catalogue_name")

    ' H: Tanggal Datang Lama
    TableManager.SetValue tbl, rowIndex, COL_OLD_ARRIVAL_DATE, _
        RecordsetDate(rs, "old_arrival_date")

    ' I: Tanggal Verifikasi Lama
    TableManager.SetValue tbl, rowIndex, COL_OLD_VERIFICATION_DATE, _
        RecordsetDate(rs, "old_verification_date")

    ' J: Satuan Lama
    TableManager.SetValue tbl, rowIndex, COL_OLD_UNIT, _
        DatabaseManager.FieldText(rs, "old_unit_of_measurement")

    ' K: Jumlah Diterima Lama
    TableManager.SetValue tbl, rowIndex, COL_OLD_TOTAL_AMOUNT, _
        DatabaseManager.FieldDouble(rs, "old_total_receive_amount")

    ' L: Jumlah OK Lama
    TableManager.SetValue tbl, rowIndex, COL_OLD_OK_AMOUNT, _
        DatabaseManager.FieldDouble(rs, "old_ok_amount")

    ' M: Jumlah NOT OK Lama
    TableManager.SetValue tbl, rowIndex, COL_OLD_NOT_OK_AMOUNT, _
        DatabaseManager.FieldDouble(rs, "old_not_ok_amount")

    ' N: Remarks Lama
    TableManager.SetValue tbl, rowIndex, COL_OLD_REMARKS, _
        DatabaseManager.FieldText(rs, "old_remarks")

    ' O: PIC Lama
    TableManager.SetValue tbl, rowIndex, COL_OLD_PIC, _
        DatabaseManager.FieldText(rs, "old_pic")

    ' =====================================================
    ' New receive information
    ' =====================================================

    ' P: UUID Baru
    TableManager.SetValue tbl, rowIndex, COL_NEW_RECEIVE_UUID, _
        DatabaseManager.FieldText(rs, "new_transaction_uuid")

    ' Q: Nama Barang Baru
    TableManager.SetValue tbl, rowIndex, COL_NEW_CATALOGUE, _
        DatabaseManager.FieldText(rs, "new_catalogue_name")

    ' R: Tanggal Datang Baru
    TableManager.SetValue tbl, rowIndex, COL_NEW_ARRIVAL_DATE, _
        RecordsetDate(rs, "new_arrival_date")

    ' S: Tanggal Verifikasi Baru
    TableManager.SetValue tbl, rowIndex, COL_NEW_VERIFICATION_DATE, _
        RecordsetDate(rs, "new_verification_date")

    ' T: Satuan Baru
    TableManager.SetValue tbl, rowIndex, COL_NEW_UNIT, _
        DatabaseManager.FieldText(rs, "new_unit_of_measurement")

    ' U: Jumlah Diterima Baru
    TableManager.SetValue tbl, rowIndex, COL_NEW_TOTAL_AMOUNT, _
        DatabaseManager.FieldDouble(rs, "new_total_receive_amount")

    ' V: Jumlah OK Baru
    TableManager.SetValue tbl, rowIndex, COL_NEW_OK_AMOUNT, _
        DatabaseManager.FieldDouble(rs, "new_ok_amount")

    ' W: Jumlah NOT OK Baru
    TableManager.SetValue tbl, rowIndex, COL_NEW_NOT_OK_AMOUNT, _
        DatabaseManager.FieldDouble(rs, "new_not_ok_amount")

    ' X: Remarks Baru
    TableManager.SetValue tbl, rowIndex, COL_NEW_REMARKS, _
        DatabaseManager.FieldText(rs, "new_remarks")

    ' Y: PIC Baru
    TableManager.SetValue tbl, rowIndex, COL_NEW_PIC, _
        DatabaseManager.FieldText(rs, "new_pic")

    ' =====================================================
    ' Request and action
    ' =====================================================

    ' Z: Pengaju
    TableManager.SetValue tbl, rowIndex, COL_REQUESTED_BY, _
        DatabaseManager.FieldText(rs, "update_request_by")

    ' AA: Action
    TableManager.SetValue tbl, rowIndex, COL_ACTION, _
        "SETUJUI"

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".WriteReceiveRequestUpdateRow", _
        "Gagal menuliskan satu baris pengajuan " & _
        "update penerimaan."

End Sub


Private Sub ClearReceiveRequestUpdateTable( _
    ByVal tbl As ListObject)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    If tbl Is Nothing Then
        ErrorManager.ThrowError _
            TableManager.TableNotFound, _
            MODULE_NAME & ".ClearReceiveRequestUpdateTable", _
            "Table pengajuan update penerimaan " & _
            "tidak ditemukan."
    End If

    ' Delete existing table rows while preserving
    ' the table header and table formatting.
    TableManager.ClearRows tbl

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".ClearReceiveRequestUpdateTable", _
        "Gagal membersihkan tabel pengajuan " & _
        "update penerimaan."

End Sub


Private Sub ApplyReceiveRequestUpdateFormatting( _
    ByVal tbl As ListObject)

    On Error GoTo ErrorHandler

    Dim actionRange As Range

    Dim errInfo As ErrorInfo

    If tbl Is Nothing Then
        ErrorManager.ThrowError _
            TableManager.TableNotFound, _
            MODULE_NAME & ".ApplyReceiveRequestUpdateFormatting", _
            "Table pengajuan update penerimaan " & _
            "tidak ditemukan."
    End If

    If Not TableManager.HasRows(tbl) Then
        Exit Sub
    End If

    ' Request date
    TableManager.FormatColumn tbl, COL_REQUEST_DATE, _
        "dd-mmm-yyyy"

    ' Old receive dates
    TableManager.FormatColumn tbl, COL_OLD_ARRIVAL_DATE, _
        "dd-mmm-yyyy"

    TableManager.FormatColumn tbl, COL_OLD_VERIFICATION_DATE, _
        "dd-mmm-yyyy"

    ' New receive dates
    TableManager.FormatColumn tbl, COL_NEW_ARRIVAL_DATE, _
        "dd-mmm-yyyy"

    TableManager.FormatColumn tbl, COL_NEW_VERIFICATION_DATE, _
        "dd-mmm-yyyy"

    ' Order amount
    TableManager.FormatColumn tbl, COL_ORDER_AMOUNT, _
        "#,##0.00"

    ' Old receive amounts
    TableManager.FormatColumn tbl, COL_OLD_TOTAL_AMOUNT, _
        "#,##0.00"

    TableManager.FormatColumn tbl, COL_OLD_OK_AMOUNT, _
        "#,##0.00"

    TableManager.FormatColumn tbl, COL_OLD_NOT_OK_AMOUNT, _
        "#,##0.00"

    ' New receive amounts
    TableManager.FormatColumn tbl, COL_NEW_TOTAL_AMOUNT, _
        "#,##0.00"

    TableManager.FormatColumn tbl, COL_NEW_OK_AMOUNT, _
        "#,##0.00"

    TableManager.FormatColumn tbl, COL_NEW_NOT_OK_AMOUNT, _
        "#,##0.00"

    ' General alignment (columns No .. SETUJUI)
    Me.Range( _
        TableManager.GetColumnRange(tbl, COL_NUMBER), _
        TableManager.GetColumnRange(tbl, COL_ACTION) _
    ).VerticalAlignment = xlCenter

    ' Format action column.
    Set actionRange = TableManager.GetColumnRange( _
        tbl, COL_ACTION)

    If Not actionRange Is Nothing Then
        With actionRange
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter

            .Font.Bold = True
            .Font.Color = RGB(0, 97, 0)

            .Interior.Pattern = xlSolid
            .Interior.Color = RGB(198, 239, 206)
        End With
    End If

CleanUp:
    Set actionRange = Nothing
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    Set actionRange = Nothing

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".ApplyReceiveRequestUpdateFormatting", _
        "Gagal menerapkan format tabel pengajuan " & _
        "update penerimaan."

End Sub


Private Sub ValidateReceiveSearchParameters( _
    ByVal manufacturingPlantName As String, _
    ByVal fromDateValue As Variant, _
    ByVal untilDateValue As Variant)

    On Error GoTo ErrorHandler

    Dim fromDate As Date
    Dim untilDate As Date

    Dim errInfo As ErrorInfo

    If ValidationText.IsBlank(manufacturingPlantName) Then
        ErrorManager.ThrowError _
            sheetUpdatePenerimaan.EmptyPlantName, _
            MODULE_NAME & ".ValidateReceiveSearchParameters", _
            "Lokasi manufacturing plant wajib diisi."
    End If

    If IsEmpty(fromDateValue) Then
        ErrorManager.ThrowError _
            sheetUpdatePenerimaan.EmptyFromDate, _
            MODULE_NAME & ".ValidateReceiveSearchParameters", _
            "Tanggal awal wajib diisi."
    End If

    If Not IsDate(fromDateValue) Then
        ErrorManager.ThrowError _
            sheetUpdatePenerimaan.InvalidFromDate, _
            MODULE_NAME & ".ValidateReceiveSearchParameters", _
            "Tanggal awal tidak valid."
    End If

    If IsEmpty(untilDateValue) Then
        ErrorManager.ThrowError _
            sheetUpdatePenerimaan.EmptyUntilDate, _
            MODULE_NAME & ".ValidateReceiveSearchParameters", _
            "Tanggal akhir wajib diisi."
    End If

    If Not IsDate(untilDateValue) Then
        ErrorManager.ThrowError _
            sheetUpdatePenerimaan.InvalidUntilDate, _
            MODULE_NAME & ".ValidateReceiveSearchParameters", _
            "Tanggal akhir tidak valid."
    End If

    fromDate = DateValue(fromDateValue)
    untilDate = DateValue(untilDateValue)

    If untilDate < fromDate Then
        ErrorManager.ThrowError _
            sheetUpdatePenerimaan.InvalidDateRange, _
            MODULE_NAME & ".ValidateReceiveSearchParameters", _
            "Tanggal akhir tidak boleh lebih kecil " & _
            "dari tanggal awal."
    End If

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".ValidateReceiveSearchParameters", _
        "Filter pengajuan update penerimaan " & _
        "tidak valid."

End Sub


' Date field for a table cell: Null -> blank cell
' (DatabaseManager.FieldDate would return date 0).
Private Function RecordsetDate( _
    ByVal rs As ADODB.recordset, _
    ByVal fieldName As String) As Variant

    If DatabaseManager.FieldIsNull(rs, fieldName) Then
        RecordsetDate = vbNullString
    Else
        RecordsetDate = DatabaseManager.FieldDate(rs, fieldName)
    End If

End Function

Private Sub HandleApproveReceiveRequestUpdate( _
    ByVal tbl As ListObject, _
    ByVal targetCell As Range)

    On Error GoTo ErrorHandler

    Dim selectedRowIndex As Long

    Dim originalTransactionUUID As String
    Dim proposedTransactionUUID As String
    Dim currentUser As String
    Dim actionText As String

    Dim oldCatalogueName As String
    Dim newCatalogueName As String

    Dim oldManufacturingPlantName As String
    Dim newManufacturingPlantName As String

    Dim oldOKAmount As Double
    Dim newOKAmount As Double

    Dim confirmationMessage As String

    ' True while the action cell shows "MEMPROSES...".
    Dim isProcessing As Boolean

    Dim errInfo As ErrorInfo

    ' =====================================================
    ' Step 1: Validate objects
    ' =====================================================
    If tbl Is Nothing Then
        ErrorManager.ThrowError _
            TableManager.TableNotFound, _
            MODULE_NAME & ".HandleApproveReceiveRequestUpdate", _
            "Tabel pengajuan update penerimaan " & _
            "belum diinisialisasi."
    End If

    If targetCell Is Nothing Then
        ErrorManager.ThrowError _
            sheetUpdatePenerimaan.EmptyApprovalCell, _
            MODULE_NAME & ".HandleApproveReceiveRequestUpdate", _
            "Cell persetujuan tidak ditemukan."
    End If

    If Not TableManager.HasRows(tbl) Then
        ErrorManager.ThrowError _
            sheetUpdatePenerimaan.EmptyRequestUpdateTable, _
            MODULE_NAME & ".HandleApproveReceiveRequestUpdate", _
            "Tabel pengajuan update penerimaan " & _
            "tidak memiliki data."
    End If

    ' =====================================================
    ' Step 2: Determine selected table row
    ' =====================================================
    selectedRowIndex = TableManager.GetRowIndex( _
        tbl, targetCell)

    If Not TableManager.IsValidRow(tbl, selectedRowIndex) Then
        ErrorManager.ThrowError _
            TableManager.InvalidRow, _
            MODULE_NAME & ".HandleApproveReceiveRequestUpdate", _
            "Baris pengajuan update penerimaan " & _
            "tidak valid."
    End If

    ' =====================================================
    ' Step 3: Validate action text
    ' =====================================================
    actionText = TableManager.GetText( _
        tbl, selectedRowIndex, COL_ACTION)

    If Not ValidationText.IsSameText( _
        actionText, _
        "SETUJUI") Then

        GoTo CleanUp
    End If

    ' =====================================================
    ' Step 4: Read receive transaction UUIDs
    ' =====================================================
    originalTransactionUUID = TableManager.GetText( _
        tbl, selectedRowIndex, COL_OLD_RECEIVE_UUID)

    proposedTransactionUUID = TableManager.GetText( _
        tbl, selectedRowIndex, COL_NEW_RECEIVE_UUID)

    If ValidationText.IsBlank(originalTransactionUUID) Then
        ErrorManager.ThrowError _
            sheetUpdatePenerimaan.EmptyOriginalReceiveUUID, _
            MODULE_NAME & ".HandleApproveReceiveRequestUpdate", _
            "UUID transaksi receive original " & _
            "tidak boleh kosong."
    End If

    If ValidationText.IsBlank(proposedTransactionUUID) Then
        ErrorManager.ThrowError _
            sheetUpdatePenerimaan.EmptyProposedReceiveUUID, _
            MODULE_NAME & ".HandleApproveReceiveRequestUpdate", _
            "UUID transaksi receive proposed " & _
            "tidak boleh kosong."
    End If

    If ValidationText.IsSameText( _
        originalTransactionUUID, _
        proposedTransactionUUID) Then

        ErrorManager.ThrowError _
            sheetUpdatePenerimaan.SameOriginalProposedReceiveUUID, _
            MODULE_NAME & ".HandleApproveReceiveRequestUpdate", _
            "UUID transaksi receive original dan " & _
            "proposed tidak boleh sama."
    End If

    ' =====================================================
    ' Step 5: Read catalogue names
    ' =====================================================
    oldCatalogueName = TableManager.GetText( _
        tbl, selectedRowIndex, COL_OLD_CATALOGUE)

    newCatalogueName = TableManager.GetText( _
        tbl, selectedRowIndex, COL_NEW_CATALOGUE)

    If ValidationText.IsBlank(oldCatalogueName) Then
        ErrorManager.ThrowError _
            sheetUpdatePenerimaan.EmptyOriginalCatalogueName, _
            MODULE_NAME & ".HandleApproveReceiveRequestUpdate", _
            "Nama catalogue receive original " & _
            "tidak boleh kosong."
    End If

    If ValidationText.IsBlank(newCatalogueName) Then
        ErrorManager.ThrowError _
            sheetUpdatePenerimaan.EmptyProposedCatalogueName, _
            MODULE_NAME & ".HandleApproveReceiveRequestUpdate", _
            "Nama catalogue receive proposed " & _
            "tidak boleh kosong."
    End If

    ' =====================================================
    ' Step 6: Read OK amounts for confirmation
    '
    ' These values are only displayed in the confirmation
    ' message. The service recalculates exact current stock.
    ' =====================================================
    oldOKAmount = TableManager.GetDouble( _
        tbl, selectedRowIndex, COL_OLD_OK_AMOUNT)

    newOKAmount = TableManager.GetDouble( _
        tbl, selectedRowIndex, COL_NEW_OK_AMOUNT)

    If ValidationNumeric.IsNegative(oldOKAmount) Then
        ErrorManager.ThrowError _
            sheetUpdatePenerimaan.NegativeOriginalOkAmount, _
            MODULE_NAME & ".HandleApproveReceiveRequestUpdate", _
            "Jumlah OK receive original tidak boleh " & _
            "kurang dari nol."
    End If

    If ValidationNumeric.IsNegative(newOKAmount) Then
        ErrorManager.ThrowError _
            sheetUpdatePenerimaan.NegativeProposedOkAmount, _
            MODULE_NAME & ".HandleApproveReceiveRequestUpdate", _
            "Jumlah OK receive proposed tidak boleh " & _
            "kurang dari nol."
    End If

    ' =====================================================
    ' Step 7: Read manufacturing plant names
    '
    ' The current table does not have separate columns for
    ' old and new manufacturing plants. Therefore, both
    ' values use the selected manufacturing plant in B4.
    ' =====================================================
    oldManufacturingPlantName = FormManager.GetCellText( _
        Me, CELL_MANUFACTURING_PLANT)

    newManufacturingPlantName = _
        oldManufacturingPlantName

    If ValidationText.IsBlank(oldManufacturingPlantName) Then
        ErrorManager.ThrowError _
            sheetUpdatePenerimaan.EmptyOriginalPlantName, _
            MODULE_NAME & ".HandleApproveReceiveRequestUpdate", _
            "Nama manufacturing plant original " & _
            "tidak boleh kosong."
    End If

    If ValidationText.IsBlank(newManufacturingPlantName) Then
        ErrorManager.ThrowError _
            sheetUpdatePenerimaan.EmptyProposedPlantName, _
            MODULE_NAME & ".HandleApproveReceiveRequestUpdate", _
            "Nama manufacturing plant proposed " & _
            "tidak boleh kosong."
    End If

    ' =====================================================
    ' Step 8: Determine current user
    ' =====================================================
    currentUser = pHelper.TrimText( _
        pConfig.currentUser())

    If ValidationText.IsBlank(currentUser) Then
        currentUser = pHelper.TrimText( _
            Environ$("Username"))
    End If

    If ValidationText.IsBlank(currentUser) Then
        ErrorManager.ThrowError _
            sheetUpdatePenerimaan.EmptyApproverUser, _
            MODULE_NAME & ".HandleApproveReceiveRequestUpdate", _
            "User yang menyetujui tidak dapat " & _
            "ditentukan."
    End If

    ' =====================================================
    ' Step 9: Build confirmation message
    ' =====================================================
    confirmationMessage = _
        "Apakah Anda yakin ingin menyetujui " & _
        "pengajuan update penerimaan ini?"

    confirmationMessage = confirmationMessage & _
        vbCrLf & vbCrLf & _
        "UUID receive original:" & vbCrLf & _
        originalTransactionUUID

    confirmationMessage = confirmationMessage & _
        vbCrLf & vbCrLf & _
        "UUID receive proposed:" & vbCrLf & _
        proposedTransactionUUID

    confirmationMessage = confirmationMessage & _
        vbCrLf & vbCrLf & _
        "Catalogue original: " & _
        oldCatalogueName

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "Jumlah OK original: " & _
        Format$(oldOKAmount, "#,##0.00")

    confirmationMessage = confirmationMessage & _
        vbCrLf & vbCrLf & _
        "Catalogue proposed: " & _
        newCatalogueName

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "Jumlah OK proposed: " & _
        Format$(newOKAmount, "#,##0.00")

    confirmationMessage = confirmationMessage & _
        vbCrLf & vbCrLf & _
        "Manufacturing plant: " & _
        oldManufacturingPlantName

    confirmationMessage = confirmationMessage & _
        vbCrLf & vbCrLf & _
        "Tindakan ini akan:"

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "1. Menyetujui pengajuan update penerimaan."

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "2. Membalik transaksi receive original."

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "3. Mengaktifkan transaksi receive proposed."

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "4. Menghitung ulang stok manufacturing plant."

    If Not FormManager.Confirm( _
        confirmationMessage, _
        "Konfirmasi Persetujuan Update Penerimaan") Then

        GoTo CleanUp
    End If

    ' =====================================================
    ' Step 10: Set temporary processing status
    ' =====================================================
    TableManager.SetValue tbl, selectedRowIndex, COL_ACTION, _
        "MEMPROSES..."

    isProcessing = True

    With TableManager.GetCell(tbl, selectedRowIndex, COL_ACTION)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

        .Font.Bold = True
        .Font.Color = RGB(128, 128, 128)

        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(230, 230, 230)
    End With

    FormManager.SetStatus _
        "Menyetujui pengajuan update penerimaan...", True

    ' =====================================================
    ' Step 11: Execute approval
    '
    ' This uses serviceTransactionRequestUpdate.
    '
    ' updateOrder = False:
    '   The proposed transaction is a receive update, so
    '   receive.order_id must not be moved.
    '
    ' updateStockMovement = True:
    '   Recalculate current stock after reversing the old
    '   receive and activating the proposed receive.
    ' =====================================================
    serviceTransactionRequestUpdate.ApproveRequestUpdate _
        originalTransactionUUID:= _
            originalTransactionUUID, _
        proposedTransactionUUID:= _
            proposedTransactionUUID, _
        user:=currentUser, _
        updateOrder:=False, _
        updateStockMovement:=True, _
        oldCatalogueName:= _
            oldCatalogueName, _
        newCatalogueName:= _
            newCatalogueName, _
        oldManufacturingPlantName:= _
            oldManufacturingPlantName, _
        newManufacturingPlantName:= _
            newManufacturingPlantName

    ' =====================================================
    ' Step 12: Approval succeeded
    ' =====================================================
    TableManager.SetValue tbl, selectedRowIndex, COL_ACTION, _
        "DISETUJUI"

    isProcessing = False

    With TableManager.GetCell(tbl, selectedRowIndex, COL_ACTION)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)

        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(0, 128, 0)
    End With

    FormManager.ClearStatus

    FormManager.ShowInfo _
        "Pengajuan update penerimaan " & _
        "berhasil disetujui.", _
        "Persetujuan Berhasil"

    ' =====================================================
    ' Step 13: Refresh pending update requests
    ' =====================================================
    CariPengajuanUpdateReceive

CleanUp:
    FormManager.ClearStatus

    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ' Restore action cell only if it had entered the
    ' processing state.
    If isProcessing Then

        TableManager.SetValue tbl, selectedRowIndex, COL_ACTION, _
            "SETUJUI"

        With TableManager.GetCell(tbl, selectedRowIndex, COL_ACTION)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter

            .Font.Bold = True
            .Font.Color = RGB(0, 97, 0)

            .Interior.Pattern = xlSolid
            .Interior.Color = RGB(198, 239, 206)
        End With

    End If

    FormManager.ClearStatus

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".HandleApproveReceiveRequestUpdate", _
        "Gagal menyetujui pengajuan " & _
        "update penerimaan."

End Sub
