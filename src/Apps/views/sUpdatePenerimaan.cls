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
' managerTable addresses columns by header text. These
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
    If managerForm.IsInRange( _
        Target, _
        Me.Range(CELL_FROM_DATE & ":" & CELL_UNTIL_DATE)) Then

        Application.EnableEvents = False

        managerForm.ShowDatePicker Target

        GoTo CleanUp

    End If

    ' =====================================================
    ' Get update request table
    ' =====================================================
    Set tbl = managerTable.GetTable( _
        Me, TABLE_REQUEST_UPDATE_RECEIVE)

    If Not managerTable.HasRows(tbl) Then
        GoTo CleanUp
    End If

    ' =====================================================
    ' Detect SETUJUI column click
    ' =====================================================
    If Not managerTable.IsInColumn( _
        Target, tbl, COL_ACTION) Then

        GoTo CleanUp
    End If

    If Not managerValidationText.IsSameText( _
        coreHelper.TrimText(CStr(Target.value)), _
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
    errInfo = managerError.CaptureError()

    Set tbl = Nothing

    managerForm.ClearStatus
    Application.EnableEvents = True

    managerError.ShowError errInfo, _
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

    state = managerForm.BeginBusy( _
        "Mengambil data pengajuan update penerimaan...")

    ' =====================================================
    ' Step 1: Read and validate filter values
    ' =====================================================
    manufacturingPlantName = managerForm.GetCellText( _
        Me, CELL_MANUFACTURING_PLANT)

    ValidateReceiveSearchParameters _
        manufacturingPlantName:=manufacturingPlantName, _
        fromDateValue:=managerForm.GetCellValue(Me, CELL_FROM_DATE), _
        untilDateValue:=managerForm.GetCellValue(Me, CELL_UNTIL_DATE)

    fromDate = DateValue( _
        managerForm.GetCellValue(Me, CELL_FROM_DATE))

    untilDate = DateValue( _
        managerForm.GetCellValue(Me, CELL_UNTIL_DATE))

    ' =====================================================
    ' Step 2: Get the target table
    ' =====================================================
    Set tbl = managerTable.GetTable( _
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
        managerError.ThrowError _
            managerDatabase.EmptyRecordset, _
            MODULE_NAME & ".CariPengajuanUpdateReceive", _
            "DAO tidak mengembalikan recordset."
    End If

    If rs.State <> adStateOpen Then
        managerError.ThrowError _
            managerDatabase.RecordsetNotOpen, _
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
        managerForm.ShowInfo _
            "Tidak ada pengajuan update penerimaan " & _
            "yang ditemukan.", _
            "Pengajuan Update Penerimaan"
    Else
        managerForm.ShowInfo _
            CStr(resultCount) & _
            " pengajuan update penerimaan ditemukan.", _
            "Pengajuan Update Penerimaan"
    End If

CleanUp:
    managerDatabase.CloseRecordset rs
    Set tbl = Nothing

    managerForm.EndBusy state
    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerDatabase.CloseRecordset rs
    Set tbl = Nothing

    managerForm.EndBusy state

    managerError.ShowError errInfo, _
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
        managerError.ThrowError _
            managerTable.TableNotFound, _
            MODULE_NAME & ".FillReceiveRequestUpdateTable", _
            "Table pengajuan update penerimaan " & _
            "tidak ditemukan."
    End If

    If rs Is Nothing Then
        managerError.ThrowError _
            managerDatabase.EmptyRecordset, _
            MODULE_NAME & ".FillReceiveRequestUpdateTable", _
            "Recordset pengajuan update penerimaan kosong."
    End If

    rowNumber = 0

    Do While Not rs.EOF

        rowNumber = rowNumber + 1

        rowIndex = managerTable.AddRow(tbl)

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
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
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
    managerTable.SetValue tbl, rowIndex, COL_NUMBER, _
        rowNumber

    ' B: Tanggal Pengajuan
    managerTable.SetValue tbl, rowIndex, COL_REQUEST_DATE, _
        RecordsetDate(rs, "update_request_date")

    ' C: Order UUID
    managerTable.SetValue tbl, rowIndex, COL_ORDER_UUID, _
        managerDatabase.FieldText(rs, "old_order_transaction_uuid")

    ' D: Order PO ID
    managerTable.SetValue tbl, rowIndex, COL_ORDER_PO_ID, _
        managerDatabase.FieldText(rs, "old_order_poid")

    ' E: Jumlah Order
    managerTable.SetValue tbl, rowIndex, COL_ORDER_AMOUNT, _
        managerDatabase.FieldDouble(rs, "old_order_amount")

    ' =====================================================
    ' Old receive information
    ' =====================================================

    ' F: UUID Lama
    managerTable.SetValue tbl, rowIndex, COL_OLD_RECEIVE_UUID, _
        managerDatabase.FieldText(rs, "old_transaction_uuid")

    ' G: Nama Barang Lama
    managerTable.SetValue tbl, rowIndex, COL_OLD_CATALOGUE, _
        managerDatabase.FieldText(rs, "old_catalogue_name")

    ' H: Tanggal Datang Lama
    managerTable.SetValue tbl, rowIndex, COL_OLD_ARRIVAL_DATE, _
        RecordsetDate(rs, "old_arrival_date")

    ' I: Tanggal Verifikasi Lama
    managerTable.SetValue tbl, rowIndex, COL_OLD_VERIFICATION_DATE, _
        RecordsetDate(rs, "old_verification_date")

    ' J: Satuan Lama
    managerTable.SetValue tbl, rowIndex, COL_OLD_UNIT, _
        managerDatabase.FieldText(rs, "old_unit_of_measurement")

    ' K: Jumlah Diterima Lama
    managerTable.SetValue tbl, rowIndex, COL_OLD_TOTAL_AMOUNT, _
        managerDatabase.FieldDouble(rs, "old_total_receive_amount")

    ' L: Jumlah OK Lama
    managerTable.SetValue tbl, rowIndex, COL_OLD_OK_AMOUNT, _
        managerDatabase.FieldDouble(rs, "old_ok_amount")

    ' M: Jumlah NOT OK Lama
    managerTable.SetValue tbl, rowIndex, COL_OLD_NOT_OK_AMOUNT, _
        managerDatabase.FieldDouble(rs, "old_not_ok_amount")

    ' N: Remarks Lama
    managerTable.SetValue tbl, rowIndex, COL_OLD_REMARKS, _
        managerDatabase.FieldText(rs, "old_remarks")

    ' O: PIC Lama
    managerTable.SetValue tbl, rowIndex, COL_OLD_PIC, _
        managerDatabase.FieldText(rs, "old_pic")

    ' =====================================================
    ' New receive information
    ' =====================================================

    ' P: UUID Baru
    managerTable.SetValue tbl, rowIndex, COL_NEW_RECEIVE_UUID, _
        managerDatabase.FieldText(rs, "new_transaction_uuid")

    ' Q: Nama Barang Baru
    managerTable.SetValue tbl, rowIndex, COL_NEW_CATALOGUE, _
        managerDatabase.FieldText(rs, "new_catalogue_name")

    ' R: Tanggal Datang Baru
    managerTable.SetValue tbl, rowIndex, COL_NEW_ARRIVAL_DATE, _
        RecordsetDate(rs, "new_arrival_date")

    ' S: Tanggal Verifikasi Baru
    managerTable.SetValue tbl, rowIndex, COL_NEW_VERIFICATION_DATE, _
        RecordsetDate(rs, "new_verification_date")

    ' T: Satuan Baru
    managerTable.SetValue tbl, rowIndex, COL_NEW_UNIT, _
        managerDatabase.FieldText(rs, "new_unit_of_measurement")

    ' U: Jumlah Diterima Baru
    managerTable.SetValue tbl, rowIndex, COL_NEW_TOTAL_AMOUNT, _
        managerDatabase.FieldDouble(rs, "new_total_receive_amount")

    ' V: Jumlah OK Baru
    managerTable.SetValue tbl, rowIndex, COL_NEW_OK_AMOUNT, _
        managerDatabase.FieldDouble(rs, "new_ok_amount")

    ' W: Jumlah NOT OK Baru
    managerTable.SetValue tbl, rowIndex, COL_NEW_NOT_OK_AMOUNT, _
        managerDatabase.FieldDouble(rs, "new_not_ok_amount")

    ' X: Remarks Baru
    managerTable.SetValue tbl, rowIndex, COL_NEW_REMARKS, _
        managerDatabase.FieldText(rs, "new_remarks")

    ' Y: PIC Baru
    managerTable.SetValue tbl, rowIndex, COL_NEW_PIC, _
        managerDatabase.FieldText(rs, "new_pic")

    ' =====================================================
    ' Request and action
    ' =====================================================

    ' Z: Pengaju
    managerTable.SetValue tbl, rowIndex, COL_REQUESTED_BY, _
        managerDatabase.FieldText(rs, "update_request_by")

    ' AA: Action
    managerTable.SetValue tbl, rowIndex, COL_ACTION, _
        "SETUJUI"

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".WriteReceiveRequestUpdateRow", _
        "Gagal menuliskan satu baris pengajuan " & _
        "update penerimaan."

End Sub


Private Sub ClearReceiveRequestUpdateTable( _
    ByVal tbl As ListObject)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    If tbl Is Nothing Then
        managerError.ThrowError _
            managerTable.TableNotFound, _
            MODULE_NAME & ".ClearReceiveRequestUpdateTable", _
            "Table pengajuan update penerimaan " & _
            "tidak ditemukan."
    End If

    ' Delete existing table rows while preserving
    ' the table header and table formatting.
    managerTable.ClearRows tbl

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
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
        managerError.ThrowError _
            managerTable.TableNotFound, _
            MODULE_NAME & ".ApplyReceiveRequestUpdateFormatting", _
            "Table pengajuan update penerimaan " & _
            "tidak ditemukan."
    End If

    If Not managerTable.HasRows(tbl) Then
        Exit Sub
    End If

    ' Request date
    managerTable.FormatColumn tbl, COL_REQUEST_DATE, _
        "dd-mmm-yyyy"

    ' Old receive dates
    managerTable.FormatColumn tbl, COL_OLD_ARRIVAL_DATE, _
        "dd-mmm-yyyy"

    managerTable.FormatColumn tbl, COL_OLD_VERIFICATION_DATE, _
        "dd-mmm-yyyy"

    ' New receive dates
    managerTable.FormatColumn tbl, COL_NEW_ARRIVAL_DATE, _
        "dd-mmm-yyyy"

    managerTable.FormatColumn tbl, COL_NEW_VERIFICATION_DATE, _
        "dd-mmm-yyyy"

    ' Order amount
    managerTable.FormatColumn tbl, COL_ORDER_AMOUNT, _
        "#,##0.00"

    ' Old receive amounts
    managerTable.FormatColumn tbl, COL_OLD_TOTAL_AMOUNT, _
        "#,##0.00"

    managerTable.FormatColumn tbl, COL_OLD_OK_AMOUNT, _
        "#,##0.00"

    managerTable.FormatColumn tbl, COL_OLD_NOT_OK_AMOUNT, _
        "#,##0.00"

    ' New receive amounts
    managerTable.FormatColumn tbl, COL_NEW_TOTAL_AMOUNT, _
        "#,##0.00"

    managerTable.FormatColumn tbl, COL_NEW_OK_AMOUNT, _
        "#,##0.00"

    managerTable.FormatColumn tbl, COL_NEW_NOT_OK_AMOUNT, _
        "#,##0.00"

    ' General alignment (columns No .. SETUJUI)
    Me.Range( _
        managerTable.GetColumnRange(tbl, COL_NUMBER), _
        managerTable.GetColumnRange(tbl, COL_ACTION) _
    ).VerticalAlignment = xlCenter

    ' Format action column.
    Set actionRange = managerTable.GetColumnRange( _
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
    errInfo = managerError.CaptureError()

    Set actionRange = Nothing

    managerError.Rethrow errInfo, _
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

    If managerValidationText.IsBlank(manufacturingPlantName) Then
        managerError.ThrowError _
            sheetUpdatePenerimaan.EmptyPlantName, _
            MODULE_NAME & ".ValidateReceiveSearchParameters", _
            "Lokasi manufacturing plant wajib diisi."
    End If

    If IsEmpty(fromDateValue) Then
        managerError.ThrowError _
            sheetUpdatePenerimaan.EmptyFromDate, _
            MODULE_NAME & ".ValidateReceiveSearchParameters", _
            "Tanggal awal wajib diisi."
    End If

    If Not IsDate(fromDateValue) Then
        managerError.ThrowError _
            sheetUpdatePenerimaan.InvalidFromDate, _
            MODULE_NAME & ".ValidateReceiveSearchParameters", _
            "Tanggal awal tidak valid."
    End If

    If IsEmpty(untilDateValue) Then
        managerError.ThrowError _
            sheetUpdatePenerimaan.EmptyUntilDate, _
            MODULE_NAME & ".ValidateReceiveSearchParameters", _
            "Tanggal akhir wajib diisi."
    End If

    If Not IsDate(untilDateValue) Then
        managerError.ThrowError _
            sheetUpdatePenerimaan.InvalidUntilDate, _
            MODULE_NAME & ".ValidateReceiveSearchParameters", _
            "Tanggal akhir tidak valid."
    End If

    fromDate = DateValue(fromDateValue)
    untilDate = DateValue(untilDateValue)

    If untilDate < fromDate Then
        managerError.ThrowError _
            sheetUpdatePenerimaan.InvalidDateRange, _
            MODULE_NAME & ".ValidateReceiveSearchParameters", _
            "Tanggal akhir tidak boleh lebih kecil " & _
            "dari tanggal awal."
    End If

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ValidateReceiveSearchParameters", _
        "Filter pengajuan update penerimaan " & _
        "tidak valid."

End Sub


' Date field for a table cell: Null -> blank cell
' (managerDatabase.FieldDate would return date 0).
Private Function RecordsetDate( _
    ByVal rs As ADODB.recordset, _
    ByVal fieldName As String) As Variant

    If managerDatabase.FieldIsNull(rs, fieldName) Then
        RecordsetDate = vbNullString
    Else
        RecordsetDate = managerDatabase.FieldDate(rs, fieldName)
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
        managerError.ThrowError _
            managerTable.TableNotFound, _
            MODULE_NAME & ".HandleApproveReceiveRequestUpdate", _
            "Tabel pengajuan update penerimaan " & _
            "belum diinisialisasi."
    End If

    If targetCell Is Nothing Then
        managerError.ThrowError _
            sheetUpdatePenerimaan.EmptyApprovalCell, _
            MODULE_NAME & ".HandleApproveReceiveRequestUpdate", _
            "Cell persetujuan tidak ditemukan."
    End If

    If Not managerTable.HasRows(tbl) Then
        managerError.ThrowError _
            sheetUpdatePenerimaan.EmptyRequestUpdateTable, _
            MODULE_NAME & ".HandleApproveReceiveRequestUpdate", _
            "Tabel pengajuan update penerimaan " & _
            "tidak memiliki data."
    End If

    ' =====================================================
    ' Step 2: Determine selected table row
    ' =====================================================
    selectedRowIndex = managerTable.GetRowIndex( _
        tbl, targetCell)

    If Not managerTable.IsValidRow(tbl, selectedRowIndex) Then
        managerError.ThrowError _
            managerTable.InvalidRow, _
            MODULE_NAME & ".HandleApproveReceiveRequestUpdate", _
            "Baris pengajuan update penerimaan " & _
            "tidak valid."
    End If

    ' =====================================================
    ' Step 3: Validate action text
    ' =====================================================
    actionText = managerTable.GetText( _
        tbl, selectedRowIndex, COL_ACTION)

    If Not managerValidationText.IsSameText( _
        actionText, _
        "SETUJUI") Then

        GoTo CleanUp
    End If

    ' =====================================================
    ' Step 4: Read receive transaction UUIDs
    ' =====================================================
    originalTransactionUUID = managerTable.GetText( _
        tbl, selectedRowIndex, COL_OLD_RECEIVE_UUID)

    proposedTransactionUUID = managerTable.GetText( _
        tbl, selectedRowIndex, COL_NEW_RECEIVE_UUID)

    If managerValidationText.IsBlank(originalTransactionUUID) Then
        managerError.ThrowError _
            sheetUpdatePenerimaan.EmptyOriginalReceiveUUID, _
            MODULE_NAME & ".HandleApproveReceiveRequestUpdate", _
            "UUID transaksi receive original " & _
            "tidak boleh kosong."
    End If

    If managerValidationText.IsBlank(proposedTransactionUUID) Then
        managerError.ThrowError _
            sheetUpdatePenerimaan.EmptyProposedReceiveUUID, _
            MODULE_NAME & ".HandleApproveReceiveRequestUpdate", _
            "UUID transaksi receive proposed " & _
            "tidak boleh kosong."
    End If

    If managerValidationText.IsSameText( _
        originalTransactionUUID, _
        proposedTransactionUUID) Then

        managerError.ThrowError _
            sheetUpdatePenerimaan.SameOriginalProposedReceiveUUID, _
            MODULE_NAME & ".HandleApproveReceiveRequestUpdate", _
            "UUID transaksi receive original dan " & _
            "proposed tidak boleh sama."
    End If

    ' =====================================================
    ' Step 5: Read catalogue names
    ' =====================================================
    oldCatalogueName = managerTable.GetText( _
        tbl, selectedRowIndex, COL_OLD_CATALOGUE)

    newCatalogueName = managerTable.GetText( _
        tbl, selectedRowIndex, COL_NEW_CATALOGUE)

    If managerValidationText.IsBlank(oldCatalogueName) Then
        managerError.ThrowError _
            sheetUpdatePenerimaan.EmptyOriginalCatalogueName, _
            MODULE_NAME & ".HandleApproveReceiveRequestUpdate", _
            "Nama catalogue receive original " & _
            "tidak boleh kosong."
    End If

    If managerValidationText.IsBlank(newCatalogueName) Then
        managerError.ThrowError _
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
    oldOKAmount = managerTable.GetDouble( _
        tbl, selectedRowIndex, COL_OLD_OK_AMOUNT)

    newOKAmount = managerTable.GetDouble( _
        tbl, selectedRowIndex, COL_NEW_OK_AMOUNT)

    If managerValidationNumeric.IsNegative(oldOKAmount) Then
        managerError.ThrowError _
            sheetUpdatePenerimaan.NegativeOriginalOkAmount, _
            MODULE_NAME & ".HandleApproveReceiveRequestUpdate", _
            "Jumlah OK receive original tidak boleh " & _
            "kurang dari nol."
    End If

    If managerValidationNumeric.IsNegative(newOKAmount) Then
        managerError.ThrowError _
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
    oldManufacturingPlantName = managerForm.GetCellText( _
        Me, CELL_MANUFACTURING_PLANT)

    newManufacturingPlantName = _
        oldManufacturingPlantName

    If managerValidationText.IsBlank(oldManufacturingPlantName) Then
        managerError.ThrowError _
            sheetUpdatePenerimaan.EmptyOriginalPlantName, _
            MODULE_NAME & ".HandleApproveReceiveRequestUpdate", _
            "Nama manufacturing plant original " & _
            "tidak boleh kosong."
    End If

    If managerValidationText.IsBlank(newManufacturingPlantName) Then
        managerError.ThrowError _
            sheetUpdatePenerimaan.EmptyProposedPlantName, _
            MODULE_NAME & ".HandleApproveReceiveRequestUpdate", _
            "Nama manufacturing plant proposed " & _
            "tidak boleh kosong."
    End If

    ' =====================================================
    ' Step 8: Determine current user
    ' =====================================================
    currentUser = coreHelper.TrimText( _
        coreConfig.currentUser())

    If managerValidationText.IsBlank(currentUser) Then
        currentUser = coreHelper.TrimText( _
            Environ$("Username"))
    End If

    If managerValidationText.IsBlank(currentUser) Then
        managerError.ThrowError _
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

    If Not managerForm.Confirm( _
        confirmationMessage, _
        "Konfirmasi Persetujuan Update Penerimaan") Then

        GoTo CleanUp
    End If

    ' =====================================================
    ' Step 10: Set temporary processing status
    ' =====================================================
    managerTable.SetValue tbl, selectedRowIndex, COL_ACTION, _
        "MEMPROSES..."

    isProcessing = True

    With managerTable.GetCell(tbl, selectedRowIndex, COL_ACTION)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

        .Font.Bold = True
        .Font.Color = RGB(128, 128, 128)

        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(230, 230, 230)
    End With

    managerForm.SetStatus _
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
    managerTable.SetValue tbl, selectedRowIndex, COL_ACTION, _
        "DISETUJUI"

    isProcessing = False

    With managerTable.GetCell(tbl, selectedRowIndex, COL_ACTION)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)

        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(0, 128, 0)
    End With

    managerForm.ClearStatus

    managerForm.ShowInfo _
        "Pengajuan update penerimaan " & _
        "berhasil disetujui.", _
        "Persetujuan Berhasil"

    ' =====================================================
    ' Step 13: Refresh pending update requests
    ' =====================================================
    CariPengajuanUpdateReceive

CleanUp:
    managerForm.ClearStatus

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    ' Restore action cell only if it had entered the
    ' processing state.
    If isProcessing Then

        managerTable.SetValue tbl, selectedRowIndex, COL_ACTION, _
            "SETUJUI"

        With managerTable.GetCell(tbl, selectedRowIndex, COL_ACTION)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter

            .Font.Bold = True
            .Font.Color = RGB(0, 97, 0)

            .Interior.Pattern = xlSolid
            .Interior.Color = RGB(198, 239, 206)
        End With

    End If

    managerForm.ClearStatus

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".HandleApproveReceiveRequestUpdate", _
        "Gagal menyetujui pengajuan " & _
        "update penerimaan."

End Sub
