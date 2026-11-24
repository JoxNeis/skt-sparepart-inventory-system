Attribute VB_Name = "sUpdateOrder"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Const MODULE_NAME As String = "SHEET_UPDATE_ORDER"

' Name of the Excel table on this worksheet (A11:X..).
Private Const TABLE_REQUEST_UPDATE_ORDER As String = _
    "TableRequestUpdateOrder"

Private Const CELL_MANUFACTURING_PLANT As String = "B4"
Private Const CELL_FROM_DATE As String = "B7"
Private Const CELL_UNTIL_DATE As String = "B8"

' =========================================================
' Table columns (TableRequestUpdateOrder)
'
' TableManager addresses columns by header text. These
' names must exactly match the Excel table headers. The
' comment shows the column position in the table.
' =========================================================
Private Const COL_NUMBER As String = "No"                               ' 1  A
Private Const COL_REQUEST_DATE As String = "Tanggal Pengajuan"          ' 2  B

' Old order columns
Private Const COL_OLD_UUID As String = "UUID LAMA"                      ' 3  C
Private Const COL_OLD_PO_ID As String = "PO ID LAMA"                    ' 4  D
Private Const COL_OLD_CATALOGUE As String = "NAMA BARANG LAMA"          ' 5  E
Private Const COL_OLD_AMOUNT As String = "JUMLAH ORDER LAMA"            ' 6  F
Private Const COL_OLD_UNIT As String = "SATUAN LAMA"                    ' 7  G
Private Const COL_OLD_VENDOR As String = "VENDOR LAMA"                  ' 8  H
Private Const COL_OLD_ORDER_DATE As String = "TANGGAL ORDER LAMA"       ' 9  I
Private Const COL_OLD_REMARKS As String = "REMARKS LAMA"                ' 10 J
Private Const COL_OLD_PIC As String = "PIC LAMA"                        ' 11 K

' New order columns
Private Const COL_NEW_UUID As String = "UUID BARU"                      ' 12 L
Private Const COL_NEW_PO_ID As String = "PO ID BARU"                    ' 13 M
Private Const COL_NEW_CATALOGUE As String = "NAMA BARANG BARU"          ' 14 N
Private Const COL_NEW_AMOUNT As String = "JUMLAH ORDER BARU"            ' 15 O
Private Const COL_NEW_UNIT As String = "SATUAN BARU"                    ' 16 P
Private Const COL_NEW_VENDOR As String = "VENDOR BARU"                  ' 17 Q
Private Const COL_NEW_ORDER_DATE As String = "TANGGAL ORDER BARU"       ' 18 R
Private Const COL_NEW_REMARKS As String = "REMARKS BARU"                ' 19 S
Private Const COL_NEW_PIC As String = "PIC BARU"                        ' 20 T

' Requester and action
Private Const COL_REQUESTED_BY As String = "PENGAJU"                    ' 21 U
Private Const COL_ACTION As String = "SETUJUI"                          ' 22 V

' Internal update request ID (hidden column, header text
' as in the workbook). 24 X "Column2" is not used.
Private Const COL_UPDATE_REQUEST_ID As String = "Column1"               ' 23 W

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    On Error GoTo ErrorHandler

    Dim tbl As ListObject

    Dim errInfo As ErrorInfo

    If FormManager.IsInRange( _
        Target, _
        Me.Range(CELL_FROM_DATE & ":" & CELL_UNTIL_DATE)) Then

        Application.EnableEvents = False

        FormManager.ShowDatePicker Target

        GoTo CleanUp

    End If

    ' The selected cell is not in the approval column.
    Set tbl = TableManager.GetTable( _
        Me, TABLE_REQUEST_UPDATE_ORDER)

    If Not TableManager.IsInColumn( _
        Target, tbl, COL_ACTION) Then

        Exit Sub
    End If

    ' Prevent the event from running recursively.
    Application.EnableEvents = False

    HandleApproveOrderRequestUpdate _
        tbl:=tbl, _
        targetCell:=Target

CleanUp:
    Set tbl = Nothing

    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    Set tbl = Nothing

    Application.EnableEvents = True

    ErrorManager.ShowError errInfo, _
        "Gagal memproses persetujuan update order."

End Sub

Public Sub CariPengajuanUpdateOrder()

    On Error GoTo ErrorHandler

    Dim tbl As ListObject
    Dim rs As ADODB.recordset

    Dim manufacturingPlantName As String
    Dim fromDate As Date
    Dim untilDate As Date

    Dim rowIndex As Long
    Dim rowNumber As Long
    Dim resultCount As Long

    Dim state As FormState
    Dim errInfo As ErrorInfo

    state = FormManager.BeginBusy( _
        "Mengambil data pengajuan update order...")

    ' =====================================================
    ' Read and validate filter values
    ' =====================================================
    manufacturingPlantName = FormManager.GetCellText( _
        Me, CELL_MANUFACTURING_PLANT)

    ValidateSearchParameters _
        manufacturingPlantName:=manufacturingPlantName, _
        fromDateValue:=FormManager.GetCellValue(Me, CELL_FROM_DATE), _
        untilDateValue:=FormManager.GetCellValue(Me, CELL_UNTIL_DATE)

    fromDate = DateValue( _
        FormManager.GetCellValue(Me, CELL_FROM_DATE))

    untilDate = DateValue( _
        FormManager.GetCellValue(Me, CELL_UNTIL_DATE))

    ' =====================================================
    ' Clear previous result
    ' =====================================================
    Set tbl = TableManager.GetTable( _
        Me, TABLE_REQUEST_UPDATE_ORDER)

    ClearOrderRequestUpdateData tbl

    ' =====================================================
    ' Retrieve data from the service
    ' =====================================================
    Set rs = serviceTransactionRequestUpdate.SelectOrderRequestUpdate( _
        manufacturingPlantName:=manufacturingPlantName, _
        fromDate:=fromDate, _
        untilDate:=untilDate)

    If rs Is Nothing Then
        ErrorManager.ThrowError _
            DatabaseManager.EmptyRecordset, _
            MODULE_NAME & ".CariPengajuanUpdateOrder", _
            "DAO tidak mengembalikan recordset."
    End If

    rowNumber = 1

    ' =====================================================
    ' Write records to the table
    ' =====================================================
    Do While Not rs.EOF

        rowIndex = TableManager.AddRow(tbl)

        WriteOrderRequestUpdateRow _
            tbl:=tbl, _
            rowIndex:=rowIndex, _
            rs:=rs, _
            rowNumber:=rowNumber

        rowNumber = rowNumber + 1

        rs.MoveNext
    Loop

    resultCount = rowNumber - 1

    ApplyOrderRequestUpdateFormatting tbl

    ' =====================================================
    ' Display result
    ' =====================================================
    If resultCount = 0 Then
        FormManager.ShowInfo _
            "Tidak ada pengajuan update order " & _
            "yang ditemukan.", _
            "Pengajuan Update Order"
    Else
        FormManager.ShowInfo _
            CStr(resultCount) & _
            " pengajuan update order ditemukan.", _
            "Pengajuan Update Order"
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
        "Gagal menampilkan pengajuan update order."

End Sub


Private Sub WriteOrderRequestUpdateRow( _
    ByVal tbl As ListObject, _
    ByVal rowIndex As Long, _
    ByVal rs As ADODB.recordset, _
    ByVal rowNumber As Long)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    ' =====================================================
    ' General information
    ' =====================================================

    ' No
    TableManager.SetValue tbl, rowIndex, COL_NUMBER, _
        rowNumber

    ' Tanggal Pengajuan
    TableManager.SetValue tbl, rowIndex, COL_REQUEST_DATE, _
        DatabaseManager.FieldDate(rs, "update_request_date")

    ' =====================================================
    ' Old order
    ' =====================================================

    ' Old UUID
    TableManager.SetValue tbl, rowIndex, COL_OLD_UUID, _
        DatabaseManager.FieldText(rs, "old_transaction_uuid")

    ' Old PO ID
    '
    ' The query currently returns old_order_id,
    ' not the actual old PO ID.
    TableManager.SetValue tbl, rowIndex, COL_OLD_PO_ID, _
        DatabaseManager.FieldText(rs, "old_order_id")

    ' Old catalogue
    TableManager.SetValue tbl, rowIndex, COL_OLD_CATALOGUE, _
        DatabaseManager.FieldText(rs, "old_catalogue_name")

    ' Old amount
    TableManager.SetValue tbl, rowIndex, COL_OLD_AMOUNT, _
        DatabaseManager.FieldDouble(rs, "old_amount")

    ' Old unit of measurement
    TableManager.SetValue tbl, rowIndex, COL_OLD_UNIT, _
        DatabaseManager.FieldText(rs, "old_unit_of_measurement")

    ' Old vendor
    TableManager.SetValue tbl, rowIndex, COL_OLD_VENDOR, _
        DatabaseManager.FieldText(rs, "old_vendor_name")

    ' Old order date
    TableManager.SetValue tbl, rowIndex, COL_OLD_ORDER_DATE, _
        RecordsetDate(rs, "old_order_date")

    ' Old remarks
    TableManager.SetValue tbl, rowIndex, COL_OLD_REMARKS, _
        DatabaseManager.FieldText(rs, "old_remarks")

    ' Old person in charge
    TableManager.SetValue tbl, rowIndex, COL_OLD_PIC, _
        DatabaseManager.FieldText(rs, "old_pic")

    ' =====================================================
    ' New order
    ' =====================================================

    ' New UUID
    TableManager.SetValue tbl, rowIndex, COL_NEW_UUID, _
        DatabaseManager.FieldText(rs, "new_transaction_uuid")

    ' New PO ID
    '
    ' The query currently returns new_order_id,
    ' not the actual new PO ID.
    TableManager.SetValue tbl, rowIndex, COL_NEW_PO_ID, _
        DatabaseManager.FieldText(rs, "new_order_id")

    ' New catalogue
    TableManager.SetValue tbl, rowIndex, COL_NEW_CATALOGUE, _
        DatabaseManager.FieldText(rs, "new_catalogue_name")

    ' New amount
    TableManager.SetValue tbl, rowIndex, COL_NEW_AMOUNT, _
        DatabaseManager.FieldDouble(rs, "new_amount")

    ' New unit of measurement
    TableManager.SetValue tbl, rowIndex, COL_NEW_UNIT, _
        DatabaseManager.FieldText(rs, "new_unit_of_measurement")

    ' New vendor
    TableManager.SetValue tbl, rowIndex, COL_NEW_VENDOR, _
        DatabaseManager.FieldText(rs, "new_vendor_name")

    ' New order date
    TableManager.SetValue tbl, rowIndex, COL_NEW_ORDER_DATE, _
        RecordsetDate(rs, "new_order_date")

    ' New remarks
    TableManager.SetValue tbl, rowIndex, COL_NEW_REMARKS, _
        DatabaseManager.FieldText(rs, "new_remarks")

    ' New person in charge
    TableManager.SetValue tbl, rowIndex, COL_NEW_PIC, _
        DatabaseManager.FieldText(rs, "new_pic")

    ' =====================================================
    ' Status and action
    ' =====================================================

    ' Requester
    TableManager.SetValue tbl, rowIndex, COL_REQUESTED_BY, _
        DatabaseManager.FieldText(rs, "update_request_by")

    ' Action
    TableManager.SetValue tbl, rowIndex, COL_ACTION, _
        "SETUJUI"

    ' Internal update request ID (hidden column)
    TableManager.SetValue tbl, rowIndex, COL_UPDATE_REQUEST_ID, _
        DatabaseManager.FieldText(rs, "update_request_id")

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".WriteOrderRequestUpdateRow", _
        "Gagal menuliskan pengajuan update order " & _
        "ke worksheet."

End Sub


Private Sub ValidateSearchParameters( _
    ByVal manufacturingPlantName As String, _
    ByVal fromDateValue As Variant, _
    ByVal untilDateValue As Variant)

    On Error GoTo ErrorHandler

    Dim fromDate As Date
    Dim untilDate As Date

    Dim errInfo As ErrorInfo

    If ValidationText.IsBlank(manufacturingPlantName) Then
        ErrorManager.ThrowError _
            sheetUpdateOrder.EmptyPlantName, _
            MODULE_NAME & ".ValidateSearchParameters", _
            "Lokasi manufacturing plant wajib diisi."
    End If

    If IsEmpty(fromDateValue) Then
        ErrorManager.ThrowError _
            sheetUpdateOrder.EmptyFromDate, _
            MODULE_NAME & ".ValidateSearchParameters", _
            "Tanggal awal wajib diisi."
    End If

    If Not IsDate(fromDateValue) Then
        ErrorManager.ThrowError _
            sheetUpdateOrder.InvalidFromDate, _
            MODULE_NAME & ".ValidateSearchParameters", _
            "Tanggal awal tidak valid."
    End If

    If IsEmpty(untilDateValue) Then
        ErrorManager.ThrowError _
            sheetUpdateOrder.EmptyUntilDate, _
            MODULE_NAME & ".ValidateSearchParameters", _
            "Tanggal akhir wajib diisi."
    End If

    If Not IsDate(untilDateValue) Then
        ErrorManager.ThrowError _
            sheetUpdateOrder.InvalidUntilDate, _
            MODULE_NAME & ".ValidateSearchParameters", _
            "Tanggal akhir tidak valid."
    End If

    fromDate = DateValue(fromDateValue)
    untilDate = DateValue(untilDateValue)

    If untilDate < fromDate Then
        ErrorManager.ThrowError _
            sheetUpdateOrder.InvalidDateRange, _
            MODULE_NAME & ".ValidateSearchParameters", _
            "Tanggal akhir tidak boleh lebih kecil " & _
            "dari tanggal awal."
    End If

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".ValidateSearchParameters", _
        "Filter pengajuan update order tidak valid."

End Sub


Private Sub ClearOrderRequestUpdateData( _
    ByVal tbl As ListObject)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    ' Delete existing table rows while preserving the
    ' table header and formatting.
    TableManager.ClearRows tbl

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".ClearOrderRequestUpdateData", _
        "Gagal membersihkan data pengajuan " & _
        "update order sebelumnya."

End Sub


Private Sub ApplyOrderRequestUpdateFormatting( _
    ByVal tbl As ListObject)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    If TableManager.HasRows(tbl) Then

        ' Request date
        TableManager.FormatColumn tbl, COL_REQUEST_DATE, _
            "dd-mmm-yyyy"

        ' Old order date
        TableManager.FormatColumn tbl, COL_OLD_ORDER_DATE, _
            "dd-mmm-yyyy"

        ' New order date
        TableManager.FormatColumn tbl, COL_NEW_ORDER_DATE, _
            "dd-mmm-yyyy"

        ' Old amount
        TableManager.FormatColumn tbl, COL_OLD_AMOUNT, _
            "#,##0.00"

        ' New amount
        TableManager.FormatColumn tbl, COL_NEW_AMOUNT, _
            "#,##0.00"

        ' General alignment (columns No .. SETUJUI)
        Me.Range( _
            TableManager.GetColumnRange(tbl, COL_NUMBER), _
            TableManager.GetColumnRange(tbl, COL_ACTION) _
        ).VerticalAlignment = xlCenter

        ' Keep the update request ID hidden.
        TableManager.GetColumnRange( _
            tbl, COL_UPDATE_REQUEST_ID _
        ).EntireColumn.Hidden = True

    Else

        ' Keep the update request ID hidden (empty table:
        ' there is no data range, use the sheet column).
        Me.Columns( _
            tbl.Range.Column + _
            TableManager.GetColumnIndex(tbl, COL_UPDATE_REQUEST_ID) - 1 _
        ).Hidden = True

    End If

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".ApplyOrderRequestUpdateFormatting", _
        "Gagal menerapkan format pengajuan update order."

End Sub


' Date field for a table cell: Null or a value that is
' not a date -> blank cell (DatabaseManager.FieldDate would
' return date 0 for Null and raise for a non-date value).
Private Function RecordsetDate( _
    ByVal rs As ADODB.recordset, _
    ByVal fieldName As String) As Variant

    If DatabaseManager.FieldIsNull(rs, fieldName) Then
        RecordsetDate = vbNullString
    ElseIf IsDate(DatabaseManager.FieldValue(rs, fieldName)) Then
        RecordsetDate = DatabaseManager.FieldDate(rs, fieldName)
    Else
        RecordsetDate = vbNullString
    End If

End Function

Private Sub HandleApproveOrderRequestUpdate( _
    ByVal tbl As ListObject, _
    ByVal targetCell As Range)

    On Error GoTo ErrorHandler

    Dim targetRow As Long
    Dim selectedRowIndex As Long

    Dim originalTransactionUUID As String
    Dim proposedTransactionUUID As String
    Dim currentUser As String
    Dim actionText As String

    Dim errInfo As ErrorInfo

    ' Worksheet row, used in the messages below.
    targetRow = targetCell.Row

    ' =====================================================
    ' Validate the selected row
    ' =====================================================

    ' Table row of the clicked cell (0 = outside the table).
    selectedRowIndex = TableManager.GetRowIndex( _
        tbl, targetCell)

    If Not TableManager.IsValidRow(tbl, selectedRowIndex) Then
        Exit Sub
    End If

    actionText = TableManager.GetText( _
        tbl, selectedRowIndex, COL_ACTION)

    If Not ValidationText.IsSameText( _
        actionText, _
        "SETUJUI") Then

        Exit Sub
    End If

    ' Original transaction UUID
    originalTransactionUUID = TableManager.GetText( _
        tbl, selectedRowIndex, COL_OLD_UUID)

    ' Proposed transaction UUID
    proposedTransactionUUID = TableManager.GetText( _
        tbl, selectedRowIndex, COL_NEW_UUID)

    ' Use the Windows username.
    '
    ' Replace this with the logged-in application user
    ' if your workbook has its own authentication system.
    currentUser = pHelper.TrimText(Environ$("Username"))

    If ValidationText.IsBlank(originalTransactionUUID) Then
        ErrorManager.ThrowError _
            sheetUpdateOrder.EmptyOriginalOrderUUID, _
            MODULE_NAME & ".HandleApproveOrderRequestUpdate", _
            "UUID transaksi original pada baris " & _
            CStr(targetRow) & " tidak boleh kosong."
    End If

    If ValidationText.IsBlank(proposedTransactionUUID) Then
        ErrorManager.ThrowError _
            sheetUpdateOrder.EmptyProposedOrderUUID, _
            MODULE_NAME & ".HandleApproveOrderRequestUpdate", _
            "UUID transaksi proposed pada baris " & _
            CStr(targetRow) & " tidak boleh kosong."
    End If

    If ValidationText.IsSameText( _
        originalTransactionUUID, _
        proposedTransactionUUID) Then

        ErrorManager.ThrowError _
            sheetUpdateOrder.SameOriginalProposedOrderUUID, _
            MODULE_NAME & ".HandleApproveOrderRequestUpdate", _
            "UUID transaksi original dan proposed " & _
            "tidak boleh sama."
    End If

    If ValidationText.IsBlank(currentUser) Then
        ErrorManager.ThrowError _
            sheetUpdateOrder.EmptyApproverUser, _
            MODULE_NAME & ".HandleApproveOrderRequestUpdate", _
            "User yang menyetujui tidak dapat ditentukan."
    End If

    ' =====================================================
    ' Ask for confirmation
    ' =====================================================
    If Not FormManager.Confirm( _
        "Apakah Anda yakin ingin menyetujui " & _
        "pengajuan update order ini?" & vbCrLf & _
        vbCrLf & _
        "Order lama:" & vbCrLf & _
        originalTransactionUUID & vbCrLf & _
        vbCrLf & _
        "Order baru:" & vbCrLf & _
        proposedTransactionUUID & vbCrLf & _
        vbCrLf & _
        "Tindakan ini akan:" & vbCrLf & _
        "1. Menyetujui pengajuan update order." & vbCrLf & _
        "2. Membalik transaksi original." & vbCrLf & _
        "3. Mengaktifkan transaksi proposed.", _
        "Konfirmasi Persetujuan") Then

        Exit Sub
    End If

    ' =====================================================
    ' Temporarily disable the action cell
    ' =====================================================
    TableManager.SetValue tbl, selectedRowIndex, COL_ACTION, _
        "MEMPROSES..."

    TableManager.GetCell( _
        tbl, selectedRowIndex, COL_ACTION _
    ).Font.Color = RGB(128, 128, 128)

    FormManager.SetStatus _
        "Menyetujui pengajuan update order...", True

    ' =====================================================
    ' Execute approval
    ' =====================================================
    serviceTransactionRequestUpdate.ApproveRequestUpdate _
        originalTransactionUUID:=originalTransactionUUID, _
        proposedTransactionUUID:=proposedTransactionUUID, _
        user:=currentUser, _
        updateOrder:=True

    ' =====================================================
    ' Approval succeeded
    ' =====================================================
    TableManager.SetValue tbl, selectedRowIndex, COL_ACTION, _
        "DISETUJUI"

    TableManager.GetCell( _
        tbl, selectedRowIndex, COL_ACTION _
    ).Font.Color = RGB(0, 128, 0)

    FormManager.ClearStatus

    FormManager.ShowInfo _
        "Pengajuan update order berhasil disetujui.", _
        "Persetujuan Berhasil"

    ' Refresh the data so the approved request disappears
    ' if the query only returns pending requests.
    CariPengajuanUpdateOrder

CleanUp:
    FormManager.ClearStatus
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ' Restore the action cell if approval failed.
    If TableManager.IsValidRow(tbl, selectedRowIndex) Then
        TableManager.SetValue tbl, selectedRowIndex, COL_ACTION, _
            "SETUJUI"

        TableManager.GetCell( _
            tbl, selectedRowIndex, COL_ACTION _
        ).Font.ColorIndex = xlAutomatic
    End If

    FormManager.ClearStatus

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".HandleApproveOrderRequestUpdate", _
        "Gagal menyetujui pengajuan update order."

End Sub
