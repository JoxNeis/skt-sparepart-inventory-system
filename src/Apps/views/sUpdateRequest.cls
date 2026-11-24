Attribute VB_Name = "sUpdateRequest"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Const MODULE_NAME As String = _
    "SHEET_UPDATE_REQUEST"

' Name of the Excel table on this worksheet (A11:W..).
Private Const TABLE_REQUEST_UPDATE_REQUEST As String = _
    "TableRequestUpdateReceive13"

Private Const CELL_MANUFACTURING_PLANT As String = "B4"
Private Const CELL_FROM_DATE As String = "B7"
Private Const CELL_UNTIL_DATE As String = "B8"

' =========================================================
' Table columns (TableRequestUpdateReceive13)
'
' TableManager addresses columns by header text. These
' names must exactly match the Excel table headers. The
' comment shows the column position in the table.
' =========================================================
Private Const COL_NUMBER As String = "No"                               ' 1  A
Private Const COL_REQUEST_DATE As String = "Tanggal Pengajuan"          ' 2  B

' Old request columns
Private Const COL_OLD_UUID As String = "UUID LAMA"                      ' 3  C
Private Const COL_OLD_MOBILE_POST As String = "MOBILE POST LAMA"        ' 4  D
Private Const COL_OLD_CATALOGUE As String = "NAMA BARANG LAMA"          ' 5  E
Private Const COL_OLD_AMOUNT As String = "JUMLAH BARANG LAMA"           ' 6  F
Private Const COL_OLD_UNIT As String = "SATUAN LAMA"                    ' 7  G
Private Const COL_OLD_REQUEST_DATE As String = "TANGGAL ORDER LAMA"     ' 8  H
Private Const COL_OLD_REMARKS As String = "REMARKS LAMA"                ' 9  I
Private Const COL_OLD_PIC As String = "PIC LAMA"                        ' 10 J

' New request columns
Private Const COL_NEW_UUID As String = "UUID BARU"                      ' 11 K
Private Const COL_NEW_MOBILE_POST As String = "MOBILE POST BARU"        ' 12 L
Private Const COL_NEW_CATALOGUE As String = "NAMA BARANG BARU"          ' 13 M
Private Const COL_NEW_AMOUNT As String = "JUMLAH BARANG BARU"           ' 14 N
Private Const COL_NEW_UNIT As String = "SATUAN BARU"                    ' 15 O
Private Const COL_NEW_REQUEST_DATE As String = "TANGGAL ORDER BARU"     ' 16 P
Private Const COL_NEW_REMARKS As String = "REMARKS BARU"                ' 17 Q
Private Const COL_NEW_PIC As String = "PIC LAMA8"                       ' 18 R (header text as in the workbook)

' 19 S "NAMA PENGAJU" is a calculated column (XLOOKUP of
' PENGAJU in TableUser). It is not written by this code.
Private Const COL_REQUESTED_BY As String = "PENGAJU"                    ' 20 T
Private Const COL_ACTION As String = "SETUJUI"                          ' 21 U


' =========================================================
' Worksheet selection event
' =========================================================
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
    ' SETUJUI column
    ' =====================================================
    Set tbl = TableManager.GetTable( _
        Me, TABLE_REQUEST_UPDATE_REQUEST)

    If Not TableManager.HasRows(tbl) Then
        GoTo CleanUp
    End If

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

    HandleApproveRequestUpdate _
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
        "update request."

End Sub


' =========================================================
' Search button procedure
' =========================================================
Public Sub CariPengajuanUpdateRequest()

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
        "Mengambil data pengajuan update request...")

    ' =====================================================
    ' Step 1: Read and validate filters
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
    ' Step 2: Clear previous results
    ' =====================================================
    Set tbl = TableManager.GetTable( _
        Me, TABLE_REQUEST_UPDATE_REQUEST)

    ClearRequestUpdateData tbl

    ' =====================================================
    ' Step 3: Call service function
    ' =====================================================
    Set rs = _
        serviceTransactionRequestUpdate.SelectRequestRequestUpdate( _
            manufacturingPlantName:= _
                manufacturingPlantName, _
            fromDate:=fromDate, _
            untilDate:=untilDate)

    If rs Is Nothing Then
        ErrorManager.ThrowError _
            DatabaseManager.EmptyRecordset, _
            MODULE_NAME & ".CariPengajuanUpdateRequest", _
            "DAO tidak mengembalikan recordset."
    End If

    If rs.State <> adStateOpen Then
        ErrorManager.ThrowError _
            DatabaseManager.RecordsetNotOpen, _
            MODULE_NAME & ".CariPengajuanUpdateRequest", _
            "Recordset pengajuan update request " & _
            "tidak terbuka."
    End If

    ' =====================================================
    ' Step 4: Write records to the table
    ' =====================================================
    rowNumber = 1

    Do While Not rs.EOF

        rowIndex = TableManager.AddRow(tbl)

        WriteRequestUpdateRow _
            tbl:=tbl, _
            rowIndex:=rowIndex, _
            rs:=rs, _
            rowNumber:=rowNumber

        rowNumber = rowNumber + 1

        rs.MoveNext
    Loop

    resultCount = rowNumber - 1

    ' =====================================================
    ' Step 5: Apply formatting
    ' =====================================================
    ApplyRequestUpdateFormatting tbl

    ' =====================================================
    ' Step 6: Display result
    ' =====================================================
    If resultCount = 0 Then

        FormManager.ShowInfo _
            "Tidak ada pengajuan update request " & _
            "yang ditemukan.", _
            "Pengajuan Update Request"

    Else

        FormManager.ShowInfo _
            CStr(resultCount) & _
            " pengajuan update request ditemukan.", _
            "Pengajuan Update Request"

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
        "Gagal menampilkan pengajuan update request."

End Sub


' =========================================================
' Write a recordset row to the table
' =========================================================
Private Sub WriteRequestUpdateRow( _
    ByVal tbl As ListObject, _
    ByVal rowIndex As Long, _
    ByVal rs As ADODB.recordset, _
    ByVal rowNumber As Long)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    ' =====================================================
    ' General request information
    ' =====================================================

    ' A: Number
    TableManager.SetValue tbl, rowIndex, COL_NUMBER, _
        rowNumber

    ' B: Update request date
    TableManager.SetValue tbl, rowIndex, COL_REQUEST_DATE, _
        RecordsetDate(rs, "update_request_date")

    ' =====================================================
    ' Old request
    ' =====================================================

    ' C: Old transaction UUID
    TableManager.SetValue tbl, rowIndex, COL_OLD_UUID, _
        DatabaseManager.FieldText(rs, "old_transaction_uuid")

    ' D: Old mobile post
    TableManager.SetValue tbl, rowIndex, COL_OLD_MOBILE_POST, _
        DatabaseManager.FieldText(rs, "old_mobile_post_name")

    ' E: Old catalogue
    TableManager.SetValue tbl, rowIndex, COL_OLD_CATALOGUE, _
        DatabaseManager.FieldText(rs, "old_catalogue_name")

    ' F: Old amount
    TableManager.SetValue tbl, rowIndex, COL_OLD_AMOUNT, _
        DatabaseManager.FieldDouble(rs, "old_amount")

    ' G: Old unit
    TableManager.SetValue tbl, rowIndex, COL_OLD_UNIT, _
        DatabaseManager.FieldText(rs, "old_unit_of_measurement")

    ' H: Old request date
    TableManager.SetValue tbl, rowIndex, COL_OLD_REQUEST_DATE, _
        RecordsetDate(rs, "old_request_date")

    ' I: Old remarks
    TableManager.SetValue tbl, rowIndex, COL_OLD_REMARKS, _
        DatabaseManager.FieldText(rs, "old_remarks")

    ' J: Old PIC
    TableManager.SetValue tbl, rowIndex, COL_OLD_PIC, _
        DatabaseManager.FieldText(rs, "old_pic")

    ' =====================================================
    ' New request
    ' =====================================================

    ' K: New transaction UUID
    TableManager.SetValue tbl, rowIndex, COL_NEW_UUID, _
        DatabaseManager.FieldText(rs, "new_transaction_uuid")

    ' L: New mobile post
    TableManager.SetValue tbl, rowIndex, COL_NEW_MOBILE_POST, _
        DatabaseManager.FieldText(rs, "new_mobile_post_name")

    ' M: New catalogue
    TableManager.SetValue tbl, rowIndex, COL_NEW_CATALOGUE, _
        DatabaseManager.FieldText(rs, "new_catalogue_name")

    ' N: New amount
    TableManager.SetValue tbl, rowIndex, COL_NEW_AMOUNT, _
        DatabaseManager.FieldDouble(rs, "new_amount")

    ' O: New unit
    TableManager.SetValue tbl, rowIndex, COL_NEW_UNIT, _
        DatabaseManager.FieldText(rs, "new_unit_of_measurement")

    ' P: New request date
    TableManager.SetValue tbl, rowIndex, COL_NEW_REQUEST_DATE, _
        RecordsetDate(rs, "new_request_date")

    ' Q: New remarks
    TableManager.SetValue tbl, rowIndex, COL_NEW_REMARKS, _
        DatabaseManager.FieldText(rs, "new_remarks")

    ' R: New PIC
    TableManager.SetValue tbl, rowIndex, COL_NEW_PIC, _
        DatabaseManager.FieldText(rs, "new_pic")

    ' =====================================================
    ' Requester and action
    ' =====================================================

    ' T: Requested by (S "NAMA PENGAJU" looks this up)
    TableManager.SetValue tbl, rowIndex, COL_REQUESTED_BY, _
        DatabaseManager.FieldText(rs, "update_request_by")

    ' U: Action
    TableManager.SetValue tbl, rowIndex, COL_ACTION, _
        "SETUJUI"

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".WriteRequestUpdateRow", _
        "Gagal menuliskan pengajuan update request " & _
        "ke worksheet."

End Sub


' =========================================================
' Approve selected update request
' =========================================================
Private Sub HandleApproveRequestUpdate( _
    ByVal tbl As ListObject, _
    ByVal targetCell As Range)

    On Error GoTo ErrorHandler

    Dim targetRow As Long
    Dim selectedRowIndex As Long

    Dim originalTransactionUUID As String
    Dim proposedTransactionUUID As String
    Dim currentUser As String
    Dim actionText As String

    Dim oldCatalogueName As String
    Dim newCatalogueName As String

    Dim oldManufacturingPlantName As String
    Dim newManufacturingPlantName As String

    Dim confirmationMessage As String

    ' True while the action cell shows "MEMPROSES...".
    Dim actionCellChanged As Boolean

    Dim errInfo As ErrorInfo

    actionCellChanged = False

    If tbl Is Nothing Then
        ErrorManager.ThrowError _
            TableManager.TableNotFound, _
            MODULE_NAME & ".HandleApproveRequestUpdate", _
            "Tabel pengajuan update request " & _
            "belum diinisialisasi."
    End If

    If targetCell Is Nothing Then
        ErrorManager.ThrowError _
            sheetUpdateRequest.EmptyApprovalCell, _
            MODULE_NAME & ".HandleApproveRequestUpdate", _
            "Cell persetujuan tidak ditemukan."
    End If

    ' Worksheet row, used in the messages below.
    targetRow = targetCell.Row

    ' Table row of the clicked cell (0 = outside the table).
    selectedRowIndex = TableManager.GetRowIndex( _
        tbl, targetCell)

    If Not TableManager.IsValidRow(tbl, selectedRowIndex) Then
        GoTo CleanUp
    End If

    actionText = TableManager.GetText( _
        tbl, selectedRowIndex, COL_ACTION)

    If Not ValidationText.IsSameText( _
        actionText, _
        "SETUJUI") Then

        GoTo CleanUp
    End If

    ' =====================================================
    ' Read transaction UUIDs
    ' =====================================================
    originalTransactionUUID = TableManager.GetText( _
        tbl, selectedRowIndex, COL_OLD_UUID)

    proposedTransactionUUID = TableManager.GetText( _
        tbl, selectedRowIndex, COL_NEW_UUID)

    If ValidationText.IsBlank(originalTransactionUUID) Then
        ErrorManager.ThrowError _
            sheetUpdateRequest.EmptyOriginalRequestUUID, _
            MODULE_NAME & ".HandleApproveRequestUpdate", _
            "UUID transaksi request original pada " & _
            "baris " & CStr(targetRow) & _
            " tidak boleh kosong."
    End If

    If ValidationText.IsBlank(proposedTransactionUUID) Then
        ErrorManager.ThrowError _
            sheetUpdateRequest.EmptyProposedRequestUUID, _
            MODULE_NAME & ".HandleApproveRequestUpdate", _
            "UUID transaksi request proposed pada " & _
            "baris " & CStr(targetRow) & _
            " tidak boleh kosong."
    End If

    If ValidationText.IsSameText( _
        originalTransactionUUID, _
        proposedTransactionUUID) Then

        ErrorManager.ThrowError _
            sheetUpdateRequest.SameOriginalProposedRequestUUID, _
            MODULE_NAME & ".HandleApproveRequestUpdate", _
            "UUID transaksi original dan proposed " & _
            "tidak boleh sama."
    End If

    ' =====================================================
    ' Read old and new catalogue
    ' =====================================================
    oldCatalogueName = TableManager.GetText( _
        tbl, selectedRowIndex, COL_OLD_CATALOGUE)

    newCatalogueName = TableManager.GetText( _
        tbl, selectedRowIndex, COL_NEW_CATALOGUE)

    If ValidationText.IsBlank(oldCatalogueName) Then
        ErrorManager.ThrowError _
            sheetUpdateRequest.EmptyOriginalCatalogueName, _
            MODULE_NAME & ".HandleApproveRequestUpdate", _
            "Nama catalogue request original " & _
            "tidak boleh kosong."
    End If

    If ValidationText.IsBlank(newCatalogueName) Then
        ErrorManager.ThrowError _
            sheetUpdateRequest.EmptyProposedCatalogueName, _
            MODULE_NAME & ".HandleApproveRequestUpdate", _
            "Nama catalogue request proposed " & _
            "tidak boleh kosong."
    End If

    ' =====================================================
    ' Read manufacturing plant
    '
    ' The saved query returns old and new plant names,
    ' but the current worksheet layout does not display
    ' separate plant columns. The current filter is used
    ' as the default for both combinations.
    ' =====================================================
    oldManufacturingPlantName = FormManager.GetCellText( _
        Me, CELL_MANUFACTURING_PLANT)

    newManufacturingPlantName = _
        oldManufacturingPlantName

    If ValidationText.IsBlank(oldManufacturingPlantName) Then
        ErrorManager.ThrowError _
            sheetUpdateRequest.EmptyOriginalPlantName, _
            MODULE_NAME & ".HandleApproveRequestUpdate", _
            "Nama manufacturing plant " & _
            "tidak boleh kosong."
    End If

    ' =====================================================
    ' Determine the approving user
    ' =====================================================
    currentUser = pHelper.TrimText( _
        pConfig.currentUser())

    If ValidationText.IsBlank(currentUser) Then
        currentUser = pHelper.TrimText( _
            Environ$("Username"))
    End If

    If ValidationText.IsBlank(currentUser) Then
        ErrorManager.ThrowError _
            sheetUpdateRequest.EmptyApproverUser, _
            MODULE_NAME & ".HandleApproveRequestUpdate", _
            "User yang menyetujui tidak dapat " & _
            "ditentukan."
    End If

    ' =====================================================
    ' Build confirmation message
    ' =====================================================
    confirmationMessage = _
        "Apakah Anda yakin ingin menyetujui " & _
        "pengajuan update request ini?"

    confirmationMessage = confirmationMessage & _
        vbCrLf & vbCrLf & _
        "UUID request original:" & vbCrLf & _
        originalTransactionUUID

    confirmationMessage = confirmationMessage & _
        vbCrLf & vbCrLf & _
        "UUID request proposed:" & vbCrLf & _
        proposedTransactionUUID

    confirmationMessage = confirmationMessage & _
        vbCrLf & vbCrLf & _
        "Catalogue original: " & _
        oldCatalogueName

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "Catalogue proposed: " & _
        newCatalogueName

    confirmationMessage = confirmationMessage & _
        vbCrLf & vbCrLf & _
        "Tindakan ini akan:"

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "1. Menyetujui pengajuan update request."

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "2. Membalik transaksi request original."

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "3. Mengaktifkan transaksi request proposed."

    confirmationMessage = confirmationMessage & _
        vbCrLf & _
        "4. Menghitung ulang stok manufacturing plant."

    If Not FormManager.Confirm( _
        confirmationMessage, _
        "Konfirmasi Persetujuan Update Request") Then

        GoTo CleanUp
    End If

    ' =====================================================
    ' Set temporary status
    ' =====================================================
    TableManager.SetValue tbl, selectedRowIndex, COL_ACTION, _
        "MEMPROSES..."

    actionCellChanged = True

    With TableManager.GetCell(tbl, selectedRowIndex, COL_ACTION)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

        .Font.Bold = True
        .Font.Color = RGB(128, 128, 128)

        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(230, 230, 230)
    End With

    FormManager.SetStatus _
        "Menyetujui pengajuan update request...", True

    ' =====================================================
    ' Execute approval
    '
    ' updateOrder = False:
    '   This is not an order update.
    '
    ' updateStockMovement = True:
    '   The original request is reversed, the proposed
    '   request is activated, and exact stock is rebuilt.
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
    ' Approval succeeded
    ' =====================================================
    TableManager.SetValue tbl, selectedRowIndex, COL_ACTION, _
        "DISETUJUI"

    ' The approval is committed: an error after this point
    ' must not put "SETUJUI" back into the cell.
    actionCellChanged = False

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
        "Pengajuan update request " & _
        "berhasil disetujui.", _
        "Persetujuan Berhasil"

    ' Refresh pending requests.
    CariPengajuanUpdateRequest

CleanUp:
    FormManager.ClearStatus

    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ' Restore the action cell only if it had entered the
    ' processing state.
    If actionCellChanged Then

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
        MODULE_NAME & ".HandleApproveRequestUpdate", _
        "Gagal menyetujui pengajuan update request."

End Sub


' =========================================================
' Filter validation
' =========================================================
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
            sheetUpdateRequest.EmptyPlantName, _
            MODULE_NAME & ".ValidateSearchParameters", _
            "Lokasi manufacturing plant wajib diisi."
    End If

    If IsEmpty(fromDateValue) Then
        ErrorManager.ThrowError _
            sheetUpdateRequest.EmptyFromDate, _
            MODULE_NAME & ".ValidateSearchParameters", _
            "Tanggal awal wajib diisi."
    End If

    If Not IsDate(fromDateValue) Then
        ErrorManager.ThrowError _
            sheetUpdateRequest.InvalidFromDate, _
            MODULE_NAME & ".ValidateSearchParameters", _
            "Tanggal awal tidak valid."
    End If

    If IsEmpty(untilDateValue) Then
        ErrorManager.ThrowError _
            sheetUpdateRequest.EmptyUntilDate, _
            MODULE_NAME & ".ValidateSearchParameters", _
            "Tanggal akhir wajib diisi."
    End If

    If Not IsDate(untilDateValue) Then
        ErrorManager.ThrowError _
            sheetUpdateRequest.InvalidUntilDate, _
            MODULE_NAME & ".ValidateSearchParameters", _
            "Tanggal akhir tidak valid."
    End If

    fromDate = DateValue(fromDateValue)
    untilDate = DateValue(untilDateValue)

    If untilDate < fromDate Then
        ErrorManager.ThrowError _
            sheetUpdateRequest.InvalidDateRange, _
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
        "Filter pengajuan update request tidak valid."

End Sub


' =========================================================
' Clear previous result
' =========================================================
Private Sub ClearRequestUpdateData( _
    ByVal tbl As ListObject)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    If tbl Is Nothing Then
        ErrorManager.ThrowError _
            TableManager.TableNotFound, _
            MODULE_NAME & ".ClearRequestUpdateData", _
            "Tabel pengajuan update request " & _
            "tidak ditemukan."
    End If

    ' Delete existing table rows while preserving the
    ' table header, formatting and calculated columns.
    TableManager.ClearRows tbl

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".ClearRequestUpdateData", _
        "Gagal membersihkan data pengajuan " & _
        "update request sebelumnya."

End Sub


' =========================================================
' Formatting
' =========================================================
Private Sub ApplyRequestUpdateFormatting( _
    ByVal tbl As ListObject)

    On Error GoTo ErrorHandler

    Dim actionRange As Range

    Dim errInfo As ErrorInfo

    If tbl Is Nothing Then
        ErrorManager.ThrowError _
            TableManager.TableNotFound, _
            MODULE_NAME & ".ApplyRequestUpdateFormatting", _
            "Tabel pengajuan update request " & _
            "tidak ditemukan."
    End If

    If Not TableManager.HasRows(tbl) Then
        GoTo CleanUp
    End If

    ' Request submission date
    TableManager.FormatColumn tbl, COL_REQUEST_DATE, _
        "dd-mmm-yyyy"

    ' Old request date
    TableManager.FormatColumn tbl, COL_OLD_REQUEST_DATE, _
        "dd-mmm-yyyy"

    ' New request date
    TableManager.FormatColumn tbl, COL_NEW_REQUEST_DATE, _
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
        MODULE_NAME & ".ApplyRequestUpdateFormatting", _
        "Gagal menerapkan format pengajuan " & _
        "update request."

End Sub


' =========================================================
' Recordset helpers
' =========================================================

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
