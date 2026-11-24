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
' managerTable addresses columns by header text. These
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
    If managerForm.IsInRange( _
        Target, _
        Me.Range(CELL_FROM_DATE & ":" & CELL_UNTIL_DATE)) Then

        Application.EnableEvents = False

        managerForm.ShowDatePicker Target

        GoTo CleanUp

    End If

    ' =====================================================
    ' SETUJUI column
    ' =====================================================
    Set tbl = managerTable.GetTable( _
        Me, TABLE_REQUEST_UPDATE_REQUEST)

    If Not managerTable.HasRows(tbl) Then
        GoTo CleanUp
    End If

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

    HandleApproveRequestUpdate _
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

    state = managerForm.BeginBusy( _
        "Mengambil data pengajuan update request...")

    ' =====================================================
    ' Step 1: Read and validate filters
    ' =====================================================
    manufacturingPlantName = managerForm.GetCellText( _
        Me, CELL_MANUFACTURING_PLANT)

    ValidateSearchParameters _
        manufacturingPlantName:=manufacturingPlantName, _
        fromDateValue:=managerForm.GetCellValue(Me, CELL_FROM_DATE), _
        untilDateValue:=managerForm.GetCellValue(Me, CELL_UNTIL_DATE)

    fromDate = DateValue( _
        managerForm.GetCellValue(Me, CELL_FROM_DATE))

    untilDate = DateValue( _
        managerForm.GetCellValue(Me, CELL_UNTIL_DATE))

    ' =====================================================
    ' Step 2: Clear previous results
    ' =====================================================
    Set tbl = managerTable.GetTable( _
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
        managerError.ThrowError _
            managerDatabase.EmptyRecordset, _
            MODULE_NAME & ".CariPengajuanUpdateRequest", _
            "DAO tidak mengembalikan recordset."
    End If

    If rs.State <> adStateOpen Then
        managerError.ThrowError _
            managerDatabase.RecordsetNotOpen, _
            MODULE_NAME & ".CariPengajuanUpdateRequest", _
            "Recordset pengajuan update request " & _
            "tidak terbuka."
    End If

    ' =====================================================
    ' Step 4: Write records to the table
    ' =====================================================
    rowNumber = 1

    Do While Not rs.EOF

        rowIndex = managerTable.AddRow(tbl)

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

        managerForm.ShowInfo _
            "Tidak ada pengajuan update request " & _
            "yang ditemukan.", _
            "Pengajuan Update Request"

    Else

        managerForm.ShowInfo _
            CStr(resultCount) & _
            " pengajuan update request ditemukan.", _
            "Pengajuan Update Request"

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
    managerTable.SetValue tbl, rowIndex, COL_NUMBER, _
        rowNumber

    ' B: Update request date
    managerTable.SetValue tbl, rowIndex, COL_REQUEST_DATE, _
        RecordsetDate(rs, "update_request_date")

    ' =====================================================
    ' Old request
    ' =====================================================

    ' C: Old transaction UUID
    managerTable.SetValue tbl, rowIndex, COL_OLD_UUID, _
        managerDatabase.FieldText(rs, "old_transaction_uuid")

    ' D: Old mobile post
    managerTable.SetValue tbl, rowIndex, COL_OLD_MOBILE_POST, _
        managerDatabase.FieldText(rs, "old_mobile_post_name")

    ' E: Old catalogue
    managerTable.SetValue tbl, rowIndex, COL_OLD_CATALOGUE, _
        managerDatabase.FieldText(rs, "old_catalogue_name")

    ' F: Old amount
    managerTable.SetValue tbl, rowIndex, COL_OLD_AMOUNT, _
        managerDatabase.FieldDouble(rs, "old_amount")

    ' G: Old unit
    managerTable.SetValue tbl, rowIndex, COL_OLD_UNIT, _
        managerDatabase.FieldText(rs, "old_unit_of_measurement")

    ' H: Old request date
    managerTable.SetValue tbl, rowIndex, COL_OLD_REQUEST_DATE, _
        RecordsetDate(rs, "old_request_date")

    ' I: Old remarks
    managerTable.SetValue tbl, rowIndex, COL_OLD_REMARKS, _
        managerDatabase.FieldText(rs, "old_remarks")

    ' J: Old PIC
    managerTable.SetValue tbl, rowIndex, COL_OLD_PIC, _
        managerDatabase.FieldText(rs, "old_pic")

    ' =====================================================
    ' New request
    ' =====================================================

    ' K: New transaction UUID
    managerTable.SetValue tbl, rowIndex, COL_NEW_UUID, _
        managerDatabase.FieldText(rs, "new_transaction_uuid")

    ' L: New mobile post
    managerTable.SetValue tbl, rowIndex, COL_NEW_MOBILE_POST, _
        managerDatabase.FieldText(rs, "new_mobile_post_name")

    ' M: New catalogue
    managerTable.SetValue tbl, rowIndex, COL_NEW_CATALOGUE, _
        managerDatabase.FieldText(rs, "new_catalogue_name")

    ' N: New amount
    managerTable.SetValue tbl, rowIndex, COL_NEW_AMOUNT, _
        managerDatabase.FieldDouble(rs, "new_amount")

    ' O: New unit
    managerTable.SetValue tbl, rowIndex, COL_NEW_UNIT, _
        managerDatabase.FieldText(rs, "new_unit_of_measurement")

    ' P: New request date
    managerTable.SetValue tbl, rowIndex, COL_NEW_REQUEST_DATE, _
        RecordsetDate(rs, "new_request_date")

    ' Q: New remarks
    managerTable.SetValue tbl, rowIndex, COL_NEW_REMARKS, _
        managerDatabase.FieldText(rs, "new_remarks")

    ' R: New PIC
    managerTable.SetValue tbl, rowIndex, COL_NEW_PIC, _
        managerDatabase.FieldText(rs, "new_pic")

    ' =====================================================
    ' Requester and action
    ' =====================================================

    ' T: Requested by (S "NAMA PENGAJU" looks this up)
    managerTable.SetValue tbl, rowIndex, COL_REQUESTED_BY, _
        managerDatabase.FieldText(rs, "update_request_by")

    ' U: Action
    managerTable.SetValue tbl, rowIndex, COL_ACTION, _
        "SETUJUI"

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
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
        managerError.ThrowError _
            managerTable.TableNotFound, _
            MODULE_NAME & ".HandleApproveRequestUpdate", _
            "Tabel pengajuan update request " & _
            "belum diinisialisasi."
    End If

    If targetCell Is Nothing Then
        managerError.ThrowError _
            sheetUpdateRequest.EmptyApprovalCell, _
            MODULE_NAME & ".HandleApproveRequestUpdate", _
            "Cell persetujuan tidak ditemukan."
    End If

    ' Worksheet row, used in the messages below.
    targetRow = targetCell.Row

    ' Table row of the clicked cell (0 = outside the table).
    selectedRowIndex = managerTable.GetRowIndex( _
        tbl, targetCell)

    If Not managerTable.IsValidRow(tbl, selectedRowIndex) Then
        GoTo CleanUp
    End If

    actionText = managerTable.GetText( _
        tbl, selectedRowIndex, COL_ACTION)

    If Not managerValidationText.IsSameText( _
        actionText, _
        "SETUJUI") Then

        GoTo CleanUp
    End If

    ' =====================================================
    ' Read transaction UUIDs
    ' =====================================================
    originalTransactionUUID = managerTable.GetText( _
        tbl, selectedRowIndex, COL_OLD_UUID)

    proposedTransactionUUID = managerTable.GetText( _
        tbl, selectedRowIndex, COL_NEW_UUID)

    If managerValidationText.IsBlank(originalTransactionUUID) Then
        managerError.ThrowError _
            sheetUpdateRequest.EmptyOriginalRequestUUID, _
            MODULE_NAME & ".HandleApproveRequestUpdate", _
            "UUID transaksi request original pada " & _
            "baris " & CStr(targetRow) & _
            " tidak boleh kosong."
    End If

    If managerValidationText.IsBlank(proposedTransactionUUID) Then
        managerError.ThrowError _
            sheetUpdateRequest.EmptyProposedRequestUUID, _
            MODULE_NAME & ".HandleApproveRequestUpdate", _
            "UUID transaksi request proposed pada " & _
            "baris " & CStr(targetRow) & _
            " tidak boleh kosong."
    End If

    If managerValidationText.IsSameText( _
        originalTransactionUUID, _
        proposedTransactionUUID) Then

        managerError.ThrowError _
            sheetUpdateRequest.SameOriginalProposedRequestUUID, _
            MODULE_NAME & ".HandleApproveRequestUpdate", _
            "UUID transaksi original dan proposed " & _
            "tidak boleh sama."
    End If

    ' =====================================================
    ' Read old and new catalogue
    ' =====================================================
    oldCatalogueName = managerTable.GetText( _
        tbl, selectedRowIndex, COL_OLD_CATALOGUE)

    newCatalogueName = managerTable.GetText( _
        tbl, selectedRowIndex, COL_NEW_CATALOGUE)

    If managerValidationText.IsBlank(oldCatalogueName) Then
        managerError.ThrowError _
            sheetUpdateRequest.EmptyOriginalCatalogueName, _
            MODULE_NAME & ".HandleApproveRequestUpdate", _
            "Nama catalogue request original " & _
            "tidak boleh kosong."
    End If

    If managerValidationText.IsBlank(newCatalogueName) Then
        managerError.ThrowError _
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
    oldManufacturingPlantName = managerForm.GetCellText( _
        Me, CELL_MANUFACTURING_PLANT)

    newManufacturingPlantName = _
        oldManufacturingPlantName

    If managerValidationText.IsBlank(oldManufacturingPlantName) Then
        managerError.ThrowError _
            sheetUpdateRequest.EmptyOriginalPlantName, _
            MODULE_NAME & ".HandleApproveRequestUpdate", _
            "Nama manufacturing plant " & _
            "tidak boleh kosong."
    End If

    ' =====================================================
    ' Determine the approving user
    ' =====================================================
    currentUser = coreHelper.TrimText( _
        coreConfig.currentUser())

    If managerValidationText.IsBlank(currentUser) Then
        currentUser = coreHelper.TrimText( _
            Environ$("Username"))
    End If

    If managerValidationText.IsBlank(currentUser) Then
        managerError.ThrowError _
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

    If Not managerForm.Confirm( _
        confirmationMessage, _
        "Konfirmasi Persetujuan Update Request") Then

        GoTo CleanUp
    End If

    ' =====================================================
    ' Set temporary status
    ' =====================================================
    managerTable.SetValue tbl, selectedRowIndex, COL_ACTION, _
        "MEMPROSES..."

    actionCellChanged = True

    With managerTable.GetCell(tbl, selectedRowIndex, COL_ACTION)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

        .Font.Bold = True
        .Font.Color = RGB(128, 128, 128)

        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(230, 230, 230)
    End With

    managerForm.SetStatus _
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
    managerTable.SetValue tbl, selectedRowIndex, COL_ACTION, _
        "DISETUJUI"

    ' The approval is committed: an error after this point
    ' must not put "SETUJUI" back into the cell.
    actionCellChanged = False

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
        "Pengajuan update request " & _
        "berhasil disetujui.", _
        "Persetujuan Berhasil"

    ' Refresh pending requests.
    CariPengajuanUpdateRequest

CleanUp:
    managerForm.ClearStatus

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    ' Restore the action cell only if it had entered the
    ' processing state.
    If actionCellChanged Then

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

    If managerValidationText.IsBlank(manufacturingPlantName) Then
        managerError.ThrowError _
            sheetUpdateRequest.EmptyPlantName, _
            MODULE_NAME & ".ValidateSearchParameters", _
            "Lokasi manufacturing plant wajib diisi."
    End If

    If IsEmpty(fromDateValue) Then
        managerError.ThrowError _
            sheetUpdateRequest.EmptyFromDate, _
            MODULE_NAME & ".ValidateSearchParameters", _
            "Tanggal awal wajib diisi."
    End If

    If Not IsDate(fromDateValue) Then
        managerError.ThrowError _
            sheetUpdateRequest.InvalidFromDate, _
            MODULE_NAME & ".ValidateSearchParameters", _
            "Tanggal awal tidak valid."
    End If

    If IsEmpty(untilDateValue) Then
        managerError.ThrowError _
            sheetUpdateRequest.EmptyUntilDate, _
            MODULE_NAME & ".ValidateSearchParameters", _
            "Tanggal akhir wajib diisi."
    End If

    If Not IsDate(untilDateValue) Then
        managerError.ThrowError _
            sheetUpdateRequest.InvalidUntilDate, _
            MODULE_NAME & ".ValidateSearchParameters", _
            "Tanggal akhir tidak valid."
    End If

    fromDate = DateValue(fromDateValue)
    untilDate = DateValue(untilDateValue)

    If untilDate < fromDate Then
        managerError.ThrowError _
            sheetUpdateRequest.InvalidDateRange, _
            MODULE_NAME & ".ValidateSearchParameters", _
            "Tanggal akhir tidak boleh lebih kecil " & _
            "dari tanggal awal."
    End If

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
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
        managerError.ThrowError _
            managerTable.TableNotFound, _
            MODULE_NAME & ".ClearRequestUpdateData", _
            "Tabel pengajuan update request " & _
            "tidak ditemukan."
    End If

    ' Delete existing table rows while preserving the
    ' table header, formatting and calculated columns.
    managerTable.ClearRows tbl

CleanUp:
    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
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
        managerError.ThrowError _
            managerTable.TableNotFound, _
            MODULE_NAME & ".ApplyRequestUpdateFormatting", _
            "Tabel pengajuan update request " & _
            "tidak ditemukan."
    End If

    If Not managerTable.HasRows(tbl) Then
        GoTo CleanUp
    End If

    ' Request submission date
    managerTable.FormatColumn tbl, COL_REQUEST_DATE, _
        "dd-mmm-yyyy"

    ' Old request date
    managerTable.FormatColumn tbl, COL_OLD_REQUEST_DATE, _
        "dd-mmm-yyyy"

    ' New request date
    managerTable.FormatColumn tbl, COL_NEW_REQUEST_DATE, _
        "dd-mmm-yyyy"

    ' Old amount
    managerTable.FormatColumn tbl, COL_OLD_AMOUNT, _
        "#,##0.00"

    ' New amount
    managerTable.FormatColumn tbl, COL_NEW_AMOUNT, _
        "#,##0.00"

    ' General alignment (columns No .. SETUJUI)
    Me.Range( _
        managerTable.GetColumnRange(tbl, COL_NUMBER), _
        managerTable.GetColumnRange(tbl, COL_ACTION) _
    ).VerticalAlignment = xlCenter

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
        MODULE_NAME & ".ApplyRequestUpdateFormatting", _
        "Gagal menerapkan format pengajuan " & _
        "update request."

End Sub


' =========================================================
' Recordset helpers
' =========================================================

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
