Attribute VB_Name = "controllerRequestUpdate"
Option Explicit

' =========================================================
' CONTROLLER - REQUEST UPDATE (sheet REQUEST_UPDATE)
'
' Pending update requests of the application location
' (SETTINGS > LOKASI), loaded when the sheet opens.
'
'   TableRequestUpdate (NO, TANGGAL REQUEST, UUID,
'   JENIS TRANSAKSI, PENGAJU, TIER PENGAJU, CHECK)
'     CHECK : double-click shows the old and new values in
'             TableUpdateDetail (KOLOM, DATA LAMA, DATA BARU);
'             changed values are bold
'   REMARKS REVIEWER  box under the detail table (Q9 in the
'                     design), or U6
'   SETUJUI / TOLAK   double-click (the big buttons under the
'                     remarks box, or U9 / U11)
'
'   The detail table grows downwards: cells are inserted in
'   columns Q:S only, so the remarks box and the buttons move
'   down with it. They are found by their labels.
'
' Only a registered user with a higher tier than the
' requester can approve or reject (checked by the service).
' =========================================================

Private Const MODULE_NAME As String = "controllerRequestUpdate"

' =========================================================
' ERROR CODES (15400 - 15499)
' =========================================================
Public Const NothingSelected As Long = 15401
Public Const EmptyLocation As Long = 15402

Private Const TABLE_REQUESTS As String = "TableRequestUpdate"
Private Const TABLE_DETAIL As String = "TableUpdateDetail"

Private Const COL_NO As String = "NO"
Private Const COL_CHECK As String = "CHECK"

Private Const REMARKS_SMALL_CELL As String = "U6"
Private Const REMARKS_LABEL As String = "REMARKS REVIEWER"
Private Const REMARKS_DEFAULT_CELL As String = "Q9"

Private Const LABEL_APPROVE As String = "SETUJUI"
Private Const LABEL_REJECT As String = "TOLAK"

Private Const RULE_CHANGED As String = "SIS_UPDATE_CHANGED"

Private mItems As Collection        ' TransactionUpdateRequest, in table order
Private mSelectedId As Long
Private mPlantName As String


' =========================================================
' EVENTS
' =========================================================

Public Sub OnActivate( _
    ByVal ws As Worksheet)

    If Len(managerSheet.SettingsPlantName()) = 0 Then
        Application.StatusBar = "Isi LOKASI di sheet SETTINGS untuk melihat request update."
        Exit Sub
    End If

    LoadRequests ws, False

End Sub


Public Sub OnDoubleClick( _
    ByVal ws As Worksheet, _
    ByVal target As Range, _
    ByRef cancel As Boolean)

    Dim tbl As ListObject
    Dim rowIndex As Long

    Select Case ButtonAt(ws, target)
        Case LABEL_APPROVE
            cancel = True
            Decide ws, True
            Exit Sub
        Case LABEL_REJECT
            cancel = True
            Decide ws, False
            Exit Sub
    End Select

    Set tbl = managerTable.GetTable(ws, TABLE_REQUESTS)

    If managerTable.IsInColumn(target.Cells(1, 1), tbl, COL_CHECK) Then
        cancel = True
        rowIndex = managerTable.GetRowIndex(tbl, target.Cells(1, 1))
        If rowIndex > 0 Then SelectRow ws, rowIndex
    End If

End Sub


' =========================================================
' ACTIONS
' =========================================================

Public Sub LoadRequests( _
    ByVal ws As Worksheet, _
    ByVal showResult As Boolean)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo
    Dim items As Collection
    Dim users As Collection
    Dim requester As User
    Dim item As TransactionUpdateRequest
    Dim data() As Variant
    Dim r As Long

    mPlantName = managerSheet.SettingsPlantName()

    If Len(mPlantName) = 0 Then
        managerError.ThrowError controllerRequestUpdate.EmptyLocation, _
            MODULE_NAME & ".LoadRequests", _
            "Isi LOKASI di sheet SETTINGS untuk melihat request update."
    End If

    state = managerForm.BeginBusy("Memuat request update...")

    Set items = serviceTransactionUpdateRequest.ListPendingUpdateRequests(mPlantName)
    Set users = serviceUser.ListUsers()

    Set mItems = New Collection
    mSelectedId = 0

    If items.Count > 0 Then

        ReDim data(1 To items.Count, 1 To 7)

        For Each item In items

            r = r + 1
            mItems.Add item, CStr(item.Id)

            Set requester = serviceUser.FindInList(users, item.RequestedBy)

            data(r, 1) = r
            data(r, 2) = item.RequestedAt
            data(r, 3) = item.OriginalUUID
            data(r, 4) = TypeLabel(item)
            data(r, 5) = item.RequestedBy
            If Not requester Is Nothing Then data(r, 6) = requester.Tier
            data(r, 7) = managerSheet.LABEL_CHECK

        Next item

        managerTable.WriteData managerTable.GetTable(ws, TABLE_REQUESTS), data, items.Count, _
            Array("0", "dd-mmm-yyyy hh:mm", managerSheet.TEXT_FORMAT, _
                  managerSheet.TEXT_FORMAT, managerSheet.TEXT_FORMAT, "0", _
                  managerSheet.TEXT_FORMAT)
    Else
        managerTable.SetRowCount managerTable.GetTable(ws, TABLE_REQUESTS), 0
    End If

    ClearDetail ws

    managerForm.EndBusy state

    If showResult Or items.Count = 0 Then
        Application.StatusBar = items.Count & " request update menunggu persetujuan (" & _
            mPlantName & ")."
    End If

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


' CHECK: show the old / new values of one request.
Public Sub SelectRow( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo
    Dim tbl As ListObject
    Dim item As TransactionUpdateRequest
    Dim lines As Collection
    Dim data() As Variant
    Dim line As Variant
    Dim r As Long

    Set tbl = managerTable.GetTable(ws, TABLE_REQUESTS)
    Set item = ItemAtRow(rowIndex)

    state = managerForm.BeginBusy()

    managerSheet.FillColumn tbl, COL_CHECK, managerSheet.LABEL_CHECK
    managerTable.SetValue tbl, rowIndex, COL_CHECK, managerSheet.LABEL_SELECTED

    mSelectedId = item.Id

    Set lines = DetailLines(item)

    ReDim data(1 To lines.Count, 1 To 3)

    managerTable.MakeRoomBelow managerTable.GetTable(ws, TABLE_DETAIL), lines.Count

    For Each line In lines
        r = r + 1
        data(r, 1) = line(0)
        data(r, 2) = line(1)
        data(r, 3) = line(2)
    Next line

    managerTable.WriteData managerTable.GetTable(ws, TABLE_DETAIL), data, lines.Count, _
        Array(managerSheet.TEXT_FORMAT, managerSheet.TEXT_FORMAT, managerSheet.TEXT_FORMAT)

    HighlightChanges ws

    RemarksCell(ws).MergeArea.ClearContents

    managerForm.EndBusy state

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


' SETUJUI (approve = True) / TOLAK (approve = False).
Public Sub Decide( _
    ByVal ws As Worksheet, _
    ByVal approve As Boolean)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo
    Dim item As TransactionUpdateRequest
    Dim note As String
    Dim verb As String

    Set item = SelectedItem(ws)

    note = Trim$(coreHelper.TextOrEmpty(RemarksCell(ws).value))
    If Len(note) = 0 Then note = managerForm.GetCellText(ws, REMARKS_SMALL_CELL)

    If approve Then
        verb = "SETUJUI"
    Else
        verb = "TOLAK"
    End If

    If Not managerForm.Confirm(verb & " request update " & TypeLabel(item) & " " & _
        item.OriginalUUID & " dari " & item.RequestedBy & "?" & vbCrLf & vbCrLf & _
        "Alasan pengaju: " & item.Reason) Then Exit Sub

    state = managerForm.BeginBusy("Memproses request update...", True)

    If approve Then
        serviceTransactionUpdateRequest.ApproveUpdateRequest mPlantName, item, note
    Else
        serviceTransactionUpdateRequest.RejectUpdateRequest mPlantName, item, note
    End If

    RemarksCell(ws).MergeArea.ClearContents
    managerForm.ClearCells ws, REMARKS_SMALL_CELL

    managerForm.EndBusy state

    If approve Then
        managerForm.ShowInfo "Request update disetujui." & vbCrLf & _
            item.OriginalUUID & " dibatalkan, " & item.ProposedUUID & " aktif."
    Else
        managerForm.ShowInfo "Request update ditolak. " & item.OriginalUUID & _
            " tetap aktif."
    End If

    LoadRequests ws, True

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


' =========================================================
' INTERNAL
' =========================================================

Private Function ItemAtRow( _
    ByVal rowIndex As Long) As TransactionUpdateRequest

    If mItems Is Nothing Then GoTo NotLoaded
    If rowIndex < 1 Or rowIndex > mItems.Count Then GoTo NotLoaded

    Set ItemAtRow = mItems(rowIndex)
    Exit Function

NotLoaded:
    managerError.ThrowError controllerRequestUpdate.NothingSelected, _
        MODULE_NAME & ".ItemAtRow", _
        "Data request update belum dimuat. Buka ulang sheet REQUEST UPDATE."

End Function


' The checked request (also after a VBA reset: the row that
' shows DIPILIH).
Private Function SelectedItem( _
    ByVal ws As Worksheet) As TransactionUpdateRequest

    Dim tbl As ListObject
    Dim rowIndex As Long

    If mSelectedId > 0 And Not mItems Is Nothing Then
        On Error Resume Next
        Set SelectedItem = mItems(CStr(mSelectedId))
        On Error GoTo 0
        If Not SelectedItem Is Nothing Then Exit Function
    End If

    Set tbl = managerTable.GetTable(ws, TABLE_REQUESTS)

    For rowIndex = 1 To managerTable.RowCount(tbl)
        If managerTable.GetText(tbl, rowIndex, COL_CHECK) = managerSheet.LABEL_SELECTED Then
            If mItems Is Nothing Then LoadRequests ws, False
            Set SelectedItem = ItemAtRow(rowIndex)
            Exit Function
        End If
    Next rowIndex

    managerError.ThrowError controllerRequestUpdate.NothingSelected, _
        MODULE_NAME & ".SelectedItem", _
        "Pilih request update dengan double-click pada kolom CHECK."

End Function


Private Sub ClearDetail( _
    ByVal ws As Worksheet)

    managerTable.SetRowCount managerTable.GetTable(ws, TABLE_DETAIL), 0
    RemarksCell(ws).MergeArea.ClearContents

End Sub


' SETUJUI / TOLAK when target is one of those buttons, ""
' otherwise. Big buttons: icon in Q, label in R (both merged
' over two rows). Small buttons: the label cell in column U.
Private Function ButtonAt( _
    ByVal ws As Worksheet, _
    ByVal target As Range) As String

    Dim topCell As Range
    Dim labelValue As Variant
    Dim labelText As String

    Set topCell = target.Cells(1, 1).MergeArea.Cells(1, 1)

    If managerTable.IsInTable(topCell, managerTable.GetTable(ws, TABLE_DETAIL)) Then Exit Function
    If managerTable.IsInTable(topCell, managerTable.GetTable(ws, TABLE_REQUESTS)) Then Exit Function

    Select Case topCell.Column
        Case 17 To 19                   ' Q:S
            labelValue = ws.Cells(topCell.Row, 18).value
        Case 21                         ' U
            labelValue = topCell.value
        Case Else
            Exit Function
    End Select

    If IsError(labelValue) Then Exit Function

    labelText = UCase$(Trim$(coreHelper.TextOrEmpty(labelValue)))

    If labelText = LABEL_APPROVE Or labelText = LABEL_REJECT Then
        ButtonAt = labelText
    End If

End Function


' The remarks input: the cell under the REMARKS REVIEWER
' label that sits below the detail table.
Private Function RemarksCell( _
    ByVal ws As Worksheet) As Range

    Dim tbl As ListObject
    Dim rowNumber As Long
    Dim firstRow As Long
    Dim columnNumber As Long

    Set tbl = managerTable.GetTable(ws, TABLE_DETAIL)

    columnNumber = tbl.Range.Column
    firstRow = tbl.Range.Row + tbl.Range.Rows.Count

    For rowNumber = firstRow To firstRow + 10
        If UCase$(Trim$(coreHelper.TextOrEmpty(ws.Cells(rowNumber, columnNumber).value))) _
            = REMARKS_LABEL Then

            Set RemarksCell = ws.Cells(rowNumber + 1, columnNumber)
            Exit Function
        End If
    Next rowNumber

    Set RemarksCell = ws.Range(REMARKS_DEFAULT_CELL)

End Function


' Bold + green fill on DATA BARU when it differs from DATA LAMA.
Private Sub HighlightChanges( _
    ByVal ws As Worksheet)

    Dim tbl As ListObject

    Set tbl = managerTable.GetTable(ws, TABLE_DETAIL)

    managerTable.EnsureFormulaRule tbl, RULE_CHANGED, _
        managerTable.RowCellFormula(tbl, "DATA LAMA") & "<>" & _
        managerTable.RowCellFormula(tbl, "DATA BARU"), _
        "DATA BARU", RGB(198, 239, 206), -1, True

End Sub


' ORDER / PENERIMAAN / REQUEST, with "(HAPUS)" for a delete.
Private Function TypeLabel( _
    ByVal item As TransactionUpdateRequest) As String

    Dim label As String

    Select Case item.TransactionType
        Case "ORDER":   label = "ORDER"
        Case "RECEIVE": label = "PENERIMAAN"
        Case "REQUEST": label = "REQUEST"
        Case Else:      label = item.TransactionType
    End Select

    If IsDeleteProposal(item) Then label = label & " (HAPUS)"

    TypeLabel = label

End Function


Private Function IsDeleteProposal( _
    ByVal item As TransactionUpdateRequest) As Boolean

    Dim quantity As Variant

    quantity = item.ProposedValue("quantity")

    If IsNull(quantity) Or IsEmpty(quantity) Then Exit Function
    IsDeleteProposal = (CDbl(quantity) = 0)

End Function


' Lines Array(KOLOM, DATA LAMA, DATA BARU) per document type.
Private Function DetailLines( _
    ByVal item As TransactionUpdateRequest) As Collection

    Dim lines As Collection

    Set lines = New Collection

    AddLine lines, item, "NO DOKUMEN", "uuid", False
    AddConstantLine lines, "AKSI", "", IIf(IsDeleteProposal(item), "HAPUS", "UPDATE")

    Select Case item.TransactionType

        Case "ORDER"
            AddLine lines, item, "NAMA SPAREPART", "sparepart_name", False
            AddLine lines, item, "SATUAN", "sparepart_unit_of_measurement", False
            AddLine lines, item, "VENDOR", "vendor_name", False
            AddLine lines, item, "PO ID", "po_number", False
            AddLine lines, item, "TANGGAL ORDER", "document_date", True
            AddLine lines, item, "JUMLAH ORDER", "quantity", False

        Case "RECEIVE"
            AddLine lines, item, "ORDER UUID", "order_uuid", False
            AddLine lines, item, "PO ID", "po_number", False
            AddLine lines, item, "NAMA SPAREPART", "sparepart_name", False
            AddLine lines, item, "TANGGAL DATANG", "document_date", True
            AddLine lines, item, "JUMLAH PENERIMAAN", "quantity", False
            AddLine lines, item, "JUMLAH OK", "ok_quantity", False
            AddLine lines, item, "JUMLAH NOT OK", "not_ok_quantity", False

        Case "REQUEST"
            AddLine lines, item, "NAMA SPAREPART", "sparepart_name", False
            AddLine lines, item, "POS MOBILE MTC", "mobile_maintenance_name", False
            AddLine lines, item, "TANGGAL REQUEST", "document_date", True
            AddLine lines, item, "JUMLAH REQUEST", "quantity", False

    End Select

    AddLine lines, item, "REMARKS", "remarks", False
    AddLine lines, item, "PIC", "person_in_charge", False
    AddConstantLine lines, "ALASAN PENGAJU", "", item.Reason
    AddConstantLine lines, "DIAJUKAN", "", _
        item.RequestedBy & " - " & Format$(item.RequestedAt, "dd-mmm-yyyy hh:mm")

    Set DetailLines = lines

End Function


Private Sub AddLine( _
    ByVal lines As Collection, _
    ByVal item As TransactionUpdateRequest, _
    ByVal label As String, _
    ByVal key As String, _
    ByVal asDate As Boolean)

    lines.Add Array(label, _
        ShowValue(item.OriginalValue(key), asDate), _
        ShowValue(item.ProposedValue(key), asDate))

End Sub


Private Sub AddConstantLine( _
    ByVal lines As Collection, _
    ByVal label As String, _
    ByVal oldText As String, _
    ByVal newText As String)

    lines.Add Array(label, oldText, newText)

End Sub


Private Function ShowValue( _
    ByVal value As Variant, _
    ByVal asDate As Boolean) As String

    If IsNull(value) Or IsEmpty(value) Then
        ShowValue = vbNullString
    ElseIf asDate And IsDate(value) Then
        ShowValue = Format$(CDate(value), managerSheet.DATE_FORMAT)
    Else
        ShowValue = CStr(value)
    End If

End Function
