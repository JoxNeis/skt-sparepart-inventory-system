Attribute VB_Name = "controllerRekapRequest"
Option Explicit

' =========================================================
' CONTROLLER - REKAP REQUEST / PENGELUARAN (sheet REKAP_REQUEST)
'
'   LOKASI K5, DARI K9 / HINGGA K10 (request date, double-click
'   = date picker), CARI S9:U9, SEBELUMNYA S7, SELANJUTNYA U7,
'   FILTER KOLOM P7 / NILAI Q7, BARIS YANG DITAMPILKAN U10
'
'   TableRekapRequest (NAMA SPAREPART and POS MOBILE MTC show
'   "code - name", the same text as their dropdowns)
'     EDIT  : row becomes editable -> AJUKAN = update proposal
'             (checked against the stock card)
'     HAPUS : delete proposal (quantity 0)
' =========================================================

Private Const MODULE_NAME As String = "controllerRekapRequest"

' =========================================================
' ERROR CODES (15300 - 15399)
' =========================================================
Public Const RequestNotLoaded As Long = 15301

Private Const LAYOUT_KEY As String = "REKAP_REQUEST"
Private Const TABLE_NAME As String = "TableRekapRequest"

Private Const COL_UUID As String = "UUID"
Private Const COL_SPAREPART As String = "NAMA SPAREPART"
Private Const COL_QUANTITY As String = "JUMLAH REQUEST"
Private Const COL_UNIT As String = "SATUAN"
Private Const COL_MOBILE As String = "POS MOBILE MTC"
Private Const COL_DATE As String = "TANGGAL REQUEST"
Private Const COL_REMARKS As String = "REMARKS"
Private Const COL_PIC As String = "PIC"
Private Const COL_EDIT As String = "EDIT"
Private Const COL_DELETE As String = "HAPUS"

Private mRequests As Collection
Private mPlantName As String
Private mFromDate As Date
Private mToDate As Date


' =========================================================
' EVENTS
' =========================================================

Public Sub OnActivate( _
    ByVal ws As Worksheet)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo

    state = managerForm.BeginBusy()

    managerPaging.PrepareSheet ws, GetLayout(), FilterHeaders()
    PrepareTable ws

    managerForm.EndBusy state
    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


Public Sub OnDoubleClick( _
    ByVal ws As Worksheet, _
    ByVal target As Range, _
    ByRef cancel As Boolean)

    Dim layout As PagingLayout
    Dim tbl As ListObject
    Dim rowIndex As Long
    Dim delta As Long
    Dim labelText As String

    Set layout = GetLayout()

    If managerPaging.IsDateCell(ws, layout, target) Then
        managerForm.HandleDatePickerDoubleClick target, target.Cells(1, 1), cancel
        Exit Sub
    End If

    delta = managerPaging.PageButtonDelta(ws, layout, target)

    If delta <> 0 Then
        cancel = True
        MovePage ws, delta
        Exit Sub
    End If

    If managerSheet.IsHit(target, ws, layout.SearchButton) Then
        cancel = True
        Search ws
        Exit Sub
    End If

    Set tbl = managerTable.GetTable(ws, TABLE_NAME)
    rowIndex = managerTable.GetRowIndex(tbl, target.Cells(1, 1))
    If rowIndex = 0 Then Exit Sub

    If managerTable.IsInColumn(target.Cells(1, 1), tbl, COL_DATE) Then
        If controllerShared.IsEditRow(tbl, rowIndex, COL_EDIT) Then
            managerForm.HandleDatePickerDoubleClick target, target.Cells(1, 1), cancel
        End If
        Exit Sub
    End If

    If managerTable.IsInColumn(target.Cells(1, 1), tbl, COL_EDIT) Then

        cancel = True
        labelText = managerTable.GetText(tbl, rowIndex, COL_EDIT)

        If labelText = managerSheet.LABEL_EDIT Then
            controllerMDPosMobile.EnsurePlant mPlantName
            controllerShared.EnterEditMode tbl, rowIndex, COL_EDIT, COL_DELETE, COL_QUANTITY
        ElseIf labelText = managerSheet.LABEL_SUBMIT Then
            SubmitUpdate ws, rowIndex
        End If

    ElseIf managerTable.IsInColumn(target.Cells(1, 1), tbl, COL_DELETE) Then

        cancel = True
        labelText = managerTable.GetText(tbl, rowIndex, COL_DELETE)

        If labelText = managerSheet.LABEL_DELETE Then
            ProposeDelete ws, rowIndex
        ElseIf labelText = managerSheet.LABEL_CANCEL Then
            RenderPage ws
        End If

    End If

End Sub


Public Sub OnChange( _
    ByVal ws As Worksheet, _
    ByVal target As Range)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo
    Dim layout As PagingLayout
    Dim tbl As ListObject

    Set layout = GetLayout()

    If managerPaging.IsViewCell(ws, layout, target) Then
        If managerPaging.HasState(LAYOUT_KEY) Then Search ws
        Exit Sub
    End If

    Set tbl = managerTable.GetTable(ws, TABLE_NAME)

    If Not managerTable.IsInTable(target, tbl) Then Exit Sub

    If controllerShared.GuardRecapChange(tbl, target, COL_EDIT, EditableColumns()) Then Exit Sub

    controllerShared.FillUnits tbl, target, COL_SPAREPART, COL_UNIT

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerError.ShowError errInfo

End Sub


' =========================================================
' ACTIONS
' =========================================================

' CARI: remember the criteria and load page 1 (count + rows,
' paged by the query).
Public Sub Search( _
    ByVal ws As Worksheet)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo
    Dim layout As PagingLayout

    Set layout = GetLayout()

    managerPaging.ReadCriteria ws, layout, mPlantName, mFromDate, mToDate

    LoadPage ws, managerPaging.FirstPageRequest(ws, layout, FilterHeaders()), 1

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerError.ShowError errInfo

End Sub

' SEBELUMNYA / SELANJUTNYA: load the neighbouring page.
Private Sub MovePage( _
    ByVal ws As Worksheet, _
    ByVal delta As Long)

    Dim request As PageRequest
    Dim newPage As Long

    Set request = managerPaging.MoveRequest(GetLayout(), delta, newPage)

    If Not request Is Nothing Then LoadPage ws, request, newPage

End Sub

' Load one page of the search and show it.
Private Sub LoadPage( _
    ByVal ws As Worksheet, _
    ByVal request As PageRequest, _
    ByVal pageNumber As Long)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo
    Dim totalRows As Long
    Dim items As Collection
    Dim requestItem As Request
    Dim data() As Variant
    Dim r As Long
    Dim keyHolder As New PageRequest
    Dim firstKey As PageRequest
    Dim lastKey As PageRequest

    state = managerForm.BeginBusy("Mengambil data request...", True)

    totalRows = serviceRequest.CountRequests(mPlantName, mFromDate, mToDate, request)
    Set items = serviceRequest.ListRequests(mPlantName, mFromDate, mToDate, request)

    Set mRequests = New Collection

    If items.Count > 0 Then

        ReDim data(1 To items.Count, 1 To 10)

        For Each requestItem In items

            r = r + 1
            mRequests.Add requestItem, requestItem.UUID

            data(r, 1) = requestItem.UUID
            data(r, 2) = requestItem.Sparepart.DisplayName
            data(r, 3) = requestItem.Quantity
            data(r, 4) = requestItem.UnitOfMeasurement
            data(r, 5) = requestItem.MobileMaintenance.DisplayName
            data(r, 6) = requestItem.RequestDate
            data(r, 7) = requestItem.Remarks
            data(r, 8) = requestItem.PersonInCharge
            data(r, 9) = managerSheet.LABEL_EDIT
            data(r, 10) = managerSheet.LABEL_DELETE

            Set lastKey = keyHolder.KeyOf(requestItem.RequestDate, requestItem.Id)
            If r = 1 Then Set firstKey = lastKey

        Next requestItem

        managerPaging.StorePage GetLayout(), data, items.Count, totalRows, _
            pageNumber, request, firstKey, lastKey
    Else
        managerPaging.StorePage GetLayout(), Empty, 0, totalRows, _
            pageNumber, request, Nothing, Nothing
    End If

    RenderPage ws

    managerForm.EndBusy state

    If totalRows = 0 Then
        Application.StatusBar = "Tidak ada request pada periode ini."
    End If

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


' AJUKAN: save the edited row as an update proposal.
Public Sub SubmitUpdate( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo
    Dim tbl As ListObject
    Dim original As Request
    Dim proposed As Request
    Dim reason As String
    Dim proposedUUID As String

    Set tbl = managerTable.GetTable(ws, TABLE_NAME)
    Set original = CachedRequest(managerTable.GetText(tbl, rowIndex, COL_UUID))

    state = managerForm.BeginBusy("Memeriksa perubahan...")

    Set proposed = New Request

    Set proposed.Sparepart = controllerShared.RequireSparepart( _
        serviceSparepart.ListSpareparts(), _
        managerSheet.TableRequiredText(tbl, rowIndex, COL_SPAREPART), _
        managerTable.GetCell(tbl, rowIndex, COL_SPAREPART))

    Set proposed.MobileMaintenance = controllerShared.RequireMobileMaintenance( _
        serviceMobileMaintenance.ListByPlant(mPlantName), _
        managerSheet.TableRequiredText(tbl, rowIndex, COL_MOBILE), _
        mPlantName, _
        managerTable.GetCell(tbl, rowIndex, COL_MOBILE))

    proposed.RequestDate = managerSheet.TableDate(tbl, rowIndex, COL_DATE)
    proposed.Quantity = managerSheet.TableWholeNumber(tbl, rowIndex, COL_QUANTITY, 0)
    proposed.Remarks = managerTable.GetText(tbl, rowIndex, COL_REMARKS)
    proposed.PersonInCharge = managerSheet.TableRequiredText(tbl, rowIndex, COL_PIC)

    managerForm.EndBusy state

    If Not controllerShared.AskReason( _
        "Alasan perubahan request " & original.UUID & ":", reason) Then Exit Sub

    state = managerForm.BeginBusy("Menyimpan usulan update...", True)

    proposedUUID = serviceRequest.ProposeRequestUpdate(mPlantName, original, proposed, reason)

    RenderPage ws

    managerForm.EndBusy state

    managerForm.ShowInfo "Usulan update request " & original.UUID & _
        " tersimpan sebagai " & proposedUUID & "." & vbCrLf & vbCrLf & _
        "Perubahan berlaku (dan stok dihitung ulang) setelah disetujui di REQUEST UPDATE."

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


' HAPUS: propose to delete the request.
Public Sub ProposeDelete( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo
    Dim tbl As ListObject
    Dim original As Request
    Dim reason As String
    Dim proposedUUID As String

    Set tbl = managerTable.GetTable(ws, TABLE_NAME)
    Set original = CachedRequest(managerTable.GetText(tbl, rowIndex, COL_UUID))

    If Not managerForm.Confirm("Ajukan penghapusan request " & original.UUID & vbCrLf & _
        original.Sparepart.DisplayName & " x " & original.Quantity & " " & _
        original.UnitOfMeasurement & " ke " & original.MobileMaintenance.Name & "?") Then Exit Sub

    If Not controllerShared.AskReason( _
        "Alasan penghapusan request " & original.UUID & ":", reason) Then Exit Sub

    state = managerForm.BeginBusy("Menyimpan usulan hapus...", True)

    proposedUUID = serviceRequest.ProposeRequestDelete(mPlantName, original, reason)

    managerForm.EndBusy state

    managerForm.ShowInfo "Usulan hapus request " & original.UUID & _
        " tersimpan (" & proposedUUID & ")." & vbCrLf & vbCrLf & _
        "Request terhapus (stok kembali) setelah disetujui di REQUEST UPDATE."

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


' =========================================================
' INTERNAL
' =========================================================

Private Function GetLayout() As PagingLayout

    Dim result As PagingLayout

    Set result = New PagingLayout
    result.UseStandardRecap LAYOUT_KEY, TABLE_NAME

    Set GetLayout = result

End Function


Private Function FilterHeaders() As Variant

    FilterHeaders = Array(COL_UUID, COL_SPAREPART, COL_QUANTITY, COL_UNIT, _
        COL_MOBILE, COL_DATE, COL_REMARKS, COL_PIC)

End Function


Private Function EditableColumns() As Variant

    EditableColumns = Array(COL_SPAREPART, COL_QUANTITY, COL_MOBILE, _
        COL_DATE, COL_REMARKS, COL_PIC)

End Function


Private Function ColumnFormats() As Variant

    ColumnFormats = Array( _
        managerSheet.TEXT_FORMAT, managerSheet.TEXT_FORMAT, managerSheet.NUMBER_FORMAT, _
        managerSheet.TEXT_FORMAT, managerSheet.TEXT_FORMAT, managerSheet.DATE_FORMAT, _
        managerSheet.TEXT_FORMAT, managerSheet.TEXT_FORMAT, managerSheet.TEXT_FORMAT, _
        managerSheet.TEXT_FORMAT)

End Function


Private Sub RenderPage( _
    ByVal ws As Worksheet)

    Dim eventsWereOn As Boolean

    eventsWereOn = Application.EnableEvents
    Application.EnableEvents = False

    managerPaging.Render ws, GetLayout(), ColumnFormats()
    PrepareTable ws

    Application.EnableEvents = eventsWereOn

End Sub


Private Sub PrepareTable( _
    ByVal ws As Worksheet)

    Dim tbl As ListObject

    Set tbl = managerTable.GetTable(ws, TABLE_NAME)

    managerList.ApplyDropdown managerTable.GetColumnRange(tbl, COL_SPAREPART), _
        managerList.LIST_SPAREPART
    managerList.ApplyDropdown managerTable.GetColumnRange(tbl, COL_MOBILE), _
        managerList.LIST_MOBILE_POST
    managerList.ApplyWholeNumberRule managerTable.GetColumnRange(tbl, COL_QUANTITY), 0

    controllerShared.EnsureEditHighlight tbl, COL_EDIT, EditableColumns()

End Sub


Private Function CachedRequest( _
    ByVal requestUUID As String) As Request

    If mRequests Is Nothing Then GoTo NotLoaded

    On Error GoTo NotLoaded
    Set CachedRequest = mRequests(requestUUID)
    Exit Function

NotLoaded:
    managerError.ThrowError controllerRekapRequest.RequestNotLoaded, _
        MODULE_NAME & ".CachedRequest", _
        "Data request belum dimuat atau sudah kedaluwarsa. Klik CARI lalu ulangi."

End Function
