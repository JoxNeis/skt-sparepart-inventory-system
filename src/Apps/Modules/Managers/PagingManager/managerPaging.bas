Attribute VB_Name = "managerPaging"
Option Explicit

' =========================================================
' PAGING MANAGER
'
' Search + paging + filter for the recap sheets and the stock
' card. The paging itself is done by the saved queries
' (keyset paging, see PageRequest and order_list.sql): a
' sheet only ever holds the page it shows.
'
'   CARI
'       Set request = managerPaging.FirstPageRequest(ws, layout, headers, indexes)
'       ...service.CountX(..., request) / service.ListX(..., request)
'       managerPaging.StorePage layout, data, rows, total, 1, request, firstKey, lastKey
'       managerPaging.Render ws, layout, columnFormats
'
'   SEBELUMNYA / SELANJUTNYA
'       Set request = managerPaging.MoveRequest(layout, delta, newPage)
'       If Not request Is Nothing Then ... load newPage ...
'
'   KOLOM / NILAI / BARIS YANG DITAMPILKAN changed
'       search again from page 1 (IsViewCell)
'
' The page lives in memory (one PageState per sheet). When
' VBA is reset (for example after editing code) it is gone;
' the user simply presses CARI again.
' =========================================================

Private Const MODULE_NAME As String = "managerPaging"

' =========================================================
' ERROR CODES (10700 - 10799)
' =========================================================
Public Const NoSearchResult As Long = 10701     ' paging before CARI

Public Const DEFAULT_PAGE_SIZE As Long = 25

Private mStates As Collection


' =========================================================
' SHEET PREPARATION
' =========================================================

' Defaults and dropdowns of the search cells:
'   LOKASI -> plant dropdown + application location
'   DARI   -> first day of this month, HINGGA -> today
'   KOLOM  -> dropdown with filterHeaders
'   BARIS YANG DITAMPILKAN -> 25 when blank
Public Sub PrepareSheet( _
    ByVal ws As Worksheet, _
    ByVal layout As PagingLayout, _
    Optional ByVal filterHeaders As Variant)

    If Len(layout.PlantCell) > 0 Then
        managerSheet.PreparePlantCell ws, layout.PlantCell
    End If

    If Len(layout.FromCell) > 0 Then
        managerSheet.PrepareDateCell ws, layout.FromCell, DateSerial(Year(Date), Month(Date), 1)
    End If

    If Len(layout.ToCell) > 0 Then
        managerSheet.PrepareDateCell ws, layout.ToCell, Date
    End If

    If Len(layout.FilterColumnCell) > 0 And Not IsMissing(filterHeaders) Then
        managerList.ApplyItemsDropdown ws.Range(layout.FilterColumnCell), filterHeaders
    End If

    If Len(layout.PageSizeCell) > 0 Then
        With ws.Range(layout.PageSizeCell).Cells(1, 1)
            If managerValidationNumeric.IsNotNumber(.value) Then
                .value = DEFAULT_PAGE_SIZE
            ElseIf CDbl(.value) < 1 Then
                .value = DEFAULT_PAGE_SIZE
            End If
        End With
        managerList.ApplyWholeNumberRule ws.Range(layout.PageSizeCell), 1
    End If

End Sub


' Plant and date range of the search cells (validated).
Public Sub ReadCriteria( _
    ByVal ws As Worksheet, _
    ByVal layout As PagingLayout, _
    ByRef plantName As String, _
    ByRef fromDate As Date, _
    ByRef toDate As Date)

    plantName = managerSheet.ReadText(ws, layout.PlantCell, "LOKASI")

    If Len(layout.FromCell) > 0 Then
        fromDate = managerSheet.ReadDate(ws, layout.FromCell, "DARI TANGGAL")
    End If

    If Len(layout.ToCell) > 0 Then
        toDate = managerSheet.ReadDate(ws, layout.ToCell, "HINGGA TANGGAL")
    End If

End Sub


' =========================================================
' REQUESTS
' =========================================================

' First page with the filter (KOLOM / NILAI) and page size
' (BARIS YANG DITAMPILKAN) of the sheet. filterHeaders lists
' the KOLOM choices; filterIndexes the query column number of
' each one (default: their position). Blank KOLOM = every
' column (index 0).
Public Function FirstPageRequest( _
    ByVal ws As Worksheet, _
    ByVal layout As PagingLayout, _
    Optional ByVal filterHeaders As Variant, _
    Optional ByVal filterIndexes As Variant) As PageRequest

    Dim request As PageRequest
    Dim columnName As String
    Dim i As Long

    Set request = New PageRequest

    request.PageSize = ReadPageSize(ws, layout)
    request.FilterValue = CellText(ws, layout.FilterValueCell)
    request.FilterIndex = 0

    columnName = CellText(ws, layout.FilterColumnCell)

    If Len(columnName) > 0 And Not IsMissing(filterHeaders) Then
        For i = LBound(filterHeaders) To UBound(filterHeaders)
            If managerValidationText.IsSameText(CStr(filterHeaders(i)), columnName) Then
                If IsMissing(filterIndexes) Then
                    request.FilterIndex = i - LBound(filterHeaders) + 1
                Else
                    request.FilterIndex = CLng(filterIndexes(i))
                End If
                Exit For
            End If
        Next i
    End If

    Set FirstPageRequest = request

End Function


' The request for the page delta (-1 / +1) away from the
' current one, and its page number. Nothing (and a status bar
' message) when there is no search result or no such page.
Public Function MoveRequest( _
    ByVal layout As PagingLayout, _
    ByVal delta As Long, _
    ByRef newPage As Long) As PageRequest

    Dim state As PageState

    Set state = FindState(layout.Key)

    If state Is Nothing Then
        Application.StatusBar = "Belum ada hasil pencarian. Klik CARI terlebih dahulu."
        Exit Function
    End If

    newPage = state.Page + delta

    If newPage < 1 Then
        Application.StatusBar = "Sudah di halaman pertama."
        Exit Function
    End If

    If newPage > state.PageCount Then
        Application.StatusBar = "Sudah di halaman terakhir."
        Exit Function
    End If

    Application.StatusBar = False

    If newPage = 1 Then
        Set MoveRequest = state.Request.FirstPage()
    ElseIf delta > 0 Then
        Set MoveRequest = state.Request.Moved(True, state.LastKey)
    Else
        Set MoveRequest = state.Request.Moved(False, state.FirstKey)
    End If

End Function


' The request that loaded the current page (to show it again
' after a change), or Nothing.
Public Function CurrentRequest( _
    ByVal layout As PagingLayout, _
    ByRef pageNumber As Long) As PageRequest

    Dim state As PageState

    Set state = FindState(layout.Key)
    If state Is Nothing Then Exit Function

    pageNumber = state.Page
    Set CurrentRequest = state.Request

End Function


' =========================================================
' STATE
' =========================================================

' Store a loaded page. firstKey / lastKey: the keys of its
' first and last row (Nothing when the page is empty).
Public Sub StorePage( _
    ByVal layout As PagingLayout, _
    ByVal data As Variant, _
    ByVal rowCount As Long, _
    ByVal totalRows As Long, _
    ByVal pageNumber As Long, _
    ByVal request As PageRequest, _
    ByVal firstKey As PageRequest, _
    ByVal lastKey As PageRequest)

    Dim state As PageState

    Set state = New PageState
    state.SetPage data, rowCount, totalRows, pageNumber, request, firstKey, lastKey

    If mStates Is Nothing Then Set mStates = New Collection

    Clear layout.Key
    mStates.Add state, layout.Key

End Sub


Public Function HasState( _
    ByVal key As String) As Boolean

    HasState = Not (FindState(key) Is Nothing)

End Function


' Raises NoSearchResult when nothing was loaded yet.
Public Function GetState( _
    ByVal key As String) As PageState

    Dim state As PageState

    Set state = FindState(key)

    If state Is Nothing Then
        managerError.ThrowError managerPaging.NoSearchResult, _
            MODULE_NAME & ".GetState", _
            "Belum ada hasil pencarian. Klik CARI terlebih dahulu."
    End If

    Set GetState = state

End Function


Public Sub Clear( _
    ByVal key As String)

    If mStates Is Nothing Then Exit Sub

    On Error Resume Next
    mStates.Remove key
    On Error GoTo 0

End Sub


' Page row behind a table row (0 when there is none).
Public Function DataRowOf( _
    ByVal key As String, _
    ByVal tableRowIndex As Long) As Long

    Dim state As PageState

    Set state = FindState(key)

    If state Is Nothing Then
        DataRowOf = 0
    Else
        DataRowOf = state.DataRowAt(tableRowIndex)
    End If

End Function


' Copy what the user typed in the given table columns back
' into the stored page. columnIndexes: Array of table column
' positions.
Public Sub CaptureColumns( _
    ByVal ws As Worksheet, _
    ByVal layout As PagingLayout, _
    ByVal columnIndexes As Variant)

    Dim state As PageState
    Dim tbl As ListObject
    Dim tableRow As Long
    Dim columnIndex As Variant

    Set state = FindState(layout.Key)
    If state Is Nothing Then Exit Sub

    Set tbl = managerTable.GetTable(ws, layout.TableName)

    For tableRow = 1 To state.PageRowCount

        If tableRow > tbl.ListRows.Count Then Exit For

        For Each columnIndex In columnIndexes
            state.CellValue(tableRow, CLng(columnIndex)) = _
                tbl.DataBodyRange.Cells(tableRow, CLng(columnIndex)).value
        Next columnIndex

    Next tableRow

End Sub


' =========================================================
' RENDER
' =========================================================

' Write the stored page into the table and the paging cells
' (BARIS DATA, TOTAL HALAMAN, SEKARANG). Without a stored page
' the table is emptied.
Public Sub Render( _
    ByVal ws As Worksheet, _
    ByVal layout As PagingLayout, _
    Optional ByVal columnFormats As Variant)

    Dim tbl As ListObject
    Dim state As PageState

    Set tbl = managerTable.GetTable(ws, layout.TableName)
    Set state = FindState(layout.Key)

    If state Is Nothing Then
        managerTable.WriteData tbl, Empty, 0, columnFormats
        WriteCell ws, layout.CountCell, 0
        WriteCell ws, layout.PagesCell, 0
        WriteCell ws, layout.PageCell, 0
        Exit Sub
    End If

    managerTable.WriteData tbl, _
        state.PageData(tbl.ListColumns.Count), _
        state.PageRowCount, _
        columnFormats

    WriteCell ws, layout.CountCell, state.TotalRows
    WriteCell ws, layout.PagesCell, state.PageCount
    WriteCell ws, layout.PageCell, state.Page

End Sub


' =========================================================
' EVENTS
' =========================================================

' -1 for SEBELUMNYA, +1 for SELANJUTNYA, 0 otherwise.
Public Function PageButtonDelta( _
    ByVal ws As Worksheet, _
    ByVal layout As PagingLayout, _
    ByVal target As Range) As Long

    If managerSheet.IsHit(target, ws, layout.PrevButton) Then
        PageButtonDelta = -1
    ElseIf managerSheet.IsHit(target, ws, layout.NextButton) Then
        PageButtonDelta = 1
    Else
        PageButtonDelta = 0
    End If

End Function


' True when target is DARI / HINGGA (date picker cells).
Public Function IsDateCell( _
    ByVal ws As Worksheet, _
    ByVal layout As PagingLayout, _
    ByVal target As Range) As Boolean

    IsDateCell = managerSheet.IsHit(target, ws, layout.FromCell) _
        Or managerSheet.IsHit(target, ws, layout.ToCell)

End Function


' True when target is KOLOM, NILAI or BARIS YANG DITAMPILKAN:
' the search must run again from page 1.
Public Function IsViewCell( _
    ByVal ws As Worksheet, _
    ByVal layout As PagingLayout, _
    ByVal target As Range) As Boolean

    IsViewCell = managerSheet.IsHit(target, ws, layout.FilterColumnCell) _
        Or managerSheet.IsHit(target, ws, layout.FilterValueCell) _
        Or managerSheet.IsHit(target, ws, layout.PageSizeCell)

End Function


' =========================================================
' INTERNAL
' =========================================================

Private Function ReadPageSize( _
    ByVal ws As Worksheet, _
    ByVal layout As PagingLayout) As Long

    Dim value As Variant

    ReadPageSize = DEFAULT_PAGE_SIZE

    If Len(layout.PageSizeCell) = 0 Then Exit Function

    value = ws.Range(layout.PageSizeCell).Cells(1, 1).value

    If managerValidationNumeric.IsNumber(value) Then
        If CDbl(value) >= 1 Then ReadPageSize = CLng(value)
    End If

End Function


Private Function CellText( _
    ByVal ws As Worksheet, _
    ByVal cellAddress As String) As String

    Dim value As Variant

    If Len(cellAddress) = 0 Then Exit Function

    value = ws.Range(cellAddress).Cells(1, 1).value

    If IsError(value) Then Exit Function
    CellText = Trim$(coreHelper.TextOrEmpty(value))

End Function


Private Sub WriteCell( _
    ByVal ws As Worksheet, _
    ByVal cellAddress As String, _
    ByVal value As Variant)

    If Len(cellAddress) = 0 Then Exit Sub
    ws.Range(cellAddress).Cells(1, 1).value = value

End Sub


Private Function FindState( _
    ByVal key As String) As PageState

    If mStates Is Nothing Then Exit Function

    On Error Resume Next
    Set FindState = mStates(key)
    On Error GoTo 0

End Function
