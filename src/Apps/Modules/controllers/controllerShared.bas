Attribute VB_Name = "controllerShared"
Option Explicit

' =========================================================
' CONTROLLER - SHARED
'
' Pieces used by several controllers:
'   - CEK SPAREPART box (INPUT_ORDER, INPUT_REQUEST)
'   - input tables (labels, dropdowns, SATUAN, delete row)
'   - EDIT mode of recap rows (REKAP_ORDER, REKAP_PENERIMAAN,
'     REKAP_REQUEST):
'         EDIT   -> double-click: the row becomes editable
'                   (yellow), EDIT shows AJUKAN, HAPUS shows
'                   BATAL
'         AJUKAN -> double-click: asks the reason and saves
'                   the update proposal
'         BATAL  -> double-click: restores the row
' =========================================================

Private Const MODULE_NAME As String = "controllerShared"

' =========================================================
' ERROR CODES (14000 - 14099)
' =========================================================
Public Const EmptySearchText As Long = 14001
Public Const UnknownSparepart As Long = 14002
Public Const UnknownVendor As Long = 14003
Public Const UnknownMobileMaintenance As Long = 14004
Public Const StaleSearchResult As Long = 14005      ' cache lost / row not found

Private Const EDIT_RULE_TAG As String = "SIS_EDIT_ROW"


' =========================================================
' CEK SPAREPART
' =========================================================

' Search the catalogue for the text in inputCell and show the
' first match (NAMA, KODE, APPLICABLE, SATUAN, STOK SEKARANG)
' in the 5 cells starting at resultFirstCell. Stock is of
' the plant in plantCell.
Public Sub SearchSparepart( _
    ByVal ws As Worksheet, _
    ByVal plantCell As String, _
    ByVal inputCell As String, _
    ByVal resultFirstCell As String)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo
    Dim searchText As String
    Dim plantName As String
    Dim results As Collection
    Dim sparepartItem As Sparepart
    Dim stockItem As ManufacturingPlantStock
    Dim stockValue As Variant

    searchText = managerForm.GetCellText(ws, inputCell)

    If Len(searchText) = 0 Then
        managerError.ThrowError controllerShared.EmptySearchText, _
            MODULE_NAME & ".SearchSparepart", _
            "Isi NAMA DISPLAY yang dicari (nama, kode, alias atau applicable)."
    End If

    plantName = managerForm.GetCellText(ws, plantCell)

    state = managerForm.BeginBusy("Mencari sparepart...")

    ws.Range(resultFirstCell).Resize(1, 5).ClearContents

    Set results = serviceSparepart.SearchInList( _
        serviceSparepart.ListSpareparts(), searchText)

    If results.Count = 0 Then
        managerForm.EndBusy state
        managerForm.ShowWarning "Sparepart '" & searchText & "' tidak ditemukan."
        Exit Sub
    End If

    Set sparepartItem = results(1)

    stockValue = "-"

    If Len(plantName) > 0 Then
        Set stockItem = servicePlantStock.FindInList( _
            servicePlantStock.ListPlantStocks(plantName), sparepartItem.Id)
        If Not stockItem Is Nothing Then stockValue = stockItem.CurrentStock
    End If

    With ws.Range(resultFirstCell)
        .Offset(0, 0).value = sparepartItem.Name
        .Offset(0, 1).NumberFormat = managerSheet.TEXT_FORMAT
        .Offset(0, 1).value = sparepartItem.Code
        .Offset(0, 2).value = sparepartItem.Applicable
        .Offset(0, 3).value = sparepartItem.UnitOfMeasurement
        .Offset(0, 4).value = stockValue
    End With

    managerForm.EndBusy state

    If results.Count > 1 Then
        Application.StatusBar = results.Count & _
            " sparepart cocok; yang pertama ditampilkan. Perjelas pencarian untuk hasil lain."
    End If

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


' =========================================================
' INPUT TABLES
' =========================================================

' Write SATUAN for every changed cell of the sparepart column.
Public Sub FillUnits( _
    ByVal tbl As ListObject, _
    ByVal target As Range, _
    ByVal sparepartColumn As String, _
    ByVal unitColumn As String)

    Dim changed As Range
    Dim cell As Range
    Dim rowIndex As Long
    Dim eventsWereOn As Boolean

    If tbl.DataBodyRange Is Nothing Then Exit Sub

    Set changed = Intersect(target, managerTable.GetColumnRange(tbl, sparepartColumn))
    If changed Is Nothing Then Exit Sub

    eventsWereOn = Application.EnableEvents
    Application.EnableEvents = False

    For Each cell In changed.Cells
        rowIndex = managerTable.GetRowIndex(tbl, cell)
        If rowIndex > 0 Then
            managerTable.SetValue tbl, rowIndex, unitColumn, _
                managerList.SparepartUnit(coreHelper.TextOrEmpty(cell.value))
        End If
    Next cell

    Application.EnableEvents = eventsWereOn

End Sub


' True when one of columns has a value on the row.
Public Function RowHasInput( _
    ByVal tbl As ListObject, _
    ByVal rowIndex As Long, _
    ByVal columns As Variant) As Boolean

    Dim columnName As Variant

    For Each columnName In columns
        If managerTable.HasValue(tbl, rowIndex, CStr(columnName)) Then
            RowHasInput = True
            Exit Function
        End If
    Next columnName

End Function


' Delete one input row (the last row is cleared) and keep the
' action label on every row.
Public Sub DeleteInputRow( _
    ByVal tbl As ListObject, _
    ByVal rowIndex As Long, _
    ByVal labelColumn As String, _
    ByVal labelText As String)

    managerTable.DeleteRow tbl, rowIndex
    managerSheet.FillColumn tbl, labelColumn, labelText

End Sub


' =========================================================
' LOOKUPS WITH A CLEAR ERROR
' =========================================================

Public Function RequireSparepart( _
    ByVal spareparts As Collection, _
    ByVal text As String, _
    ByVal sourceCell As Range) As Sparepart

    Dim result As Sparepart

    Set result = serviceSparepart.FindInList(spareparts, text)

    If result Is Nothing Then
        managerSheet.FailCell sourceCell, controllerShared.UnknownSparepart, _
            "Sparepart '" & text & "' tidak ditemukan di MASTER DATA SPAREPART. " & _
            "Pilih dari daftar."
    End If

    Set RequireSparepart = result

End Function

' Blank text -> Nothing (vendor is optional).
Public Function FindVendor( _
    ByVal vendors As Collection, _
    ByVal text As String, _
    ByVal sourceCell As Range) As Vendor

    Dim result As Vendor

    If Len(Trim$(text)) = 0 Then Exit Function

    Set result = serviceVendor.FindInList(vendors, text)

    If result Is Nothing Then
        managerSheet.FailCell sourceCell, controllerShared.UnknownVendor, _
            "Vendor '" & text & "' tidak ditemukan di MASTER DATA VENDOR. " & _
            "Pilih dari daftar."
    End If

    Set FindVendor = result

End Function

Public Function RequireMobileMaintenance( _
    ByVal items As Collection, _
    ByVal text As String, _
    ByVal plantName As String, _
    ByVal sourceCell As Range) As MobileMaintenance

    Dim result As MobileMaintenance

    Set result = serviceMobileMaintenance.FindInList(items, text)

    If result Is Nothing Then
        managerSheet.FailCell sourceCell, controllerShared.UnknownMobileMaintenance, _
            "Pos mobile '" & text & "' tidak ditemukan di plant " & plantName & _
            ". Pilih dari daftar."
    End If

    Set RequireMobileMaintenance = result

End Function


' =========================================================
' EDIT MODE OF RECAP ROWS
' =========================================================

' Yellow fill on the editable cells of rows in edit mode
' (conditional format, so it never leaks into the design).
Public Sub EnsureEditHighlight( _
    ByVal tbl As ListObject, _
    ByVal editColumn As String, _
    ByVal editableColumns As Variant)

    managerTable.EnsureRowHighlight tbl, EDIT_RULE_TAG, editColumn, _
        managerSheet.LABEL_SUBMIT, editableColumns, managerSheet.EDIT_FILL_COLOR

End Sub


Public Function IsEditRow( _
    ByVal tbl As ListObject, _
    ByVal rowIndex As Long, _
    ByVal editColumn As String) As Boolean

    If Not managerTable.IsValidRow(tbl, rowIndex) Then Exit Function

    IsEditRow = (managerTable.GetText(tbl, rowIndex, editColumn) = _
        managerSheet.LABEL_SUBMIT)

End Function


' EDIT -> AJUKAN, HAPUS -> BATAL, and select the first
' editable cell of the row.
Public Sub EnterEditMode( _
    ByVal tbl As ListObject, _
    ByVal rowIndex As Long, _
    ByVal editColumn As String, _
    ByVal deleteColumn As String, _
    ByVal firstEditableColumn As String)

    Dim eventsWereOn As Boolean

    eventsWereOn = Application.EnableEvents
    Application.EnableEvents = False

    managerTable.SetValue tbl, rowIndex, editColumn, managerSheet.LABEL_SUBMIT
    managerTable.SetValue tbl, rowIndex, deleteColumn, managerSheet.LABEL_CANCEL

    Application.EnableEvents = eventsWereOn

    managerTable.SelectCell tbl, rowIndex, firstEditableColumn

    Application.StatusBar = "Ubah sel kuning pada baris ini, lalu double-click " & _
        managerSheet.LABEL_SUBMIT & " untuk mengajukan update (atau " & _
        managerSheet.LABEL_CANCEL & " untuk membatalkan)."

End Sub


' Undo edits outside the editable cells of rows in edit
' mode (and outside alwaysEditableColumns, e.g. PO ID).
' Returns True when the edit was undone.
Public Function GuardRecapChange( _
    ByVal tbl As ListObject, _
    ByVal target As Range, _
    ByVal editColumn As String, _
    ByVal editableColumns As Variant, _
    Optional ByVal alwaysEditableColumns As Variant) As Boolean

    Dim allowed As Range
    Dim rowIndex As Long
    Dim columnName As Variant

    If tbl.DataBodyRange Is Nothing Then Exit Function
    If Intersect(target, tbl.DataBodyRange) Is Nothing Then Exit Function

    If Not IsMissing(alwaysEditableColumns) Then
        For Each columnName In alwaysEditableColumns
            Set allowed = UnionOf(allowed, managerTable.GetColumnRange(tbl, CStr(columnName)))
        Next columnName
    End If

    For rowIndex = 1 To managerTable.RowCount(tbl)
        If IsEditRow(tbl, rowIndex, editColumn) Then
            For Each columnName In editableColumns
                Set allowed = UnionOf(allowed, _
                    managerTable.GetCell(tbl, rowIndex, CStr(columnName)))
            Next columnName
        End If
    Next rowIndex

    GuardRecapChange = managerSheet.UndoIfReadOnly( _
        target, tbl.DataBodyRange, allowed, _
        "Baris ini hanya bisa dibaca. Double-click EDIT untuk mengubah.")

End Function


' Ask the reason of an update / delete proposal.
' Returns False when the user cancelled.
Public Function AskReason( _
    ByVal message As String, _
    ByRef reason As String) As Boolean

    Dim cancelled As Boolean

    Do
        reason = managerForm.AskText(message, cancelled, "Alasan")

        If cancelled Then
            AskReason = False
            Exit Function
        End If

        If Len(reason) > 0 Then
            AskReason = True
            Exit Function
        End If

        managerForm.ShowWarning "Alasan wajib diisi."
    Loop

End Function


' =========================================================
' INTERNAL
' =========================================================

Private Function UnionOf( _
    ByVal first As Range, _
    ByVal second As Range) As Range

    If first Is Nothing Then
        Set UnionOf = second
    ElseIf second Is Nothing Then
        Set UnionOf = first
    Else
        Set UnionOf = Union(first, second)
    End If

End Function
