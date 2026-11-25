Attribute VB_Name = "managerSheet"
Option Explicit

' =========================================================
' SHEET MANAGER
'
' Helpers shared by the controllers:
'   - sheet names of the workbook
'   - "buttons": clickable cells are double-clicked
'   - reading required / optional input cells with clear
'     messages (text, whole numbers, dates)
'   - the application location (SETTINGS > LOKASI)
'   - Indonesian month names
'   - undoing edits in read-only cells
' =========================================================

Private Const MODULE_NAME As String = "managerSheet"

' =========================================================
' ERROR CODES (10500 - 10599)
' Usage: managerError.ThrowError managerSheet.EmptyRequiredCell, ...
' =========================================================
Public Const EmptyRequiredCell As Long = 10501
Public Const InvalidDateCell As Long = 10502
Public Const InvalidNumberCell As Long = 10503
Public Const SheetNotFound As Long = 10504
Public Const InvalidMonthName As Long = 10505

' =========================================================
' SHEET NAMES
' =========================================================
Public Const SHEET_DASHBOARD As String = "DASHBOARD"
Public Const SHEET_INPUT_ORDER As String = "INPUT_ORDER"
Public Const SHEET_REKAP_ORDER As String = "REKAP_ORDER"
Public Const SHEET_REKAP_PENDING_ORDER As String = "REKAP_PENDING_ORDER"
Public Const SHEET_INPUT_PENERIMAAN As String = "INPUT_PENERIMAAN"
Public Const SHEET_REKAP_PENERIMAAN As String = "REKAP_PENERIMAAN"
Public Const SHEET_INPUT_REQUEST As String = "INPUT_REQUEST"
Public Const SHEET_REKAP_REQUEST As String = "REKAP_REQUEST"
Public Const SHEET_KARTU_STOCK As String = "KARTU_STOCK"
Public Const SHEET_REKAP_INVENTORY As String = "REKAP_INVENTORY"
Public Const SHEET_MD_SPAREPART As String = "MD_SPAREPART"
Public Const SHEET_MD_POS_MOBILE As String = "MD_POS_MOBILE"
Public Const SHEET_MD_VENDOR As String = "MD_VENDOR"
Public Const SHEET_MD_USER As String = "MD_USER"
Public Const SHEET_REQUEST_UPDATE As String = "REQUEST_UPDATE"
Public Const SHEET_SETTINGS As String = "SETTINGS"

' SETTINGS > APLIKASI > LOKASI (merged L8:M8)
Public Const SETTINGS_PLANT_CELL As String = "L8"

' Number and date formats used on every sheet.
Public Const DATE_FORMAT As String = "dd-mmm-yyyy"
Public Const TEXT_FORMAT As String = "@"
Public Const NUMBER_FORMAT As String = "#,##0"
Public Const DECIMAL_FORMAT As String = "#,##0.00"

' Labels written into the action columns of tables.
Public Const LABEL_EDIT As String = "EDIT"
Public Const LABEL_SUBMIT As String = "AJUKAN"
Public Const LABEL_DELETE As String = "HAPUS"
Public Const LABEL_CANCEL As String = "BATAL"
Public Const LABEL_SAVE As String = "SIMPAN"
Public Const LABEL_SAVE_PO As String = "SIMPAN PO"
Public Const LABEL_CHECK As String = "CHECK"
Public Const LABEL_SELECTED As String = "DIPILIH"

' Hidden sheet-level name holding the id edited in an MD form.
Private Const EDIT_ID_NAME As String = "SIS_EDIT_ID"

' Light yellow used for editable cells (same as the design).
Public Const EDIT_FILL_COLOR As Long = 13431551     ' RGB(255, 242, 204)


' =========================================================
' SHEETS
' =========================================================

Public Function GetSheet( _
    ByVal sheetName As String) As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        managerError.ThrowError managerSheet.SheetNotFound, _
            MODULE_NAME & ".GetSheet", _
            "Sheet '" & sheetName & "' tidak ditemukan."
    End If

    Set GetSheet = ws

End Function


' True when the (first cell of the) double-clicked target is
' inside the button area of the sheet.
Public Function IsHit( _
    ByVal target As Range, _
    ByVal ws As Worksheet, _
    ByVal areaAddress As String) As Boolean

    IsHit = False
    If Len(areaAddress) = 0 Then Exit Function

    IsHit = managerForm.IsInRange(target.Cells(1, 1), ws.Range(areaAddress))

End Function


' =========================================================
' LOCATION (SETTINGS > LOKASI)
' =========================================================

' The application location, "" when not set.
Public Function SettingsPlantName() As String

    Dim value As Variant

    value = GetSheet(SHEET_SETTINGS).Range(SETTINGS_PLANT_CELL).value

    If IsError(value) Then Exit Function
    SettingsPlantName = Trim$(coreHelper.TextOrEmpty(value))

End Function


' Put the plant dropdown on a LOKASI cell and fill it with
' the application location when it is still blank.
Public Sub PreparePlantCell( _
    ByVal ws As Worksheet, _
    ByVal cellAddress As String)

    managerList.ApplyDropdown ws.Range(cellAddress), managerList.LIST_PLANT

    If managerValidationText.IsBlank(ws.Range(cellAddress).Cells(1, 1).value) Then
        ws.Range(cellAddress).Cells(1, 1).value = SettingsPlantName()
    End If

End Sub


' Write a default date into a blank date cell.
Public Sub PrepareDateCell( _
    ByVal ws As Worksheet, _
    ByVal cellAddress As String, _
    ByVal defaultDate As Date)

    With ws.Range(cellAddress).Cells(1, 1)
        If managerValidationText.IsBlank(.value) Then
            .value = defaultDate
        End If
        .NumberFormat = DATE_FORMAT
    End With

End Sub


' =========================================================
' READING INPUT CELLS
' =========================================================

' Trimmed text of a cell. Raises EmptyRequiredCell when it is
' blank and required.
Public Function ReadText( _
    ByVal ws As Worksheet, _
    ByVal cellAddress As String, _
    ByVal label As String, _
    Optional ByVal required As Boolean = True) As String

    Dim result As String

    result = managerForm.GetCellText(ws, cellAddress)

    If required And Len(result) = 0 Then
        SelectCell ws.Range(cellAddress)
        managerError.ThrowError managerSheet.EmptyRequiredCell, _
            MODULE_NAME & ".ReadText", _
            label & " wajib diisi."
    End If

    ReadText = result

End Function


' Date of a cell. Blank -> 0 when not required.
Public Function ReadDate( _
    ByVal ws As Worksheet, _
    ByVal cellAddress As String, _
    ByVal label As String, _
    Optional ByVal required As Boolean = True) As Date

    ReadDate = ParseDate( _
        ws.Range(cellAddress).Cells(1, 1).value, label, required, _
        ws.Range(cellAddress))

End Function


' Whole number >= minimum of a cell. Blank -> blankValue
' when not required.
Public Function ReadWholeNumber( _
    ByVal ws As Worksheet, _
    ByVal cellAddress As String, _
    ByVal label As String, _
    Optional ByVal minimum As Long = 0, _
    Optional ByVal required As Boolean = True, _
    Optional ByVal blankValue As Long = 0) As Long

    ReadWholeNumber = ParseWholeNumber( _
        ws.Range(cellAddress).Cells(1, 1).value, label, minimum, _
        required, blankValue, ws.Range(cellAddress))

End Function


' Value -> date. Raises InvalidDateCell / EmptyRequiredCell.
Public Function ParseDate( _
    ByVal value As Variant, _
    ByVal label As String, _
    Optional ByVal required As Boolean = True, _
    Optional ByVal sourceCell As Range) As Date

    If IsError(value) Then
        FailCell sourceCell, managerSheet.InvalidDateCell, label & " berisi nilai error."
    End If

    If managerValidationText.IsBlank(value) Then
        If required Then
            FailCell sourceCell, managerSheet.EmptyRequiredCell, _
                label & " wajib diisi (double-click untuk memilih tanggal)."
        End If
        ParseDate = 0
        Exit Function
    End If

    If Not IsDate(value) Then
        FailCell sourceCell, managerSheet.InvalidDateCell, _
            label & " harus berupa tanggal (double-click untuk memilih tanggal)."
    End If

    ParseDate = DateValue(CDate(value))

End Function


' Value -> whole number >= minimum.
Public Function ParseWholeNumber( _
    ByVal value As Variant, _
    ByVal label As String, _
    Optional ByVal minimum As Long = 0, _
    Optional ByVal required As Boolean = True, _
    Optional ByVal blankValue As Long = 0, _
    Optional ByVal sourceCell As Range) As Long

    If IsError(value) Then
        FailCell sourceCell, managerSheet.InvalidNumberCell, label & " berisi nilai error."
    End If

    If managerValidationText.IsBlank(value) Then
        If required Then
            FailCell sourceCell, managerSheet.EmptyRequiredCell, label & " wajib diisi."
        End If
        ParseWholeNumber = blankValue
        Exit Function
    End If

    If managerValidationNumeric.IsNotNumber(value) Then
        FailCell sourceCell, managerSheet.InvalidNumberCell, label & " harus berupa angka."
    End If

    If Not managerValidationNumeric.IsWholeNumber(CDbl(value)) Then
        FailCell sourceCell, managerSheet.InvalidNumberCell, label & " harus bilangan bulat."
    End If

    If CDbl(value) < minimum Then
        FailCell sourceCell, managerSheet.InvalidNumberCell, _
            label & " minimal " & minimum & "."
    End If

    ParseWholeNumber = CLng(value)

End Function


' Value -> number >= 0 (decimals allowed).
Public Function ParseDecimal( _
    ByVal value As Variant, _
    ByVal label As String, _
    Optional ByVal sourceCell As Range) As Double

    If IsError(value) Then
        FailCell sourceCell, managerSheet.InvalidNumberCell, label & " berisi nilai error."
    End If

    If managerValidationText.IsBlank(value) Then
        ParseDecimal = 0
        Exit Function
    End If

    If managerValidationNumeric.IsNotNumber(value) Then
        FailCell sourceCell, managerSheet.InvalidNumberCell, label & " harus berupa angka."
    End If

    If CDbl(value) < 0 Then
        FailCell sourceCell, managerSheet.InvalidNumberCell, label & " tidak boleh negatif."
    End If

    ParseDecimal = CDbl(value)

End Function


' Table cell -> text / whole number / date with the table
' row number in the message.
Public Function TableText( _
    ByVal tbl As ListObject, _
    ByVal rowIndex As Long, _
    ByVal columnName As String) As String

    TableText = managerTable.GetText(tbl, rowIndex, columnName)

End Function

Public Function TableWholeNumber( _
    ByVal tbl As ListObject, _
    ByVal rowIndex As Long, _
    ByVal columnName As String, _
    Optional ByVal minimum As Long = 0, _
    Optional ByVal required As Boolean = True, _
    Optional ByVal blankValue As Long = 0) As Long

    TableWholeNumber = ParseWholeNumber( _
        managerTable.GetValue(tbl, rowIndex, columnName), _
        RowLabel(rowIndex, columnName), minimum, required, blankValue, _
        managerTable.GetCell(tbl, rowIndex, columnName))

End Function

Public Function TableDate( _
    ByVal tbl As ListObject, _
    ByVal rowIndex As Long, _
    ByVal columnName As String, _
    Optional ByVal required As Boolean = True) As Date

    TableDate = ParseDate( _
        managerTable.GetValue(tbl, rowIndex, columnName), _
        RowLabel(rowIndex, columnName), required, _
        managerTable.GetCell(tbl, rowIndex, columnName))

End Function

Public Function TableDecimal( _
    ByVal tbl As ListObject, _
    ByVal rowIndex As Long, _
    ByVal columnName As String) As Double

    TableDecimal = ParseDecimal( _
        managerTable.GetValue(tbl, rowIndex, columnName), _
        RowLabel(rowIndex, columnName), _
        managerTable.GetCell(tbl, rowIndex, columnName))

End Function

' Raise EmptyRequiredCell for a blank table cell.
Public Function TableRequiredText( _
    ByVal tbl As ListObject, _
    ByVal rowIndex As Long, _
    ByVal columnName As String) As String

    Dim result As String

    result = managerTable.GetText(tbl, rowIndex, columnName)

    If Len(result) = 0 Then
        FailCell managerTable.GetCell(tbl, rowIndex, columnName), _
            managerSheet.EmptyRequiredCell, _
            RowLabel(rowIndex, columnName) & " wajib diisi."
    End If

    TableRequiredText = result

End Function


' Raise errorCode with message, after selecting the cell
' (when it is on the active sheet), so the user sees where.
Public Sub FailCell( _
    ByVal sourceCell As Range, _
    ByVal errorCode As Long, _
    ByVal message As String)

    SelectCell sourceCell

    managerError.ThrowError errorCode, MODULE_NAME, message

End Sub


' =========================================================
' TABLE ACTION LABELS
' =========================================================

' Write labelText into every body row of an action column.
Public Sub FillColumn( _
    ByVal tbl As ListObject, _
    ByVal columnName As String, _
    ByVal labelText As String)

    Dim columnRange As Range

    Set columnRange = managerTable.GetColumnRange(tbl, columnName)
    If columnRange Is Nothing Then Exit Sub

    columnRange.value = labelText

End Sub


' =========================================================
' READ-ONLY CELLS
' =========================================================

' For Worksheet_Change: when target touches a protected area
' and is not inside an allowed area, undo the user's edit.
' Returns True when the edit was undone.
Public Function UndoIfReadOnly( _
    ByVal target As Range, _
    ByVal protectedArea As Range, _
    Optional ByVal allowedArea As Range, _
    Optional ByVal message As String = vbNullString) As Boolean

    Dim cell As Range

    UndoIfReadOnly = False

    If protectedArea Is Nothing Then Exit Function
    If Intersect(target, protectedArea) Is Nothing Then Exit Function

    For Each cell In Intersect(target, protectedArea).Cells

        If allowedArea Is Nothing Then
            UndoIfReadOnly = True
        ElseIf Intersect(cell, allowedArea) Is Nothing Then
            UndoIfReadOnly = True
        End If

        If UndoIfReadOnly Then Exit For

    Next cell

    If Not UndoIfReadOnly Then Exit Function

    Application.EnableEvents = False
    On Error Resume Next
    Application.Undo
    On Error GoTo 0
    Application.EnableEvents = True

    If Len(message) = 0 Then
        message = "Kolom ini hanya bisa dibaca."
    End If

    Application.StatusBar = message

End Function


' =========================================================
' EDIT MODE OF A MASTER DATA FORM
'
' The id of the row being edited is kept in a hidden,
' sheet-level name, so it survives a VBA reset.
' 0 = the form creates a new row.
' =========================================================

Public Function GetEditId( _
    ByVal ws As Worksheet) As Long

    Dim refersToText As String

    On Error Resume Next
    refersToText = ws.Names(EDIT_ID_NAME).RefersTo
    On Error GoTo 0

    If Len(refersToText) > 1 Then
        GetEditId = CLng(Val(Mid$(refersToText, 2)))
    End If

End Function

Public Sub SetEditId( _
    ByVal ws As Worksheet, _
    ByVal editId As Long)

    ws.Names.Add Name:=EDIT_ID_NAME, RefersTo:="=" & CStr(editId), Visible:=False

End Sub


' =========================================================
' MONTHS (Indonesian, upper case)
' =========================================================

Public Function IndonesianMonths() As Variant

    IndonesianMonths = Array("JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", _
        "JUNI", "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", _
        "DESEMBER")

End Function

Public Function IndonesianMonth( _
    ByVal monthNumber As Long) As String

    Dim names As Variant

    names = IndonesianMonths()
    IndonesianMonth = CStr(names(monthNumber - 1))

End Function

' "JANUARI" -> 1. Also accepts a number 1 - 12.
Public Function ParseMonth( _
    ByVal value As Variant, _
    Optional ByVal sourceCell As Range) As Long

    Dim names As Variant
    Dim i As Long

    If managerValidationNumeric.IsNumber(value) Then
        If CLng(value) >= 1 And CLng(value) <= 12 Then
            ParseMonth = CLng(value)
            Exit Function
        End If
    End If

    names = IndonesianMonths()

    For i = LBound(names) To UBound(names)
        If managerValidationText.IsSameText(CStr(names(i)), Trim$(coreHelper.TextOrEmpty(value))) Then
            ParseMonth = i + 1
            Exit Function
        End If
    Next i

    FailCell sourceCell, managerSheet.InvalidMonthName, _
        "Bulan '" & coreHelper.TextOrEmpty(value) & "' tidak dikenal."

End Function


' "NOVEMBER 2026"
Public Function PeriodText( _
    ByVal value As Date) As String

    PeriodText = IndonesianMonth(Month(value)) & " " & Year(value)

End Function


' =========================================================
' INTERNAL
' =========================================================

' (SelectCell is Public: selects a cell only when its sheet
' is the active sheet, never raises.)

Private Function RowLabel( _
    ByVal rowIndex As Long, _
    ByVal columnName As String) As String

    RowLabel = "Baris ke-" & rowIndex & ", kolom " & columnName

End Function


Public Sub SelectCell( _
    ByVal sourceCell As Range)

    If sourceCell Is Nothing Then Exit Sub

    On Error Resume Next
    If sourceCell.Worksheet Is ActiveSheet Then
        sourceCell.Cells(1, 1).Select
    End If
    On Error GoTo 0

End Sub
