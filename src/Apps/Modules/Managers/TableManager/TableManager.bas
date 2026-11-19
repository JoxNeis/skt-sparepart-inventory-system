Attribute VB_Name = "TableManager"
Option Explicit

' =========================================================
' TABLE MANAGER
'
' Every access to an Excel table (ListObject) goes through
' this module, so all sheets read, write, clear and format
' tables the same way.
'
' Row numbers are always the 1-based row index INSIDE the
' table body (ListRow.Index), never the worksheet row.
' Columns are always addressed by their header text.
'
'   Set tbl = TableManager.GetTable(Me, TABLE_NAME)
'   TableManager.ValidateColumns tbl, Array(COLUMN_ID, COLUMN_NAME)
'
'   For rowIndex = 1 To TableManager.RowCount(tbl)
'       nameValue = TableManager.GetText(tbl, rowIndex, COLUMN_NAME)
'   Next rowIndex
'
'   rowIndex = TableManager.AddRow(tbl)
'   TableManager.SetValue tbl, rowIndex, COLUMN_NAME, "Bolt"
'
'   ' Worksheet_SelectionChange
'   If TableManager.IsInColumn(Target, tbl, COLUMN_UPDATE) Then
'       rowIndex = TableManager.GetRowIndex(tbl, Target)
'   End If
' =========================================================

Private Const MODULE_NAME As String = "TableManager"

' =========================================================
' ERROR CODES (10400 - 10499)
' Usage: ErrorManager.ThrowError TableManager.TableNotFound, ...
' =========================================================
Public Const TableNotFound As Long = 10401           ' table name not on the sheet / tbl is Nothing
Public Const ColumnNotFound As Long = 10402          ' header text not in the table
Public Const InvalidRow As Long = 10403              ' row index outside the table body
Public Const CellContainsError As Long = 10404       ' cell holds #N/A, #VALUE!, ...
Public Const InvalidNumber As Long = 10405           ' GetDouble / GetLong on non-numeric value
Public Const InvalidDate As Long = 10406             ' GetDate on a value that is not a date
Public Const EmptyWorksheet As Long = 10407          ' ws parameter is Nothing


' =========================================================
' TABLE
' =========================================================

' Get a table by name or raise TableManager.TableNotFound.
Public Function GetTable( _
    ByVal ws As Worksheet, _
    ByVal tableName As String) As ListObject

    Dim tbl As ListObject

    If ws Is Nothing Then
        ErrorManager.ThrowError _
            TableManager.EmptyWorksheet, _
            MODULE_NAME & ".GetTable", _
            "Worksheet tidak boleh kosong."
    End If

    On Error Resume Next
    Set tbl = ws.ListObjects(tableName)
    On Error GoTo 0

    If tbl Is Nothing Then
        ErrorManager.ThrowError _
            TableManager.TableNotFound, _
            MODULE_NAME & ".GetTable", _
            "Tabel '" & tableName & "' tidak ditemukan " & _
            "pada worksheet '" & ws.Name & "'."
    End If

    Set GetTable = tbl

End Function


' Raise TableManager.TableNotFound when tbl is Nothing.
Public Sub RequireTable( _
    ByVal tbl As ListObject, _
    ByVal errorSource As String)

    If tbl Is Nothing Then
        ErrorManager.ThrowError _
            TableManager.TableNotFound, _
            errorSource, _
            "Tabel belum diinisialisasi."
    End If

End Sub


Public Function RowCount( _
    ByVal tbl As ListObject) As Long

    RequireTable tbl, MODULE_NAME & ".RowCount"

    RowCount = tbl.ListRows.Count

End Function


Public Function HasRows( _
    ByVal tbl As ListObject) As Boolean

    HasRows = (RowCount(tbl) > 0)

End Function


' =========================================================
' COLUMNS
' =========================================================

Public Function HasColumn( _
    ByVal tbl As ListObject, _
    ByVal columnName As String) As Boolean

    Dim targetColumn As ListColumn

    HasColumn = False

    If tbl Is Nothing Then Exit Function

    On Error Resume Next
    Set targetColumn = tbl.ListColumns(columnName)
    On Error GoTo 0

    HasColumn = Not (targetColumn Is Nothing)

End Function


' Raise TableManager.ColumnNotFound for the first header in
' requiredColumns (an Array) that is missing.
Public Sub ValidateColumns( _
    ByVal tbl As ListObject, _
    ByVal requiredColumns As Variant)

    Dim columnName As Variant

    RequireTable tbl, MODULE_NAME & ".ValidateColumns"

    For Each columnName In requiredColumns
        RequireColumn tbl, CStr(columnName), MODULE_NAME & ".ValidateColumns"
    Next columnName

End Sub


' Position of the column inside the table (1-based).
Public Function GetColumnIndex( _
    ByVal tbl As ListObject, _
    ByVal columnName As String) As Long

    RequireColumn tbl, columnName, MODULE_NAME & ".GetColumnIndex"

    GetColumnIndex = tbl.ListColumns(columnName).Index

End Function


' Data cells of one column. Nothing when the table is empty.
Public Function GetColumnRange( _
    ByVal tbl As ListObject, _
    ByVal columnName As String) As Range

    RequireColumn tbl, columnName, MODULE_NAME & ".GetColumnRange"

    If tbl.ListRows.Count = 0 Then
        Set GetColumnRange = Nothing
    Else
        Set GetColumnRange = tbl.ListColumns(columnName).DataBodyRange
    End If

End Function


' Apply number format and/or horizontal alignment to the
' data cells of a column. Does nothing when the table is
' empty. Pass vbNullString / 0 to skip a setting.
'
'   TableManager.FormatColumn tbl, COLUMN_ID, "0", xlCenter
'   TableManager.FormatColumn tbl, COLUMN_NAME, "@"
Public Sub FormatColumn( _
    ByVal tbl As ListObject, _
    ByVal columnName As String, _
    Optional ByVal numberFormat As String = vbNullString, _
    Optional ByVal horizontalAlignment As Long = 0)

    Dim columnRange As Range

    Set columnRange = GetColumnRange(tbl, columnName)

    If columnRange Is Nothing Then Exit Sub

    If Len(numberFormat) > 0 Then
        columnRange.NumberFormat = numberFormat
    End If

    If horizontalAlignment <> 0 Then
        columnRange.HorizontalAlignment = horizontalAlignment
    End If

End Sub


' =========================================================
' ROWS
' =========================================================

' Add an empty row at the bottom. Returns its row index.
Public Function AddRow( _
    ByVal tbl As ListObject) As Long

    Dim newRow As ListRow

    RequireTable tbl, MODULE_NAME & ".AddRow"

    Set newRow = tbl.ListRows.Add

    AddRow = newRow.Index

End Function


' Delete every data row (the table and its header stay).
Public Sub ClearRows( _
    ByVal tbl As ListObject)

    RequireTable tbl, MODULE_NAME & ".ClearRows"

    Do While tbl.ListRows.Count > 0
        tbl.ListRows(tbl.ListRows.Count).Delete
    Loop

End Sub


' Clear the values of every data row (the rows stay).
Public Sub ClearContents( _
    ByVal tbl As ListObject)

    RequireTable tbl, MODULE_NAME & ".ClearContents"

    If tbl.ListRows.Count > 0 Then
        tbl.DataBodyRange.ClearContents
    End If

End Sub


' Delete rows that have no value in any column.
Public Sub DeleteEmptyRows( _
    ByVal tbl As ListObject)

    Dim rowIndex As Long

    RequireTable tbl, MODULE_NAME & ".DeleteEmptyRows"

    For rowIndex = tbl.ListRows.Count To 1 Step -1
        If IsRowEmpty(tbl, rowIndex) Then
            tbl.ListRows(rowIndex).Delete
        End If
    Next rowIndex

End Sub


' True when the row has no value in any column.
Public Function IsRowEmpty( _
    ByVal tbl As ListObject, _
    ByVal rowIndex As Long) As Boolean

    RequireRow tbl, rowIndex, MODULE_NAME & ".IsRowEmpty"

    IsRowEmpty = _
        (Application.WorksheetFunction.CountA( _
            tbl.ListRows(rowIndex).Range) = 0)

End Function


' True when rowIndex is a data row of the table.
Public Function IsValidRow( _
    ByVal tbl As ListObject, _
    ByVal rowIndex As Long) As Boolean

    IsValidRow = False

    If tbl Is Nothing Then Exit Function

    IsValidRow = (rowIndex >= 1 And rowIndex <= tbl.ListRows.Count)

End Function


' =========================================================
' SELECTION (Worksheet_SelectionChange)
' =========================================================

' True when target touches the data cells of the column.
Public Function IsInColumn( _
    ByVal target As Range, _
    ByVal tbl As ListObject, _
    ByVal columnName As String) As Boolean

    IsInColumn = False

    If target Is Nothing Then Exit Function
    If tbl Is Nothing Then Exit Function
    If tbl.ListRows.Count = 0 Then Exit Function

    IsInColumn = FormManager.IsInRange( _
        target, GetColumnRange(tbl, columnName))

End Function


' True when target touches the table body.
Public Function IsInTable( _
    ByVal target As Range, _
    ByVal tbl As ListObject) As Boolean

    IsInTable = False

    If target Is Nothing Then Exit Function
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function

    IsInTable = FormManager.IsInRange(target, tbl.DataBodyRange)

End Function


' Table row index of the first cell of target.
' Returns 0 when target is outside the table body.
Public Function GetRowIndex( _
    ByVal tbl As ListObject, _
    ByVal target As Range) As Long

    Dim rowIndex As Long

    GetRowIndex = 0

    If Not IsInTable(target, tbl) Then Exit Function

    rowIndex = target.Cells(1, 1).Row - tbl.DataBodyRange.Row + 1

    If IsValidRow(tbl, rowIndex) Then
        GetRowIndex = rowIndex
    End If

End Function


' =========================================================
' CELLS
' =========================================================

Public Function GetCell( _
    ByVal tbl As ListObject, _
    ByVal rowIndex As Long, _
    ByVal columnName As String) As Range

    RequireRow tbl, rowIndex, MODULE_NAME & ".GetCell"
    RequireColumn tbl, columnName, MODULE_NAME & ".GetCell"

    Set GetCell = tbl.ListRows(rowIndex).Range.Cells( _
        1, tbl.ListColumns(columnName).Index)

End Function


' Raw value (Empty, text, number, date or error value).
Public Function GetValue( _
    ByVal tbl As ListObject, _
    ByVal rowIndex As Long, _
    ByVal columnName As String) As Variant

    GetValue = GetCell(tbl, rowIndex, columnName).value

End Function


' True when the cell is not blank and not an error value.
Public Function HasValue( _
    ByVal tbl As ListObject, _
    ByVal rowIndex As Long, _
    ByVal columnName As String) As Boolean

    Dim value As Variant

    value = GetValue(tbl, rowIndex, columnName)

    If IsError(value) Then
        HasValue = False
    Else
        HasValue = ValidationText.IsNotBlank(value)
    End If

End Function


' Trimmed text. Empty -> "". Error value -> CellContainsError.
Public Function GetText( _
    ByVal tbl As ListObject, _
    ByVal rowIndex As Long, _
    ByVal columnName As String) As String

    Dim value As Variant

    value = GetValue(tbl, rowIndex, columnName)

    RequireNoCellError value, tbl, rowIndex, columnName, MODULE_NAME & ".GetText"

    If IsNull(value) Or IsEmpty(value) Then
        GetText = vbNullString
    Else
        GetText = Trim$(CStr(value))
    End If

End Function


' Number. Blank -> 0. Not numeric -> TableManager.InvalidNumber.
Public Function GetDouble( _
    ByVal tbl As ListObject, _
    ByVal rowIndex As Long, _
    ByVal columnName As String) As Double

    Dim value As Variant

    value = GetValue(tbl, rowIndex, columnName)

    RequireNoCellError value, tbl, rowIndex, columnName, MODULE_NAME & ".GetDouble"

    If ValidationText.IsBlank(value) Then
        GetDouble = 0
        Exit Function
    End If

    If ValidationNumeric.IsNotNumber(value) Then
        ErrorManager.ThrowError _
            TableManager.InvalidNumber, _
            MODULE_NAME & ".GetDouble", _
            "Kolom '" & columnName & "' baris ke-" & _
            CStr(rowIndex) & " harus berupa angka."
    End If

    GetDouble = CDbl(value)

End Function


' Whole number. Blank -> 0. Not numeric -> InvalidNumber.
Public Function GetLong( _
    ByVal tbl As ListObject, _
    ByVal rowIndex As Long, _
    ByVal columnName As String) As Long

    GetLong = CLng(GetDouble(tbl, rowIndex, columnName))

End Function


' Date. Blank -> 0 (empty date). Not a date -> InvalidDate.
Public Function GetDate( _
    ByVal tbl As ListObject, _
    ByVal rowIndex As Long, _
    ByVal columnName As String) As Date

    Dim value As Variant

    value = GetValue(tbl, rowIndex, columnName)

    RequireNoCellError value, tbl, rowIndex, columnName, MODULE_NAME & ".GetDate"

    If ValidationText.IsBlank(value) Then
        GetDate = 0
        Exit Function
    End If

    If Not IsDate(value) Then
        ErrorManager.ThrowError _
            TableManager.InvalidDate, _
            MODULE_NAME & ".GetDate", _
            "Kolom '" & columnName & "' baris ke-" & _
            CStr(rowIndex) & " harus berupa tanggal."
    End If

    GetDate = CDate(value)

End Function


Public Sub SetValue( _
    ByVal tbl As ListObject, _
    ByVal rowIndex As Long, _
    ByVal columnName As String, _
    ByVal value As Variant)

    GetCell(tbl, rowIndex, columnName).value = value

End Sub


Public Sub SelectCell( _
    ByVal tbl As ListObject, _
    ByVal rowIndex As Long, _
    ByVal columnName As String)

    GetCell(tbl, rowIndex, columnName).Select

End Sub


' =========================================================
' INTERNAL
' =========================================================

Private Sub RequireColumn( _
    ByVal tbl As ListObject, _
    ByVal columnName As String, _
    ByVal errorSource As String)

    RequireTable tbl, errorSource

    If Not HasColumn(tbl, columnName) Then
        ErrorManager.ThrowError _
            TableManager.ColumnNotFound, _
            errorSource, _
            "Kolom '" & columnName & _
            "' tidak ditemukan pada tabel '" & tbl.Name & "'."
    End If

End Sub


Private Sub RequireRow( _
    ByVal tbl As ListObject, _
    ByVal rowIndex As Long, _
    ByVal errorSource As String)

    RequireTable tbl, errorSource

    If Not IsValidRow(tbl, rowIndex) Then
        ErrorManager.ThrowError _
            TableManager.InvalidRow, _
            errorSource, _
            "Baris ke-" & CStr(rowIndex) & _
            " tidak ada pada tabel '" & tbl.Name & "'."
    End If

End Sub


Private Sub RequireNoCellError( _
    ByVal value As Variant, _
    ByVal tbl As ListObject, _
    ByVal rowIndex As Long, _
    ByVal columnName As String, _
    ByVal errorSource As String)

    If IsError(value) Then
        ErrorManager.ThrowError _
            TableManager.CellContainsError, _
            errorSource, _
            "Kolom '" & columnName & "' baris ke-" & _
            CStr(rowIndex) & " pada tabel '" & tbl.Name & _
            "' berisi nilai error."
    End If

End Sub
