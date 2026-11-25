Attribute VB_Name = "managerTable"
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
'   Set tbl = managerTable.GetTable(Me, TABLE_NAME)
'   managerTable.ValidateColumns tbl, Array(COLUMN_ID, COLUMN_NAME)
'
'   For rowIndex = 1 To managerTable.RowCount(tbl)
'       nameValue = managerTable.GetText(tbl, rowIndex, COLUMN_NAME)
'   Next rowIndex
'
'   rowIndex = managerTable.AddRow(tbl)
'   managerTable.SetValue tbl, rowIndex, COLUMN_NAME, "Bolt"
'
'   ' Worksheet_SelectionChange
'   If managerTable.IsInColumn(Target, tbl, COLUMN_UPDATE) Then
'       rowIndex = managerTable.GetRowIndex(tbl, Target)
'   End If
' =========================================================

Private Const MODULE_NAME As String = "managerTable"

' =========================================================
' ERROR CODES (10400 - 10499)
' Usage: managerError.ThrowError managerTable.TableNotFound, ...
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

' Get a table by name or raise managerTable.TableNotFound.
Public Function GetTable( _
    ByVal ws As Worksheet, _
    ByVal tableName As String) As ListObject

    Dim tbl As ListObject

    If ws Is Nothing Then
        managerError.ThrowError _
            managerTable.EmptyWorksheet, _
            MODULE_NAME & ".GetTable", _
            "Worksheet tidak boleh kosong."
    End If

    On Error Resume Next
    Set tbl = ws.ListObjects(tableName)
    On Error GoTo 0

    If tbl Is Nothing Then
        managerError.ThrowError _
            managerTable.TableNotFound, _
            MODULE_NAME & ".GetTable", _
            "Tabel '" & tableName & "' tidak ditemukan " & _
            "pada worksheet '" & ws.Name & "'."
    End If

    Set GetTable = tbl

End Function


' Raise managerTable.TableNotFound when tbl is Nothing.
Public Sub RequireTable( _
    ByVal tbl As ListObject, _
    ByVal errorSource As String)

    If tbl Is Nothing Then
        managerError.ThrowError _
            managerTable.TableNotFound, _
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


' Raise managerTable.ColumnNotFound for the first header in
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
'   managerTable.FormatColumn tbl, COLUMN_ID, "0", xlCenter
'   managerTable.FormatColumn tbl, COLUMN_NAME, "@"
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

    IsInColumn = managerForm.IsInRange( _
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

    IsInTable = managerForm.IsInRange(target, tbl.DataBodyRange)

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
        HasValue = managerValidationText.IsNotBlank(value)
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


' Number. Blank -> 0. Not numeric -> managerTable.InvalidNumber.
Public Function GetDouble( _
    ByVal tbl As ListObject, _
    ByVal rowIndex As Long, _
    ByVal columnName As String) As Double

    Dim value As Variant

    value = GetValue(tbl, rowIndex, columnName)

    RequireNoCellError value, tbl, rowIndex, columnName, MODULE_NAME & ".GetDouble"

    If managerValidationText.IsBlank(value) Then
        GetDouble = 0
        Exit Function
    End If

    If managerValidationNumeric.IsNotNumber(value) Then
        managerError.ThrowError _
            managerTable.InvalidNumber, _
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

    If managerValidationText.IsBlank(value) Then
        GetDate = 0
        Exit Function
    End If

    If Not IsDate(value) Then
        managerError.ThrowError _
            managerTable.InvalidDate, _
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

    Dim targetCell As Range

    Set targetCell = GetCell(tbl, rowIndex, columnName)

    If targetCell.Worksheet Is ActiveSheet Then
        targetCell.Select
    End If

End Sub


' =========================================================
' BULK (whole table)
'
' The designed look of a table lives in its FIRST body row.
' SetRowCount never deletes that row: an "empty" table keeps
' one blank row, so the formatting is never lost, and new
' rows copy the look of the first row.
' =========================================================

' Resize the body to rowCount rows and clear all values.
' Formatting of row 1 is copied to added rows. rowCount 0
' keeps one blank row (or leaves an empty table empty).
Public Sub SetRowCount( _
    ByVal tbl As ListObject, _
    ByVal rowCount As Long)

    Dim targetCount As Long
    Dim currentCount As Long
    Dim firstHeaderCell As Range

    RequireTable tbl, MODULE_NAME & ".SetRowCount"

    targetCount = rowCount
    If targetCount < 1 Then targetCount = 1

    If tbl.ListRows.Count = 0 Then

        ' Empty table: only Excel's insert row is shown.
        If rowCount <= 0 Then Exit Sub

        GrowEmptyTable tbl, targetCount
        Exit Sub

    End If

    tbl.DataBodyRange.ClearContents

    currentCount = tbl.ListRows.Count

    If currentCount > targetCount Then

        tbl.DataBodyRange.Rows(targetCount + 1) _
            .Resize(currentCount - targetCount).Delete Shift:=xlUp

    ElseIf currentCount < targetCount Then

        Set firstHeaderCell = tbl.HeaderRowRange.Cells(1, 1)

        tbl.Resize firstHeaderCell.Resize( _
            targetCount + 1, tbl.ListColumns.Count)

        CopyRowLook _
            tbl.DataBodyRange.Rows(1), _
            tbl.DataBodyRange.Rows(currentCount + 1) _
                .Resize(targetCount - currentCount)

    End If

End Sub


' Add one row at the bottom with the look of row 1 and keep
' the values of the other rows. Returns its row index.
Public Function AddStyledRow( _
    ByVal tbl As ListObject) As Long

    RequireTable tbl, MODULE_NAME & ".AddStyledRow"

    If tbl.ListRows.Count = 0 Then
        SetRowCount tbl, 1
        AddStyledRow = 1
        Exit Function
    End If

    SetRowCountKeep tbl, tbl.ListRows.Count + 1

    AddStyledRow = tbl.ListRows.Count

End Function


' Make room for rowCount body rows when cells directly below
' the table are in use (for example a merged block designed
' right under it): blank cells are inserted in the table's
' columns only, so everything below in those columns moves
' down. SetRowCount shrinking moves it back up.
Public Sub MakeRoomBelow( _
    ByVal tbl As ListObject, _
    ByVal rowCount As Long)

    Dim currentCount As Long
    Dim extra As Long

    RequireTable tbl, MODULE_NAME & ".MakeRoomBelow"

    currentCount = tbl.ListRows.Count
    If currentCount < 1 Then currentCount = 1

    extra = rowCount - currentCount
    If extra <= 0 Then Exit Sub

    tbl.HeaderRowRange.Offset(currentCount + 1).Resize(extra) _
        .Insert Shift:=xlDown

End Sub


' Delete one body row. The last remaining row is cleared
' instead, so the table keeps its look.
Public Sub DeleteRow( _
    ByVal tbl As ListObject, _
    ByVal rowIndex As Long)

    RequireRow tbl, rowIndex, MODULE_NAME & ".DeleteRow"

    If tbl.ListRows.Count = 1 Then
        tbl.DataBodyRange.ClearContents
    Else
        tbl.ListRows(rowIndex).Delete
    End If

End Sub


' Write a 2D array (1-based rows and columns) into the body.
' The body is resized to the number of rows first. Columns
' of data beyond the table width are ignored, so callers may
' keep hidden key columns at the end of their arrays.
'
' columnFormats (optional): Array with one number format per
' table column ("" = keep). Applied BEFORE the values are
' written, so text columns ("@") keep leading zeros.
Public Sub WriteData( _
    ByVal tbl As ListObject, _
    ByVal data As Variant, _
    Optional ByVal rowCount As Long = -1, _
    Optional ByVal columnFormats As Variant)

    Dim columnCount As Long
    Dim dataColumns As Long
    Dim output() As Variant
    Dim r As Long
    Dim c As Long

    RequireTable tbl, MODULE_NAME & ".WriteData"

    columnCount = tbl.ListColumns.Count

    If rowCount < 0 Then
        If IsArray(data) Then
            rowCount = UBound(data, 1)
        Else
            rowCount = 0
        End If
    End If

    SetRowCount tbl, rowCount

    If Not IsMissing(columnFormats) Then
        ApplyColumnFormats tbl, columnFormats
    End If

    If rowCount = 0 Then Exit Sub

    dataColumns = UBound(data, 2)
    If dataColumns > columnCount Then dataColumns = columnCount

    ReDim output(1 To rowCount, 1 To columnCount)

    For r = 1 To rowCount
        For c = 1 To dataColumns
            output(r, c) = data(r, c)
        Next c
    Next r

    tbl.DataBodyRange.value = output

End Sub


' Apply one number format per table column ("" = keep).
Public Sub ApplyColumnFormats( _
    ByVal tbl As ListObject, _
    ByVal columnFormats As Variant)

    Dim c As Long
    Dim position As Long

    RequireTable tbl, MODULE_NAME & ".ApplyColumnFormats"

    If Not IsArray(columnFormats) Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    position = 0

    For c = LBound(columnFormats) To UBound(columnFormats)

        position = position + 1
        If position > tbl.ListColumns.Count Then Exit For

        If Len(CStr(columnFormats(c))) > 0 Then
            tbl.ListColumns(position).DataBodyRange.NumberFormat = _
                CStr(columnFormats(c))
        End If

    Next c

End Sub


' Column letter of a table column on the worksheet ("T").
Public Function ColumnLetter( _
    ByVal tbl As ListObject, _
    ByVal columnName As String) As String

    Dim columnNumber As Long

    RequireColumn tbl, columnName, MODULE_NAME & ".ColumnLetter"

    columnNumber = tbl.ListColumns(columnName).Range.Column

    ColumnLetter = Split( _
        tbl.Parent.Cells(1, columnNumber).Address(True, False), "$")(0)

End Function


' =========================================================
' CONDITIONAL FORMAT RULES
'
' Rules use INDEX($X:$X,ROW()) instead of relative cell
' references, so they work for every row the table will
' ever have, no matter which cell is active when the rule is
' created. Each rule carries a tag (a harmless text
' comparison) so it can be found and replaced: calling the
' Ensure* procedures again never stacks duplicate rules.
' =========================================================

' Fill the cells of highlightColumns on every row whose
' markerColumn shows markerText. Used for "edit mode" rows.
Public Sub EnsureRowHighlight( _
    ByVal tbl As ListObject, _
    ByVal ruleTag As String, _
    ByVal markerColumn As String, _
    ByVal markerText As String, _
    ByVal highlightColumns As Variant, _
    ByVal fillColor As Long)

    Dim formulaText As String
    Dim columnName As Variant

    RequireTable tbl, MODULE_NAME & ".EnsureRowHighlight"

    If tbl.DataBodyRange Is Nothing Then Exit Sub

    formulaText = "=AND(""" & ruleTag & """<>""""," & _
        "INDEX($" & ColumnLetter(tbl, markerColumn) & ":$" & _
        ColumnLetter(tbl, markerColumn) & ",ROW())=""" & _
        markerText & """)"

    RemoveRules tbl, ruleTag

    For Each columnName In highlightColumns
        AddFillRule GetColumnRange(tbl, CStr(columnName)), _
            formulaText, fillColor, False, -1
    Next columnName

End Sub


' Add one formula rule to the whole body (or to one column
' when columnName is given). fontColor -1 = keep.
Public Sub EnsureFormulaRule( _
    ByVal tbl As ListObject, _
    ByVal ruleTag As String, _
    ByVal conditionFormula As String, _
    Optional ByVal columnName As String = vbNullString, _
    Optional ByVal fillColor As Long = -1, _
    Optional ByVal fontColor As Long = -1, _
    Optional ByVal makeBold As Boolean = False)

    Dim targetRange As Range
    Dim formulaText As String

    RequireTable tbl, MODULE_NAME & ".EnsureFormulaRule"

    If tbl.DataBodyRange Is Nothing Then Exit Sub

    RemoveRules tbl, ruleTag

    If Len(columnName) = 0 Then
        Set targetRange = tbl.DataBodyRange
    Else
        Set targetRange = GetColumnRange(tbl, columnName)
    End If

    formulaText = "=AND(""" & ruleTag & """<>""""," & conditionFormula & ")"

    AddFillRule targetRange, formulaText, fillColor, makeBold, fontColor

End Sub


' INDEX($X:$X,ROW()) for a table column, for use inside
' EnsureFormulaRule conditions.
Public Function RowCellFormula( _
    ByVal tbl As ListObject, _
    ByVal columnName As String) As String

    Dim letter As String

    letter = ColumnLetter(tbl, columnName)

    RowCellFormula = "INDEX($" & letter & ":$" & letter & ",ROW())"

End Function


' =========================================================
' INTERNAL
' =========================================================

' Turn the insert row of an empty table into rowCount real
' rows. ListRows.Add is NOT used: on an empty table Excel then
' paints the columns that have no column fill with a dark
' fill. Resize keeps the designed insert row. A resize to one
' row keeps it an insert row, so grow by one more and drop it.
Private Sub GrowEmptyTable( _
    ByVal tbl As ListObject, _
    ByVal rowCount As Long)

    Dim firstHeaderCell As Range
    Dim resizedRows As Long

    Set firstHeaderCell = tbl.HeaderRowRange.Cells(1, 1)

    resizedRows = rowCount
    If resizedRows < 2 Then resizedRows = 2

    tbl.Resize firstHeaderCell.Resize(resizedRows + 1, tbl.ListColumns.Count)

    Do While tbl.ListRows.Count > rowCount
        tbl.ListRows(tbl.ListRows.Count).Delete
    Loop

    tbl.DataBodyRange.ClearContents

    If rowCount > 1 Then
        CopyRowLook _
            tbl.DataBodyRange.Rows(1), _
            tbl.DataBodyRange.Rows(2).Resize(rowCount - 1)
    End If

End Sub


' Grow (never shrink) the body to rowCount rows while KEEPING
' the existing values. New rows get the look of row 1.
Private Sub SetRowCountKeep( _
    ByVal tbl As ListObject, _
    ByVal rowCount As Long)

    Dim currentCount As Long
    Dim firstHeaderCell As Range

    currentCount = tbl.ListRows.Count
    If rowCount <= currentCount Then Exit Sub

    Set firstHeaderCell = tbl.HeaderRowRange.Cells(1, 1)

    tbl.Resize firstHeaderCell.Resize(rowCount + 1, tbl.ListColumns.Count)

    tbl.DataBodyRange.Rows(currentCount + 1) _
        .Resize(rowCount - currentCount).ClearContents

    CopyRowLook _
        tbl.DataBodyRange.Rows(1), _
        tbl.DataBodyRange.Rows(currentCount + 1) _
            .Resize(rowCount - currentCount)

End Sub


' Copy fill, font, number format, alignment and borders of
' one source row to every row of target, column by column.
' (PasteSpecial xlPasteFormats is not used: it also copies
' conditional formats, which would stack up on every render.)
Private Sub CopyRowLook( _
    ByVal sourceRow As Range, _
    ByVal target As Range)

    Dim c As Long
    Dim sourceCell As Range
    Dim targetColumn As Range
    Dim edge As Variant

    For c = 1 To sourceRow.Columns.Count

        Set sourceCell = sourceRow.Cells(1, c)
        Set targetColumn = target.Columns(c)

        With targetColumn
            .NumberFormat = sourceCell.NumberFormat
            .HorizontalAlignment = sourceCell.HorizontalAlignment
            .VerticalAlignment = sourceCell.VerticalAlignment
            .WrapText = sourceCell.WrapText

            .Font.Name = sourceCell.Font.Name
            .Font.Size = sourceCell.Font.Size
            .Font.Bold = sourceCell.Font.Bold
            .Font.Italic = sourceCell.Font.Italic
            .Font.Color = sourceCell.Font.Color

            If sourceCell.Interior.ColorIndex = xlColorIndexNone Then
                .Interior.ColorIndex = xlColorIndexNone
            Else
                .Interior.Pattern = sourceCell.Interior.Pattern
                .Interior.Color = sourceCell.Interior.Color
            End If
        End With

        For Each edge In Array(xlEdgeLeft, xlEdgeRight, xlEdgeBottom)
            CopyBorder sourceCell.Borders(edge), targetColumn.Borders(edge)
        Next edge

        CopyBorder sourceCell.Borders(xlEdgeBottom), _
            targetColumn.Borders(xlEdgeTop)

        If targetColumn.Rows.Count > 1 Then
            CopyBorder sourceCell.Borders(xlEdgeBottom), _
                targetColumn.Borders(xlInsideHorizontal)
        End If

    Next c

End Sub


Private Sub CopyBorder( _
    ByVal source As Border, _
    ByVal target As Border)

    If source.LineStyle = xlNone Or IsNull(source.LineStyle) Then
        target.LineStyle = xlNone
    Else
        target.LineStyle = source.LineStyle
        target.Weight = source.Weight
        target.Color = source.Color
    End If

End Sub


' Delete every conditional format rule on the table's sheet
' that carries ruleTag. The whole sheet is scanned: a rule
' that covers only part of the body (one column) is not
' always listed by the body range itself.
Private Sub RemoveRules( _
    ByVal tbl As ListObject, _
    ByVal ruleTag As String)

    Dim rules As FormatConditions
    Dim i As Long
    Dim formulaText As String

    Set rules = tbl.Parent.Cells.FormatConditions

    For i = rules.Count To 1 Step -1

        formulaText = vbNullString

        On Error Resume Next
        formulaText = rules(i).Formula1
        On Error GoTo 0

        If InStr(1, formulaText, """" & ruleTag & """", vbTextCompare) > 0 Then
            rules(i).Delete
        End If

    Next i

End Sub


Private Sub AddFillRule( _
    ByVal targetRange As Range, _
    ByVal formulaText As String, _
    ByVal fillColor As Long, _
    ByVal makeBold As Boolean, _
    ByVal fontColor As Long)

    Dim rule As FormatCondition

    If targetRange Is Nothing Then Exit Sub

    Set rule = targetRange.FormatConditions.Add( _
        Type:=xlExpression, Formula1:=formulaText)

    If fillColor >= 0 Then
        rule.Interior.Color = fillColor
    End If

    If fontColor >= 0 Then
        rule.Font.Color = fontColor
    End If

    If makeBold Then
        rule.Font.Bold = True
    End If

End Sub


Private Sub RequireColumn( _
    ByVal tbl As ListObject, _
    ByVal columnName As String, _
    ByVal errorSource As String)

    RequireTable tbl, errorSource

    If Not HasColumn(tbl, columnName) Then
        managerError.ThrowError _
            managerTable.ColumnNotFound, _
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
        managerError.ThrowError _
            managerTable.InvalidRow, _
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
        managerError.ThrowError _
            managerTable.CellContainsError, _
            errorSource, _
            "Kolom '" & columnName & "' baris ke-" & _
            CStr(rowIndex) & " pada tabel '" & tbl.Name & _
            "' berisi nilai error."
    End If

End Sub
