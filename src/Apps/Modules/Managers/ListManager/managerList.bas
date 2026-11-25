Attribute VB_Name = "managerList"
Option Explicit

' =========================================================
' LIST MANAGER (dropdowns)
'
' Every vendor, sparepart, mobile post and plant on the
' sheets is picked from a dropdown (data validation list).
' Sparepart, vendor and mobile post dropdowns show (and put
' in the cell) the DISPLAY value "code - name"; the services
' resolve that text back to the row.
'
'   ListSparepart   = TableMDSparepart[DISPLAY]         (MD_SPAREPART)
'   ListVendor      = TableListVendor[DISPLAY]          (LISTS, hidden)
'   ListMobilePost  = TableListMobilePost[DISPLAY]      (LISTS, hidden)
'   ListPlant       = TableListPlant[NAMA]              (LISTS, hidden)
'
' The MD_VENDOR and MD_POS_MOBILE tables have no DISPLAY
' column, so their "code - name" lists live on the very
' hidden sheet LISTS and are rewritten every time those MD
' tables are refreshed (WriteVendors / WriteMobilePosts).
' Plants are maintained in Access and have no MD sheet; their
' list is refreshed by WritePlants.
' =========================================================

Private Const MODULE_NAME As String = "managerList"

' =========================================================
' ERROR CODES (10600 - 10699)
' =========================================================
Public Const ListSheetFailed As Long = 10601

Public Const LIST_SHEET_NAME As String = "LISTS"
Public Const TABLE_PLANT As String = "TableListPlant"
Public Const TABLE_VENDOR As String = "TableListVendor"
Public Const TABLE_MOBILE_POST As String = "TableListMobilePost"

Public Const LIST_PLANT As String = "ListPlant"
Public Const LIST_SPAREPART As String = "ListSparepart"
Public Const LIST_VENDOR As String = "ListVendor"
Public Const LIST_MOBILE_POST As String = "ListMobilePost"

Private Const PLANT_HEADERS As String = "ID,KODE,NAMA,TIPE,REGION"
Private Const DISPLAY_HEADERS As String = "ID,KODE,NAMA,DISPLAY"


' Create the LISTS sheet, its plant table and the workbook
' names when they are missing. Safe to call again.
Public Sub EnsureInfrastructure()

    Dim ws As Worksheet

    Set ws = EnsureListSheet()
    EnsureListTable ws, TABLE_PLANT, PLANT_HEADERS, "A1"
    EnsureListTable ws, TABLE_VENDOR, DISPLAY_HEADERS, "G1"
    EnsureListTable ws, TABLE_MOBILE_POST, DISPLAY_HEADERS, "L1"

    EnsureName LIST_PLANT, "=" & TABLE_PLANT & "[NAMA]"
    EnsureName LIST_SPAREPART, "=TableMDSparepart[DISPLAY]"
    EnsureName LIST_VENDOR, "=" & TABLE_VENDOR & "[DISPLAY]"
    EnsureName LIST_MOBILE_POST, "=" & TABLE_MOBILE_POST & "[DISPLAY]"

End Sub


' Replace the plant list with plants (ManufacturingPlant).
Public Sub WritePlants( _
    ByVal plants As Collection)

    Dim tbl As ListObject
    Dim data() As Variant
    Dim plantItem As ManufacturingPlant
    Dim r As Long

    Set tbl = managerTable.GetTable(EnsureListSheet(), TABLE_PLANT)

    If plants.Count = 0 Then
        managerTable.SetRowCount tbl, 0
        Exit Sub
    End If

    ReDim data(1 To plants.Count, 1 To 5)

    For Each plantItem In plants
        r = r + 1
        data(r, 1) = plantItem.Id
        data(r, 2) = plantItem.Code
        data(r, 3) = plantItem.Name
        data(r, 4) = plantItem.PlantType
        data(r, 5) = plantItem.Region
    Next plantItem

    managerTable.WriteData tbl, data, plants.Count, Array("0", "@", "@", "@", "@")

End Sub


' Replace the vendor dropdown list ("code - name").
Public Sub WriteVendors( _
    ByVal vendors As Collection)

    Dim data() As Variant
    Dim vendorItem As Vendor
    Dim r As Long

    If vendors.Count > 0 Then
        ReDim data(1 To vendors.Count, 1 To 4)
        For Each vendorItem In vendors
            r = r + 1
            data(r, 1) = vendorItem.Id
            data(r, 2) = vendorItem.Code
            data(r, 3) = vendorItem.Name
            data(r, 4) = vendorItem.DisplayName
        Next vendorItem
    End If

    WriteDisplayList TABLE_VENDOR, data, vendors.Count

End Sub


' Replace the mobile post dropdown list ("code - name") with
' the mobile maintenances of one plant.
Public Sub WriteMobilePosts( _
    ByVal items As Collection)

    Dim data() As Variant
    Dim item As MobileMaintenance
    Dim r As Long

    If Not items Is Nothing Then
        If items.Count > 0 Then
            ReDim data(1 To items.Count, 1 To 4)
            For Each item In items
                r = r + 1
                data(r, 1) = item.Id
                data(r, 2) = item.Code
                data(r, 3) = item.Name
                data(r, 4) = item.DisplayName
            Next item
        End If
    End If

    WriteDisplayList TABLE_MOBILE_POST, data, r

End Sub


' Plant type (TIPE) of a plant name from the plant list, ""
' when unknown. Used for the LOKASI line on KARTU_STOCK.
Public Function PlantType( _
    ByVal plantName As String) As String

    Dim tbl As ListObject
    Dim rowIndex As Long

    Set tbl = managerTable.GetTable(EnsureListSheet(), TABLE_PLANT)

    For rowIndex = 1 To managerTable.RowCount(tbl)
        If managerValidationText.IsSameText( _
            managerTable.GetText(tbl, rowIndex, "NAMA"), plantName) Then

            PlantType = managerTable.GetText(tbl, rowIndex, "TIPE")
            Exit Function
        End If
    Next rowIndex

End Function


' SATUAN of a sparepart picked in a dropdown, read from the
' MD_SPAREPART table (no database call). text may be the
' display name, the code or the name. "" when unknown.
Public Function SparepartUnit( _
    ByVal text As String) As String

    Dim tbl As ListObject
    Dim rowIndex As Long

    text = Trim$(text)
    If Len(text) = 0 Then Exit Function

    Set tbl = managerTable.GetTable( _
        managerSheet.GetSheet(managerSheet.SHEET_MD_SPAREPART), "TableMDSparepart")

    For rowIndex = 1 To managerTable.RowCount(tbl)
        If managerValidationText.IsSameText(managerTable.GetText(tbl, rowIndex, "DISPLAY"), text) _
            Or managerValidationText.IsSameText(managerTable.GetText(tbl, rowIndex, "KODE"), text) _
            Or managerValidationText.IsSameText(managerTable.GetText(tbl, rowIndex, "NAMA SPAREPART"), text) Then

            SparepartUnit = managerTable.GetText(tbl, rowIndex, "SATUAN")
            Exit Function
        End If
    Next rowIndex

End Function


' Dropdown from a workbook list name (ListVendor ...).
' strict = False lets the user type text that is not in the
' list (used for search boxes).
Public Sub ApplyDropdown( _
    ByVal target As Range, _
    ByVal listName As String, _
    Optional ByVal strict As Boolean = True)

    If target Is Nothing Then Exit Sub
    Set target = WholeArea(target)

    With target.Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:="=" & listName
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = strict
        .ErrorTitle = coreConfig.ApplicationName()
        .ErrorMessage = "Pilih nilai dari daftar."
    End With

End Sub


' Dropdown from fixed items, e.g. month names. Items may not
' contain commas (Formula1 always uses the English comma,
' also on Windows with an Indonesian list separator).
Public Sub ApplyItemsDropdown( _
    ByVal target As Range, _
    ByVal items As Variant, _
    Optional ByVal strict As Boolean = True)

    If target Is Nothing Then Exit Sub
    Set target = WholeArea(target)

    With target.Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:=Join(items, ",")
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = strict
        .ErrorTitle = coreConfig.ApplicationName()
        .ErrorMessage = "Pilih nilai dari daftar."
    End With

End Sub


' Whole numbers >= minimum only.
Public Sub ApplyWholeNumberRule( _
    ByVal target As Range, _
    Optional ByVal minimum As Long = 0)

    If target Is Nothing Then Exit Sub
    Set target = WholeArea(target)

    With target.Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlGreaterEqual, _
             Formula1:=CStr(minimum)
        .IgnoreBlank = True
        .ShowError = True
        .ErrorTitle = coreConfig.ApplicationName()
        .ErrorMessage = "Isi dengan bilangan bulat minimal " & minimum & "."
    End With

End Sub


Public Sub ClearValidation( _
    ByVal target As Range)

    If target Is Nothing Then Exit Sub
    target.Validation.Delete

End Sub


' Last 10 years up to next year, newest last.
Public Function YearItems() As Variant

    Dim items() As String
    Dim i As Long

    ReDim items(0 To 11)

    For i = 0 To 11
        items(i) = CStr(Year(Date) - 10 + i)
    Next i

    YearItems = items

End Function


' =========================================================
' INTERNAL
' =========================================================

' A single cell inside a merged area -> the whole area.
Private Function WholeArea( _
    ByVal target As Range) As Range

    If target.Cells.Count = 1 Then
        Set WholeArea = target.MergeArea
    Else
        Set WholeArea = target
    End If

End Function

Private Function EnsureListSheet() As Worksheet

    Dim ws As Worksheet
    Dim activeWs As Object
    Dim eventsWereOn As Boolean

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LIST_SHEET_NAME)
    On Error GoTo 0

    If Not ws Is Nothing Then
        Set EnsureListSheet = ws
        Exit Function
    End If

    eventsWereOn = Application.EnableEvents
    Application.EnableEvents = False

    Set activeWs = ActiveSheet

    Set ws = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = LIST_SHEET_NAME
    ws.Visible = xlSheetVeryHidden

    If Not activeWs Is Nothing Then activeWs.Activate

    Application.EnableEvents = eventsWereOn

    Set EnsureListSheet = ws

End Function


Private Sub EnsureListTable( _
    ByVal ws As Worksheet, _
    ByVal tableName As String, _
    ByVal headerText As String, _
    ByVal firstCellAddress As String)

    Dim tbl As ListObject
    Dim headers As Variant

    On Error Resume Next
    Set tbl = ws.ListObjects(tableName)
    On Error GoTo 0

    If Not tbl Is Nothing Then Exit Sub

    headers = Split(headerText, ",")

    ws.Range(firstCellAddress).Resize(1, UBound(headers) + 1).value = headers

    Set tbl = ws.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=ws.Range(firstCellAddress).Resize(2, UBound(headers) + 1), _
        XlListObjectHasHeaders:=xlYes)

    tbl.Name = tableName

End Sub


' Write rowCount rows (ID, KODE, NAMA, DISPLAY) into one of
' the display lists. Codes stay text ("0012").
Private Sub WriteDisplayList( _
    ByVal tableName As String, _
    ByVal data As Variant, _
    ByVal rowCount As Long)

    Dim tbl As ListObject

    Set tbl = managerTable.GetTable(EnsureListSheet(), tableName)

    If rowCount = 0 Then
        managerTable.SetRowCount tbl, 0
    Else
        managerTable.WriteData tbl, data, rowCount, Array("0", "@", "@", "@")
    End If

End Sub


Private Sub EnsureName( _
    ByVal nameText As String, _
    ByVal refersTo As String)

    Dim existing As Name

    On Error Resume Next
    Set existing = ThisWorkbook.Names(nameText)
    On Error GoTo 0

    If existing Is Nothing Then
        ThisWorkbook.Names.Add Name:=nameText, refersTo:=refersTo, Visible:=False
    ElseIf existing.refersTo <> refersTo Then
        existing.refersTo = refersTo
    End If

End Sub
