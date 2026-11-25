Attribute VB_Name = "controllerDashboard"
Option Explicit

' =========================================================
' CONTROLLER - DASHBOARD (sheet DASHBOARD)
'
' Totals of the application location (SETTINGS > LOKASI) for
' the current month, refreshed every time the sheet opens.
'
'   ORDER        TOTAL ORDER I9, TOTAL ONGOING ORDER K9,
'                TOTAL FINISH ORDER M9, JUMLAH BARANG O9
'   PENERIMAAN   TOTAL PENERIMAAN I16, TOTAL BARANG DITERIMA K16,
'                TOTAL BARANG OK M16, TOTAL NOT OK O16
'   PENGELUARAN  TOTAL PENGELUARAN I24, TOTAL BARANG KELUAR K24,
'                TOTAL JENIS BARANG M24
'   SPAREPART URGENT  TableSparepartUrgent: stock at or below
'                the reorder point
' =========================================================

Private Const MODULE_NAME As String = "controllerDashboard"

' =========================================================
' ERROR CODES (15500 - 15599)
' =========================================================
' (errors come from serviceDashboard)

Private Const TITLE_CELL As String = "J2"
Private Const TITLE_TEXT As String = "DASHBOARD"

Private Const TABLE_URGENT As String = "TableSparepartUrgent"
Private Const TABLE_URGENT_DESIGN As String = "Table1"


' =========================================================
' EVENTS
' =========================================================

Public Sub OnActivate( _
    ByVal ws As Worksheet)

    Refresh ws

End Sub


' =========================================================
' ACTIONS
' =========================================================

Public Sub Refresh( _
    ByVal ws As Worksheet)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo
    Dim plantName As String
    Dim fromDate As Date
    Dim summary As Object
    Dim urgent As Collection
    Dim stockItem As ManufacturingPlantStock
    Dim data() As Variant
    Dim r As Long
    Dim tbl As ListObject

    FixTableName ws

    plantName = managerSheet.SettingsPlantName()

    If Len(plantName) = 0 Then
        ws.Range(TITLE_CELL).value = TITLE_TEXT
        Application.StatusBar = "Isi LOKASI di sheet SETTINGS untuk melihat dashboard."
        Exit Sub
    End If

    fromDate = DateSerial(Year(Date), Month(Date), 1)

    state = managerForm.BeginBusy("Memuat dashboard...")

    Set summary = serviceDashboard.GetSummary(plantName, fromDate, Date)

    ws.Range(TITLE_CELL).value = TITLE_TEXT & "  |  " & plantName & "  |  " & _
        managerSheet.PeriodText(Date)

    ws.Range("I9").value = summary("order_count")
    ws.Range("K9").value = summary("order_ongoing")
    ws.Range("M9").value = summary("order_finished")
    ws.Range("O9").value = summary("order_quantity")

    ws.Range("I16").value = summary("receive_count")
    ws.Range("K16").value = summary("receive_quantity")
    ws.Range("M16").value = summary("receive_ok")
    ws.Range("O16").value = summary("receive_not_ok")

    ws.Range("I24").value = summary("request_count")
    ws.Range("K24").value = summary("request_quantity")
    ws.Range("M24").value = summary("request_kinds")

    Set urgent = summary("urgent")
    Set tbl = managerTable.GetTable(ws, TABLE_URGENT)

    If urgent.Count = 0 Then
        managerTable.SetRowCount tbl, 0
    Else
        ReDim data(1 To urgent.Count, 1 To 4)

        For Each stockItem In urgent
            r = r + 1
            data(r, 1) = r
            data(r, 2) = stockItem.Sparepart.DisplayName
            data(r, 3) = stockItem.CurrentStock
            data(r, 4) = stockItem.ReorderPoint
        Next stockItem

        managerTable.WriteData tbl, data, urgent.Count, _
            Array("0", managerSheet.TEXT_FORMAT, managerSheet.NUMBER_FORMAT, _
                  managerSheet.NUMBER_FORMAT)
    End If

    managerForm.EndBusy state

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


' =========================================================
' INTERNAL
' =========================================================

' The design calls the urgent table "Table1".
Private Sub FixTableName( _
    ByVal ws As Worksheet)

    Dim tbl As ListObject

    On Error Resume Next
    Set tbl = ws.ListObjects(TABLE_URGENT)
    On Error GoTo 0

    If Not tbl Is Nothing Then Exit Sub

    On Error Resume Next
    Set tbl = ws.ListObjects(TABLE_URGENT_DESIGN)
    On Error GoTo 0

    If Not tbl Is Nothing Then tbl.Name = TABLE_URGENT

End Sub
