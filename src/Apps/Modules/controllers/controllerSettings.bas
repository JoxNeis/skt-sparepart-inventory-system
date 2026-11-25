Attribute VB_Name = "controllerSettings"
Option Explicit

' =========================================================
' CONTROLLER - SETTINGS (sheet SETTINGS)
'
'   APLIKASI > LOKASI  L8 (dropdown): the application location.
'     It is the default LOKASI of every sheet, the plant of the
'     DASHBOARD and REQUEST UPDATE, and the plant of new users.
'     Changing it moves every LOKASI cell to the new plant.
'
'   EMAIL > TO / CC / BCC tables: recipients of KIRIM EMAIL.
' =========================================================

Private Const MODULE_NAME As String = "controllerSettings"

' =========================================================
' ERROR CODES (15600 - 15699)
' =========================================================
' (none)


' =========================================================
' EVENTS
' =========================================================

Public Sub OnActivate( _
    ByVal ws As Worksheet)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    managerList.ApplyDropdown ws.Range(managerSheet.SETTINGS_PLANT_CELL), managerList.LIST_PLANT

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerError.ShowError errInfo

End Sub


Public Sub OnDoubleClick( _
    ByVal ws As Worksheet, _
    ByVal target As Range, _
    ByRef cancel As Boolean)

    ' No buttons on this sheet.

End Sub


Public Sub OnChange( _
    ByVal ws As Worksheet, _
    ByVal target As Range)

    If managerSheet.IsHit(target, ws, managerSheet.SETTINGS_PLANT_CELL) Then
        ApplyLocation managerSheet.SettingsPlantName()
    End If

End Sub


' =========================================================
' ACTIONS
' =========================================================

' Write plantName into the LOKASI cell of every sheet and
' reload the mobile posts of that plant.
Public Sub ApplyLocation( _
    ByVal plantName As String)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo
    Dim target As Variant

    If Len(plantName) = 0 Then Exit Sub

    state = managerForm.BeginBusy("Mengganti lokasi...")

    For Each target In LocationCells()
        managerSheet.GetSheet(CStr(target(0))).Range(CStr(target(1))).Cells(1, 1).value = plantName
    Next target

    managerForm.EndBusy state

    controllerMDPosMobile.EnsurePlant plantName

    Application.StatusBar = "Lokasi aplikasi: " & plantName & "."

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


' =========================================================
' INTERNAL
' =========================================================

' Array(sheet, cell) of every LOKASI cell.
Private Function LocationCells() As Variant

    LocationCells = Array( _
        Array(managerSheet.SHEET_INPUT_ORDER, "K5"), _
        Array(managerSheet.SHEET_REKAP_ORDER, "K5"), _
        Array(managerSheet.SHEET_REKAP_PENDING_ORDER, "K5"), _
        Array(managerSheet.SHEET_INPUT_PENERIMAAN, "K5"), _
        Array(managerSheet.SHEET_REKAP_PENERIMAAN, "K5"), _
        Array(managerSheet.SHEET_INPUT_REQUEST, "K5"), _
        Array(managerSheet.SHEET_REKAP_REQUEST, "K5"), _
        Array(managerSheet.SHEET_KARTU_STOCK, "K5"), _
        Array(managerSheet.SHEET_REKAP_INVENTORY, "K5"))

End Function
