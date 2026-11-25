Attribute VB_Name = "coreApp"
Option Explicit

' =========================================================
' APPLICATION START
'
' Called by ThisWorkbook.Workbook_Open:
'   1. LISTS sheet + workbook names for the dropdowns
'   2. master data from the database into the MD tables and
'      the plant list (so every dropdown is filled)
'   3. SETTINGS > LOKASI defaults to the plant of the current
'      user (or the first plant)
'   4. open the DASHBOARD
'
' A database problem is shown once; the workbook stays usable
' and REFRESH on the MD sheets loads the data later.
'
' LAYERING
'   Sheet (event) -> Controller (reads / writes cells)
'                 -> Service (rules, one DB transaction)
'                 -> DAO (one saved Access query)
' =========================================================

Private Const MODULE_NAME As String = "coreApp"

' =========================================================
' ERROR CODES (10800 - 10899)
' =========================================================
Public Const StartFailed As Long = 10801


Public Sub Initialize()

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo

    state = managerForm.BeginBusy("Menyiapkan " & coreConfig.ApplicationName() & "...", True)

    managerList.EnsureInfrastructure

    SyncMasterData

    managerForm.EndBusy state

    managerNavigation.GoToSheet managerSheet.SHEET_DASHBOARD

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo, _
        "Data master belum dapat dimuat dari database. " & _
        "Periksa file database lalu klik REFRESH di sheet MASTER DATA."

End Sub


' Plants, spareparts, vendors, users and the mobile posts of
' the application location, from the database into the
' workbook. Raises on a database error.
Public Sub SyncMasterData()

    Dim plants As Collection
    Dim sessionUser As User
    Dim plantName As String
    Dim settingsSheet As Worksheet

    Set plants = serviceManufacturingPlant.ListPlants()
    managerList.WritePlants plants

    controllerMDSparepart.RefreshTable managerSheet.GetSheet(managerSheet.SHEET_MD_SPAREPART)
    controllerMDVendor.RefreshTable managerSheet.GetSheet(managerSheet.SHEET_MD_VENDOR)
    controllerMDUser.RefreshTable managerSheet.GetSheet(managerSheet.SHEET_MD_USER)

    Set settingsSheet = managerSheet.GetSheet(managerSheet.SHEET_SETTINGS)

    plantName = managerSheet.SettingsPlantName()

    ' The stored location must still exist.
    If Len(plantName) > 0 Then
        If serviceManufacturingPlant.FindInList(plants, plantName) Is Nothing Then
            plantName = vbNullString
        End If
    End If

    If Len(plantName) = 0 Then

        Set sessionUser = serviceUser.GetCurrentUser()

        If Not sessionUser Is Nothing Then
            plantName = sessionUser.ManufacturingPlant.Name
        ElseIf plants.Count > 0 Then
            plantName = plants(1).Name
        End If

        settingsSheet.Range(managerSheet.SETTINGS_PLANT_CELL).value = plantName

        If Len(plantName) > 0 Then
            controllerSettings.ApplyLocation plantName
        End If

    End If

    managerList.ApplyDropdown _
        settingsSheet.Range(managerSheet.SETTINGS_PLANT_CELL), managerList.LIST_PLANT

    controllerMDPosMobile.EnsurePlant plantName

    If plants.Count = 0 Then
        managerForm.ShowWarning _
            "Belum ada plant di database. Tambahkan plant di Microsoft Access " & _
            "(tabel manufacturing_plants), lalu buka ulang workbook ini."
    End If

End Sub
