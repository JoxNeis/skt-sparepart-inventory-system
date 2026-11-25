Attribute VB_Name = "serviceManufacturingPlant"
Option Explicit

' =========================================================
' SERVICE - MANUFACTURING PLANT
'
' Plants are maintained directly in Access. The application
' reads them for the LOKASI dropdowns and resolves the plant
' a sheet works with by its name.
' =========================================================

Private Const MODULE_NAME As String = "serviceManufacturingPlant"

' =========================================================
' ERROR CODES (18300 - 18399)
' Usage: managerError.ThrowError serviceManufacturingPlant.PlantNotFound, ...
' =========================================================
Public Const EmptyPlantName As Long = 18301
Public Const PlantNotFound As Long = 18302


' Active plants, ordered by name.
Public Function ListPlants() As Collection

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    Set ListPlants = daoManufacturingPlant.ListPlants( _
        managerDatabase.GetConnection())

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ListPlants", _
        "Gagal mengambil data plant."

End Function


' The active plant with this name (or code, or "code - name").
' Raises PlantNotFound.
Public Function GetPlant( _
    ByVal plantName As String) As ManufacturingPlant

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    Set GetPlant = FindPlant(managerDatabase.GetConnection(), plantName)

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".GetPlant", _
        "Gagal mengambil data plant."

End Function


' Same as GetPlant, on the caller's connection (inside a
' running transaction).
Public Function FindPlant( _
    ByVal conn As ADODB.connection, _
    ByVal plantName As String) As ManufacturingPlant

    Dim plantItem As ManufacturingPlant

    plantName = Trim$(plantName)

    If managerValidationText.IsBlank(plantName) Then
        managerError.ThrowError _
            serviceManufacturingPlant.EmptyPlantName, _
            MODULE_NAME & ".FindPlant", _
            "Lokasi / plant belum dipilih."
    End If

    Set plantItem = FindInList(daoManufacturingPlant.ListPlants(conn), plantName)

    If plantItem Is Nothing Then
        managerError.ThrowError _
            serviceManufacturingPlant.PlantNotFound, _
            MODULE_NAME & ".FindPlant", _
            "Plant '" & plantName & "' tidak ditemukan atau sudah diarsipkan."
    End If

    Set FindPlant = plantItem

End Function


' The plant in plants whose name, code or display name equals
' text (ignoring case). Nothing when there is none.
Public Function FindInList( _
    ByVal plants As Collection, _
    ByVal text As String) As ManufacturingPlant

    Dim plantItem As ManufacturingPlant

    text = Trim$(text)
    If Len(text) = 0 Then Exit Function

    For Each plantItem In plants
        If managerValidationText.IsSameText(plantItem.Name, text) _
            Or managerValidationText.IsSameText(plantItem.Code, text) _
            Or managerValidationText.IsSameText(plantItem.DisplayName, text) Then

            Set FindInList = plantItem
            Exit Function
        End If
    Next plantItem

End Function
