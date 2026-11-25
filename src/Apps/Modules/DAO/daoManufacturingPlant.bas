Attribute VB_Name = "daoManufacturingPlant"
Option Explicit

' =========================================================
' DAO - MANUFACTURING PLANT
' Table: manufacturing_plants
'
' One function = one saved Access query.
' Plants are maintained in Access; the application only
' reads them (dropdown lists, document numbers).
' =========================================================

Private Const MODULE_NAME As String = "daoManufacturingPlant"


' manufacturing_plant_list
' Active plants, ordered by name.
Public Function ListPlants( _
    ByVal conn As ADODB.connection) As Collection

    Dim cmd As ADODB.Command
    Dim rs As ADODB.recordset
    Dim plants As Collection

    Set cmd = managerDatabase.CreateCommand(conn, "manufacturing_plant_list")
    Set rs = managerDatabase.ExecuteRecordset(cmd)

    Set plants = New Collection

    Do While Not rs.EOF
        plants.Add MapPlant(rs)
        rs.MoveNext
    Loop

    managerDatabase.CloseRecordset rs

    Set ListPlants = plants

End Function


' manufacturing_plant_create
Public Function CreatePlant( _
    ByVal conn As ADODB.connection, _
    ByVal plantItem As ManufacturingPlant) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "manufacturing_plant_create")

    managerDatabase.AddTextParameter cmd, "p_code", plantItem.Code, 15
    managerDatabase.AddTextParameter cmd, "p_name", plantItem.Name, 255
    managerDatabase.AddNullableTextParameter cmd, "p_address", plantItem.Address, 255
    managerDatabase.AddNullableTextParameter cmd, "p_type", plantItem.PlantType, 50
    managerDatabase.AddNullableTextParameter cmd, "p_region", plantItem.Region, 100

    CreatePlant = managerDatabase.ExecuteNonQuery(cmd)

End Function


' manufacturing_plant_update
' 0 rows when the plant is archived, or when the code would
' change while the plant already has a transaction.
Public Function UpdatePlant( _
    ByVal conn As ADODB.connection, _
    ByVal plantItem As ManufacturingPlant) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "manufacturing_plant_update")

    managerDatabase.AddLongParameter cmd, "p_manufacturing_plant_id", plantItem.Id
    managerDatabase.AddTextParameter cmd, "p_code", plantItem.Code, 15
    managerDatabase.AddTextParameter cmd, "p_name", plantItem.Name, 255
    managerDatabase.AddNullableTextParameter cmd, "p_address", plantItem.Address, 255
    managerDatabase.AddNullableTextParameter cmd, "p_type", plantItem.PlantType, 50
    managerDatabase.AddNullableTextParameter cmd, "p_region", plantItem.Region, 100

    UpdatePlant = managerDatabase.ExecuteNonQuery(cmd)

End Function


' manufacturing_plant_archive
Public Function ArchivePlant( _
    ByVal conn As ADODB.connection, _
    ByVal plantId As Long) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "manufacturing_plant_archive")

    managerDatabase.AddLongParameter cmd, "p_manufacturing_plant_id", plantId

    ArchivePlant = managerDatabase.ExecuteNonQuery(cmd)

End Function


' =========================================================
' INTERNAL
' =========================================================

Private Function MapPlant( _
    ByVal rs As ADODB.recordset) As ManufacturingPlant

    Dim plantItem As ManufacturingPlant

    Set plantItem = New ManufacturingPlant

    With plantItem
        .Id = managerDatabase.FieldLong(rs, "id")
        .Code = managerDatabase.FieldText(rs, "code")
        .Name = managerDatabase.FieldText(rs, "name")
        .Address = managerDatabase.FieldText(rs, "address")
        .PlantType = managerDatabase.FieldText(rs, "type")
        .Region = managerDatabase.FieldText(rs, "region")
    End With

    Set MapPlant = plantItem

End Function
