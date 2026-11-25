Attribute VB_Name = "daoMobileMaintenance"
Option Explicit

' =========================================================
' DAO - MOBILE MAINTENANCE (pos mobile)
' Table: mobile_maintenances
'
' One function = one saved Access query.
' Validation, transaction, row-count checks and activity
' log live in serviceMobileMaintenance.
' =========================================================

Private Const MODULE_NAME As String = "daoMobileMaintenance"


' mobile_maintenance_list
' Active mobile maintenances of one plant, ordered by name.
Public Function ListByPlant( _
    ByVal conn As ADODB.connection, _
    ByVal plantItem As ManufacturingPlant) As Collection

    Dim cmd As ADODB.Command
    Dim rs As ADODB.recordset
    Dim items As Collection
    Dim item As MobileMaintenance

    Set cmd = managerDatabase.CreateCommand(conn, "mobile_maintenance_list")

    managerDatabase.AddLongParameter cmd, "p_manufacturing_plant_id", plantItem.Id

    Set rs = managerDatabase.ExecuteRecordset(cmd)

    Set items = New Collection

    Do While Not rs.EOF

        Set item = New MobileMaintenance

        With item
            .Id = managerDatabase.FieldLong(rs, "id")
            Set .ManufacturingPlant = plantItem
            .Code = managerDatabase.FieldText(rs, "code")
            .Name = managerDatabase.FieldText(rs, "name")
        End With

        items.Add item
        rs.MoveNext

    Loop

    managerDatabase.CloseRecordset rs

    Set ListByPlant = items

End Function


' mobile_maintenance_create
Public Function CreateMobileMaintenance( _
    ByVal conn As ADODB.connection, _
    ByVal item As MobileMaintenance) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "mobile_maintenance_create")

    managerDatabase.AddLongParameter cmd, "p_manufacturing_plant_id", item.ManufacturingPlant.Id
    managerDatabase.AddTextParameter cmd, "p_code", item.Code, 40
    managerDatabase.AddTextParameter cmd, "p_name", item.Name, 255

    CreateMobileMaintenance = managerDatabase.ExecuteNonQuery(cmd)

End Function


' mobile_maintenance_update
Public Function UpdateMobileMaintenance( _
    ByVal conn As ADODB.connection, _
    ByVal item As MobileMaintenance) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "mobile_maintenance_update")

    managerDatabase.AddLongParameter cmd, "p_mobile_maintenance_id", item.Id
    managerDatabase.AddTextParameter cmd, "p_code", item.Code, 40
    managerDatabase.AddTextParameter cmd, "p_name", item.Name, 255

    UpdateMobileMaintenance = managerDatabase.ExecuteNonQuery(cmd)

End Function


' mobile_maintenance_archive
Public Function ArchiveMobileMaintenance( _
    ByVal conn As ADODB.connection, _
    ByVal mobileMaintenanceId As Long) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "mobile_maintenance_archive")

    managerDatabase.AddLongParameter cmd, "p_mobile_maintenance_id", mobileMaintenanceId

    ArchiveMobileMaintenance = managerDatabase.ExecuteNonQuery(cmd)

End Function
