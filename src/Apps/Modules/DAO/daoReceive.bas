Attribute VB_Name = "daoReceive"
Option Explicit

' =========================================================
' DAO - RECEIVE
'
' One function = one saved Access query.
' The connection (and its running transaction) is owned by
' the calling service. Validation, row-count checks, stock
' changes and activity logs are done by serviceReceive.
' =========================================================

Private Const MODULE_NAME As String = "DAO_RECEIVE"


' create_receive
' Returns the affected row count.
Public Function CreateReceive( _
    ByVal conn As ADODB.connection, _
    ByVal transactionUUID As String, _
    ByVal orderTransactionUUID As String, _
    ByVal receiveItem As Receive) As Long

    Dim cmd As ADODB.Command

    Set cmd = DatabaseManager.CreateCommand(conn, "create_receive")

    DatabaseManager.AddTextParameter cmd, "p_receive_transaction_uuid", transactionUUID
    DatabaseManager.AddTextParameter cmd, "p_order_transaction_uuid", orderTransactionUUID
    DatabaseManager.AddDateParameter cmd, "p_arrival_date", receiveItem.ArrivalDate
    DatabaseManager.AddDateParameter cmd, "p_verification_date", receiveItem.VerificationDate
    DatabaseManager.AddDoubleParameter cmd, "p_receive_total_amount", receiveItem.totalReceiveAmount
    DatabaseManager.AddDoubleParameter cmd, "p_receive_ok_amount", receiveItem.okAmount
    DatabaseManager.AddDoubleParameter cmd, "p_receive_not_ok_amount", receiveItem.notOKAmount
    DatabaseManager.AddTextParameter cmd, "p_receive_remarks", receiveItem.remarks
    DatabaseManager.AddTextParameter cmd, "p_receive_system_remarks", receiveItem.SystemRemarks
    DatabaseManager.AddTextParameter cmd, "p_receive_person_in_charge", receiveItem.personInCharge

    CreateReceive = DatabaseManager.ExecuteNonQuery(cmd)

End Function


' select_receive_by_manufacturing_plant_name_and_arrival_date
' Returns Receive objects keyed by receive transaction UUID.
Public Function SelectReceives( _
    ByVal conn As ADODB.connection, _
    ByVal manufacturingPlantName As String, _
    ByVal fromDate As Date, _
    ByVal untilDate As Date) As Collection

    Dim cmd As ADODB.Command
    Dim rs As ADODB.recordset

    Dim receives As Collection
    Dim receiveItem As Receive
    Dim orderItem As Order
    Dim plantItem As Plant
    Dim catalogueItem As Catalogue
    Dim vendorItem As Vendor

    Dim poIdValue As String

    Set cmd = DatabaseManager.CreateCommand(conn, _
        "select_receive_by_manufacturing_plant_name_and_arrival_date")

    DatabaseManager.AddTextParameter cmd, "p_manufacturing_plant_name", manufacturingPlantName
    DatabaseManager.AddDateParameter cmd, "p_from_date", fromDate
    DatabaseManager.AddDateParameter cmd, "p_until_date", untilDate

    Set rs = DatabaseManager.ExecuteRecordset(cmd)

    Set receives = New Collection

    Do While Not rs.EOF

        Set plantItem = New Plant
        Set catalogueItem = New Catalogue
        Set vendorItem = New Vendor
        Set orderItem = New Order
        Set receiveItem = New Receive

        ' -------------------------------------------------
        ' Plant object
        ' -------------------------------------------------
        plantItem.Name = DatabaseManager.FieldText(rs, "manufacturing_plant_name")

        ' -------------------------------------------------
        ' Catalogue object
        ' -------------------------------------------------
        catalogueItem.Name = DatabaseManager.FieldText(rs, "catalogue_name")
        catalogueItem.unitOfMeasurement = DatabaseManager.FieldText(rs, "catalogue_unit_of_measurement")

        ' -------------------------------------------------
        ' Vendor object
        ' -------------------------------------------------
        vendorItem.Name = DatabaseManager.FieldText(rs, "vendor_name")

        ' -------------------------------------------------
        ' Order object
        ' -------------------------------------------------
        orderItem.Id = DatabaseManager.FieldLong(rs, "order_id")
        orderItem.UUID = DatabaseManager.FieldText(rs, "order_uuid")

        poIdValue = DatabaseManager.FieldText(rs, "order_poid")
        If Len(poIdValue) > 0 Then
            orderItem.poID = poIdValue
        End If

        Set orderItem.Plant = plantItem
        Set orderItem.Catalogue = catalogueItem
        Set orderItem.Vendor = vendorItem

        orderItem.orderDate = DatabaseManager.FieldDate(rs, "order_date")
        orderItem.Amount = DatabaseManager.FieldDouble(rs, "order_amount")
        orderItem.remarks = DatabaseManager.FieldText(rs, "order_remarks")

        ' -------------------------------------------------
        ' Receive object
        ' -------------------------------------------------
        receiveItem.Id = DatabaseManager.FieldLong(rs, "receive_id")
        receiveItem.UUID = DatabaseManager.FieldText(rs, "transaction_uuid")

        Set receiveItem.Order = orderItem

        receiveItem.unitOfMeasurement = DatabaseManager.FieldText(rs, "catalogue_unit_of_measurement")
        receiveItem.RemainingOrder = DatabaseManager.FieldDouble(rs, "total_remaining_order")
        receiveItem.totalReceiveAmount = DatabaseManager.FieldDouble(rs, "total_receive_amount")
        receiveItem.okAmount = DatabaseManager.FieldDouble(rs, "total_ok_amount")
        receiveItem.notOKAmount = DatabaseManager.FieldDouble(rs, "total_not_ok_amount")

        If Not DatabaseManager.FieldIsNull(rs, "receive_arrival_date") Then
            receiveItem.ArrivalDate = DatabaseManager.FieldDate(rs, "receive_arrival_date")
        End If

        If Not DatabaseManager.FieldIsNull(rs, "receive_verification_date") Then
            receiveItem.VerificationDate = DatabaseManager.FieldDate(rs, "receive_verification_date")
        End If

        receiveItem.remarks = DatabaseManager.FieldText(rs, "receive_remarks")
        receiveItem.SystemRemarks = DatabaseManager.FieldText(rs, "receive_system_remarks")
        receiveItem.personInCharge = DatabaseManager.FieldText(rs, "receive_person_in_charge")

        ' Transaction UUID should be unique for each receive.
        receives.Add _
            item:=receiveItem, _
            Key:=receiveItem.UUID

        Set receiveItem = Nothing
        Set orderItem = Nothing
        Set plantItem = Nothing
        Set catalogueItem = Nothing
        Set vendorItem = Nothing

        rs.MoveNext
    Loop

    DatabaseManager.CloseRecordset rs

    Set SelectReceives = receives

End Function


' select_pending_order_by_manufacturing_plant_name_and_order_date
' Returns Receive objects (Order + remaining order amount)
' keyed by order UUID.
Public Function SelectPendingOrderReceive( _
    ByVal conn As ADODB.connection, _
    ByVal manufacturingPlantName As String, _
    ByVal fromDate As Date, _
    ByVal untilDate As Date) As Collection

    Dim cmd As ADODB.Command
    Dim rs As ADODB.recordset

    Dim receives As Collection
    Dim orderItem As Order
    Dim receiveItem As Receive
    Dim plantItem As Plant
    Dim catalogueItem As Catalogue
    Dim vendorItem As Vendor

    Dim poIdValue As String

    Set cmd = DatabaseManager.CreateCommand(conn, _
        "select_pending_order_by_manufacturing_plant_name_and_order_date")

    DatabaseManager.AddTextParameter cmd, "p_manufacturing_plant_name", manufacturingPlantName
    DatabaseManager.AddDateParameter cmd, "p_from_date", fromDate
    DatabaseManager.AddDateParameter cmd, "p_until_date", untilDate

    Set rs = DatabaseManager.ExecuteRecordset(cmd)

    Set receives = New Collection

    Do While Not rs.EOF

        Set plantItem = New Plant
        Set catalogueItem = New Catalogue
        Set vendorItem = New Vendor
        Set orderItem = New Order
        Set receiveItem = New Receive

        plantItem.Name = DatabaseManager.FieldText(rs, "manufacturing_plant_name")

        catalogueItem.Name = DatabaseManager.FieldText(rs, "catalogue_name")
        catalogueItem.unitOfMeasurement = DatabaseManager.FieldText(rs, "catalogue_unit_of_measurement")

        vendorItem.Name = DatabaseManager.FieldText(rs, "vendor_name")

        orderItem.Id = DatabaseManager.FieldLong(rs, "order_id")
        orderItem.UUID = DatabaseManager.FieldText(rs, "order_uuid")

        poIdValue = DatabaseManager.FieldText(rs, "order_poid")
        If Len(poIdValue) > 0 Then
            orderItem.poID = poIdValue
        End If

        Set orderItem.Plant = plantItem
        orderItem.orderDate = DatabaseManager.FieldDate(rs, "order_date")
        Set orderItem.Catalogue = catalogueItem
        orderItem.Amount = DatabaseManager.FieldDouble(rs, "order_amount")
        orderItem.Status = DatabaseManager.FieldText(rs, "order_status")
        Set orderItem.Vendor = vendorItem
        orderItem.remarks = DatabaseManager.FieldText(rs, "order_remarks")
        orderItem.personInCharge = DatabaseManager.FieldText(rs, "order_person_in_charge")

        Set receiveItem.Order = orderItem
        receiveItem.RemainingOrder = DatabaseManager.FieldDouble(rs, "remaining_order_amount")

        receives.Add _
            item:=receiveItem, _
            Key:=orderItem.UUID

        Set receiveItem = Nothing
        Set orderItem = Nothing
        Set plantItem = Nothing
        Set catalogueItem = Nothing
        Set vendorItem = Nothing

        rs.MoveNext
    Loop

    DatabaseManager.CloseRecordset rs

    Set SelectPendingOrderReceive = receives

End Function
