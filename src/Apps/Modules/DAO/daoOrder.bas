Attribute VB_Name = "daoOrder"
Option Explicit

' =========================================================
' DAO - ORDER
'
' One function = one saved Access query.
' The connection (and its running transaction) is owned by
' the calling service. Validation, row-count checks and
' activity logs are done by serviceOrder.
' =========================================================

Private Const MODULE_NAME As String = "DAO_ORDER"


' create_order
' Returns the affected row count.
Public Function CreateOrder( _
    ByVal conn As ADODB.connection, _
    ByVal transactionUUID As String, _
    ByVal orderItem As Order) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "create_order")

    managerDatabase.AddTextParameter cmd, "p_transaction_uuid", transactionUUID
    managerDatabase.AddDateParameter cmd, "p_order_date", orderItem.orderDate
    managerDatabase.AddTextParameter cmd, "p_catalogue_name", orderItem.Catalogue.Name
    managerDatabase.AddDoubleParameter cmd, "p_amount", orderItem.Amount
    managerDatabase.AddTextParameter cmd, "p_manufacturing_name", orderItem.Plant.Name
    managerDatabase.AddTextParameter cmd, "p_vendor_name", orderItem.Vendor.Name
    managerDatabase.AddTextParameter cmd, "p_remarks", orderItem.remarks
    managerDatabase.AddTextParameter cmd, "person_in_charge", orderItem.personInCharge

    CreateOrder = managerDatabase.ExecuteNonQuery(cmd)

End Function


' select_order_receive_by_manufacturing_plant_name_and_order_date
' Returns Order objects keyed by order UUID.
Public Function SelectOrder( _
    ByVal conn As ADODB.connection, _
    ByVal manufacturingPlantName As String, _
    ByVal fromDate As Date, _
    ByVal untilDate As Date) As Collection

    Dim cmd As ADODB.Command
    Dim rs As ADODB.recordset
    Dim orders As Collection
    Dim orderItem As Order

    Set cmd = managerDatabase.CreateCommand(conn, _
        "select_order_receive_by_manufacturing_plant_name_and_order_date")

    managerDatabase.AddTextParameter cmd, "p_manufacturing_plant_name", manufacturingPlantName
    managerDatabase.AddDateParameter cmd, "p_from_date", fromDate
    managerDatabase.AddDateParameter cmd, "p_until_date", untilDate

    Set rs = managerDatabase.ExecuteRecordset(cmd)

    Set orders = New Collection

    Do While Not rs.EOF

        Set orderItem = MapOrderRow(rs)

        orders.Add _
            item:=orderItem, _
            Key:=orderItem.UUID

        Set orderItem = Nothing

        rs.MoveNext
    Loop

    managerDatabase.CloseRecordset rs

    Set SelectOrder = orders

End Function


' select_pending_order_by_manufacturing_plant_name_and_order_date
' Returns Order objects keyed by order UUID.
Public Function SelectPendingOrder( _
    ByVal conn As ADODB.connection, _
    ByVal manufacturingPlantName As String, _
    ByVal fromDate As Date, _
    ByVal untilDate As Date) As Collection

    Dim cmd As ADODB.Command
    Dim rs As ADODB.recordset
    Dim orders As Collection
    Dim orderItem As Order

    Set cmd = managerDatabase.CreateCommand(conn, _
        "select_pending_order_by_manufacturing_plant_name_and_order_date")

    managerDatabase.AddTextParameter cmd, "p_manufacturing_plant_name", manufacturingPlantName
    managerDatabase.AddDateParameter cmd, "p_from_date", fromDate
    managerDatabase.AddDateParameter cmd, "p_until_date", untilDate

    Set rs = managerDatabase.ExecuteRecordset(cmd)

    Set orders = New Collection

    Do While Not rs.EOF

        Set orderItem = MapOrderRow(rs)

        orders.Add _
            item:=orderItem, _
            Key:=orderItem.UUID

        Set orderItem = Nothing

        rs.MoveNext
    Loop

    managerDatabase.CloseRecordset rs

    Set SelectPendingOrder = orders

End Function


' update_order_po_id
' Returns the affected row count.
Public Function UpdateOrderPOID( _
    ByVal conn As ADODB.connection, _
    ByVal orderUUID As String, _
    ByVal poID As String) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "update_order_po_id")

    managerDatabase.AddTextParameter cmd, "p_order_uuid", orderUUID
    managerDatabase.AddTextParameter cmd, "p_order_po_id", poID

    UpdateOrderPOID = managerDatabase.ExecuteNonQuery(cmd)

End Function


' =========================================================
' INTERNAL
' =========================================================

' Current recordset row -> Order (with Plant, Catalogue and
' Vendor). Used by SelectOrder and SelectPendingOrder, which
' return the same columns.
Private Function MapOrderRow( _
    ByVal rs As ADODB.recordset) As Order

    Dim orderItem As Order
    Dim plantItem As Plant
    Dim catalogueItem As Catalogue
    Dim vendorItem As Vendor
    Dim poIdValue As String

    Set plantItem = New Plant
    Set catalogueItem = New Catalogue
    Set vendorItem = New Vendor
    Set orderItem = New Order

    plantItem.Name = managerDatabase.FieldText(rs, "manufacturing_plant_name")
    catalogueItem.Name = managerDatabase.FieldText(rs, "catalogue_name")
    catalogueItem.unitOfMeasurement = managerDatabase.FieldText(rs, "catalogue_unit_of_measurement")
    vendorItem.Name = managerDatabase.FieldText(rs, "vendor_name")

    orderItem.Id = managerDatabase.FieldLong(rs, "order_id")
    orderItem.UUID = managerDatabase.FieldText(rs, "order_uuid")

    poIdValue = managerDatabase.FieldText(rs, "order_poid")
    If Len(poIdValue) > 0 Then
        orderItem.poID = poIdValue
    End If

    Set orderItem.Plant = plantItem
    orderItem.orderDate = managerDatabase.FieldDate(rs, "order_date")
    Set orderItem.Catalogue = catalogueItem
    orderItem.Amount = managerDatabase.FieldDouble(rs, "order_amount")
    orderItem.Status = managerDatabase.FieldText(rs, "order_status")
    Set orderItem.Vendor = vendorItem
    orderItem.remarks = managerDatabase.FieldText(rs, "order_remarks")
    orderItem.personInCharge = managerDatabase.FieldText(rs, "order_person_in_charge")

    Set MapOrderRow = orderItem

End Function
