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

    Set cmd = DatabaseManager.CreateCommand(conn, "create_order")

    DatabaseManager.AddTextParameter cmd, "p_transaction_uuid", transactionUUID
    DatabaseManager.AddDateParameter cmd, "p_order_date", orderItem.orderDate
    DatabaseManager.AddTextParameter cmd, "p_catalogue_name", orderItem.Catalogue.Name
    DatabaseManager.AddDoubleParameter cmd, "p_amount", orderItem.Amount
    DatabaseManager.AddTextParameter cmd, "p_manufacturing_name", orderItem.Plant.Name
    DatabaseManager.AddTextParameter cmd, "p_vendor_name", orderItem.Vendor.Name
    DatabaseManager.AddTextParameter cmd, "p_remarks", orderItem.remarks
    DatabaseManager.AddTextParameter cmd, "person_in_charge", orderItem.personInCharge

    CreateOrder = DatabaseManager.ExecuteNonQuery(cmd)

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

    Set cmd = DatabaseManager.CreateCommand(conn, _
        "select_order_receive_by_manufacturing_plant_name_and_order_date")

    DatabaseManager.AddTextParameter cmd, "p_manufacturing_plant_name", manufacturingPlantName
    DatabaseManager.AddDateParameter cmd, "p_from_date", fromDate
    DatabaseManager.AddDateParameter cmd, "p_until_date", untilDate

    Set rs = DatabaseManager.ExecuteRecordset(cmd)

    Set orders = New Collection

    Do While Not rs.EOF

        Set orderItem = MapOrderRow(rs)

        orders.Add _
            item:=orderItem, _
            Key:=orderItem.UUID

        Set orderItem = Nothing

        rs.MoveNext
    Loop

    DatabaseManager.CloseRecordset rs

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

    Set cmd = DatabaseManager.CreateCommand(conn, _
        "select_pending_order_by_manufacturing_plant_name_and_order_date")

    DatabaseManager.AddTextParameter cmd, "p_manufacturing_plant_name", manufacturingPlantName
    DatabaseManager.AddDateParameter cmd, "p_from_date", fromDate
    DatabaseManager.AddDateParameter cmd, "p_until_date", untilDate

    Set rs = DatabaseManager.ExecuteRecordset(cmd)

    Set orders = New Collection

    Do While Not rs.EOF

        Set orderItem = MapOrderRow(rs)

        orders.Add _
            item:=orderItem, _
            Key:=orderItem.UUID

        Set orderItem = Nothing

        rs.MoveNext
    Loop

    DatabaseManager.CloseRecordset rs

    Set SelectPendingOrder = orders

End Function


' update_order_po_id
' Returns the affected row count.
Public Function UpdateOrderPOID( _
    ByVal conn As ADODB.connection, _
    ByVal orderUUID As String, _
    ByVal poID As String) As Long

    Dim cmd As ADODB.Command

    Set cmd = DatabaseManager.CreateCommand(conn, "update_order_po_id")

    DatabaseManager.AddTextParameter cmd, "p_order_uuid", orderUUID
    DatabaseManager.AddTextParameter cmd, "p_order_po_id", poID

    UpdateOrderPOID = DatabaseManager.ExecuteNonQuery(cmd)

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

    Set MapOrderRow = orderItem

End Function
