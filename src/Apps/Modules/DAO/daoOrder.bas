Attribute VB_Name = "daoOrder"
Option Explicit

' =========================================================
' DAO - ORDER
' Table: orders (progress from vw_order_progress)
'
' One function = one saved Access query.
' The connection (and its running transaction) is owned by
' the calling service. Validation, row-count checks and
' activity logs are done by serviceOrder.
' =========================================================

Private Const MODULE_NAME As String = "daoOrder"


' order_create
' The ORDER header (transactionUUID) must exist. 0 rows when
' the header is not an ORDER, or the sparepart / vendor is
' archived or unknown.
Public Function CreateOrder( _
    ByVal conn As ADODB.connection, _
    ByVal transactionUUID As String, _
    ByVal orderItem As Order) As Long

    Dim cmd As ADODB.Command
    Dim vendorId As Variant

    vendorId = Null
    If orderItem.HasVendor Then vendorId = orderItem.Vendor.Id

    Set cmd = managerDatabase.CreateCommand(conn, "order_create")

    managerDatabase.AddTextParameter cmd, "p_uuid", transactionUUID, 36
    managerDatabase.AddLongParameter cmd, "p_sparepart_id", orderItem.Sparepart.Id
    managerDatabase.AddNullableLongParameter cmd, "p_vendor_id", vendorId
    managerDatabase.AddNullableTextParameter cmd, "p_po_number", orderItem.PoNumber, 50
    managerDatabase.AddDateParameter cmd, "p_order_date", DateValue(orderItem.OrderDate)
    managerDatabase.AddLongParameter cmd, "p_quantity", orderItem.Quantity
    managerDatabase.AddNullableLongTextParameter cmd, "p_remarks", orderItem.Remarks
    managerDatabase.AddNullableLongTextParameter cmd, "p_system_remarks", orderItem.SystemRemarks
    managerDatabase.AddNullableTextParameter cmd, "p_person_in_charge", orderItem.PersonInCharge, 255

    CreateOrder = managerDatabase.ExecuteNonQuery(cmd)

End Function


' order_list
' Active, not deleted orders of one plant with an order date
' between fromDate and toDate (whole days), keyed by UUID, in
' order date order. page: filter and page (Nothing = every
' row).
Public Function ListOrders( _
    ByVal conn As ADODB.connection, _
    ByVal plantItem As ManufacturingPlant, _
    ByVal fromDate As Date, _
    ByVal toDate As Date, _
    Optional ByVal page As PageRequest = Nothing) As Collection

    Set ListOrders = RunList(conn, "order_list", plantItem, fromDate, toDate, page)

End Function


' order_list_count
' Number of rows ListOrders returns over all pages.
Public Function CountOrders( _
    ByVal conn As ADODB.connection, _
    ByVal plantItem As ManufacturingPlant, _
    ByVal fromDate As Date, _
    ByVal toDate As Date, _
    Optional ByVal page As PageRequest = Nothing) As Long

    CountOrders = RunCount(conn, "order_list_count", plantItem, fromDate, toDate, page)

End Function


' order_list_pending
' Same as ListOrders, only orders with a remaining quantity.
Public Function ListPendingOrders( _
    ByVal conn As ADODB.connection, _
    ByVal plantItem As ManufacturingPlant, _
    ByVal fromDate As Date, _
    ByVal toDate As Date, _
    Optional ByVal page As PageRequest = Nothing) As Collection

    Set ListPendingOrders = RunList(conn, "order_list_pending", plantItem, fromDate, toDate, page)

End Function


' order_list_pending_count
' Number of rows ListPendingOrders returns over all pages.
Public Function CountPendingOrders( _
    ByVal conn As ADODB.connection, _
    ByVal plantItem As ManufacturingPlant, _
    ByVal fromDate As Date, _
    ByVal toDate As Date, _
    Optional ByVal page As PageRequest = Nothing) As Long

    CountPendingOrders = RunCount(conn, "order_list_pending_count", plantItem, fromDate, toDate, page)

End Function


' order_set_po_number
Public Function SetPoNumber( _
    ByVal conn As ADODB.connection, _
    ByVal orderUUID As String, _
    ByVal poNumber As String) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "order_set_po_number")

    managerDatabase.AddTextParameter cmd, "p_uuid", orderUUID, 36
    managerDatabase.AddNullableTextParameter cmd, "p_po_number", poNumber, 50

    SetPoNumber = managerDatabase.ExecuteNonQuery(cmd)

End Function


' =========================================================
' INTERNAL
' =========================================================

Private Function RunList( _
    ByVal conn As ADODB.connection, _
    ByVal queryName As String, _
    ByVal plantItem As ManufacturingPlant, _
    ByVal fromDate As Date, _
    ByVal toDate As Date, _
    ByVal page As PageRequest) As Collection

    Dim cmd As ADODB.Command
    Dim rs As ADODB.recordset
    Dim orders As Collection
    Dim orderItem As Order

    If page Is Nothing Then Set page = New PageRequest

    Set cmd = managerDatabase.CreateCommand(conn, queryName)

    AddCriteria cmd, plantItem, fromDate, toDate
    managerDatabase.AddPageFilterParameters cmd, page
    managerDatabase.AddPageDateKeyParameters cmd, page

    Set rs = managerDatabase.ExecuteRecordset(cmd)
    managerDatabase.MoveToPageStart rs, page

    Set orders = New Collection

    Do While Not rs.EOF

        If managerDatabase.IsPageFull(orders.Count, page) Then Exit Do

        Set orderItem = MapOrder(rs, plantItem)
        orders.Add orderItem, orderItem.UUID

        rs.MoveNext

    Loop

    managerDatabase.CloseRecordset rs

    Set RunList = orders

End Function


Private Function RunCount( _
    ByVal conn As ADODB.connection, _
    ByVal queryName As String, _
    ByVal plantItem As ManufacturingPlant, _
    ByVal fromDate As Date, _
    ByVal toDate As Date, _
    ByVal page As PageRequest) As Long

    Dim cmd As ADODB.Command

    If page Is Nothing Then Set page = New PageRequest

    Set cmd = managerDatabase.CreateCommand(conn, queryName)

    AddCriteria cmd, plantItem, fromDate, toDate
    managerDatabase.AddPageFilterParameters cmd, page

    RunCount = managerDatabase.ExecuteCount(cmd)

End Function


Private Sub AddCriteria( _
    ByVal cmd As ADODB.Command, _
    ByVal plantItem As ManufacturingPlant, _
    ByVal fromDate As Date, _
    ByVal toDate As Date)

    managerDatabase.AddLongParameter cmd, "p_manufacturing_plant_id", plantItem.Id
    managerDatabase.AddDateParameter cmd, "p_date_from", DateValue(fromDate)
    managerDatabase.AddDateParameter cmd, "p_date_to", DateValue(toDate)

End Sub


' Current row of order_list / order_list_pending -> Order
' with its ORDER header, sparepart and vendor.
Private Function MapOrder( _
    ByVal rs As ADODB.recordset, _
    ByVal plantItem As ManufacturingPlant) As Order

    Dim orderItem As Order
    Dim header As Transaction
    Dim sparepartItem As Sparepart
    Dim vendorItem As Vendor
    Dim textValue As String

    Set header = New Transaction
    header.UUID = managerDatabase.FieldText(rs, "uuid")
    header.TransactionType = "ORDER"
    Set header.ManufacturingPlant = plantItem
    header.ActivatedAt = managerDatabase.FieldDate(rs, "activated_at")

    Set sparepartItem = New Sparepart
    sparepartItem.Id = managerDatabase.FieldLong(rs, "sparepart_id")
    sparepartItem.Code = managerDatabase.FieldText(rs, "sparepart_code")
    sparepartItem.Name = managerDatabase.FieldText(rs, "sparepart_name")
    sparepartItem.UnitOfMeasurement = managerDatabase.FieldText(rs, "sparepart_unit_of_measurement")

    Set orderItem = New Order

    With orderItem
        .Id = managerDatabase.FieldLong(rs, "id")
        Set .Transaction = header
        Set .Sparepart = sparepartItem

        If Not managerDatabase.FieldIsNull(rs, "vendor_id") Then
            Set vendorItem = New Vendor
            vendorItem.Id = managerDatabase.FieldLong(rs, "vendor_id")
            vendorItem.Name = managerDatabase.FieldText(rs, "vendor_name")
            textValue = managerDatabase.FieldText(rs, "vendor_code")
            If Len(textValue) > 0 Then vendorItem.Code = textValue
            Set .Vendor = vendorItem
        End If

        .PoNumber = managerDatabase.FieldText(rs, "po_number")
        .OrderDate = managerDatabase.FieldDate(rs, "order_date")
        .Quantity = managerDatabase.FieldLong(rs, "quantity")
        .Remarks = managerDatabase.FieldText(rs, "remarks")
        .SystemRemarks = managerDatabase.FieldText(rs, "system_remarks")

        textValue = managerDatabase.FieldText(rs, "person_in_charge")
        If Len(textValue) > 0 Then .PersonInCharge = textValue

        .OkQuantityTotal = managerDatabase.FieldLong(rs, "ok_quantity_total")
        .ReceivedQuantityTotal = managerDatabase.FieldLong(rs, "received_quantity_total")
        .RemainingQuantity = managerDatabase.FieldLong(rs, "remaining_quantity")
        .Status = managerDatabase.FieldText(rs, "status")
        .EstimatedArrival = managerDatabase.FieldDate(rs, "estimated_arrival")
    End With

    Set MapOrder = orderItem

End Function
