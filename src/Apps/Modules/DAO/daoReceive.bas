Attribute VB_Name = "daoReceive"
Option Explicit

' =========================================================
' DAO - RECEIVE (penerimaan)
' Table: receives
'
' One function = one saved Access query.
' The connection (and its running transaction) is owned by
' the calling service. Validation, stock recalculation and
' activity logs are done by serviceReceive.
' =========================================================

Private Const MODULE_NAME As String = "daoReceive"


' receive_create
' The RECEIVE header (transactionUUID) must exist. 0 rows
' when the order is not active or belongs to another plant.
Public Function CreateReceive( _
    ByVal conn As ADODB.connection, _
    ByVal transactionUUID As String, _
    ByVal orderUUID As String, _
    ByVal receiveItem As Receive) As Long

    Dim cmd As ADODB.Command
    Dim verificationDate As Variant

    verificationDate = Null
    If receiveItem.IsVerified Then
        verificationDate = DateValue(receiveItem.VerificationDate)
    End If

    Set cmd = managerDatabase.CreateCommand(conn, "receive_create")

    managerDatabase.AddTextParameter cmd, "p_uuid", transactionUUID, 36
    managerDatabase.AddTextParameter cmd, "p_order_uuid", orderUUID, 36
    managerDatabase.AddDateParameter cmd, "p_arrival_date", DateValue(receiveItem.ArrivalDate)
    managerDatabase.AddDateParameter cmd, "p_verification_date", verificationDate
    managerDatabase.AddLongParameter cmd, "p_received_quantity", receiveItem.ReceivedQuantity
    managerDatabase.AddLongParameter cmd, "p_ok_quantity", receiveItem.OkQuantity
    managerDatabase.AddLongParameter cmd, "p_not_ok_quantity", receiveItem.NotOkQuantity
    managerDatabase.AddNullableLongTextParameter cmd, "p_remarks", receiveItem.Remarks
    managerDatabase.AddNullableLongTextParameter cmd, "p_system_remarks", receiveItem.SystemRemarks
    managerDatabase.AddNullableTextParameter cmd, "p_person_in_charge", receiveItem.PersonInCharge, 255

    CreateReceive = managerDatabase.ExecuteNonQuery(cmd)

End Function


' receive_list
' Active, not deleted receives of one plant with an arrival
' date between fromDate and toDate (whole days), keyed by
' UUID, in arrival date order. page: filter and page
' (Nothing = every row).
Public Function ListReceives( _
    ByVal conn As ADODB.connection, _
    ByVal plantItem As ManufacturingPlant, _
    ByVal fromDate As Date, _
    ByVal toDate As Date, _
    Optional ByVal page As PageRequest = Nothing) As Collection

    Dim cmd As ADODB.Command
    Dim rs As ADODB.recordset
    Dim receives As Collection
    Dim receiveItem As Receive

    If page Is Nothing Then Set page = New PageRequest

    Set cmd = managerDatabase.CreateCommand(conn, "receive_list")

    AddCriteria cmd, plantItem, fromDate, toDate
    managerDatabase.AddPageFilterParameters cmd, page
    managerDatabase.AddPageDateKeyParameters cmd, page

    Set rs = managerDatabase.ExecuteRecordset(cmd)
    managerDatabase.MoveToPageStart rs, page

    Set receives = New Collection

    Do While Not rs.EOF

        If managerDatabase.IsPageFull(receives.Count, page) Then Exit Do

        Set receiveItem = MapReceive(rs, plantItem)
        receives.Add receiveItem, receiveItem.UUID

        rs.MoveNext

    Loop

    managerDatabase.CloseRecordset rs

    Set ListReceives = receives

End Function


' receive_reassign_order
' Moves every receive (active and draft) of the old order to
' the new order. Returns the number of receives moved
' (0 is valid).
Public Function ReassignOrder( _
    ByVal conn As ADODB.connection, _
    ByVal oldOrderUUID As String, _
    ByVal newOrderUUID As String) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "receive_reassign_order")

    managerDatabase.AddTextParameter cmd, "p_old_order_uuid", oldOrderUUID, 36
    managerDatabase.AddTextParameter cmd, "p_new_order_uuid", newOrderUUID, 36

    ReassignOrder = managerDatabase.ExecuteNonQuery(cmd)

End Function


' receive_list_count
' Number of rows ListReceives returns over all pages.
Public Function CountReceives( _
    ByVal conn As ADODB.connection, _
    ByVal plantItem As ManufacturingPlant, _
    ByVal fromDate As Date, _
    ByVal toDate As Date, _
    Optional ByVal page As PageRequest = Nothing) As Long

    Dim cmd As ADODB.Command

    If page Is Nothing Then Set page = New PageRequest

    Set cmd = managerDatabase.CreateCommand(conn, "receive_list_count")

    AddCriteria cmd, plantItem, fromDate, toDate
    managerDatabase.AddPageFilterParameters cmd, page

    CountReceives = managerDatabase.ExecuteCount(cmd)

End Function


' =========================================================
' INTERNAL
' =========================================================

Private Sub AddCriteria( _
    ByVal cmd As ADODB.Command, _
    ByVal plantItem As ManufacturingPlant, _
    ByVal fromDate As Date, _
    ByVal toDate As Date)

    managerDatabase.AddLongParameter cmd, "p_manufacturing_plant_id", plantItem.Id
    managerDatabase.AddDateParameter cmd, "p_date_from", DateValue(fromDate)
    managerDatabase.AddDateParameter cmd, "p_date_to", DateValue(toDate)

End Sub


Private Function MapReceive( _
    ByVal rs As ADODB.recordset, _
    ByVal plantItem As ManufacturingPlant) As Receive

    Dim receiveItem As Receive
    Dim receiveHeader As Transaction
    Dim orderHeader As Transaction
    Dim orderItem As Order
    Dim sparepartItem As Sparepart
    Dim vendorItem As Vendor
    Dim textValue As String

    ' ---- order ------------------------------------------
    Set orderHeader = New Transaction
    orderHeader.UUID = managerDatabase.FieldText(rs, "order_uuid")
    orderHeader.TransactionType = "ORDER"
    Set orderHeader.ManufacturingPlant = plantItem

    Set sparepartItem = New Sparepart
    sparepartItem.Id = managerDatabase.FieldLong(rs, "sparepart_id")
    sparepartItem.Code = managerDatabase.FieldText(rs, "sparepart_code")
    sparepartItem.Name = managerDatabase.FieldText(rs, "sparepart_name")
    sparepartItem.UnitOfMeasurement = managerDatabase.FieldText(rs, "sparepart_unit_of_measurement")

    Set orderItem = New Order

    With orderItem
        .Id = managerDatabase.FieldLong(rs, "order_id")
        Set .Transaction = orderHeader
        Set .Sparepart = sparepartItem

        If Not managerDatabase.FieldIsNull(rs, "vendor_id") Then
            Set vendorItem = New Vendor
            vendorItem.Id = managerDatabase.FieldLong(rs, "vendor_id")
            vendorItem.Name = managerDatabase.FieldText(rs, "vendor_name")
            textValue = managerDatabase.FieldText(rs, "vendor_code")
            If Len(textValue) > 0 Then vendorItem.Code = textValue
            Set .Vendor = vendorItem
        End If

        .PoNumber = managerDatabase.FieldText(rs, "order_po_number")
        .OrderDate = managerDatabase.FieldDate(rs, "order_date")
        .Quantity = managerDatabase.FieldLong(rs, "order_quantity")
        .Remarks = managerDatabase.FieldText(rs, "order_remarks")
    End With

    ' ---- receive ----------------------------------------
    Set receiveHeader = New Transaction
    receiveHeader.UUID = managerDatabase.FieldText(rs, "uuid")
    receiveHeader.TransactionType = "RECEIVE"
    Set receiveHeader.ManufacturingPlant = plantItem
    receiveHeader.ActivatedAt = managerDatabase.FieldDate(rs, "activated_at")

    Set receiveItem = New Receive

    With receiveItem
        .Id = managerDatabase.FieldLong(rs, "id")
        Set .Transaction = receiveHeader
        Set .Order = orderItem

        .ArrivalDate = managerDatabase.FieldDate(rs, "arrival_date")
        .VerificationDate = managerDatabase.FieldDate(rs, "verification_date")

        ' Received first: OK + NOT OK may not exceed it.
        .ReceivedQuantity = managerDatabase.FieldLong(rs, "received_quantity")
        .OkQuantity = managerDatabase.FieldLong(rs, "ok_quantity")
        .NotOkQuantity = managerDatabase.FieldLong(rs, "not_ok_quantity")

        .Remarks = managerDatabase.FieldText(rs, "remarks")
        .SystemRemarks = managerDatabase.FieldText(rs, "system_remarks")

        textValue = managerDatabase.FieldText(rs, "person_in_charge")
        If Len(textValue) > 0 Then .PersonInCharge = textValue

        .RemainingQuantityBefore = managerDatabase.FieldLong(rs, "remaining_quantity_before")
    End With

    Set MapReceive = receiveItem

End Function
