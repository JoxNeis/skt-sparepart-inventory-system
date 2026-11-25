Attribute VB_Name = "daoRequest"
Option Explicit

' =========================================================
' DAO - REQUEST (pengeluaran)
' Table: requests
'
' One function = one saved Access query.
' The connection (and its running transaction) is owned by
' the calling service. Stock checks, stock recalculation and
' activity logs are done by serviceRequest.
' =========================================================

Private Const MODULE_NAME As String = "daoRequest"


' request_create
' The REQUEST header (transactionUUID) must exist. 0 rows
' when the mobile maintenance belongs to another plant or the
' mobile maintenance / sparepart is archived or unknown.
Public Function CreateRequest( _
    ByVal conn As ADODB.connection, _
    ByVal transactionUUID As String, _
    ByVal requestItem As Request) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "request_create")

    managerDatabase.AddTextParameter cmd, "p_uuid", transactionUUID, 36
    managerDatabase.AddLongParameter cmd, "p_mobile_maintenance_id", requestItem.MobileMaintenance.Id
    managerDatabase.AddLongParameter cmd, "p_sparepart_id", requestItem.Sparepart.Id
    managerDatabase.AddDateParameter cmd, "p_request_date", DateValue(requestItem.RequestDate)
    managerDatabase.AddLongParameter cmd, "p_quantity", requestItem.Quantity
    managerDatabase.AddNullableLongTextParameter cmd, "p_remarks", requestItem.Remarks
    managerDatabase.AddNullableLongTextParameter cmd, "p_system_remarks", requestItem.SystemRemarks
    managerDatabase.AddNullableTextParameter cmd, "p_person_in_charge", requestItem.PersonInCharge, 255

    CreateRequest = managerDatabase.ExecuteNonQuery(cmd)

End Function


' request_list
' Active, not deleted requests of one plant with a request
' date between fromDate and toDate (whole days), keyed by
' UUID, in request date order. page: filter and page
' (Nothing = every row).
Public Function ListRequests( _
    ByVal conn As ADODB.connection, _
    ByVal plantItem As ManufacturingPlant, _
    ByVal fromDate As Date, _
    ByVal toDate As Date, _
    Optional ByVal page As PageRequest = Nothing) As Collection

    Dim cmd As ADODB.Command
    Dim rs As ADODB.recordset
    Dim requests As Collection
    Dim requestItem As Request

    If page Is Nothing Then Set page = New PageRequest

    Set cmd = managerDatabase.CreateCommand(conn, "request_list")

    AddCriteria cmd, plantItem, fromDate, toDate
    managerDatabase.AddPageFilterParameters cmd, page
    managerDatabase.AddPageDateKeyParameters cmd, page

    Set rs = managerDatabase.ExecuteRecordset(cmd)
    managerDatabase.MoveToPageStart rs, page

    Set requests = New Collection

    Do While Not rs.EOF

        If managerDatabase.IsPageFull(requests.Count, page) Then Exit Do

        Set requestItem = MapRequest(rs, plantItem)
        requests.Add requestItem, requestItem.UUID

        rs.MoveNext

    Loop

    managerDatabase.CloseRecordset rs

    Set ListRequests = requests

End Function


' request_list_count
' Number of rows ListRequests returns over all pages.
Public Function CountRequests( _
    ByVal conn As ADODB.connection, _
    ByVal plantItem As ManufacturingPlant, _
    ByVal fromDate As Date, _
    ByVal toDate As Date, _
    Optional ByVal page As PageRequest = Nothing) As Long

    Dim cmd As ADODB.Command

    If page Is Nothing Then Set page = New PageRequest

    Set cmd = managerDatabase.CreateCommand(conn, "request_list_count")

    AddCriteria cmd, plantItem, fromDate, toDate
    managerDatabase.AddPageFilterParameters cmd, page

    CountRequests = managerDatabase.ExecuteCount(cmd)

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


Private Function MapRequest( _
    ByVal rs As ADODB.recordset, _
    ByVal plantItem As ManufacturingPlant) As Request

    Dim requestItem As Request
    Dim header As Transaction
    Dim sparepartItem As Sparepart
    Dim mobileItem As MobileMaintenance
    Dim textValue As String

    Set header = New Transaction
    header.UUID = managerDatabase.FieldText(rs, "uuid")
    header.TransactionType = "REQUEST"
    Set header.ManufacturingPlant = plantItem
    header.ActivatedAt = managerDatabase.FieldDate(rs, "activated_at")

    Set sparepartItem = New Sparepart
    sparepartItem.Id = managerDatabase.FieldLong(rs, "sparepart_id")
    sparepartItem.Code = managerDatabase.FieldText(rs, "sparepart_code")
    sparepartItem.Name = managerDatabase.FieldText(rs, "sparepart_name")
    sparepartItem.UnitOfMeasurement = managerDatabase.FieldText(rs, "sparepart_unit_of_measurement")

    Set mobileItem = New MobileMaintenance
    mobileItem.Id = managerDatabase.FieldLong(rs, "mobile_maintenance_id")
    Set mobileItem.ManufacturingPlant = plantItem
    mobileItem.Name = managerDatabase.FieldText(rs, "mobile_maintenance_name")
    textValue = managerDatabase.FieldText(rs, "mobile_maintenance_code")
    If Len(textValue) > 0 Then mobileItem.Code = textValue

    Set requestItem = New Request

    With requestItem
        .Id = managerDatabase.FieldLong(rs, "id")
        Set .Transaction = header
        Set .MobileMaintenance = mobileItem
        Set .Sparepart = sparepartItem
        .RequestDate = managerDatabase.FieldDate(rs, "request_date")
        .Quantity = managerDatabase.FieldLong(rs, "quantity")
        .Remarks = managerDatabase.FieldText(rs, "remarks")
        .SystemRemarks = managerDatabase.FieldText(rs, "system_remarks")

        textValue = managerDatabase.FieldText(rs, "person_in_charge")
        If Len(textValue) > 0 Then .PersonInCharge = textValue
    End With

    Set MapRequest = requestItem

End Function
