Attribute VB_Name = "daoTransactionUpdateRequest"
Option Explicit

' =========================================================
' DAO - TRANSACTION UPDATE REQUEST
' Table: transaction_update_requests
'
' One function = one saved Access query.
' The approval workflow lives in
' serviceTransactionUpdateRequest.
' =========================================================

Private Const MODULE_NAME As String = "daoTransactionUpdateRequest"

' Document columns of transaction_update_request_list that
' exist once with old_ and once with new_ in front.
Private Const SNAPSHOT_COLUMNS As String = _
    "transaction_uuid,document_date,quantity,ok_quantity," & _
    "not_ok_quantity,remarks,person_in_charge,order_uuid," & _
    "po_number,sparepart_id,sparepart_name," & _
    "sparepart_unit_of_measurement,vendor_name," & _
    "mobile_maintenance_name"


' transaction_update_request_create
' 0 rows when the original is not active, the proposal is
' not a draft of the same type and plant, or the original
' already has a pending update request.
Public Function CreateUpdateRequest( _
    ByVal conn As ADODB.connection, _
    ByVal originalUUID As String, _
    ByVal proposedUUID As String, _
    ByVal reason As String, _
    ByVal requestedBy As String) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "transaction_update_request_create")

    managerDatabase.AddTextParameter cmd, "p_original_uuid", originalUUID, 36
    managerDatabase.AddTextParameter cmd, "p_proposed_uuid", proposedUUID, 36
    managerDatabase.AddNullableLongTextParameter cmd, "p_reason", reason
    managerDatabase.AddTextParameter cmd, "p_requested_by", requestedBy, 64

    CreateUpdateRequest = managerDatabase.ExecuteNonQuery(cmd)

End Function


' transaction_update_request_list
' Update requests (any verdict) of one plant requested
' between fromDate and toDate (whole days).
Public Function ListUpdateRequests( _
    ByVal conn As ADODB.connection, _
    ByVal plantId As Long, _
    ByVal fromDate As Date, _
    ByVal toDate As Date) As Collection

    Dim cmd As ADODB.Command
    Dim rs As ADODB.recordset
    Dim items As Collection
    Dim item As TransactionUpdateRequest

    Set cmd = managerDatabase.CreateCommand(conn, "transaction_update_request_list")

    managerDatabase.AddLongParameter cmd, "p_manufacturing_plant_id", plantId
    managerDatabase.AddDateParameter cmd, "p_date_from", DateValue(fromDate)
    managerDatabase.AddDateParameter cmd, "p_date_to", DateValue(toDate)

    Set rs = managerDatabase.ExecuteRecordset(cmd)

    Set items = New Collection

    Do While Not rs.EOF

        Set item = MapUpdateRequest(rs)
        items.Add item, CStr(item.Id)

        rs.MoveNext

    Loop

    managerDatabase.CloseRecordset rs

    Set ListUpdateRequests = items

End Function


' transaction_update_request_update
' Sets the verdict (APPROVED / REJECTED). 0 rows when the
' request is not pending, or the reviewer is archived or
' does not have a higher tier than the requester.
Public Function SetVerdict( _
    ByVal conn As ADODB.connection, _
    ByVal updateRequestId As Long, _
    ByVal verdict As String, _
    ByVal reviewer As String, _
    ByVal decisionNote As String) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "transaction_update_request_update")

    managerDatabase.AddLongParameter cmd, "p_update_request_id", updateRequestId
    managerDatabase.AddTextParameter cmd, "p_verdict", verdict, 10
    managerDatabase.AddTextParameter cmd, "p_reviewer", reviewer, 64
    managerDatabase.AddNullableLongTextParameter cmd, "p_decision_note", decisionNote

    SetVerdict = managerDatabase.ExecuteNonQuery(cmd)

End Function


' =========================================================
' INTERNAL
' =========================================================

Private Function MapUpdateRequest( _
    ByVal rs As ADODB.recordset) As TransactionUpdateRequest

    Dim item As TransactionUpdateRequest
    Dim columnName As Variant
    Dim key As String

    Set item = New TransactionUpdateRequest

    With item
        .Id = managerDatabase.FieldLong(rs, "id")
        .TransactionType = managerDatabase.FieldText(rs, "transaction_type")
        .Reason = managerDatabase.FieldText(rs, "reason")
        .RequestedBy = managerDatabase.FieldText(rs, "requested_by")
        .RequestedAt = managerDatabase.FieldDate(rs, "requested_at")
        .Verdict = managerDatabase.FieldText(rs, "verdict")
        .Reviewer = managerDatabase.FieldText(rs, "reviewer")
        .ReviewedAt = managerDatabase.FieldDate(rs, "reviewed_at")
        .DecisionNote = managerDatabase.FieldText(rs, "decision_note")
        .ManufacturingPlantName = managerDatabase.FieldText(rs, "manufacturing_plant_name")
    End With

    For Each columnName In Split(SNAPSHOT_COLUMNS, ",")

        ' old_transaction_uuid / new_transaction_uuid are
        ' stored as "uuid".
        key = CStr(columnName)
        If key = "transaction_uuid" Then key = "uuid"

        item.Original.Item(key) = managerDatabase.FieldValue(rs, "old_" & CStr(columnName))
        item.Proposed.Item(key) = managerDatabase.FieldValue(rs, "new_" & CStr(columnName))

    Next columnName

    Set MapUpdateRequest = item

End Function
