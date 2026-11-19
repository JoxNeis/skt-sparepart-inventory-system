Attribute VB_Name = "daoRequest"
Option Explicit

' =========================================================
' DAO - REQUEST (WAREHOUSE REQUEST)
'
' One function = one saved Access query.
' The connection (and its running transaction) is owned by
' the calling service. Validation, stock check, row-count
' checks, stock changes and activity logs are done by
' serviceRequest.
' =========================================================

Private Const MODULE_NAME As String = "DAO_REQUEST"


' create_request
' Returns the affected row count.
Public Function CreateRequest( _
    ByVal conn As ADODB.connection, _
    ByVal transactionUUID As String, _
    ByVal requestItem As Request) As Long

    Dim cmd As ADODB.Command

    Set cmd = DatabaseManager.CreateCommand(conn, "create_request")

    DatabaseManager.AddTextParameter cmd, "p_transaction_uuid", transactionUUID
    DatabaseManager.AddTextParameter cmd, "p_mobile_post_name", requestItem.MobilePost.Name
    DatabaseManager.AddDateParameter cmd, "p_request_date", requestItem.orderDate
    DatabaseManager.AddTextParameter cmd, "p_catalogue_name", requestItem.Catalogue.Name
    DatabaseManager.AddDoubleParameter cmd, "p_request_amount", requestItem.Amount
    DatabaseManager.AddLongTextParameter cmd, "p_request_remarks", requestItem.remarks
    DatabaseManager.AddLongTextParameter cmd, "p_request_system_remarks", requestItem.SystemRemarks
    DatabaseManager.AddTextParameter cmd, "p_request_person_in_charge", requestItem.personInCharge

    CreateRequest = DatabaseManager.ExecuteNonQuery(cmd)

End Function


' select_request_by_manufacturing_plant_name_and_order_date
' Returns Request objects (no collection key).
Public Function GetRequestsByManufacturingPlantAndOrderDate( _
    ByVal conn As ADODB.connection, _
    ByVal manufacturingPlantName As String, _
    ByVal fromDate As Date, _
    ByVal untilDate As Date) As Collection

    Dim cmd As ADODB.Command
    Dim rs As ADODB.recordset

    Dim result As Collection

    Dim requestItem As Request
    Dim mobilePostItem As MobilePost
    Dim catalogueItem As Catalogue

    Set cmd = DatabaseManager.CreateCommand(conn, _
        "select_request_by_manufacturing_plant_name_and_order_date")

    ' Access resolves saved-query parameters by position.
    ' Append them in exactly the same order as the
    ' PARAMETERS declaration in the saved query.
    DatabaseManager.AddTextParameter cmd, "p_manufacturing_plant_name", manufacturingPlantName
    DatabaseManager.AddDateParameter cmd, "p_from_date", fromDate
    DatabaseManager.AddDateParameter cmd, "p_until_date", untilDate

    Set rs = DatabaseManager.ExecuteRecordset(cmd)

    Set result = New Collection

    Do While Not rs.EOF

        Set requestItem = New Request
        Set mobilePostItem = New MobilePost
        Set catalogueItem = New Catalogue

        mobilePostItem.Name = DatabaseManager.FieldText(rs, "mobile_post_name")

        catalogueItem.Name = DatabaseManager.FieldText(rs, "catalogue_name")
        catalogueItem.unitOfMeasurement = DatabaseManager.FieldText(rs, "catalogue_unit_of_measurement")

        With requestItem
            .UUID = DatabaseManager.FieldText(rs, "request_uuid")

            Set .MobilePost = mobilePostItem
            Set .Catalogue = catalogueItem

            If Not DatabaseManager.FieldIsNull(rs, "request_date") Then
                .orderDate = DatabaseManager.FieldDate(rs, "request_date")
            End If

            If Not DatabaseManager.FieldIsNull(rs, "request_amount") Then
                .Amount = DatabaseManager.FieldDouble(rs, "request_amount")
            End If

            .unitOfMeasurement = DatabaseManager.FieldText(rs, "catalogue_unit_of_measurement")
            .remarks = DatabaseManager.FieldText(rs, "request_remarks")
            .SystemRemarks = DatabaseManager.FieldText(rs, "request_system_remarks")
            .personInCharge = DatabaseManager.FieldText(rs, "request_person_in_charge")
        End With

        result.Add requestItem

        Set requestItem = Nothing
        Set mobilePostItem = Nothing
        Set catalogueItem = Nothing

        rs.MoveNext
    Loop

    DatabaseManager.CloseRecordset rs

    Set GetRequestsByManufacturingPlantAndOrderDate = result

End Function
