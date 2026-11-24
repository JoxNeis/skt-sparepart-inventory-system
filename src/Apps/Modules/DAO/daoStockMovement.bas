Attribute VB_Name = "daoStockMovement"
Option Explicit

' =========================================================
' DAO - STOCK MOVEMENT
'
' One function = one saved Access query.
' The running beginning / end stock is calculated by
' serviceStockMovement.
' =========================================================

Private Const MODULE_NAME As String = "DAO_STOCK_MOVEMENT"


' select_stock_movement_by_manufacturing_plant_name_and_order_date
' Returns StockMovement objects with the movement fields
' filled in. BeginningStock / EndStock are left at 0.
Public Function SelectStockMovement( _
    ByVal conn As ADODB.connection, _
    ByVal manufacturingPlantName As String, _
    ByVal catalogueName As String, _
    ByVal movementMonth As Integer, _
    ByVal movementYear As Long) As Collection

    Dim cmd As ADODB.Command
    Dim rs As ADODB.recordset
    Dim stockMovements As Collection
    Dim stockMovementItem As StockMovement

    Set cmd = managerDatabase.CreateCommand(conn, _
        "select_stock_movement_by_manufacturing_plant_name_and_order_date")

    managerDatabase.AddTextParameter cmd, "p_manufacturing_plant_name", manufacturingPlantName
    managerDatabase.AddTextParameter cmd, "p_catalogue_name", catalogueName
    managerDatabase.AddIntegerParameter cmd, "p_month", movementMonth
    managerDatabase.AddLongParameter cmd, "p_year", movementYear

    Set rs = managerDatabase.ExecuteRecordset(cmd)

    Set stockMovements = New Collection

    Do While Not rs.EOF

        Set stockMovementItem = New StockMovement

        With stockMovementItem
            .MovementDate = CLng(managerDatabase.FieldDate(rs, "movement_date"))
            .IncomingReferenceNumber = managerDatabase.FieldText(rs, "incoming_reference_number")
            .IncomingFrom = managerDatabase.FieldText(rs, "incoming_from")
            .incomingAmount = managerDatabase.FieldDouble(rs, "incoming_amount")
            .OutgoingReferenceNumber = managerDatabase.FieldText(rs, "outgoing_reference_number")
            .OutgoingTo = managerDatabase.FieldText(rs, "outgoing_to")
            .outgoingAmount = managerDatabase.FieldDouble(rs, "outgoing_amount")
            .remarks = managerDatabase.FieldText(rs, "remarks")
            .personInCharge = managerDatabase.FieldText(rs, "person_in_charge")
        End With

        stockMovements.Add item:=stockMovementItem

        Set stockMovementItem = Nothing

        rs.MoveNext
    Loop

    managerDatabase.CloseRecordset rs

    Set SelectStockMovement = stockMovements

End Function


' select_beginning_stock
' Returns total_stock, or 0 when there is no row / Null.
Public Function SelectBeginningStock( _
    ByVal conn As ADODB.connection, _
    ByVal stockDate As Date, _
    ByVal manufacturingPlantName As String, _
    ByVal catalogueName As String) As Double

    Dim cmd As ADODB.Command
    Dim rs As ADODB.recordset

    Set cmd = managerDatabase.CreateCommand(conn, "select_beginning_stock")

    managerDatabase.AddDateParameter cmd, "p_date", stockDate
    managerDatabase.AddTextParameter cmd, "p_manufacturing_plant_name", manufacturingPlantName
    managerDatabase.AddTextParameter cmd, "p_catalogue_name", catalogueName

    Set rs = managerDatabase.ExecuteRecordset(cmd)

    If rs.EOF Then
        SelectBeginningStock = 0
    Else
        SelectBeginningStock = managerDatabase.FieldDouble(rs, "total_stock")
    End If

    managerDatabase.CloseRecordset rs

End Function
