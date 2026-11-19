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

    Set cmd = DatabaseManager.CreateCommand(conn, _
        "select_stock_movement_by_manufacturing_plant_name_and_order_date")

    DatabaseManager.AddTextParameter cmd, "p_manufacturing_plant_name", manufacturingPlantName
    DatabaseManager.AddTextParameter cmd, "p_catalogue_name", catalogueName
    DatabaseManager.AddIntegerParameter cmd, "p_month", movementMonth
    DatabaseManager.AddLongParameter cmd, "p_year", movementYear

    Set rs = DatabaseManager.ExecuteRecordset(cmd)

    Set stockMovements = New Collection

    Do While Not rs.EOF

        Set stockMovementItem = New StockMovement

        With stockMovementItem
            .MovementDate = CLng(DatabaseManager.FieldDate(rs, "movement_date"))
            .IncomingReferenceNumber = DatabaseManager.FieldText(rs, "incoming_reference_number")
            .IncomingFrom = DatabaseManager.FieldText(rs, "incoming_from")
            .incomingAmount = DatabaseManager.FieldDouble(rs, "incoming_amount")
            .OutgoingReferenceNumber = DatabaseManager.FieldText(rs, "outgoing_reference_number")
            .OutgoingTo = DatabaseManager.FieldText(rs, "outgoing_to")
            .outgoingAmount = DatabaseManager.FieldDouble(rs, "outgoing_amount")
            .remarks = DatabaseManager.FieldText(rs, "remarks")
            .personInCharge = DatabaseManager.FieldText(rs, "person_in_charge")
        End With

        stockMovements.Add item:=stockMovementItem

        Set stockMovementItem = Nothing

        rs.MoveNext
    Loop

    DatabaseManager.CloseRecordset rs

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

    Set cmd = DatabaseManager.CreateCommand(conn, "select_beginning_stock")

    DatabaseManager.AddDateParameter cmd, "p_date", stockDate
    DatabaseManager.AddTextParameter cmd, "p_manufacturing_plant_name", manufacturingPlantName
    DatabaseManager.AddTextParameter cmd, "p_catalogue_name", catalogueName

    Set rs = DatabaseManager.ExecuteRecordset(cmd)

    If rs.EOF Then
        SelectBeginningStock = 0
    Else
        SelectBeginningStock = DatabaseManager.FieldDouble(rs, "total_stock")
    End If

    DatabaseManager.CloseRecordset rs

End Function
