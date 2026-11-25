Attribute VB_Name = "daoStockLedger"
Option Explicit

' =========================================================
' DAO - STOCK LEDGER (kartu stok)
' View: vw_stock_ledger (active receives IN, requests OUT)
'
' One function = one saved Access query.
' =========================================================

Private Const MODULE_NAME As String = "daoStockLedger"


' stock_ledger_get_balance
' IN - OUT of every movement BEFORE beforeDate.
Public Function GetBalance( _
    ByVal conn As ADODB.connection, _
    ByVal plantId As Long, _
    ByVal sparepartId As Long, _
    ByVal beforeDate As Date) As Long

    Dim cmd As ADODB.Command
    Dim rs As ADODB.recordset

    Set cmd = managerDatabase.CreateCommand(conn, "stock_ledger_get_balance")

    managerDatabase.AddLongParameter cmd, "p_manufacturing_plant_id", plantId
    managerDatabase.AddLongParameter cmd, "p_sparepart_id", sparepartId
    managerDatabase.AddDateParameter cmd, "p_date", beforeDate

    Set rs = managerDatabase.ExecuteRecordset(cmd)

    If rs.EOF Then
        GetBalance = 0
    Else
        GetBalance = managerDatabase.FieldLong(rs, "balance")
    End If

    managerDatabase.CloseRecordset rs

End Function


' stock_ledger_get_available
' The lowest end-of-day balance from onDate onwards: the
' quantity that can still leave the store on onDate without
' the stock card going negative on that day or any later
' day. Negative when the card is already negative.
Public Function GetAvailableQuantity( _
    ByVal conn As ADODB.connection, _
    ByVal plantId As Long, _
    ByVal sparepartId As Long, _
    ByVal onDate As Date) As Long

    Dim cmd As ADODB.Command
    Dim rs As ADODB.recordset

    Set cmd = managerDatabase.CreateCommand(conn, "stock_ledger_get_available")

    managerDatabase.AddLongParameter cmd, "p_manufacturing_plant_id", plantId
    managerDatabase.AddLongParameter cmd, "p_sparepart_id", sparepartId
    managerDatabase.AddDateParameter cmd, "p_date", onDate

    Set rs = managerDatabase.ExecuteRecordset(cmd)

    If rs.EOF Then
        GetAvailableQuantity = 0
    Else
        GetAvailableQuantity = managerDatabase.FieldLong(rs, "available_quantity")
    End If

    managerDatabase.CloseRecordset rs

End Function


' stock_ledger_list_by_month
' Movements of one month, by date (IN before OUT on the same
' day, then by reference). BeginningStock / EndStock are
' filled by the service. page: keyed on (KeyDate = movement
' date, KeyText = IN / OUT, KeyText2 = reference)
' (Nothing = every row).
Public Function ListByMonth( _
    ByVal conn As ADODB.connection, _
    ByVal plantId As Long, _
    ByVal sparepartId As Long, _
    ByVal yearNumber As Integer, _
    ByVal monthNumber As Integer, _
    Optional ByVal page As PageRequest = Nothing) As Collection

    Dim cmd As ADODB.Command
    Dim rs As ADODB.recordset
    Dim entries As Collection
    Dim entry As StockLedgerEntry

    If page Is Nothing Then Set page = New PageRequest

    Set cmd = managerDatabase.CreateCommand(conn, "stock_ledger_list_by_month")

    managerDatabase.AddLongParameter cmd, "p_manufacturing_plant_id", plantId
    managerDatabase.AddLongParameter cmd, "p_sparepart_id", sparepartId
    managerDatabase.AddIntegerParameter cmd, "p_year", yearNumber
    managerDatabase.AddIntegerParameter cmd, "p_month", monthNumber
    managerDatabase.AddTextParameter cmd, "p_direction", page.Direction, 4
    managerDatabase.AddDateParameter cmd, "p_key_date", page.KeyDate
    managerDatabase.AddTextParameter cmd, "p_key_direction", page.KeyText, 3
    managerDatabase.AddTextParameter cmd, "p_key_reference", page.KeyText2, 36

    Set rs = managerDatabase.ExecuteRecordset(cmd)
    managerDatabase.MoveToPageStart rs, page

    Set entries = New Collection

    Do While Not rs.EOF

        If managerDatabase.IsPageFull(entries.Count, page) Then Exit Do

        Set entry = New StockLedgerEntry

        With entry
            .MovementDate = managerDatabase.FieldDate(rs, "movement_date")
            .IncomingReference = managerDatabase.FieldText(rs, "incoming_reference")
            .IncomingFrom = managerDatabase.FieldText(rs, "incoming_from")
            .IncomingQuantity = managerDatabase.FieldLong(rs, "incoming_quantity")
            .OutgoingReference = managerDatabase.FieldText(rs, "outgoing_reference")
            .OutgoingTo = managerDatabase.FieldText(rs, "outgoing_to")
            .OutgoingQuantity = managerDatabase.FieldLong(rs, "outgoing_quantity")
            .Remarks = managerDatabase.FieldText(rs, "remarks")
            .PersonInCharge = managerDatabase.FieldText(rs, "person_in_charge")
        End With

        entries.Add entry
        rs.MoveNext

    Loop

    managerDatabase.CloseRecordset rs

    Set ListByMonth = entries

End Function


' stock_ledger_count_by_month
' Number of movements of one month.
Public Function CountByMonth( _
    ByVal conn As ADODB.connection, _
    ByVal plantId As Long, _
    ByVal sparepartId As Long, _
    ByVal yearNumber As Integer, _
    ByVal monthNumber As Integer) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "stock_ledger_count_by_month")

    managerDatabase.AddLongParameter cmd, "p_manufacturing_plant_id", plantId
    managerDatabase.AddLongParameter cmd, "p_sparepart_id", sparepartId
    managerDatabase.AddIntegerParameter cmd, "p_year", yearNumber
    managerDatabase.AddIntegerParameter cmd, "p_month", monthNumber

    CountByMonth = managerDatabase.ExecuteCount(cmd)

End Function


' stock_ledger_get_balance_before
' Stock just before one stock card row (its STOK AWAL), in
' the order of ListByMonth.
Public Function GetBalanceBefore( _
    ByVal conn As ADODB.connection, _
    ByVal plantId As Long, _
    ByVal sparepartId As Long, _
    ByVal keyDate As Date, _
    ByVal keyDirection As String, _
    ByVal keyReference As String) As Long

    Dim cmd As ADODB.Command
    Dim rs As ADODB.recordset

    Set cmd = managerDatabase.CreateCommand(conn, "stock_ledger_get_balance_before")

    managerDatabase.AddLongParameter cmd, "p_manufacturing_plant_id", plantId
    managerDatabase.AddLongParameter cmd, "p_sparepart_id", sparepartId
    managerDatabase.AddDateParameter cmd, "p_key_date", keyDate
    managerDatabase.AddTextParameter cmd, "p_key_direction", keyDirection, 3
    managerDatabase.AddTextParameter cmd, "p_key_reference", keyReference, 36

    Set rs = managerDatabase.ExecuteRecordset(cmd)

    If rs.EOF Then
        GetBalanceBefore = 0
    Else
        GetBalanceBefore = managerDatabase.FieldLong(rs, "balance")
    End If

    managerDatabase.CloseRecordset rs

End Function
