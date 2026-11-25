Attribute VB_Name = "serviceStockLedger"
Option Explicit

' =========================================================
' SERVICE - STOCK LEDGER (kartu stok)
'
' Built only from ACTIVE documents:
'   IN  = OK quantity of a receive, on its verification date
'   OUT = quantity of a request, on its request date
' =========================================================

Private Const MODULE_NAME As String = "serviceStockLedger"

' =========================================================
' ERROR CODES (17800 - 17899)
' Usage: managerError.ThrowError serviceStockLedger.InvalidMonth, ...
' =========================================================
Public Const InvalidMonth As Long = 17801
Public Const InvalidYear As Long = 17802
Public Const EmptySparepart As Long = 17803


' Stock card of one sparepart for one month. Every entry has
' BeginningStock and EndStock filled; the first BeginningStock
' is the balance at the start of the month.
' beginningStock returns that opening balance (also when the
' month has no movement).
Public Function GetStockCard( _
    ByVal plantName As String, _
    ByVal sparepartId As Long, _
    ByVal yearNumber As Integer, _
    ByVal monthNumber As Integer, _
    ByRef beginningStock As Long) As Collection

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim plantItem As ManufacturingPlant
    Dim entries As Collection
    Dim entry As StockLedgerEntry
    Dim runningStock As Long
    Dim errInfo As ErrorInfo

    If monthNumber < 1 Or monthNumber > 12 Then
        managerError.ThrowError serviceStockLedger.InvalidMonth, _
            MODULE_NAME & ".GetStockCard", _
            "Bulan harus 1 - 12."
    End If

    If yearNumber < 1900 Or yearNumber > 9998 Then
        managerError.ThrowError serviceStockLedger.InvalidYear, _
            MODULE_NAME & ".GetStockCard", _
            "Tahun tidak valid."
    End If

    If sparepartId <= 0 Then
        managerError.ThrowError serviceStockLedger.EmptySparepart, _
            MODULE_NAME & ".GetStockCard", _
            "Sparepart belum dipilih."
    End If

    Set conn = managerDatabase.GetConnection()
    Set plantItem = serviceManufacturingPlant.FindPlant(conn, plantName)

    beginningStock = daoStockLedger.GetBalance( _
        conn, plantItem.Id, sparepartId, DateSerial(yearNumber, monthNumber, 1))

    Set entries = daoStockLedger.ListByMonth( _
        conn, plantItem.Id, sparepartId, yearNumber, monthNumber)

    runningStock = beginningStock

    For Each entry In entries
        entry.ApplyBeginningStock runningStock
        runningStock = entry.EndStock
    Next entry

    Set GetStockCard = entries

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".GetStockCard", _
        "Gagal mengambil kartu stok."

End Function


' One page of the stock card of one sparepart for one month
' (see PageRequest; keys KeyDate / KeyText / KeyText2 from
' StockCardKey). Every entry has BeginningStock and EndStock
' filled; the first BeginningStock is the stock just before
' the first row of the page.
'   totalRows      : movements of the whole month
'   beginningStock : stock at the start of the page (the
'                    month's opening stock when the month has
'                    no movement)
Public Function GetStockCardPage( _
    ByVal plantName As String, _
    ByVal sparepartId As Long, _
    ByVal yearNumber As Integer, _
    ByVal monthNumber As Integer, _
    ByVal page As PageRequest, _
    ByRef totalRows As Long, _
    ByRef beginningStock As Long) As Collection

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim plantItem As ManufacturingPlant
    Dim entries As Collection
    Dim entry As StockLedgerEntry
    Dim firstEntry As StockLedgerEntry
    Dim runningStock As Long
    Dim errInfo As ErrorInfo

    RequireCardInput sparepartId, yearNumber, monthNumber, MODULE_NAME & ".GetStockCardPage"

    Set conn = managerDatabase.GetConnection()
    Set plantItem = serviceManufacturingPlant.FindPlant(conn, plantName)

    totalRows = daoStockLedger.CountByMonth( _
        conn, plantItem.Id, sparepartId, yearNumber, monthNumber)

    Set entries = daoStockLedger.ListByMonth( _
        conn, plantItem.Id, sparepartId, yearNumber, monthNumber, page)

    If entries.Count = 0 Then
        beginningStock = daoStockLedger.GetBalance( _
            conn, plantItem.Id, sparepartId, DateSerial(yearNumber, monthNumber, 1))
    Else
        Set firstEntry = entries(1)
        beginningStock = daoStockLedger.GetBalanceBefore( _
            conn, plantItem.Id, sparepartId, _
            firstEntry.MovementDate, firstEntry.Direction, EntryReference(firstEntry))
    End If

    runningStock = beginningStock

    For Each entry In entries
        entry.ApplyBeginningStock runningStock
        runningStock = entry.EndStock
    Next entry

    Set GetStockCardPage = entries

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".GetStockCardPage", _
        "Gagal mengambil kartu stok."

End Function


' The paging key of one stock card row.
Public Function StockCardKey( _
    ByVal entry As StockLedgerEntry) As PageRequest

    Dim keyHolder As New PageRequest

    Set StockCardKey = keyHolder.KeyOf( _
        entry.MovementDate, 0, entry.Direction, EntryReference(entry))

End Function


' Quantity that can leave the store on onDate without the
' stock card going negative then or later (caller's
' connection, so uncommitted changes count).
Public Function GetAvailableQuantity( _
    ByVal conn As ADODB.connection, _
    ByVal plantId As Long, _
    ByVal sparepartId As Long, _
    ByVal onDate As Date) As Long

    GetAvailableQuantity = daoStockLedger.GetAvailableQuantity( _
        conn, plantId, sparepartId, DateValue(onDate))

End Function


' =========================================================
' INTERNAL
' =========================================================

Private Sub RequireCardInput( _
    ByVal sparepartId As Long, _
    ByVal yearNumber As Integer, _
    ByVal monthNumber As Integer, _
    ByVal errorSource As String)

    If monthNumber < 1 Or monthNumber > 12 Then
        managerError.ThrowError serviceStockLedger.InvalidMonth, _
            errorSource, "Bulan harus 1 - 12."
    End If

    If yearNumber < 1900 Or yearNumber > 9998 Then
        managerError.ThrowError serviceStockLedger.InvalidYear, _
            errorSource, "Tahun tidak valid."
    End If

    If sparepartId <= 0 Then
        managerError.ThrowError serviceStockLedger.EmptySparepart, _
            errorSource, "Sparepart belum dipilih."
    End If

End Sub


' Reference of the row (receive or request number).
Private Function EntryReference( _
    ByVal entry As StockLedgerEntry) As String

    If entry.Direction = "IN" Then
        EntryReference = entry.IncomingReference
    Else
        EntryReference = entry.OutgoingReference
    End If

End Function
