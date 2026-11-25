Attribute VB_Name = "serviceDashboard"
Option Explicit

' =========================================================
' SERVICE - DASHBOARD
'
' Totals of one plant for a period, read in one go:
'   ORDER       total, ongoing, finished, ordered quantity
'   PENERIMAAN  total, received, OK, NOT OK
'   PENGELUARAN total, quantity issued, sparepart kinds
'   URGENT      spareparts at or below their reorder point
' Deleted documents (quantity 0) are not counted.
' =========================================================

Private Const MODULE_NAME As String = "serviceDashboard"

' =========================================================
' ERROR CODES (18400 - 18499)
' =========================================================
Public Const InvalidDateRange As Long = 18401


' Returns a Scripting.Dictionary with the keys
'   order_count, order_ongoing, order_finished, order_quantity,
'   receive_count, receive_quantity, receive_ok, receive_not_ok,
'   request_count, request_quantity, request_kinds,
'   urgent (Collection of ManufacturingPlantStock at or
'           below the reorder point, by sparepart name)
Public Function GetSummary( _
    ByVal plantName As String, _
    ByVal fromDate As Date, _
    ByVal toDate As Date) As Object

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim plantItem As ManufacturingPlant
    Dim summary As Object
    Dim kinds As Object
    Dim urgent As Collection
    Dim orderItem As Order
    Dim receiveItem As Receive
    Dim requestItem As Request
    Dim stockItem As ManufacturingPlantStock
    Dim errInfo As ErrorInfo

    If DateValue(toDate) < DateValue(fromDate) Then
        managerError.ThrowError serviceDashboard.InvalidDateRange, _
            MODULE_NAME & ".GetSummary", _
            "Tanggal akhir tidak boleh lebih kecil dari tanggal awal."
    End If

    Set conn = managerDatabase.GetConnection()
    Set plantItem = serviceManufacturingPlant.FindPlant(conn, plantName)

    Set summary = CreateObject("Scripting.Dictionary")

    summary("order_count") = 0
    summary("order_ongoing") = 0
    summary("order_finished") = 0
    summary("order_quantity") = 0

    For Each orderItem In daoOrder.ListOrders(conn, plantItem, fromDate, toDate)
        If orderItem.Quantity > 0 Then
            summary("order_count") = summary("order_count") + 1
            summary("order_quantity") = summary("order_quantity") + orderItem.Quantity
            If orderItem.IsFullyReceived Then
                summary("order_finished") = summary("order_finished") + 1
            Else
                summary("order_ongoing") = summary("order_ongoing") + 1
            End If
        End If
    Next orderItem

    summary("receive_count") = 0
    summary("receive_quantity") = 0
    summary("receive_ok") = 0
    summary("receive_not_ok") = 0

    For Each receiveItem In daoReceive.ListReceives(conn, plantItem, fromDate, toDate)
        If receiveItem.ReceivedQuantity > 0 Then
            summary("receive_count") = summary("receive_count") + 1
            summary("receive_quantity") = summary("receive_quantity") + receiveItem.ReceivedQuantity
            summary("receive_ok") = summary("receive_ok") + receiveItem.OkQuantity
            summary("receive_not_ok") = summary("receive_not_ok") + receiveItem.NotOkQuantity
        End If
    Next receiveItem

    Set kinds = CreateObject("Scripting.Dictionary")

    summary("request_count") = 0
    summary("request_quantity") = 0

    For Each requestItem In daoRequest.ListRequests(conn, plantItem, fromDate, toDate)
        If requestItem.Quantity > 0 Then
            summary("request_count") = summary("request_count") + 1
            summary("request_quantity") = summary("request_quantity") + requestItem.Quantity
            kinds(CStr(requestItem.Sparepart.Id)) = True
        End If
    Next requestItem

    summary("request_kinds") = kinds.Count

    Set urgent = New Collection

    For Each stockItem In daoPlantStock.ListByPlant(conn, plantItem)
        If stockItem.ReorderPoint > 0 And stockItem.IsAtReorderPoint Then
            urgent.Add stockItem
        End If
    Next stockItem

    Set summary("urgent") = urgent

    Set GetSummary = summary

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".GetSummary", _
        "Gagal mengambil data dashboard."

End Function
