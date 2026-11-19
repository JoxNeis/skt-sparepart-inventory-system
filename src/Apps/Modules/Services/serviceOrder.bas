Attribute VB_Name = "serviceOrder"
Option Explicit

' =========================================================
' SERVICE - ORDER
'
' Public entry points (called by sheets) get the
' connection and own the transaction:
'   CreateOrder, SelectOrder, SelectPendingOrder,
'   UpdateOrderPOID
' =========================================================

Private Const MODULE_NAME As String = "SERVICE_ORDER"

' =========================================================
' ERROR CODES (17400 - 17499)
' Usage: ErrorManager.ThrowError serviceOrder.EmptyOrder, ...
' =========================================================
Public Const EmptyOrder As Long = 17401                     ' Order object is Nothing
Public Const EmptyTransactionUUID As Long = 17402           ' CreateTransaction returned no UUID
Public Const SameOriginalProposedUUID As Long = 17403       ' original UUID = proposed UUID
Public Const NoRowAffectedCreateOrder As Long = 17404       ' create_order inserted 0 rows
Public Const EmptyManufacturingPlantName As Long = 17405
Public Const InvalidDateRange As Long = 17406               ' untilDate < fromDate
Public Const EmptyOrderUUID As Long = 17407
Public Const EmptyPoId As Long = 17408
Public Const NoRowAffectedUpdatePoId As Long = 17409        ' update_order_po_id updated 0 rows


' =========================================================
' PUBLIC ENTRY POINTS
' =========================================================

' Create an order in one transaction.
'
' originalUUID = ""  : normal order (active transaction).
' originalUUID <> "" : proposed update of that order
'                      (inactive transaction + update
'                      request + receives moved to the
'                      proposed order).
Public Sub CreateOrder( _
    ByVal orderItem As Order, _
    Optional ByVal originalUUID As String = "")

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim affected As Long
    Dim receiveAffected As Long
    Dim transactionUUID As String
    Dim isCreatedForUpdate As Boolean
    Dim activityCode As String
    Dim inTransaction As Boolean
    Dim errInfo As ErrorInfo

    inTransaction = False

    ' =====================================================
    ' Step 1: Validate
    ' =====================================================
    If orderItem Is Nothing Then
        ErrorManager.ThrowError _
            serviceOrder.EmptyOrder, _
            MODULE_NAME & ".CreateOrder", _
            "Order tidak boleh kosong."
    End If

    originalUUID = Trim$(originalUUID)
    isCreatedForUpdate = (Len(originalUUID) > 0)

    ' =====================================================
    ' Step 2: Begin transaction
    ' =====================================================
    Set conn = DatabaseManager.GetConnection()

    DatabaseManager.BeginTransaction conn
    inTransaction = True

    ' =====================================================
    ' Step 3: Create the transaction
    '
    ' Normal order:
    '   active = True
    '
    ' Proposed update:
    '   active = False
    ' =====================================================
    transactionUUID = serviceTransaction.CreateTransaction( _
        conn:=conn, _
        transactionType:="ORDER", _
        transactionDate:=orderItem.orderDate, _
        isActive:=Not isCreatedForUpdate)

    transactionUUID = Trim$(transactionUUID)

    If ValidationText.IsBlank(transactionUUID) Then
        ErrorManager.ThrowError _
            serviceOrder.EmptyTransactionUUID, _
            MODULE_NAME & ".CreateOrder", _
            "UUID transaksi order tidak berhasil dibuat."
    End If

    If isCreatedForUpdate Then
        If ValidationText.IsSameText(originalUUID, transactionUUID) Then
            ErrorManager.ThrowError _
                serviceOrder.SameOriginalProposedUUID, _
                MODULE_NAME & ".CreateOrder", _
                "UUID transaksi original dan proposed " & _
                "tidak boleh sama."
        End If
    End If

    ' =====================================================
    ' Step 4: Create the order
    ' =====================================================
    affected = daoOrder.CreateOrder(conn, transactionUUID, orderItem)

    If ValidationNumeric.IsLessThan(affected, 1) Then
        ErrorManager.ThrowError _
            serviceOrder.NoRowAffectedCreateOrder, _
            MODULE_NAME & ".CreateOrder", _
            "Order tidak tersimpan (0 baris terpengaruh)."
    End If

    ' =====================================================
    ' Step 5: Create update request and move receives
    ' (0 moved receives is valid)
    ' =====================================================
    receiveAffected = 0

    If isCreatedForUpdate Then
        serviceTransactionRequestUpdate.CreateRequestUpdate _
            conn, originalUUID, transactionUUID

        receiveAffected = daoUpdateOrder.UpdateOrderReceives( _
            conn, originalUUID, transactionUUID)
    End If

    ' =====================================================
    ' Step 6: Write activity log
    ' =====================================================
    If isCreatedForUpdate Then
        activityCode = "Order.CreateUpdateProposal"
    Else
        activityCode = "Order.Create"
    End If

    serviceActivityLog.WriteActivityLog _
        conn, _
        activityCode, _
        BuildOrderDetail(orderItem, transactionUUID) & _
        BuildOrderUpdateDetail(isCreatedForUpdate, originalUUID) & _
        BuildReceiveTransferDetail(isCreatedForUpdate, receiveAffected)

    ' =====================================================
    ' Step 7: Commit
    ' =====================================================
    DatabaseManager.CommitTransaction conn
    inTransaction = False

    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    If inTransaction Then
        DatabaseManager.RollbackTransaction conn
    End If

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".CreateOrder", _
        "Gagal membuat order."

End Sub


' Orders (with receive information) of one manufacturing
' plant between fromDate and untilDate, keyed by order UUID.
Public Function SelectOrder( _
    ByVal manufacturingPlantName As String, _
    ByVal fromDate As Date, _
    ByVal untilDate As Date) As Collection

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim errInfo As ErrorInfo

    manufacturingPlantName = Trim$(manufacturingPlantName)

    RequireSearchParameters manufacturingPlantName, fromDate, untilDate, _
        MODULE_NAME & ".SelectOrder"

    Set conn = DatabaseManager.GetConnection()

    Set SelectOrder = daoOrder.SelectOrder( _
        conn, manufacturingPlantName, DateValue(fromDate), DateValue(untilDate))

    Exit Function

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".SelectOrder", _
        "Gagal mengambil data order dan penerimaan."

End Function


' Pending orders of one manufacturing plant between
' fromDate and untilDate, keyed by order UUID.
Public Function SelectPendingOrder( _
    ByVal manufacturingPlantName As String, _
    ByVal fromDate As Date, _
    ByVal untilDate As Date) As Collection

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim errInfo As ErrorInfo

    manufacturingPlantName = Trim$(manufacturingPlantName)

    RequireSearchParameters manufacturingPlantName, fromDate, untilDate, _
        MODULE_NAME & ".SelectPendingOrder"

    Set conn = DatabaseManager.GetConnection()

    Set SelectPendingOrder = daoOrder.SelectPendingOrder( _
        conn, manufacturingPlantName, DateValue(fromDate), DateValue(untilDate))

    Exit Function

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".SelectPendingOrder", _
        "Gagal mengambil data order dan penerimaan."

End Function


' Set the PO ID of one order and write the activity log,
' in one transaction.
Public Sub UpdateOrderPOID( _
    ByVal orderUUID As String, _
    ByVal poID As String)

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim affected As Long
    Dim inTransaction As Boolean
    Dim errInfo As ErrorInfo

    inTransaction = False

    orderUUID = Trim$(orderUUID)
    poID = Trim$(poID)

    ' =====================================================
    ' Step 1: Validate
    ' =====================================================
    If ValidationText.IsBlank(orderUUID) Then
        ErrorManager.ThrowError _
            serviceOrder.EmptyOrderUUID, _
            MODULE_NAME & ".UpdateOrderPOID", _
            "UUID order tidak boleh kosong."
    End If

    If ValidationText.IsBlank(poID) Then
        ErrorManager.ThrowError _
            serviceOrder.EmptyPoId, _
            MODULE_NAME & ".UpdateOrderPOID", _
            "PO ID tidak boleh kosong."
    End If

    ' =====================================================
    ' Step 2: Begin transaction
    ' =====================================================
    Set conn = DatabaseManager.GetConnection()

    DatabaseManager.BeginTransaction conn
    inTransaction = True

    ' =====================================================
    ' Step 3: Update the order PO ID
    ' =====================================================
    affected = daoOrder.UpdateOrderPOID(conn, orderUUID, poID)

    If ValidationNumeric.IsLessThan(affected, 1) Then
        ErrorManager.ThrowError _
            serviceOrder.NoRowAffectedUpdatePoId, _
            MODULE_NAME & ".UpdateOrderPOID", _
            "PO ID order tidak tersimpan (0 baris terpengaruh)."
    End If

    ' =====================================================
    ' Step 4: Write activity log
    ' =====================================================
    serviceActivityLog.WriteActivityLog _
        conn, _
        "Order.UpdatePoId", _
        BuildUpdateOrderPoIdDetail(orderUUID, poID)

    ' =====================================================
    ' Step 5: Commit
    ' =====================================================
    DatabaseManager.CommitTransaction conn
    inTransaction = False

    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    If inTransaction Then
        DatabaseManager.RollbackTransaction conn
    End If

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".UpdateOrderPOID", _
        "Gagal memperbarui PO ID order."

End Sub


' =========================================================
' INTERNAL
' =========================================================

Private Sub RequireSearchParameters( _
    ByVal manufacturingPlantName As String, _
    ByVal fromDate As Date, _
    ByVal untilDate As Date, _
    ByVal errorSource As String)

    If ValidationText.IsBlank(manufacturingPlantName) Then
        ErrorManager.ThrowError _
            serviceOrder.EmptyManufacturingPlantName, _
            errorSource, _
            "Nama manufacturing plant tidak boleh kosong."
    End If

    If DateValue(untilDate) < DateValue(fromDate) Then
        ErrorManager.ThrowError _
            serviceOrder.InvalidDateRange, _
            errorSource, _
            "Tanggal akhir tidak boleh lebih kecil dari tanggal awal."
    End If

End Sub


Private Function BuildOrderDetail( _
    ByVal orderItem As Order, _
    ByVal transactionUUID As String) As String

    Dim detail As String

    detail = "UUID: " & transactionUUID & vbCrLf & _
             "Order Date: " & Format$(orderItem.orderDate, "yyyy-mm-dd") & vbCrLf & _
             "Catalogue: " & orderItem.Catalogue.Name & vbCrLf & _
             "Amount: " & orderItem.Amount & vbCrLf & _
             "Plant: " & orderItem.Plant.Name & vbCrLf & _
             "Vendor: " & orderItem.Vendor.Name & vbCrLf & _
             "Remarks: " & Trim$(orderItem.remarks) & vbCrLf & _
             "Person In Charge: " & orderItem.personInCharge

    BuildOrderDetail = detail

End Function


Private Function BuildUpdateOrderPoIdDetail( _
    ByVal orderUUID As String, _
    ByVal poID As String) As String

    Dim detail As String

    detail = _
        "UUID: " & Trim$(orderUUID) & vbCrLf & _
        "PO ID: " & Trim$(poID)

    BuildUpdateOrderPoIdDetail = detail

End Function


Private Function BuildOrderUpdateDetail( _
    ByVal isCreatedForUpdate As Boolean, _
    ByVal originalUUID As String) As String

    If isCreatedForUpdate Then
        BuildOrderUpdateDetail = _
            vbCrLf & _
            "Created For Update: Yes" & vbCrLf & _
            "Original UUID: " & Trim$(originalUUID)
    Else
        BuildOrderUpdateDetail = _
            vbCrLf & _
            "Created For Update: No"
    End If

End Function


Private Function BuildReceiveTransferDetail( _
    ByVal isCreatedForUpdate As Boolean, _
    ByVal receiveAffected As Long) As String

    If Not isCreatedForUpdate Then
        BuildReceiveTransferDetail = vbNullString
        Exit Function
    End If

    BuildReceiveTransferDetail = _
        vbCrLf & _
        "Receive Rows Transferred: " & CStr(receiveAffected)

End Function
