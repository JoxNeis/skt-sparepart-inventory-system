Attribute VB_Name = "serviceOrder"
Option Explicit

' =========================================================
' SERVICE - ORDER
'
' An order does not move stock. Public entry points get the
' connection and own the transaction:
'   CreateOrders         new active orders (one batch, all or nothing)
'   ListOrders           order recap
'   ListPendingOrders    orders with a remaining quantity
'   SetPoNumber          the only field changed directly
'   ProposeOrderUpdate   correction -> draft + update request
'   ProposeOrderDelete   same, with quantity 0
' =========================================================

Private Const MODULE_NAME As String = "serviceOrder"

' =========================================================
' ERROR CODES (17400 - 17499)
' Usage: managerError.ThrowError serviceOrder.EmptyOrderList, ...
' =========================================================
Public Const EmptyOrderList As Long = 17401
Public Const EmptyOrder As Long = 17402
Public Const EmptySparepart As Long = 17403
Public Const ZeroQuantity As Long = 17404           ' new order with quantity 0
Public Const EmptyOrderDate As Long = 17405
Public Const EmptyPersonInCharge As Long = 17406
Public Const CreateFailed As Long = 17407           ' order_create inserted 0 rows
Public Const InvalidDateRange As Long = 17408
Public Const EmptyOrderUUID As Long = 17409
Public Const SetPoNumberFailed As Long = 17410
Public Const EmptyReason As Long = 17411
Public Const NothingChanged As Long = 17412
Public Const AlreadyDeleted As Long = 17413


' =========================================================
' CREATE
' =========================================================

' Create every order of the list as an active document of
' the plant, in one transaction. Each Order needs Sparepart
' (with Id), Quantity > 0, OrderDate and PersonInCharge;
' Vendor is optional. On success every order has its
' Transaction (document number) set. Returns the numbers.
Public Function CreateOrders( _
    ByVal plantName As String, _
    ByVal orders As Collection) As Collection

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim plantItem As ManufacturingPlant
    Dim creator As User
    Dim orderItem As Order
    Dim uuids As Collection
    Dim transactionUUID As String
    Dim affected As Long
    Dim inTransaction As Boolean
    Dim errInfo As ErrorInfo

    If orders Is Nothing Then
        managerError.ThrowError serviceOrder.EmptyOrderList, _
            MODULE_NAME & ".CreateOrders", "Tidak ada order untuk disimpan."
    End If

    If orders.Count = 0 Then
        managerError.ThrowError serviceOrder.EmptyOrderList, _
            MODULE_NAME & ".CreateOrders", "Tidak ada order untuk disimpan."
    End If

    For Each orderItem In orders
        ValidateOrder orderItem, True, MODULE_NAME & ".CreateOrders"
    Next orderItem

    Set conn = managerDatabase.GetConnection()
    Set plantItem = serviceManufacturingPlant.FindPlant(conn, plantName)
    Set creator = serviceUser.RequireRegisteredUser(conn)

    Set uuids = New Collection

    managerDatabase.BeginTransaction conn
    inTransaction = True

    For Each orderItem In orders

        transactionUUID = serviceTransaction.CreateTransaction( _
            conn, plantItem, "ORDER", orderItem.OrderDate, True, creator.Username)

        affected = daoOrder.CreateOrder(conn, transactionUUID, orderItem)

        managerDatabase.RequireAffectedRows affected, 1, _
            serviceOrder.CreateFailed, _
            MODULE_NAME & ".CreateOrders", _
            "Order " & orderItem.Sparepart.DisplayName & _
            " tidak tersimpan: sparepart atau vendor tidak aktif."

        Set orderItem.Transaction = BuildHeader( _
            transactionUUID, plantItem, creator.Username, True)

        serviceActivityLog.WriteActivityLog conn, _
            "Order.Create", _
            "Order " & transactionUUID & ": " & orderItem.Sparepart.DisplayName & _
            " x " & orderItem.Quantity & " " & orderItem.UnitOfMeasurement & ".", _
            BuildOrderJson(orderItem, plantItem)

        uuids.Add transactionUUID

    Next orderItem

    managerDatabase.CommitTransaction conn
    inTransaction = False

    Set CreateOrders = uuids

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()
    If inTransaction Then managerDatabase.RollbackTransaction conn

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".CreateOrders", _
        "Gagal menyimpan order. Tidak ada order yang tersimpan."

End Function


' =========================================================
' LIST
' =========================================================

' Active, not deleted orders of the plant with an order date
' in the range, keyed by document number. page: filter and
' page (see PageRequest); Nothing = every row.
Public Function ListOrders( _
    ByVal plantName As String, _
    ByVal fromDate As Date, _
    ByVal toDate As Date, _
    Optional ByVal page As PageRequest = Nothing) As Collection

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim errInfo As ErrorInfo

    RequireDateRange fromDate, toDate, MODULE_NAME & ".ListOrders"

    Set conn = managerDatabase.GetConnection()

    Set ListOrders = daoOrder.ListOrders( _
        conn, serviceManufacturingPlant.FindPlant(conn, plantName), _
        fromDate, toDate, page)

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ListOrders", _
        "Gagal mengambil data order."

End Function

' Number of rows ListOrders gives over all pages (same filter).
Public Function CountOrders( _
    ByVal plantName As String, _
    ByVal fromDate As Date, _
    ByVal toDate As Date, _
    Optional ByVal page As PageRequest = Nothing) As Long

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim errInfo As ErrorInfo

    Set conn = managerDatabase.GetConnection()

    CountOrders = daoOrder.CountOrders( _
        conn, serviceManufacturingPlant.FindPlant(conn, plantName), _
        fromDate, toDate, page)

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".CountOrders", _
        "Gagal menghitung data order."

End Function



' Active orders with a remaining quantity, keyed by document
' number.
Public Function ListPendingOrders( _
    ByVal plantName As String, _
    ByVal fromDate As Date, _
    ByVal toDate As Date, _
    Optional ByVal page As PageRequest = Nothing) As Collection

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim errInfo As ErrorInfo

    RequireDateRange fromDate, toDate, MODULE_NAME & ".ListPendingOrders"

    Set conn = managerDatabase.GetConnection()

    Set ListPendingOrders = daoOrder.ListPendingOrders( _
        conn, serviceManufacturingPlant.FindPlant(conn, plantName), _
        fromDate, toDate, page)

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ListPendingOrders", _
        "Gagal mengambil data pending order."

End Function

' Number of rows ListPendingOrders gives over all pages.
Public Function CountPendingOrders( _
    ByVal plantName As String, _
    ByVal fromDate As Date, _
    ByVal toDate As Date, _
    Optional ByVal page As PageRequest = Nothing) As Long

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim errInfo As ErrorInfo

    Set conn = managerDatabase.GetConnection()

    CountPendingOrders = daoOrder.CountPendingOrders( _
        conn, serviceManufacturingPlant.FindPlant(conn, plantName), _
        fromDate, toDate, page)

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".CountPendingOrders", _
        "Gagal menghitung data pending order."

End Function



' =========================================================
' PO NUMBER (no approval needed)
' =========================================================

Public Sub SetPoNumber( _
    ByVal orderUUID As String, _
    ByVal poNumber As String)

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim affected As Long
    Dim inTransaction As Boolean
    Dim errInfo As ErrorInfo

    orderUUID = Trim$(orderUUID)
    poNumber = Trim$(poNumber)

    If managerValidationText.IsBlank(orderUUID) Then
        managerError.ThrowError serviceOrder.EmptyOrderUUID, _
            MODULE_NAME & ".SetPoNumber", "Nomor order tidak boleh kosong."
    End If

    Set conn = managerDatabase.GetConnection()

    managerDatabase.BeginTransaction conn
    inTransaction = True

    affected = daoOrder.SetPoNumber(conn, orderUUID, poNumber)

    managerDatabase.RequireAffectedRows affected, 1, _
        serviceOrder.SetPoNumberFailed, _
        MODULE_NAME & ".SetPoNumber", _
        "Order " & orderUUID & " tidak ditemukan."

    serviceActivityLog.WriteActivityLog conn, _
        "Order.SetPoNumber", _
        "PO order " & orderUUID & " = '" & poNumber & "'.", _
        coreJson.JsonObject( _
            "uuid", coreJson.JsonText(orderUUID), _
            "po_number", coreJson.JsonText(poNumber))

    managerDatabase.CommitTransaction conn
    inTransaction = False

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    If inTransaction Then managerDatabase.RollbackTransaction conn

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".SetPoNumber", _
        "Gagal menyimpan PO ID."

End Sub


' =========================================================
' UPDATE PROPOSAL
' =========================================================

' Save the corrected order as a DRAFT and create a pending
' update request for the original. Nothing else changes
' until a higher-tier user approves it. Returns the number
' of the proposed (draft) order.
Public Function ProposeOrderUpdate( _
    ByVal plantName As String, _
    ByVal original As Order, _
    ByVal proposed As Order, _
    ByVal reason As String) As String

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim plantItem As ManufacturingPlant
    Dim requester As User
    Dim proposedUUID As String
    Dim affected As Long
    Dim inTransaction As Boolean
    Dim errInfo As ErrorInfo

    If original Is Nothing Then
        managerError.ThrowError serviceOrder.EmptyOrder, _
            MODULE_NAME & ".ProposeOrderUpdate", "Order asal tidak boleh kosong."
    End If

    ValidateOrder proposed, False, MODULE_NAME & ".ProposeOrderUpdate"
    RequireReason reason, MODULE_NAME & ".ProposeOrderUpdate"

    If Not HasChanges(original, proposed) Then
        managerError.ThrowError serviceOrder.NothingChanged, _
            MODULE_NAME & ".ProposeOrderUpdate", _
            "Tidak ada perubahan pada order " & original.UUID & "."
    End If

    ' The PO number is not part of an update: keep the
    ' original one (it is changed directly with SetPoNumber).
    proposed.PoNumber = original.PoNumber

    Set conn = managerDatabase.GetConnection()
    Set plantItem = serviceManufacturingPlant.FindPlant(conn, plantName)
    Set requester = serviceUser.RequireRegisteredUser(conn)

    managerDatabase.BeginTransaction conn
    inTransaction = True

    proposedUUID = serviceTransaction.CreateTransaction( _
        conn, plantItem, "ORDER", proposed.OrderDate, False, requester.Username)

    affected = daoOrder.CreateOrder(conn, proposedUUID, proposed)

    managerDatabase.RequireAffectedRows affected, 1, _
        serviceOrder.CreateFailed, _
        MODULE_NAME & ".ProposeOrderUpdate", _
        "Order usulan tidak tersimpan: sparepart atau vendor tidak aktif."

    Set proposed.Transaction = BuildHeader( _
        proposedUUID, plantItem, requester.Username, False)

    serviceTransactionUpdateRequest.CreateUpdateRequest _
        conn, original.UUID, proposedUUID, reason, requester.Username

    serviceActivityLog.WriteActivityLog conn, _
        "Order.CreateUpdateProposal", _
        "Usulan update order " & original.UUID & " -> " & proposedUUID & _
        ". Alasan: " & reason, _
        coreJson.JsonAppend( _
            coreJson.JsonAppend(BuildOrderJson(proposed, plantItem), _
                "original_uuid", coreJson.JsonText(original.UUID)), _
            "reason", coreJson.JsonText(reason))

    managerDatabase.CommitTransaction conn
    inTransaction = False

    ProposeOrderUpdate = proposedUUID

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()
    If inTransaction Then managerDatabase.RollbackTransaction conn

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ProposeOrderUpdate", _
        "Gagal membuat usulan update order."

End Function


' Propose to delete an order: an update with quantity 0.
Public Function ProposeOrderDelete( _
    ByVal plantName As String, _
    ByVal original As Order, _
    ByVal reason As String) As String

    Dim proposed As Order

    If original Is Nothing Then
        managerError.ThrowError serviceOrder.EmptyOrder, _
            MODULE_NAME & ".ProposeOrderDelete", "Order asal tidak boleh kosong."
    End If

    If original.Quantity = 0 Then
        managerError.ThrowError serviceOrder.AlreadyDeleted, _
            MODULE_NAME & ".ProposeOrderDelete", _
            "Order " & original.UUID & " sudah dihapus (jumlah 0)."
    End If

    Set proposed = CopyOrder(original)
    proposed.Quantity = 0

    ProposeOrderDelete = ProposeOrderUpdate(plantName, original, proposed, reason)

End Function


' A new Order with the same detail values as source (no
' document header, no progress).
Public Function CopyOrder( _
    ByVal source As Order) As Order

    Dim result As Order

    Set result = New Order

    With result
        Set .Sparepart = source.Sparepart
        If source.HasVendor Then Set .Vendor = source.Vendor
        .PoNumber = source.PoNumber
        .OrderDate = source.OrderDate
        .Quantity = source.Quantity
        .Remarks = source.Remarks
        .SystemRemarks = source.SystemRemarks
        If Len(source.PersonInCharge) > 0 Then .PersonInCharge = source.PersonInCharge
    End With

    Set CopyOrder = result

End Function


' Activity log JSON of an order (README "Log JSON").
Public Function BuildOrderJson( _
    ByVal orderItem As Order, _
    ByVal plantItem As ManufacturingPlant) As String

    Dim vendorJson As String

    If orderItem.HasVendor Then
        vendorJson = serviceVendor.BuildVendorReferenceJson(orderItem.Vendor)
    Else
        vendorJson = "null"
    End If

    BuildOrderJson = coreJson.JsonObject( _
        "uuid", coreJson.JsonText(orderItem.UUID), _
        "transaction_type", coreJson.JsonText("ORDER"), _
        "manufacturing_plant", coreJson.JsonObject( _
            "code", coreJson.JsonText(plantItem.Code), _
            "name", coreJson.JsonText(plantItem.Name)), _
        "sparepart", serviceSparepart.BuildSparepartReferenceJson(orderItem.Sparepart), _
        "vendor", vendorJson, _
        "po_number", coreJson.JsonText(orderItem.PoNumber), _
        "order_date", coreJson.JsonDate(orderItem.OrderDate), _
        "quantity", coreJson.JsonNumber(orderItem.Quantity), _
        "remarks", coreJson.JsonText(orderItem.Remarks), _
        "system_remarks", coreJson.JsonText(orderItem.SystemRemarks), _
        "person_in_charge", coreJson.JsonText(orderItem.PersonInCharge))

End Function


' =========================================================
' INTERNAL
' =========================================================

Private Sub ValidateOrder( _
    ByVal orderItem As Order, _
    ByVal isNew As Boolean, _
    ByVal errorSource As String)

    If orderItem Is Nothing Then
        managerError.ThrowError serviceOrder.EmptyOrder, _
            errorSource, "Order tidak boleh kosong."
    End If

    If orderItem.Sparepart Is Nothing Then
        managerError.ThrowError serviceOrder.EmptySparepart, _
            errorSource, "Sparepart order belum dipilih."
    End If

    If orderItem.Sparepart.Id <= 0 Then
        managerError.ThrowError serviceOrder.EmptySparepart, _
            errorSource, "Sparepart order tidak dikenal."
    End If

    If isNew And orderItem.Quantity <= 0 Then
        managerError.ThrowError serviceOrder.ZeroQuantity, _
            errorSource, _
            "Jumlah order " & orderItem.Sparepart.DisplayName & _
            " harus lebih besar dari 0."
    End If

    If orderItem.OrderDate = 0 Then
        managerError.ThrowError serviceOrder.EmptyOrderDate, _
            errorSource, "Tanggal order belum diisi."
    End If

    If Len(orderItem.PersonInCharge) = 0 Then
        managerError.ThrowError serviceOrder.EmptyPersonInCharge, _
            errorSource, _
            "PIC order " & orderItem.Sparepart.DisplayName & " belum diisi."
    End If

End Sub


Private Function HasChanges( _
    ByVal original As Order, _
    ByVal proposed As Order) As Boolean

    Dim originalVendorId As Long
    Dim proposedVendorId As Long

    If original.HasVendor Then originalVendorId = original.Vendor.Id
    If proposed.HasVendor Then proposedVendorId = proposed.Vendor.Id

    HasChanges = _
        original.Sparepart.Id <> proposed.Sparepart.Id _
        Or originalVendorId <> proposedVendorId _
        Or DateValue(original.OrderDate) <> DateValue(proposed.OrderDate) _
        Or original.Quantity <> proposed.Quantity _
        Or original.Remarks <> proposed.Remarks _
        Or original.PersonInCharge <> proposed.PersonInCharge

End Function


Private Sub RequireReason( _
    ByVal reason As String, _
    ByVal errorSource As String)

    If managerValidationText.IsBlank(reason) Then
        managerError.ThrowError serviceOrder.EmptyReason, _
            errorSource, "Alasan perubahan wajib diisi."
    End If

End Sub


Private Sub RequireDateRange( _
    ByVal fromDate As Date, _
    ByVal toDate As Date, _
    ByVal errorSource As String)

    If DateValue(toDate) < DateValue(fromDate) Then
        managerError.ThrowError serviceOrder.InvalidDateRange, _
            errorSource, _
            "Tanggal akhir tidak boleh lebih kecil dari tanggal awal."
    End If

End Sub


Private Function BuildHeader( _
    ByVal transactionUUID As String, _
    ByVal plantItem As ManufacturingPlant, _
    ByVal createdBy As String, _
    ByVal isActive As Boolean) As Transaction

    Dim header As Transaction

    Set header = New Transaction

    header.UUID = transactionUUID
    header.TransactionType = "ORDER"
    Set header.ManufacturingPlant = plantItem
    header.CreatedBy = createdBy
    header.CreatedAt = Now
    If isActive Then header.ActivatedAt = Now

    Set BuildHeader = header

End Function
