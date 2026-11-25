Attribute VB_Name = "serviceReceive"
Option Explicit

' =========================================================
' SERVICE - RECEIVE (penerimaan)
'
' Goods arriving against an active order. Only the OK
' quantity enters stock, on the verification date; a receive
' without verification date does not move stock yet.
'
' Rules:
'   received = OK + NOT OK
'   OK quantity <= the order's remaining quantity
'
' Public entry points get the connection and own the
' transaction:
'   CreateReceives, ListReceives,
'   ProposeReceiveUpdate, ProposeReceiveDelete
' =========================================================

Private Const MODULE_NAME As String = "serviceReceive"

' =========================================================
' ERROR CODES (17600 - 17699)
' Usage: managerError.ThrowError serviceReceive.OkExceedsRemaining, ...
' =========================================================
Public Const EmptyReceiveList As Long = 17601
Public Const EmptyReceive As Long = 17602
Public Const EmptyOrder As Long = 17603
Public Const EmptyArrivalDate As Long = 17604
Public Const ZeroReceivedQuantity As Long = 17605
Public Const QuantityNotBalanced As Long = 17606
Public Const OkExceedsRemaining As Long = 17607
Public Const EmptyPersonInCharge As Long = 17608
Public Const CreateFailed As Long = 17609          ' order not active / other plant
Public Const InvalidDateRange As Long = 17610
Public Const EmptyReason As Long = 17611
Public Const NothingChanged As Long = 17612
Public Const AlreadyDeleted As Long = 17613


' =========================================================
' CREATE
' =========================================================

' Create every receive of the list as an active document, in
' one transaction. Each Receive needs Order (UUID, Sparepart
' and RemainingQuantity, as loaded by
' serviceOrder.ListPendingOrders), ArrivalDate, quantities
' and PersonInCharge. Returns the new document numbers.
Public Function CreateReceives( _
    ByVal plantName As String, _
    ByVal receives As Collection) As Collection

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim plantItem As ManufacturingPlant
    Dim creator As User
    Dim receiveItem As Receive
    Dim uuids As Collection
    Dim okPerOrder As Object
    Dim touched As Object
    Dim sparepartKey As Variant
    Dim transactionUUID As String
    Dim affected As Long
    Dim inTransaction As Boolean
    Dim errInfo As ErrorInfo

    If receives Is Nothing Then
        managerError.ThrowError serviceReceive.EmptyReceiveList, _
            MODULE_NAME & ".CreateReceives", "Tidak ada penerimaan untuk disimpan."
    End If

    If receives.Count = 0 Then
        managerError.ThrowError serviceReceive.EmptyReceiveList, _
            MODULE_NAME & ".CreateReceives", "Tidak ada penerimaan untuk disimpan."
    End If

    ' Validate every row, and the OK total per order against
    ' its remaining quantity (several rows may share an order).
    Set okPerOrder = CreateObject("Scripting.Dictionary")

    For Each receiveItem In receives

        ValidateReceive receiveItem, True, MODULE_NAME & ".CreateReceives"

        okPerOrder(receiveItem.Order.UUID) = _
            okPerOrder(receiveItem.Order.UUID) + receiveItem.OkQuantity

        If okPerOrder(receiveItem.Order.UUID) > receiveItem.Order.RemainingQuantity Then
            managerError.ThrowError serviceReceive.OkExceedsRemaining, _
                MODULE_NAME & ".CreateReceives", _
                "Jumlah OK untuk order " & receiveItem.Order.UUID & " (" & _
                okPerOrder(receiveItem.Order.UUID) & _
                ") melebihi jumlah yang harus dipenuhi (" & _
                receiveItem.Order.RemainingQuantity & ")."
        End If

    Next receiveItem

    Set conn = managerDatabase.GetConnection()
    Set plantItem = serviceManufacturingPlant.FindPlant(conn, plantName)
    Set creator = serviceUser.RequireRegisteredUser(conn)

    Set uuids = New Collection
    Set touched = CreateObject("Scripting.Dictionary")

    managerDatabase.BeginTransaction conn
    inTransaction = True

    For Each receiveItem In receives

        transactionUUID = serviceTransaction.CreateTransaction( _
            conn, plantItem, "RECEIVE", receiveItem.ArrivalDate, True, creator.Username)

        affected = daoReceive.CreateReceive( _
            conn, transactionUUID, receiveItem.Order.UUID, receiveItem)

        managerDatabase.RequireAffectedRows affected, 1, _
            serviceReceive.CreateFailed, _
            MODULE_NAME & ".CreateReceives", _
            "Penerimaan untuk order " & receiveItem.Order.UUID & _
            " tidak tersimpan: order tidak aktif atau milik plant lain."

        Set receiveItem.Transaction = BuildHeader( _
            transactionUUID, plantItem, creator.Username, True)

        If receiveItem.IsVerified And receiveItem.OkQuantity > 0 Then
            touched(CStr(receiveItem.Order.Sparepart.Id)) = _
                receiveItem.Order.Sparepart.DisplayName
        End If

        serviceActivityLog.WriteActivityLog conn, _
            "Receive.Create", _
            "Penerimaan " & transactionUUID & " untuk order " & _
            receiveItem.Order.UUID & ": diterima " & receiveItem.ReceivedQuantity & _
            ", OK " & receiveItem.OkQuantity & ", NOT OK " & receiveItem.NotOkQuantity & ".", _
            BuildReceiveJson(receiveItem, plantItem)

        uuids.Add transactionUUID

    Next receiveItem

    For Each sparepartKey In touched.Keys
        servicePlantStock.RecalculateCurrentStock conn, plantItem.Id, CLng(sparepartKey)
    Next sparepartKey

    managerDatabase.CommitTransaction conn
    inTransaction = False

    Set CreateReceives = uuids

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()
    If inTransaction Then managerDatabase.RollbackTransaction conn

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".CreateReceives", _
        "Gagal menyimpan penerimaan. Tidak ada penerimaan yang tersimpan."

End Function


' =========================================================
' LIST
' =========================================================

' Active receives of the plant with an arrival date in the
' range, keyed by document number.
Public Function ListReceives( _
    ByVal plantName As String, _
    ByVal fromDate As Date, _
    ByVal toDate As Date, _
    Optional ByVal page As PageRequest = Nothing) As Collection

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim errInfo As ErrorInfo

    If DateValue(toDate) < DateValue(fromDate) Then
        managerError.ThrowError serviceReceive.InvalidDateRange, _
            MODULE_NAME & ".ListReceives", _
            "Tanggal akhir tidak boleh lebih kecil dari tanggal awal."
    End If

    Set conn = managerDatabase.GetConnection()

    Set ListReceives = daoReceive.ListReceives( _
        conn, serviceManufacturingPlant.FindPlant(conn, plantName), _
        fromDate, toDate, page)

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ListReceives", _
        "Gagal mengambil data penerimaan."

End Function

' Number of rows ListReceives gives over all pages (same filter).
Public Function CountReceives( _
    ByVal plantName As String, _
    ByVal fromDate As Date, _
    ByVal toDate As Date, _
    Optional ByVal page As PageRequest = Nothing) As Long

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim errInfo As ErrorInfo

    Set conn = managerDatabase.GetConnection()

    CountReceives = daoReceive.CountReceives( _
        conn, serviceManufacturingPlant.FindPlant(conn, plantName), _
        fromDate, toDate, page)

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".CountReceives", _
        "Gagal menghitung data penerimaan."

End Function



' =========================================================
' UPDATE PROPOSAL
' =========================================================

' Save the corrected receive (same order) as a DRAFT and
' create a pending update request for the original. Stock
' does not change until it is approved. Returns the number
' of the proposed (draft) receive.
Public Function ProposeReceiveUpdate( _
    ByVal plantName As String, _
    ByVal original As Receive, _
    ByVal proposed As Receive, _
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
        managerError.ThrowError serviceReceive.EmptyReceive, _
            MODULE_NAME & ".ProposeReceiveUpdate", "Penerimaan asal tidak boleh kosong."
    End If

    If proposed Is Nothing Then
        managerError.ThrowError serviceReceive.EmptyReceive, _
            MODULE_NAME & ".ProposeReceiveUpdate", "Penerimaan usulan tidak boleh kosong."
    End If

    ' A receive update keeps its order.
    Set proposed.Order = original.Order
    proposed.RemainingQuantityBefore = original.RemainingQuantityBefore

    ValidateReceive proposed, False, MODULE_NAME & ".ProposeReceiveUpdate"

    If original.Order.Quantity > 0 Then
        If proposed.OkQuantity > original.RemainingQuantityBefore Then
            managerError.ThrowError serviceReceive.OkExceedsRemaining, _
                MODULE_NAME & ".ProposeReceiveUpdate", _
                "Jumlah OK (" & proposed.OkQuantity & _
                ") melebihi jumlah yang harus dipenuhi (" & _
                original.RemainingQuantityBefore & ")."
        End If
    End If

    If managerValidationText.IsBlank(reason) Then
        managerError.ThrowError serviceReceive.EmptyReason, _
            MODULE_NAME & ".ProposeReceiveUpdate", "Alasan perubahan wajib diisi."
    End If

    If Not HasChanges(original, proposed) Then
        managerError.ThrowError serviceReceive.NothingChanged, _
            MODULE_NAME & ".ProposeReceiveUpdate", _
            "Tidak ada perubahan pada penerimaan " & original.UUID & "."
    End If

    Set conn = managerDatabase.GetConnection()
    Set plantItem = serviceManufacturingPlant.FindPlant(conn, plantName)
    Set requester = serviceUser.RequireRegisteredUser(conn)

    managerDatabase.BeginTransaction conn
    inTransaction = True

    proposedUUID = serviceTransaction.CreateTransaction( _
        conn, plantItem, "RECEIVE", proposed.ArrivalDate, False, requester.Username)

    affected = daoReceive.CreateReceive( _
        conn, proposedUUID, original.Order.UUID, proposed)

    managerDatabase.RequireAffectedRows affected, 1, _
        serviceReceive.CreateFailed, _
        MODULE_NAME & ".ProposeReceiveUpdate", _
        "Penerimaan usulan tidak tersimpan: order " & original.Order.UUID & _
        " tidak aktif."

    Set proposed.Transaction = BuildHeader( _
        proposedUUID, plantItem, requester.Username, False)

    serviceTransactionUpdateRequest.CreateUpdateRequest _
        conn, original.UUID, proposedUUID, reason, requester.Username

    serviceActivityLog.WriteActivityLog conn, _
        "Receive.CreateUpdateProposal", _
        "Usulan update penerimaan " & original.UUID & " -> " & proposedUUID & _
        ". Alasan: " & reason, _
        coreJson.JsonAppend( _
            coreJson.JsonAppend(BuildReceiveJson(proposed, plantItem), _
                "original_uuid", coreJson.JsonText(original.UUID)), _
            "reason", coreJson.JsonText(reason))

    managerDatabase.CommitTransaction conn
    inTransaction = False

    ProposeReceiveUpdate = proposedUUID

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()
    If inTransaction Then managerDatabase.RollbackTransaction conn

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ProposeReceiveUpdate", _
        "Gagal membuat usulan update penerimaan."

End Function


' Propose to delete a receive: an update with every quantity 0.
Public Function ProposeReceiveDelete( _
    ByVal plantName As String, _
    ByVal original As Receive, _
    ByVal reason As String) As String

    Dim proposed As Receive

    If original Is Nothing Then
        managerError.ThrowError serviceReceive.EmptyReceive, _
            MODULE_NAME & ".ProposeReceiveDelete", "Penerimaan asal tidak boleh kosong."
    End If

    If original.ReceivedQuantity = 0 Then
        managerError.ThrowError serviceReceive.AlreadyDeleted, _
            MODULE_NAME & ".ProposeReceiveDelete", _
            "Penerimaan " & original.UUID & " sudah dihapus (jumlah 0)."
    End If

    Set proposed = New Receive

    With proposed
        .ArrivalDate = original.ArrivalDate
        .VerificationDate = original.VerificationDate
        .ReceivedQuantity = 0
        .OkQuantity = 0
        .NotOkQuantity = 0
        .Remarks = original.Remarks
        .SystemRemarks = original.SystemRemarks
        If Len(original.PersonInCharge) > 0 Then .PersonInCharge = original.PersonInCharge
    End With

    ProposeReceiveDelete = ProposeReceiveUpdate(plantName, original, proposed, reason)

End Function


' Activity log JSON of a receive (README "Log JSON").
Public Function BuildReceiveJson( _
    ByVal receiveItem As Receive, _
    ByVal plantItem As ManufacturingPlant) As String

    BuildReceiveJson = coreJson.JsonObject( _
        "uuid", coreJson.JsonText(receiveItem.UUID), _
        "transaction_type", coreJson.JsonText("RECEIVE"), _
        "order", coreJson.JsonObject( _
            "uuid", coreJson.JsonText(receiveItem.Order.UUID), _
            "po_number", coreJson.JsonText(receiveItem.Order.PoNumber)), _
        "manufacturing_plant", coreJson.JsonObject( _
            "code", coreJson.JsonText(plantItem.Code), _
            "name", coreJson.JsonText(plantItem.Name)), _
        "sparepart", serviceSparepart.BuildSparepartReferenceJson(receiveItem.Order.Sparepart), _
        "arrival_date", coreJson.JsonDate(receiveItem.ArrivalDate), _
        "verification_date", coreJson.JsonDate(receiveItem.VerificationDate), _
        "received_quantity", coreJson.JsonNumber(receiveItem.ReceivedQuantity), _
        "ok_quantity", coreJson.JsonNumber(receiveItem.OkQuantity), _
        "not_ok_quantity", coreJson.JsonNumber(receiveItem.NotOkQuantity), _
        "remarks", coreJson.JsonText(receiveItem.Remarks), _
        "system_remarks", coreJson.JsonText(receiveItem.SystemRemarks), _
        "person_in_charge", coreJson.JsonText(receiveItem.PersonInCharge))

End Function


' =========================================================
' INTERNAL
' =========================================================

Private Sub ValidateReceive( _
    ByVal receiveItem As Receive, _
    ByVal isNew As Boolean, _
    ByVal errorSource As String)

    If receiveItem Is Nothing Then
        managerError.ThrowError serviceReceive.EmptyReceive, _
            errorSource, "Penerimaan tidak boleh kosong."
    End If

    If receiveItem.Order Is Nothing Then
        managerError.ThrowError serviceReceive.EmptyOrder, _
            errorSource, "Order penerimaan belum dipilih."
    End If

    If Len(receiveItem.Order.UUID) = 0 Then
        managerError.ThrowError serviceReceive.EmptyOrder, _
            errorSource, "Nomor order penerimaan kosong."
    End If

    If receiveItem.ArrivalDate = 0 Then
        managerError.ThrowError serviceReceive.EmptyArrivalDate, _
            errorSource, _
            "Tanggal datang untuk order " & receiveItem.Order.UUID & " belum diisi."
    End If

    If isNew And receiveItem.ReceivedQuantity <= 0 Then
        managerError.ThrowError serviceReceive.ZeroReceivedQuantity, _
            errorSource, _
            "Jumlah penerimaan untuk order " & receiveItem.Order.UUID & _
            " harus lebih besar dari 0."
    End If

    If Not receiveItem.IsQuantityBalanced Then
        managerError.ThrowError serviceReceive.QuantityNotBalanced, _
            errorSource, _
            "Order " & receiveItem.Order.UUID & _
            ": jumlah penerimaan harus sama dengan OK + NOT OK."
    End If

    If Len(receiveItem.PersonInCharge) = 0 Then
        managerError.ThrowError serviceReceive.EmptyPersonInCharge, _
            errorSource, _
            "PIC penerimaan untuk order " & receiveItem.Order.UUID & " belum diisi."
    End If

End Sub


Private Function HasChanges( _
    ByVal original As Receive, _
    ByVal proposed As Receive) As Boolean

    HasChanges = _
        DateValue(original.ArrivalDate) <> DateValue(proposed.ArrivalDate) _
        Or original.VerificationDate <> proposed.VerificationDate _
        Or original.ReceivedQuantity <> proposed.ReceivedQuantity _
        Or original.OkQuantity <> proposed.OkQuantity _
        Or original.NotOkQuantity <> proposed.NotOkQuantity _
        Or original.Remarks <> proposed.Remarks _
        Or original.PersonInCharge <> proposed.PersonInCharge

End Function


Private Function BuildHeader( _
    ByVal transactionUUID As String, _
    ByVal plantItem As ManufacturingPlant, _
    ByVal createdBy As String, _
    ByVal isActive As Boolean) As Transaction

    Dim header As Transaction

    Set header = New Transaction

    header.UUID = transactionUUID
    header.TransactionType = "RECEIVE"
    Set header.ManufacturingPlant = plantItem
    header.CreatedBy = createdBy
    header.CreatedAt = Now
    If isActive Then header.ActivatedAt = Now

    Set BuildHeader = header

End Function
