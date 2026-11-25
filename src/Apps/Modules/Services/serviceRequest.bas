Attribute VB_Name = "serviceRequest"
Option Explicit

' =========================================================
' SERVICE - REQUEST (pengeluaran)
'
' Spareparts issued from the plant store to a mobile
' maintenance of the same plant. A request reduces stock on
' its request date.
'
' Rule: the stock card may not go negative on the request
' date or ANY later day (stock_ledger_get_available), which
' also protects back-dated requests.
'
' Public entry points get the connection and own the
' transaction:
'   CreateRequests, ListRequests,
'   ProposeRequestUpdate, ProposeRequestDelete
' =========================================================

Private Const MODULE_NAME As String = "serviceRequest"

' =========================================================
' ERROR CODES (17700 - 17799)
' Usage: managerError.ThrowError serviceRequest.InsufficientStock, ...
' =========================================================
Public Const EmptyRequestList As Long = 17701
Public Const EmptyRequest As Long = 17702
Public Const EmptySparepart As Long = 17703
Public Const EmptyMobileMaintenance As Long = 17704
Public Const EmptyRequestDate As Long = 17705
Public Const ZeroQuantity As Long = 17706
Public Const EmptyPersonInCharge As Long = 17707
Public Const InsufficientStock As Long = 17708
Public Const CreateFailed As Long = 17709          ' other plant / archived
Public Const InvalidDateRange As Long = 17710
Public Const EmptyReason As Long = 17711
Public Const NothingChanged As Long = 17712
Public Const AlreadyDeleted As Long = 17713


' =========================================================
' CREATE
' =========================================================

' Create every request of the list as an active document, in
' one transaction. Each Request needs Sparepart (with Id),
' MobileMaintenance (with Id, same plant), RequestDate,
' Quantity > 0 and PersonInCharge. Returns the numbers.
Public Function CreateRequests( _
    ByVal plantName As String, _
    ByVal requests As Collection) As Collection

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim plantItem As ManufacturingPlant
    Dim creator As User
    Dim requestItem As Request
    Dim uuids As Collection
    Dim touched As Object
    Dim sparepartKey As Variant
    Dim transactionUUID As String
    Dim available As Long
    Dim affected As Long
    Dim inTransaction As Boolean
    Dim errInfo As ErrorInfo

    If requests Is Nothing Then
        managerError.ThrowError serviceRequest.EmptyRequestList, _
            MODULE_NAME & ".CreateRequests", "Tidak ada request untuk disimpan."
    End If

    If requests.Count = 0 Then
        managerError.ThrowError serviceRequest.EmptyRequestList, _
            MODULE_NAME & ".CreateRequests", "Tidak ada request untuk disimpan."
    End If

    For Each requestItem In requests
        ValidateRequest requestItem, True, MODULE_NAME & ".CreateRequests"
    Next requestItem

    Set conn = managerDatabase.GetConnection()
    Set plantItem = serviceManufacturingPlant.FindPlant(conn, plantName)
    Set creator = serviceUser.RequireRegisteredUser(conn)

    Set uuids = New Collection
    Set touched = CreateObject("Scripting.Dictionary")

    managerDatabase.BeginTransaction conn
    inTransaction = True

    For Each requestItem In requests

        ' Same connection: the rows saved earlier in this batch
        ' already count.
        available = serviceStockLedger.GetAvailableQuantity( _
            conn, plantItem.Id, requestItem.Sparepart.Id, requestItem.RequestDate)

        If requestItem.Quantity > available Then
            RaiseInsufficientStock requestItem, available, _
                MODULE_NAME & ".CreateRequests"
        End If

        transactionUUID = serviceTransaction.CreateTransaction( _
            conn, plantItem, "REQUEST", requestItem.RequestDate, True, creator.Username)

        affected = daoRequest.CreateRequest(conn, transactionUUID, requestItem)

        managerDatabase.RequireAffectedRows affected, 1, _
            serviceRequest.CreateFailed, _
            MODULE_NAME & ".CreateRequests", _
            "Request " & requestItem.Sparepart.DisplayName & _
            " tidak tersimpan: pos mobile bukan milik plant ini, " & _
            "atau sparepart / pos mobile sudah diarsipkan."

        Set requestItem.Transaction = BuildHeader( _
            transactionUUID, plantItem, creator.Username, True)

        touched(CStr(requestItem.Sparepart.Id)) = requestItem.Sparepart.DisplayName

        serviceActivityLog.WriteActivityLog conn, _
            "Request.Create", _
            "Request " & transactionUUID & ": " & requestItem.Sparepart.DisplayName & _
            " x " & requestItem.Quantity & " ke " & requestItem.MobileMaintenance.Name & ".", _
            BuildRequestJson(requestItem, plantItem)

        uuids.Add transactionUUID

    Next requestItem

    For Each sparepartKey In touched.Keys
        servicePlantStock.RequireNonNegativeLedger _
            conn, plantItem.Id, CLng(sparepartKey), CStr(touched(sparepartKey))
        servicePlantStock.RecalculateCurrentStock conn, plantItem.Id, CLng(sparepartKey)
    Next sparepartKey

    managerDatabase.CommitTransaction conn
    inTransaction = False

    Set CreateRequests = uuids

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()
    If inTransaction Then managerDatabase.RollbackTransaction conn

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".CreateRequests", _
        "Gagal menyimpan request. Tidak ada request yang tersimpan."

End Function


' =========================================================
' LIST
' =========================================================

' Active requests of the plant with a request date in the
' range, keyed by document number.
Public Function ListRequests( _
    ByVal plantName As String, _
    ByVal fromDate As Date, _
    ByVal toDate As Date, _
    Optional ByVal page As PageRequest = Nothing) As Collection

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim errInfo As ErrorInfo

    If DateValue(toDate) < DateValue(fromDate) Then
        managerError.ThrowError serviceRequest.InvalidDateRange, _
            MODULE_NAME & ".ListRequests", _
            "Tanggal akhir tidak boleh lebih kecil dari tanggal awal."
    End If

    Set conn = managerDatabase.GetConnection()

    Set ListRequests = daoRequest.ListRequests( _
        conn, serviceManufacturingPlant.FindPlant(conn, plantName), _
        fromDate, toDate, page)

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ListRequests", _
        "Gagal mengambil data request."

End Function

' Number of rows ListRequests gives over all pages (same filter).
Public Function CountRequests( _
    ByVal plantName As String, _
    ByVal fromDate As Date, _
    ByVal toDate As Date, _
    Optional ByVal page As PageRequest = Nothing) As Long

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim errInfo As ErrorInfo

    Set conn = managerDatabase.GetConnection()

    CountRequests = daoRequest.CountRequests( _
        conn, serviceManufacturingPlant.FindPlant(conn, plantName), _
        fromDate, toDate, page)

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".CountRequests", _
        "Gagal menghitung data request."

End Function



' =========================================================
' UPDATE PROPOSAL
' =========================================================

' Save the corrected request as a DRAFT and create a pending
' update request for the original. The proposal is checked
' against the stock card; the original quantity is added
' back only when the original is dated on or before the
' proposed date (and it is the same sparepart).
Public Function ProposeRequestUpdate( _
    ByVal plantName As String, _
    ByVal original As Request, _
    ByVal proposed As Request, _
    ByVal reason As String) As String

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim plantItem As ManufacturingPlant
    Dim requester As User
    Dim proposedUUID As String
    Dim available As Long
    Dim affected As Long
    Dim inTransaction As Boolean
    Dim errInfo As ErrorInfo

    If original Is Nothing Then
        managerError.ThrowError serviceRequest.EmptyRequest, _
            MODULE_NAME & ".ProposeRequestUpdate", "Request asal tidak boleh kosong."
    End If

    ValidateRequest proposed, False, MODULE_NAME & ".ProposeRequestUpdate"

    If managerValidationText.IsBlank(reason) Then
        managerError.ThrowError serviceRequest.EmptyReason, _
            MODULE_NAME & ".ProposeRequestUpdate", "Alasan perubahan wajib diisi."
    End If

    If Not HasChanges(original, proposed) Then
        managerError.ThrowError serviceRequest.NothingChanged, _
            MODULE_NAME & ".ProposeRequestUpdate", _
            "Tidak ada perubahan pada request " & original.UUID & "."
    End If

    Set conn = managerDatabase.GetConnection()
    Set plantItem = serviceManufacturingPlant.FindPlant(conn, plantName)
    Set requester = serviceUser.RequireRegisteredUser(conn)

    If proposed.Quantity > 0 Then

        available = serviceStockLedger.GetAvailableQuantity( _
            conn, plantItem.Id, proposed.Sparepart.Id, proposed.RequestDate)

        If original.Sparepart.Id = proposed.Sparepart.Id Then
            If DateValue(original.RequestDate) <= DateValue(proposed.RequestDate) Then
                available = available + original.Quantity
            End If
        End If

        If proposed.Quantity > available Then
            RaiseInsufficientStock proposed, available, _
                MODULE_NAME & ".ProposeRequestUpdate"
        End If

    End If

    managerDatabase.BeginTransaction conn
    inTransaction = True

    proposedUUID = serviceTransaction.CreateTransaction( _
        conn, plantItem, "REQUEST", proposed.RequestDate, False, requester.Username)

    affected = daoRequest.CreateRequest(conn, proposedUUID, proposed)

    managerDatabase.RequireAffectedRows affected, 1, _
        serviceRequest.CreateFailed, _
        MODULE_NAME & ".ProposeRequestUpdate", _
        "Request usulan tidak tersimpan: pos mobile bukan milik plant ini, " & _
        "atau sparepart / pos mobile sudah diarsipkan."

    Set proposed.Transaction = BuildHeader( _
        proposedUUID, plantItem, requester.Username, False)

    serviceTransactionUpdateRequest.CreateUpdateRequest _
        conn, original.UUID, proposedUUID, reason, requester.Username

    serviceActivityLog.WriteActivityLog conn, _
        "Request.CreateUpdateProposal", _
        "Usulan update request " & original.UUID & " -> " & proposedUUID & _
        ". Alasan: " & reason, _
        coreJson.JsonAppend( _
            coreJson.JsonAppend(BuildRequestJson(proposed, plantItem), _
                "original_uuid", coreJson.JsonText(original.UUID)), _
            "reason", coreJson.JsonText(reason))

    managerDatabase.CommitTransaction conn
    inTransaction = False

    ProposeRequestUpdate = proposedUUID

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()
    If inTransaction Then managerDatabase.RollbackTransaction conn

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ProposeRequestUpdate", _
        "Gagal membuat usulan update request."

End Function


' Propose to delete a request: an update with quantity 0.
Public Function ProposeRequestDelete( _
    ByVal plantName As String, _
    ByVal original As Request, _
    ByVal reason As String) As String

    Dim proposed As Request

    If original Is Nothing Then
        managerError.ThrowError serviceRequest.EmptyRequest, _
            MODULE_NAME & ".ProposeRequestDelete", "Request asal tidak boleh kosong."
    End If

    If original.Quantity = 0 Then
        managerError.ThrowError serviceRequest.AlreadyDeleted, _
            MODULE_NAME & ".ProposeRequestDelete", _
            "Request " & original.UUID & " sudah dihapus (jumlah 0)."
    End If

    Set proposed = CopyRequest(original)
    proposed.Quantity = 0

    ProposeRequestDelete = ProposeRequestUpdate(plantName, original, proposed, reason)

End Function


' A new Request with the same detail values as source (no
' document header).
Public Function CopyRequest( _
    ByVal source As Request) As Request

    Dim result As Request

    Set result = New Request

    With result
        Set .MobileMaintenance = source.MobileMaintenance
        Set .Sparepart = source.Sparepart
        .RequestDate = source.RequestDate
        .Quantity = source.Quantity
        .Remarks = source.Remarks
        .SystemRemarks = source.SystemRemarks
        If Len(source.PersonInCharge) > 0 Then .PersonInCharge = source.PersonInCharge
    End With

    Set CopyRequest = result

End Function


' Activity log JSON of a request (README "Log JSON").
Public Function BuildRequestJson( _
    ByVal requestItem As Request, _
    ByVal plantItem As ManufacturingPlant) As String

    BuildRequestJson = coreJson.JsonObject( _
        "uuid", coreJson.JsonText(requestItem.UUID), _
        "transaction_type", coreJson.JsonText("REQUEST"), _
        "manufacturing_plant", coreJson.JsonObject( _
            "code", coreJson.JsonText(plantItem.Code), _
            "name", coreJson.JsonText(plantItem.Name)), _
        "mobile_maintenance", serviceMobileMaintenance.BuildReferenceJson(requestItem.MobileMaintenance), _
        "sparepart", serviceSparepart.BuildSparepartReferenceJson(requestItem.Sparepart), _
        "request_date", coreJson.JsonDate(requestItem.RequestDate), _
        "quantity", coreJson.JsonNumber(requestItem.Quantity), _
        "remarks", coreJson.JsonText(requestItem.Remarks), _
        "system_remarks", coreJson.JsonText(requestItem.SystemRemarks), _
        "person_in_charge", coreJson.JsonText(requestItem.PersonInCharge))

End Function


' =========================================================
' INTERNAL
' =========================================================

Private Sub ValidateRequest( _
    ByVal requestItem As Request, _
    ByVal isNew As Boolean, _
    ByVal errorSource As String)

    If requestItem Is Nothing Then
        managerError.ThrowError serviceRequest.EmptyRequest, _
            errorSource, "Request tidak boleh kosong."
    End If

    If requestItem.Sparepart Is Nothing Then
        managerError.ThrowError serviceRequest.EmptySparepart, _
            errorSource, "Sparepart request belum dipilih."
    End If

    If requestItem.Sparepart.Id <= 0 Then
        managerError.ThrowError serviceRequest.EmptySparepart, _
            errorSource, "Sparepart request tidak dikenal."
    End If

    If requestItem.MobileMaintenance Is Nothing Then
        managerError.ThrowError serviceRequest.EmptyMobileMaintenance, _
            errorSource, _
            "Pos mobile untuk " & requestItem.Sparepart.DisplayName & " belum dipilih."
    End If

    If requestItem.MobileMaintenance.Id <= 0 Then
        managerError.ThrowError serviceRequest.EmptyMobileMaintenance, _
            errorSource, _
            "Pos mobile untuk " & requestItem.Sparepart.DisplayName & " tidak dikenal."
    End If

    If requestItem.RequestDate = 0 Then
        managerError.ThrowError serviceRequest.EmptyRequestDate, _
            errorSource, "Tanggal request belum diisi."
    End If

    If isNew And requestItem.Quantity <= 0 Then
        managerError.ThrowError serviceRequest.ZeroQuantity, _
            errorSource, _
            "Jumlah request " & requestItem.Sparepart.DisplayName & _
            " harus lebih besar dari 0."
    End If

    If Len(requestItem.PersonInCharge) = 0 Then
        managerError.ThrowError serviceRequest.EmptyPersonInCharge, _
            errorSource, _
            "PIC request " & requestItem.Sparepart.DisplayName & " belum diisi."
    End If

End Sub


Private Sub RaiseInsufficientStock( _
    ByVal requestItem As Request, _
    ByVal available As Long, _
    ByVal errorSource As String)

    If available < 0 Then available = 0

    managerError.ThrowError serviceRequest.InsufficientStock, _
        errorSource, _
        "Stok " & requestItem.Sparepart.DisplayName & " tidak cukup pada " & _
        Format$(requestItem.RequestDate, "dd-mmm-yyyy") & _
        " (atau hari setelahnya): tersedia " & available & _
        ", diminta " & requestItem.Quantity & "."

End Sub


Private Function HasChanges( _
    ByVal original As Request, _
    ByVal proposed As Request) As Boolean

    HasChanges = _
        original.Sparepart.Id <> proposed.Sparepart.Id _
        Or original.MobileMaintenance.Id <> proposed.MobileMaintenance.Id _
        Or DateValue(original.RequestDate) <> DateValue(proposed.RequestDate) _
        Or original.Quantity <> proposed.Quantity _
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
    header.TransactionType = "REQUEST"
    Set header.ManufacturingPlant = plantItem
    header.CreatedBy = createdBy
    header.CreatedAt = Now
    If isActive Then header.ActivatedAt = Now

    Set BuildHeader = header

End Function
