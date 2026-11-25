Attribute VB_Name = "serviceTransactionUpdateRequest"
Option Explicit

' =========================================================
' SERVICE - TRANSACTION UPDATE REQUEST
'
' An active document is never changed in place. A
' correction is a DRAFT document plus a pending update
' request (created by serviceOrder / serviceReceive /
' serviceRequest through CreateUpdateRequest).
'
' A reviewer with a HIGHER tier than the requester decides:
'
'   APPROVE (one database transaction):
'     verdict APPROVED, original reversed, proposal activated
'     ORDER   -> every receive of the original moves to the
'                new order (receive_reassign_order)
'     all     -> the stock card of the touched spareparts may
'                not go negative; current_stock recalculated
'
'   REJECT:
'     verdict REJECTED, proposal reversed (never activated)
' =========================================================

Private Const MODULE_NAME As String = "serviceTransactionUpdateRequest"

Private Const FIRST_REQUEST_DATE As Date = #1/1/2000#

' =========================================================
' ERROR CODES (18000 - 18099)
' Usage: managerError.ThrowError serviceTransactionUpdateRequest.TierTooLow, ...
' =========================================================
Public Const CreateFailed As Long = 18001          ' original not active / pending exists
Public Const EmptyUpdateRequest As Long = 18002
Public Const NotPending As Long = 18003
Public Const ReviewerIsRequester As Long = 18004
Public Const TierTooLow As Long = 18005
Public Const VerdictFailed As Long = 18006


' Called INSIDE the proposing service's transaction.
Public Sub CreateUpdateRequest( _
    ByVal conn As ADODB.connection, _
    ByVal originalUUID As String, _
    ByVal proposedUUID As String, _
    ByVal reason As String, _
    ByVal requestedBy As String)

    Dim affected As Long

    affected = daoTransactionUpdateRequest.CreateUpdateRequest( _
        conn, originalUUID, proposedUUID, reason, requestedBy)

    managerDatabase.RequireAffectedRows affected, 1, _
        serviceTransactionUpdateRequest.CreateFailed, _
        MODULE_NAME & ".CreateUpdateRequest", _
        "Dokumen " & originalUUID & " tidak dapat diajukan untuk update: " & _
        "dokumen tidak aktif, atau masih ada pengajuan update yang " & _
        "menunggu persetujuan."

End Sub


' Pending update requests of the plant, oldest first.
Public Function ListPendingUpdateRequests( _
    ByVal plantName As String) As Collection

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim plantItem As ManufacturingPlant
    Dim items As Collection
    Dim pending As Collection
    Dim item As TransactionUpdateRequest
    Dim errInfo As ErrorInfo

    Set conn = managerDatabase.GetConnection()
    Set plantItem = serviceManufacturingPlant.FindPlant(conn, plantName)

    Set items = daoTransactionUpdateRequest.ListUpdateRequests( _
        conn, plantItem.Id, FIRST_REQUEST_DATE, Date)

    Set pending = New Collection

    For Each item In items
        If item.IsPending Then pending.Add item, CStr(item.Id)
    Next item

    Set ListPendingUpdateRequests = pending

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ListPendingUpdateRequests", _
        "Gagal mengambil data request update."

End Function


' Approve: reverse the original, activate the proposal and
' run the follow-up of its document type.
Public Sub ApproveUpdateRequest( _
    ByVal plantName As String, _
    ByVal item As TransactionUpdateRequest, _
    ByVal decisionNote As String)

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim plantItem As ManufacturingPlant
    Dim reviewer As User
    Dim affected As Long
    Dim receivesMoved As Long
    Dim sparepartIds As Object
    Dim sparepartKey As Variant
    Dim inTransaction As Boolean
    Dim errInfo As ErrorInfo

    Set conn = managerDatabase.GetConnection()
    Set plantItem = serviceManufacturingPlant.FindPlant(conn, plantName)
    Set reviewer = RequireReviewer(conn, item, MODULE_NAME & ".ApproveUpdateRequest")

    managerDatabase.BeginTransaction conn
    inTransaction = True

    affected = daoTransactionUpdateRequest.SetVerdict( _
        conn, item.Id, "APPROVED", reviewer.Username, decisionNote)

    managerDatabase.RequireAffectedRows affected, 1, _
        serviceTransactionUpdateRequest.VerdictFailed, _
        MODULE_NAME & ".ApproveUpdateRequest", _
        "Request update sudah diproses atau tier Anda tidak cukup."

    serviceTransaction.ReverseTransaction conn, item.OriginalUUID, reviewer.Username
    serviceTransaction.ActivateTransaction conn, item.ProposedUUID

    If item.TransactionType = "ORDER" Then
        receivesMoved = daoReceive.ReassignOrder( _
            conn, item.OriginalUUID, item.ProposedUUID)
    End If

    ' Spareparts whose stock card may have changed.
    Set sparepartIds = CreateObject("Scripting.Dictionary")

    If item.OriginalSparepartId > 0 Then
        sparepartIds(CStr(item.OriginalSparepartId)) = _
            coreHelper.TextOrEmpty(item.OriginalValue("sparepart_name"))
    End If

    If item.ProposedSparepartId > 0 Then
        sparepartIds(CStr(item.ProposedSparepartId)) = _
            coreHelper.TextOrEmpty(item.ProposedValue("sparepart_name"))
    End If

    For Each sparepartKey In sparepartIds.Keys
        servicePlantStock.RequireNonNegativeLedger _
            conn, plantItem.Id, CLng(sparepartKey), CStr(sparepartIds(sparepartKey))
        servicePlantStock.RecalculateCurrentStock conn, plantItem.Id, CLng(sparepartKey)
    Next sparepartKey

    serviceActivityLog.WriteActivityLog conn, _
        ActivityPrefix(item.TransactionType) & ".UpdateRequest.Approve", _
        "Update " & item.OriginalUUID & " -> " & item.ProposedUUID & _
        " disetujui oleh " & reviewer.Username & "." & _
        IIf(item.TransactionType = "ORDER", _
            " Penerimaan dipindahkan: " & receivesMoved & ".", ""), _
        BuildDecisionJson(item, "APPROVED", reviewer.Username, decisionNote, receivesMoved)

    managerDatabase.CommitTransaction conn
    inTransaction = False

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    If inTransaction Then managerDatabase.RollbackTransaction conn

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ApproveUpdateRequest", _
        "Gagal menyetujui request update. Tidak ada perubahan yang tersimpan."

End Sub


' Reject: the proposal is reversed, the original stays active.
Public Sub RejectUpdateRequest( _
    ByVal plantName As String, _
    ByVal item As TransactionUpdateRequest, _
    ByVal decisionNote As String)

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim reviewer As User
    Dim affected As Long
    Dim inTransaction As Boolean
    Dim errInfo As ErrorInfo

    Set conn = managerDatabase.GetConnection()

    ' Resolves (and so checks) the plant of the sheet.
    serviceManufacturingPlant.FindPlant conn, plantName

    Set reviewer = RequireReviewer(conn, item, MODULE_NAME & ".RejectUpdateRequest")

    managerDatabase.BeginTransaction conn
    inTransaction = True

    affected = daoTransactionUpdateRequest.SetVerdict( _
        conn, item.Id, "REJECTED", reviewer.Username, decisionNote)

    managerDatabase.RequireAffectedRows affected, 1, _
        serviceTransactionUpdateRequest.VerdictFailed, _
        MODULE_NAME & ".RejectUpdateRequest", _
        "Request update sudah diproses atau tier Anda tidak cukup."

    serviceTransaction.ReverseTransaction conn, item.ProposedUUID, reviewer.Username

    serviceActivityLog.WriteActivityLog conn, _
        ActivityPrefix(item.TransactionType) & ".UpdateRequest.Reject", _
        "Update " & item.OriginalUUID & " -> " & item.ProposedUUID & _
        " ditolak oleh " & reviewer.Username & ".", _
        BuildDecisionJson(item, "REJECTED", reviewer.Username, decisionNote, 0)

    managerDatabase.CommitTransaction conn
    inTransaction = False

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    If inTransaction Then managerDatabase.RollbackTransaction conn

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".RejectUpdateRequest", _
        "Gagal menolak request update."

End Sub


' =========================================================
' INTERNAL
' =========================================================

' The registered current user, when allowed to review item.
Private Function RequireReviewer( _
    ByVal conn As ADODB.connection, _
    ByVal item As TransactionUpdateRequest, _
    ByVal errorSource As String) As User

    Dim reviewer As User

    If item Is Nothing Then
        managerError.ThrowError serviceTransactionUpdateRequest.EmptyUpdateRequest, _
            errorSource, "Pilih request update terlebih dahulu."
    End If

    If Not item.IsPending Then
        managerError.ThrowError serviceTransactionUpdateRequest.NotPending, _
            errorSource, "Request update ini sudah diproses."
    End If

    Set reviewer = serviceUser.RequireRegisteredUser(conn)

    If managerValidationText.IsSameText(reviewer.Username, item.RequestedBy) Then
        managerError.ThrowError serviceTransactionUpdateRequest.ReviewerIsRequester, _
            errorSource, "Anda tidak dapat memproses request update Anda sendiri."
    End If

    If Not daoUser.CheckCanApprove(conn, item.RequestedBy, reviewer.Username) Then
        managerError.ThrowError serviceTransactionUpdateRequest.TierTooLow, _
            errorSource, _
            "Tier Anda (" & reviewer.Tier & ") harus lebih tinggi dari tier " & _
            "pengaju (" & item.RequestedBy & ")."
    End If

    Set RequireReviewer = reviewer

End Function


' ORDER -> Order, RECEIVE -> Receive, REQUEST -> Request
Private Function ActivityPrefix( _
    ByVal transactionType As String) As String

    ActivityPrefix = UCase$(Left$(transactionType, 1)) & _
        LCase$(Mid$(transactionType, 2))

End Function


Private Function BuildDecisionJson( _
    ByVal item As TransactionUpdateRequest, _
    ByVal verdict As String, _
    ByVal reviewer As String, _
    ByVal decisionNote As String, _
    ByVal receivesMoved As Long) As String

    Dim json As String

    json = coreJson.JsonObject( _
        "update_request_id", coreJson.JsonNumber(item.Id), _
        "transaction_type", coreJson.JsonText(item.TransactionType), _
        "original_uuid", coreJson.JsonText(item.OriginalUUID), _
        "proposed_uuid", coreJson.JsonText(item.ProposedUUID), _
        "requested_by", coreJson.JsonText(item.RequestedBy), _
        "reason", coreJson.JsonText(item.Reason), _
        "verdict", coreJson.JsonText(verdict), _
        "reviewer", coreJson.JsonText(reviewer), _
        "decision_note", coreJson.JsonText(decisionNote))

    If item.TransactionType = "ORDER" And verdict = "APPROVED" Then
        json = coreJson.JsonAppend(json, "receives_moved", coreJson.JsonNumber(receivesMoved))
    End If

    BuildDecisionJson = json

End Function
