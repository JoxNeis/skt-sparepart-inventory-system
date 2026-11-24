Attribute VB_Name = "serviceTransactionRequestUpdate"
Option Explicit

' =========================================================
' SERVICE - TRANSACTION REQUEST UPDATE
'
' Public entry points (called by sheets) get the
' connection and own the transaction:
'   ApproveRequestUpdate
'   SelectOrderRequestUpdate
'   SelectReceiveRequestUpdate
'   SelectRequestRequestUpdate
'
' Helper for other services (runs INSIDE the caller's
' transaction, receives conn, never commits / rolls back):
'   CreateRequestUpdate
' =========================================================

Private Const MODULE_NAME As String = "SERVICE_TRANSACTION_REQUEST_UPDATE"

' =========================================================
' ERROR CODES (18000 - 18099)
' Usage: managerError.ThrowError serviceTransactionRequestUpdate.EmptyOriginalUUID, ...
' =========================================================
Public Const EmptyOriginalUUID As Long = 18001
Public Const EmptyProposedUUID As Long = 18002
Public Const SameOriginalProposedUUID As Long = 18003
Public Const EmptyRequestedBy As Long = 18004                    ' current user is blank
Public Const NoRowAffectedCreateRequestUpdate As Long = 18005    ' request not inserted
Public Const EmptyApproverUser As Long = 18006
Public Const EmptyOldCatalogueName As Long = 18007
Public Const EmptyNewCatalogueName As Long = 18008
Public Const EmptyOldManufacturingPlantName As Long = 18009
Public Const EmptyNewManufacturingPlantName As Long = 18010
Public Const NoRowAffectedApproveRequestUpdate As Long = 18011   ' not found / processed / tier too low
Public Const MultipleRowsAffectedApproveRequestUpdate As Long = 18012
Public Const EmptyManufacturingPlantName As Long = 18013
Public Const EmptyFromDate As Long = 18014                       ' fromDate <= 0
Public Const EmptyUntilDate As Long = 18015                      ' untilDate <= 0
Public Const InvalidDateRange As Long = 18016                    ' untilDate < fromDate


' =========================================================
' HELPER FOR OTHER SERVICES (inside caller's transaction)
' =========================================================

' Register that proposedUUID is a proposed update of
' originalUUID, requested by the current user.
Public Sub CreateRequestUpdate( _
    ByVal conn As ADODB.connection, _
    ByVal originalUUID As String, _
    ByVal proposedUUID As String)

    On Error GoTo ErrorHandler

    Dim requestedBy As String
    Dim affected As Long
    Dim errInfo As ErrorInfo

    originalUUID = Trim$(originalUUID)
    proposedUUID = Trim$(proposedUUID)
    requestedBy = Trim$(coreConfig.currentUser())

    RequireUUIDPair originalUUID, proposedUUID, _
        MODULE_NAME & ".CreateRequestUpdate"

    If managerValidationText.IsBlank(requestedBy) Then
        managerError.ThrowError _
            serviceTransactionRequestUpdate.EmptyRequestedBy, _
            MODULE_NAME & ".CreateRequestUpdate", _
            "User yang meminta perubahan tidak boleh kosong."
    End If

    affected = daoTransactionRequestUpdate.CreateTransactionRequestUpdate( _
        conn, originalUUID, proposedUUID, requestedBy)

    If managerValidationNumeric.IsLessThan(affected, 1) Then
        managerError.ThrowError _
            serviceTransactionRequestUpdate.NoRowAffectedCreateRequestUpdate, _
            MODULE_NAME & ".CreateRequestUpdate", _
            "Permintaan perubahan transaksi tidak tersimpan. " & _
            "Pastikan UUID transaksi original dan proposed tersedia."
    End If

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".CreateRequestUpdate", _
        "Gagal membuat permintaan perubahan transaksi."

End Sub


' =========================================================
' PUBLIC ENTRY POINTS
' =========================================================

' Approve a proposed update in one transaction:
'   1. mark the request as approved
'   2. reverse the original transaction
'   3. activate the proposed transaction
'   4. (updateOrder) move receives to the proposed order
'   5. (updateStockMovement) recalculate the stock of the
'      old and, when different, the new catalogue / plant
'   6. write the activity log
Public Sub ApproveRequestUpdate( _
    ByVal originalTransactionUUID As String, _
    ByVal proposedTransactionUUID As String, _
    ByVal user As String, _
    Optional ByVal updateOrder As Boolean = False, _
    Optional ByVal updateStockMovement As Boolean = False, _
    Optional ByVal oldCatalogueName As String = "", _
    Optional ByVal newCatalogueName As String = "", _
    Optional ByVal oldManufacturingPlantName As String = "", _
    Optional ByVal newManufacturingPlantName As String = "")

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim affected As Long
    Dim activityCode As String
    Dim activityDetail As String
    Dim isSameStockRecord As Boolean
    Dim inTransaction As Boolean
    Dim errInfo As ErrorInfo

    originalTransactionUUID = Trim$(originalTransactionUUID)
    proposedTransactionUUID = Trim$(proposedTransactionUUID)
    user = Trim$(user)
    oldCatalogueName = Trim$(oldCatalogueName)
    newCatalogueName = Trim$(newCatalogueName)
    oldManufacturingPlantName = Trim$(oldManufacturingPlantName)
    newManufacturingPlantName = Trim$(newManufacturingPlantName)

    ' =====================================================
    ' Step 1: Validate
    ' =====================================================
    RequireUUIDPair originalTransactionUUID, proposedTransactionUUID, _
        MODULE_NAME & ".ApproveRequestUpdate"

    If managerValidationText.IsBlank(user) Then
        managerError.ThrowError _
            serviceTransactionRequestUpdate.EmptyApproverUser, _
            MODULE_NAME & ".ApproveRequestUpdate", _
            "User tidak boleh kosong."
    End If

    If updateStockMovement Then

        If managerValidationText.IsBlank(oldCatalogueName) Then
            managerError.ThrowError _
                serviceTransactionRequestUpdate.EmptyOldCatalogueName, _
                MODULE_NAME & ".ApproveRequestUpdate", _
                "Nama catalogue original tidak boleh " & _
                "kosong ketika update stock aktif."
        End If

        If managerValidationText.IsBlank(newCatalogueName) Then
            managerError.ThrowError _
                serviceTransactionRequestUpdate.EmptyNewCatalogueName, _
                MODULE_NAME & ".ApproveRequestUpdate", _
                "Nama catalogue proposed tidak boleh " & _
                "kosong ketika update stock aktif."
        End If

        If managerValidationText.IsBlank(oldManufacturingPlantName) Then
            managerError.ThrowError _
                serviceTransactionRequestUpdate.EmptyOldManufacturingPlantName, _
                MODULE_NAME & ".ApproveRequestUpdate", _
                "Nama manufacturing plant original " & _
                "tidak boleh kosong ketika update stock aktif."
        End If

        If managerValidationText.IsBlank(newManufacturingPlantName) Then
            managerError.ThrowError _
                serviceTransactionRequestUpdate.EmptyNewManufacturingPlantName, _
                MODULE_NAME & ".ApproveRequestUpdate", _
                "Nama manufacturing plant proposed " & _
                "tidak boleh kosong ketika update stock aktif."
        End If

    End If

    ' =====================================================
    ' Step 2: Activity log content
    ' =====================================================
    If updateStockMovement Then
        activityCode = "Receive.UpdateRequest.Approve"
    ElseIf updateOrder Then
        activityCode = "Order.UpdateRequest.Approve"
    Else
        activityCode = "Transaction.UpdateRequest.Approve"
    End If

    activityDetail = BuildApproveDetail( _
        originalTransactionUUID, proposedTransactionUUID, user, _
        updateOrder, updateStockMovement, _
        oldCatalogueName, newCatalogueName, _
        oldManufacturingPlantName, newManufacturingPlantName)

    ' =====================================================
    ' Step 3: Begin transaction
    ' =====================================================
    Set conn = managerDatabase.GetConnection()

    managerDatabase.BeginTransaction conn
    inTransaction = True

    ' =====================================================
    ' Step 4: Approve the request
    ' =====================================================
    affected = daoTransactionRequestUpdate.ApproveRequestUpdate( _
        conn, originalTransactionUUID, proposedTransactionUUID, user)

    If managerValidationNumeric.IsLessThan(affected, 1) Then
        managerError.ThrowError _
            serviceTransactionRequestUpdate.NoRowAffectedApproveRequestUpdate, _
            MODULE_NAME & ".ApproveRequestUpdate", _
            "Pengajuan update tidak disetujui. " & _
            "Data tidak ditemukan, sudah diproses, " & _
            "atau tier user tidak memenuhi persyaratan."
    End If

    If managerValidationNumeric.IsGreaterThan(affected, 1) Then
        managerError.ThrowError _
            serviceTransactionRequestUpdate.MultipleRowsAffectedApproveRequestUpdate, _
            MODULE_NAME & ".ApproveRequestUpdate", _
            "Lebih dari satu pengajuan update " & _
            "terpengaruh. Transaksi dibatalkan untuk " & _
            "mencegah perubahan data yang tidak sesuai."
    End If

    ' =====================================================
    ' Step 5: Reverse original, activate proposed
    ' =====================================================
    serviceTransaction.ReverseTransaction conn, originalTransactionUUID

    serviceTransaction.ActivateTransaction conn, proposedTransactionUUID

    ' =====================================================
    ' Step 6: Move receives to the proposed order
    ' (0 rows is valid: the order has no receive yet)
    ' =====================================================
    If updateOrder Then
        affected = daoUpdateOrder.UpdateOrderReceives( _
            conn, originalTransactionUUID, proposedTransactionUUID)
    End If

    ' =====================================================
    ' Step 7: Recalculate stock
    ' Runs after reversal + activation, on the same
    ' connection, so a failure rolls everything back.
    ' =====================================================
    If updateStockMovement Then

        isSameStockRecord = _
            managerValidationText.IsSameText(oldCatalogueName, newCatalogueName) And _
            managerValidationText.IsSameText(oldManufacturingPlantName, newManufacturingPlantName)

        servicePlantStock.RecalculateStock _
            conn, oldCatalogueName, oldManufacturingPlantName

        If Not isSameStockRecord Then
            servicePlantStock.RecalculateStock _
                conn, newCatalogueName, newManufacturingPlantName
        End If

    End If

    ' =====================================================
    ' Step 8: Activity log (approving user)
    ' =====================================================
    serviceActivityLog.WriteActivityLog _
        conn, activityCode, activityDetail, user

    ' =====================================================
    ' Step 9: Commit
    ' =====================================================
    managerDatabase.CommitTransaction conn
    inTransaction = False

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    If inTransaction Then
        managerDatabase.RollbackTransaction conn
    End If

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ApproveRequestUpdate", _
        "Gagal menyetujui pengajuan update."

End Sub


' Pending order update requests. Disconnected recordset.
Public Function SelectOrderRequestUpdate( _
    ByVal manufacturingPlantName As String, _
    ByVal fromDate As Date, _
    ByVal untilDate As Date) As ADODB.recordset

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim errInfo As ErrorInfo

    manufacturingPlantName = Trim$(manufacturingPlantName)

    RequireSearchParameters manufacturingPlantName, fromDate, untilDate, _
        MODULE_NAME & ".SelectOrderRequestUpdate"

    Set conn = managerDatabase.GetConnection()

    Set SelectOrderRequestUpdate = daoUpdateOrder.SelectOrderRequestUpdate( _
        conn, manufacturingPlantName, DateValue(fromDate), DateValue(untilDate))

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".SelectOrderRequestUpdate", _
        "Gagal mengambil data pengajuan update order."

End Function


' Pending receive update requests. Disconnected recordset.
Public Function SelectReceiveRequestUpdate( _
    ByVal manufacturingPlantName As String, _
    ByVal fromDate As Date, _
    ByVal untilDate As Date) As ADODB.recordset

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim errInfo As ErrorInfo

    manufacturingPlantName = Trim$(manufacturingPlantName)

    RequireSearchParameters manufacturingPlantName, fromDate, untilDate, _
        MODULE_NAME & ".SelectReceiveRequestUpdate"

    Set conn = managerDatabase.GetConnection()

    Set SelectReceiveRequestUpdate = daoUpdateReceive.SelectReceiveRequestUpdate( _
        conn, manufacturingPlantName, DateValue(fromDate), DateValue(untilDate))

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".SelectReceiveRequestUpdate", _
        "Gagal mengambil pengajuan update penerimaan."

End Function


' Pending request (warehouse request) update requests.
' Disconnected recordset.
Public Function SelectRequestRequestUpdate( _
    ByVal manufacturingPlantName As String, _
    ByVal fromDate As Date, _
    ByVal untilDate As Date) As ADODB.recordset

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim errInfo As ErrorInfo

    manufacturingPlantName = Trim$(manufacturingPlantName)

    RequireSearchParameters manufacturingPlantName, fromDate, untilDate, _
        MODULE_NAME & ".SelectRequestRequestUpdate"

    Set conn = managerDatabase.GetConnection()

    Set SelectRequestRequestUpdate = daoUpdateRequest.SelectRequestRequestUpdate( _
        conn, manufacturingPlantName, DateValue(fromDate), DateValue(untilDate))

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".SelectRequestRequestUpdate", _
        "Gagal mengambil pengajuan update request."

End Function


' =========================================================
' INTERNAL
' =========================================================

Private Sub RequireUUIDPair( _
    ByVal originalUUID As String, _
    ByVal proposedUUID As String, _
    ByVal errorSource As String)

    If managerValidationText.IsBlank(originalUUID) Then
        managerError.ThrowError _
            serviceTransactionRequestUpdate.EmptyOriginalUUID, _
            errorSource, _
            "UUID transaksi original tidak boleh kosong."
    End If

    If managerValidationText.IsBlank(proposedUUID) Then
        managerError.ThrowError _
            serviceTransactionRequestUpdate.EmptyProposedUUID, _
            errorSource, _
            "UUID transaksi proposed tidak boleh kosong."
    End If

    If managerValidationText.IsSameText(originalUUID, proposedUUID) Then
        managerError.ThrowError _
            serviceTransactionRequestUpdate.SameOriginalProposedUUID, _
            errorSource, _
            "UUID transaksi original dan proposed tidak boleh sama."
    End If

End Sub


Private Sub RequireSearchParameters( _
    ByVal manufacturingPlantName As String, _
    ByVal fromDate As Date, _
    ByVal untilDate As Date, _
    ByVal errorSource As String)

    If managerValidationText.IsBlank(manufacturingPlantName) Then
        managerError.ThrowError _
            serviceTransactionRequestUpdate.EmptyManufacturingPlantName, _
            errorSource, _
            "Nama manufacturing plant tidak boleh kosong."
    End If

    If managerValidationNumeric.IsZeroOrNegative(fromDate) Then
        managerError.ThrowError _
            serviceTransactionRequestUpdate.EmptyFromDate, _
            errorSource, _
            "Tanggal awal tidak boleh kosong."
    End If

    If managerValidationNumeric.IsZeroOrNegative(untilDate) Then
        managerError.ThrowError _
            serviceTransactionRequestUpdate.EmptyUntilDate, _
            errorSource, _
            "Tanggal akhir tidak boleh kosong."
    End If

    If DateValue(untilDate) < DateValue(fromDate) Then
        managerError.ThrowError _
            serviceTransactionRequestUpdate.InvalidDateRange, _
            errorSource, _
            "Tanggal akhir tidak boleh lebih kecil dari tanggal awal."
    End If

End Sub


Private Function BuildApproveDetail( _
    ByVal originalTransactionUUID As String, _
    ByVal proposedTransactionUUID As String, _
    ByVal user As String, _
    ByVal updateOrder As Boolean, _
    ByVal updateStockMovement As Boolean, _
    ByVal oldCatalogueName As String, _
    ByVal newCatalogueName As String, _
    ByVal oldManufacturingPlantName As String, _
    ByVal newManufacturingPlantName As String) As String

    Dim detail As String

    detail = _
        "Original Transaction UUID: " & originalTransactionUUID & vbCrLf & _
        "Proposed Transaction UUID: " & proposedTransactionUUID & vbCrLf & _
        "Approved By: " & user & vbCrLf & _
        "Original Transaction Reversed: Yes" & vbCrLf & _
        "Proposed Transaction Activated: Yes"

    If updateOrder Then
        detail = detail & "; update_order=true"
    End If

    If updateStockMovement Then
        detail = detail & "; update_stock_movement=true"
        detail = detail & "; old_catalogue=" & oldCatalogueName
        detail = detail & "; new_catalogue=" & newCatalogueName
        detail = detail & "; old_manufacturing_plant=" & oldManufacturingPlantName
        detail = detail & "; new_manufacturing_plant=" & newManufacturingPlantName
    End If

    BuildApproveDetail = detail

End Function
