Attribute VB_Name = "serviceReceive"
Option Explicit

' =========================================================
' SERVICE - RECEIVE
'
' Public entry points (called by sheets) get the
' connection and own the transaction:
'   CreateReceive, SelectReceives, SelectPendingOrderReceive
' =========================================================

Private Const MODULE_NAME As String = "SERVICE_RECEIVE"

' =========================================================
' ERROR CODES (17600 - 17699)
' Usage: managerError.ThrowError serviceReceive.EmptyReceive, ...
' =========================================================
Public Const EmptyReceive As Long = 17601                       ' Receive object is Nothing
Public Const EmptyOrderTransactionUUID As Long = 17602
Public Const InvalidTotalReceiveAmount As Long = 17603          ' total <> OK + not OK
Public Const EmptyTransactionUUID As Long = 17604               ' CreateTransaction returned no UUID
Public Const SameOriginalProposedUUID As Long = 17605           ' original UUID = proposed UUID
Public Const NoRowAffectedCreateReceive As Long = 17606         ' create_receive inserted 0 rows
Public Const MultipleRowsAffectedCreateReceive As Long = 17607  ' create_receive inserted > 1 row
Public Const EmptyManufacturingPlantName As Long = 17608
Public Const EmptyFromDate As Long = 17609                      ' fromDate <= 0
Public Const EmptyUntilDate As Long = 17610                     ' untilDate <= 0
Public Const InvalidDateRange As Long = 17611                   ' untilDate < fromDate


' =========================================================
' PUBLIC ENTRY POINTS
' =========================================================

' Create a receive in one transaction.
'
' originalUUID = ""  : normal receive (active transaction,
'                      current stock + OK amount).
' originalUUID <> "" : proposed update of that receive
'                      (inactive transaction + update
'                      request, stock is not changed).
Public Sub CreateReceive( _
    ByVal receiveItem As Receive, _
    ByVal orderTransactionUUID As String, _
    Optional ByVal originalUUID As String = "")

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim affected As Long
    Dim transactionUUID As String
    Dim isCreatedForUpdate As Boolean
    Dim activityCode As String
    Dim inTransaction As Boolean
    Dim errInfo As ErrorInfo

    inTransaction = False

    ' =====================================================
    ' Step 1: Validate
    ' =====================================================
    If receiveItem Is Nothing Then
        managerError.ThrowError _
            serviceReceive.EmptyReceive, _
            MODULE_NAME & ".CreateReceive", _
            "Receive tidak boleh kosong."
    End If

    orderTransactionUUID = Trim$(orderTransactionUUID)
    originalUUID = Trim$(originalUUID)

    isCreatedForUpdate = (Len(originalUUID) > 0)

    If managerValidationText.IsBlank(orderTransactionUUID) Then
        managerError.ThrowError _
            serviceReceive.EmptyOrderTransactionUUID, _
            MODULE_NAME & ".CreateReceive", _
            "UUID transaksi order tidak boleh kosong."
    End If

    If receiveItem.totalReceiveAmount <> _
        receiveItem.okAmount + receiveItem.notOKAmount Then

        managerError.ThrowError _
            serviceReceive.InvalidTotalReceiveAmount, _
            MODULE_NAME & ".CreateReceive", _
            "Total penerimaan harus sama dengan jumlah OK " & _
            "ditambah jumlah tidak OK."
    End If

    ' =====================================================
    ' Step 2: Begin transaction
    ' =====================================================
    Set conn = managerDatabase.GetConnection()

    managerDatabase.BeginTransaction conn
    inTransaction = True

    ' =====================================================
    ' Step 3: Create the RECEIVE transaction
    '
    ' Normal receive:
    '   active = True
    '
    ' Proposed update:
    '   active = False
    ' =====================================================
    transactionUUID = serviceTransaction.CreateTransaction( _
        conn:=conn, _
        transactionType:="RECEIVE", _
        transactionDate:=receiveItem.ArrivalDate, _
        isActive:=Not isCreatedForUpdate)

    transactionUUID = Trim$(transactionUUID)

    If managerValidationText.IsBlank(transactionUUID) Then
        managerError.ThrowError _
            serviceReceive.EmptyTransactionUUID, _
            MODULE_NAME & ".CreateReceive", _
            "UUID transaksi receive tidak berhasil dibuat."
    End If

    If isCreatedForUpdate Then
        If managerValidationText.IsSameText(originalUUID, transactionUUID) Then
            managerError.ThrowError _
                serviceReceive.SameOriginalProposedUUID, _
                MODULE_NAME & ".CreateReceive", _
                "UUID transaksi original dan proposed " & _
                "tidak boleh sama."
        End If
    End If

    ' =====================================================
    ' Step 4: Create the receive
    ' =====================================================
    affected = daoReceive.CreateReceive( _
        conn, transactionUUID, orderTransactionUUID, receiveItem)

    If managerValidationNumeric.IsLessThan(affected, 1) Then
        managerError.ThrowError _
            serviceReceive.NoRowAffectedCreateReceive, _
            MODULE_NAME & ".CreateReceive", _
            "Penerimaan tidak tersimpan. " & _
            "Order dengan UUID tersebut mungkin " & _
            "tidak ditemukan."
    End If

    If managerValidationNumeric.IsGreaterThan(affected, 1) Then
        managerError.ThrowError _
            serviceReceive.MultipleRowsAffectedCreateReceive, _
            MODULE_NAME & ".CreateReceive", _
            "Lebih dari satu data penerimaan tersimpan. " & _
            "Transaksi dibatalkan untuk mencegah " & _
            "perubahan data yang tidak sesuai."
    End If

    ' =====================================================
    ' Step 5: Create update request or update stock
    '
    ' A proposed receive is inactive and must not change
    ' the plant stock until the proposal is approved.
    ' =====================================================
    If isCreatedForUpdate Then
        serviceTransactionRequestUpdate.CreateRequestUpdate _
            conn, originalUUID, transactionUUID
    Else
        servicePlantStock.ApplyStockChange _
            conn, _
            receiveItem.okAmount, _
            receiveItem.Order.Catalogue.Name, _
            receiveItem.Order.Plant.Name
    End If

    ' =====================================================
    ' Step 6: Write activity log
    ' =====================================================
    If isCreatedForUpdate Then
        activityCode = "Receive.CreateUpdateProposal"
    Else
        activityCode = "Receive.Create"
    End If

    serviceActivityLog.WriteActivityLog _
        conn, _
        activityCode, _
        BuildReceiveDetail(receiveItem, transactionUUID, orderTransactionUUID) & _
        BuildReceiveUpdateDetail(isCreatedForUpdate, originalUUID)

    ' =====================================================
    ' Step 7: Commit
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
        MODULE_NAME & ".CreateReceive", _
        "Gagal membuat data penerimaan."

End Sub


' Receives of one manufacturing plant with an arrival date
' between fromDate and untilDate, keyed by receive UUID.
Public Function SelectReceives( _
    ByVal manufacturingPlantName As String, _
    ByVal fromDate As Date, _
    ByVal untilDate As Date) As Collection

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim errInfo As ErrorInfo

    manufacturingPlantName = Trim$(manufacturingPlantName)

    ' =====================================================
    ' Step 1: Validate
    ' =====================================================
    If managerValidationText.IsBlank(manufacturingPlantName) Then
        managerError.ThrowError _
            serviceReceive.EmptyManufacturingPlantName, _
            MODULE_NAME & ".SelectReceives", _
            "Nama manufacturing plant tidak boleh kosong."
    End If

    If managerValidationNumeric.IsZeroOrNegative(fromDate) Then
        managerError.ThrowError _
            serviceReceive.EmptyFromDate, _
            MODULE_NAME & ".SelectReceives", _
            "Tanggal awal tidak boleh kosong."
    End If

    If managerValidationNumeric.IsZeroOrNegative(untilDate) Then
        managerError.ThrowError _
            serviceReceive.EmptyUntilDate, _
            MODULE_NAME & ".SelectReceives", _
            "Tanggal akhir tidak boleh kosong."
    End If

    If DateValue(untilDate) < DateValue(fromDate) Then
        managerError.ThrowError _
            serviceReceive.InvalidDateRange, _
            MODULE_NAME & ".SelectReceives", _
            "Tanggal akhir tidak boleh lebih kecil dari tanggal awal."
    End If

    ' =====================================================
    ' Step 2: Query
    ' =====================================================
    Set conn = managerDatabase.GetConnection()

    Set SelectReceives = daoReceive.SelectReceives( _
        conn, manufacturingPlantName, DateValue(fromDate), DateValue(untilDate))

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".SelectReceives", _
        "Gagal mengambil data penerimaan."

End Function


' Pending orders of one manufacturing plant as Receive
' objects (Order + remaining order amount), keyed by order
' UUID.
Public Function SelectPendingOrderReceive( _
    ByVal manufacturingPlantName As String, _
    ByVal fromDate As Date, _
    ByVal untilDate As Date) As Collection

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim errInfo As ErrorInfo

    manufacturingPlantName = Trim$(manufacturingPlantName)

    ' =====================================================
    ' Step 1: Validate
    ' =====================================================
    If managerValidationText.IsBlank(manufacturingPlantName) Then
        managerError.ThrowError _
            serviceReceive.EmptyManufacturingPlantName, _
            MODULE_NAME & ".SelectPendingOrderReceive", _
            "Nama manufacturing plant tidak boleh kosong."
    End If

    If DateValue(untilDate) < DateValue(fromDate) Then
        managerError.ThrowError _
            serviceReceive.InvalidDateRange, _
            MODULE_NAME & ".SelectPendingOrderReceive", _
            "Tanggal akhir tidak boleh lebih kecil dari tanggal awal."
    End If

    ' =====================================================
    ' Step 2: Query
    ' =====================================================
    Set conn = managerDatabase.GetConnection()

    Set SelectPendingOrderReceive = daoReceive.SelectPendingOrderReceive( _
        conn, manufacturingPlantName, DateValue(fromDate), DateValue(untilDate))

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".SelectPendingOrderReceive", _
        "Gagal mengambil data order dan penerimaan."

End Function


' =========================================================
' INTERNAL
' =========================================================

Private Function BuildReceiveDetail( _
    ByVal receiveItem As Receive, _
    ByVal receiveTransactionUuid As String, _
    ByVal orderTransactionUUID As String) As String

    Dim detail As String

    detail = _
        "Receive UUID: " & _
            Trim$(receiveTransactionUuid) & vbCrLf & _
        "Order UUID: " & _
            Trim$(orderTransactionUUID) & vbCrLf & _
        "Arrival Date: " & _
            Format$(receiveItem.ArrivalDate, _
                    "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
        "Verification Date: " & _
            Format$(receiveItem.VerificationDate, _
                    "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
        "Total Receive Amount: " & _
            CStr(receiveItem.totalReceiveAmount) & vbCrLf & _
        "OK Amount: " & _
            CStr(receiveItem.okAmount) & vbCrLf & _
        "Not OK Amount: " & _
            CStr(receiveItem.notOKAmount) & vbCrLf & _
        "Remarks: " & _
            Trim$(receiveItem.remarks) & vbCrLf & _
        "System Remarks: " & _
            Trim$(receiveItem.SystemRemarks) & vbCrLf & _
        "Person In Charge: " & _
            Trim$(receiveItem.personInCharge)

    BuildReceiveDetail = detail

End Function


Private Function BuildReceiveUpdateDetail( _
    ByVal isCreatedForUpdate As Boolean, _
    ByVal originalUUID As String) As String

    originalUUID = Trim$(originalUUID)

    If Not isCreatedForUpdate Then
        BuildReceiveUpdateDetail = vbNullString
        Exit Function
    End If

    BuildReceiveUpdateDetail = _
        "; update_proposal=true" & _
        "; original_transaction_uuid=" & _
        originalUUID

End Function
