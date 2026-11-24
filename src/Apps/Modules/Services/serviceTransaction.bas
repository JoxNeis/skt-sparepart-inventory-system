Attribute VB_Name = "serviceTransaction"
Option Explicit

' =========================================================
' SERVICE - TRANSACTION
'
' Helpers used by other services INSIDE their transaction.
' They receive the caller's connection and never begin,
' commit or roll back by themselves.
' =========================================================

Private Const MODULE_NAME As String = "SERVICE_TRANSACTION"

' =========================================================
' ERROR CODES (17900 - 17999)
' Usage: managerError.ThrowError serviceTransaction.InvalidTransactionType, ...
' =========================================================
Public Const InvalidTransactionType As Long = 17901         ' not ORDER / REQUEST / RECEIVE / MAINTENANCE
Public Const EmptyGeneratedUUID As Long = 17902             ' select_uuid returned no row / blank
Public Const NoRowAffectedCreateTransaction As Long = 17903 ' 0 rows inserted
Public Const EmptyTransactionCreator As Long = 17904        ' current user is blank
Public Const TransactionReversalFailed As Long = 17905      ' reversal affected <> 1 row
Public Const TransactionActivationFailed As Long = 17906    ' activation affected <> 1 row
Public Const EmptyTransactionUUID As Long = 17907


' Generate a UUID and insert the transaction row.
'
' transactionType : ORDER, REQUEST or RECEIVE
' isActive        : True  -> active_at = transactionDate
'                   False -> active_at = Null (proposed update)
'
' Returns the new transaction UUID.
Public Function CreateTransaction( _
    ByVal conn As ADODB.connection, _
    ByVal transactionType As String, _
    ByVal transactionDate As Date, _
    ByVal isActive As Boolean) As String

    On Error GoTo ErrorHandler

    Dim transactionCode As String
    Dim transactionUUID As String
    Dim createdBy As String
    Dim activeAt As Variant
    Dim affected As Long
    Dim errInfo As ErrorInfo

    transactionType = UCase$(Trim$(transactionType))

    ' =====================================================
    ' Step 1: Map transaction type to UUID prefix
    ' =====================================================
    transactionCode = GetTransactionCode(transactionType)

    If managerValidationText.IsBlank(transactionCode) Then
        managerError.ThrowError _
            serviceTransaction.InvalidTransactionType, _
            MODULE_NAME & ".CreateTransaction", _
            "Transaction type harus ORDER, RECEIVE, REQUEST, atau MAINTENANCE."
    End If

    createdBy = Trim$(coreConfig.currentUser())

    If managerValidationText.IsBlank(createdBy) Then
        managerError.ThrowError _
            serviceTransaction.EmptyTransactionCreator, _
            MODULE_NAME & ".CreateTransaction", _
            "User pembuat transaction tidak dapat ditentukan."
    End If

    ' =====================================================
    ' Step 2: Generate UUID
    ' =====================================================
    transactionUUID = daoTransaction.SelectUUID( _
        conn, transactionCode, transactionDate)

    If managerValidationText.IsBlank(transactionUUID) Then
        managerError.ThrowError _
            serviceTransaction.EmptyGeneratedUUID, _
            MODULE_NAME & ".CreateTransaction", _
            "UUID gagal dibuat karena query tidak menghasilkan data."
    End If

    ' =====================================================
    ' Step 3: Insert transaction
    ' =====================================================
    If isActive Then
        activeAt = transactionDate
    Else
        activeAt = Null
    End If

    affected = daoTransaction.CreateTransaction( _
        conn, transactionUUID, transactionType, createdBy, activeAt)

    If managerValidationNumeric.IsLessThan(affected, 1) Then
        managerError.ThrowError _
            serviceTransaction.NoRowAffectedCreateTransaction, _
            MODULE_NAME & ".CreateTransaction", _
            "Transaction tidak tersimpan karena 0 baris terpengaruh."
    End If

    CreateTransaction = transactionUUID

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".CreateTransaction", _
        "Gagal membuat transaction."

End Function


' Mark the original transaction as reversed. Exactly one
' row must change.
Public Sub ReverseTransaction( _
    ByVal conn As ADODB.connection, _
    ByVal originalTransactionUUID As String)

    On Error GoTo ErrorHandler

    Dim affected As Long
    Dim errInfo As ErrorInfo

    originalTransactionUUID = Trim$(originalTransactionUUID)

    If managerValidationText.IsBlank(originalTransactionUUID) Then
        managerError.ThrowError _
            serviceTransaction.EmptyTransactionUUID, _
            MODULE_NAME & ".ReverseTransaction", _
            "UUID transaksi original tidak boleh kosong."
    End If

    affected = daoTransaction.ReverseTransaction( _
        conn, originalTransactionUUID)

    If affected <> 1 Then
        managerError.ThrowError _
            serviceTransaction.TransactionReversalFailed, _
            MODULE_NAME & ".ReverseTransaction", _
            "Transaksi original tidak berhasil dibalik. " & _
            "Jumlah baris terpengaruh: " & CStr(affected) & "."
    End If

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ReverseTransaction", _
        "Gagal membalik transaksi original."

End Sub


' Activate the proposed transaction. Exactly one row must
' change.
Public Sub ActivateTransaction( _
    ByVal conn As ADODB.connection, _
    ByVal transactionUUID As String)

    On Error GoTo ErrorHandler

    Dim affected As Long
    Dim errInfo As ErrorInfo

    transactionUUID = Trim$(transactionUUID)

    If managerValidationText.IsBlank(transactionUUID) Then
        managerError.ThrowError _
            serviceTransaction.EmptyTransactionUUID, _
            MODULE_NAME & ".ActivateTransaction", _
            "UUID transaksi proposed tidak boleh kosong."
    End If

    affected = daoTransaction.ActivateTransaction( _
        conn, transactionUUID)

    If affected <> 1 Then
        managerError.ThrowError _
            serviceTransaction.TransactionActivationFailed, _
            MODULE_NAME & ".ActivateTransaction", _
            "Transaksi proposed tidak berhasil diaktifkan. " & _
            "Jumlah baris terpengaruh: " & CStr(affected) & "."
    End If

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".ActivateTransaction", _
        "Gagal mengaktifkan transaksi proposed."

End Sub


' ORDER -> ORD, REQUEST -> WRQ, RECEIVE -> RCV,
' MAINTENANCE -> MTW, else "".
Private Function GetTransactionCode( _
    ByVal transactionType As String) As String

    Select Case transactionType
        Case "ORDER"
            GetTransactionCode = "ORD"
        Case "REQUEST"
            GetTransactionCode = "WRQ"
        Case "RECEIVE"
            GetTransactionCode = "RCV"
        Case "MAINTENANCE"
            GetTransactionCode = "MTW"
        Case Else
            GetTransactionCode = vbNullString
    End Select

End Function
