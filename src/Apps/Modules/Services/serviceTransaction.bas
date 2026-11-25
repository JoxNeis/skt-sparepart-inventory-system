Attribute VB_Name = "serviceTransaction"
Option Explicit

' =========================================================
' SERVICE - TRANSACTION (document header)
'
' Helpers used by the document services (order, receive,
' request, update request) INSIDE their transaction. Every
' function receives the caller's connection.
'
' Document number: <PLANT>-<PREFIX>-<yyyymmdd>-<nnn>
'   ORDER       -> ORD      RECEIVE     -> RCV
'   REQUEST     -> WRQ      MAINTENANCE -> MTW
' The date is the document date (order / arrival / request
' date); the sequence restarts every day per plant and prefix.
' =========================================================

Private Const MODULE_NAME As String = "serviceTransaction"

' =========================================================
' ERROR CODES (17900 - 17999)
' Usage: managerError.ThrowError serviceTransaction.InvalidTransactionType, ...
' =========================================================
Public Const InvalidTransactionType As Long = 17901
Public Const EmptyPlant As Long = 17902
Public Const SequenceExhausted As Long = 17903      ' 999 documents on one day
Public Const CreateFailed As Long = 17904
Public Const ActivateFailed As Long = 17905         ' not a draft / reversed
Public Const ReverseFailed As Long = 17906          ' already reversed
Public Const EmptyUUID As Long = 17907


' Create a document header and return its new number.
'   isActive = True  : active document (activated now)
'   isActive = False : draft (a proposed update)
Public Function CreateTransaction( _
    ByVal conn As ADODB.connection, _
    ByVal plantItem As ManufacturingPlant, _
    ByVal transactionType As String, _
    ByVal documentDate As Date, _
    ByVal isActive As Boolean, _
    ByVal createdBy As String) As String

    On Error GoTo ErrorHandler

    Dim transactionUUID As String
    Dim activatedAt As Variant
    Dim affected As Long
    Dim errInfo As ErrorInfo

    transactionType = UCase$(Trim$(transactionType))

    If plantItem Is Nothing Then
        managerError.ThrowError _
            serviceTransaction.EmptyPlant, _
            MODULE_NAME & ".CreateTransaction", _
            "Plant transaksi tidak boleh kosong."
    End If

    transactionUUID = daoTransaction.GetNextUUID( _
        conn, plantItem.Code, UuidPrefix(transactionType), documentDate)

    If managerValidationText.IsBlank(transactionUUID) Then
        managerError.ThrowError _
            serviceTransaction.SequenceExhausted, _
            MODULE_NAME & ".CreateTransaction", _
            "Nomor dokumen " & transactionType & " untuk tanggal " & _
            Format$(documentDate, "dd-mm-yyyy") & _
            " sudah habis (maksimal 999 dokumen per hari)."
    End If

    If isActive Then
        activatedAt = Now
    Else
        activatedAt = Null
    End If

    affected = daoTransaction.CreateTransaction( _
        conn, transactionUUID, transactionType, plantItem.Id, _
        createdBy, activatedAt)

    managerDatabase.RequireAffectedRows affected, 1, _
        serviceTransaction.CreateFailed, _
        MODULE_NAME & ".CreateTransaction", _
        "Header transaksi " & transactionUUID & " tidak tersimpan."

    CreateTransaction = transactionUUID

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".CreateTransaction", _
        "Gagal membuat header transaksi " & transactionType & "."

End Function


' Draft -> active.
Public Sub ActivateTransaction( _
    ByVal conn As ADODB.connection, _
    ByVal transactionUUID As String)

    Dim affected As Long

    RequireUUID transactionUUID, MODULE_NAME & ".ActivateTransaction"

    affected = daoTransaction.ActivateTransaction(conn, transactionUUID)

    managerDatabase.RequireAffectedRows affected, 1, _
        serviceTransaction.ActivateFailed, _
        MODULE_NAME & ".ActivateTransaction", _
        "Transaksi " & transactionUUID & _
        " tidak dapat diaktifkan (bukan draft atau sudah dibatalkan)."

End Sub


' Active or draft -> reversed.
Public Sub ReverseTransaction( _
    ByVal conn As ADODB.connection, _
    ByVal transactionUUID As String, _
    ByVal reversedBy As String)

    Dim affected As Long

    RequireUUID transactionUUID, MODULE_NAME & ".ReverseTransaction"

    affected = daoTransaction.ReverseTransaction( _
        conn, transactionUUID, reversedBy)

    managerDatabase.RequireAffectedRows affected, 1, _
        serviceTransaction.ReverseFailed, _
        MODULE_NAME & ".ReverseTransaction", _
        "Transaksi " & transactionUUID & _
        " tidak dapat dibatalkan (sudah dibatalkan sebelumnya)."

End Sub


' ORD / RCV / WRQ / MTW
Public Function UuidPrefix( _
    ByVal transactionType As String) As String

    Select Case UCase$(Trim$(transactionType))
        Case "ORDER":       UuidPrefix = "ORD"
        Case "RECEIVE":     UuidPrefix = "RCV"
        Case "REQUEST":     UuidPrefix = "WRQ"
        Case "MAINTENANCE": UuidPrefix = "MTW"
        Case Else
            managerError.ThrowError _
                serviceTransaction.InvalidTransactionType, _
                MODULE_NAME & ".UuidPrefix", _
                "Tipe transaksi tidak dikenal: " & transactionType & "."
    End Select

End Function


' =========================================================
' INTERNAL
' =========================================================

Private Sub RequireUUID( _
    ByVal transactionUUID As String, _
    ByVal errorSource As String)

    If managerValidationText.IsBlank(transactionUUID) Then
        managerError.ThrowError _
            serviceTransaction.EmptyUUID, _
            errorSource, _
            "Nomor dokumen tidak boleh kosong."
    End If

End Sub
