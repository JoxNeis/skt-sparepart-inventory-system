Attribute VB_Name = "Receive"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' =========================================================
' MODEL - RECEIVE
' Table      : receives
' Error codes: entitiesReceive (ErrorCodesEntities.bas)
'
' Detail of a RECEIVE transaction - goods arriving against
' an order, split into OK and NOT OK after verification.
' Only OkQuantity enters stock, on VerificationDate.
'
' Table rule (ck_receives_quantity):
'   ReceivedQuantity = OkQuantity + NotOkQuantity
' Set ReceivedQuantity first, then OkQuantity and
' NotOkQuantity; call Validate before saving to check the
' total balances.
'
' Renamed from the old model:
'   totalReceiveAmount -> ReceivedQuantity (Long)
'   okAmount           -> OkQuantity       (Long)
'   notOKAmount        -> NotOkQuantity    (Long)
'   RemainingOrder     -> RemainingQuantityBefore
'   UUID               -> Transaction.UUID
'   unitOfMeasurement  -> Order.Sparepart.UnitOfMeasurement
' =========================================================

Private Const CLASS_NAME As String = "Receive"

Private Const TRANSACTION_TYPE As String = "RECEIVE"

Private Const PERSON_IN_CHARGE_MAX_LENGTH As Long = 255

Private mId As Long
Private mTransaction As Transaction
Private mOrder As Order
Private mArrivalDate As Date
Private mVerificationDate As Date
Private mReceivedQuantity As Long
Private mOkQuantity As Long
Private mNotOkQuantity As Long
Private mRemarks As String
Private mSystemRemarks As String
Private mPersonInCharge As String

' receive_list.remaining_quantity_before
Private mRemainingQuantityBefore As Long


' =========================================================
' ID                                    [id]
' =========================================================
Public Property Get Id() As Long
    Id = mId
End Property

Public Property Let Id(ByVal value As Long)

    If ValidationNumeric.IsZeroOrNegative(value) Then
        ErrorManager.ThrowError _
            entitiesReceive.NegativeId, _
            CLASS_NAME & ".Id", _
            CLASS_NAME & ": Id harus lebih besar dari 0."
    End If

    mId = value

End Property


' =========================================================
' TRANSACTION                           [transaction_id] NOT NULL UNIQUE
' Must be a RECEIVE header (when its type is set).
' =========================================================
Public Property Get Transaction() As Transaction
    Set Transaction = mTransaction
End Property

Public Property Set Transaction(ByVal value As Transaction)

    If value Is Nothing Then
        ErrorManager.ThrowError _
            entitiesReceive.EmptyTransaction, _
            CLASS_NAME & ".Transaction", _
            CLASS_NAME & ": Transaksi tidak boleh kosong."
    End If

    If Len(value.TransactionType) > 0 Then
        If value.TransactionType <> TRANSACTION_TYPE Then
            ErrorManager.ThrowError _
                entitiesReceive.WrongTransactionType, _
                CLASS_NAME & ".Transaction", _
                CLASS_NAME & ": Tipe transaksi harus " & _
                TRANSACTION_TYPE & "."
        End If
    End If

    Set mTransaction = value

End Property


' =========================================================
' ORDER                                 [order_id] NOT NULL
' =========================================================
Public Property Get Order() As Order
    Set Order = mOrder
End Property

Public Property Set Order(ByVal value As Order)

    If value Is Nothing Then
        ErrorManager.ThrowError _
            entitiesReceive.EmptyOrder, _
            CLASS_NAME & ".Order", _
            CLASS_NAME & ": Order tidak boleh kosong."
    End If

    If value.Quantity > 0 Then

        If mReceivedQuantity > value.Quantity Then
            ErrorManager.ThrowError _
                entitiesReceive.ReceivedQuantityExceedsOrder, _
                CLASS_NAME & ".Order", _
                CLASS_NAME & _
                ": Total penerimaan tidak boleh melebihi jumlah order."
        End If

        If mOkQuantity > value.Quantity Then
            ErrorManager.ThrowError _
                entitiesReceive.OkQuantityExceedsOrder, _
                CLASS_NAME & ".Order", _
                CLASS_NAME & _
                ": Jumlah barang OK tidak boleh melebihi jumlah order."
        End If

    End If

    Set mOrder = value

End Property


' =========================================================
' ARRIVAL DATE                          [arrival_date] NOT NULL
' =========================================================
Public Property Get ArrivalDate() As Date
    ArrivalDate = mArrivalDate
End Property

Public Property Let ArrivalDate(ByVal value As Date)

    If ValidationNumeric.IsZeroOrNegative(value) Then
        ErrorManager.ThrowError _
            entitiesReceive.EmptyArrivalDate, _
            CLASS_NAME & ".ArrivalDate", _
            CLASS_NAME & ": Tanggal kedatangan tidak boleh kosong."
    End If

    If DateValue(value) > VBA.Date Then
        ErrorManager.ThrowError _
            entitiesReceive.FutureArrivalDate, _
            CLASS_NAME & ".ArrivalDate", _
            CLASS_NAME & _
            ": Tanggal kedatangan tidak boleh berada di masa depan."
    End If

    If mVerificationDate <> 0 Then
        If DateValue(value) > DateValue(mVerificationDate) Then
            ErrorManager.ThrowError _
                entitiesReceive.ArrivalAfterVerification, _
                CLASS_NAME & ".ArrivalDate", _
                CLASS_NAME & _
                ": Tanggal kedatangan tidak boleh setelah tanggal verifikasi."
        End If
    End If

    mArrivalDate = value

End Property


' =========================================================
' VERIFICATION DATE                     [verification_date] (optional)
' 0 = NULL = not verified yet (nothing enters stock).
' =========================================================
Public Property Get VerificationDate() As Date
    VerificationDate = mVerificationDate
End Property

Public Property Let VerificationDate(ByVal value As Date)

    If ValidationNumeric.IsNegative(value) Then
        ErrorManager.ThrowError _
            entitiesReceive.InvalidVerificationDate, _
            CLASS_NAME & ".VerificationDate", _
            CLASS_NAME & ": Tanggal verifikasi tidak valid."
    End If

    If value <> 0 Then

        If DateValue(value) > VBA.Date Then
            ErrorManager.ThrowError _
                entitiesReceive.FutureVerificationDate, _
                CLASS_NAME & ".VerificationDate", _
                CLASS_NAME & _
                ": Tanggal verifikasi tidak boleh berada di masa depan."
        End If

        If mArrivalDate <> 0 Then
            If DateValue(value) < DateValue(mArrivalDate) Then
                ErrorManager.ThrowError _
                    entitiesReceive.VerificationBeforeArrival, _
                    CLASS_NAME & ".VerificationDate", _
                    CLASS_NAME & _
                    ": Tanggal verifikasi tidak boleh sebelum tanggal kedatangan."
            End If
        End If

    End If

    mVerificationDate = value

End Property


' =========================================================
' RECEIVED QUANTITY                     [received_quantity] LONG NOT NULL
' =========================================================
Public Property Get ReceivedQuantity() As Long
    ReceivedQuantity = mReceivedQuantity
End Property

Public Property Let ReceivedQuantity(ByVal value As Long)

    If ValidationNumeric.IsNegative(value) Then
        ErrorManager.ThrowError _
            entitiesReceive.NegativeReceivedQuantity, _
            CLASS_NAME & ".ReceivedQuantity", _
            CLASS_NAME & ": Total penerimaan tidak boleh negatif."
    End If

    If Not mOrder Is Nothing Then
        If mOrder.Quantity > 0 And value > mOrder.Quantity Then
            ErrorManager.ThrowError _
                entitiesReceive.ReceivedQuantityExceedsOrder, _
                CLASS_NAME & ".ReceivedQuantity", _
                CLASS_NAME & _
                ": Total penerimaan tidak boleh melebihi jumlah order."
        End If
    End If

    If mOkQuantity + mNotOkQuantity > value Then
        ErrorManager.ThrowError _
            entitiesReceive.OkNotOkExceedsReceived, _
            CLASS_NAME & ".ReceivedQuantity", _
            CLASS_NAME & _
            ": Total barang OK dan tidak OK melebihi total penerimaan."
    End If

    mReceivedQuantity = value

End Property


' =========================================================
' OK QUANTITY                           [ok_quantity] LONG NOT NULL >= 0
' =========================================================
Public Property Get OkQuantity() As Long
    OkQuantity = mOkQuantity
End Property

Public Property Let OkQuantity(ByVal value As Long)

    If ValidationNumeric.IsNegative(value) Then
        ErrorManager.ThrowError _
            entitiesReceive.NegativeOkQuantity, _
            CLASS_NAME & ".OkQuantity", _
            CLASS_NAME & ": Jumlah barang OK tidak boleh negatif."
    End If

    If value + mNotOkQuantity > mReceivedQuantity Then
        ErrorManager.ThrowError _
            entitiesReceive.OkNotOkExceedsReceived, _
            CLASS_NAME & ".OkQuantity", _
            CLASS_NAME & _
            ": Total barang OK dan tidak OK melebihi total penerimaan."
    End If

    If Not mOrder Is Nothing Then
        If mOrder.Quantity > 0 And value > mOrder.Quantity Then
            ErrorManager.ThrowError _
                entitiesReceive.OkQuantityExceedsOrder, _
                CLASS_NAME & ".OkQuantity", _
                CLASS_NAME & _
                ": Jumlah barang OK tidak boleh melebihi jumlah order."
        End If
    End If

    mOkQuantity = value

End Property


' =========================================================
' NOT OK QUANTITY                       [not_ok_quantity] LONG NOT NULL >= 0
' =========================================================
Public Property Get NotOkQuantity() As Long
    NotOkQuantity = mNotOkQuantity
End Property

Public Property Let NotOkQuantity(ByVal value As Long)

    If ValidationNumeric.IsNegative(value) Then
        ErrorManager.ThrowError _
            entitiesReceive.NegativeNotOkQuantity, _
            CLASS_NAME & ".NotOkQuantity", _
            CLASS_NAME & ": Jumlah barang tidak OK tidak boleh negatif."
    End If

    If mOkQuantity + value > mReceivedQuantity Then
        ErrorManager.ThrowError _
            entitiesReceive.OkNotOkExceedsReceived, _
            CLASS_NAME & ".NotOkQuantity", _
            CLASS_NAME & _
            ": Total barang OK dan tidak OK melebihi total penerimaan."
    End If

    mNotOkQuantity = value

End Property


' =========================================================
' REMARKS / SYSTEM REMARKS              [remarks] [system_remarks] LONGTEXT
' =========================================================
Public Property Get Remarks() As String
    Remarks = mRemarks
End Property

Public Property Let Remarks(ByVal value As String)
    mRemarks = Trim$(value)
End Property

Public Property Get SystemRemarks() As String
    SystemRemarks = mSystemRemarks
End Property

Public Property Let SystemRemarks(ByVal value As String)
    mSystemRemarks = Trim$(value)
End Property


' =========================================================
' PERSON IN CHARGE                      [person_in_charge] TEXT(255)
' =========================================================
Public Property Get PersonInCharge() As String
    PersonInCharge = mPersonInCharge
End Property

Public Property Let PersonInCharge(ByVal value As String)

    value = Trim$(value)

    If ValidationText.IsBlank(value) Then
        ErrorManager.ThrowError _
            entitiesReceive.EmptyPersonInCharge, _
            CLASS_NAME & ".PersonInCharge", _
            CLASS_NAME & ": Person in charge tidak boleh kosong."
    End If

    If ValidationText.IsLongerThan(value, PERSON_IN_CHARGE_MAX_LENGTH) Then
        ErrorManager.ThrowError _
            entitiesReceive.PersonInChargeTooLong, _
            CLASS_NAME & ".PersonInCharge", _
            CLASS_NAME & ": Person in charge maksimal " & _
            PERSON_IN_CHARGE_MAX_LENGTH & " karakter."
    End If

    mPersonInCharge = value

End Property


' =========================================================
' REMAINING QUANTITY BEFORE (read, not stored)
' Order quantity minus the OK quantity of the receives before
' this one (receive_list.remaining_quantity_before), or the
' order's remaining quantity for a new receive.
' =========================================================
Public Property Get RemainingQuantityBefore() As Long
    RemainingQuantityBefore = mRemainingQuantityBefore
End Property

Public Property Let RemainingQuantityBefore(ByVal value As Long)

    If ValidationNumeric.IsNegative(value) Then
        ErrorManager.ThrowError _
            entitiesReceive.NegativeRemainingQuantity, _
            CLASS_NAME & ".RemainingQuantityBefore", _
            CLASS_NAME & ": Sisa jumlah order tidak boleh negatif."
    End If

    If Not mOrder Is Nothing Then
        If mOrder.Quantity > 0 And value > mOrder.Quantity Then
            ErrorManager.ThrowError _
                entitiesReceive.RemainingQuantityExceedsOrder, _
                CLASS_NAME & ".RemainingQuantityBefore", _
                CLASS_NAME & _
                ": Sisa jumlah order tidak boleh melebihi jumlah order."
        End If
    End If

    mRemainingQuantityBefore = value

End Property


' =========================================================
' CALCULATED PROPERTIES
' =========================================================

' Document number of the RECEIVE header ("" when not set).
Public Property Get UUID() As String
    If Not mTransaction Is Nothing Then UUID = mTransaction.UUID
End Property

' Plant of the RECEIVE header (Nothing when not set).
Public Property Get ManufacturingPlant() As ManufacturingPlant
    If Not mTransaction Is Nothing Then
        Set ManufacturingPlant = mTransaction.ManufacturingPlant
    End If
End Property

Public Property Get OrderQuantity() As Long
    If Not mOrder Is Nothing Then OrderQuantity = mOrder.Quantity
End Property

Public Property Get UnitOfMeasurement() As String
    If Not mOrder Is Nothing Then
        UnitOfMeasurement = mOrder.UnitOfMeasurement
    End If
End Property

' RemainingQuantityBefore minus this receive's OK quantity,
' never below 0.
Public Property Get RemainingQuantityAfter() As Long

    Dim result As Long

    result = mRemainingQuantityBefore - mOkQuantity

    If result < 0 Then result = 0

    RemainingQuantityAfter = result

End Property

Public Property Get IsVerified() As Boolean
    IsVerified = (mVerificationDate <> 0)
End Property

' ck_receives_quantity: received = OK + NOT OK.
Public Property Get IsQuantityBalanced() As Boolean
    IsQuantityBalanced = (mReceivedQuantity = mOkQuantity + mNotOkQuantity)
End Property


' =========================================================
' VALIDATE
' Checks the rules that span several properties. Call it
' before saving.
' =========================================================
Public Sub Validate()

    If Not IsQuantityBalanced Then
        ErrorManager.ThrowError _
            entitiesReceive.QuantityNotBalanced, _
            CLASS_NAME & ".Validate", _
            CLASS_NAME & _
            ": Total penerimaan harus sama dengan barang OK + tidak OK."
    End If

End Sub
