Attribute VB_Name = "Receive"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' Error codes: entitiesReceive (ErrorCodesEntities.bas)

Private Const CLASS_NAME As String = "Receive"

' Amounts are compared after rounding to this many decimals.
Private Const AMOUNT_DECIMALS As Long = 4

Private mId As Long
Private mUUID As String
Private mOrder As Order
Private mRemainingOrder As Double
Private mTotalReceiveAmount As Double
Private mOKAmount As Double
Private mNotOKAmount As Double
Private mUnitOfMeasurement As String
Private mArrivalDate As Date
Private mVerificationDate As Date
Private mRemarks As String
Private mSystemRemarks As String
Private mPIC As String


' =========================================================
' ID
' =========================================================

Public Property Get Id() As Long
    Id = mId
End Property

Public Property Let Id(ByVal value As Long)

    If ValidationNumeric.IsNegative(value) Then
        ErrorManager.ThrowError _
            entitiesReceive.NegativeId, _
            CLASS_NAME & ".Id", _
            CLASS_NAME & ": ID tidak boleh negatif."
    End If

    mId = value

End Property


' =========================================================
' UUID
' =========================================================

Public Property Get UUID() As String
    UUID = mUUID
End Property

Public Property Let UUID(ByVal value As String)

    value = Trim$(value)

    If ValidationText.IsBlank(value) Then
        ErrorManager.ThrowError _
            entitiesReceive.EmptyUUID, _
            CLASS_NAME & ".UUID", _
            CLASS_NAME & ": UUID tidak boleh kosong."
    End If

    mUUID = value

End Property


' =========================================================
' ORDER
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

    If ValidationNumeric.IsPositive(mTotalReceiveAmount) Then
        If ValidationNumeric.IsGreaterThan( _
            mTotalReceiveAmount, value.Amount, AMOUNT_DECIMALS) Then

            ErrorManager.ThrowError _
                entitiesReceive.TotalReceiveExceedsOrder, _
                CLASS_NAME & ".Order", _
                CLASS_NAME & _
                ": Total penerimaan tidak boleh melebihi jumlah order."
        End If
    End If

    Set mOrder = value

End Property


' =========================================================
' TOTAL RECEIVE AMOUNT
' =========================================================

Public Property Get totalReceiveAmount() As Double
    totalReceiveAmount = mTotalReceiveAmount
End Property

Public Property Let totalReceiveAmount(ByVal value As Double)

    If ValidationNumeric.IsNegative(value) Then
        ErrorManager.ThrowError _
            entitiesReceive.NegativeTotalReceive, _
            CLASS_NAME & ".TotalReceiveAmount", _
            CLASS_NAME & ": Total penerimaan tidak boleh negatif."
    End If

    If Not mOrder Is Nothing Then
        If ValidationNumeric.IsGreaterThan( _
            value, mOrder.Amount, AMOUNT_DECIMALS) Then

            ErrorManager.ThrowError _
                entitiesReceive.TotalReceiveExceedsOrder, _
                CLASS_NAME & ".TotalReceiveAmount", _
                CLASS_NAME & _
                ": Total penerimaan tidak boleh melebihi jumlah order."
        End If
    End If

    If ValidationNumeric.IsGreaterThan( _
        mOKAmount + mNotOKAmount, value, AMOUNT_DECIMALS) Then

        ErrorManager.ThrowError _
            entitiesReceive.OkNotOkExceedsTotalReceive, _
            CLASS_NAME & ".TotalReceiveAmount", _
            CLASS_NAME & _
            ": Total barang OK dan tidak OK melebihi total penerimaan."
    End If

    mTotalReceiveAmount = value

End Property


' =========================================================
' OK AMOUNT
' =========================================================

Public Property Get okAmount() As Double
    okAmount = mOKAmount
End Property

Public Property Let okAmount(ByVal value As Double)

    If ValidationNumeric.IsNegative(value) Then
        ErrorManager.ThrowError _
            entitiesReceive.NegativeOkAmount, _
            CLASS_NAME & ".OKAmount", _
            CLASS_NAME & ": Jumlah barang OK tidak boleh negatif."
    End If

    If ValidationNumeric.IsGreaterThan( _
        value + mNotOKAmount, mTotalReceiveAmount, AMOUNT_DECIMALS) Then

        ErrorManager.ThrowError _
            entitiesReceive.OkNotOkExceedsTotalReceive, _
            CLASS_NAME & ".OKAmount", _
            CLASS_NAME & _
            ": Total barang OK dan tidak OK melebihi total penerimaan."
    End If

    If Not mOrder Is Nothing Then
        If ValidationNumeric.IsGreaterThan( _
            value, mOrder.Amount, AMOUNT_DECIMALS) Then

            ErrorManager.ThrowError _
                entitiesReceive.OkAmountExceedsOrder, _
                CLASS_NAME & ".OKAmount", _
                CLASS_NAME & _
                ": Jumlah barang OK tidak boleh melebihi jumlah order."
        End If
    End If

    mOKAmount = value

End Property


' =========================================================
' NOT OK AMOUNT
' =========================================================

Public Property Get notOKAmount() As Double
    notOKAmount = mNotOKAmount
End Property

Public Property Let notOKAmount(ByVal value As Double)

    If ValidationNumeric.IsNegative(value) Then
        ErrorManager.ThrowError _
            entitiesReceive.NegativeNotOkAmount, _
            CLASS_NAME & ".NotOKAmount", _
            CLASS_NAME & ": Jumlah barang tidak OK tidak boleh negatif."
    End If

    If ValidationNumeric.IsGreaterThan( _
        mOKAmount + value, mTotalReceiveAmount, AMOUNT_DECIMALS) Then

        ErrorManager.ThrowError _
            entitiesReceive.OkNotOkExceedsTotalReceive, _
            CLASS_NAME & ".NotOKAmount", _
            CLASS_NAME & _
            ": Total barang OK dan tidak OK melebihi total penerimaan."
    End If

    mNotOKAmount = value

End Property


' =========================================================
' UNIT OF MEASUREMENT
' =========================================================

Public Property Get unitOfMeasurement() As String
    unitOfMeasurement = mUnitOfMeasurement
End Property

Public Property Let unitOfMeasurement(ByVal value As String)

    value = Trim$(value)

    If ValidationText.IsBlank(value) Then
        ErrorManager.ThrowError _
            entitiesReceive.EmptyUnitOfMeasurement, _
            CLASS_NAME & ".UnitOfMeasurement", _
            CLASS_NAME & ": Satuan jumlah barang tidak boleh kosong."
    End If

    mUnitOfMeasurement = value

End Property


' =========================================================
' ARRIVAL DATE
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
' VERIFICATION DATE
' =========================================================

Public Property Get VerificationDate() As Date
    VerificationDate = mVerificationDate
End Property

Public Property Let VerificationDate(ByVal value As Date)

    If ValidationNumeric.IsZeroOrNegative(value) Then
        ErrorManager.ThrowError _
            entitiesReceive.EmptyVerificationDate, _
            CLASS_NAME & ".VerificationDate", _
            CLASS_NAME & ": Tanggal verifikasi tidak boleh kosong."
    End If

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

    mVerificationDate = value

End Property


' =========================================================
' REMARKS
' =========================================================

Public Property Get remarks() As String
    remarks = mRemarks
End Property

Public Property Let remarks(ByVal value As String)
    mRemarks = Trim$(value)
End Property


' =========================================================
' SYSTEM REMARKS
' =========================================================

Public Property Get SystemRemarks() As String
    SystemRemarks = mSystemRemarks
End Property

Public Property Let SystemRemarks(ByVal value As String)
    mSystemRemarks = Trim$(value)
End Property


' =========================================================
' PERSON IN CHARGE
' =========================================================

Public Property Get personInCharge() As String
    personInCharge = mPIC
End Property

Public Property Let personInCharge(ByVal value As String)

    value = Trim$(value)

    If ValidationText.IsBlank(value) Then
        ErrorManager.ThrowError _
            entitiesReceive.EmptyPersonInCharge, _
            CLASS_NAME & ".PersonInCharge", _
            CLASS_NAME & ": PIC tidak boleh kosong."
    End If

    mPIC = value

End Property

' =========================================================
' REMAINING ORDER
' =========================================================

Public Property Get RemainingOrder() As Double
    RemainingOrder = mRemainingOrder
End Property

Public Property Let RemainingOrder(ByVal value As Double)

    If ValidationNumeric.IsNegative(value) Then
        ErrorManager.ThrowError _
            entitiesReceive.NegativeRemainingOrder, _
            CLASS_NAME & ".RemainingOrder", _
            CLASS_NAME & ": Sisa jumlah order tidak boleh negatif."
    End If

    If Not mOrder Is Nothing Then
        If ValidationNumeric.IsGreaterThan( _
            value, mOrder.Amount, AMOUNT_DECIMALS) Then

            ErrorManager.ThrowError _
                entitiesReceive.RemainingOrderExceedsOrder, _
                CLASS_NAME & ".RemainingOrder", _
                CLASS_NAME & _
                ": Sisa jumlah order tidak boleh melebihi jumlah order."
        End If
    End If

    mRemainingOrder = value

End Property

' =========================================================
' CALCULATED PROPERTIES
' =========================================================
Public Property Get orderAmount() As Double

    If mOrder Is Nothing Then
        orderAmount = 0
    Else
        orderAmount = mOrder.Amount
    End If

End Property


Public Property Get RemainingAmount() As Double

    Dim result As Double

    If mOrder Is Nothing Then
        RemainingAmount = 0
        Exit Property
    End If

    result = mOrder.Amount - mOKAmount

    If result < 0 Then
        result = 0
    End If

    RemainingAmount = result

End Property
