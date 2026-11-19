Attribute VB_Name = "StockMovement"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' Error codes: entitiesStockMovement (ErrorCodesEntities.bas)

Private Const CLASS_NAME As String = "StockMovement"

Private mDate As Long
Private mBeginningStock As Double
Private mIncomingReferenceNumber As String
Private mIncomingFrom As String
Private mIncomingAmount As Double
Private mOutgoingReferenceNumber As String
Private mOutgoingTo As String
Private mOutgoingAmount As Double
Private mEndStock As Double
Private mRemarks As String
Private mPersonInCharge As String

Public Property Get MovementDate() As Long
    MovementDate = mDate
End Property

Public Property Let MovementDate(ByVal value As Long)

    If ValidationNumeric.IsZeroOrNegative(value) Then
        ErrorManager.ThrowError _
            entitiesStockMovement.InvalidMovementDate, _
            CLASS_NAME & ".MovementDate", _
            CLASS_NAME & ": Tanggal harus lebih besar dari 0."
    End If

    mDate = value

End Property

Public Property Get beginningStock() As Double
    beginningStock = mBeginningStock
End Property

Public Property Let beginningStock(ByVal value As Double)

    If ValidationNumeric.IsNegative(value) Then
        ErrorManager.ThrowError _
            entitiesStockMovement.NegativeBeginningStock, _
            CLASS_NAME & ".BeginningStock", _
            CLASS_NAME & ": Stok awal tidak boleh kurang dari 0."
    End If

    mBeginningStock = value

End Property

Public Property Get IncomingReferenceNumber() As String
    IncomingReferenceNumber = mIncomingReferenceNumber
End Property

Public Property Let IncomingReferenceNumber(ByVal value As String)

    value = Trim$(value)

    mIncomingReferenceNumber = value

End Property

Public Property Get IncomingFrom() As String
    IncomingFrom = mIncomingFrom
End Property

Public Property Let IncomingFrom(ByVal value As String)

    value = Trim$(value)

    mIncomingFrom = value

End Property

Public Property Get incomingAmount() As Double
    incomingAmount = mIncomingAmount
End Property

Public Property Let incomingAmount(ByVal value As Double)

    If ValidationNumeric.IsNegative(value) Then
        ErrorManager.ThrowError _
            entitiesStockMovement.NegativeIncomingAmount, _
            CLASS_NAME & ".IncomingAmount", _
            CLASS_NAME & ": Jumlah barang masuk tidak boleh kurang dari 0."
    End If

    mIncomingAmount = value

End Property

Public Property Get OutgoingReferenceNumber() As String
    OutgoingReferenceNumber = mOutgoingReferenceNumber
End Property

Public Property Let OutgoingReferenceNumber(ByVal value As String)

    value = Trim$(value)

    mOutgoingReferenceNumber = value

End Property

Public Property Get OutgoingTo() As String
    OutgoingTo = mOutgoingTo
End Property

Public Property Let OutgoingTo(ByVal value As String)

    value = Trim$(value)

    mOutgoingTo = value

End Property

Public Property Get outgoingAmount() As Double
    outgoingAmount = mOutgoingAmount
End Property

Public Property Let outgoingAmount(ByVal value As Double)

    If ValidationNumeric.IsNegative(value) Then
        ErrorManager.ThrowError _
            entitiesStockMovement.NegativeOutgoingAmount, _
            CLASS_NAME & ".OutgoingAmount", _
            CLASS_NAME & ": Jumlah barang keluar tidak boleh kurang dari 0."
    End If

    mOutgoingAmount = value

End Property

Public Property Get EndStock() As Double
    EndStock = mEndStock
End Property

Public Property Let EndStock(ByVal value As Double)

    If ValidationNumeric.IsNegative(value) Then
        ErrorManager.ThrowError _
            entitiesStockMovement.NegativeEndStock, _
            CLASS_NAME & ".EndStock", _
            CLASS_NAME & ": Stok akhir tidak boleh kurang dari 0."
    End If

    mEndStock = value

End Property

Public Property Get remarks() As String
    remarks = mRemarks
End Property

Public Property Let remarks(ByVal value As String)

    value = Trim$(value)

    mRemarks = value

End Property

Public Property Get personInCharge() As String
    personInCharge = mPersonInCharge
End Property

Public Property Let personInCharge(ByVal value As String)

    value = Trim$(value)

    If ValidationText.IsBlank(value) Then
        ErrorManager.ThrowError _
            entitiesStockMovement.EmptyPersonInCharge, _
            CLASS_NAME & ".PersonInCharge", _
            CLASS_NAME & ": Penanggung jawab tidak boleh kosong."
    End If

    mPersonInCharge = value

End Property
