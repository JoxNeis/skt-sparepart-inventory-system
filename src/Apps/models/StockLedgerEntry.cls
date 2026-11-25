Attribute VB_Name = "StockLedgerEntry"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' =========================================================
' MODEL - STOCK LEDGER ENTRY (read only, no table)
' Source     : stock_ledger_list_by_month (vw_stock_ledger)
' Error codes: entitiesStockLedgerEntry (managerErrorCodesEntities.bas)
'
' One line of the stock card (kartu stok):
'   IN  = verified OK quantity of an active receive
'   OUT = quantity of an active request
' BeginningStock and EndStock are not query columns; the
' service fills them while running the balance.
'
' Was: StockMovement
'   MovementDate (Long)      -> MovementDate (Date)
'   IncomingReferenceNumber  -> IncomingReference
'   OutgoingReferenceNumber  -> OutgoingReference
'   incomingAmount / outgoingAmount -> IncomingQuantity /
'                               OutgoingQuantity (Long)
' =========================================================

Private Const CLASS_NAME As String = "StockLedgerEntry"

Private mMovementDate As Date
Private mBeginningStock As Long
Private mIncomingReference As String
Private mIncomingFrom As String
Private mIncomingQuantity As Long
Private mOutgoingReference As String
Private mOutgoingTo As String
Private mOutgoingQuantity As Long
Private mEndStock As Long
Private mRemarks As String
Private mPersonInCharge As String


' =========================================================
' MOVEMENT DATE                         movement_date
' =========================================================
Public Property Get MovementDate() As Date
    MovementDate = mMovementDate
End Property

Public Property Let MovementDate(ByVal value As Date)

    If managerValidationNumeric.IsZeroOrNegative(value) Then
        managerError.ThrowError _
            entitiesStockLedgerEntry.EmptyMovementDate, _
            CLASS_NAME & ".MovementDate", _
            CLASS_NAME & ": Tanggal pergerakan stok tidak boleh kosong."
    End If

    mMovementDate = value

End Property


' =========================================================
' BEGINNING STOCK                       (running balance)
' =========================================================
Public Property Get BeginningStock() As Long
    BeginningStock = mBeginningStock
End Property

Public Property Let BeginningStock(ByVal value As Long)

    If managerValidationNumeric.IsNegative(value) Then
        managerError.ThrowError _
            entitiesStockLedgerEntry.NegativeBeginningStock, _
            CLASS_NAME & ".BeginningStock", _
            CLASS_NAME & ": Stok awal tidak boleh negatif."
    End If

    mBeginningStock = value

End Property


' =========================================================
' INCOMING                              incoming_reference
'                                       incoming_from
'                                       incoming_quantity
' =========================================================
Public Property Get IncomingReference() As String
    IncomingReference = mIncomingReference
End Property

Public Property Let IncomingReference(ByVal value As String)
    mIncomingReference = Trim$(value)
End Property

Public Property Get IncomingFrom() As String
    IncomingFrom = mIncomingFrom
End Property

Public Property Let IncomingFrom(ByVal value As String)
    mIncomingFrom = Trim$(value)
End Property

Public Property Get IncomingQuantity() As Long
    IncomingQuantity = mIncomingQuantity
End Property

Public Property Let IncomingQuantity(ByVal value As Long)

    If managerValidationNumeric.IsNegative(value) Then
        managerError.ThrowError _
            entitiesStockLedgerEntry.NegativeIncomingQuantity, _
            CLASS_NAME & ".IncomingQuantity", _
            CLASS_NAME & ": Jumlah barang masuk tidak boleh negatif."
    End If

    mIncomingQuantity = value

End Property


' =========================================================
' OUTGOING                              outgoing_reference
'                                       outgoing_to
'                                       outgoing_quantity
' =========================================================
Public Property Get OutgoingReference() As String
    OutgoingReference = mOutgoingReference
End Property

Public Property Let OutgoingReference(ByVal value As String)
    mOutgoingReference = Trim$(value)
End Property

Public Property Get OutgoingTo() As String
    OutgoingTo = mOutgoingTo
End Property

Public Property Let OutgoingTo(ByVal value As String)
    mOutgoingTo = Trim$(value)
End Property

Public Property Get OutgoingQuantity() As Long
    OutgoingQuantity = mOutgoingQuantity
End Property

Public Property Let OutgoingQuantity(ByVal value As Long)

    If managerValidationNumeric.IsNegative(value) Then
        managerError.ThrowError _
            entitiesStockLedgerEntry.NegativeOutgoingQuantity, _
            CLASS_NAME & ".OutgoingQuantity", _
            CLASS_NAME & ": Jumlah barang keluar tidak boleh negatif."
    End If

    mOutgoingQuantity = value

End Property


' =========================================================
' END STOCK                             (running balance)
' =========================================================
Public Property Get EndStock() As Long
    EndStock = mEndStock
End Property

Public Property Let EndStock(ByVal value As Long)

    If managerValidationNumeric.IsNegative(value) Then
        managerError.ThrowError _
            entitiesStockLedgerEntry.NegativeEndStock, _
            CLASS_NAME & ".EndStock", _
            CLASS_NAME & ": Stok akhir tidak boleh negatif."
    End If

    mEndStock = value

End Property


' =========================================================
' REMARKS / PERSON IN CHARGE            remarks, person_in_charge
' Both may be NULL in the source rows.
' =========================================================
Public Property Get Remarks() As String
    Remarks = mRemarks
End Property

Public Property Let Remarks(ByVal value As String)
    mRemarks = Trim$(value)
End Property

Public Property Get PersonInCharge() As String
    PersonInCharge = mPersonInCharge
End Property

Public Property Let PersonInCharge(ByVal value As String)
    mPersonInCharge = Trim$(value)
End Property


' =========================================================
' CALCULATED PROPERTIES
' =========================================================

' IN / OUT, as vw_stock_ledger.direction.
Public Property Get Direction() As String

    If Len(mOutgoingReference) > 0 Or mOutgoingQuantity > 0 Then
        Direction = "OUT"
    Else
        Direction = "IN"
    End If

End Property

' IncomingQuantity - OutgoingQuantity.
Public Property Get NetQuantity() As Long
    NetQuantity = mIncomingQuantity - mOutgoingQuantity
End Property

' Sets EndStock = BeginningStock + NetQuantity.
Public Sub ApplyBeginningStock(ByVal value As Long)

    BeginningStock = value
    EndStock = value + NetQuantity

End Sub
