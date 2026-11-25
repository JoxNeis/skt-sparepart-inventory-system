Attribute VB_Name = "Order"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' =========================================================
' MODEL - ORDER
' Table      : orders (+ progress columns of vw_order_progress)
' Error codes: entitiesOrder (managerErrorCodesEntities.bas)
'
' Detail of an ORDER transaction - spareparts ordered from a
' vendor. The document number (UUID) and the plant live on
' the Transaction header.
'
' Renamed from the old model:
'   Catalogue -> Sparepart       Amount -> Quantity (Long)
'   poID      -> PoNumber        Plant  -> Transaction.ManufacturingPlant
'   UUID      -> Transaction.UUID
' =========================================================

Private Const CLASS_NAME As String = "Order"

Private Const TRANSACTION_TYPE As String = "ORDER"

Private Const PO_NUMBER_MAX_LENGTH As Long = 50
Private Const PERSON_IN_CHARGE_MAX_LENGTH As Long = 255

Private mId As Long
Private mTransaction As Transaction
Private mSparepart As Sparepart
Private mVendor As Vendor
Private mPoNumber As String
Private mOrderDate As Date
Private mQuantity As Long
Private mRemarks As String
Private mSystemRemarks As String
Private mPersonInCharge As String

' vw_order_progress
Private mOkQuantityTotal As Long
Private mReceivedQuantityTotal As Long
Private mRemainingQuantity As Long
Private mStatus As String
Private mEstimatedArrival As Date


' =========================================================
' ID                                    [id]
' =========================================================
Public Property Get Id() As Long
    Id = mId
End Property

Public Property Let Id(ByVal value As Long)

    If managerValidationNumeric.IsZeroOrNegative(value) Then
        managerError.ThrowError _
            entitiesOrder.NegativeId, _
            CLASS_NAME & ".Id", _
            CLASS_NAME & ": Id harus lebih besar dari 0."
    End If

    mId = value

End Property


' =========================================================
' TRANSACTION                           [transaction_id] NOT NULL UNIQUE
' Must be an ORDER header (when its type is set).
' =========================================================
Public Property Get Transaction() As Transaction
    Set Transaction = mTransaction
End Property

Public Property Set Transaction(ByVal value As Transaction)

    If value Is Nothing Then
        managerError.ThrowError _
            entitiesOrder.EmptyTransaction, _
            CLASS_NAME & ".Transaction", _
            CLASS_NAME & ": Transaksi tidak boleh kosong."
    End If

    If Len(value.TransactionType) > 0 Then
        If value.TransactionType <> TRANSACTION_TYPE Then
            managerError.ThrowError _
                entitiesOrder.WrongTransactionType, _
                CLASS_NAME & ".Transaction", _
                CLASS_NAME & ": Tipe transaksi harus " & _
                TRANSACTION_TYPE & "."
        End If
    End If

    Set mTransaction = value

End Property


' =========================================================
' SPAREPART                             [sparepart_id] NOT NULL
' =========================================================
Public Property Get Sparepart() As Sparepart
    Set Sparepart = mSparepart
End Property

Public Property Set Sparepart(ByVal value As Sparepart)

    If value Is Nothing Then
        managerError.ThrowError _
            entitiesOrder.EmptySparepart, _
            CLASS_NAME & ".Sparepart", _
            CLASS_NAME & ": Sparepart tidak boleh kosong."
    End If

    Set mSparepart = value

End Property


' =========================================================
' VENDOR                                [vendor_id] (optional)
' Nothing = NULL.
' =========================================================
Public Property Get Vendor() As Vendor
    Set Vendor = mVendor
End Property

Public Property Set Vendor(ByVal value As Vendor)
    Set mVendor = value
End Property

Public Property Get HasVendor() As Boolean
    HasVendor = Not (mVendor Is Nothing)
End Property


' =========================================================
' PO NUMBER                             [po_number] TEXT(50) (optional)
' Free-text SAP PO number, filled in later
' (order_set_po_number).
' =========================================================
Public Property Get PoNumber() As String
    PoNumber = mPoNumber
End Property

Public Property Let PoNumber(ByVal value As String)

    value = Trim$(value)

    If managerValidationText.IsLongerThan(value, PO_NUMBER_MAX_LENGTH) Then
        managerError.ThrowError _
            entitiesOrder.PoNumberTooLong, _
            CLASS_NAME & ".PoNumber", _
            CLASS_NAME & ": Nomor PO maksimal " & _
            PO_NUMBER_MAX_LENGTH & " karakter."
    End If

    mPoNumber = value

End Property

Public Property Get HasPoNumber() As Boolean
    HasPoNumber = (Len(mPoNumber) > 0)
End Property


' =========================================================
' ORDER DATE                            [order_date] NOT NULL
' =========================================================
Public Property Get OrderDate() As Date
    OrderDate = mOrderDate
End Property

Public Property Let OrderDate(ByVal value As Date)

    If managerValidationNumeric.IsZeroOrNegative(value) Then
        managerError.ThrowError _
            entitiesOrder.EmptyOrderDate, _
            CLASS_NAME & ".OrderDate", _
            CLASS_NAME & ": Tanggal order tidak boleh kosong."
    End If

    If DateValue(value) > VBA.Date Then
        managerError.ThrowError _
            entitiesOrder.FutureOrderDate, _
            CLASS_NAME & ".OrderDate", _
            CLASS_NAME & ": Tanggal order tidak boleh berada di masa depan."
    End If

    mOrderDate = value

End Property


' =========================================================
' QUANTITY                              [quantity] LONG NOT NULL >= 0
' 0 = deleted through an approved update. A NEW order must
' be above 0: serviceOrder checks that.
' =========================================================
Public Property Get Quantity() As Long
    Quantity = mQuantity
End Property

Public Property Let Quantity(ByVal value As Long)

    If managerValidationNumeric.IsNegative(value) Then
        managerError.ThrowError _
            entitiesOrder.InvalidQuantity, _
            CLASS_NAME & ".Quantity", _
            CLASS_NAME & ": Jumlah order tidak boleh negatif."
    End If

    mQuantity = value

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

    If managerValidationText.IsBlank(value) Then
        managerError.ThrowError _
            entitiesOrder.EmptyPersonInCharge, _
            CLASS_NAME & ".PersonInCharge", _
            CLASS_NAME & ": Person in charge tidak boleh kosong."
    End If

    If managerValidationText.IsLongerThan(value, PERSON_IN_CHARGE_MAX_LENGTH) Then
        managerError.ThrowError _
            entitiesOrder.PersonInChargeTooLong, _
            CLASS_NAME & ".PersonInCharge", _
            CLASS_NAME & ": Person in charge maksimal " & _
            PERSON_IN_CHARGE_MAX_LENGTH & " karakter."
    End If

    mPersonInCharge = value

End Property


' =========================================================
' PROGRESS (read from vw_order_progress, not stored)
'   ok_quantity_total, received_quantity_total,
'   remaining_quantity, status
' =========================================================
Public Property Get OkQuantityTotal() As Long
    OkQuantityTotal = mOkQuantityTotal
End Property

Public Property Let OkQuantityTotal(ByVal value As Long)

    If managerValidationNumeric.IsNegative(value) Then
        managerError.ThrowError _
            entitiesOrder.NegativeOkQuantityTotal, _
            CLASS_NAME & ".OkQuantityTotal", _
            CLASS_NAME & ": Total barang OK tidak boleh negatif."
    End If

    mOkQuantityTotal = value

End Property

Public Property Get ReceivedQuantityTotal() As Long
    ReceivedQuantityTotal = mReceivedQuantityTotal
End Property

Public Property Let ReceivedQuantityTotal(ByVal value As Long)

    If managerValidationNumeric.IsNegative(value) Then
        managerError.ThrowError _
            entitiesOrder.NegativeReceivedQuantityTotal, _
            CLASS_NAME & ".ReceivedQuantityTotal", _
            CLASS_NAME & ": Total penerimaan tidak boleh negatif."
    End If

    mReceivedQuantityTotal = value

End Property

Public Property Get RemainingQuantity() As Long
    RemainingQuantity = mRemainingQuantity
End Property

Public Property Let RemainingQuantity(ByVal value As Long)

    If managerValidationNumeric.IsNegative(value) Then
        managerError.ThrowError _
            entitiesOrder.NegativeRemainingQuantity, _
            CLASS_NAME & ".RemainingQuantity", _
            CLASS_NAME & ": Sisa jumlah order tidak boleh negatif."
    End If

    If mQuantity > 0 Then
        If value > mQuantity Then
            managerError.ThrowError _
                entitiesOrder.RemainingQuantityExceedsQuantity, _
                CLASS_NAME & ".RemainingQuantity", _
                CLASS_NAME & _
                ": Sisa jumlah order tidak boleh melebihi jumlah order."
        End If
    End If

    mRemainingQuantity = value

End Property

' BELUM DATANG / DATANG SEBAGIAN / DATANG SEMUA
Public Property Get Status() As String
    Status = mStatus
End Property

Public Property Let Status(ByVal value As String)

    value = UCase$(Trim$(value))

    If Not managerValidationText.IsInList(value, _
        Array("BELUM DATANG", "DATANG SEBAGIAN", "DATANG SEMUA")) Then

        managerError.ThrowError _
            entitiesOrder.InvalidStatus, _
            CLASS_NAME & ".Status", _
            CLASS_NAME & _
            ": Status harus BELUM DATANG, DATANG SEBAGIAN, atau DATANG SEMUA."
    End If

    mStatus = value

End Property


' =========================================================
' ESTIMATED ARRIVAL (read from order_list, not stored)
' order_date + lead time of the plant stock row; 0 = unknown
' (lead time 0).
' =========================================================
Public Property Get EstimatedArrival() As Date
    EstimatedArrival = mEstimatedArrival
End Property

Public Property Let EstimatedArrival(ByVal value As Date)
    mEstimatedArrival = value
End Property


' =========================================================
' CALCULATED PROPERTIES
' =========================================================

' Document number of the ORDER header ("" when not set).
Public Property Get UUID() As String
    If Not mTransaction Is Nothing Then UUID = mTransaction.UUID
End Property

' Plant of the ORDER header (Nothing when not set).
Public Property Get ManufacturingPlant() As ManufacturingPlant
    If Not mTransaction Is Nothing Then
        Set ManufacturingPlant = mTransaction.ManufacturingPlant
    End If
End Property

Public Property Get UnitOfMeasurement() As String
    If Not mSparepart Is Nothing Then
        UnitOfMeasurement = mSparepart.UnitOfMeasurement
    End If
End Property

Public Property Get IsFullyReceived() As Boolean
    IsFullyReceived = (mStatus = "DATANG SEMUA")
End Property
