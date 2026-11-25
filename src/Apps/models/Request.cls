Attribute VB_Name = "Request"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' =========================================================
' MODEL - REQUEST
' Table      : requests
' Error codes: entitiesRequest (managerErrorCodesEntities.bas)
'
' Detail of a REQUEST transaction - spareparts issued from
' the plant store to a mobile maintenance. Leaves stock on
' RequestDate.
'
' Renamed from the old model:
'   MobilePost        -> MobileMaintenance
'   Catalogue         -> Sparepart
'   orderDate         -> RequestDate
'   Amount            -> Quantity (Long)
'   UUID              -> Transaction.UUID
'   unitOfMeasurement -> Sparepart.UnitOfMeasurement
'   ArriveDate        -> removed (no column)
' =========================================================

Private Const CLASS_NAME As String = "Request"

Private Const TRANSACTION_TYPE As String = "REQUEST"

Private Const PERSON_IN_CHARGE_MAX_LENGTH As Long = 255

Private mId As Long
Private mTransaction As Transaction
Private mMobileMaintenance As MobileMaintenance
Private mSparepart As Sparepart
Private mRequestDate As Date
Private mQuantity As Long
Private mRemarks As String
Private mSystemRemarks As String
Private mPersonInCharge As String


' =========================================================
' ID                                    [id]
' =========================================================
Public Property Get Id() As Long
    Id = mId
End Property

Public Property Let Id(ByVal value As Long)

    If managerValidationNumeric.IsZeroOrNegative(value) Then
        managerError.ThrowError _
            entitiesRequest.NegativeId, _
            CLASS_NAME & ".Id", _
            CLASS_NAME & ": Id harus lebih besar dari 0."
    End If

    mId = value

End Property


' =========================================================
' TRANSACTION                           [transaction_id] NOT NULL UNIQUE
' Must be a REQUEST header (when its type is set).
' =========================================================
Public Property Get Transaction() As Transaction
    Set Transaction = mTransaction
End Property

Public Property Set Transaction(ByVal value As Transaction)

    If value Is Nothing Then
        managerError.ThrowError _
            entitiesRequest.EmptyTransaction, _
            CLASS_NAME & ".Transaction", _
            CLASS_NAME & ": Transaksi tidak boleh kosong."
    End If

    If Len(value.TransactionType) > 0 Then
        If value.TransactionType <> TRANSACTION_TYPE Then
            managerError.ThrowError _
                entitiesRequest.WrongTransactionType, _
                CLASS_NAME & ".Transaction", _
                CLASS_NAME & ": Tipe transaksi harus " & _
                TRANSACTION_TYPE & "."
        End If
    End If

    Set mTransaction = value

End Property


' =========================================================
' MOBILE MAINTENANCE                    [mobile_maintenance_id] NOT NULL
' =========================================================
Public Property Get MobileMaintenance() As MobileMaintenance
    Set MobileMaintenance = mMobileMaintenance
End Property

Public Property Set MobileMaintenance(ByVal value As MobileMaintenance)

    If value Is Nothing Then
        managerError.ThrowError _
            entitiesRequest.EmptyMobileMaintenance, _
            CLASS_NAME & ".MobileMaintenance", _
            CLASS_NAME & ": Mobile maintenance tidak boleh kosong."
    End If

    Set mMobileMaintenance = value

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
            entitiesRequest.EmptySparepart, _
            CLASS_NAME & ".Sparepart", _
            CLASS_NAME & ": Sparepart tidak boleh kosong."
    End If

    Set mSparepart = value

End Property


' =========================================================
' REQUEST DATE                          [request_date] NOT NULL
' =========================================================
Public Property Get RequestDate() As Date
    RequestDate = mRequestDate
End Property

Public Property Let RequestDate(ByVal value As Date)

    If managerValidationNumeric.IsZeroOrNegative(value) Then
        managerError.ThrowError _
            entitiesRequest.EmptyRequestDate, _
            CLASS_NAME & ".RequestDate", _
            CLASS_NAME & ": Tanggal request tidak boleh kosong."
    End If

    If DateValue(value) > VBA.Date Then
        managerError.ThrowError _
            entitiesRequest.FutureRequestDate, _
            CLASS_NAME & ".RequestDate", _
            CLASS_NAME & ": Tanggal request tidak boleh berada di masa depan."
    End If

    mRequestDate = value

End Property


' =========================================================
' QUANTITY                              [quantity] LONG NOT NULL >= 0
' 0 = deleted through an approved update. A NEW request must
' be above 0: serviceRequest checks that.
' =========================================================
Public Property Get Quantity() As Long
    Quantity = mQuantity
End Property

Public Property Let Quantity(ByVal value As Long)

    If managerValidationNumeric.IsNegative(value) Then
        managerError.ThrowError _
            entitiesRequest.InvalidQuantity, _
            CLASS_NAME & ".Quantity", _
            CLASS_NAME & ": Jumlah barang tidak boleh negatif."
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
            entitiesRequest.EmptyPersonInCharge, _
            CLASS_NAME & ".PersonInCharge", _
            CLASS_NAME & ": Person in charge tidak boleh kosong."
    End If

    If managerValidationText.IsLongerThan(value, PERSON_IN_CHARGE_MAX_LENGTH) Then
        managerError.ThrowError _
            entitiesRequest.PersonInChargeTooLong, _
            CLASS_NAME & ".PersonInCharge", _
            CLASS_NAME & ": Person in charge maksimal " & _
            PERSON_IN_CHARGE_MAX_LENGTH & " karakter."
    End If

    mPersonInCharge = value

End Property


' =========================================================
' CALCULATED PROPERTIES
' =========================================================

' Document number of the REQUEST header ("" when not set).
Public Property Get UUID() As String
    If Not mTransaction Is Nothing Then UUID = mTransaction.UUID
End Property

' Plant of the REQUEST header (Nothing when not set).
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
