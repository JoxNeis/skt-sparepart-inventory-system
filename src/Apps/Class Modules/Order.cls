Attribute VB_Name = "Order"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' Error codes: entitiesOrder (ErrorCodesEntities.bas)

Private Const CLASS_NAME As String = "Order"

Private mId As Long
Private mUUID As String
Private mPoId As String
Private mPlant As Plant
Private mOrderDate As Date
Private mCatalogue As Catalogue
Private mAmount As Double
Private mStatus As String
Private mVendor As Vendor
Private mRemarks As String
Private mPIC As String

' =========================================================
' ID
' =========================================================
Public Property Get Id() As Long
    Id = mId
End Property

Public Property Let Id(ByVal value As Long)

    If ValidationNumeric.IsZeroOrNegative(value) Then
        ErrorManager.ThrowError _
            entitiesOrder.NegativeId, _
            CLASS_NAME & ".Id", _
            CLASS_NAME & ": ID harus lebih besar dari 0."
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
            entitiesOrder.EmptyUUID, _
            CLASS_NAME & ".UUID", _
            CLASS_NAME & ": UUID tidak boleh kosong."
    End If

    mUUID = value

End Property

' =========================================================
' PO ID
' =========================================================
Public Property Get poID() As String
    poID = mPoId
End Property

Public Property Let poID(ByVal value As String)

    value = Trim$(value)

    If ValidationText.IsBlank(value) Then
        ErrorManager.ThrowError _
            entitiesOrder.EmptyPoId, _
            CLASS_NAME & ".PoId", _
            CLASS_NAME & ": Nomor PO tidak boleh kosong."
    End If

    mPoId = value

End Property

' =========================================================
' PLANT
' =========================================================
Public Property Get Plant() As Plant
    Set Plant = mPlant
End Property

Public Property Set Plant(ByVal value As Plant)

    If value Is Nothing Then
        ErrorManager.ThrowError _
            entitiesOrder.EmptyPlant, _
            CLASS_NAME & ".Plant", _
            CLASS_NAME & ": Plant tidak boleh kosong."
    End If

    Set mPlant = value

End Property

' =========================================================
' ORDER DATE
' =========================================================
Public Property Get orderDate() As Date
    orderDate = mOrderDate
End Property

Public Property Let orderDate(ByVal value As Date)

    If ValidationNumeric.IsZero(value) Then
        ErrorManager.ThrowError _
            entitiesOrder.EmptyOrderDate, _
            CLASS_NAME & ".OrderDate", _
            CLASS_NAME & ": Tanggal order tidak boleh kosong."
    End If

    If DateValue(value) > VBA.Date Then
        ErrorManager.ThrowError _
            entitiesOrder.FutureOrderDate, _
            CLASS_NAME & ".OrderDate", _
            CLASS_NAME & ": Tanggal order tidak boleh berada di masa depan."
    End If

    mOrderDate = value

End Property

' =========================================================
' CATALOGUE
' =========================================================
Public Property Get Catalogue() As Catalogue
    Set Catalogue = mCatalogue
End Property

Public Property Set Catalogue(ByVal value As Catalogue)

    If value Is Nothing Then
        ErrorManager.ThrowError _
            entitiesOrder.EmptyCatalogue, _
            CLASS_NAME & ".Catalogue", _
            CLASS_NAME & ": Catalogue tidak boleh kosong."
    End If

    Set mCatalogue = value

End Property

' =========================================================
' AMOUNT
' =========================================================
Public Property Get Amount() As Double
    Amount = mAmount
End Property

Public Property Let Amount(ByVal value As Double)

    If ValidationNumeric.IsZeroOrNegative(value) Then
        ErrorManager.ThrowError _
            entitiesOrder.InvalidAmount, _
            CLASS_NAME & ".Amount", _
            CLASS_NAME & ": Jumlah order harus lebih besar dari 0."
    End If

    mAmount = value

End Property

' =========================================================
' STATUS
' =========================================================
Public Property Get Status() As String
    Status = mStatus
End Property

Public Property Let Status(ByVal value As String)

    value = UCase$(Trim$(value))

    If Not ValidationText.IsInList(value, _
        Array("DATANG", "DATANG SEBAGIAN", "BELUM DATANG")) Then

        ErrorManager.ThrowError _
            entitiesOrder.InvalidStatus, _
            CLASS_NAME & ".Status", _
            CLASS_NAME & _
            ": Status harus DATANG, DATANG SEBAGIAN, atau BELUM DATANG."
    End If

    mStatus = value

End Property

' =========================================================
' VENDOR
' =========================================================
Public Property Get Vendor() As Vendor
    Set Vendor = mVendor
End Property

Public Property Set Vendor(ByVal value As Vendor)

    If value Is Nothing Then
        ErrorManager.ThrowError _
            entitiesOrder.EmptyVendor, _
            CLASS_NAME & ".Vendor", _
            CLASS_NAME & ": Vendor tidak boleh kosong."
    End If

    Set mVendor = value

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
' PERSON IN CHARGE
' =========================================================
Public Property Get personInCharge() As String
    personInCharge = mPIC
End Property

Public Property Let personInCharge(ByVal value As String)

    value = Trim$(value)

    If ValidationText.IsBlank(value) Then
        ErrorManager.ThrowError _
            entitiesOrder.EmptyPersonInCharge, _
            CLASS_NAME & ".PersonInCharge", _
            CLASS_NAME & ": Person in charge tidak boleh kosong."
    End If

    mPIC = value

End Property
