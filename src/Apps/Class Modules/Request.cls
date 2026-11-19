Attribute VB_Name = "Request"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' Error codes: entitiesRequest (ErrorCodesEntities.bas)

Private Const CLASS_NAME As String = "REQUEST"

Private mId As Long
Private mUUID As String
Private mMobilePost As MobilePost
Private mOrderDate As Date
Private mCatalogue As Catalogue
Private mAmount As Double
Private mUnitOfMeasurement As String
Private mArriveDate As Date
Private mRemarks As String
Private mSystemRemarks As String
Private mPersonInCharge As String


' =========================================================
' ID
' =========================================================
Public Property Get Id() As Long
    Id = mId
End Property

Public Property Let Id(ByVal value As Long)

    If ValidationNumeric.IsNegative(value) Then
        ErrorManager.ThrowError _
            entitiesRequest.NegativeId, _
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
            entitiesRequest.EmptyUUID, _
            CLASS_NAME & ".UUID", _
            CLASS_NAME & ": UUID tidak boleh kosong."
    End If

    mUUID = value

End Property


' =========================================================
' MOBILE POST
' =========================================================
Public Property Get MobilePost() As MobilePost
    Set MobilePost = mMobilePost
End Property

Public Property Set MobilePost(ByVal value As MobilePost)

    If value Is Nothing Then
        ErrorManager.ThrowError _
            entitiesRequest.EmptyMobilePost, _
            CLASS_NAME & ".MobilePost", _
            CLASS_NAME & ": Mobile post tidak boleh kosong."
    End If

    Set mMobilePost = value

End Property


' =========================================================
' ORDER DATE
' =========================================================
Public Property Get orderDate() As Date
    orderDate = mOrderDate
End Property

Public Property Let orderDate(ByVal value As Date)

    If ValidationNumeric.IsZeroOrNegative(value) Then
        ErrorManager.ThrowError _
            entitiesRequest.EmptyOrderDate, _
            CLASS_NAME & ".OrderDate", _
            CLASS_NAME & ": Tanggal order tidak boleh kosong."
    End If

    If DateValue(value) > VBA.Date Then
        ErrorManager.ThrowError _
            entitiesRequest.FutureOrderDate, _
            CLASS_NAME & ".OrderDate", _
            CLASS_NAME & ": Tanggal order tidak boleh berada di masa depan."
    End If

    If mArriveDate <> 0 Then
        If DateValue(value) > DateValue(mArriveDate) Then
            ErrorManager.ThrowError _
                entitiesRequest.OrderDateAfterArriveDate, _
                CLASS_NAME & ".OrderDate", _
                CLASS_NAME & _
                ": Tanggal order tidak boleh setelah tanggal kedatangan."
        End If
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
            entitiesRequest.EmptyCatalogue, _
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
            entitiesRequest.InvalidAmount, _
            CLASS_NAME & ".Amount", _
            CLASS_NAME & ": Jumlah barang harus lebih besar dari nol."
    End If

    mAmount = value

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
            entitiesRequest.EmptyUnitOfMeasurement, _
            CLASS_NAME & ".UnitOfMeasurement", _
            CLASS_NAME & ": Satuan jumlah barang tidak boleh kosong."
    End If

    mUnitOfMeasurement = value

End Property


' =========================================================
' ARRIVE DATE
' =========================================================
Public Property Get ArriveDate() As Date
    ArriveDate = mArriveDate
End Property

Public Property Let ArriveDate(ByVal value As Date)

    If ValidationNumeric.IsZeroOrNegative(value) Then
        ErrorManager.ThrowError _
            entitiesRequest.EmptyArriveDate, _
            CLASS_NAME & ".ArriveDate", _
            CLASS_NAME & ": Tanggal kedatangan tidak boleh kosong."
    End If

    If DateValue(value) > VBA.Date Then
        ErrorManager.ThrowError _
            entitiesRequest.FutureArriveDate, _
            CLASS_NAME & ".ArriveDate", _
            CLASS_NAME & ": Tanggal kedatangan tidak boleh berada di masa depan."
    End If

    If mOrderDate <> 0 Then
        If DateValue(value) < DateValue(mOrderDate) Then
            ErrorManager.ThrowError _
                entitiesRequest.ArriveDateBeforeOrderDate, _
                CLASS_NAME & ".ArriveDate", _
                CLASS_NAME & _
                ": Tanggal kedatangan tidak boleh sebelum tanggal order."
        End If
    End If

    mArriveDate = value

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
    personInCharge = mPersonInCharge
End Property

Public Property Let personInCharge(ByVal value As String)

    value = Trim$(value)

    If ValidationText.IsBlank(value) Then
        ErrorManager.ThrowError _
            entitiesRequest.EmptyPersonInCharge, _
            CLASS_NAME & ".PersonInCharge", _
            CLASS_NAME & ": Person in charge tidak boleh kosong."
    End If

    mPersonInCharge = value

End Property
