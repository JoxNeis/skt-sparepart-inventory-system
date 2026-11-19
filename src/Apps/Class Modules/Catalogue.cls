Attribute VB_Name = "Catalogue"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' Error codes: entitiesCatalogue (ErrorCodesEntities.bas)

Private Const CLASS_NAME As String = "Catalogue"

Private mId As Long
Private mName As String
Private mAlias As String
Private mCode As String
Private mCategory As String
Private mUnitOfMeasurement As String

' =========================================================
' ID
' =========================================================
Public Property Get Id() As Long
    Id = mId
End Property

Public Property Let Id(ByVal value As Long)

    If ValidationNumeric.IsZeroOrNegative(value) Then
        ErrorManager.ThrowError _
            entitiesCatalogue.NegativeId, _
            CLASS_NAME & ".Id", _
            CLASS_NAME & ": Id harus lebih besar dari 0."
    End If

    mId = value

End Property

' =========================================================
' NAME
' =========================================================
Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(ByVal value As String)

    value = Trim$(value)

    If ValidationText.IsBlank(value) Then
        ErrorManager.ThrowError _
            entitiesCatalogue.EmptyName, _
            CLASS_NAME & ".Name", _
            CLASS_NAME & ": Nama barang tidak boleh berupa teks kosong."
    End If

    mName = value

End Property

' =========================================================
' ALIAS
' =========================================================
Public Property Get Alias() As String

    Alias = mAlias

End Property


Public Property Let Alias( _
    ByVal value As String)

    value = Trim$(value)

    mAlias = value

End Property

' =========================================================
' CODE
' =========================================================
Public Property Get Code() As String
    Code = mCode
End Property

Public Property Let Code(ByVal value As String)

    value = Trim$(value)

    mCode = value

End Property

' =========================================================
' CATEGORY
' =========================================================
Public Property Get Category() As String
    Category = mCategory
End Property

Public Property Let Category(ByVal value As String)

    value = Trim$(value)

    mCategory = value

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
            entitiesCatalogue.EmptyUnitOfMeasurement, _
            CLASS_NAME & ".UnitOfMeasurement", _
            CLASS_NAME & ": Satuan jumlah barang tidak boleh kosong."
    End If

    mUnitOfMeasurement = value

End Property
