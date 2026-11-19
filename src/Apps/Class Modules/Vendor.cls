Attribute VB_Name = "Vendor"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' Error codes: entitiesVendor (ErrorCodesEntities.bas)

Private Const CLASS_NAME As String = "Vendor"

Private mId As Long
Private mName As String
Private mCompanyName As String
Private mEmail As String
Private mCode As String

Public Property Get Id() As Long
    Id = mId
End Property

Public Property Let Id(ByVal value As Long)
    If ValidationNumeric.IsZeroOrNegative(value) Then
        ErrorManager.ThrowError _
            entitiesVendor.NegativeId, _
            CLASS_NAME & ".Id", _
            CLASS_NAME & ": Id harus lebih besar dari 0"
    End If

    mId = value
End Property

Public Property Get Name() As String
    Name = mName
End Property
Public Property Let Name(ByVal value As String)
    value = Trim$(value)

    If ValidationText.IsBlank(value) Then
        ErrorManager.ThrowError _
            entitiesVendor.EmptyName, _
            CLASS_NAME & ".Name", _
            CLASS_NAME & ": Nama barang tidak boleh teks kosong."
    End If
    mName = value
End Property

Public Property Get Code() As String
    Code = mCode
End Property
Public Property Let Code(ByVal value As String)
    value = Trim$(value)

    If ValidationText.IsBlank(value) Then
        ErrorManager.ThrowError _
            entitiesVendor.EmptyCode, _
            CLASS_NAME & ".Code", _
            CLASS_NAME & ": Kode tidak dapat kosong."
    End If

    mCode = value
End Property



Public Property Get CompanyName() As String
    CompanyName = mCompanyName
End Property
Public Property Let CompanyName(ByVal value As String)
    value = Trim$(value)

    If ValidationText.IsBlank(value) Then
        ErrorManager.ThrowError _
            entitiesVendor.EmptyCompanyName, _
            CLASS_NAME & ".CompanyName", _
            CLASS_NAME & ": Nama perusahaan tidak boleh teks kosong."
    End If
    mCompanyName = value
End Property

Public Property Get Email() As String
    Email = mEmail
End Property
Public Property Let Email(ByVal value As String)
    value = LCase$(Trim$(value))
    If ValidationText.IsBlank(value) Then
        ErrorManager.ThrowError _
            entitiesVendor.EmptyEmail, _
            CLASS_NAME & ".Email", _
            CLASS_NAME & ": Email tidak boleh teks kosong."
    End If
    mEmail = value
End Property
