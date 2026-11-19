Attribute VB_Name = "MobilePost"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' Error codes: entitiesMobilePost (ErrorCodesEntities.bas)

Private Const CLASS_NAME As String = "MobilePost"

Private mId As Long
Private mPlant As Plant
Private mName As String
Private mCode As String

Public Property Get Id() As Long
    Id = mId
End Property
Public Property Let Id(ByVal value As Long)

    If ValidationNumeric.IsZeroOrNegative(value) Then
        ErrorManager.ThrowError _
            entitiesMobilePost.NegativeId, _
            CLASS_NAME & ".Id", _
            CLASS_NAME & ": ID harus lebih besar dari 0."
    End If

    mId = value

End Property

Public Property Get Plant() As Plant
    Set Plant = mPlant
End Property

Public Property Set Plant(ByVal value As Plant)

    If value Is Nothing Then
        ErrorManager.ThrowError _
            entitiesMobilePost.EmptyPlant, _
            CLASS_NAME & ".Plant", _
            CLASS_NAME & ": Plant tidak boleh kosong."
    End If

    Set mPlant = value

End Property

Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(ByVal value As String)

    value = Trim$(value)

    If ValidationText.IsBlank(value) Then
        ErrorManager.ThrowError _
            entitiesMobilePost.EmptyName, _
            CLASS_NAME & ".Name", _
            CLASS_NAME & ": Nama tidak boleh kosong."
    End If

    mName = value

End Property


' =========================================================
' CODE
' =========================================================
Public Property Get Code() As String

    Code = mCode

End Property


Public Property Let Code( _
    ByVal value As String)

    value = Trim$(value)

    If ValidationText.IsBlank(value) Then
        ErrorManager.ThrowError _
            entitiesMobilePost.EmptyCode, _
            CLASS_NAME & ".Code", _
            CLASS_NAME & ": Kode mobile post tidak boleh kosong."
    End If

    mCode = value

End Property
