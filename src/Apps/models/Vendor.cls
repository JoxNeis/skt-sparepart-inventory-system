Attribute VB_Name = "Vendor"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' =========================================================
' MODEL - VENDOR
' Table      : vendors
' Error codes: entitiesVendor (ErrorCodesEntities.bas)
'
' [company_name] and [email] are optional in the table, so
' they may be left blank here too.
' =========================================================

Private Const CLASS_NAME As String = "Vendor"

Private Const CODE_MAX_LENGTH As Long = 40
Private Const NAME_MAX_LENGTH As Long = 255
Private Const COMPANY_NAME_MAX_LENGTH As Long = 255
Private Const EMAIL_MAX_LENGTH As Long = 255

Private mId As Long
Private mCode As String
Private mName As String
Private mCompanyName As String
Private mEmail As String
Private mCreatedAt As Date
Private mUpdatedAt As Date
Private mDeletedAt As Date


' =========================================================
' ID                                    [id]
' =========================================================
Public Property Get Id() As Long
    Id = mId
End Property

Public Property Let Id(ByVal value As Long)

    If ValidationNumeric.IsZeroOrNegative(value) Then
        ErrorManager.ThrowError _
            entitiesVendor.NegativeId, _
            CLASS_NAME & ".Id", _
            CLASS_NAME & ": Id harus lebih besar dari 0."
    End If

    mId = value

End Property


' =========================================================
' CODE                                  [code] TEXT(40) NOT NULL
' =========================================================
Public Property Get Code() As String
    Code = mCode
End Property

Public Property Let Code(ByVal value As String)

    mCode = CheckText(value, "Code", "Kode vendor", CODE_MAX_LENGTH, _
        entitiesVendor.EmptyCode, _
        entitiesVendor.CodeTooLong)

End Property


' =========================================================
' NAME                                  [name] TEXT(255) NOT NULL
' =========================================================
Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(ByVal value As String)

    mName = CheckText(value, "Name", "Nama vendor", NAME_MAX_LENGTH, _
        entitiesVendor.EmptyName, _
        entitiesVendor.NameTooLong)

End Property


' =========================================================
' COMPANY NAME                          [company_name] TEXT(255)
' =========================================================
Public Property Get CompanyName() As String
    CompanyName = mCompanyName
End Property

Public Property Let CompanyName(ByVal value As String)

    mCompanyName = CheckText(value, "CompanyName", "Nama perusahaan", _
        COMPANY_NAME_MAX_LENGTH, 0, entitiesVendor.CompanyNameTooLong)

End Property


' =========================================================
' EMAIL                                 [email] TEXT(255)
' =========================================================
Public Property Get Email() As String
    Email = mEmail
End Property

Public Property Let Email(ByVal value As String)

    value = LCase$(CheckText(value, "Email", "Email", _
        EMAIL_MAX_LENGTH, 0, entitiesVendor.EmailTooLong))

    If Len(value) > 0 Then
        If Not (value Like "?*@?*.?*") Then
            ErrorManager.ThrowError _
                entitiesVendor.InvalidEmail, _
                CLASS_NAME & ".Email", _
                CLASS_NAME & ": Format email tidak valid."
        End If
    End If

    mEmail = value

End Property


' =========================================================
' TIMESTAMPS        [created_at] [updated_at] [deleted_at]
' Set by the database; 0 = NULL.
' =========================================================
Public Property Get CreatedAt() As Date
    CreatedAt = mCreatedAt
End Property

Public Property Let CreatedAt(ByVal value As Date)
    mCreatedAt = value
End Property

Public Property Get UpdatedAt() As Date
    UpdatedAt = mUpdatedAt
End Property

Public Property Let UpdatedAt(ByVal value As Date)
    mUpdatedAt = value
End Property

Public Property Get DeletedAt() As Date
    DeletedAt = mDeletedAt
End Property

Public Property Let DeletedAt(ByVal value As Date)
    mDeletedAt = value
End Property


' =========================================================
' CALCULATED PROPERTIES
' =========================================================
Public Property Get IsArchived() As Boolean
    IsArchived = (mDeletedAt <> 0)
End Property


' =========================================================
' INTERNAL
' =========================================================

' Trims value. Raises emptyCode when it is blank (pass 0 for
' an optional column) and tooLongCode when it is longer than
' the column.
Private Function CheckText( _
    ByVal value As String, _
    ByVal propertyName As String, _
    ByVal label As String, _
    ByVal maxLength As Long, _
    ByVal emptyCode As Long, _
    ByVal tooLongCode As Long) As String

    value = Trim$(value)

    If emptyCode <> 0 Then
        If ValidationText.IsBlank(value) Then
            ErrorManager.ThrowError _
                emptyCode, _
                CLASS_NAME & "." & propertyName, _
                CLASS_NAME & ": " & label & " tidak boleh kosong."
        End If
    End If

    If ValidationText.IsLongerThan(value, maxLength) Then
        ErrorManager.ThrowError _
            tooLongCode, _
            CLASS_NAME & "." & propertyName, _
            CLASS_NAME & ": " & label & " maksimal " & _
            maxLength & " karakter."
    End If

    CheckText = value

End Function
