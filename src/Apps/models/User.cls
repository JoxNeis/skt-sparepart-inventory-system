Attribute VB_Name = "User"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' =========================================================
' MODEL - USER
' Table      : users
' Error codes: entitiesUser (ErrorCodesEntities.bas)
'
' [username] is the Windows login and the value stored in
' every *_by column. [tier] drives approvals: the approver
' needs a higher tier than the requester.
' =========================================================

Private Const CLASS_NAME As String = "User"

Private Const USERNAME_MAX_LENGTH As Long = 64
Private Const NAME_MAX_LENGTH As Long = 100
Private Const JOBDESK_MAX_LENGTH As Long = 255

Private mId As Long
Private mManufacturingPlant As ManufacturingPlant
Private mUsername As String
Private mFirstName As String
Private mMiddleName As String
Private mLastName As String
Private mJobdesk As String
Private mTier As Long
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
            entitiesUser.NegativeId, _
            CLASS_NAME & ".Id", _
            CLASS_NAME & ": Id harus lebih besar dari 0."
    End If

    mId = value

End Property


' =========================================================
' MANUFACTURING PLANT                   [manufacturing_plant_id] NOT NULL
' =========================================================
Public Property Get ManufacturingPlant() As ManufacturingPlant
    Set ManufacturingPlant = mManufacturingPlant
End Property

Public Property Set ManufacturingPlant(ByVal value As ManufacturingPlant)

    If value Is Nothing Then
        ErrorManager.ThrowError _
            entitiesUser.EmptyManufacturingPlant, _
            CLASS_NAME & ".ManufacturingPlant", _
            CLASS_NAME & ": Plant tidak boleh kosong."
    End If

    Set mManufacturingPlant = value

End Property


' =========================================================
' USERNAME                              [username] TEXT(64) NOT NULL
' =========================================================
Public Property Get Username() As String
    Username = mUsername
End Property

Public Property Let Username(ByVal value As String)

    mUsername = CheckText(value, "Username", "Username", _
        USERNAME_MAX_LENGTH, _
        entitiesUser.EmptyUsername, _
        entitiesUser.UsernameTooLong)

End Property


' =========================================================
' FIRST NAME                            [first_name] TEXT(100) NOT NULL
' =========================================================
Public Property Get FirstName() As String
    FirstName = mFirstName
End Property

Public Property Let FirstName(ByVal value As String)

    mFirstName = CheckText(value, "FirstName", "Nama depan", _
        NAME_MAX_LENGTH, _
        entitiesUser.EmptyFirstName, _
        entitiesUser.FirstNameTooLong)

End Property


' =========================================================
' MIDDLE NAME                           [middle_name] TEXT(100)
' =========================================================
Public Property Get MiddleName() As String
    MiddleName = mMiddleName
End Property

Public Property Let MiddleName(ByVal value As String)

    mMiddleName = CheckText(value, "MiddleName", "Nama tengah", _
        NAME_MAX_LENGTH, 0, entitiesUser.MiddleNameTooLong)

End Property


' =========================================================
' LAST NAME                             [last_name] TEXT(100)
' =========================================================
Public Property Get LastName() As String
    LastName = mLastName
End Property

Public Property Let LastName(ByVal value As String)

    mLastName = CheckText(value, "LastName", "Nama belakang", _
        NAME_MAX_LENGTH, 0, entitiesUser.LastNameTooLong)

End Property


' =========================================================
' JOBDESK                               [jobdesk] TEXT(255)
' =========================================================
Public Property Get Jobdesk() As String
    Jobdesk = mJobdesk
End Property

Public Property Let Jobdesk(ByVal value As String)

    mJobdesk = CheckText(value, "Jobdesk", "Jobdesk", _
        JOBDESK_MAX_LENGTH, 0, entitiesUser.JobdeskTooLong)

End Property


' =========================================================
' TIER                                  [tier] LONG NOT NULL DEFAULT 0
' =========================================================
Public Property Get Tier() As Long
    Tier = mTier
End Property

Public Property Let Tier(ByVal value As Long)

    If ValidationNumeric.IsNegative(value) Then
        ErrorManager.ThrowError _
            entitiesUser.NegativeTier, _
            CLASS_NAME & ".Tier", _
            CLASS_NAME & ": Tier tidak boleh negatif."
    End If

    mTier = value

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

' First, middle and last name joined by single spaces.
Public Property Get FullName() As String

    Dim result As String

    result = mFirstName

    If Len(mMiddleName) > 0 Then result = result & " " & mMiddleName
    If Len(mLastName) > 0 Then result = result & " " & mLastName

    FullName = Trim$(result)

End Property

Public Property Get IsArchived() As Boolean
    IsArchived = (mDeletedAt <> 0)
End Property

' Same rule as user_check_can_approve: the approver needs a
' strictly higher tier than the requester.
Public Function CanApprove(ByVal requester As User) As Boolean

    If requester Is Nothing Then Exit Function
    If IsArchived Then Exit Function

    CanApprove = (mTier > requester.Tier)

End Function


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
