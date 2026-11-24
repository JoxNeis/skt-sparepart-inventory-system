Attribute VB_Name = "MobileMaintenance"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' =========================================================
' MODEL - MOBILE MAINTENANCE
' Table      : mobile_maintenances
' Error codes: entitiesMobileMaintenance (ErrorCodesEntities.bas)
'
' Was: MobilePost
' Code and name are unique per plant, not globally.
' =========================================================

Private Const CLASS_NAME As String = "MobileMaintenance"

Private Const CODE_MAX_LENGTH As Long = 40
Private Const NAME_MAX_LENGTH As Long = 255

Private mId As Long
Private mManufacturingPlant As ManufacturingPlant
Private mCode As String
Private mName As String
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
            entitiesMobileMaintenance.NegativeId, _
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
            entitiesMobileMaintenance.EmptyManufacturingPlant, _
            CLASS_NAME & ".ManufacturingPlant", _
            CLASS_NAME & ": Plant tidak boleh kosong."
    End If

    Set mManufacturingPlant = value

End Property


' =========================================================
' CODE                                  [code] TEXT(40) NOT NULL
' =========================================================
Public Property Get Code() As String
    Code = mCode
End Property

Public Property Let Code(ByVal value As String)

    mCode = CheckText(value, "Code", "Kode mobile maintenance", _
        CODE_MAX_LENGTH, _
        entitiesMobileMaintenance.EmptyCode, _
        entitiesMobileMaintenance.CodeTooLong)

End Property


' =========================================================
' NAME                                  [name] TEXT(255) NOT NULL
' =========================================================
Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(ByVal value As String)

    mName = CheckText(value, "Name", "Nama mobile maintenance", _
        NAME_MAX_LENGTH, _
        entitiesMobileMaintenance.EmptyName, _
        entitiesMobileMaintenance.NameTooLong)

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

' "code - name", same as mobile_maintenance_list.display.
Public Property Get DisplayName() As String

    If Len(mCode) = 0 Then
        DisplayName = mName
    Else
        DisplayName = mCode & " - " & mName
    End If

End Property

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
