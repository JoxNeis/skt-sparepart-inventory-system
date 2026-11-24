Attribute VB_Name = "ManufacturingPlant"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' =========================================================
' MODEL - MANUFACTURING PLANT
' Table      : manufacturing_plants
' Error codes: entitiesManufacturingPlant (ErrorCodesEntities.bas)
'
' Was: Plant
' =========================================================

Private Const CLASS_NAME As String = "ManufacturingPlant"

Private Const CODE_MAX_LENGTH As Long = 15
Private Const NAME_MAX_LENGTH As Long = 255
Private Const ADDRESS_MAX_LENGTH As Long = 255
Private Const TYPE_MAX_LENGTH As Long = 50
Private Const REGION_MAX_LENGTH As Long = 100

Private mId As Long
Private mCode As String
Private mName As String
Private mAddress As String
Private mPlantType As String
Private mRegion As String
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
            entitiesManufacturingPlant.NegativeId, _
            CLASS_NAME & ".Id", _
            CLASS_NAME & ": Id harus lebih besar dari 0."
    End If

    mId = value

End Property


' =========================================================
' CODE                                  [code] TEXT(15) NOT NULL
' =========================================================
Public Property Get Code() As String
    Code = mCode
End Property

Public Property Let Code(ByVal value As String)

    mCode = CheckText(value, "Code", "Kode plant", CODE_MAX_LENGTH, _
        entitiesManufacturingPlant.EmptyCode, _
        entitiesManufacturingPlant.CodeTooLong)

End Property


' =========================================================
' NAME                                  [name] TEXT(255) NOT NULL
' =========================================================
Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(ByVal value As String)

    mName = CheckText(value, "Name", "Nama plant", NAME_MAX_LENGTH, _
        entitiesManufacturingPlant.EmptyName, _
        entitiesManufacturingPlant.NameTooLong)

End Property


' =========================================================
' ADDRESS                               [address] TEXT(255)
' =========================================================
Public Property Get Address() As String
    Address = mAddress
End Property

Public Property Let Address(ByVal value As String)

    mAddress = CheckText(value, "Address", "Alamat", ADDRESS_MAX_LENGTH, _
        0, entitiesManufacturingPlant.AddressTooLong)

End Property


' =========================================================
' PLANT TYPE                            [type] TEXT(50)
' "Type" is a VBA keyword, so the property is PlantType.
' Free text, e.g. PLANT / WAREHOUSE.
' =========================================================
Public Property Get PlantType() As String
    PlantType = mPlantType
End Property

Public Property Let PlantType(ByVal value As String)

    mPlantType = UCase$(CheckText(value, "PlantType", "Tipe plant", _
        TYPE_MAX_LENGTH, 0, entitiesManufacturingPlant.TypeTooLong))

End Property


' =========================================================
' REGION                                [region] TEXT(100)
' =========================================================
Public Property Get Region() As String
    Region = mRegion
End Property

Public Property Let Region(ByVal value As String)

    mRegion = UCase$(CheckText(value, "Region", "Region", _
        REGION_MAX_LENGTH, 0, entitiesManufacturingPlant.RegionTooLong))

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

' "code - name", same as manufacturing_plant_list.display_name.
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
