Attribute VB_Name = "Sparepart"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' =========================================================
' MODEL - SPAREPART
' Table      : spareparts
' Error codes: entitiesSparepart (ErrorCodesEntities.bas)
'
' Was: Catalogue
' [display] is a calculated field in the table ("code - name");
' here it is the read-only DisplayName property.
' =========================================================

Private Const CLASS_NAME As String = "Sparepart"

Private Const CODE_MAX_LENGTH As Long = 40
Private Const NAME_MAX_LENGTH As Long = 255
Private Const CATEGORY_MAX_LENGTH As Long = 50
Private Const APPLICABLE_MAX_LENGTH As Long = 255
Private Const UNIT_OF_MEASUREMENT_MAX_LENGTH As Long = 15
Private Const ALIAS_MAX_LENGTH As Long = 255
Private Const STATUS_MAX_LENGTH As Long = 50

Private mId As Long
Private mCode As String
Private mName As String
Private mCategory As String
Private mApplicable As String
Private mUnitOfMeasurement As String
Private mAlias As String
Private mStatus As String
Private mCreatedAt As Date
Private mUpdatedAt As Date
Private mDeletedAt As Date
Private mLastSyncedAt As Date


' =========================================================
' ID                                    [id]
' =========================================================
Public Property Get Id() As Long
    Id = mId
End Property

Public Property Let Id(ByVal value As Long)

    If ValidationNumeric.IsZeroOrNegative(value) Then
        ErrorManager.ThrowError _
            entitiesSparepart.NegativeId, _
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

    mCode = CheckText(value, "Code", "Kode barang", CODE_MAX_LENGTH, _
        entitiesSparepart.EmptyCode, _
        entitiesSparepart.CodeTooLong)

End Property


' =========================================================
' NAME                                  [name] TEXT(255) NOT NULL
' =========================================================
Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(ByVal value As String)

    mName = CheckText(value, "Name", "Nama barang", NAME_MAX_LENGTH, _
        entitiesSparepart.EmptyName, _
        entitiesSparepart.NameTooLong)

End Property


' =========================================================
' CATEGORY                              [category] TEXT(50)
' =========================================================
Public Property Get Category() As String
    Category = mCategory
End Property

Public Property Let Category(ByVal value As String)

    mCategory = CheckText(value, "Category", "Kategori", _
        CATEGORY_MAX_LENGTH, 0, entitiesSparepart.CategoryTooLong)

End Property


' =========================================================
' APPLICABLE                            [applicable] TEXT(255)
' The units / machines the sparepart fits.
' =========================================================
Public Property Get Applicable() As String
    Applicable = mApplicable
End Property

Public Property Let Applicable(ByVal value As String)

    mApplicable = CheckText(value, "Applicable", "Applicable", _
        APPLICABLE_MAX_LENGTH, 0, entitiesSparepart.ApplicableTooLong)

End Property


' =========================================================
' UNIT OF MEASUREMENT                   [unit_of_measurement] TEXT(15) NOT NULL
' =========================================================
Public Property Get UnitOfMeasurement() As String
    UnitOfMeasurement = mUnitOfMeasurement
End Property

Public Property Let UnitOfMeasurement(ByVal value As String)

    mUnitOfMeasurement = CheckText(value, "UnitOfMeasurement", _
        "Satuan barang", UNIT_OF_MEASUREMENT_MAX_LENGTH, _
        entitiesSparepart.EmptyUnitOfMeasurement, _
        entitiesSparepart.UnitOfMeasurementTooLong)

End Property


' =========================================================
' ALIAS                                 [alias] TEXT(255) NOT NULL
' The one other name a sparepart is known by.
' =========================================================
Public Property Get Alias() As String
    Alias = mAlias
End Property

Public Property Let Alias(ByVal value As String)

    mAlias = CheckText(value, "Alias", "Alias barang", ALIAS_MAX_LENGTH, _
        entitiesSparepart.EmptyAlias, _
        entitiesSparepart.AliasTooLong)

End Property


' =========================================================
' STATUS                                [status] TEXT(50)
' =========================================================
Public Property Get Status() As String
    Status = mStatus
End Property

Public Property Let Status(ByVal value As String)

    mStatus = UCase$(CheckText(value, "Status", "Status", _
        STATUS_MAX_LENGTH, 0, entitiesSparepart.StatusTooLong))

End Property


' =========================================================
' TIMESTAMPS        [created_at] [updated_at] [deleted_at]
'                   [last_synced_at]
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

Public Property Get LastSyncedAt() As Date
    LastSyncedAt = mLastSyncedAt
End Property

Public Property Let LastSyncedAt(ByVal value As Date)
    mLastSyncedAt = value
End Property


' =========================================================
' CALCULATED PROPERTIES
' =========================================================

' "code - name", same as spareparts.display.
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
