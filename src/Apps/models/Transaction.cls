Attribute VB_Name = "Transaction"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' =========================================================
' MODEL - TRANSACTION (document header)
' Table      : transactions
' Error codes: entitiesTransaction (managerErrorCodesEntities.bas)
'
' Shared by every document (ORDER, RECEIVE, REQUEST,
' MAINTENANCE). Each header has exactly one detail (1:1):
' Order, Receive, Request or MaintenanceWork.
'
' Life:  DRAFT    -> ActivatedAt = 0
'        ACTIVE   -> ActivatedAt set, ReversedAt = 0
'        REVERSED -> ReversedAt set
' A transaction is never edited or deleted. An edit is a new
' proposed transaction (see TransactionUpdateRequest).
'
' [uuid] is the document number, e.g. ORD-20261124-001
' (transaction_get_next_uuid); it is not an RFC 4122 UUID.
' =========================================================

Private Const CLASS_NAME As String = "Transaction"

' Class modules cannot hold Public Const, so these stay
' private; callers use the same literal text.
Private Const TYPE_ORDER As String = "ORDER"
Private Const TYPE_RECEIVE As String = "RECEIVE"
Private Const TYPE_REQUEST As String = "REQUEST"
Private Const TYPE_MAINTENANCE As String = "MAINTENANCE"

Private Const STATUS_DRAFT As String = "DRAFT"
Private Const STATUS_ACTIVE As String = "ACTIVE"
Private Const STATUS_REVERSED As String = "REVERSED"

Private Const UUID_MAX_LENGTH As Long = 36
Private Const USERNAME_MAX_LENGTH As Long = 64

Private mId As Long
Private mUUID As String
Private mTransactionType As String
Private mManufacturingPlant As ManufacturingPlant
Private mCreatedBy As String
Private mCreatedAt As Date
Private mActivatedAt As Date
Private mReversedBy As String
Private mReversedAt As Date
Private mSignedAt As Date
Private mDigitalSignature As String
Private mSyncedBy As String
Private mSyncedAt As Date


' =========================================================
' ID                                    [id]
' =========================================================
Public Property Get Id() As Long
    Id = mId
End Property

Public Property Let Id(ByVal value As Long)

    If managerValidationNumeric.IsZeroOrNegative(value) Then
        managerError.ThrowError _
            entitiesTransaction.NegativeId, _
            CLASS_NAME & ".Id", _
            CLASS_NAME & ": Id harus lebih besar dari 0."
    End If

    mId = value

End Property


' =========================================================
' UUID                                  [uuid] TEXT(36) NOT NULL
' =========================================================
Public Property Get UUID() As String
    UUID = mUUID
End Property

Public Property Let UUID(ByVal value As String)

    mUUID = UCase$(CheckText(value, "UUID", "Nomor dokumen", _
        UUID_MAX_LENGTH, _
        entitiesTransaction.EmptyUUID, _
        entitiesTransaction.UUIDTooLong))

End Property


' =========================================================
' TRANSACTION TYPE                      [transaction_type] TEXT(15) NOT NULL
' ORDER / RECEIVE / REQUEST / MAINTENANCE
' =========================================================
Public Property Get TransactionType() As String
    TransactionType = mTransactionType
End Property

Public Property Let TransactionType(ByVal value As String)

    value = UCase$(Trim$(value))

    If Not managerValidationText.IsInList(value, _
        Array(TYPE_ORDER, TYPE_RECEIVE, TYPE_REQUEST, TYPE_MAINTENANCE)) Then

        managerError.ThrowError _
            entitiesTransaction.InvalidTransactionType, _
            CLASS_NAME & ".TransactionType", _
            CLASS_NAME & _
            ": Tipe transaksi harus ORDER, RECEIVE, REQUEST, atau " & _
            "MAINTENANCE."
    End If

    mTransactionType = value

End Property


' =========================================================
' MANUFACTURING PLANT                   [manufacturing_plant_id] NOT NULL
' =========================================================
Public Property Get ManufacturingPlant() As ManufacturingPlant
    Set ManufacturingPlant = mManufacturingPlant
End Property

Public Property Set ManufacturingPlant(ByVal value As ManufacturingPlant)

    If value Is Nothing Then
        managerError.ThrowError _
            entitiesTransaction.EmptyManufacturingPlant, _
            CLASS_NAME & ".ManufacturingPlant", _
            CLASS_NAME & ": Plant tidak boleh kosong."
    End If

    Set mManufacturingPlant = value

End Property


' =========================================================
' CREATED BY / AT                       [created_by] TEXT(64) NOT NULL
'                                       [created_at]
' =========================================================
Public Property Get CreatedBy() As String
    CreatedBy = mCreatedBy
End Property

Public Property Let CreatedBy(ByVal value As String)

    mCreatedBy = CheckText(value, "CreatedBy", "Pembuat transaksi", _
        USERNAME_MAX_LENGTH, _
        entitiesTransaction.EmptyCreatedBy, _
        entitiesTransaction.CreatedByTooLong)

End Property

Public Property Get CreatedAt() As Date
    CreatedAt = mCreatedAt
End Property

Public Property Let CreatedAt(ByVal value As Date)
    mCreatedAt = value
End Property


' =========================================================
' ACTIVATED AT                          [activated_at]
' 0 = NULL = still a draft.
' =========================================================
Public Property Get ActivatedAt() As Date
    ActivatedAt = mActivatedAt
End Property

Public Property Let ActivatedAt(ByVal value As Date)
    mActivatedAt = value
End Property


' =========================================================
' REVERSED BY / AT                      [reversed_by] TEXT(64)
'                                       [reversed_at]
' =========================================================
Public Property Get ReversedBy() As String
    ReversedBy = mReversedBy
End Property

Public Property Let ReversedBy(ByVal value As String)

    mReversedBy = CheckText(value, "ReversedBy", "Pembatal transaksi", _
        USERNAME_MAX_LENGTH, 0, entitiesTransaction.ReversedByTooLong)

End Property

Public Property Get ReversedAt() As Date
    ReversedAt = mReversedAt
End Property

Public Property Let ReversedAt(ByVal value As Date)
    mReversedAt = value
End Property


' =========================================================
' SIGNATURE                             [signed_at]
'                                       [digital_signature] LONGTEXT
' =========================================================
Public Property Get SignedAt() As Date
    SignedAt = mSignedAt
End Property

Public Property Let SignedAt(ByVal value As Date)
    mSignedAt = value
End Property

Public Property Get DigitalSignature() As String
    DigitalSignature = mDigitalSignature
End Property

Public Property Let DigitalSignature(ByVal value As String)
    mDigitalSignature = Trim$(value)
End Property


' =========================================================
' SYNCED BY / AT                        [synced_by] TEXT(64)
'                                       [synced_at]
' =========================================================
Public Property Get SyncedBy() As String
    SyncedBy = mSyncedBy
End Property

Public Property Let SyncedBy(ByVal value As String)

    mSyncedBy = CheckText(value, "SyncedBy", "Penyinkron transaksi", _
        USERNAME_MAX_LENGTH, 0, entitiesTransaction.SyncedByTooLong)

End Property

Public Property Get SyncedAt() As Date
    SyncedAt = mSyncedAt
End Property

Public Property Let SyncedAt(ByVal value As Date)
    mSyncedAt = value
End Property


' =========================================================
' CALCULATED PROPERTIES
' =========================================================

' DRAFT / ACTIVE / REVERSED
Public Property Get Status() As String

    If mReversedAt <> 0 Then
        Status = STATUS_REVERSED
    ElseIf mActivatedAt <> 0 Then
        Status = STATUS_ACTIVE
    Else
        Status = STATUS_DRAFT
    End If

End Property

Public Property Get IsDraft() As Boolean
    IsDraft = (Status = STATUS_DRAFT)
End Property

Public Property Get IsActive() As Boolean
    IsActive = (Status = STATUS_ACTIVE)
End Property

Public Property Get IsReversed() As Boolean
    IsReversed = (Status = STATUS_REVERSED)
End Property

Public Property Get IsSigned() As Boolean
    IsSigned = (mSignedAt <> 0)
End Property

Public Property Get IsSynced() As Boolean
    IsSynced = (mSyncedAt <> 0)
End Property

' True when both objects stand for the same row: same Id when
' both have one, otherwise the same UUID.
Public Function IsSameAs(ByVal other As Transaction) As Boolean

    If other Is Nothing Then Exit Function

    If mId > 0 And other.Id > 0 Then
        IsSameAs = (mId = other.Id)
    ElseIf Len(mUUID) > 0 And Len(other.UUID) > 0 Then
        IsSameAs = managerValidationText.IsSameText(mUUID, other.UUID)
    End If

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
        If managerValidationText.IsBlank(value) Then
            managerError.ThrowError _
                emptyCode, _
                CLASS_NAME & "." & propertyName, _
                CLASS_NAME & ": " & label & " tidak boleh kosong."
        End If
    End If

    If managerValidationText.IsLongerThan(value, maxLength) Then
        managerError.ThrowError _
            tooLongCode, _
            CLASS_NAME & "." & propertyName, _
            CLASS_NAME & ": " & label & " maksimal " & _
            maxLength & " karakter."
    End If

    CheckText = value

End Function
