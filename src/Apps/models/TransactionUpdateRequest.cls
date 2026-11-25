Attribute VB_Name = "TransactionUpdateRequest"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' =========================================================
' MODEL - TRANSACTION UPDATE REQUEST
' Table      : transaction_update_requests
'              (+ the original / proposed document columns of
'              transaction_update_request_list)
' Error codes: entitiesTransactionUpdateRequest
'              (managerErrorCodesEntities.bas)
'
' A proposed edit of an active document. The edit itself is
' a DRAFT transaction (the proposed document); approving it
' reverses the original and activates the proposal.
'
' Life:  PENDING  -> Verdict = ""
'        APPROVED -> Verdict = "APPROVED"
'        REJECTED -> Verdict = "REJECTED"
'
' Original and Proposed hold the document columns of
' transaction_update_request_list without the old_ / new_
' prefix, as a Scripting.Dictionary:
'   uuid, document_date, quantity, ok_quantity,
'   not_ok_quantity, remarks, person_in_charge, order_uuid,
'   po_number, sparepart_id, sparepart_name,
'   sparepart_unit_of_measurement, vendor_name,
'   mobile_maintenance_name
' =========================================================

Private Const CLASS_NAME As String = "TransactionUpdateRequest"

Private Const VERDICT_APPROVED As String = "APPROVED"
Private Const VERDICT_REJECTED As String = "REJECTED"

Private Const USERNAME_MAX_LENGTH As Long = 64

Private mId As Long
Private mTransactionType As String
Private mReason As String
Private mRequestedBy As String
Private mRequestedAt As Date
Private mVerdict As String
Private mReviewer As String
Private mReviewedAt As Date
Private mDecisionNote As String
Private mManufacturingPlantName As String
Private mOriginal As Object
Private mProposed As Object


Private Sub Class_Initialize()
    Set mOriginal = CreateObject("Scripting.Dictionary")
    Set mProposed = CreateObject("Scripting.Dictionary")
    mOriginal.CompareMode = vbTextCompare
    mProposed.CompareMode = vbTextCompare
End Sub


' =========================================================
' ID                                    [id]
' =========================================================
Public Property Get Id() As Long
    Id = mId
End Property

Public Property Let Id(ByVal value As Long)

    If managerValidationNumeric.IsZeroOrNegative(value) Then
        managerError.ThrowError _
            entitiesTransactionUpdateRequest.NegativeId, _
            CLASS_NAME & ".Id", _
            CLASS_NAME & ": Id harus lebih besar dari 0."
    End If

    mId = value

End Property


' =========================================================
' TRANSACTION TYPE                      ORDER / RECEIVE / REQUEST
' (type of the original document)
' =========================================================
Public Property Get TransactionType() As String
    TransactionType = mTransactionType
End Property

Public Property Let TransactionType(ByVal value As String)

    value = UCase$(Trim$(value))

    If Not managerValidationText.IsInList(value, _
        Array("ORDER", "RECEIVE", "REQUEST", "MAINTENANCE")) Then

        managerError.ThrowError _
            entitiesTransactionUpdateRequest.TransactionTypeMismatch, _
            CLASS_NAME & ".TransactionType", _
            CLASS_NAME & ": Tipe transaksi tidak dikenal: " & value & "."
    End If

    mTransactionType = value

End Property


' =========================================================
' REASON                                [reason] LONGTEXT
' =========================================================
Public Property Get Reason() As String
    Reason = mReason
End Property

Public Property Let Reason(ByVal value As String)
    mReason = Trim$(value)
End Property


' =========================================================
' REQUESTED BY / AT                     [requested_by] TEXT(64) NOT NULL
'                                       [requested_at]
' =========================================================
Public Property Get RequestedBy() As String
    RequestedBy = mRequestedBy
End Property

Public Property Let RequestedBy(ByVal value As String)

    value = Trim$(value)

    If managerValidationText.IsBlank(value) Then
        managerError.ThrowError _
            entitiesTransactionUpdateRequest.EmptyRequestedBy, _
            CLASS_NAME & ".RequestedBy", _
            CLASS_NAME & ": Pengaju tidak boleh kosong."
    End If

    If managerValidationText.IsLongerThan(value, USERNAME_MAX_LENGTH) Then
        managerError.ThrowError _
            entitiesTransactionUpdateRequest.RequestedByTooLong, _
            CLASS_NAME & ".RequestedBy", _
            CLASS_NAME & ": Pengaju maksimal " & _
            USERNAME_MAX_LENGTH & " karakter."
    End If

    mRequestedBy = value

End Property

Public Property Get RequestedAt() As Date
    RequestedAt = mRequestedAt
End Property

Public Property Let RequestedAt(ByVal value As Date)
    mRequestedAt = value
End Property


' =========================================================
' VERDICT                               [verdict] TEXT(10)
' "" = pending, APPROVED, REJECTED
' =========================================================
Public Property Get Verdict() As String
    Verdict = mVerdict
End Property

Public Property Let Verdict(ByVal value As String)

    value = UCase$(Trim$(value))

    If Len(value) > 0 Then
        If Not managerValidationText.IsInList(value, _
            Array(VERDICT_APPROVED, VERDICT_REJECTED)) Then

            managerError.ThrowError _
                entitiesTransactionUpdateRequest.InvalidVerdict, _
                CLASS_NAME & ".Verdict", _
                CLASS_NAME & ": Verdict harus APPROVED atau REJECTED."
        End If
    End If

    mVerdict = value

End Property


' =========================================================
' REVIEWER / REVIEWED AT / DECISION NOTE
' =========================================================
Public Property Get Reviewer() As String
    Reviewer = mReviewer
End Property

Public Property Let Reviewer(ByVal value As String)

    value = Trim$(value)

    If managerValidationText.IsLongerThan(value, USERNAME_MAX_LENGTH) Then
        managerError.ThrowError _
            entitiesTransactionUpdateRequest.ReviewerTooLong, _
            CLASS_NAME & ".Reviewer", _
            CLASS_NAME & ": Reviewer maksimal " & _
            USERNAME_MAX_LENGTH & " karakter."
    End If

    mReviewer = value

End Property

Public Property Get ReviewedAt() As Date
    ReviewedAt = mReviewedAt
End Property

Public Property Let ReviewedAt(ByVal value As Date)
    mReviewedAt = value
End Property

Public Property Get DecisionNote() As String
    DecisionNote = mDecisionNote
End Property

Public Property Let DecisionNote(ByVal value As String)
    mDecisionNote = Trim$(value)
End Property


' =========================================================
' MANUFACTURING PLANT NAME              (plant of the original)
' =========================================================
Public Property Get ManufacturingPlantName() As String
    ManufacturingPlantName = mManufacturingPlantName
End Property

Public Property Let ManufacturingPlantName(ByVal value As String)
    mManufacturingPlantName = Trim$(value)
End Property


' =========================================================
' ORIGINAL / PROPOSED DOCUMENT
' =========================================================
Public Property Get Original() As Object
    Set Original = mOriginal
End Property

Public Property Get Proposed() As Object
    Set Proposed = mProposed
End Property

' Value of one original column (Null when not present).
Public Function OriginalValue(ByVal key As String) As Variant

    If mOriginal.Exists(key) Then
        OriginalValue = mOriginal(key)
    Else
        OriginalValue = Null
    End If

End Function

' Value of one proposed column (Null when not present).
Public Function ProposedValue(ByVal key As String) As Variant

    If mProposed.Exists(key) Then
        ProposedValue = mProposed(key)
    Else
        ProposedValue = Null
    End If

End Function


' =========================================================
' CALCULATED PROPERTIES
' =========================================================

Public Property Get OriginalUUID() As String
    OriginalUUID = coreHelper.TextOrEmpty(OriginalValue("uuid"))
End Property

Public Property Get ProposedUUID() As String
    ProposedUUID = coreHelper.TextOrEmpty(ProposedValue("uuid"))
End Property

Public Property Get OriginalSparepartId() As Long
    OriginalSparepartId = ValueAsLong(OriginalValue("sparepart_id"))
End Property

Public Property Get ProposedSparepartId() As Long
    ProposedSparepartId = ValueAsLong(ProposedValue("sparepart_id"))
End Property

Public Property Get IsPending() As Boolean
    IsPending = (Len(mVerdict) = 0)
End Property

Public Property Get IsApproved() As Boolean
    IsApproved = (mVerdict = VERDICT_APPROVED)
End Property

Public Property Get IsRejected() As Boolean
    IsRejected = (mVerdict = VERDICT_REJECTED)
End Property


' =========================================================
' INTERNAL
' =========================================================

Private Function ValueAsLong(ByVal value As Variant) As Long

    If IsNull(value) Or IsEmpty(value) Then
        ValueAsLong = 0
    ElseIf IsNumeric(value) Then
        ValueAsLong = CLng(value)
    Else
        ValueAsLong = 0
    End If

End Function
