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
' Error codes: entitiesTransactionUpdateRequest (ErrorCodesEntities.bas)
'
' Approval workflow for editing an active transaction. The
' edit is stored as a new draft (proposed) transaction.
' Approving reverses the original and activates the proposal.
'
' Life:  PENDING  -> Verdict = ""
'        APPROVED -> Verdict = "APPROVED"
'        REJECTED -> Verdict = "REJECTED"
' Reviewer and ReviewedAt are set together with Verdict and
' are blank while the request is pending (call Validate).
' =========================================================

Private Const CLASS_NAME As String = "TransactionUpdateRequest"

Private Const VERDICT_APPROVED As String = "APPROVED"
Private Const VERDICT_REJECTED As String = "REJECTED"

Private Const USERNAME_MAX_LENGTH As Long = 64

Private mId As Long
Private mOriginalTransaction As Transaction
Private mProposedTransaction As Transaction
Private mReason As String
Private mRequestedBy As String
Private mRequestedAt As Date
Private mVerdict As String
Private mReviewer As String
Private mReviewedAt As Date
Private mDecisionNote As String


' =========================================================
' ID                                    [id]
' =========================================================
Public Property Get Id() As Long
    Id = mId
End Property

Public Property Let Id(ByVal value As Long)

    If ValidationNumeric.IsZeroOrNegative(value) Then
        ErrorManager.ThrowError _
            entitiesTransactionUpdateRequest.NegativeId, _
            CLASS_NAME & ".Id", _
            CLASS_NAME & ": Id harus lebih besar dari 0."
    End If

    mId = value

End Property


' =========================================================
' ORIGINAL TRANSACTION                  [original_transaction_id] NOT NULL
' =========================================================
Public Property Get OriginalTransaction() As Transaction
    Set OriginalTransaction = mOriginalTransaction
End Property

Public Property Set OriginalTransaction(ByVal value As Transaction)

    If value Is Nothing Then
        ErrorManager.ThrowError _
            entitiesTransactionUpdateRequest.EmptyOriginalTransaction, _
            CLASS_NAME & ".OriginalTransaction", _
            CLASS_NAME & ": Transaksi asal tidak boleh kosong."
    End If

    CheckPair value, mProposedTransaction, "OriginalTransaction"

    Set mOriginalTransaction = value

End Property


' =========================================================
' PROPOSED TRANSACTION                  [proposed_transaction_id] NOT NULL UNIQUE
' =========================================================
Public Property Get ProposedTransaction() As Transaction
    Set ProposedTransaction = mProposedTransaction
End Property

Public Property Set ProposedTransaction(ByVal value As Transaction)

    If value Is Nothing Then
        ErrorManager.ThrowError _
            entitiesTransactionUpdateRequest.EmptyProposedTransaction, _
            CLASS_NAME & ".ProposedTransaction", _
            CLASS_NAME & ": Transaksi usulan tidak boleh kosong."
    End If

    CheckPair mOriginalTransaction, value, "ProposedTransaction"

    Set mProposedTransaction = value

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

    If ValidationText.IsBlank(value) Then
        ErrorManager.ThrowError _
            entitiesTransactionUpdateRequest.EmptyRequestedBy, _
            CLASS_NAME & ".RequestedBy", _
            CLASS_NAME & ": Pemohon tidak boleh kosong."
    End If

    If ValidationText.IsLongerThan(value, USERNAME_MAX_LENGTH) Then
        ErrorManager.ThrowError _
            entitiesTransactionUpdateRequest.RequestedByTooLong, _
            CLASS_NAME & ".RequestedBy", _
            CLASS_NAME & ": Pemohon maksimal " & _
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
' "" (pending) / APPROVED / REJECTED
' =========================================================
Public Property Get Verdict() As String
    Verdict = mVerdict
End Property

Public Property Let Verdict(ByVal value As String)

    value = UCase$(Trim$(value))

    If Len(value) > 0 Then
        If Not ValidationText.IsInList(value, _
            Array(VERDICT_APPROVED, VERDICT_REJECTED)) Then

            ErrorManager.ThrowError _
                entitiesTransactionUpdateRequest.InvalidVerdict, _
                CLASS_NAME & ".Verdict", _
                CLASS_NAME & ": Keputusan harus APPROVED atau REJECTED."
        End If
    End If

    mVerdict = value

End Property


' =========================================================
' REVIEWER / REVIEWED AT                [reviewer] TEXT(64)
'                                       [reviewed_at]
' =========================================================
Public Property Get Reviewer() As String
    Reviewer = mReviewer
End Property

Public Property Let Reviewer(ByVal value As String)

    value = Trim$(value)

    If ValidationText.IsLongerThan(value, USERNAME_MAX_LENGTH) Then
        ErrorManager.ThrowError _
            entitiesTransactionUpdateRequest.ReviewerTooLong, _
            CLASS_NAME & ".Reviewer", _
            CLASS_NAME & ": Reviewer maksimal " & _
            USERNAME_MAX_LENGTH & " karakter."
    End If

    If Len(value) > 0 And Len(mRequestedBy) > 0 Then
        If ValidationText.IsSameText(value, mRequestedBy) Then
            ErrorManager.ThrowError _
                entitiesTransactionUpdateRequest.ReviewerIsRequester, _
                CLASS_NAME & ".Reviewer", _
                CLASS_NAME & ": Reviewer tidak boleh sama dengan pemohon."
        End If
    End If

    mReviewer = value

End Property

Public Property Get ReviewedAt() As Date
    ReviewedAt = mReviewedAt
End Property

Public Property Let ReviewedAt(ByVal value As Date)
    mReviewedAt = value
End Property


' =========================================================
' DECISION NOTE                         [decision_note] LONGTEXT
' =========================================================
Public Property Get DecisionNote() As String
    DecisionNote = mDecisionNote
End Property

Public Property Let DecisionNote(ByVal value As String)
    mDecisionNote = Trim$(value)
End Property


' =========================================================
' CALCULATED PROPERTIES
' =========================================================
Public Property Get IsPending() As Boolean
    IsPending = (Len(mVerdict) = 0)
End Property

Public Property Get IsApproved() As Boolean
    IsApproved = (mVerdict = VERDICT_APPROVED)
End Property

Public Property Get IsRejected() As Boolean
    IsRejected = (mVerdict = VERDICT_REJECTED)
End Property

' ORDER / RECEIVE / REQUEST, from the original transaction.
Public Property Get TransactionType() As String
    If Not mOriginalTransaction Is Nothing Then
        TransactionType = mOriginalTransaction.TransactionType
    End If
End Property


' =========================================================
' VALIDATE
' ck_transaction_update_requests_verdict: Reviewer and
' ReviewedAt are both blank while pending, and both set once
' there is a verdict. Call it before saving.
' =========================================================
Public Sub Validate()

    If IsPending Then

        If Len(mReviewer) > 0 Or mReviewedAt <> 0 Then
            ErrorManager.ThrowError _
                entitiesTransactionUpdateRequest.ReviewWithoutVerdict, _
                CLASS_NAME & ".Validate", _
                CLASS_NAME & _
                ": Reviewer dan tanggal review harus kosong selama belum ada keputusan."
        End If

    Else

        If Len(mReviewer) = 0 Or mReviewedAt = 0 Then
            ErrorManager.ThrowError _
                entitiesTransactionUpdateRequest.VerdictWithoutReview, _
                CLASS_NAME & ".Validate", _
                CLASS_NAME & _
                ": Keputusan harus disertai reviewer dan tanggal review."
        End If

    End If

End Sub


' =========================================================
' INTERNAL
' =========================================================

' The original and proposed transactions must be different
' rows of the same type and plant (transaction_update_request_create).
' Only checks what both objects have filled in.
Private Sub CheckPair( _
    ByVal original As Transaction, _
    ByVal proposed As Transaction, _
    ByVal propertyName As String)

    If original Is Nothing Or proposed Is Nothing Then Exit Sub

    If original.IsSameAs(proposed) Then
        ErrorManager.ThrowError _
            entitiesTransactionUpdateRequest.SameTransaction, _
            CLASS_NAME & "." & propertyName, _
            CLASS_NAME & _
            ": Transaksi usulan tidak boleh sama dengan transaksi asal."
    End If

    If Len(original.TransactionType) > 0 And _
       Len(proposed.TransactionType) > 0 Then

        If original.TransactionType <> proposed.TransactionType Then
            ErrorManager.ThrowError _
                entitiesTransactionUpdateRequest.TransactionTypeMismatch, _
                CLASS_NAME & "." & propertyName, _
                CLASS_NAME & _
                ": Tipe transaksi usulan harus sama dengan transaksi asal."
        End If

    End If

    If Not original.ManufacturingPlant Is Nothing And _
       Not proposed.ManufacturingPlant Is Nothing Then

        If original.ManufacturingPlant.Id > 0 And _
           proposed.ManufacturingPlant.Id > 0 Then

            If original.ManufacturingPlant.Id <> _
               proposed.ManufacturingPlant.Id Then

                ErrorManager.ThrowError _
                    entitiesTransactionUpdateRequest.ManufacturingPlantMismatch, _
                    CLASS_NAME & "." & propertyName, _
                    CLASS_NAME & _
                    ": Plant transaksi usulan harus sama dengan transaksi asal."
            End If

        End If

    End If

End Sub
