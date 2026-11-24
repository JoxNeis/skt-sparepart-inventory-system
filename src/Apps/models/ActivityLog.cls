Attribute VB_Name = "ActivityLog"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' =========================================================
' MODEL - ACTIVITY LOG
' Table      : activity_logs
' Error codes: entitiesActivityLog (ErrorCodesEntities.bas)
'
' Append-only audit trail. [username] is deliberately not a
' User object: a log row must outlive the user row.
' =========================================================

Private Const CLASS_NAME As String = "ActivityLog"

Private Const ACTIVITY_CODE_MAX_LENGTH As Long = 50
Private Const USERNAME_MAX_LENGTH As Long = 64

Private mId As Long
Private mActivityCode As String
Private mDetail As String
Private mJson As String
Private mUsername As String
Private mCreatedAt As Date


' =========================================================
' ID                                    [id]
' =========================================================
Public Property Get Id() As Long
    Id = mId
End Property

Public Property Let Id(ByVal value As Long)

    If ValidationNumeric.IsZeroOrNegative(value) Then
        ErrorManager.ThrowError _
            entitiesActivityLog.NegativeId, _
            CLASS_NAME & ".Id", _
            CLASS_NAME & ": Id harus lebih besar dari 0."
    End If

    mId = value

End Property


' =========================================================
' ACTIVITY CODE                         [activity_code] TEXT(50) NOT NULL
' =========================================================
Public Property Get ActivityCode() As String
    ActivityCode = mActivityCode
End Property

Public Property Let ActivityCode(ByVal value As String)

    mActivityCode = UCase$(CheckText(value, "ActivityCode", _
        "Kode aktivitas", ACTIVITY_CODE_MAX_LENGTH, _
        entitiesActivityLog.EmptyActivityCode, _
        entitiesActivityLog.ActivityCodeTooLong))

End Property


' =========================================================
' DETAIL                                [detail] LONGTEXT
' =========================================================
Public Property Get Detail() As String
    Detail = mDetail
End Property

Public Property Let Detail(ByVal value As String)
    mDetail = Trim$(value)
End Property


' =========================================================
' JSON                                  [json] LONGTEXT
' =========================================================
Public Property Get Json() As String
    Json = mJson
End Property

Public Property Let Json(ByVal value As String)
    mJson = Trim$(value)
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
        entitiesActivityLog.EmptyUsername, _
        entitiesActivityLog.UsernameTooLong)

End Property


' =========================================================
' CREATED AT                            [created_at]
' Set by the database.
' =========================================================
Public Property Get CreatedAt() As Date
    CreatedAt = mCreatedAt
End Property

Public Property Let CreatedAt(ByVal value As Date)
    mCreatedAt = value
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
