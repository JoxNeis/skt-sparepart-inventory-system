Attribute VB_Name = "serviceActivityLog"
Option Explicit

' =========================================================
' SERVICE - ACTIVITY LOG
'
' Helper used by other services INSIDE their transaction.
' It receives the caller's connection and never begins,
' commits or rolls back by itself.
' =========================================================

Private Const MODULE_NAME As String = "SERVICE_ACTIVITY_LOG"

' =========================================================
' ERROR CODES (17100 - 17199)
' Usage: managerError.ThrowError serviceActivityLog.InsertFailed, ...
' =========================================================
Public Const EmptyActivityCode As Long = 17101
Public Const EmptyActivityUser As Long = 17102     ' user cannot be determined
Public Const InsertFailed As Long = 17103          ' 0 rows inserted


' Insert one activity log row on the caller's connection.
' user defaults to coreConfig.currentUser().
Public Sub WriteActivityLog( _
    ByVal conn As ADODB.connection, _
    ByVal activityCode As String, _
    ByVal detail As String, _
    Optional ByVal user As String = vbNullString)

    On Error GoTo ErrorHandler

    Dim affected As Long
    Dim errInfo As ErrorInfo

    activityCode = Trim$(activityCode)
    user = Trim$(user)

    If managerValidationText.IsBlank(user) Then
        user = Trim$(coreConfig.currentUser())
    End If

    If managerValidationText.IsBlank(activityCode) Then
        managerError.ThrowError _
            serviceActivityLog.EmptyActivityCode, _
            MODULE_NAME & ".WriteActivityLog", _
            "Activity code tidak boleh kosong."
    End If

    If managerValidationText.IsBlank(user) Then
        managerError.ThrowError _
            serviceActivityLog.EmptyActivityUser, _
            MODULE_NAME & ".WriteActivityLog", _
            "User activity log tidak dapat ditentukan."
    End If

    affected = daoActivityLog.CreateActivityLog( _
        conn, activityCode, detail, user)

    If managerValidationNumeric.IsLessThan(affected, 1) Then
        managerError.ThrowError _
            serviceActivityLog.InsertFailed, _
            MODULE_NAME & ".WriteActivityLog", _
            "Activity log tidak tersimpan (0 baris terpengaruh)."
    End If

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".WriteActivityLog", _
        "Gagal membuat activity log."

End Sub
