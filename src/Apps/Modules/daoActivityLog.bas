Attribute VB_Name = "daoActivityLog"
Option Explicit

Private Const MODULE_NAME As String = "DAO_ACTIVITY_LOG"

' =========================================================
' ERROR CODES (12100 - 12199)
' Usage: ErrorManager.ThrowError daoActivityLog.InsertFailed, ...
' =========================================================
Public Const EmptyActivityCode As Long = 12101
Public Const EmptyConnection As Long = 12102
Public Const InsertFailed As Long = 12103

Public Sub createActivityLog( _
    ByVal activityCode As String, _
    ByVal detail As String, _
    ByVal user As String, _
    ByVal conn As ADODB.connection)

    On Error GoTo ErrorHandler

    Dim cmd As ADODB.Command
    Dim affected As Long
    Dim detailSize As Long
    Dim errInfo As ErrorInfo

    If ValidationText.IsBlank(activityCode) Then
        ErrorManager.ThrowError _
            daoActivityLog.EmptyActivityCode, _
            MODULE_NAME & ".createActivityLog", _
            "Activity code tidak boleh kosong."
    End If

    If conn Is Nothing Then
        ErrorManager.ThrowError _
            daoActivityLog.EmptyConnection, _
            MODULE_NAME & ".createActivityLog", _
            "Koneksi database tidak boleh kosong."
    End If

    detailSize = Len(detail)
    If detailSize = 0 Then detailSize = 1

    Set cmd = New ADODB.Command

    With cmd
        Set .ActiveConnection = conn
        .CommandType = adCmdStoredProc
        .CommandText = "create_activity_log"
        .Parameters.Append .CreateParameter( _
            "p_activity_code", adVarWChar, adParamInput, 255, TrimText(activityCode))
        .Parameters.Append .CreateParameter( _
            "p_detail", adLongVarWChar, adParamInput, detailSize, detail)
        .Parameters.Append .CreateParameter( _
            "p_user", adVarWChar, adParamInput, 255, user)
        .Execute recordsAffected:=affected, Options:=adExecuteNoRecords
    End With

    If ValidationNumeric.IsLessThan(affected, 1) Then
        ErrorManager.ThrowError _
            daoActivityLog.InsertFailed, _
            MODULE_NAME & ".createActivityLog", _
            "Activity log tidak tersimpan (0 baris terpengaruh)."
    End If

CleanUp:
    Set cmd = Nothing
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    Set cmd = Nothing

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".createActivityLog", _
        "Gagal membuat activity log."

End Sub
