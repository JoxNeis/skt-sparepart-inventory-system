Attribute VB_Name = "MaintenanceWorkDetail"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' =========================================================
' MODEL - MAINTENANCE WORK DETAIL
' Table      : maintenance_work_details
' Error codes: entitiesMaintenanceWorkDetail (ErrorCodesEntities.bas)
'
' One line of a MaintenanceWork - a task performed, optionally
' with the sparepart and amount used for it.
' Sparepart = Nothing and Amount = 0 (both NULL) for a task
' that used no sparepart; call Validate before saving.
'
' MaintenanceTaskId is a plain Long: there is no
' maintenance_tasks table (and no model) yet.
' =========================================================

Private Const CLASS_NAME As String = "MaintenanceWorkDetail"

Private mId As Long
Private mMaintenanceWorkId As Long
Private mMaintenanceTaskId As Long
Private mSparepart As Sparepart
Private mAmount As Long


' =========================================================
' ID                                    [id]
' =========================================================
Public Property Get Id() As Long
    Id = mId
End Property

Public Property Let Id(ByVal value As Long)

    If ValidationNumeric.IsZeroOrNegative(value) Then
        ErrorManager.ThrowError _
            entitiesMaintenanceWorkDetail.NegativeId, _
            CLASS_NAME & ".Id", _
            CLASS_NAME & ": Id harus lebih besar dari 0."
    End If

    mId = value

End Property


' =========================================================
' MAINTENANCE WORK ID                   [maintenance_work_id] NOT NULL
' =========================================================
Public Property Get MaintenanceWorkId() As Long
    MaintenanceWorkId = mMaintenanceWorkId
End Property

Public Property Let MaintenanceWorkId(ByVal value As Long)

    If ValidationNumeric.IsZeroOrNegative(value) Then
        ErrorManager.ThrowError _
            entitiesMaintenanceWorkDetail.InvalidMaintenanceWorkId, _
            CLASS_NAME & ".MaintenanceWorkId", _
            CLASS_NAME & ": Id maintenance work harus lebih besar dari 0."
    End If

    mMaintenanceWorkId = value

End Property


' =========================================================
' MAINTENANCE TASK ID                   [maintenance_task_id] NOT NULL
' =========================================================
Public Property Get MaintenanceTaskId() As Long
    MaintenanceTaskId = mMaintenanceTaskId
End Property

Public Property Let MaintenanceTaskId(ByVal value As Long)

    If ValidationNumeric.IsZeroOrNegative(value) Then
        ErrorManager.ThrowError _
            entitiesMaintenanceWorkDetail.InvalidMaintenanceTaskId, _
            CLASS_NAME & ".MaintenanceTaskId", _
            CLASS_NAME & ": Id task maintenance harus lebih besar dari 0."
    End If

    mMaintenanceTaskId = value

End Property


' =========================================================
' SPAREPART                             [sparepart_id] (optional)
' Nothing = NULL.
' =========================================================
Public Property Get Sparepart() As Sparepart
    Set Sparepart = mSparepart
End Property

Public Property Set Sparepart(ByVal value As Sparepart)
    Set mSparepart = value
End Property


' =========================================================
' AMOUNT                                [amount] LONG (optional)
' 0 = NULL.
' =========================================================
Public Property Get Amount() As Long
    Amount = mAmount
End Property

Public Property Let Amount(ByVal value As Long)

    If ValidationNumeric.IsNegative(value) Then
        ErrorManager.ThrowError _
            entitiesMaintenanceWorkDetail.NegativeAmount, _
            CLASS_NAME & ".Amount", _
            CLASS_NAME & ": Jumlah sparepart tidak boleh negatif."
    End If

    mAmount = value

End Property


' =========================================================
' CALCULATED PROPERTIES
' =========================================================
Public Property Get UsesSparepart() As Boolean
    UsesSparepart = Not (mSparepart Is Nothing)
End Property


' =========================================================
' VALIDATE
' A sparepart needs an amount > 0, and an amount needs a
' sparepart. Call it before saving.
' =========================================================
Public Sub Validate()

    If UsesSparepart And mAmount = 0 Then
        ErrorManager.ThrowError _
            entitiesMaintenanceWorkDetail.SparepartWithoutAmount, _
            CLASS_NAME & ".Validate", _
            CLASS_NAME & _
            ": Jumlah sparepart harus lebih besar dari 0."
    End If

    If Not UsesSparepart And mAmount > 0 Then
        ErrorManager.ThrowError _
            entitiesMaintenanceWorkDetail.AmountWithoutSparepart, _
            CLASS_NAME & ".Validate", _
            CLASS_NAME & _
            ": Sparepart tidak boleh kosong bila jumlah diisi."
    End If

End Sub
