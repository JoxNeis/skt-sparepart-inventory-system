Attribute VB_Name = "MaintenanceWork"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' =========================================================
' MODEL - MAINTENANCE WORK
' Table      : maintenance_works
' Error codes: entitiesMaintenanceWork (managerErrorCodesEntities.bas)
'
' Detail of a MAINTENANCE transaction - one mobile
' maintenance job done on a date. Its lines (tasks and the
' spareparts used) are the MaintenanceWorkDetail objects in
' Details. The document number (UUID, MTW-...) and the plant
' live on the Transaction header (1:1).
'
' Removed from the old model:
'   UUID -> Transaction.UUID ([uuid] column dropped)
' =========================================================

Private Const CLASS_NAME As String = "MaintenanceWork"

Private Const TRANSACTION_TYPE As String = "MAINTENANCE"

Private Const NAME_MAX_LENGTH As Long = 255
Private Const TYPE_MAX_LENGTH As Long = 50

Private mId As Long
Private mTransaction As Transaction
Private mMaintenanceName As String
Private mMaintenanceDate As Date
Private mMaintenanceType As String
Private mDetails As Collection


Private Sub Class_Initialize()
    Set mDetails = New Collection
End Sub

Private Sub Class_Terminate()
    Set mDetails = Nothing
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
            entitiesMaintenanceWork.NegativeId, _
            CLASS_NAME & ".Id", _
            CLASS_NAME & ": Id harus lebih besar dari 0."
    End If

    mId = value

End Property


' =========================================================
' TRANSACTION                           [transaction_id] NOT NULL UNIQUE
' Must be a MAINTENANCE header (when its type is set).
' =========================================================
Public Property Get Transaction() As Transaction
    Set Transaction = mTransaction
End Property

Public Property Set Transaction(ByVal value As Transaction)

    If value Is Nothing Then
        managerError.ThrowError _
            entitiesMaintenanceWork.EmptyTransaction, _
            CLASS_NAME & ".Transaction", _
            CLASS_NAME & ": Transaksi tidak boleh kosong."
    End If

    If Len(value.TransactionType) > 0 Then
        If value.TransactionType <> TRANSACTION_TYPE Then
            managerError.ThrowError _
                entitiesMaintenanceWork.WrongTransactionType, _
                CLASS_NAME & ".Transaction", _
                CLASS_NAME & ": Tipe transaksi harus " & _
                TRANSACTION_TYPE & "."
        End If
    End If

    Set mTransaction = value

End Property


' =========================================================
' MAINTENANCE NAME                      [maintenance_name] TEXT(255) NOT NULL
' =========================================================
Public Property Get MaintenanceName() As String
    MaintenanceName = mMaintenanceName
End Property

Public Property Let MaintenanceName(ByVal value As String)

    mMaintenanceName = CheckText(value, "MaintenanceName", _
        "Nama maintenance", NAME_MAX_LENGTH, _
        entitiesMaintenanceWork.EmptyMaintenanceName, _
        entitiesMaintenanceWork.MaintenanceNameTooLong)

End Property


' =========================================================
' MAINTENANCE DATE                      [maintenance_date] NOT NULL
' =========================================================
Public Property Get MaintenanceDate() As Date
    MaintenanceDate = mMaintenanceDate
End Property

Public Property Let MaintenanceDate(ByVal value As Date)

    If managerValidationNumeric.IsZeroOrNegative(value) Then
        managerError.ThrowError _
            entitiesMaintenanceWork.EmptyMaintenanceDate, _
            CLASS_NAME & ".MaintenanceDate", _
            CLASS_NAME & ": Tanggal maintenance tidak boleh kosong."
    End If

    If DateValue(value) > VBA.Date Then
        managerError.ThrowError _
            entitiesMaintenanceWork.FutureMaintenanceDate, _
            CLASS_NAME & ".MaintenanceDate", _
            CLASS_NAME & _
            ": Tanggal maintenance tidak boleh berada di masa depan."
    End If

    mMaintenanceDate = value

End Property


' =========================================================
' MAINTENANCE TYPE                      [maintenance_type] TEXT(50) NOT NULL
' Free text for now, e.g. PREVENTIVE or CORRECTIVE.
' =========================================================
Public Property Get MaintenanceType() As String
    MaintenanceType = mMaintenanceType
End Property

Public Property Let MaintenanceType(ByVal value As String)

    mMaintenanceType = UCase$(CheckText(value, "MaintenanceType", _
        "Tipe maintenance", TYPE_MAX_LENGTH, _
        entitiesMaintenanceWork.EmptyMaintenanceType, _
        entitiesMaintenanceWork.MaintenanceTypeTooLong))

End Property


' =========================================================
' DETAILS                               maintenance_work_details
' The detail keeps only MaintenanceWorkId (no reference back
' to this object), so there is no circular reference.
' =========================================================
Public Property Get Details() As Collection
    Set Details = mDetails
End Property

Public Property Get DetailCount() As Long
    DetailCount = mDetails.Count
End Property

Public Sub AddDetail(ByVal detail As MaintenanceWorkDetail)

    If detail Is Nothing Then
        managerError.ThrowError _
            entitiesMaintenanceWork.EmptyDetail, _
            CLASS_NAME & ".AddDetail", _
            CLASS_NAME & ": Detail maintenance tidak boleh kosong."
    End If

    If mId > 0 Then
        If detail.MaintenanceWorkId = 0 Then
            detail.MaintenanceWorkId = mId
        ElseIf detail.MaintenanceWorkId <> mId Then
            managerError.ThrowError _
                entitiesMaintenanceWork.DetailOfOtherWork, _
                CLASS_NAME & ".AddDetail", _
                CLASS_NAME & _
                ": Detail milik maintenance work yang lain."
        End If
    End If

    mDetails.Add detail

End Sub

Public Sub ClearDetails()
    Set mDetails = New Collection
End Sub


' =========================================================
' CALCULATED PROPERTIES
' =========================================================

' Document number of the MAINTENANCE header ("" when not set).
Public Property Get UUID() As String
    If Not mTransaction Is Nothing Then UUID = mTransaction.UUID
End Property

' Plant of the MAINTENANCE header (Nothing when not set).
Public Property Get ManufacturingPlant() As ManufacturingPlant
    If Not mTransaction Is Nothing Then
        Set ManufacturingPlant = mTransaction.ManufacturingPlant
    End If
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
