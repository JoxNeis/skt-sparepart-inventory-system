Attribute VB_Name = "managerError"
Option Explicit

' =========================================================
' ERROR MANAGER
'
' Central place to raise, propagate, inspect and display
' application errors.
'
' Every application error has a numeric code. The code is
' referenced through its owner, so the owner is visible at
' the place where the error is raised or checked:
'
'   Standard modules declare Public Const codes inside the
'   module itself (the module name is the namespace):
'       serviceActivityLog.InsertFailed
'       serviceOrder.EmptyOrder
'       managerDatabase.ConnectionFailed
'       managerTable.ColumnNotFound
'
'   Class modules and worksheet modules cannot hold Public
'   Const, so their codes are Public Enums in:
'       managerErrorCodesEntities  ->  entitiesOrder.NegativeId
'       managerErrorCodesSheets    ->  sheetInputOrder.EmptyPlantName
'
' Raising an error (validation functions return Boolean,
' the caller decides to throw):
'
'   If managerValidationNumeric.IsZeroOrNegative(value) Then
'       managerError.ThrowError _
'           entitiesOrder.NegativeId, _
'           CLASS_NAME & ".Id", _
'           "ID harus lebih besar dari 0."
'   End If
'
' Propagating an error from an error handler:
'
'   ErrorHandler:
'       errInfo = managerError.CaptureError()   ' first line
'       ' ... cleanup ...
'       managerError.Rethrow errInfo, _
'           MODULE_NAME & ".CreateOrder", _
'           "Gagal membuat order."
'
' Getting the error (the original code is preserved while
' the error travels up through Rethrow):
'
'   If managerError.HasErrorCode(errInfo, entitiesOrder.NegativeId) Then
'       ...
'   End If
'
'   Select Case errInfo.Number
'       Case serviceOrder.EmptyOrder
'       Case entitiesOrder.NegativeId
'   End Select
'
' ---------------------------------------------------------
' ERROR CODE RANGES (each owner gets a block of 100)
'
'   10000 - 10999  Managers and infrastructure
'       10000  managerError
'       10100  managerDatabase
'       10200  coreHelper
'       10300  managerForm
'       10400  managerTable
'
'   11000 - 11999  Entities (class modules, blocks of 50)
'       11000  entitiesManufacturingPlant
'       11050  entitiesUser
'       11100  entitiesActivityLog
'       11150  entitiesVendor
'       11200  entitiesSparepart
'       11250  entitiesMobileMaintenance
'       11300  entitiesManufacturingPlantStock
'       11350  entitiesTransaction
'       11400  entitiesOrder
'       11450  entitiesReceive
'       11500  entitiesRequest
'       11550  entitiesTransactionUpdateRequest
'       11600  entitiesMaintenanceWork
'       11650  entitiesMaintenanceWorkDetail
'       11700  entitiesAssembly
'       11750  entitiesStockLedgerEntry
'
'   12000 - 13999  DAO modules
'       12100  daoActivityLog
'       12200  daoCatalogue
'       12300  daoMobilePost
'       12400  daoOrder
'       12500  daoPlantStock
'       12600  daoReceive
'       12700  daoRequest
'       12800  daoStockMovement
'       12900  daoTransaction
'       13000  daoTransactionRequestUpdate
'       13100  daoUpdateOrder
'       13200  daoUpdateReceive
'       13300  daoUpdateRequest
'       13400  daoUser
'       13500  daoVendor
'
'   14000 - 15999  Worksheets and user forms
'       14100  sheetInputOrder
'       14200  sheetInputPenerimaan
'       14300  sheetInputRequest
'       14400  sheetInventory
'       14500  sheetKartuStock
'       14600  sheetMasterDataCatalogue
'       14700  sheetMasterDataMobilePost
'       14800  sheetMasterDataUser
'       14900  sheetMasterDataVendor
'       15000  sheetRecapOrder
'       15100  sheetRecapPendingOrder
'       15200  sheetRecapPenerimaan
'       15300  sheetRecapRequest
'       15400  sheetUpdateOrder
'       15500  sheetUpdatePenerimaan
'       15600  sheetUpdateRequest
'
'   16000 - 16999  Test modules
'       16100  testDaoOrder
'       16200  testDaoReceive
'       16300  testDaoRequest
'       16400  testDaoStockMovement
'       16500  testDaoTransaction
'
'   17000 - 18999  Services (business logic + transactions)
'       17100  serviceActivityLog
'       17200  serviceCatalogue
'       17300  serviceMobilePost
'       17400  serviceOrder
'       17500  servicePlantStock
'       17600  serviceReceive
'       17700  serviceRequest
'       17800  serviceStockMovement
'       17900  serviceTransaction
'       18000  serviceTransactionRequestUpdate
'       18100  serviceUser
'       18200  serviceVendor
'
' DAO modules only run one saved query per function, so
' they normally raise nothing themselves: connection,
' command and field errors come from managerDatabase, and
' validation / row-count errors come from the services.
' =========================================================

Private Const MODULE_NAME As String = "managerError"

Public Const APP_ERROR_FIRST As Long = 10000
Public Const APP_ERROR_LAST As Long = 19999

' =========================================================
' ERROR CODES (10000 - 10099)
' =========================================================

' An error was propagated without any error number.
Public Const UnknownError As Long = 10001

' =========================================================
' ERROR SNAPSHOT
'
' The Err object is cleared by every On Error statement,
' Resume, Exit Sub/Function and by many cleanup calls.
' Capture it as the first statement of an error handler.
' =========================================================
Public Type ErrorInfo
    Number As Long
    Source As String
    Description As String
End Type


' =========================================================
' Take a snapshot of the current Err object.
' =========================================================
Public Function CaptureError() As ErrorInfo

    Dim info As ErrorInfo

    info.Number = Err.Number
    info.Source = Err.Source
    info.Description = Err.Description

    CaptureError = info

End Function


' =========================================================
' Raise a new application error.
'
' errorCode    : owner code, e.g. entitiesOrder.NegativeId
' errorSource  : MODULE_NAME / CLASS_NAME & ".ProcedureName"
' errorMessage : message for the user
' =========================================================
Public Sub ThrowError( _
    ByVal errorCode As Long, _
    ByVal errorSource As String, _
    ByVal errorMessage As String)

    If errorCode = 0 Then
        errorCode = UnknownError
    End If

    Err.Raise _
        Number:=errorCode, _
        Source:=errorSource, _
        Description:=errorMessage

End Sub


' =========================================================
' Raise a captured error again, adding the current
' procedure to the source chain and an optional message on
' top of the original description.
'
' The original error number is kept, so the caller at the
' top can still find out which error happened.
'
' fallbackCode : optional application code used when the
'                captured error is NOT an application error
'                (for example a raw ADO or VBA runtime error).
' =========================================================
Public Sub Rethrow( _
    ByRef info As ErrorInfo, _
    ByVal errorSource As String, _
    Optional ByVal errorMessage As String = vbNullString, _
    Optional ByVal fallbackCode As Long = 0)

    Dim finalNumber As Long
    Dim finalSource As String
    Dim finalDescription As String

    finalNumber = info.Number

    If fallbackCode <> 0 Then
        If Not IsAppError(finalNumber) Then
            finalNumber = fallbackCode
        End If
    End If

    If finalNumber = 0 Then
        finalNumber = UnknownError
    End If

    finalSource = errorSource

    If Len(Trim$(info.Source)) > 0 Then
        finalSource = finalSource & " -> " & info.Source
    End If

    finalDescription = errorMessage

    If Len(Trim$(info.Description)) > 0 Then
        If Len(finalDescription) > 0 Then
            finalDescription = finalDescription & vbCrLf & _
                               info.Description
        Else
            finalDescription = info.Description
        End If
    End If

    Err.Raise _
        Number:=finalNumber, _
        Source:=finalSource, _
        Description:=finalDescription

End Sub


' =========================================================
' True when the number belongs to the application code
' ranges (10000 - 19999).
' =========================================================
Public Function IsAppError( _
    ByVal errorNumber As Long) As Boolean

    IsAppError = _
        (errorNumber >= APP_ERROR_FIRST And _
         errorNumber <= APP_ERROR_LAST)

End Function


' =========================================================
' True when the captured error has the given code.
' =========================================================
Public Function HasErrorCode( _
    ByRef info As ErrorInfo, _
    ByVal errorCode As Long) As Boolean

    HasErrorCode = (info.Number = errorCode)

End Function


' =========================================================
' Show a captured error to the user.
'
' context : optional extra line, e.g. "Baris ke-3 pada tabel."
' =========================================================
Public Sub ShowError( _
    ByRef info As ErrorInfo, _
    Optional ByVal context As String = vbNullString, _
    Optional ByVal title As String = vbNullString)

    Dim message As String

    If Len(title) = 0 Then
        title = "Aplikasi " & coreConfig.ApplicationName() & _
                " Mengalami Error!"
    End If

    message = "Error !!!" & vbCrLf & vbCrLf

    If Len(context) > 0 Then
        message = message & context & vbCrLf
    End If

    message = message & _
        "Kode: " & CStr(info.Number) & vbCrLf & _
        "Sumber: " & info.Source & vbCrLf & _
        "Keterangan: " & info.Description

    Debug.Print "[" & MODULE_NAME & "] " & _
        CStr(info.Number) & " | " & _
        info.Source & " | " & _
        Replace(info.Description, vbCrLf, " / ")

    MsgBox _
        Prompt:=message, _
        Buttons:=vbCritical + vbOKOnly, _
        title:=title

End Sub
