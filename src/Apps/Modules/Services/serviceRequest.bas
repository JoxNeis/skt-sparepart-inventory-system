Attribute VB_Name = "serviceRequest"
Option Explicit

' =========================================================
' SERVICE - REQUEST (WAREHOUSE REQUEST)
'
' Public entry points (called by sheets) get the
' connection and own the transaction:
'   CreateRequest, GetRequestsByManufacturingPlantAndOrderDate
' =========================================================

Private Const MODULE_NAME As String = "SERVICE_REQUEST"

' =========================================================
' ERROR CODES (17700 - 17799)
' Usage: managerError.ThrowError serviceRequest.EmptyRequest, ...
' =========================================================
Public Const EmptyRequest As Long = 17701                       ' Request object is Nothing
Public Const EmptyMobilePost As Long = 17702                    ' MobilePost object is Nothing
Public Const EmptyPlant As Long = 17703                         ' MobilePost.Plant object is Nothing
Public Const EmptyCatalogue As Long = 17704                     ' Catalogue object is Nothing
Public Const EmptyRequestDate As Long = 17705                   ' orderDate <= 0
Public Const InvalidRequestAmount As Long = 17706               ' Amount <= 0
Public Const EmptyMobilePostName As Long = 17707
Public Const EmptyManufacturingPlantName As Long = 17708
Public Const EmptyCatalogueName As Long = 17709
Public Const EmptyPersonInCharge As Long = 17710
Public Const RequestAmountExceedsStock As Long = 17711          ' stock after request < 0
Public Const EmptyTransactionUUID As Long = 17712               ' CreateTransaction returned no UUID
Public Const SameOriginalProposedUUID As Long = 17713           ' original UUID = proposed UUID
Public Const NoRowAffectedCreateRequest As Long = 17714         ' create_request inserted 0 rows
Public Const MultipleRowsAffectedCreateRequest As Long = 17715  ' create_request inserted > 1 row
Public Const EmptyFromDate As Long = 17716                      ' fromDate <= 0
Public Const EmptyUntilDate As Long = 17717                     ' untilDate <= 0
Public Const InvalidDateRange As Long = 17718                   ' fromDate > untilDate


' =========================================================
' PUBLIC ENTRY POINTS
' =========================================================

' Create a warehouse request in one transaction and return
' the new transaction UUID (also stored in requestItem.UUID).
'
' originalUUID = ""  : normal request (stock check, active
'                      transaction, current stock - amount).
' originalUUID <> "" : proposed update of that request
'                      (inactive transaction + update
'                      request, stock is not changed).
Public Function CreateRequest( _
    ByVal requestItem As Request, _
    Optional ByVal originalUUID As String = "") As String

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim affected As Long

    Dim transactionUUID As String
    Dim isCreatedForUpdate As Boolean
    Dim inTransaction As Boolean

    Dim currentStock As Double
    Dim availableStockAfterRequest As Double

    Dim activityCode As String
    Dim activityDetail As String
    Dim validationMessage As String

    Dim errInfo As ErrorInfo

    inTransaction = False

    ' =====================================================
    ' Step 1: Validate request object
    ' =====================================================
    If requestItem Is Nothing Then
        managerError.ThrowError _
            serviceRequest.EmptyRequest, _
            MODULE_NAME & ".CreateRequest", _
            "Request tidak boleh kosong."
    End If

    If requestItem.MobilePost Is Nothing Then
        managerError.ThrowError _
            serviceRequest.EmptyMobilePost, _
            MODULE_NAME & ".CreateRequest", _
            "Mobile post request tidak boleh kosong."
    End If

    If requestItem.MobilePost.Plant Is Nothing Then
        managerError.ThrowError _
            serviceRequest.EmptyPlant, _
            MODULE_NAME & ".CreateRequest", _
            "Manufacturing plant request " & _
            "tidak boleh kosong."
    End If

    If requestItem.Catalogue Is Nothing Then
        managerError.ThrowError _
            serviceRequest.EmptyCatalogue, _
            MODULE_NAME & ".CreateRequest", _
            "Catalogue request tidak boleh kosong."
    End If

    ' =====================================================
    ' Step 2: Normalize and determine creation mode
    ' =====================================================
    originalUUID = Trim$(originalUUID)

    isCreatedForUpdate = (Len(originalUUID) > 0)

    ' =====================================================
    ' Step 3: Validate request properties
    ' =====================================================
    If managerValidationNumeric.IsZeroOrNegative(requestItem.orderDate) Then
        managerError.ThrowError _
            serviceRequest.EmptyRequestDate, _
            MODULE_NAME & ".CreateRequest", _
            "Tanggal request tidak boleh kosong."
    End If

    If managerValidationNumeric.IsZeroOrNegative(requestItem.Amount) Then
        managerError.ThrowError _
            serviceRequest.InvalidRequestAmount, _
            MODULE_NAME & ".CreateRequest", _
            "Jumlah request harus lebih besar dari nol."
    End If

    If managerValidationText.IsBlank(requestItem.MobilePost.Name) Then
        managerError.ThrowError _
            serviceRequest.EmptyMobilePostName, _
            MODULE_NAME & ".CreateRequest", _
            "Nama mobile post tidak boleh kosong."
    End If

    If managerValidationText.IsBlank(requestItem.MobilePost.Plant.Name) Then
        managerError.ThrowError _
            serviceRequest.EmptyManufacturingPlantName, _
            MODULE_NAME & ".CreateRequest", _
            "Nama manufacturing plant " & _
            "tidak boleh kosong."
    End If

    If managerValidationText.IsBlank(requestItem.Catalogue.Name) Then
        managerError.ThrowError _
            serviceRequest.EmptyCatalogueName, _
            MODULE_NAME & ".CreateRequest", _
            "Nama catalogue tidak boleh kosong."
    End If

    If managerValidationText.IsBlank(requestItem.personInCharge) Then
        managerError.ThrowError _
            serviceRequest.EmptyPersonInCharge, _
            MODULE_NAME & ".CreateRequest", _
            "Person in charge tidak boleh kosong."
    End If

    ' =====================================================
    ' Step 4: Open connection and begin transaction
    ' =====================================================
    Set conn = managerDatabase.GetConnection()

    If conn Is Nothing Then
        managerError.ThrowError _
            managerDatabase.ConnectionFailed, _
            MODULE_NAME & ".CreateRequest", _
            "Connection database tidak berhasil dibuat."
    End If

    If conn.State <> adStateOpen Then
        managerError.ThrowError _
            managerDatabase.ConnectionNotOpen, _
            MODULE_NAME & ".CreateRequest", _
            "Connection database tidak terbuka."
    End If

    managerDatabase.BeginTransaction conn
    inTransaction = True

    ' =====================================================
    ' Step 5: Check available stock
    '
    ' Normal request:
    '   Check available stock because the transaction will
    '   be active and immediately reduce current stock.
    '   Read inside the transaction, so the check and the
    '   stock change see the same data.
    '
    ' Proposed update:
    '   Do not apply a stock movement yet. The existing
    '   original request remains active until approval.
    ' =====================================================
    If Not isCreatedForUpdate Then

        currentStock = daoStockMovement.SelectBeginningStock( _
            conn, _
            requestItem.orderDate, _
            requestItem.MobilePost.Plant.Name, _
            requestItem.Catalogue.Name)

        availableStockAfterRequest = _
            currentStock - requestItem.Amount

        If managerValidationNumeric.IsNegative(availableStockAfterRequest) Then

            validationMessage = _
                "Jumlah request melebihi stok yang tersedia."

            validationMessage = validationMessage & _
                vbCrLf & _
                "Tanggal Request: " & _
                Format$(requestItem.orderDate, "yyyy-mm-dd")

            validationMessage = validationMessage & _
                vbCrLf & _
                "Nama Sparepart: " & _
                requestItem.Catalogue.Name

            validationMessage = validationMessage & _
                vbCrLf & _
                "Stok Tersedia: " & _
                Format$(currentStock, "#,##0.00")

            validationMessage = validationMessage & _
                vbCrLf & _
                "Jumlah Request: " & _
                Format$(requestItem.Amount, "#,##0.00")

            validationMessage = validationMessage & _
                vbCrLf & _
                "Kekurangan Stok: " & _
                Format$(Abs(availableStockAfterRequest), "#,##0.00")

            managerError.ThrowError _
                serviceRequest.RequestAmountExceedsStock, _
                MODULE_NAME & ".CreateRequest", _
                validationMessage
        End If

    End If

    ' =====================================================
    ' Step 6: Create REQUEST transaction
    '
    ' Normal request:
    '   active = True
    '
    ' Proposed update:
    '   active = False
    ' =====================================================
    transactionUUID = serviceTransaction.CreateTransaction( _
        conn:=conn, _
        transactionType:="REQUEST", _
        transactionDate:=requestItem.orderDate, _
        isActive:=Not isCreatedForUpdate)

    transactionUUID = Trim$(transactionUUID)

    If managerValidationText.IsBlank(transactionUUID) Then
        managerError.ThrowError _
            serviceRequest.EmptyTransactionUUID, _
            MODULE_NAME & ".CreateRequest", _
            "UUID transaksi request tidak berhasil dibuat."
    End If

    If isCreatedForUpdate Then
        If managerValidationText.IsSameText(originalUUID, transactionUUID) Then
            managerError.ThrowError _
                serviceRequest.SameOriginalProposedUUID, _
                MODULE_NAME & ".CreateRequest", _
                "UUID transaksi original dan proposed " & _
                "tidak boleh sama."
        End If
    End If

    ' =====================================================
    ' Step 7: Create request row
    ' =====================================================
    affected = daoRequest.CreateRequest(conn, transactionUUID, requestItem)

    If managerValidationNumeric.IsLessThan(affected, 1) Then
        managerError.ThrowError _
            serviceRequest.NoRowAffectedCreateRequest, _
            MODULE_NAME & ".CreateRequest", _
            "Request tidak tersimpan. Transaction, " & _
            "mobile post, atau catalogue mungkin " & _
            "tidak ditemukan."
    End If

    If managerValidationNumeric.IsGreaterThan(affected, 1) Then
        managerError.ThrowError _
            serviceRequest.MultipleRowsAffectedCreateRequest, _
            MODULE_NAME & ".CreateRequest", _
            "Penyimpanan request memengaruhi lebih " & _
            "dari satu baris."
    End If

    ' =====================================================
    ' Step 8: Create update request or update stock
    ' =====================================================
    If isCreatedForUpdate Then

        ' The proposed transaction remains inactive.
        ' Stock is not updated until approval.
        serviceTransactionRequestUpdate.CreateRequestUpdate _
            conn, originalUUID, transactionUUID

    Else

        ' Normal active request immediately reduces stock.
        servicePlantStock.ApplyStockChange _
            conn, _
            -requestItem.Amount, _
            requestItem.Catalogue.Name, _
            requestItem.MobilePost.Plant.Name

    End If

    ' =====================================================
    ' Step 9: Build activity log information
    ' =====================================================
    If isCreatedForUpdate Then
        activityCode = "Request.CreateUpdateProposal"
    Else
        activityCode = "Request.Create"
    End If

    activityDetail = BuildRequestDetail(requestItem, transactionUUID)

    activityDetail = activityDetail & _
        BuildRequestUpdateDetail(isCreatedForUpdate, originalUUID)

    ' =====================================================
    ' Step 10: Write activity log
    ' =====================================================
    serviceActivityLog.WriteActivityLog _
        conn, activityCode, activityDetail

    ' =====================================================
    ' Step 11: Commit
    ' =====================================================
    managerDatabase.CommitTransaction conn
    inTransaction = False

    requestItem.UUID = transactionUUID

    CreateRequest = transactionUUID

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()

    If inTransaction Then
        managerDatabase.RollbackTransaction conn
    End If

    managerError.Rethrow errInfo, _
        MODULE_NAME & ".CreateRequest", _
        "Gagal membuat request."

End Function


' Requests of one manufacturing plant between fromDate and
' untilDate (collection without key).
Public Function GetRequestsByManufacturingPlantAndOrderDate( _
    ByVal manufacturingPlantName As String, _
    ByVal fromDate As Date, _
    ByVal untilDate As Date) As Collection

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim errInfo As ErrorInfo

    manufacturingPlantName = Trim$(manufacturingPlantName)

    ' =====================================================
    ' Step 1: Validate
    ' =====================================================
    If managerValidationText.IsBlank(manufacturingPlantName) Then
        managerError.ThrowError _
            serviceRequest.EmptyManufacturingPlantName, _
            MODULE_NAME & ".GetRequestsByManufacturingPlantAndOrderDate", _
            "Nama manufacturing plant tidak boleh kosong."
    End If

    If managerValidationNumeric.IsZeroOrNegative(fromDate) Then
        managerError.ThrowError _
            serviceRequest.EmptyFromDate, _
            MODULE_NAME & ".GetRequestsByManufacturingPlantAndOrderDate", _
            "Tanggal awal tidak boleh kosong."
    End If

    If managerValidationNumeric.IsZeroOrNegative(untilDate) Then
        managerError.ThrowError _
            serviceRequest.EmptyUntilDate, _
            MODULE_NAME & ".GetRequestsByManufacturingPlantAndOrderDate", _
            "Tanggal akhir tidak boleh kosong."
    End If

    If fromDate > untilDate Then
        managerError.ThrowError _
            serviceRequest.InvalidDateRange, _
            MODULE_NAME & ".GetRequestsByManufacturingPlantAndOrderDate", _
            "Tanggal awal tidak boleh lebih besar " & _
            "dari tanggal akhir."
    End If

    ' =====================================================
    ' Step 2: Query (dates are passed as given, as before)
    ' =====================================================
    Set conn = managerDatabase.GetConnection()

    Set GetRequestsByManufacturingPlantAndOrderDate = _
        daoRequest.GetRequestsByManufacturingPlantAndOrderDate( _
            conn, manufacturingPlantName, fromDate, untilDate)

    Exit Function

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.Rethrow errInfo, _
        MODULE_NAME & _
        ".GetRequestsByManufacturingPlantAndOrderDate", _
        "Gagal mengambil data request."

End Function


' =========================================================
' INTERNAL
' =========================================================

Private Function BuildRequestDetail( _
    ByVal requestItem As Request, _
    ByVal transactionUUID As String) As String

    Dim detail As String

    detail = _
        "UUID: " & Trim$(transactionUUID) & vbCrLf & _
        "Request Date: " & _
            Format$(requestItem.orderDate, "yyyy-mm-dd") & vbCrLf & _
        "Mobile Post: " & _
            Trim$(requestItem.MobilePost.Name) & vbCrLf & _
        "Catalogue: " & _
            Trim$(requestItem.Catalogue.Name) & vbCrLf & _
        "Amount: " & _
            Format$(requestItem.Amount, "#,##0.####") & vbCrLf & _
        "Unit: " & _
            Trim$(requestItem.unitOfMeasurement) & vbCrLf & _
        "Remarks: " & _
            Trim$(requestItem.remarks) & vbCrLf & _
        "System Remarks: " & _
            Trim$(requestItem.SystemRemarks) & vbCrLf & _
        "Person In Charge: " & _
            Trim$(requestItem.personInCharge)

    BuildRequestDetail = detail

End Function


Private Function BuildRequestUpdateDetail( _
    ByVal isCreatedForUpdate As Boolean, _
    ByVal originalUUID As String) As String

    originalUUID = Trim$(originalUUID)

    If Not isCreatedForUpdate Then
        BuildRequestUpdateDetail = vbNullString
        Exit Function
    End If

    BuildRequestUpdateDetail = _
        "; update_proposal=true"

    BuildRequestUpdateDetail = _
        BuildRequestUpdateDetail & _
        "; original_transaction_uuid=" & _
        originalUUID

End Function
