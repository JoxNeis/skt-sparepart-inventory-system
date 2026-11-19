Attribute VB_Name = "testDaoRequest"
Option Explicit

Private Const MODULE_NAME As String = "TEST_REQUEST"

' =========================================================
' ERROR CODES (16300 - 16399)
' Usage: ErrorManager.ThrowError testDaoRequest.EmptyCreatedRequestUUID, ...
' =========================================================
Public Const EmptyCreatedRequestUUID As Long = 16301    ' CreateRequest returned no UUID

Public Sub TestWarehouseRequest()

    On Error GoTo ErrorHandler

    Dim requestItem As Request
    Dim mobilePostItem As MobilePost
    Dim catalogueItem As Catalogue
    Dim createdUuid As String
    Dim errInfo As ErrorInfo

    ' =====================================================
    ' Prepare related objects
    ' =====================================================
    Set mobilePostItem = New MobilePost

    mobilePostItem.Name = "Test Mobile Post"

    Set catalogueItem = New Catalogue

    catalogueItem.Name = "GTP - Main Device"
    catalogueItem.unitOfMeasurement = "PCS"

    ' =====================================================
    ' Create Request object
    ' =====================================================
    Set requestItem = New Request

    With requestItem
        Set .MobilePost = mobilePostItem
        Set .Catalogue = catalogueItem

        .orderDate = DateSerial(2026, 9, 1)
        .Amount = 100
        .unitOfMeasurement = catalogueItem.unitOfMeasurement
        .ArriveDate = DateSerial(2026, 9, 5)
        .remarks = "Warehouse request untuk pengujian."
        .SystemRemarks = "Created by TestWarehouseRequest."
        .personInCharge = "Test User"
    End With

    ' =====================================================
    ' Execute database operation
    ' =====================================================
    Debug.Print String$(80, "=")
    Debug.Print "CREATING WAREHOUSE REQUEST"
    Debug.Print String$(80, "=")

    createdUuid = serviceRequest.CreateRequest( _
        requestItem:=requestItem)

    If ValidationText.IsBlank(createdUuid) Then
        ErrorManager.ThrowError _
            testDaoRequest.EmptyCreatedRequestUUID, _
            MODULE_NAME & ".TestWarehouseRequest", _
            "CreateRequest tidak mengembalikan UUID."
    End If

    ' =====================================================
    ' Print test result only after CreateRequest succeeds
    ' =====================================================
    Debug.Print "CREATED UUID      : " & createdUuid
    Debug.Print "OBJECT UUID       : " & requestItem.UUID
    Debug.Print "Mobile Post       : " & requestItem.MobilePost.Name
    Debug.Print "Request Date      : " & _
        Format$(requestItem.orderDate, "yyyy-mm-dd")
    Debug.Print "Catalogue         : " & requestItem.Catalogue.Name
    Debug.Print "Amount            : " & _
        Format$(requestItem.Amount, "#,##0.0000")
    Debug.Print "Unit              : " & requestItem.unitOfMeasurement
    Debug.Print "Remarks           : " & requestItem.remarks
    Debug.Print "System Remarks    : " & requestItem.SystemRemarks
    Debug.Print "Person In Charge  : " & requestItem.personInCharge

    Debug.Print String$(80, "=")
    Debug.Print "TEST RESULT: SUCCESS"
    Debug.Print "DATA HAS BEEN COMMITTED TO ACCESS"
    Debug.Print String$(80, "=")

    FormManager.ShowInfo _
        "Request berhasil dibuat." & _
        vbCrLf & vbCrLf & _
        "UUID: " & createdUuid, _
        "Test Request"

CleanUp:
    Set requestItem = Nothing
    Set mobilePostItem = Nothing
    Set catalogueItem = Nothing
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.ShowError errInfo, _
        "Gagal menguji penyimpanan Request.", _
        "Test Request"

    Resume CleanUp

End Sub

Public Sub TestGetRequestsByManufacturingPlantAndOrderDate()

    On Error GoTo ErrorHandler

    Dim requestItems As Collection
    Dim requestItem As Request

    Dim manufacturingPlantName As String
    Dim fromDate As Date
    Dim untilDate As Date

    Dim rowNumber As Long
    Dim errInfo As ErrorInfo

    manufacturingPlantName = "Test Plant"
    fromDate = DateSerial(2026, 9, 1)
    untilDate = DateSerial(2026, 9, 30)

    Set requestItems = _
        serviceRequest.GetRequestsByManufacturingPlantAndOrderDate( _
            manufacturingPlantName:=manufacturingPlantName, _
            fromDate:=fromDate, _
            untilDate:=untilDate)

    Debug.Print String$(160, "=")
    Debug.Print "REQUEST SELECT TEST"
    Debug.Print String$(160, "=")

    Debug.Print "Manufacturing plant : " & _
        manufacturingPlantName

    Debug.Print "From date           : " & _
        Format$(fromDate, "yyyy-mm-dd")

    Debug.Print "Until date          : " & _
        Format$(untilDate, "yyyy-mm-dd")

    Debug.Print "Rows returned       : " & _
        requestItems.Count

    Debug.Print String$(160, "-")

    If requestItems.Count = 0 Then
        Debug.Print "No request data found."
    Else
        PrintRequestHeader

        For Each requestItem In requestItems
            rowNumber = rowNumber + 1

            PrintRequestRow _
                rowNumber:=rowNumber, _
                requestItem:=requestItem
        Next requestItem
    End If

    Debug.Print String$(160, "=")
    Debug.Print "TEST RESULT: SUCCESS"
    Debug.Print String$(160, "=")

CleanUp:
    Set requestItem = Nothing
    Set requestItems = Nothing
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.ShowError errInfo, _
        "Gagal menguji select request.", _
        "Test Select Request"

    Resume CleanUp

End Sub

Private Sub PrintRequestHeader()

    Debug.Print _
        FixedWidth("NO", 5) & _
        FixedWidth("REQUEST UUID", 22) & _
        FixedWidth("DATE", 12) & _
        FixedWidth("MOBILE POST", 25) & _
        FixedWidth("CATALOGUE", 30) & _
        FixedWidth("AMOUNT", 14) & _
        FixedWidth("UNIT", 10) & _
        FixedWidth("PERSON IN CHARGE", 24) & _
        FixedWidth("REMARKS", 35)

    Debug.Print String$(177, "-")

End Sub

Private Sub PrintRequestRow( _
    ByVal rowNumber As Long, _
    ByVal requestItem As Request)

    Debug.Print _
        FixedWidth(CStr(rowNumber), 5) & _
        FixedWidth(requestItem.UUID, 22) & _
        FixedWidth( _
            Format$(requestItem.orderDate, "yyyy-mm-dd"), _
            12) & _
        FixedWidth(requestItem.MobilePost.Name, 25) & _
        FixedWidth(requestItem.Catalogue.Name, 30) & _
        FixedWidth( _
            Format$(requestItem.Amount, "#,##0.####"), _
            14) & _
        FixedWidth(requestItem.unitOfMeasurement, 10) & _
        FixedWidth(requestItem.personInCharge, 24) & _
        FixedWidth(requestItem.remarks, 35)

End Sub
