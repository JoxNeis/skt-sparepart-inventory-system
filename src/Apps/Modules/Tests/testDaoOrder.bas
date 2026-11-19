Attribute VB_Name = "testDaoOrder"
Option Explicit

Private Const MODULE_NAME As String = "TEST_ORDER"

Public Sub TestCreateOrder()

    On Error GoTo ErrorHandler

    Dim orderItem As Order
    Dim catalogueItem As Catalogue
    Dim plantItem As Plant
    Dim vendorItem As Vendor
    Dim errInfo As ErrorInfo

    ' --- Arrange: build the dependent objects ---
    Set catalogueItem = New Catalogue
    catalogueItem.Name = "GTP - Main Device"

    Set plantItem = New Plant
    plantItem.Name = "Test Plant"

    Set vendorItem = New Vendor
    vendorItem.Name = "Test Vendor"

    ' --- Arrange: build the order ---
    Set orderItem = New Order
    With orderItem
        .orderDate = Date
        Set .Catalogue = catalogueItem
        .Amount = 100#
        Set .Plant = plantItem
        Set .Vendor = vendorItem
        .remarks = "Test remarks"
        .personInCharge = "Tester"
    End With

    ' --- Act ---
    serviceOrder.CreateOrder orderItem

    ' --- Assert ---
    Debug.Print "PASS: " & MODULE_NAME & ".TestCreateOrder - order created"

    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()
    ErrorManager.ShowError errInfo

End Sub

Public Sub TestSelectOrder()

    On Error GoTo ErrorHandler

    Dim orders As Collection
    Dim orderItem As Order
    Dim index As Long
    Dim errInfo As ErrorInfo

    Set orders = serviceOrder.SelectPendingOrder( _
        manufacturingPlantName:="Test Plant", _
        fromDate:=DateSerial(2026, 9, 1), _
        untilDate:=DateSerial(2026, 9, 30))

    Debug.Print String$(188, "=")
    Debug.Print "TOTAL ORDER: " & orders.Count
    Debug.Print String$(188, "=")

    If orders.Count = 0 Then
        Debug.Print "Tidak ada data order."
        GoTo CleanUp
    End If

    Debug.Print _
        FixedWidth("NO.", 6) & _
        FixedWidth("ORDER ID", 10) & _
        FixedWidth("ORDER UUID", 24) & _
        FixedWidth("PO ID", 18) & _
        FixedWidth("PLANT", 20) & _
        FixedWidth("ORDER DATE", 14) & _
        FixedWidth("CATALOGUE", 22) & _
        FixedWidth("UNIT", 10) & _
        FixedWidth("AMOUNT", 16) & _
        FixedWidth("STATUS", 14) & _
        FixedWidth("VENDOR", 18) & _
        FixedWidth("PIC", 16)

    Debug.Print String$(188, "-")

    For index = 1 To orders.Count

        Set orderItem = orders(index)

        Debug.Print _
            FixedWidth(CStr(index), 6) & _
            FixedWidth(CStr(orderItem.Id), 10) & _
            FixedWidth(orderItem.UUID, 24) & _
            FixedWidth(orderItem.poID, 18) & _
            FixedWidth(orderItem.Plant.Name, 20) & _
            FixedWidth(Format$(orderItem.orderDate, "yyyy-mm-dd"), 14) & _
            FixedWidth(orderItem.Catalogue.Name, 22) & _
            FixedWidth( _
                orderItem.Catalogue.unitOfMeasurement, _
                10) & _
            FixedWidth( _
                Format$(orderItem.Amount, "#,##0.0000"), _
                16) & _
            FixedWidth(orderItem.Status, 14) & _
            FixedWidth(orderItem.Vendor.Name, 18) & _
            FixedWidth(orderItem.personInCharge, 16)

        Debug.Print _
            FixedWidth("REMARKS:", 16) & _
            orderItem.remarks

        Debug.Print String$(188, "-")

    Next index
    
    Debug.Print String$(188, "=")

CleanUp:
    Set orderItem = Nothing
    Set orders = Nothing
    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.ShowError errInfo, _
        "Gagal menguji pemilihan order.", _
        "Test Select Order"

    Resume CleanUp

End Sub

Public Sub TestUpdateOrderPoId()

    On Error GoTo ErrorHandler

    Dim orderUUID As String
    Dim poID As String
    Dim errInfo As ErrorInfo

    ' --- Arrange ---
    orderUUID = "ORD-20260909-003"
    poID = "PO-TEST-001"

    ' --- Act ---
    serviceOrder.UpdateOrderPOID orderUUID, poID

    ' --- Assert ---
    Debug.Print "PASS: " & MODULE_NAME & ".TestUpdateOrderPoId - po_id updated (UUID=" & orderUUID & ", PO=" & poID & ")"

    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()
    ErrorManager.ShowError errInfo

End Sub
