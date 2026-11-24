Attribute VB_Name = "testDaoStockMovement"

Option Explicit

Public Sub TestSelectStockMovement()

    On Error GoTo ErrorHandler

    Dim stockMovements As Collection
    Dim stockMovementItem As StockMovement

    Dim headerText As String
    Dim movementText As String

    Dim errInfo As ErrorInfo

    Set stockMovements = serviceStockMovement.SelectStockMovement( _
        manufacturingPlantName:="Test Plant", _
        catalogueName:="GTP - Main Device", _
        movementMonth:=9, _
        movementYear:=2026, _
        beginningStock:=0)

    Debug.Print String$(150, "=")

    headerText = FixedWidth("DATE", 12)
    headerText = headerText & FixedWidth("BEGIN", 12)
    headerText = headerText & FixedWidth("IN REF", 25)
    headerText = headerText & FixedWidth("IN FROM", 20)
    headerText = headerText & FixedWidth("IN", 12)
    headerText = headerText & FixedWidth("OUT REF", 25)
    headerText = headerText & FixedWidth("OUT TO", 20)
    headerText = headerText & FixedWidth("OUT", 12)
    headerText = headerText & FixedWidth("END", 12)

    Debug.Print headerText
    Debug.Print String$(150, "-")

    For Each stockMovementItem In stockMovements

        movementText = FixedWidth( _
            Format$( _
                CDate(stockMovementItem.MovementDate), _
                "yyyy-mm-dd"), _
            12)

        movementText = movementText & FixedWidth( _
            Format$(stockMovementItem.beginningStock, "0.00"), _
            12)

        movementText = movementText & FixedWidth( _
            stockMovementItem.IncomingReferenceNumber, _
            25)

        movementText = movementText & FixedWidth( _
            stockMovementItem.IncomingFrom, _
            20)

        movementText = movementText & FixedWidth( _
            Format$(stockMovementItem.incomingAmount, "0.00"), _
            12)

        movementText = movementText & FixedWidth( _
            stockMovementItem.OutgoingReferenceNumber, _
            25)

        movementText = movementText & FixedWidth( _
            stockMovementItem.OutgoingTo, _
            20)

        movementText = movementText & FixedWidth( _
            Format$(stockMovementItem.outgoingAmount, "0.00"), _
            12)

        movementText = movementText & FixedWidth( _
            Format$(stockMovementItem.EndStock, "0.00"), _
            12)

        Debug.Print movementText

    Next stockMovementItem

    Debug.Print String$(150, "=")
    Debug.Print "Jumlah stock movement: " & stockMovements.Count

CleanUp:
    Set stockMovementItem = Nothing
    Set stockMovements = Nothing
    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.ShowError errInfo, _
        "Gagal melakukan pengujian stock movement.", _
        "Stock Movement"

    Resume CleanUp

End Sub

Private Function FixedWidth( _
    ByVal value As String, _
    ByVal width As Long) As String

    value = TrimText(value)

    If width <= 0 Then
        FixedWidth = vbNullString
        Exit Function
    End If

    If Len(value) >= width Then
        FixedWidth = Left$(value, width - 1) & " "
    Else
        FixedWidth = value & Space$(width - Len(value))
    End If

End Function
