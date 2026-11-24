Attribute VB_Name = "testDaoReceive"
Option Explicit

Public Sub TestSelectReceives()

    On Error GoTo ErrorHandler

    Dim receiveItems As Collection
    Dim receiveItem As Receive
    Dim errInfo As ErrorInfo

    Set receiveItems = serviceReceive.SelectReceives( _
        manufacturingPlantName:="Test Plant", _
        fromDate:=DateSerial(2026, 9, 1), _
        untilDate:=DateSerial(2026, 9, 30))

    Debug.Print "Jumlah receive: " & receiveItems.Count
    Debug.Print String$(155, "=")

    Debug.Print _
        FixedWidth("RECEIVE UUID", 24) & _
        FixedWidth("ORDER UUID", 24) & _
        FixedWidth("ORDER PO ID", 18) & _
        FixedWidth("REMAINING", 16, True) & _
        FixedWidth("RECEIVE", 16, True) & _
        FixedWidth("OK", 14, True) & _
        FixedWidth("NOT OK", 14, True)

    Debug.Print String$(155, "-")

    For Each receiveItem In receiveItems

        Debug.Print _
            FixedWidth(receiveItem.UUID, 24) & _
            FixedWidth(receiveItem.Order.UUID, 24) & _
            FixedWidth(receiveItem.Order.poID, 18) & _
            FixedWidth(Format$(receiveItem.RemainingOrder, "#,##0.0000"), 16, True) & _
            FixedWidth(Format$(receiveItem.totalReceiveAmount, "#,##0.0000"), 16, True) & _
            FixedWidth(Format$(receiveItem.okAmount, "#,##0.0000"), 14, True) & _
            FixedWidth(Format$(receiveItem.notOKAmount, "#,##0.0000"), 14, True)

    Next receiveItem

    Debug.Print String$(155, "=")

CleanUp:
    Set receiveItem = Nothing
    Set receiveItems = Nothing
    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()

    managerError.ShowError errInfo, _
        "Gagal mengambil receive.", _
        "TestSelectReceives"

    Resume CleanUp

End Sub
