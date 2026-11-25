Attribute VB_Name = "controllerInputOrder"
Option Explicit

' =========================================================
' CONTROLLER - INPUT ORDER (sheet INPUT_ORDER)
'
'   LOKASI        K5   dropdown (plant)
'   TANGGAL ORDER K7   double-click = date picker
'   BERSIHKAN TABEL  N5:O5   double-click
'   TAMBAH BARIS     N7:O7   double-click
'   SIMPAN           N9:O9   double-click
'   KIRIM EMAIL      I9:J9   double-click (.eml email draft)
'   CEK SPAREPART    R6 (dropdown, free text) + CARI T6:U6
'
'   TableInputOrder: NAMA SPAREPART (dropdown), JUMLAH ORDER,
'   SATUAN (automatic), VENDOR (dropdown), REMARKS, PIC,
'   HAPUS (double-click deletes the row)
' =========================================================

Private Const MODULE_NAME As String = "controllerInputOrder"

' =========================================================
' ERROR CODES (14100 - 14199)
' =========================================================
Public Const NoOrderRow As Long = 14101
Public Const NothingToEmail As Long = 14102

Private Const PLANT_CELL As String = "K5"
Private Const DATE_CELL As String = "K7"

Private Const BUTTON_CLEAR As String = "N5:O5"
Private Const BUTTON_ADD_ROW As String = "N7:O7"
Private Const BUTTON_SAVE As String = "N9:O9"
Private Const BUTTON_EMAIL As String = "I9:J9"

Private Const SEARCH_CELL As String = "R6"
Private Const BUTTON_SEARCH As String = "T6:U6"
Private Const RESULT_FIRST_CELL As String = "Q9"

Private Const TABLE_NAME As String = "TableInputOrder"
Private Const COL_SPAREPART As String = "NAMA SPAREPART"
Private Const COL_QUANTITY As String = "JUMLAH ORDER"
Private Const COL_UNIT As String = "SATUAN"
Private Const COL_VENDOR As String = "VENDOR"
Private Const COL_REMARKS As String = "REMARKS"
Private Const COL_PIC As String = "PIC"
Private Const COL_DELETE As String = "HAPUS"

' Orders of the last SIMPAN, for KIRIM EMAIL.
Private mLastSaved As Collection
Private mLastSavedPlant As String


' =========================================================
' EVENTS
' =========================================================

Public Sub OnActivate( _
    ByVal ws As Worksheet)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo

    state = managerForm.BeginBusy()

    managerSheet.PreparePlantCell ws, PLANT_CELL
    managerSheet.PrepareDateCell ws, DATE_CELL, Date
    managerList.ApplyDropdown ws.Range(SEARCH_CELL), managerList.LIST_SPAREPART, False
    PrepareTable GetTable(ws)

    managerForm.EndBusy state
    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


Public Sub OnDoubleClick( _
    ByVal ws As Worksheet, _
    ByVal target As Range, _
    ByRef cancel As Boolean)

    Dim tbl As ListObject

    If managerForm.HandleDatePickerDoubleClick(target, ws.Range(DATE_CELL), cancel) Then
        Exit Sub
    End If

    If managerSheet.IsHit(target, ws, BUTTON_SAVE) Then
        cancel = True
        SaveOrders ws
    ElseIf managerSheet.IsHit(target, ws, BUTTON_ADD_ROW) Then
        cancel = True
        AddRow ws
    ElseIf managerSheet.IsHit(target, ws, BUTTON_CLEAR) Then
        cancel = True
        ClearTable ws, True
    ElseIf managerSheet.IsHit(target, ws, BUTTON_EMAIL) Then
        cancel = True
        SendEmail ws
    ElseIf managerSheet.IsHit(target, ws, BUTTON_SEARCH) Then
        cancel = True
        controllerShared.SearchSparepart ws, PLANT_CELL, SEARCH_CELL, RESULT_FIRST_CELL
    Else
        Set tbl = GetTable(ws)
        If managerTable.IsInColumn(target.Cells(1, 1), tbl, COL_DELETE) Then
            cancel = True
            DeleteRow ws, managerTable.GetRowIndex(tbl, target.Cells(1, 1))
        End If
    End If

End Sub


Public Sub OnChange( _
    ByVal ws As Worksheet, _
    ByVal target As Range)

    controllerShared.FillUnits GetTable(ws), target, COL_SPAREPART, COL_UNIT

End Sub


' =========================================================
' ACTIONS
' =========================================================

Public Sub AddRow( _
    ByVal ws As Worksheet)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo
    Dim tbl As ListObject
    Dim newRow As Long
    Dim previousPic As String

    state = managerForm.BeginBusy()

    Set tbl = GetTable(ws)

    previousPic = managerTable.GetText(tbl, managerTable.RowCount(tbl), COL_PIC)

    newRow = managerTable.AddStyledRow(tbl)

    managerTable.SetValue tbl, newRow, COL_DELETE, managerSheet.LABEL_DELETE
    managerTable.SetValue tbl, newRow, COL_PIC, previousPic

    PrepareTable tbl

    managerForm.EndBusy state

    managerTable.SelectCell tbl, newRow, COL_SPAREPART

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


Public Sub ClearTable( _
    ByVal ws As Worksheet, _
    ByVal askFirst As Boolean)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo

    If askFirst Then
        If Not managerForm.Confirm("Hapus semua baris pada tabel order?") Then Exit Sub
    End If

    state = managerForm.BeginBusy()

    managerTable.SetRowCount GetTable(ws), 1
    ws.Range(RESULT_FIRST_CELL).Resize(1, 5).ClearContents
    PrepareTable GetTable(ws)

    managerForm.EndBusy state
    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


Public Sub DeleteRow( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo
    Dim tbl As ListObject

    Set tbl = GetTable(ws)
    If rowIndex = 0 Then Exit Sub

    If controllerShared.RowHasInput(tbl, rowIndex, InputColumns()) Then
        If Not managerForm.Confirm("Hapus baris ke-" & rowIndex & "?") Then Exit Sub
    End If

    state = managerForm.BeginBusy()
    controllerShared.DeleteInputRow tbl, rowIndex, COL_DELETE, managerSheet.LABEL_DELETE
    managerForm.EndBusy state

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


Public Sub SaveOrders( _
    ByVal ws As Worksheet)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo
    Dim tbl As ListObject
    Dim plantName As String
    Dim orderDate As Date
    Dim spareparts As Collection
    Dim vendors As Collection
    Dim orders As Collection
    Dim orderItem As Order
    Dim uuids As Collection
    Dim rowIndex As Long
    Dim message As String
    Dim uuid As Variant

    Set tbl = GetTable(ws)

    plantName = managerSheet.ReadText(ws, PLANT_CELL, "LOKASI")
    orderDate = managerSheet.ReadDate(ws, DATE_CELL, "TANGGAL ORDER")

    state = managerForm.BeginBusy("Memeriksa order...")

    Set spareparts = serviceSparepart.ListSpareparts()
    Set vendors = serviceVendor.ListVendors()
    Set orders = New Collection

    For rowIndex = 1 To managerTable.RowCount(tbl)

        If controllerShared.RowHasInput(tbl, rowIndex, InputColumns()) Then

            Set orderItem = New Order

            Set orderItem.Sparepart = controllerShared.RequireSparepart( _
                spareparts, _
                managerSheet.TableRequiredText(tbl, rowIndex, COL_SPAREPART), _
                managerTable.GetCell(tbl, rowIndex, COL_SPAREPART))

            Set orderItem.Vendor = controllerShared.FindVendor( _
                vendors, _
                managerTable.GetText(tbl, rowIndex, COL_VENDOR), _
                managerTable.GetCell(tbl, rowIndex, COL_VENDOR))

            orderItem.OrderDate = orderDate
            orderItem.Quantity = managerSheet.TableWholeNumber(tbl, rowIndex, COL_QUANTITY, 1)
            orderItem.Remarks = managerTable.GetText(tbl, rowIndex, COL_REMARKS)
            orderItem.PersonInCharge = managerSheet.TableRequiredText(tbl, rowIndex, COL_PIC)

            orders.Add orderItem

        End If

    Next rowIndex

    managerForm.EndBusy state

    If orders.Count = 0 Then
        managerError.ThrowError controllerInputOrder.NoOrderRow, _
            MODULE_NAME & ".SaveOrders", _
            "Isi minimal satu baris order."
    End If

    If Not managerForm.Confirm( _
        "Simpan " & orders.Count & " order untuk " & plantName & _
        " tanggal " & Format$(orderDate, managerSheet.DATE_FORMAT) & "?") Then Exit Sub

    state = managerForm.BeginBusy("Menyimpan order...", True)

    Set uuids = serviceOrder.CreateOrders(plantName, orders)

    Set mLastSaved = orders
    mLastSavedPlant = plantName

    managerTable.SetRowCount tbl, 1
    PrepareTable tbl

    managerForm.EndBusy state

    message = orders.Count & " order tersimpan:" & vbCrLf
    For Each uuid In uuids
        message = message & "  " & CStr(uuid) & vbCrLf
    Next uuid

    managerForm.ShowInfo message

    If managerForm.Confirm("Buat email order sekarang?") Then
        OpenOrderEmail mLastSaved, mLastSavedPlant
    End If

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


' Email draft (.eml) with the rows in the table, or (when the
' table is empty) with the orders of the last SIMPAN.
Public Sub SendEmail( _
    ByVal ws As Worksheet)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo
    Dim tbl As ListObject
    Dim orders As Collection
    Dim orderItem As Order
    Dim vendors As Collection
    Dim rowIndex As Long
    Dim sparepartItem As Sparepart

    Set tbl = GetTable(ws)
    Set orders = New Collection

    For rowIndex = 1 To managerTable.RowCount(tbl)

        If controllerShared.RowHasInput(tbl, rowIndex, InputColumns()) Then

            ' Unsaved rows: only what is needed for the mail.
            Set sparepartItem = New Sparepart
            sparepartItem.Name = managerSheet.TableRequiredText(tbl, rowIndex, COL_SPAREPART)
            sparepartItem.UnitOfMeasurement = managerTable.GetText(tbl, rowIndex, COL_UNIT)

            Set orderItem = New Order
            Set orderItem.Sparepart = sparepartItem

            If managerTable.HasValue(tbl, rowIndex, COL_VENDOR) Then
                Set orderItem.Vendor = New Vendor
                orderItem.Vendor.Name = managerTable.GetText(tbl, rowIndex, COL_VENDOR)
            End If

            orderItem.Quantity = managerSheet.TableWholeNumber(tbl, rowIndex, COL_QUANTITY, 0)
            orderItem.Remarks = managerTable.GetText(tbl, rowIndex, COL_REMARKS)
            If managerTable.HasValue(tbl, rowIndex, COL_PIC) Then
                orderItem.PersonInCharge = managerTable.GetText(tbl, rowIndex, COL_PIC)
            End If

            orders.Add orderItem

        End If

    Next rowIndex

    If orders.Count > 0 Then
        OpenOrderEmail orders, managerForm.GetCellText(ws, PLANT_CELL)
        Exit Sub
    End If

    If Not mLastSaved Is Nothing Then
        If mLastSaved.Count > 0 Then
            OpenOrderEmail mLastSaved, mLastSavedPlant
            Exit Sub
        End If
    End If

    managerError.ThrowError controllerInputOrder.NothingToEmail, _
        MODULE_NAME & ".SendEmail", _
        "Tidak ada order untuk dikirim. Isi tabel order atau simpan order terlebih dahulu."

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerError.ShowError errInfo

End Sub


' =========================================================
' INTERNAL
' =========================================================

Private Function GetTable( _
    ByVal ws As Worksheet) As ListObject

    Set GetTable = managerTable.GetTable(ws, TABLE_NAME)

End Function


' Columns that make a row count as an order. PIC is left out:
' TAMBAH BARIS copies it from the row above.
Private Function InputColumns() As Variant

    InputColumns = Array(COL_SPAREPART, COL_QUANTITY, COL_VENDOR, COL_REMARKS)

End Function


' Labels, dropdowns and number rule of the input table.
Private Sub PrepareTable( _
    ByVal tbl As ListObject)

    managerSheet.FillColumn tbl, COL_DELETE, managerSheet.LABEL_DELETE

    managerList.ApplyDropdown managerTable.GetColumnRange(tbl, COL_SPAREPART), _
        managerList.LIST_SPAREPART
    managerList.ApplyDropdown managerTable.GetColumnRange(tbl, COL_VENDOR), _
        managerList.LIST_VENDOR
    managerList.ApplyWholeNumberRule managerTable.GetColumnRange(tbl, COL_QUANTITY), 1

End Sub


Private Sub OpenOrderEmail( _
    ByVal orders As Collection, _
    ByVal plantName As String)

    Dim html As String
    Dim orderItem As Order
    Dim number As Long
    Dim vendorName As String
    Dim greetingName As String

    greetingName = managerEmail.FirstRecipientName()
    If Len(greetingName) = 0 Then greetingName = "Bapak/Ibu"

    html = "<p>Dengan hormat " & managerEmail.HtmlText(greetingName) & ",</p>" & _
        "<p>Berikut order sparepart " & managerEmail.HtmlText(plantName) & _
        " tanggal " & Format$(Date, managerSheet.DATE_FORMAT) & ":</p>" & _
        "<table border=""1"" cellpadding=""4"" cellspacing=""0"" " & _
        "style=""border-collapse:collapse;font-family:Calibri,Arial;font-size:11pt"">" & _
        "<tr style=""background:#DCE6F1""><th>NO</th><th>NO ORDER</th>" & _
        "<th>NAMA SPAREPART</th><th>JUMLAH</th><th>SATUAN</th><th>VENDOR</th>" & _
        "<th>REMARKS</th><th>PIC</th></tr>"

    For Each orderItem In orders

        number = number + 1

        vendorName = vbNullString
        If orderItem.HasVendor Then vendorName = orderItem.Vendor.DisplayName

        html = html & "<tr><td>" & number & "</td>" & _
            "<td>" & managerEmail.HtmlText(orderItem.UUID) & "</td>" & _
            "<td>" & managerEmail.HtmlText(orderItem.Sparepart.DisplayName) & "</td>" & _
            "<td style=""text-align:right"">" & orderItem.Quantity & "</td>" & _
            "<td>" & managerEmail.HtmlText(orderItem.UnitOfMeasurement) & "</td>" & _
            "<td>" & managerEmail.HtmlText(vendorName) & "</td>" & _
            "<td>" & managerEmail.HtmlText(orderItem.Remarks) & "</td>" & _
            "<td>" & managerEmail.HtmlText(orderItem.PersonInCharge) & "</td></tr>"

    Next orderItem

    html = html & "</table><p>Terima kasih.</p>"

    managerEmail.OpenDraft _
        "Order Sparepart " & plantName & " - " & Format$(Date, managerSheet.DATE_FORMAT), _
        html

End Sub
