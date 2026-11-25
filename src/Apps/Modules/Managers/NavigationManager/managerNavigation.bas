Attribute VB_Name = "managerNavigation"
Option Explicit

' =========================================================
' NAVIGATION MANAGER (sidebar)
'
' Every sheet has the same sidebar in columns B:F. A menu
' row holds its target as text, e.g. "REKAP_ORDER!A1", in
' one of its cells. Double-clicking anywhere on the row opens
' that sheet. Rows without a target text fall back to the
' fixed menu map below.
'
'   Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
'       If managerNavigation.HandleDoubleClick(Target, Cancel) Then Exit Sub
'       ...
'   End Sub
' =========================================================

Private Const MODULE_NAME As String = "managerNavigation"

Private Const SIDEBAR_FIRST_COLUMN As Long = 2     ' B
Private Const SIDEBAR_LAST_COLUMN As Long = 6      ' F
Private Const SIDEBAR_FIRST_ROW As Long = 5
Private Const SIDEBAR_LAST_ROW As Long = 32


' Returns True (and sets cancel) when target is a sidebar
' menu row; the target sheet is then activated.
Public Function HandleDoubleClick( _
    ByVal target As Range, _
    ByRef cancel As Boolean) As Boolean

    Dim sheetName As String

    HandleDoubleClick = False

    With target.Cells(1, 1)
        If .Column < SIDEBAR_FIRST_COLUMN Or .Column > SIDEBAR_LAST_COLUMN Then Exit Function
        If .Row < SIDEBAR_FIRST_ROW Or .Row > SIDEBAR_LAST_ROW Then Exit Function
    End With

    sheetName = TargetSheetName(target.Worksheet, target.Cells(1, 1).Row)

    If Len(sheetName) = 0 Then Exit Function

    cancel = True
    HandleDoubleClick = True

    GoToSheet sheetName

End Function


' Activate a sheet by name (silently ignored when missing).
Public Sub GoToSheet( _
    ByVal sheetName As String)

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        managerForm.ShowWarning "Sheet '" & sheetName & "' tidak ditemukan."
        Exit Sub
    End If

    If ws.Visible <> xlSheetVisible Then Exit Sub

    ws.Activate
    ws.Range("A1").Select

End Sub


' =========================================================
' INTERNAL
' =========================================================

Private Function TargetSheetName( _
    ByVal ws As Worksheet, _
    ByVal rowNumber As Long) As String

    Dim columnNumber As Long
    Dim value As Variant
    Dim text As String
    Dim position As Long

    For columnNumber = SIDEBAR_FIRST_COLUMN To SIDEBAR_LAST_COLUMN

        value = ws.Cells(rowNumber, columnNumber).value

        If Not IsError(value) Then

            text = Trim$(coreHelper.TextOrEmpty(value))
            position = InStr(1, text, "!")

            If position > 1 Then
                text = Replace(Left$(text, position - 1), "'", vbNullString)
                If SheetExists(text) Then
                    TargetSheetName = text
                    Exit Function
                End If
            End If

        End If

    Next columnNumber

    TargetSheetName = FallbackSheetName(rowNumber)

End Function


' The fixed menu of the sidebar design.
Private Function FallbackSheetName( _
    ByVal rowNumber As Long) As String

    Select Case rowNumber
        Case 5:      FallbackSheetName = managerSheet.SHEET_DASHBOARD
        Case 7, 9:   FallbackSheetName = managerSheet.SHEET_REKAP_ORDER
        Case 8:      FallbackSheetName = managerSheet.SHEET_INPUT_ORDER
        Case 10:     FallbackSheetName = managerSheet.SHEET_REKAP_PENDING_ORDER
        Case 12, 14: FallbackSheetName = managerSheet.SHEET_REKAP_PENERIMAAN
        Case 13:     FallbackSheetName = managerSheet.SHEET_INPUT_PENERIMAAN
        Case 16, 18: FallbackSheetName = managerSheet.SHEET_REKAP_REQUEST
        Case 17:     FallbackSheetName = managerSheet.SHEET_INPUT_REQUEST
        Case 20, 21: FallbackSheetName = managerSheet.SHEET_REKAP_INVENTORY
        Case 22:     FallbackSheetName = managerSheet.SHEET_KARTU_STOCK
        Case 25:     FallbackSheetName = managerSheet.SHEET_MD_SPAREPART
        Case 26:     FallbackSheetName = managerSheet.SHEET_MD_POS_MOBILE
        Case 27:     FallbackSheetName = managerSheet.SHEET_MD_VENDOR
        Case 28:     FallbackSheetName = managerSheet.SHEET_MD_USER
        Case 30:     FallbackSheetName = managerSheet.SHEET_REQUEST_UPDATE
        Case 32:     FallbackSheetName = managerSheet.SHEET_SETTINGS
        Case Else:   FallbackSheetName = vbNullString
    End Select

End Function


Private Function SheetExists( _
    ByVal sheetName As String) As Boolean

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    SheetExists = Not (ws Is Nothing)

End Function
