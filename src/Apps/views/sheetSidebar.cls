Attribute VB_Name = "sheetSidebar"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

' =========================================================
' SHEET - SIDEBAR
'
' Events only. Hidden template of the sidebar.
' Clickable cells are double-clicked; the sidebar (B:F)
' opens the other sheets.
' =========================================================


Private Sub Worksheet_BeforeDoubleClick( _
    ByVal Target As Range, _
    Cancel As Boolean)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    If managerNavigation.HandleDoubleClick(Target, Cancel) Then Exit Sub

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    Application.EnableEvents = True
    managerError.ShowError errInfo

End Sub
