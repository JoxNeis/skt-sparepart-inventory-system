Attribute VB_Name = "sheetRekapInventory"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

' =========================================================
' SHEET - REKAP_INVENTORY
'
' Events only. The work is done by controllerRekapInventory.
' Clickable cells are double-clicked; the sidebar (B:F)
' opens the other sheets.
' =========================================================


Private Sub Worksheet_Activate()

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    controllerRekapInventory.OnActivate Me

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    Application.EnableEvents = True
    managerError.ShowError errInfo

End Sub


Private Sub Worksheet_BeforeDoubleClick( _
    ByVal Target As Range, _
    Cancel As Boolean)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    If managerNavigation.HandleDoubleClick(Target, Cancel) Then Exit Sub
    controllerRekapInventory.OnDoubleClick Me, Target, Cancel

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    Application.EnableEvents = True
    managerError.ShowError errInfo

End Sub


Private Sub Worksheet_Change( _
    ByVal Target As Range)

    On Error GoTo ErrorHandler

    Dim errInfo As ErrorInfo

    controllerRekapInventory.OnChange Me, Target

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    Application.EnableEvents = True
    managerError.ShowError errInfo

End Sub
