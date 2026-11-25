Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

' =========================================================
' WORKBOOK
'
' Open : prepare the dropdown lists, load the master data
'        from the database and show the DASHBOARD
'        (coreApp.Initialize).
' Close: close the database connection.
' =========================================================


Private Sub Workbook_Open()

    coreApp.Initialize

End Sub


Private Sub Workbook_BeforeClose( _
    Cancel As Boolean)

    managerDatabase.DisconnectFromDatabase

End Sub
