Attribute VB_Name = "pConfig"
Option Explicit

' Error codes live in each owner module (see ErrorManager.bas).

Public Function currentUser() As String
    currentUser = Environ$("USERNAME")
End Function

Public Function databasePath() As String
    databasePath = "C:\Users\" & currentUser() & "\OneDrive - Philip Morris International\Sparepart Inventory System\System\Sparepart Inventory System Ver 5\Sparepart Inventory System Ver 5.accdb"
End Function

Public Function ApplicationName() As String
    ApplicationName = "Sparepart Inventory System"
End Function
