Attribute VB_Name = "coreConfig"
Option Explicit

' Error codes live in each owner module (see managerError.bas).

Private Const DATABASE_FILE_NAME As String = "SIS.accdb"

Public Function currentUser() As String
    currentUser = Environ$("USERNAME")
End Function

Public Function currentExcelPath() As String
    currentExcelPath = ThisWorkbook.Path
End Function

Public Function databasePath() As String

    Dim excelPath As String
    Dim pathCol As Collection
    Dim part As String
    Dim result As String
    Dim i As Long

    excelPath = currentExcelPath()

    If Len(excelPath) = 0 Or LCase$(Left$(excelPath, 4)) = "http" Then
        databasePath = "C:\Users\" & currentUser() & "\OneDrive - Philip Morris International\Sparepart Inventory System\System\Sparepart Inventory System Ver 5\Sparepart Inventory System Ver 5.accdb"
        Exit Function
    End If

    Set pathCol = coreHelper.SplitString(excelPath, "\")

    For i = 1 To pathCol.Count
        part = pathCol(i)

        If i > 1 Then
            If StrComp(pathCol(i - 1), "Users", vbTextCompare) = 0 Then
                part = currentUser()
            End If
        End If

        If i = pathCol.Count And LCase$(part) Like "*.xls?" Then Exit For

        result = result & part & "\"
    Next i

    databasePath = result & DATABASE_FILE_NAME

End Function

Public Function ApplicationName() As String
    ApplicationName = "Sparepart Inventory System"
End Function
