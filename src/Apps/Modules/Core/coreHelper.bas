Attribute VB_Name = "coreHelper"
Option Explicit

Private Const MODULE_NAME As String = "HELPER"

' =========================================================
' ERROR CODES (10200 - 10299)
' Usage: managerError.ThrowError coreHelper.InvalidRecordsetNumber, ...
' =========================================================
Public Const InvalidRecordsetNumber As Long = 10201  ' database value is not numeric

Public Function TrimText(ByVal value As String) As String
    TrimText = Trim$(value)
End Function

Public Function RecordsetText(ByVal value As Variant) As String

    If IsNull(value) Or IsEmpty(value) Then
        RecordsetText = vbNullString
    Else
        RecordsetText = Trim$(CStr(value))
    End If

End Function

Public Function RecordsetDouble(ByVal value As Variant) As Double

    If IsNull(value) Or IsEmpty(value) Then
        RecordsetDouble = 0
        Exit Function
    End If

    ' Not IsNumeric (instead of managerValidationNumeric.IsNotNumber) keeps
    ' accepting Yes/No (Boolean) fields as -1 / 0, as before.
    If Not IsNumeric(value) Then
        managerError.ThrowError _
            coreHelper.InvalidRecordsetNumber, _
            MODULE_NAME & ".RecordsetDouble", _
            "Nilai database bukan angka yang valid."
    End If

    RecordsetDouble = CDbl(value)

End Function


Public Function FixedWidth( _
    ByVal value As Variant, _
    ByVal width As Long, _
    Optional ByVal alignRight As Boolean = False) As String

    Dim text As String

    If IsNull(value) Or IsEmpty(value) Then
        text = vbNullString
    Else
        text = CStr(value)
    End If

    If Len(text) > width - 1 Then
        text = Left$(text, width - 4) & "..."
    End If

    If alignRight Then
        FixedWidth = Space$(width - Len(text)) & text
    Else
        FixedWidth = text & Space$(width - Len(text))
    End If

End Function

Public Function TextOrEmpty( _
    ByVal value As Variant) As String

    If IsNull(value) Or IsEmpty(value) Then
        TextOrEmpty = vbNullString
    Else
        TextOrEmpty = CStr(value)
    End If

End Function

Public Function SplitString( _
    ByVal value As String, _
    ByVal character As String) As Collection

    Dim result As Collection
    Dim parts() As String
    Dim i As Long

    Set result = New Collection
    parts = Split(value, character)

    For i = LBound(parts) To UBound(parts)
        result.Add parts(i)
    Next i

    Set SplitString = result

End Function


