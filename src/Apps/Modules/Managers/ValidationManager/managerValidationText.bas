Attribute VB_Name = "managerValidationText"
Option Explicit

' =========================================================
' VALIDATION MANAGER - TEXT
'
' Every function returns a Boolean and never raises an
' error. The caller decides what to do with the result:
'
'   If managerValidationText.IsBlank(value) Then
'       managerError.ThrowError _
'           entitiesOrder.EmptyUUID, _
'           CLASS_NAME & ".UUID", _
'           "UUID tidak boleh kosong."
'   End If
' =========================================================


' =========================================================
' True when the value is Empty, Null, an error value, or
' text that contains only spaces.
' =========================================================
Public Function IsBlank( _
    ByVal value As Variant) As Boolean

    If IsObject(value) Then
        IsBlank = (value Is Nothing)
        Exit Function
    End If

    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        IsBlank = True
        Exit Function
    End If

    If IsArray(value) Then
        IsBlank = False
        Exit Function
    End If

    IsBlank = (Len(Trim$(CStr(value))) = 0)

End Function


' =========================================================
' Opposite of IsBlank.
' =========================================================
Public Function IsNotBlank( _
    ByVal value As Variant) As Boolean

    IsNotBlank = Not IsBlank(value)

End Function


' =========================================================
' True when the trimmed text is longer than maxLength.
' =========================================================
Public Function IsLongerThan( _
    ByVal value As Variant, _
    ByVal maxLength As Long) As Boolean

    If IsBlank(value) Then
        IsLongerThan = False
        Exit Function
    End If

    IsLongerThan = (Len(Trim$(CStr(value))) > maxLength)

End Function


' =========================================================
' True when the trimmed text is shorter than minLength.
' A blank value counts as length 0.
' =========================================================
Public Function IsShorterThan( _
    ByVal value As Variant, _
    ByVal minLength As Long) As Boolean

    If IsBlank(value) Then
        IsShorterThan = (minLength > 0)
        Exit Function
    End If

    IsShorterThan = (Len(Trim$(CStr(value))) < minLength)

End Function


' =========================================================
' True when both texts are equal.
' ignoreCase = True uses vbTextCompare.
' =========================================================
Public Function IsSameText( _
    ByVal firstValue As String, _
    ByVal secondValue As String, _
    Optional ByVal ignoreCase As Boolean = True) As Boolean

    If ignoreCase Then
        IsSameText = (StrComp(firstValue, secondValue, vbTextCompare) = 0)
    Else
        IsSameText = (StrComp(firstValue, secondValue, vbBinaryCompare) = 0)
    End If

End Function


' =========================================================
' True when the trimmed value matches one of the items in
' allowedValues (an Array).
'
'   If Not managerValidationText.IsInList(value, _
'       Array("DATANG", "DATANG SEBAGIAN", "BELUM DATANG")) Then
' =========================================================
Public Function IsInList( _
    ByVal value As Variant, _
    ByVal allowedValues As Variant, _
    Optional ByVal ignoreCase As Boolean = True) As Boolean

    Dim item As Variant
    Dim textValue As String

    IsInList = False

    If IsBlank(value) Then Exit Function
    If Not IsArray(allowedValues) Then Exit Function

    textValue = Trim$(CStr(value))

    For Each item In allowedValues
        If Not IsBlank(item) Then
            If IsSameText(textValue, Trim$(CStr(item)), ignoreCase) Then
                IsInList = True
                Exit Function
            End If
        End If
    Next item

End Function
