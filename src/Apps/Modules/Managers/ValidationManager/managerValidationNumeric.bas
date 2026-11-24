Attribute VB_Name = "managerValidationNumeric"
Option Explicit

' =========================================================
' VALIDATION MANAGER - NUMERIC
'
' Every function returns a Boolean and never raises an
' error. The caller decides what to do with the result:
'
'   If managerValidationNumeric.IsNegative(value) Then
'       managerError.ThrowError _
'           entitiesReceive.NegativeOkAmount, _
'           CLASS_NAME & ".OKAmount", _
'           "Jumlah barang OK tidak boleh negatif."
'   End If
'
' decimals (optional): round both sides before comparing.
' Use it for Double values that can carry tiny floating
' point differences, e.g. decimals:=4. The default (-1)
' compares the values as they are.
' =========================================================


' =========================================================
' True when the value can be used as a number.
' Empty, Null, blank text, error values and Booleans are
' not numbers.
' =========================================================
Public Function IsNumber( _
    ByVal value As Variant) As Boolean

    IsNumber = False

    If IsObject(value) Then Exit Function
    If IsArray(value) Then Exit Function
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    If VarType(value) = vbBoolean Then Exit Function

    If VarType(value) = vbString Then
        If Len(Trim$(value)) = 0 Then Exit Function
    End If

    IsNumber = IsNumeric(value)

End Function


' =========================================================
' Opposite of IsNumber.
' =========================================================
Public Function IsNotNumber( _
    ByVal value As Variant) As Boolean

    IsNotNumber = Not IsNumber(value)

End Function


' =========================================================
' value < 0
' =========================================================
Public Function IsNegative( _
    ByVal value As Double, _
    Optional ByVal decimals As Long = -1) As Boolean

    IsNegative = (RoundValue(value, decimals) < 0)

End Function


' =========================================================
' value = 0
' =========================================================
Public Function IsZero( _
    ByVal value As Double, _
    Optional ByVal decimals As Long = -1) As Boolean

    IsZero = (RoundValue(value, decimals) = 0)

End Function


' =========================================================
' value <= 0
' =========================================================
Public Function IsZeroOrNegative( _
    ByVal value As Double, _
    Optional ByVal decimals As Long = -1) As Boolean

    IsZeroOrNegative = (RoundValue(value, decimals) <= 0)

End Function


' =========================================================
' value > 0
' =========================================================
Public Function IsPositive( _
    ByVal value As Double, _
    Optional ByVal decimals As Long = -1) As Boolean

    IsPositive = (RoundValue(value, decimals) > 0)

End Function


' =========================================================
' value > limit
' =========================================================
Public Function IsGreaterThan( _
    ByVal value As Double, _
    ByVal limit As Double, _
    Optional ByVal decimals As Long = -1) As Boolean

    IsGreaterThan = _
        (RoundValue(value, decimals) > RoundValue(limit, decimals))

End Function


' =========================================================
' value < limit
' =========================================================
Public Function IsLessThan( _
    ByVal value As Double, _
    ByVal limit As Double, _
    Optional ByVal decimals As Long = -1) As Boolean

    IsLessThan = _
        (RoundValue(value, decimals) < RoundValue(limit, decimals))

End Function


' =========================================================
' value < minValue Or value > maxValue
' =========================================================
Public Function IsOutOfRange( _
    ByVal value As Double, _
    ByVal minValue As Double, _
    ByVal maxValue As Double, _
    Optional ByVal decimals As Long = -1) As Boolean

    IsOutOfRange = _
        IsLessThan(value, minValue, decimals) Or _
        IsGreaterThan(value, maxValue, decimals)

End Function


' =========================================================
' True when the value has no fractional part.
' =========================================================
Public Function IsWholeNumber( _
    ByVal value As Double) As Boolean

    IsWholeNumber = (value = Fix(value))

End Function


' =========================================================
' Internal: round only when decimals >= 0.
' =========================================================
Private Function RoundValue( _
    ByVal value As Double, _
    ByVal decimals As Long) As Double

    If decimals >= 0 Then
        RoundValue = Round(value, decimals)
    Else
        RoundValue = value
    End If

End Function
