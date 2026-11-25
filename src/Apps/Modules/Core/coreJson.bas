Attribute VB_Name = "coreJson"
Option Explicit

' =========================================================
' JSON
'
' Small builders for activity_logs.json. Every function
' returns a ready JSON fragment, so they nest:
'
'   json = coreJson.JsonObject( _
'       "uuid", coreJson.JsonText(uuid), _
'       "sparepart", coreJson.JsonObject( _
'           "code", coreJson.JsonText(code), _
'           "name", coreJson.JsonText(name)), _
'       "order_date", coreJson.JsonDate(orderDate), _
'       "quantity", coreJson.JsonNumber(quantity))
'
' Numbers are written with a dot as decimal separator, also
' on Windows with an Indonesian locale (decimal comma).
' =========================================================


' "text" (escaped). Null / Empty -> "".
Public Function JsonText( _
    ByVal value As Variant) As String

    If IsNull(value) Or IsEmpty(value) Then
        JsonText = """"""
    Else
        JsonText = """" & Escape(CStr(value)) & """"
    End If

End Function


' 12 / 1.65. Null / Empty / non-numeric -> null.
Public Function JsonNumber( _
    ByVal value As Variant) As String

    If IsNull(value) Or IsEmpty(value) Then
        JsonNumber = "null"
    ElseIf Not IsNumeric(value) Then
        JsonNumber = "null"
    Else
        JsonNumber = Trim$(Str$(CDbl(value)))
    End If

End Function


' "yyyy-mm-dd". 0 (empty date) -> null.
Public Function JsonDate( _
    ByVal value As Date) As String

    If value = 0 Then
        JsonDate = "null"
    Else
        JsonDate = """" & Format$(value, "yyyy-mm-dd") & """"
    End If

End Function


' "yyyy-mm-dd hh:nn:ss". 0 (empty date) -> null.
Public Function JsonDateTime( _
    ByVal value As Date) As String

    If value = 0 Then
        JsonDateTime = "null"
    Else
        JsonDateTime = """" & Format$(value, "yyyy-mm-dd hh:nn:ss") & """"
    End If

End Function


Public Function JsonBoolean( _
    ByVal value As Boolean) As String

    If value Then
        JsonBoolean = "true"
    Else
        JsonBoolean = "false"
    End If

End Function


' {"key": value, ...} from alternating key / JSON fragment
' arguments. Fragments must already be JSON (use the
' functions above).
Public Function JsonObject( _
    ParamArray keyValuePairs() As Variant) As String

    Dim result As String
    Dim i As Long

    result = "{"

    For i = LBound(keyValuePairs) To UBound(keyValuePairs) - 1 Step 2

        If i > LBound(keyValuePairs) Then result = result & ","

        result = result & _
            """" & Escape(CStr(keyValuePairs(i))) & """:" & _
            CStr(keyValuePairs(i + 1))

    Next i

    JsonObject = result & "}"

End Function


' Add one more "key": value pair to an object made by
' JsonObject.
Public Function JsonAppend( _
    ByVal jsonObjectText As String, _
    ByVal key As String, _
    ByVal fragment As String) As String

    Dim body As String

    body = Left$(jsonObjectText, Len(jsonObjectText) - 1)

    If Len(body) > 1 Then body = body & ","

    JsonAppend = body & """" & Escape(key) & """:" & fragment & "}"

End Function


' =========================================================
' INTERNAL
' =========================================================

Private Function Escape( _
    ByVal value As String) As String

    Dim result As String
    Dim i As Long
    Dim ch As String
    Dim code As Long

    For i = 1 To Len(value)

        ch = Mid$(value, i, 1)
        code = AscW(ch)

        Select Case code
            Case 34
                result = result & "\"""
            Case 92
                result = result & "\\"
            Case 8
                result = result & "\b"
            Case 9
                result = result & "\t"
            Case 10
                result = result & "\n"
            Case 12
                result = result & "\f"
            Case 13
                result = result & "\r"
            Case 0 To 31
                result = result & "\u" & Right$("000" & Hex$(code), 4)
            Case Else
                result = result & ch
        End Select

    Next i

    Escape = result

End Function
