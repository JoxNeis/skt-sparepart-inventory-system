Attribute VB_Name = "managerEmail"
Option Explicit

' =========================================================
' EMAIL MANAGER
'
' Writes the mail as an unsent .eml file and opens it with the
' default mail app (the new Outlook, classic Outlook, ...).
' It never sends by itself: the user checks it and presses
' Send. Recipients come from the TO / CC / BCC tables on the
' SETTINGS sheet.
'
' The file has "X-Unsent: 1", so mail apps open it as a new
' message instead of a received one. Files are kept in
' %TEMP%\SIS Email.
' =========================================================

Private Const MODULE_NAME As String = "managerEmail"

' =========================================================
' ERROR CODES (10900 - 10999)
' =========================================================
Public Const EmailFileFailed As Long = 10901
Public Const NoRecipient As Long = 10902

Public Const TABLE_EMAIL_TO As String = "TableEmailTo"
Public Const TABLE_EMAIL_CC As String = "TableEmailCC"
Public Const TABLE_EMAIL_BCC As String = "TableEmailBCC"

Private Const EMAIL_FOLDER As String = "SIS Email"
Private Const AD_TYPE_BINARY As Long = 1
Private Const AD_TYPE_TEXT As Long = 2
Private Const AD_SAVE_CREATE_OVERWRITE As Long = 2


' Write a new mail with the SETTINGS recipients to an .eml
' file and open it.
Public Sub OpenDraft( _
    ByVal subject As String, _
    ByVal htmlBody As String)

    Dim toList As String
    Dim filePath As String
    Dim message As String

    toList = Recipients(TABLE_EMAIL_TO)

    If Len(toList) = 0 Then
        managerError.ThrowError managerEmail.NoRecipient, _
            MODULE_NAME & ".OpenDraft", _
            "Belum ada penerima email (TO) di sheet SETTINGS."
    End If

    message = BuildMessage(toList, _
        Recipients(TABLE_EMAIL_CC), _
        Recipients(TABLE_EMAIL_BCC), _
        subject, htmlBody)

    filePath = NewFilePath(subject)

    On Error GoTo FileFailed
    WriteAscii filePath, message
    CreateObject("Shell.Application").ShellExecute filePath, "", "", "open", 1
    Exit Sub

FileFailed:
    managerError.ThrowError managerEmail.EmailFileFailed, _
        MODULE_NAME & ".OpenDraft", _
        "File email tidak dapat dibuat atau dibuka: " & filePath & _
        " (" & Err.Description & ")"

End Sub


' The whole .eml text (RFC 5322 / MIME, HTML body in UTF-8).
Private Function BuildMessage( _
    ByVal toList As String, _
    ByVal ccList As String, _
    ByVal bccList As String, _
    ByVal subject As String, _
    ByVal htmlBody As String) As String

    Dim result As String

    result = "X-Unsent: 1" & vbCrLf & _
        "To: " & toList & vbCrLf

    If Len(ccList) > 0 Then result = result & "Cc: " & ccList & vbCrLf
    If Len(bccList) > 0 Then result = result & "Bcc: " & bccList & vbCrLf

    result = result & _
        "Subject: =?UTF-8?B?" & Base64(Utf8Bytes(subject), False) & "?=" & vbCrLf & _
        "MIME-Version: 1.0" & vbCrLf & _
        "Content-Type: text/html; charset=""utf-8""" & vbCrLf & _
        "Content-Transfer-Encoding: base64" & vbCrLf & _
        vbCrLf & _
        Base64(Utf8Bytes( _
            "<html><head><meta charset=""utf-8""></head><body>" & _
            htmlBody & "</body></html>"), True) & vbCrLf

    BuildMessage = result

End Function


' %TEMP%\SIS Email\<subject> yyyymmdd-hhnnss.eml
Private Function NewFilePath( _
    ByVal subject As String) As String

    Dim folder As String
    Dim fileName As String
    Dim badChars As Variant
    Dim i As Long

    folder = Environ$("TEMP") & "\" & EMAIL_FOLDER
    If Len(Dir$(folder, vbDirectory)) = 0 Then MkDir folder

    fileName = subject
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    For i = LBound(badChars) To UBound(badChars)
        fileName = Replace(fileName, badChars(i), "-")
    Next i

    If Len(fileName) > 80 Then fileName = Left$(fileName, 80)

    NewFilePath = folder & "\" & Trim$(fileName) & " " & _
        Format$(Now, "yyyymmdd-hhnnss") & ".eml"

End Function


' UTF-8 bytes of text (without the byte order mark).
Private Function Utf8Bytes( _
    ByVal text As String) As Variant

    Dim stream As Object

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = AD_TYPE_TEXT
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText text
    stream.Position = 0
    stream.Type = AD_TYPE_BINARY
    stream.Position = 3
    Utf8Bytes = stream.Read
    stream.Close

End Function


' Base64 of bytes; with lineBreaks, CRLF every 76 characters
' (as MIME bodies need).
Private Function Base64( _
    ByVal bytes As Variant, _
    ByVal lineBreaks As Boolean) As String

    Dim node As Object
    Dim text As String
    Dim result As String
    Dim i As Long

    If IsNull(bytes) Then Exit Function
    If LenB(bytes) = 0 Then Exit Function

    Set node = CreateObject("MSXML2.DOMDocument").createElement("b64")
    node.DataType = "bin.base64"
    node.nodeTypedValue = bytes

    text = Replace(Replace(node.text, vbCr, vbNullString), vbLf, vbNullString)

    If Not lineBreaks Then
        Base64 = text
        Exit Function
    End If

    For i = 1 To Len(text) Step 76
        If Len(result) > 0 Then result = result & vbCrLf
        result = result & Mid$(text, i, 76)
    Next i

    Base64 = result

End Function


' The message is plain ASCII (non-ASCII text is base64), so it
' is written byte for byte.
Private Sub WriteAscii( _
    ByVal filePath As String, _
    ByVal text As String)

    Dim stream As Object

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = AD_TYPE_TEXT
    stream.Charset = "us-ascii"
    stream.Open
    stream.WriteText text
    stream.SaveToFile filePath, AD_SAVE_CREATE_OVERWRITE
    stream.Close

End Sub


' "a@x.com; b@y.com" from the EMAIL column of one table.
Public Function Recipients( _
    ByVal tableName As String) As String

    Dim tbl As ListObject
    Dim rowIndex As Long
    Dim address As String
    Dim result As String

    Set tbl = managerTable.GetTable( _
        managerSheet.GetSheet(managerSheet.SHEET_SETTINGS), tableName)

    For rowIndex = 1 To managerTable.RowCount(tbl)

        address = managerTable.GetText(tbl, rowIndex, "EMAIL")

        If Len(address) > 0 Then
            If Len(result) > 0 Then result = result & "; "
            result = result & address
        End If

    Next rowIndex

    Recipients = result

End Function


' First name (NAMA) of the TO table, for the greeting.
Public Function FirstRecipientName() As String

    Dim tbl As ListObject

    Set tbl = managerTable.GetTable( _
        managerSheet.GetSheet(managerSheet.SHEET_SETTINGS), TABLE_EMAIL_TO)

    If managerTable.RowCount(tbl) > 0 Then
        FirstRecipientName = managerTable.GetText(tbl, 1, "NAMA")
    End If

End Function


' &, <, > and quotes escaped for HTML text.
Public Function HtmlText( _
    ByVal value As Variant) As String

    Dim text As String

    text = coreHelper.TextOrEmpty(value)
    text = Replace(text, "&", "&amp;")
    text = Replace(text, "<", "&lt;")
    text = Replace(text, ">", "&gt;")
    text = Replace(text, """", "&quot;")

    HtmlText = text

End Function
