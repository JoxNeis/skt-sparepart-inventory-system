Option Explicit

' =========================================================
' EMAIL MANAGER
'
' Creates an unsent .eml file and opens it using the Windows
' default application for .eml files.
'
' The message is never sent automatically. The user reviews
' the message and presses Send.
'
' Recipients are read from the TO, CC, and BCC tables on the
' SETTINGS sheet.
'
' The file contains "X-Unsent: 1", so compatible mail apps
' open it as a new outgoing message rather than as a received
' message.
'
' Temporary files are stored in:
'   %TEMP%\SIS Email
'
' Public interface:
'   managerEmail.OpenDraft
'   managerEmail.Recipients
'   managerEmail.FirstRecipientName
'   managerEmail.HtmlText
' =========================================================

Private Const MODULE_NAME As String = "managerEmail"

' =========================================================
' ERROR CODES (10900 - 10999)
' =========================================================

Public Const EmailFileFailed As Long = 10901
Public Const NoRecipient As Long = 10902
Public Const InvalidEmailHeader As Long = 10903
Public Const EmailFolderFailed As Long = 10904
Public Const EmailOpenFailed As Long = 10905

' =========================================================
' SETTINGS TABLES
' =========================================================

Public Const TABLE_EMAIL_TO As String = "TableEmailTo"
Public Const TABLE_EMAIL_CC As String = "TableEmailCC"
Public Const TABLE_EMAIL_BCC As String = "TableEmailBCC"

' =========================================================
' FILE SETTINGS
' =========================================================

Private Const EMAIL_FOLDER As String = "SIS Email"
Private Const MAX_FILE_NAME_LENGTH As Long = 80

' =========================================================
' ADODB STREAM CONSTANTS
' =========================================================

Private Const AD_TYPE_BINARY As Long = 1
Private Const AD_TYPE_TEXT As Long = 2
Private Const AD_SAVE_CREATE_OVERWRITE As Long = 2

' =========================================================
' PUBLIC ENTRY POINT
' =========================================================

' Creates an unsent .eml draft using recipients stored in the
' SETTINGS tables, then opens it using the Windows default
' application for .eml files.
Public Sub OpenDraft( _
    ByVal subject As String, _
    ByVal htmlBody As String)

    Dim toList As String
    Dim ccList As String
    Dim bccList As String

    Dim filePath As String
    Dim message As String

    toList = Recipients(TABLE_EMAIL_TO)
    ccList = Recipients(TABLE_EMAIL_CC)
    bccList = Recipients(TABLE_EMAIL_BCC)

    If Len(Trim$(toList)) = 0 Then

        managerError.ThrowError _
            managerEmail.NoRecipient, _
            MODULE_NAME & ".OpenDraft", _
            "Belum ada penerima email (TO) di sheet SETTINGS."

        Exit Sub

    End If

    On Error GoTo FileFailed

    message = BuildMessage( _
        toList, _
        ccList, _
        bccList, _
        subject, _
        htmlBody)

    filePath = NewFilePath(subject)

    ' Create the physical .eml file as UTF-8 without BOM.
    WriteUtf8NoBom filePath, message

    ' Open the .eml through Windows Explorer.
    '
    ' Windows will use the default application registered for
    ' .eml files, such as New Outlook.
    OpenEmailFile filePath

    Exit Sub

FileFailed:
    managerError.ThrowError _
        managerEmail.EmailFileFailed, _
        MODULE_NAME & ".OpenDraft", _
        "File email tidak dapat dibuat atau dibuka: " & _
        filePath & " (" & Err.Description & ")"

End Sub

' =========================================================
' MESSAGE BUILDING
' =========================================================

' Creates the complete RFC 5322/MIME message.
'
' This follows the legacy flow:
'   - The complete .eml file is written as UTF-8.
'   - The HTML body is stored directly in the message.
'   - Base64 encoding is not required.
Private Function BuildMessage( _
    ByVal toList As String, _
    ByVal ccList As String, _
    ByVal bccList As String, _
    ByVal subject As String, _
    ByVal htmlBody As String) As String

    Dim result As String
    Dim completeHtml As String

    ValidateHeaderValue toList, "To"
    ValidateHeaderValue ccList, "Cc"
    ValidateHeaderValue bccList, "Bcc"
    ValidateHeaderValue subject, "Subject"

    completeHtml = BuildHtmlDocument(htmlBody)

    result = _
        "To: " & toList & vbCrLf

    If Len(Trim$(ccList)) > 0 Then

        result = result & _
            "Cc: " & ccList & vbCrLf

    End If

    If Len(Trim$(bccList)) > 0 Then

        result = result & _
            "Bcc: " & bccList & vbCrLf

    End If

    result = result & _
        "Subject: " & subject & vbCrLf & _
        "X-Unsent: 1" & vbCrLf & _
        "MIME-Version: 1.0" & vbCrLf & _
        "Content-Type: text/html; charset=""UTF-8""" & vbCrLf & _
        "Content-Transfer-Encoding: 8bit" & vbCrLf & _
        vbCrLf & _
        completeHtml

    BuildMessage = result

End Function

' Wraps the supplied body in a complete HTML document.
'
' If the calling module already supplies a complete HTML
' document, it is returned without adding another wrapper.
Private Function BuildHtmlDocument( _
    ByVal htmlBody As String) As String

    If InStr(1, htmlBody, "<html", vbTextCompare) > 0 Then

        BuildHtmlDocument = htmlBody
        Exit Function

    End If

    BuildHtmlDocument = _
        "<html>" & _
        "<head>" & _
        "<meta http-equiv=""Content-Type"" " & _
        "content=""text/html; charset=UTF-8"">" & _
        "<meta charset=""UTF-8"">" & _
        "</head>" & _
        "<body>" & _
        htmlBody & _
        "</body>" & _
        "</html>"

End Function

' Prevents CR/LF header injection through recipient or
' subject values.
Private Sub ValidateHeaderValue( _
    ByVal value As String, _
    ByVal fieldName As String)

    If InStr(1, value, vbCr, vbBinaryCompare) > 0 Or _
       InStr(1, value, vbLf, vbBinaryCompare) > 0 Then

        managerError.ThrowError _
            managerEmail.InvalidEmailHeader, _
            MODULE_NAME & ".ValidateHeaderValue", _
            "Field email """ & fieldName & _
            """ mengandung line break yang tidak valid."

    End If

End Sub

' =========================================================
' FILE PATH
' =========================================================

' Returns a file path in this format:
'
' %TEMP%\SIS Email\<subject> yyyymmdd-hhnnss.eml
Private Function NewFilePath( _
    ByVal subject As String) As String

    Dim folder As String
    Dim fileName As String

    folder = EmailFolder()
    fileName = SafeFileName(subject)

    If Len(fileName) = 0 Then
        fileName = "Email"
    End If

    NewFilePath = _
        folder & "\" & _
        fileName & " " & _
        Format$(Now, "yyyymmdd-hhnnss") & _
        ".eml"

End Function

' Returns the temporary email folder and creates it when
' necessary.
Private Function EmailFolder() As String

    Dim folder As String

    On Error GoTo FolderFailed

    folder = Environ$("TEMP") & "\" & EMAIL_FOLDER

    If Len(Dir$(folder, vbDirectory)) = 0 Then
        MkDir folder
    End If

    EmailFolder = folder

    Exit Function

FolderFailed:
    managerError.ThrowError _
        managerEmail.EmailFolderFailed, _
        MODULE_NAME & ".EmailFolder", _
        "Folder email tidak dapat dibuat: " & _
        folder & " (" & Err.Description & ")"

End Function

' Removes characters that Windows does not permit in a file
' name.
Private Function SafeFileName( _
    ByVal value As String) As String

    Dim result As String
    Dim badChars As Variant
    Dim character As Variant

    result = Trim$(value)

    badChars = Array( _
        "\", _
        "/", _
        ":", _
        "*", _
        "?", _
        """", _
        "<", _
        ">", _
        "|")

    For Each character In badChars

        result = Replace( _
            result, _
            CStr(character), _
            "-")

    Next character

    result = Replace(result, vbCr, " ")
    result = Replace(result, vbLf, " ")
    result = Replace(result, vbTab, " ")

    ' Remove repeated spaces.
    Do While InStr(1, result, "  ", vbBinaryCompare) > 0
        result = Replace(result, "  ", " ")
    Loop

    ' Windows file names cannot end with a period or space.
    Do While Len(result) > 0 And _
        (Right$(result, 1) = "." Or Right$(result, 1) = " ")

        result = Left$(result, Len(result) - 1)

    Loop

    If Len(result) > MAX_FILE_NAME_LENGTH Then
        result = Left$(result, MAX_FILE_NAME_LENGTH)
    End If

    SafeFileName = Trim$(result)

End Function

' =========================================================
' UTF-8 FILE WRITING
' =========================================================

' Writes the complete .eml message as UTF-8 without a BOM.
'
' ADODB.Stream normally adds the UTF-8 BOM:
'   EF BB BF
'
' The second binary stream copies the content starting after
' those first three bytes.
Private Sub WriteUtf8NoBom( _
    ByVal filePath As String, _
    ByVal content As String)

    Dim sourceStream As Object
    Dim destinationStream As Object

    On Error GoTo WriteFailed

    Set sourceStream = CreateObject("ADODB.Stream")

    With sourceStream

        .Type = AD_TYPE_TEXT
        .Charset = "utf-8"
        .Open
        .WriteText content

        ' Position must be zero before changing the stream
        ' from text mode to binary mode.
        .Position = 0
        .Type = AD_TYPE_BINARY

        ' Skip the three-byte UTF-8 BOM.
        .Position = 3

    End With

    Set destinationStream = CreateObject("ADODB.Stream")

    With destinationStream

        .Type = AD_TYPE_BINARY
        .Open

        sourceStream.CopyTo destinationStream

        .SaveToFile _
            filePath, _
            AD_SAVE_CREATE_OVERWRITE

        .Close

    End With

    sourceStream.Close

    Set destinationStream = Nothing
    Set sourceStream = Nothing

    Exit Sub

WriteFailed:
    On Error Resume Next

    If Not destinationStream Is Nothing Then
        destinationStream.Close
    End If

    If Not sourceStream Is Nothing Then
        sourceStream.Close
    End If

    Set destinationStream = Nothing
    Set sourceStream = Nothing

    On Error GoTo 0

    Err.Raise _
        managerEmail.EmailFileFailed, _
        MODULE_NAME & ".WriteUtf8NoBom", _
        "File .eml tidak dapat ditulis: " & _
        filePath

End Sub

' =========================================================
' OPEN EMAIL FILE
' =========================================================

' Opens the .eml file through Windows Explorer.
'
' Explorer uses the default Windows file association for
' .eml files. Therefore, this can open New Outlook without
' using Outlook COM automation.
Private Sub OpenEmailFile( _
    ByVal filePath As String)

    Dim expectedFolder As String
    Dim command As String

    On Error GoTo OpenFailed

    expectedFolder = EmailFolder()

    ' Only allow this procedure to open an .eml file.
    If StrComp( _
            Right$(filePath, 4), _
            ".eml", _
            vbTextCompare) <> 0 Then

        Err.Raise _
            vbObjectError + EmailOpenFailed, _
            MODULE_NAME & ".OpenEmailFile", _
            "Hanya file .eml yang dapat dibuka."

    End If

    ' Only allow a file inside the expected temporary email
    ' folder.
    If StrComp( _
            Left$(filePath, Len(expectedFolder) + 1), _
            expectedFolder & "\", _
            vbTextCompare) <> 0 Then

        Err.Raise _
            vbObjectError + EmailOpenFailed, _
            MODULE_NAME & ".OpenEmailFile", _
            "Lokasi file email tidak valid."

    End If

    command = _
        "explorer.exe " & _
        Chr$(34) & filePath & Chr$(34)

    Shell command, vbNormalFocus

    Exit Sub

OpenFailed:
    Err.Raise _
        managerEmail.EmailOpenFailed, _
        MODULE_NAME & ".OpenEmailFile", _
        "File .eml tidak dapat dibuka: " & _
        filePath & " (" & Err.Description & ")"

End Sub

' =========================================================
' RECIPIENTS
' =========================================================

' Returns email addresses from the EMAIL column in this
' format:
'
' a@x.com; b@y.com
Public Function Recipients( _
    ByVal tableName As String) As String

    Dim tbl As ListObject
    Dim rowIndex As Long
    Dim address As String
    Dim result As String

    Set tbl = managerTable.GetTable( _
        managerSheet.GetSheet( _
            managerSheet.SHEET_SETTINGS), _
        tableName)

    For rowIndex = 1 To managerTable.rowCount(tbl)

        address = Trim$( _
            managerTable.GetText( _
                tbl, _
                rowIndex, _
                "EMAIL"))

        If Len(address) > 0 Then

            ValidateHeaderValue _
                address, _
                tableName & ".EMAIL"

            If Len(result) > 0 Then
                result = result & "; "
            End If

            result = result & address

        End If

    Next rowIndex

    Recipients = result

End Function

' Returns the first NAMA value from the TO table for use in
' the email greeting.
Public Function FirstRecipientName() As String

    Dim tbl As ListObject

    Set tbl = managerTable.GetTable( _
        managerSheet.GetSheet( _
            managerSheet.SHEET_SETTINGS), _
        TABLE_EMAIL_TO)

    If managerTable.rowCount(tbl) > 0 Then

        FirstRecipientName = managerTable.GetText( _
            tbl, _
            1, _
            "NAMA")

    End If

End Function

' =========================================================
' HTML ENCODING
' =========================================================

' Converts a value into HTML-safe text.
Public Function HtmlText( _
    ByVal value As Variant) As String

    Dim text As String

    text = coreHelper.TextOrEmpty(value)

    ' Ampersand must be replaced first.
    text = Replace(text, "&", "&amp;")
    text = Replace(text, "<", "&lt;")
    text = Replace(text, ">", "&gt;")
    text = Replace(text, """", "&quot;")
    text = Replace(text, "'", "&#39;")

    ' Preserve line breaks from worksheet cells.
    text = Replace(text, vbCrLf, "<br>")
    text = Replace(text, vbCr, "<br>")
    text = Replace(text, vbLf, "<br>")

    HtmlText = text

End Function