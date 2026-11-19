Attribute VB_Name = "serviceMobilePost"
Option Explicit

' =========================================================
' SERVICE - MOBILE POST
'
' Public entry points (called by sheets) get the
' connection and own the transaction:
'   SelectMobilePostByManufacturingPlantName, CreateMobilePost
' =========================================================

Private Const MODULE_NAME As String = "SERVICE_MOBILE_POST"

' =========================================================
' ERROR CODES (17300 - 17399)
' Usage: ErrorManager.ThrowError serviceMobilePost.EmptyMobilePostObject, ...
' =========================================================
Public Const EmptyManufacturingPlantName As Long = 17301
Public Const EmptyRecordsetMobilePostName As Long = 17302         ' query returned a blank mobile post name
Public Const EmptyMobilePostObject As Long = 17303                ' MobilePost object is Nothing
Public Const EmptyMobilePostPlant As Long = 17304                 ' MobilePost.Plant is Nothing
Public Const EmptyMobilePostName As Long = 17305
Public Const EmptyMobilePostCode As Long = 17306
Public Const EmptyCurrentUser As Long = 17307                     ' creating user cannot be determined
Public Const NoRowAffectedCreateMobilePost As Long = 17308        ' 0 rows inserted
Public Const MultipleRowsAffectedCreateMobilePost As Long = 17309 ' > 1 row inserted


' =========================================================
' PUBLIC ENTRY POINTS
' =========================================================

' Mobile posts of one manufacturing plant. Read only, no
' transaction.
Public Function SelectMobilePostByManufacturingPlantName( _
    ByVal manufacturingPlantName As String) As Collection

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim mobilePosts As Collection
    Dim mobilePostItem As MobilePost
    Dim errInfo As ErrorInfo

    manufacturingPlantName = Trim$(manufacturingPlantName)

    ' =====================================================
    ' Step 1: Parameter validation
    ' =====================================================
    If ValidationText.IsBlank(manufacturingPlantName) Then
        ErrorManager.ThrowError _
            serviceMobilePost.EmptyManufacturingPlantName, _
            MODULE_NAME & ".SelectMobilePostByManufacturingPlantName", _
            "Nama manufacturing plant " & _
            "tidak boleh kosong."
    End If

    ' =====================================================
    ' Step 2: Query
    ' =====================================================
    Set conn = DatabaseManager.GetConnection()

    Set mobilePosts = daoMobilePost.SelectMobilePostByManufacturingPlantName( _
        conn, manufacturingPlantName)

    ' =====================================================
    ' Step 3: Every returned mobile post must have a name
    ' =====================================================
    For Each mobilePostItem In mobilePosts

        If ValidationText.IsBlank(mobilePostItem.Name) Then
            ErrorManager.ThrowError _
                serviceMobilePost.EmptyRecordsetMobilePostName, _
                MODULE_NAME & ".SelectMobilePostByManufacturingPlantName", _
                "Query mengembalikan nama mobile post " & _
                "yang kosong."
        End If

    Next mobilePostItem

    Set SelectMobilePostByManufacturingPlantName = mobilePosts

    Exit Function

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".SelectMobilePostByManufacturingPlantName", _
        "Gagal mengambil daftar mobile post."

End Function


' Insert one mobile post and write the activity log, in one
' transaction.
Public Sub CreateMobilePost( _
    ByVal mobilePostItem As MobilePost)

    On Error GoTo ErrorHandler

    Dim conn As ADODB.connection
    Dim affected As Long
    Dim inTransaction As Boolean

    Dim manufacturingPlantName As String
    Dim mobilePostName As String
    Dim mobilePostCode As String

    Dim currentUser As String
    Dim activityDetail As String

    Dim errInfo As ErrorInfo

    inTransaction = False

    ' =====================================================
    ' Step 1: Validate MobilePost object
    ' =====================================================
    If mobilePostItem Is Nothing Then
        ErrorManager.ThrowError _
            serviceMobilePost.EmptyMobilePostObject, _
            MODULE_NAME & ".CreateMobilePost", _
            "Objek MobilePost tidak boleh kosong."
    End If

    If mobilePostItem.Plant Is Nothing Then
        ErrorManager.ThrowError _
            serviceMobilePost.EmptyMobilePostPlant, _
            MODULE_NAME & ".CreateMobilePost", _
            "Manufacturing plant pada MobilePost " & _
            "tidak boleh kosong."
    End If

    ' =====================================================
    ' Step 2: Read and normalize properties
    ' =====================================================
    manufacturingPlantName = Trim$(mobilePostItem.Plant.Name)
    mobilePostName = Trim$(mobilePostItem.Name)
    mobilePostCode = Trim$(mobilePostItem.Code)

    ' =====================================================
    ' Step 3: Validate property values
    ' =====================================================
    If ValidationText.IsBlank(manufacturingPlantName) Then
        ErrorManager.ThrowError _
            serviceMobilePost.EmptyManufacturingPlantName, _
            MODULE_NAME & ".CreateMobilePost", _
            "Nama manufacturing plant " & _
            "tidak boleh kosong."
    End If

    If ValidationText.IsBlank(mobilePostName) Then
        ErrorManager.ThrowError _
            serviceMobilePost.EmptyMobilePostName, _
            MODULE_NAME & ".CreateMobilePost", _
            "Nama mobile post tidak boleh kosong."
    End If

    If ValidationText.IsBlank(mobilePostCode) Then
        ErrorManager.ThrowError _
            serviceMobilePost.EmptyMobilePostCode, _
            MODULE_NAME & ".CreateMobilePost", _
            "Kode mobile post tidak boleh kosong."
    End If

    ' =====================================================
    ' Step 4: Determine current user
    ' =====================================================
    currentUser = Trim$(pConfig.currentUser())

    If ValidationText.IsBlank(currentUser) Then
        ErrorManager.ThrowError _
            serviceMobilePost.EmptyCurrentUser, _
            MODULE_NAME & ".CreateMobilePost", _
            "User yang membuat mobile post " & _
            "tidak dapat ditentukan."
    End If

    ' =====================================================
    ' Step 5: Begin transaction
    ' =====================================================
    Set conn = DatabaseManager.GetConnection()

    DatabaseManager.BeginTransaction conn
    inTransaction = True

    ' =====================================================
    ' Step 6: Insert mobile post
    ' =====================================================
    affected = daoMobilePost.CreateMobilePost(conn, mobilePostItem)

    ' =====================================================
    ' Step 7: Validate affected records
    ' =====================================================
    If ValidationNumeric.IsLessThan(affected, 1) Then
        ErrorManager.ThrowError _
            serviceMobilePost.NoRowAffectedCreateMobilePost, _
            MODULE_NAME & ".CreateMobilePost", _
            "Mobile post tidak berhasil dibuat." & _
            vbCrLf & _
            "Manufacturing plant mungkin tidak " & _
            "ditemukan, atau nama/kode mobile post " & _
            "sudah digunakan."
    End If

    If ValidationNumeric.IsGreaterThan(affected, 1) Then
        ErrorManager.ThrowError _
            serviceMobilePost.MultipleRowsAffectedCreateMobilePost, _
            MODULE_NAME & ".CreateMobilePost", _
            "Pembuatan mobile post memengaruhi " & _
            "lebih dari satu record. " & _
            "Transaksi dibatalkan."
    End If

    ' =====================================================
    ' Step 8: Build activity detail
    ' =====================================================
    activityDetail = BuildMobilePostCreateDetail(mobilePostItem)

    ' =====================================================
    ' Step 9: Write activity log
    ' =====================================================
    serviceActivityLog.WriteActivityLog _
        conn, _
        "MobilePost.Create", _
        activityDetail

    ' =====================================================
    ' Step 10: Commit
    ' =====================================================
    DatabaseManager.CommitTransaction conn
    inTransaction = False

    Exit Sub

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    If inTransaction Then
        DatabaseManager.RollbackTransaction conn
    End If

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".CreateMobilePost", _
        "Gagal membuat data mobile post."

End Sub


' =========================================================
' INTERNAL
' =========================================================

Private Function BuildMobilePostCreateDetail( _
    ByVal mobilePostItem As MobilePost) As String

    On Error GoTo ErrorHandler

    Dim detail As String
    Dim errInfo As ErrorInfo

    If mobilePostItem Is Nothing Then
        ErrorManager.ThrowError _
            serviceMobilePost.EmptyMobilePostObject, _
            MODULE_NAME & ".BuildMobilePostCreateDetail", _
            "Objek MobilePost tidak boleh kosong."
    End If

    If mobilePostItem.Plant Is Nothing Then
        ErrorManager.ThrowError _
            serviceMobilePost.EmptyMobilePostPlant, _
            MODULE_NAME & ".BuildMobilePostCreateDetail", _
            "Manufacturing plant pada MobilePost " & _
            "tidak boleh kosong."
    End If

    detail = "mobile_post_name=" & Trim$(mobilePostItem.Name)
    detail = detail & "; mobile_post_code=" & Trim$(mobilePostItem.Code)
    detail = detail & "; manufacturing_plant_name=" & Trim$(mobilePostItem.Plant.Name)

    BuildMobilePostCreateDetail = detail

    Exit Function

ErrorHandler:
    errInfo = ErrorManager.CaptureError()

    ErrorManager.Rethrow errInfo, _
        MODULE_NAME & ".BuildMobilePostCreateDetail"

End Function
