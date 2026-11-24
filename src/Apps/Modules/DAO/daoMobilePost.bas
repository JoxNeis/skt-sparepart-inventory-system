Attribute VB_Name = "daoMobilePost"
Option Explicit

' =========================================================
' DAO - MOBILE POST
'
' One function = one saved Access query.
' Validation, transaction, row-count checks and activity
' log live in serviceMobilePost.
' =========================================================

Private Const MODULE_NAME As String = "DAO_MOBILE_POST"


' select_mobile_post_by_manufacturing_plant_name
' Returns MobilePost objects (no collection key, so a
' duplicate name does not cause error 457). Plant.Name is
' the given manufacturingPlantName.
' Name is only set when the query returns a value, so
' serviceMobilePost can report a blank name itself.
Public Function SelectMobilePostByManufacturingPlantName( _
    ByVal conn As ADODB.connection, _
    ByVal manufacturingPlantName As String) As Collection

    Dim cmd As ADODB.Command
    Dim rs As ADODB.recordset
    Dim mobilePosts As Collection
    Dim mobilePostItem As MobilePost
    Dim plantItem As Plant
    Dim mobilePostName As String

    Set cmd = managerDatabase.CreateCommand(conn, _
        "select_mobile_post_by_manufacturing_plant_name")

    managerDatabase.AddTextParameter cmd, "p_manufacturing_plant_name", manufacturingPlantName

    Set rs = managerDatabase.ExecuteRecordset(cmd)

    Set mobilePosts = New Collection

    Do While Not rs.EOF

        mobilePostName = managerDatabase.FieldText(rs, "name")

        Set plantItem = New Plant
        plantItem.Name = manufacturingPlantName

        Set mobilePostItem = New MobilePost

        With mobilePostItem
            If Len(mobilePostName) > 0 Then .Name = mobilePostName

            Set .Plant = plantItem
        End With

        mobilePosts.Add item:=mobilePostItem

        Set mobilePostItem = Nothing
        Set plantItem = Nothing

        rs.MoveNext
    Loop

    managerDatabase.CloseRecordset rs

    Set SelectMobilePostByManufacturingPlantName = mobilePosts

End Function


' create_mobile_post
' Parameter order: p_manufacturing_plant_name,
' p_mobile_post_name, p_mobile_post_code.
' Returns the affected row count.
Public Function CreateMobilePost( _
    ByVal conn As ADODB.connection, _
    ByVal mobilePostItem As MobilePost) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "create_mobile_post")

    managerDatabase.AddTextParameter cmd, "p_manufacturing_plant_name", mobilePostItem.Plant.Name
    managerDatabase.AddTextParameter cmd, "p_mobile_post_name", mobilePostItem.Name
    managerDatabase.AddTextParameter cmd, "p_mobile_post_code", mobilePostItem.Code

    CreateMobilePost = managerDatabase.ExecuteNonQuery(cmd)

End Function
