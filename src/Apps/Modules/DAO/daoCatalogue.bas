Attribute VB_Name = "daoCatalogue"
Option Explicit

' =========================================================
' DAO - CATALOGUE
'
' One function = one saved Access query.
' Validation, transaction, row-count checks and activity
' log live in serviceCatalogue.
' =========================================================

Private Const MODULE_NAME As String = "DAO_CATALOGUE"


' select_catalogues
' Returns Catalogue objects (no collection key, so duplicate
' names, codes or aliases do not cause error 457).
Public Function GetAllCatalogues( _
    ByVal conn As ADODB.connection) As Collection

    Dim cmd As ADODB.Command
    Dim rs As ADODB.recordset
    Dim catalogues As Collection
    Dim catalogueItem As Catalogue

    Set cmd = managerDatabase.CreateCommand(conn, "select_catalogues")

    Set rs = managerDatabase.ExecuteRecordset(cmd)

    Set catalogues = New Collection

    Do While Not rs.EOF

        Set catalogueItem = New Catalogue

        With catalogueItem
            .Id = managerDatabase.FieldLong(rs, "id")
            .Name = managerDatabase.FieldText(rs, "name")
            .Alias = managerDatabase.FieldText(rs, "alias")
            .Code = managerDatabase.FieldText(rs, "code")
            .Category = managerDatabase.FieldText(rs, "category")
            .unitOfMeasurement = managerDatabase.FieldText(rs, "unit_of_measurement")
        End With

        catalogues.Add item:=catalogueItem

        Set catalogueItem = Nothing

        rs.MoveNext
    Loop

    managerDatabase.CloseRecordset rs

    Set GetAllCatalogues = catalogues

End Function


' create_catalogue
' Returns the affected row count.
Public Function CreateCatalogue( _
    ByVal conn As ADODB.connection, _
    ByVal catalogueItem As Catalogue) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "create_catalogue")

    managerDatabase.AddTextParameter cmd, "p_name", catalogueItem.Name
    managerDatabase.AddTextParameter cmd, "p_alias", catalogueItem.Alias
    managerDatabase.AddTextParameter cmd, "p_code", catalogueItem.Code
    managerDatabase.AddTextParameter cmd, "p_category", catalogueItem.Category
    managerDatabase.AddTextParameter cmd, "p_unit_of_measurement", catalogueItem.unitOfMeasurement

    CreateCatalogue = managerDatabase.ExecuteNonQuery(cmd)

End Function
