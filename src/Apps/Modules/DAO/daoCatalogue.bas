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

    Set cmd = DatabaseManager.CreateCommand(conn, "select_catalogues")

    Set rs = DatabaseManager.ExecuteRecordset(cmd)

    Set catalogues = New Collection

    Do While Not rs.EOF

        Set catalogueItem = New Catalogue

        With catalogueItem
            .Id = DatabaseManager.FieldLong(rs, "id")
            .Name = DatabaseManager.FieldText(rs, "name")
            .Alias = DatabaseManager.FieldText(rs, "alias")
            .Code = DatabaseManager.FieldText(rs, "code")
            .Category = DatabaseManager.FieldText(rs, "category")
            .unitOfMeasurement = DatabaseManager.FieldText(rs, "unit_of_measurement")
        End With

        catalogues.Add item:=catalogueItem

        Set catalogueItem = Nothing

        rs.MoveNext
    Loop

    DatabaseManager.CloseRecordset rs

    Set GetAllCatalogues = catalogues

End Function


' create_catalogue
' Returns the affected row count.
Public Function CreateCatalogue( _
    ByVal conn As ADODB.connection, _
    ByVal catalogueItem As Catalogue) As Long

    Dim cmd As ADODB.Command

    Set cmd = DatabaseManager.CreateCommand(conn, "create_catalogue")

    DatabaseManager.AddTextParameter cmd, "p_name", catalogueItem.Name
    DatabaseManager.AddTextParameter cmd, "p_alias", catalogueItem.Alias
    DatabaseManager.AddTextParameter cmd, "p_code", catalogueItem.Code
    DatabaseManager.AddTextParameter cmd, "p_category", catalogueItem.Category
    DatabaseManager.AddTextParameter cmd, "p_unit_of_measurement", catalogueItem.unitOfMeasurement

    CreateCatalogue = DatabaseManager.ExecuteNonQuery(cmd)

End Function
