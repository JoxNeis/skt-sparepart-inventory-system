Attribute VB_Name = "daoVendor"
Option Explicit

' =========================================================
' DAO - VENDOR
'
' One function = one saved Access query.
' Validation, transaction, row-count checks and activity
' log live in serviceVendor.
' =========================================================

Private Const MODULE_NAME As String = "DAO_VENDOR"


' select_vendors
' Returns Vendor objects (no collection key, so duplicate
' names do not cause error 457).
' Name, Code and CompanyName are only set when the query
' returns a value, so serviceVendor can report a blank
' value with the record number itself.
Public Function SelectVendors( _
    ByVal conn As ADODB.connection) As Collection

    Dim cmd As ADODB.Command
    Dim rs As ADODB.recordset
    Dim vendors As Collection
    Dim vendorItem As Vendor
    Dim vendorName As String
    Dim vendorCode As String
    Dim vendorCompanyName As String

    Set cmd = DatabaseManager.CreateCommand(conn, "select_vendors")

    Set rs = DatabaseManager.ExecuteRecordset(cmd)

    ' Fields required by the Vendor mapping below.
    DatabaseManager.ValidateFields rs, Array("name", "code", "company_name")

    Set vendors = New Collection

    Do While Not rs.EOF

        vendorName = DatabaseManager.FieldText(rs, "name")
        vendorCode = DatabaseManager.FieldText(rs, "code")
        vendorCompanyName = DatabaseManager.FieldText(rs, "company_name")

        Set vendorItem = New Vendor

        With vendorItem
            .Id = DatabaseManager.FieldLong(rs, "id")
            If Len(vendorName) > 0 Then .Name = vendorName
            If Len(vendorCode) > 0 Then .Code = vendorCode
            If Len(vendorCompanyName) > 0 Then .CompanyName = vendorCompanyName
        End With

        vendors.Add item:=vendorItem

        Set vendorItem = Nothing

        rs.MoveNext
    Loop

    DatabaseManager.CloseRecordset rs

    Set SelectVendors = vendors

End Function


' create_vendor
' Parameter order: p_vendor_name, p_vendor_company_name,
' p_vendor_code.
' Returns the affected row count.
Public Function CreateVendor( _
    ByVal conn As ADODB.connection, _
    ByVal vendorItem As Vendor) As Long

    Dim cmd As ADODB.Command

    Set cmd = DatabaseManager.CreateCommand(conn, "create_vendor")

    DatabaseManager.AddTextParameter cmd, "p_vendor_name", vendorItem.Name
    DatabaseManager.AddTextParameter cmd, "p_vendor_company_name", vendorItem.CompanyName
    DatabaseManager.AddTextParameter cmd, "p_vendor_code", vendorItem.Code

    CreateVendor = DatabaseManager.ExecuteNonQuery(cmd)

End Function
