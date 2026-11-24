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

    Set cmd = managerDatabase.CreateCommand(conn, "select_vendors")

    Set rs = managerDatabase.ExecuteRecordset(cmd)

    ' Fields required by the Vendor mapping below.
    managerDatabase.ValidateFields rs, Array("name", "code", "company_name")

    Set vendors = New Collection

    Do While Not rs.EOF

        vendorName = managerDatabase.FieldText(rs, "name")
        vendorCode = managerDatabase.FieldText(rs, "code")
        vendorCompanyName = managerDatabase.FieldText(rs, "company_name")

        Set vendorItem = New Vendor

        With vendorItem
            .Id = managerDatabase.FieldLong(rs, "id")
            If Len(vendorName) > 0 Then .Name = vendorName
            If Len(vendorCode) > 0 Then .Code = vendorCode
            If Len(vendorCompanyName) > 0 Then .CompanyName = vendorCompanyName
        End With

        vendors.Add item:=vendorItem

        Set vendorItem = Nothing

        rs.MoveNext
    Loop

    managerDatabase.CloseRecordset rs

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

    Set cmd = managerDatabase.CreateCommand(conn, "create_vendor")

    managerDatabase.AddTextParameter cmd, "p_vendor_name", vendorItem.Name
    managerDatabase.AddTextParameter cmd, "p_vendor_company_name", vendorItem.CompanyName
    managerDatabase.AddTextParameter cmd, "p_vendor_code", vendorItem.Code

    CreateVendor = managerDatabase.ExecuteNonQuery(cmd)

End Function
