Attribute VB_Name = "daoVendor"
Option Explicit

' =========================================================
' DAO - VENDOR
' Table: vendors
'
' One function = one saved Access query.
' Validation, transaction, row-count checks and activity
' log live in serviceVendor.
' =========================================================

Private Const MODULE_NAME As String = "daoVendor"


' vendor_list
' Active vendors, ordered by name.
Public Function ListVendors( _
    ByVal conn As ADODB.connection) As Collection

    Dim cmd As ADODB.Command
    Dim rs As ADODB.recordset
    Dim vendors As Collection

    Set cmd = managerDatabase.CreateCommand(conn, "vendor_list")
    Set rs = managerDatabase.ExecuteRecordset(cmd)

    Set vendors = New Collection

    Do While Not rs.EOF
        vendors.Add MapVendor(rs)
        rs.MoveNext
    Loop

    managerDatabase.CloseRecordset rs

    Set ListVendors = vendors

End Function


' vendor_create
Public Function CreateVendor( _
    ByVal conn As ADODB.connection, _
    ByVal vendorItem As Vendor) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "vendor_create")

    managerDatabase.AddTextParameter cmd, "p_code", vendorItem.Code, 40
    managerDatabase.AddTextParameter cmd, "p_name", vendorItem.Name, 255
    managerDatabase.AddNullableTextParameter cmd, "p_company_name", vendorItem.CompanyName, 255
    managerDatabase.AddNullableTextParameter cmd, "p_email", vendorItem.Email, 255

    CreateVendor = managerDatabase.ExecuteNonQuery(cmd)

End Function


' vendor_update
Public Function UpdateVendor( _
    ByVal conn As ADODB.connection, _
    ByVal vendorItem As Vendor) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "vendor_update")

    managerDatabase.AddLongParameter cmd, "p_vendor_id", vendorItem.Id
    managerDatabase.AddTextParameter cmd, "p_code", vendorItem.Code, 40
    managerDatabase.AddTextParameter cmd, "p_name", vendorItem.Name, 255
    managerDatabase.AddNullableTextParameter cmd, "p_company_name", vendorItem.CompanyName, 255
    managerDatabase.AddNullableTextParameter cmd, "p_email", vendorItem.Email, 255

    UpdateVendor = managerDatabase.ExecuteNonQuery(cmd)

End Function


' vendor_archive
Public Function ArchiveVendor( _
    ByVal conn As ADODB.connection, _
    ByVal vendorId As Long) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "vendor_archive")

    managerDatabase.AddLongParameter cmd, "p_vendor_id", vendorId

    ArchiveVendor = managerDatabase.ExecuteNonQuery(cmd)

End Function


' =========================================================
' INTERNAL
' =========================================================

Private Function MapVendor( _
    ByVal rs As ADODB.recordset) As Vendor

    Dim vendorItem As Vendor

    Set vendorItem = New Vendor

    With vendorItem
        .Id = managerDatabase.FieldLong(rs, "id")
        .Code = managerDatabase.FieldText(rs, "code")
        .Name = managerDatabase.FieldText(rs, "name")
        .CompanyName = managerDatabase.FieldText(rs, "company_name")
        .Email = managerDatabase.FieldText(rs, "email")
    End With

    Set MapVendor = vendorItem

End Function
