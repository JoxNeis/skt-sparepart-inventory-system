Attribute VB_Name = "daoSparepart"
Option Explicit

' =========================================================
' DAO - SPAREPART (catalogue)
' Table: spareparts
'
' One function = one saved Access query.
' Validation, transaction, row-count checks and activity
' log live in serviceSparepart.
' =========================================================

Private Const MODULE_NAME As String = "daoSparepart"


' sparepart_list
' Active spareparts, ordered by name.
Public Function ListSpareparts( _
    ByVal conn As ADODB.connection) As Collection

    Dim cmd As ADODB.Command
    Dim rs As ADODB.recordset
    Dim spareparts As Collection

    Set cmd = managerDatabase.CreateCommand(conn, "sparepart_list")
    Set rs = managerDatabase.ExecuteRecordset(cmd)

    Set spareparts = New Collection

    Do While Not rs.EOF
        spareparts.Add MapSparepart(rs)
        rs.MoveNext
    Loop

    managerDatabase.CloseRecordset rs

    Set ListSpareparts = spareparts

End Function


' sparepart_find_by_name
' Active spareparts whose name, code or alias equals the
' text (matched_on: CATALOGUE or ALIAS is not mapped).
Public Function FindSparepartsByName( _
    ByVal conn As ADODB.connection, _
    ByVal nameText As String) As Collection

    Dim cmd As ADODB.Command
    Dim rs As ADODB.recordset
    Dim spareparts As Collection
    Dim sparepartItem As Sparepart

    Set cmd = managerDatabase.CreateCommand(conn, "sparepart_find_by_name")

    managerDatabase.AddTextParameter cmd, "p_name", nameText, 255

    Set rs = managerDatabase.ExecuteRecordset(cmd)

    Set spareparts = New Collection

    Do While Not rs.EOF

        Set sparepartItem = New Sparepart

        With sparepartItem
            .Id = managerDatabase.FieldLong(rs, "id")
            .Code = managerDatabase.FieldText(rs, "code")
            .Name = managerDatabase.FieldText(rs, "name")
            .Alias = managerDatabase.FieldText(rs, "alias")
            .UnitOfMeasurement = managerDatabase.FieldText(rs, "unit_of_measurement")
        End With

        spareparts.Add sparepartItem
        rs.MoveNext

    Loop

    managerDatabase.CloseRecordset rs

    Set FindSparepartsByName = spareparts

End Function


' sparepart_create
Public Function CreateSparepart( _
    ByVal conn As ADODB.connection, _
    ByVal sparepartItem As Sparepart) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "sparepart_create")

    AddSparepartParameters cmd, sparepartItem

    CreateSparepart = managerDatabase.ExecuteNonQuery(cmd)

End Function


' sparepart_update
Public Function UpdateSparepart( _
    ByVal conn As ADODB.connection, _
    ByVal sparepartItem As Sparepart) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "sparepart_update")

    managerDatabase.AddLongParameter cmd, "p_sparepart_id", sparepartItem.Id
    AddSparepartParameters cmd, sparepartItem

    UpdateSparepart = managerDatabase.ExecuteNonQuery(cmd)

End Function


' sparepart_archive
Public Function ArchiveSparepart( _
    ByVal conn As ADODB.connection, _
    ByVal sparepartId As Long) As Long

    Dim cmd As ADODB.Command

    Set cmd = managerDatabase.CreateCommand(conn, "sparepart_archive")

    managerDatabase.AddLongParameter cmd, "p_sparepart_id", sparepartId

    ArchiveSparepart = managerDatabase.ExecuteNonQuery(cmd)

End Function


' =========================================================
' INTERNAL
' =========================================================

' p_code ... p_status, in the order of the PARAMETERS
' declaration of sparepart_create / sparepart_update.
Private Sub AddSparepartParameters( _
    ByVal cmd As ADODB.Command, _
    ByVal sparepartItem As Sparepart)

    managerDatabase.AddTextParameter cmd, "p_code", sparepartItem.Code, 40
    managerDatabase.AddTextParameter cmd, "p_name", sparepartItem.Name, 255
    managerDatabase.AddNullableTextParameter cmd, "p_category", sparepartItem.Category, 50
    managerDatabase.AddNullableTextParameter cmd, "p_applicable", sparepartItem.Applicable, 255
    managerDatabase.AddTextParameter cmd, "p_unit_of_measurement", sparepartItem.UnitOfMeasurement, 15
    managerDatabase.AddTextParameter cmd, "p_alias", sparepartItem.Alias, 255
    managerDatabase.AddNullableTextParameter cmd, "p_status", sparepartItem.Status, 50

End Sub


Private Function MapSparepart( _
    ByVal rs As ADODB.recordset) As Sparepart

    Dim sparepartItem As Sparepart

    Set sparepartItem = New Sparepart

    With sparepartItem
        .Id = managerDatabase.FieldLong(rs, "id")
        .Code = managerDatabase.FieldText(rs, "code")
        .Name = managerDatabase.FieldText(rs, "name")
        .Alias = managerDatabase.FieldText(rs, "alias")
        .Category = managerDatabase.FieldText(rs, "category")
        .Applicable = managerDatabase.FieldText(rs, "applicable")
        .UnitOfMeasurement = managerDatabase.FieldText(rs, "unit_of_measurement")
        .Status = managerDatabase.FieldText(rs, "status")
    End With

    Set MapSparepart = sparepartItem

End Function
