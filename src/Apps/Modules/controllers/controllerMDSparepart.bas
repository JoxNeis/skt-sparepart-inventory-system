Attribute VB_Name = "controllerMDSparepart"
Option Explicit

' =========================================================
' CONTROLLER - MASTER DATA SPAREPART (sheet MD_SPAREPART)
'
'   Form: NAMA SPAREPART K6, KODE K7, APPLICABLE K8,
'         SATUAN K9, ALIAS K10 (blank = same as the name)
'   BATAL I12:J13, SIMPAN K12:L13, REFRESH I15:L16
'
'   TableMDSparepart (ID, NAMA SPAREPART, KODE, DISPLAY,
'   SATUAN, ALIAS) is also the source of every sparepart
'   dropdown (ListSparepart = TableMDSparepart[DISPLAY]).
' =========================================================

Private Const MODULE_NAME As String = "controllerMDSparepart"

' =========================================================
' ERROR CODES (14600 - 14699)
' =========================================================
' (validation errors come from managerSheet / the Sparepart model)

Private Const HEADER_CELL As String = "I5"
Private Const NAME_CELL As String = "K6"
Private Const CODE_CELL As String = "K7"
Private Const APPLICABLE_CELL As String = "K8"
Private Const UNIT_CELL As String = "K9"
Private Const ALIAS_CELL As String = "K10"

Private Const BUTTON_CANCEL As String = "I12:J13"
Private Const BUTTON_SAVE As String = "K12:L13"
Private Const BUTTON_REFRESH As String = "I15:L16"

Private Const TABLE_NAME As String = "TableMDSparepart"

Private Const HEADER_NEW As String = "FORM PENGISIAN SPAREPART BARU"


' =========================================================
' EVENTS
' =========================================================

Public Sub OnActivate( _
    ByVal ws As Worksheet)

    ws.Range(HEADER_CELL).value = HEADER_NEW

End Sub


Public Sub OnDoubleClick( _
    ByVal ws As Worksheet, _
    ByVal target As Range, _
    ByRef cancel As Boolean)

    If managerSheet.IsHit(target, ws, BUTTON_SAVE) Then
        cancel = True
        Save ws
    ElseIf managerSheet.IsHit(target, ws, BUTTON_CANCEL) Then
        cancel = True
        ClearForm ws
    ElseIf managerSheet.IsHit(target, ws, BUTTON_REFRESH) Then
        cancel = True
        Refresh ws, True
    End If

End Sub


' =========================================================
' ACTIONS
' =========================================================

Public Sub Save( _
    ByVal ws As Worksheet)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo
    Dim sparepartItem As Sparepart
    Dim aliasText As String

    Set sparepartItem = New Sparepart

    With sparepartItem
        .Name = managerSheet.ReadText(ws, NAME_CELL, "NAMA SPAREPART")
        .Code = managerSheet.ReadText(ws, CODE_CELL, "KODE")
        .Applicable = managerSheet.ReadText(ws, APPLICABLE_CELL, "APPLICABLE", False)
        .UnitOfMeasurement = managerSheet.ReadText(ws, UNIT_CELL, "SATUAN")

        aliasText = managerSheet.ReadText(ws, ALIAS_CELL, "ALIAS", False)
        If Len(aliasText) = 0 Then aliasText = .Name
        .Alias = aliasText
    End With

    state = managerForm.BeginBusy("Menyimpan sparepart...")

    serviceSparepart.CreateSparepart sparepartItem
    RefreshTable ws
    ClearFormCells ws

    managerForm.EndBusy state

    managerForm.ShowInfo "Sparepart " & sparepartItem.DisplayName & " tersimpan."

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


Public Sub ClearForm( _
    ByVal ws As Worksheet)

    ClearFormCells ws

End Sub


' REFRESH: reload the table from the database.
Public Sub Refresh( _
    ByVal ws As Worksheet, _
    Optional ByVal showResult As Boolean = False)

    On Error GoTo ErrorHandler

    Dim state As FormState
    Dim errInfo As ErrorInfo
    Dim count As Long

    state = managerForm.BeginBusy("Memuat data sparepart...")
    count = RefreshTable(ws)
    managerForm.EndBusy state

    If showResult Then Application.StatusBar = count & " sparepart dimuat."

    Exit Sub

ErrorHandler:
    errInfo = managerError.CaptureError()
    managerForm.EndBusy state
    managerError.ShowError errInfo

End Sub


' Fill the table from the database (no error handling: the
' caller decides). Returns the number of spareparts.
Public Function RefreshTable( _
    ByVal ws As Worksheet) As Long

    Dim spareparts As Collection
    Dim sparepartItem As Sparepart
    Dim data() As Variant
    Dim r As Long
    Dim tbl As ListObject

    Set tbl = managerTable.GetTable(ws, TABLE_NAME)
    Set spareparts = serviceSparepart.ListSpareparts()

    If spareparts.Count = 0 Then
        managerTable.SetRowCount tbl, 0
        Exit Function
    End If

    ReDim data(1 To spareparts.Count, 1 To 6)

    For Each sparepartItem In spareparts
        r = r + 1
        data(r, 1) = sparepartItem.Id
        data(r, 2) = sparepartItem.Name
        data(r, 3) = sparepartItem.Code
        data(r, 4) = sparepartItem.DisplayName
        data(r, 5) = sparepartItem.UnitOfMeasurement
        data(r, 6) = sparepartItem.Alias
    Next sparepartItem

    managerTable.WriteData tbl, data, spareparts.Count, _
        Array("0", managerSheet.TEXT_FORMAT, managerSheet.TEXT_FORMAT, _
              managerSheet.TEXT_FORMAT, managerSheet.TEXT_FORMAT, managerSheet.TEXT_FORMAT)

    RefreshTable = spareparts.Count

End Function


' =========================================================
' INTERNAL
' =========================================================

Private Sub ClearFormCells( _
    ByVal ws As Worksheet)

    managerForm.ClearCells ws, NAME_CELL, CODE_CELL, APPLICABLE_CELL, UNIT_CELL, ALIAS_CELL
    ws.Range(HEADER_CELL).value = HEADER_NEW

End Sub
