Attribute VB_Name = "PagingLayout"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' =========================================================
' PAGING LAYOUT (used by managerPaging)
'
' Where the search / paging cells of one sheet are. Blank
' addresses mean the sheet does not have that part.
'
'   LOKASI          PlantCell
'   DARI / HINGGA   FromCell / ToCell (double-click = date picker)
'   HASIL PENCARIAN CountCell (BARIS DATA), PagesCell (TOTAL HALAMAN)
'   FILTER          FilterColumnCell (KOLOM), FilterValueCell (NILAI)
'   HALAMAN         PrevButton, PageCell (SEKARANG), NextButton
'   CARI            SearchButton
'   BARIS YANG DITAMPILKAN  PageSizeCell
' =========================================================

Public Key As String
Public TableName As String

Public PlantCell As String
Public FromCell As String
Public ToCell As String

Public CountCell As String
Public PagesCell As String

Public FilterColumnCell As String
Public FilterValueCell As String

Public PrevButton As String
Public PageCell As String
Public NextButton As String

Public SearchButton As String
Public PageSizeCell As String


' Layout of REKAP_ORDER, REKAP_PENDING_ORDER, REKAP_REQUEST
' and INPUT_PENERIMAAN (same positions on every sheet).
Public Sub UseStandardRecap( _
    ByVal layoutKey As String, _
    ByVal tableNameText As String)

    Key = layoutKey
    TableName = tableNameText

    PlantCell = "K5"
    FromCell = "K9"
    ToCell = "K10"

    CountCell = "M7"
    PagesCell = "N7"

    FilterColumnCell = "P7"
    FilterValueCell = "Q7"

    PrevButton = "S7"
    PageCell = "T7"
    NextButton = "U7"

    SearchButton = "S9:U9"
    PageSizeCell = "U10"

End Sub
