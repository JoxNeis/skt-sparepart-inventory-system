Attribute VB_Name = "PageRequest"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' =========================================================
' PAGE REQUEST
'
' Which rows a list query returns (keyset paging, done by the
' saved queries; see src/database/queries/order_list.sql):
'
'   FilterIndex / FilterValue
'       "contains" filter on column number FilterIndex of the
'       query (0 = every column). Blank value = no filter.
'   Direction NEXT + key  -> the rows AFTER the key
'   Direction PREV + key  -> the rows BEFORE the key
'       The key is the sort key of the last (NEXT) or first
'       (PREV) row of the current page: KeyDate + KeyId for the
'       document lists, KeyText + KeyId for the stock list,
'       KeyDate + KeyText + KeyText2 for the stock card.
'   PageSize
'       rows per page; 0 = every row (no paging).
'
' A new PageRequest is the first page of an unfiltered list.
' =========================================================

Public FilterIndex As Long
Public FilterValue As String
Public Direction As String
Public KeyDate As Date
Public KeyId As Long
Public KeyText As String
Public KeyText2 As String
Public PageSize As Long

Private Const DIRECTION_NEXT As String = "NEXT"
Private Const DIRECTION_PREV As String = "PREV"


Private Sub Class_Initialize()
    Direction = DIRECTION_NEXT
    KeyDate = #1/1/1900#
End Sub


Public Property Get IsPrevious() As Boolean
    IsPrevious = (Direction = DIRECTION_PREV)
End Property


' Same filter and page size, from the start of the list.
Public Function FirstPage() As PageRequest

    Dim result As PageRequest

    Set result = New PageRequest
    result.FilterIndex = FilterIndex
    result.FilterValue = FilterValue
    result.PageSize = PageSize

    Set FirstPage = result

End Function


' Same filter and page size; the rows after (NEXT) or before
' (PREV) the key held by keyRow.
Public Function Moved( _
    ByVal goForward As Boolean, _
    ByVal keyRow As PageRequest) As PageRequest

    Dim result As PageRequest

    Set result = FirstPage()

    If goForward Then
        result.Direction = DIRECTION_NEXT
    Else
        result.Direction = DIRECTION_PREV
    End If

    result.KeyDate = keyRow.KeyDate
    result.KeyId = keyRow.KeyId
    result.KeyText = keyRow.KeyText
    result.KeyText2 = keyRow.KeyText2

    Set Moved = result

End Function


' A key holder for one row (used for the first / last row of
' a page).
Public Function KeyOf( _
    ByVal rowDate As Date, _
    ByVal rowId As Long, _
    Optional ByVal rowText As String = vbNullString, _
    Optional ByVal rowText2 As String = vbNullString) As PageRequest

    Dim result As PageRequest

    Set result = New PageRequest
    result.KeyDate = rowDate
    result.KeyId = rowId
    result.KeyText = rowText
    result.KeyText2 = rowText2

    Set KeyOf = result

End Function
