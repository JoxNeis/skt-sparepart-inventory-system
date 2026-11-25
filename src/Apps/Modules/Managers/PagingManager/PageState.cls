Attribute VB_Name = "PageState"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' =========================================================
' PAGE STATE (used by managerPaging)
'
' The page a recap sheet shows. Paging is done by the saved
' queries (keyset), so only the current page is kept:
'
'   Data       2D array (1-based) of the page. The first
'              columns match the table columns; extra columns
'              at the end may hold hidden keys (ids).
'   TotalRows  all rows of the search (the *_count query)
'   Page       1-based page number, PageSize rows per page
'   Request    the PageRequest that loaded this page
'   FirstKey / LastKey
'              sort keys of the first / last row: the start of
'              the previous / next page
' =========================================================

Private mData As Variant
Private mRowCount As Long
Private mColumnCount As Long
Private mTotalRows As Long
Private mPage As Long
Private mPageSize As Long
Private mRequest As PageRequest
Private mFirstKey As PageRequest
Private mLastKey As PageRequest


Public Sub SetPage( _
    ByVal data As Variant, _
    ByVal rowCount As Long, _
    ByVal totalRows As Long, _
    ByVal pageNumber As Long, _
    ByVal request As PageRequest, _
    ByVal firstKey As PageRequest, _
    ByVal lastKey As PageRequest)

    mData = data
    mRowCount = rowCount
    mTotalRows = totalRows
    mPage = pageNumber
    mPageSize = request.PageSize

    If IsArray(data) And rowCount > 0 Then
        mColumnCount = UBound(data, 2)
    Else
        mColumnCount = 0
    End If

    Set mRequest = request
    Set mFirstKey = firstKey
    Set mLastKey = lastKey

End Sub


' =========================================================
' PAGE
' =========================================================

Public Property Get TotalRows() As Long
    TotalRows = mTotalRows
End Property

Public Property Get Page() As Long
    Page = mPage
End Property

Public Property Get PageSize() As Long
    PageSize = mPageSize
End Property

Public Property Get PageCount() As Long

    If mTotalRows = 0 Or mPageSize <= 0 Then
        PageCount = 1
    Else
        PageCount = (mTotalRows + mPageSize - 1) \ mPageSize
    End If

End Property

Public Property Get Request() As PageRequest
    Set Request = mRequest
End Property

Public Property Get FirstKey() As PageRequest
    Set FirstKey = mFirstKey
End Property

Public Property Get LastKey() As PageRequest
    Set LastKey = mLastKey
End Property


' =========================================================
' ROWS OF THE PAGE
' =========================================================

Public Property Get PageRowCount() As Long
    PageRowCount = mRowCount
End Property

Public Property Get ColumnCount() As Long
    ColumnCount = mColumnCount
End Property

Public Property Get CellValue( _
    ByVal pageRow As Long, _
    ByVal columnIndex As Long) As Variant

    CellValue = mData(pageRow, columnIndex)

End Property

Public Property Let CellValue( _
    ByVal pageRow As Long, _
    ByVal columnIndex As Long, _
    ByVal newValue As Variant)

    mData(pageRow, columnIndex) = newValue

End Property

' Row of the page behind table row tableRow (0 when outside).
Public Function DataRowAt( _
    ByVal tableRow As Long) As Long

    If tableRow < 1 Or tableRow > mRowCount Then
        DataRowAt = 0
    Else
        DataRowAt = tableRow
    End If

End Function

' The page as a 2D array of columnCount columns (Empty when
' the page has no rows).
Public Function PageData( _
    ByVal columnCount As Long) As Variant

    Dim result() As Variant
    Dim r As Long
    Dim c As Long

    If mRowCount = 0 Then
        PageData = Empty
        Exit Function
    End If

    If columnCount > mColumnCount Then columnCount = mColumnCount

    ReDim result(1 To mRowCount, 1 To columnCount)

    For r = 1 To mRowCount
        For c = 1 To columnCount
            result(r, c) = mData(r, c)
        Next c
    Next r

    PageData = result

End Function
