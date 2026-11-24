Attribute VB_Name = "Assembly"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' =========================================================
' MODEL - ASSEMBLY
' Table      : assemblies
' Error codes: entitiesAssembly (ErrorCodesEntities.bas)
'
' Bill of materials: links a parent sparepart to one child
' sparepart it is assembled from. A sparepart cannot be its
' own child (ck_assemblies_not_self).
' =========================================================

Private Const CLASS_NAME As String = "Assembly"

Private mId As Long
Private mParent As Sparepart
Private mChild As Sparepart
Private mCreatedAt As Date
Private mUpdatedAt As Date


' =========================================================
' ID                                    [id]
' =========================================================
Public Property Get Id() As Long
    Id = mId
End Property

Public Property Let Id(ByVal value As Long)

    If ValidationNumeric.IsZeroOrNegative(value) Then
        ErrorManager.ThrowError _
            entitiesAssembly.NegativeId, _
            CLASS_NAME & ".Id", _
            CLASS_NAME & ": Id harus lebih besar dari 0."
    End If

    mId = value

End Property


' =========================================================
' PARENT                                [parent_id] NOT NULL
' =========================================================
Public Property Get Parent() As Sparepart
    Set Parent = mParent
End Property

Public Property Set Parent(ByVal value As Sparepart)

    If value Is Nothing Then
        ErrorManager.ThrowError _
            entitiesAssembly.EmptyParent, _
            CLASS_NAME & ".Parent", _
            CLASS_NAME & ": Sparepart induk tidak boleh kosong."
    End If

    CheckNotSelf value, mChild, "Parent"

    Set mParent = value

End Property


' =========================================================
' CHILD                                 [child_id] NOT NULL
' =========================================================
Public Property Get Child() As Sparepart
    Set Child = mChild
End Property

Public Property Set Child(ByVal value As Sparepart)

    If value Is Nothing Then
        ErrorManager.ThrowError _
            entitiesAssembly.EmptyChild, _
            CLASS_NAME & ".Child", _
            CLASS_NAME & ": Sparepart komponen tidak boleh kosong."
    End If

    CheckNotSelf mParent, value, "Child"

    Set mChild = value

End Property


' =========================================================
' TIMESTAMPS        [created_at] [updated_at]
' Set by the database; 0 = NULL.
' =========================================================
Public Property Get CreatedAt() As Date
    CreatedAt = mCreatedAt
End Property

Public Property Let CreatedAt(ByVal value As Date)
    mCreatedAt = value
End Property

Public Property Get UpdatedAt() As Date
    UpdatedAt = mUpdatedAt
End Property

Public Property Let UpdatedAt(ByVal value As Date)
    mUpdatedAt = value
End Property


' =========================================================
' INTERNAL
' =========================================================

' Same sparepart = same Id when both have one, otherwise the
' same code.
Private Sub CheckNotSelf( _
    ByVal parentItem As Sparepart, _
    ByVal childItem As Sparepart, _
    ByVal propertyName As String)

    Dim isSame As Boolean

    If parentItem Is Nothing Or childItem Is Nothing Then Exit Sub

    If parentItem Is childItem Then
        isSame = True
    ElseIf parentItem.Id > 0 And childItem.Id > 0 Then
        isSame = (parentItem.Id = childItem.Id)
    ElseIf Len(parentItem.Code) > 0 And Len(childItem.Code) > 0 Then
        isSame = ValidationText.IsSameText(parentItem.Code, childItem.Code)
    End If

    If isSame Then
        ErrorManager.ThrowError _
            entitiesAssembly.ParentIsChild, _
            CLASS_NAME & "." & propertyName, _
            CLASS_NAME & _
            ": Sparepart tidak boleh menjadi komponen dirinya sendiri."
    End If

End Sub
