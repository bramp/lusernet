VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cUploads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cUpload"
Attribute VB_Ext_KEY = "Member0" ,"cUpload"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Const maxObjects = 10

'local variable to hold collection
Dim objects(maxObjects) As cUpload
Dim topPointer As Long

Public Function Add(Optional sKey As String) As cUpload
    'create a new object
    Dim objNewMember As cUpload
    Set objNewMember = New cUpload

    'set the properties passed into the method
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As cUpload
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Dim i As Long
    Dim lCount As Long
    
    For i = 0 To maxObjects
        If isNothing(objects(i)) Then lCount = lCount + 1
    Next i
    
    count = lCount
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)
    objects [vntIndexKey] = Nothing
End Sub

Private Sub Class_Initialize()
    'creates the collection when this class is created
    topPointer = 0
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Dim i As Long
    For i = 0 To topPointer
        objects(i) = Nothing
    Next i
End Sub

