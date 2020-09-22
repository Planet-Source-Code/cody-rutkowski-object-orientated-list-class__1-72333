VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
' -----------------------------------------------------------------
' 2009 AUG 02
'
' class written to create an object-orientated approach
' to array handling. Supplies nice re-usable methods that
' can help keep programs clean and free from common programming
' errors when working with VB6 arrays.
'
' Feel free to use this code as you please. Distribute in programs
' as you wish as long as you do not claim that you wrote this code.
' No credit needs to be given to the original author.
'
'       Thanks and Happy Coding,
'       Cody Rutkowski
'
' -----------------------------------------------------------------

Private m_count As Long ' store count of items in collection.
Private m_collection()  ' array that contains items
Private m_key()         ' array that contains key for each item in collection array.
Private m_allowDupes As Boolean ' determines whether array can hold the same value more than once.

Public Property Get Count() As Long
Attribute Count.VB_Description = "Gets count of items in array."
    Count = m_count
End Property

Public Property Get Item(ByVal index As Integer) As Variant
Attribute Item.VB_Description = "Get an item's value at a specific index."
    ' Grab the item at the selected index.
    If IsObject(m_collection(index)) Then
        Set Item = m_collection(index)
    Else
        Let Item = m_collection(index)
    End If
End Property

Public Property Get Key(ByVal index As Integer) As Variant
Attribute Key.VB_Description = "Get the key of an item at a specific index."
    ' Return the key of an item at position defined by index.
    Key = m_key(index)
End Property

Public Property Get GetItemIndex(ByVal Key) As Integer
Attribute GetItemIndex.VB_Description = "Get the index of an item by the items key."
    ' Locate the index of an item in the array by its key
    Dim index As Integer
    For index = 0 To m_count - 1
        If Key = m_key(index) Then
            GetItemIndex = index
            Exit Property
        End If
    Next
    
    ' No item found by key. Return -1.
    GetItemIndex = -1
End Property

Public Function Contains(ByRef var) As Boolean
Attribute Contains.VB_Description = "Check if an item is contained in the array by checking for the same value."
    ' Loop through array and check to see if item is already
    ' contained within the array of items.
    Dim index As Long
    For index = 0 To m_count - 1
        If m_collection(index) = var Then
            Contains = True
            Exit Function
        End If
    Next
    
    ' item not contained in array.
    Contains = False
    
End Function
  
Public Sub Add(ByRef var, Optional ByVal Key)
Attribute Add.VB_Description = "Add an item to the array with a specific key."

    ' make sure item key is not already in use.
    If IsMissing(Key) = False Then
        If (GetItemIndex(Key) > -1) Then Exit Sub
    End If
    If AllowDuplicates = False Then
        If Contains(var) = True Then
            Err.Raise vbObjectError, , "ArrayList already contains item with value. Duplicates not allowed."
            Exit Sub
        End If
    End If
       
    ' resize the array for items and keys.
    m_count = m_count + 1
    ReDim Preserve m_collection(m_count - 1)
    ReDim Preserve m_key(m_count - 1)
    
    ' check to see if the item being added is an item. if it is
    ' set it to the new position in the array. Others use let.
    If IsObject(var) Then
        Set m_collection(m_count - 1) = var
    Else
        Let m_collection(m_count - 1) = var
    End If
    m_key(m_count - 1) = Key
    
End Sub

Public Function Equals(ByRef obj As arrayList) As Boolean
Attribute Equals.VB_Description = "Check to see if array object is equal to another array object."

    On Error GoTo errHandler
 
    ' Compare Count.
    If Not (obj.Count = Me.Count) Then GoTo NotEqual
    ' Compare items.
    Dim index As Long
    For index = 0 To obj.Count - 1
        If Not (obj.Item(index) = Me.Item(index)) Then GoTo NotEqual
        If Not (obj.Key(index) = Me.Key(index)) Then GoTo NotEqual
    Next
    
    ' They are equal.
    Equals = True
    
    Exit Function
errHandler:
    Err.Clear
NotEqual:
    ' At least 1 item did not match up between both array classes.
    Equals = False
    
End Function

Public Property Get IsEmpty() As Boolean
Attribute IsEmpty.VB_Description = "Gets whether array is empty (when count = 0)."
     IsEmpty = (m_count = 0)
End Property

Public Function ToArray() As Variant
Attribute ToArray.VB_Description = "Return the array class as a standard vb6 array data type."

    If IsEmpty Then
        ' return empty array.
        ToArray = Array()
    Else
        ' return array of items.
        ToArray = m_collection
    End If

End Function

Public Sub Clear()
Attribute Clear.VB_Description = "Clear all items from array and all keys. Resets count of array items to zero."
    ' resize arrays and set count to 0.
    m_count = 0
    ReDim m_collection(0)
    ReDim m_key(0)
End Sub
 
Public Sub Delete(ByVal index As Integer)
Attribute Delete.VB_Description = "Delete an item in the array at a given index."

    ' move all items in array from index +1 to the item before it.
    Dim i As Integer
    For i = index To m_count - 2
        m_collection(i) = m_collection(i + 1)
        m_key(i) = m_key(i + 1)
    Next
    m_count = m_count - 1
    
    ' resize collection and key array.
    If m_count > 0 Then
        ReDim Preserve m_collection(m_count - 1)
        ReDim Preserve m_key(m_count - 1)
    Else
        Call Clear
    End If
    
End Sub

Public Property Get AllowDuplicates() As Boolean
Attribute AllowDuplicates.VB_Description = "Get/Set whether duplicate values can be added to array. If set to false, all duplicates added previously will remain."
    AllowDuplicates = m_allowDupes
End Property
Public Property Let AllowDuplicates(vNewValue As Boolean)
    m_allowDupes = vNewValue
End Property

Private Sub Class_Initialize()
    m_allowDupes = True
End Sub
