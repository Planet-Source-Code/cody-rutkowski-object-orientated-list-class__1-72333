VERSION 5.00
Begin VB.Form frmDemo 
   Caption         =   "List Class DEMO"
   ClientHeight    =   3810
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   ScaleHeight     =   254
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRandomizeItems 
      Caption         =   "Randomize"
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdFindItem 
      Caption         =   "Find Item"
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdEditItem 
      Caption         =   "Edit Item"
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdDeleteItem 
      Caption         =   "Delete item"
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdAddItem 
      Caption         =   "Add item"
      Height          =   375
      Left            =   5520
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.ListBox lstListItems 
      Height          =   3180
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5295
   End
   Begin VB.Label lblHeader 
      AutoSize        =   -1  'True
      Caption         =   "Enumerated List Items:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1950
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private m_arrayList As arrayList

Private Sub EnumerateList(ByRef listbox As listbox, ByRef arrayList As arrayList)

    ' clear the listbox initially.
    listbox.Clear

    ' list all items with their keys.
    Dim lItem As Long
    For lItem = 0 To arrayList.Count - 1
    
        listbox.AddItem "key=" & arrayList.Key(lItem) & " value=" & arrayList.Item(lItem)
        listbox.ItemData(listbox.NewIndex) = lItem
        
    Next

End Sub

Private Sub cmdAddItem_Click()

    ' add new item to arraylist.
    m_arrayList.Add "new item " & Time$, "new_" & m_arrayList.Count

    ' re-enumerate items.
    Dim lSelItemIndex As Long
    lSelItemIndex = Me.lstListItems.ListIndex
    
    Call EnumerateList(lstListItems, m_arrayList)
    
    If (lSelItemIndex <= lstListItems.ListCount - 1) Then
        lstListItems.ListIndex = lSelItemIndex
    End If
    
End Sub

Private Sub cmdDeleteItem_Click()

    ' verify item has been selected.
    If lstListItems.ListIndex = -1 Then
        MsgBox "You must select an item to delete.", vbExclamation, "List Demo"
        Exit Sub
    End If
    
    ' delete item.
    m_arrayList.Delete lstListItems.ItemData(lstListItems.ListIndex)
    
    ' re-enumerate items.
    Dim lSelItemIndex As Long
    lSelItemIndex = Me.lstListItems.ListIndex
    
    Call EnumerateList(lstListItems, m_arrayList)
    
    If (lSelItemIndex <= lstListItems.ListCount - 1) Then
        lstListItems.ListIndex = lSelItemIndex
    End If
    
End Sub

Private Sub cmdEditItem_Click()

    ' verify item has been selected.
    If lstListItems.ListIndex = -1 Then
        MsgBox "You must select an item to edit.", vbExclamation, "List Demo"
        Exit Sub
    End If
    
    Dim sNewValue As String
    sNewValue = InputBox("Enter new value for item.", "Edit Item")
    
    ' check pointer of new value string. if 0, user canceled prompt.
    If StrPtr(sNewValue) > 0 Then
     
        ' update value.
        Dim sItemValue As String
        Dim lItemIndex As Long
        lItemIndex = lstListItems.ItemData(lstListItems.ListIndex)
        sItemValue = m_arrayList.Item(lItemIndex)
        m_arrayList.Item(lItemIndex) = sNewValue
        
        ' re-enumerate items.
        Dim lSelItemIndex As Long
        lSelItemIndex = Me.lstListItems.ListIndex
         
        Call EnumerateList(lstListItems, m_arrayList)
        
        If (lSelItemIndex <= lstListItems.ListCount - 1) Then
            lstListItems.ListIndex = lSelItemIndex
        End If
    
        ' notify user.
        MsgBox "Value for item updated.", vbInformation, "List Demo"
         
    
    Else
        ' user cancelled input prompt
    End If
    
End Sub

Private Sub cmdFindItem_Click()

    Dim sKey As String
    sKey = InputBox("Enter key of item that you would like to locate.", "Find Item By Key.")
    
    ' check pointer of new value string. if 0, user canceled prompt.
    If StrPtr(sKey) > 0 Then
        
        ' locate item by key.
        Dim sItemValue As String
        Dim lItemIndex As Long
        lItemIndex = m_arrayList.GetItemIndex(sKey)
        sItemValue = m_arrayList.Item(lItemIndex)
         
        ' prompt user on result.
        If lItemIndex > -1 Then
            MsgBox "Found item with value: " & sItemValue, vbExclamation, "List Demo"
        Else
            MsgBox "Unable to find item by supplied key.", vbExclamation, "List Demo"
        End If
    
    Else
        ' user cancelled input prompt
    End If
End Sub

Private Sub cmdRandomizeItems_Click()

    ' generate list again.
    Call RandomlyCreateList
    
End Sub

Private Sub Form_Load()
    ' object initialization.
    Set m_arrayList = New arrayList
    
    Call RandomlyCreateList
End Sub


Private Sub RandomlyCreateList()

    Randomize Timer

    ' generate random number of items to add
    Dim iItemsToAdd As Integer
    iItemsToAdd = 5 + (5 * Rnd) ' Min: 5 Max: 10
     
    m_arrayList.Clear
    
    While (m_arrayList.Count < iItemsToAdd)
    
        m_arrayList.Add Round(100 * Rnd), "itm_" & m_arrayList.Count
    
    Wend
    
    ' enumerate items.
    Call EnumerateList(lstListItems, m_arrayList)
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' clean up
    Set m_arrayList = Nothing
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    ' nicely re-arrange controls so that the form looks good even when resized.
    cmdAddItem.Left = Me.ScaleWidth - cmdAddItem.Width - 8
    cmdDeleteItem.Left = Me.ScaleWidth - cmdAddItem.Width - 8
    cmdAddItem.Left = Me.ScaleWidth - cmdAddItem.Width - 8
    cmdEditItem.Left = Me.ScaleWidth - cmdAddItem.Width - 8
    cmdFindItem.Left = Me.ScaleWidth - cmdAddItem.Width - 8
    cmdRandomizeItems.Move Me.ScaleWidth - cmdAddItem.Width - 8, ScaleHeight - 8 - cmdRandomizeItems.Height
    
    Me.lblHeader.Move 8, 8
    Me.lstListItems.Move 8, 24, ScaleWidth - cmdAddItem.Width - 24, ScaleHeight - 32
     
End Sub
