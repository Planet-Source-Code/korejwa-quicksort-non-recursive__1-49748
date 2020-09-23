VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Quicksort Demo"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   4935
      Left            =   4320
      TabIndex        =   4
      Top             =   120
      Width           =   4095
      Begin VB.CommandButton cmdSortFilteredIndex 
         Caption         =   "Sort Elements Containing ""A"" By Index"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   4320
         Width           =   3495
      End
      Begin VB.ListBox List2 
         Height          =   2400
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   3495
      End
      Begin VB.CommandButton cmdSortByIndex 
         Caption         =   "Sort Complete Array By Index"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   3840
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   $"Form1.frx":0000
         Height          =   855
         Left            =   120
         TabIndex        =   7
         Top             =   2880
         Width           =   3855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.CommandButton cmdCreateList 
         Caption         =   "Create Random List"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   4320
         Width           =   1695
      End
      Begin VB.CommandButton cmdSortArray 
         Caption         =   "Sort String Array"
         Height          =   375
         Left            =   2160
         TabIndex        =   2
         Top             =   4320
         Width           =   1695
      End
      Begin VB.ListBox List1 
         Height          =   2400
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label2 
         Caption         =   "Quicksort is the fastest known general sorting algorithm for large arrays."
         Height          =   855
         Left            =   240
         TabIndex        =   9
         Top             =   2880
         Width           =   3615
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Base 0


Private Const NumElements          As Long = 100  'Total number of strings to sort
Private Const MinChars             As Long = 6    'Minimum number of characters per string
Private Const MaxChars             As Long = 16   'Maximum number of characters per string

Private MyStrings(NumElements - 1) As String      'An array of Strings



Private Sub cmdCreateList_Click()
    Dim i      As Long
    Dim j      As Long    'j term to avoid long pauses due to unreasonably large NumElements
    Dim s      As String  'Current string
    Dim Chars  As Long    'Number of characters in current string


   'Make MyStrings an array of random strings
    Randomize Timer
    For i = 0 To NumElements - 1
        Chars = Int((MaxChars - MinChars) * Rnd + MinChars)
        s = ""
        For j = 1 To Chars
            s = s & Chr$(Rnd * 25 + 65)
        Next j
        MyStrings(i) = s
    Next i

   'Display the array of strings
    If NumElements > 500 Then j = NumElements \ 100 Else j = 1
    List1.Clear
    For i = 0 To NumElements - 1 Step j
        List1.AddItem MyStrings(i)
    Next i

End Sub


Private Sub cmdSortArray_Click()
    Dim i      As Long
    Dim j      As Long    'j term to avoid long pauses due to unreasonably large NumElements


   'Sort the MyStrings array
    SortStringArray MyStrings 'Sort and display the MyStrings array

   'Display the MyStrings array
    If NumElements > 500 Then j = NumElements \ 100 Else j = 1
    List1.Clear
    For i = 0 To NumElements - 1 Step j
        List1.AddItem MyStrings(i)
    Next i

End Sub


Private Sub cmdSortByIndex_Click()
    Dim i     As Long
    Dim j     As Long 'j term to avoid long pauses due to unreasonably large NumElements
    Dim idx() As Long 'Index  to elements in MyStrings
    Dim idxs  As Long 'Number of elements in idx()


   'Create an array of indexes
    idxs = NumElements
    ReDim idx(idxs - 1)
    For i = 0 To idxs - 1
        idx(i) = i
    Next i

   'Sort the index array
    SortStringIndexArray MyStrings, idx, 0, idxs - 1

   'Display the sorted version
    If idxs > 500 Then j = idxs \ 100 Else j = 1
    List2.Clear
    For i = 0 To idxs - 1 Step j
        List2.AddItem MyStrings(idx(i))
    Next i

End Sub


'The Indexed example above includes an index to all the strings in the array.
'However, you'd usually Index an array if you:
'    Need to work with only some elements in the string array
'    Do not want to change the string array
'    Do not want to make an inefficient copy of the array elements

'The code below creates an index array that contains indexes only to strings
' that contain an "A" in them.  The list is sorted and displayed.
' The MyStrings array is not changed in any way.
Private Sub cmdSortFilteredIndex_Click()
    Dim i       As Long
    Dim j       As Long
    Dim idx()   As Long 'Index  to elements in MyStrings
    Dim idxs    As Long 'Number of elements in idx()


   'Create a filtered array of indexes
    idxs = 0
    ReDim idx(NumElements - 1)
    For i = 0 To NumElements - 1
        If InStr(1, MyStrings(i), "A") > 0 Then 'Include in the Index array
            idx(idxs) = i
            idxs = idxs + 1
        End If
    Next i

   'Sort the index array
    SortStringIndexArray MyStrings, idx, 0, idxs - 1

   'Display the sorted version
    If idxs > 500 Then j = idxs \ 100 Else j = 1 'j to avoid long delays for large arrays
    List2.Clear
    For i = 0 To idxs - 1 Step j
        List2.AddItem MyStrings(idx(i))
    Next i

End Sub


Private Sub Form_Load()
    cmdCreateList_Click
End Sub
