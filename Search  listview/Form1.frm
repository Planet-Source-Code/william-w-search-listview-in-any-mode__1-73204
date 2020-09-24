VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form Form1 
   Caption         =   "Search Listview"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   11700
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option1 
      Caption         =   "None"
      Height          =   255
      Index           =   7
      Left            =   10680
      TabIndex        =   14
      ToolTipText     =   "Restrict to column"
      Top             =   5880
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Column 6"
      Height          =   255
      Index           =   6
      Left            =   10680
      TabIndex        =   13
      ToolTipText     =   "Restrict to column"
      Top             =   5640
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Column 5"
      Height          =   255
      Index           =   5
      Left            =   9600
      TabIndex        =   12
      ToolTipText     =   "Restrict to column"
      Top             =   5880
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Column 4"
      Height          =   255
      Index           =   4
      Left            =   9600
      TabIndex        =   11
      ToolTipText     =   "Restrict to column"
      Top             =   5640
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Column 3"
      Height          =   255
      Index           =   3
      Left            =   8520
      TabIndex        =   10
      ToolTipText     =   "Restrict to column"
      Top             =   5880
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Column 2"
      Height          =   255
      Index           =   2
      Left            =   8520
      TabIndex        =   9
      ToolTipText     =   "Restrict to column"
      Top             =   5640
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Column 1"
      Height          =   255
      Index           =   1
      Left            =   7440
      TabIndex        =   8
      ToolTipText     =   "Restrict to column"
      Top             =   5880
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Column 0"
      Height          =   255
      Index           =   0
      Left            =   7440
      TabIndex        =   7
      ToolTipText     =   "Restrict to column"
      Top             =   5640
      Width           =   975
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Multi Selection"
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Case Sensitive"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Search"
      Height          =   255
      Left            =   6240
      TabIndex        =   3
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   5640
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Fill Listview"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   5640
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9340
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Column 1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Column 2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Column 3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Column 4"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Column 5"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Column 6"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Column 7"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Found 0"
      Height          =   195
      Left            =   4920
      TabIndex        =   4
      Top             =   6000
      Width           =   585
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lBegin As Long
'Search Listview with Subitems William W 2010


Private Sub Command1_Click()

   'Load a bunch of random stuff into the listview
  Dim LvItem As ListItem
  Dim A As Long
  Dim B As Long
  Dim RandomText As String
  Dim LastTimer As Single

   LastTimer = Timer

   Do Until Timer >= LastTimer + 5

      For A = 35 To 123
         RandomText = Chr$(A) & Chr$(A) & Chr$(A) & Chr$(A) & Chr$(A) & Chr$(A)

         Set LvItem = ListView1.ListItems.Add(, , RandomText)

         For B = 1 To ListView1.ColumnHeaders.Count - 1

            If B = 6 Then
               LvItem.SubItems(B) = Val(Round(Timer - LastTimer, 1))
             Else
               LvItem.SubItems(B) = "Sub Item " & B & " " & Chr$(A) & Chr$(A) & Chr$(A) & Chr$(B + _
                  100)
            End If

         Next
      Next
      DoEvents
   Loop

End Sub

Private Sub Command2_Click()

  Dim A As Long
  Dim RestrictedCol As Long

   For A = 0 To 7

      If Option1(A) = True Then
         If A = 7 Then
            RestrictedCol = -1
          Else
            RestrictedCol = A
            Exit For
         End If

      End If
   Next A

   lBegin = lBegin + 1

   SearchListVw ListView1, Text1.Text, lBegin, Check1.Value, Check2.Value, RestrictedCol

End Sub

Private Sub InitSearchListVw(LV As ListView)

   'Fills the Tag property of the Main item with the data contained in each column/Subitem
   On Error GoTo InitErr
  Dim lItem As Long
  Dim lSubitem As Long

   For lItem = 1 To LV.ListItems.Count
      For lSubitem = 0 To LV.ColumnHeaders.Count - 1

         If lSubitem = 0 Then
            'You could use LV.ListItems(a).listSubItems(1).tag also if you were already using the
            'main item tags
            LV.ListItems(lItem).Tag = LV.ListItems(lItem).Text
            ' index 0 stands for the first column or the Main
            ' item in the listview as there is no subitem 0 this will have to be
            ' conditional
          Else
            'Item;Subitem1;Subitem2;Subitem3......
            LV.ListItems(lItem).Tag = LV.ListItems(lItem).Tag & ";" & _
               LV.ListItems(lItem).SubItems(lSubitem)

         End If
      Next lSubitem

   Next lItem
   'Delete the data in lv.tag to make it re-search or add or remove an item
   '(Lv.tag="")
   LV.Tag = "Loaded " & lItem
   Debug.Print LV.Tag & " Items"

InitErr:

   Select Case Err.Number
    Case 0:
      'No Error
    Case 91: 'Empty List
    Case Else:
      Debug.Print Err.Number; Err.Description

   End Select

End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)

   lBegin = Item.Index

End Sub

Private Sub SearchListVw(LV As ListView, _
                         Stext As String, _
                         Optional ByRef Start As Long = 1, _
                         Optional CaseSens As Boolean = False, _
                         Optional MultiSelect As Boolean = True, _
                         Optional RestrictedColumn As Long = -1)

   'Search a Listview subitems included
   On Error GoTo SearchErr
  Dim lFound As Long
  Dim lIndex As Long
  Dim lItem As Long
  Dim lPos As Long

   LV.MultiSelect = MultiSelect
   LV.HideSelection = False 'Needed
   LV.FullRowSelect = True 'optional....

   LV.ListItems(LV.SelectedItem.Index).Selected = False
   lIndex = 0
   If Start < 1 Then Start = 1

   'Here we prevent having to load the items and sub items in to the tag properties multiple times
   '   unless there are more items added...
   If InStr(1, LV.Tag, LV.ListItems.Count + 1) = 0 Then InitSearchListVw LV

   For lItem = Start To LV.ListItems.Count

      If CaseSens = True Then
         'Search is case sensitive so no text formatting..
         lPos = InStr(1, LV.ListItems(lItem).Tag, Stext)

         If RestrictedColumn <> -1 Then
            If InStr(1, Split(LV.ListItems(lItem).Tag, ";")(RestrictedColumn), Stext) = 0 Then lPos _
               = 0
         End If

       Else
         ' if its not to be case sensitive then make
         '   everything uppercase
         Stext = UCase$(Stext)
         lPos = InStr(1, UCase$(LV.ListItems(lItem).Tag), Stext)

         If RestrictedColumn <> -1 Then
            'Split the tag into an array check the desired colum for a match
            If InStr(1, UCase$(Split(LV.ListItems(lItem).Tag, ";")(RestrictedColumn)), Stext) = 0 _
               Then lPos = 0 'Not found..
         End If

      End If

      If lPos <> 0 Then

         If lIndex = 0 Then lIndex = lItem 'If the first item hasnt been selected select it now

         lFound = lFound + 1
         If MultiSelect = True Then LV.ListItems(lItem).Selected = True
         'Select All Matches in the listview

       Else
         'Not Found
         LV.ListItems(lItem).Selected = False
      End If

   Next

   If lIndex = 0 Then
      Label1.Caption = "Not Found"
      Start = 1
    Else
      Label1.Caption = "Found " & lFound
      Start = lIndex
      LV.ListItems(lIndex).Selected = True
      LV.ListItems(lIndex).EnsureVisible
   End If

SearchErr:

   Select Case Err.Number
    Case 0:
      'No Error
    Case 91: 'Empty List
    Case Else:
      Debug.Print Err.Number; Err.Description

   End Select

End Sub

Private Sub Text1_Change()
'Checks for a match as you type
  Dim A As Long
  Dim RestrictedCol As Long

   For A = 0 To 7

      If Option1(A) = True Then
         If A = 7 Then
            RestrictedCol = -1
          Else
            RestrictedCol = A
            Exit For
         End If

      End If
   Next A

   SearchListVw ListView1, Text1.Text, lBegin, Check1.Value, Check2.Value, RestrictedCol

End Sub

