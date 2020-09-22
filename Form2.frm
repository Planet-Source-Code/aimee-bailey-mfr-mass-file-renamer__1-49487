VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mass File Renamer - Summary..."
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8145
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Accept >>"
      Default         =   -1  'True
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   5160
      Width           =   1335
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   8705
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Current Filename"
         Object.Width           =   6703
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "New Filename"
         Object.Width           =   6703
      EndProperty
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Cancel = True
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
With Form1
    For i = 0 To .List1.ListCount - 1
        ListView1.ListItems.Add , , .ListView1.ListItems(i + 1).Text
        ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , .List1.List(i)
    Next i
End With
End Sub
