VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Renaming File(s)..."
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6240
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Stop"
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar PBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   2143
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
   Begin VB.Label Label6 
      Caption         =   "..."
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   600
      Width           =   5175
   End
   Begin VB.Label Label5 
      Caption         =   "..."
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   360
      Width           =   5175
   End
   Begin VB.Label Label4 
      Caption         =   "..."
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label3 
      Caption         =   "To:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "From:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Progress:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Cancel As Boolean

Private Sub Command1_Click()
If Me.Cancel = True Then Unload Me
End Sub

Private Sub Form_Load()
Me.Show
Me.Visible = True
Me.Cancel = False
With Form1
    For i = 0 To .List1.ListCount - 1
        ListView1.ListItems.Add , , .ListView1.ListItems(i + 1).Text
        ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , .List1.List(i)
    Next i
End With
PBar1.Max = ListView1.ListItems.Count

StartProcess

End Sub

Public Function Perc(index As Integer) As String
Dim x As Integer
x = 100 / PBar1.Max * index
Perc = Int(x) & "%"
End Function

Public Function GetFilename(file As String) As String
Dim x, i As Integer
For i = 1 To Len(file)
    If Mid(file, i, 1) = "\" Or Mid(file, i, 1) = "/" Then x = i
Next i
GetFilename = Mid(file, x + 1, 255)
End Function

Public Function StartProcess()
Dim fiin As String
Dim fiout As String
Dim i As Integer
For i = 1 To ListView1.ListItems.Count
    If Me.Cancel = True Then GoTo err
    fiin = ListView1.ListItems(i).Text
    fiout = ListView1.ListItems(i).ListSubItems(1).Text
        Label4.Caption = GetFilename(fiin)
        Label5.Caption = GetFilename(fiout)
        Label6.Caption = i & " of " & PBar1.Max & " (" & Perc(i) & ")"
        PBar1.Value = i
        Name fiin As fiout
Next i
MsgBox "Complete!"
Command1.Caption = "&Close"
Me.Cancel = True
Exit Function
err:
MsgBox "Canceled!"
Command1.Caption = "&Close"
Me.Cancel = True
End Function
