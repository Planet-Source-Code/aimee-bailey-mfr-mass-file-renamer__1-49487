VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mass File Renamer"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5985
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check4 
      Caption         =   "Show Summary Before Process"
      Height          =   255
      Left            =   1800
      TabIndex        =   18
      Top             =   5640
      Width           =   2535
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   6000
      TabIndex        =   15
      Top             =   360
      Width           =   3375
   End
   Begin MSComDlg.CommonDialog CMD 
      Left            =   5160
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Process Files >"
      Height          =   375
      Left            =   4440
      TabIndex        =   12
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Rename Options"
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   5775
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   1200
         Top             =   960
      End
      Begin VB.CheckBox Check3 
         Caption         =   "LCase Ext's"
         Height          =   255
         Left            =   4200
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Remove ""_"" Charicter"
         Height          =   255
         Left            =   2160
         TabIndex        =   16
         Top             =   240
         Width           =   3135
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "Form1.frx":0442
         Left            =   4080
         List            =   "Form1.frx":0449
         TabIndex        =   8
         Text            =   "[Keep The Same]"
         Top             =   600
         Width           =   1575
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Form1.frx":045E
         Left            =   1440
         List            =   "Form1.frx":0468
         TabIndex        =   7
         Text            =   "[Keep The Same]"
         Top             =   600
         Width           =   2535
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form1.frx":048E
         Left            =   120
         List            =   "Form1.frx":049B
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "First Letter Uppercased"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label2 
         Height          =   255
         Left            =   960
         TabIndex        =   10
         Top             =   1080
         Width           =   4695
      End
      Begin VB.Label Label1 
         Caption         =   "Example:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Files To Rename"
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin VB.CommandButton Command6 
         Caption         =   "&Clear"
         Height          =   285
         Left            =   4440
         TabIndex        =   19
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Add Directory..."
         Height          =   285
         Left            =   1320
         TabIndex        =   13
         Top             =   3360
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Remove Selected"
         Height          =   285
         Left            =   2880
         TabIndex        =   3
         Top             =   3360
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add File..."
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   3360
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3015
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   5318
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Filename(s)"
            Object.Width           =   9234
         EndProperty
      End
   End
   Begin VB.FileListBox File1 
      Height          =   1650
      Left            =   960
      TabIndex        =   14
      Top             =   960
      Width           =   3615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Create Summary"
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   5640
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Cancel As Boolean

Private Sub Command1_Click()
Dim aa() As String
Dim bb As String
On Error GoTo err
CMD.CancelError = True
CMD.Flags = cdlOFNExplorer 'Or cdlOFNAllowMultiselect
CMD.Filter = "All Files [*.*]|*.**"
CMD.DialogTitle = "Add file..."
CMD.ShowOpen

Me.Caption = GetEXT(CMD.filename)

ListView1.ListItems.Add , , CMD.filename
err:
End Sub

Private Sub Command2_Click()
On Error Resume Next
ListView1.ListItems.Remove ListView1.SelectedItem.index
End Sub

Public Function DoExample(lst As ListView)
Dim currFilename As String
Dim currExt As String
Dim currPath As String
Dim a, b, c As String
a = "": b = "": c = ""
    If ListView1.ListItems.Count <= 0 Then Exit Function
    currExt = GetEXT(lst.ListItems(1).Text)
    currFilename = GetFile(lst.ListItems(1).Text)
    currPath = GetDIR(lst.ListItems(1).Text)
    
    If Combo1.ListIndex = 1 Then
        a = "1. "
    End If
    
    b = Trim(GetNewFilename(currFilename, Combo2.Text, lst.ListItems(1).Text))
    
    If Trim(RTrim(b)) = "" Then b = GetFile(lst.ListItems(1).Text)
    
    If Combo3.Text = "[Keep The Same]" Then
        c = GetEXT(lst.ListItems(1).Text)
    Else
        If Check3.Value = 1 Then
            c = LCase(DoDot(Combo3.Text))
        Else
            c = DoDot(Combo3.Text)
        End If
    End If
    
    If Check2.Value = 1 Then
        b = Replace(b, "_", " ")
    End If
    
    If Check1.Value = 1 Then
        b = UCase(Mid(b, 1, 1)) & Mid(b, 2, 255)
    End If
    
    Label2.Caption = currPath & a & b & c
End Function

Public Function GetNewFilenames(lst As ListView, tolist As ListBox)
Dim currFilename As String
Dim currExt As String
Dim currPath As String
Dim a, b, c As String
tolist.Clear
For i = 1 To lst.ListItems.Count
    a = "": b = "": c = ""
    currExt = GetEXT(lst.ListItems(i).Text)
    currFilename = GetFile(lst.ListItems(i).Text)
    currPath = GetDIR(lst.ListItems(i).Text)
    
    If Combo1.ListIndex = 1 Then
        a = i & ". "
    End If
    
    b = Trim(GetNewFilename(currFilename, Combo2.Text, lst.ListItems(i).Text))
    
    If Trim(RTrim(b)) = "" Then b = GetFile(lst.ListItems(i).Text)
    
    If Combo3.Text = "[Keep The Same]" Then
        c = GetEXT(lst.ListItems(i).Text)
    Else
        If Check3.Value = 1 Then
            c = LCase(DoDot(Combo3.Text))
        Else
            c = DoDot(Combo3.Text)
        End If
    End If
    
    If Check2.Value = 1 Then
        b = Replace(b, "_", " ")
    End If
    
    If Check1.Value = 1 Then
        b = UCase(Mid(b, 1, 1)) & Mid(b, 2, 255)
    End If
    
    tolist.AddItem currPath & a & b & c
    
Next i

End Function

Public Function GetNewFilename(current As String, args As String, filename As String) As String
If args = "[Grab From ID3]" Then
    xMP3.getTag filename
    If xMP3.valid = True Then
        GetNewFilename = xMP3.Artist & " - " & xMP3.Title
    Else
        GetNewFilename = current
    End If
ElseIf args = "[Keep The Same]" Then
    GetNewFilename = current
Else
    GetNewFilename = args
End If

End Function

Public Function GetNewExt(default As String, Optional newExt As String)
If newExt = "" Then
    GetNewExt = DoDot(default)
Else
    GetNewExt = DoDot(newExt)
End If
End Function

Public Function DoDot(str As String) As String
If Check3.Value = 1 Then
    x = "." & LCase(Replace(str, ".", ""))
Else
    x = "." & Replace(str, ".", "")
End If
DoDot = x
End Function
Private Sub Command3_Click()
If Combo2.Text = "[Grab From ID3]" Then
    x = MsgBox("Grabbing Info From The ID3 Tags Should only be used when all files already have an ID3 tag!! do you want to proceed?", vbYesNo, "Question")
    If x = vbYes Then
    GetNewFilenames ListView1, List1
    Else
        Exit Sub
    End If
Else
GetNewFilenames ListView1, List1
End If
End Sub

Private Sub Command4_Click()
On Error Resume Next
Command3_Click
Me.Cancel = False
If Check4.Value = 1 Then
    Form2.Show vbModal, Me
    Do Until Form2.Visible = False
        DoEvents
    Loop
End If
If Me.Cancel = True Then Exit Sub
    Form3.Show vbModal, Me
End Sub

Private Sub Command5_Click()
On Error GoTo err
CMD.CancelError = True
CMD.Flags = cdlOFNExplorer 'Or cdlOFNAllowMultiselect
CMD.Filter = "All Files [*.*]|*.**"
CMD.DialogTitle = "Choose file from directory..."
CMD.ShowOpen

File1.path = GetDIR(CMD.filename)
File1.Refresh


For i = 0 To File1.ListCount - 1
    ListView1.ListItems.Add , , DoDir(GetDIR(CMD.filename)) & File1.List(i)
Next i
err:
End Sub

Public Function DoDir(path As String) As String
Dim x As String

If InStr(1, path, "\") > 1 Then
    If Right(path, 1) <> "\" Then x = "\" Else x = ""
ElseIf InStr(1, path, "/") >= 1 Then
    If Right(path, 1) <> "/" Then x = "/" Else x = ""
End If

DoDir = path & x
End Function

Public Function GetFile(file As String) As String
Dim x As String
Dim a, b, i As Integer

For i = 1 To Len(file)
    If Mid(file, i, 1) = "/" Or Mid(file, i, 1) = "\" Then
        a = i
    End If
Next i

xfile = Mid(file, a + 1, 255)

b = InStrRev(xfile, ".")

GetFile = Mid(xfile, 1, b - 1)

End Function

Public Function GetEXT(file As String) As String
Dim x As Integer
x = InStrRev(file, ".")
GetEXT = Mid(file, x, 255)
End Function

Public Function GetDIR(file As String) As String
Dim x, i As Integer
'On Error GoTo err
For i = 1 To Len(file)
    If Mid(file, i, 1) = "\" Or Mid(file, i, 1) = "/" Then
        x = i
    End If
Next i
GetDIR = Mid(file, 1, x)
Exit Function
err:
GetDIR = file
End Function

Private Sub List1_Click()
Me.Caption = List1.Text
End Sub

Private Sub Timer1_Timer()
DoExample ListView1
End Sub
