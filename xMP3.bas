Attribute VB_Name = "xMP3"
Public ValidTag As String
Public Title As String
Public Artist As String
Public Year As String
Public Album As String
Public Comment As String
Public Genre As Byte
Public valid As Boolean


Private Type ID3v1
    ValidTag As String * 3
    Title As String * 30
    Artist As String * 30
    Album As String * 30
    Year As String * 4
    Comment As String * 30
    Genre As Byte
    End Type


Public Sub getTag(MP3 As String)
On Error Resume Next
    Dim ID3 As ID3v1
    If Mid(GetFile(MP3), 1, 12) = "__INCOMPLETE" Then GoTo err
    Open MP3 For Binary As #1
    Get #1, FileLen(MP3) - 127, ID3
    Close #1


    With ID3
        ValidTag = .ValidTag
        Title = Trim(.Title)
        Artist = Trim(.Artist)
        Album = .Album
        Comment = .Comment
        Year = .Year
        Genre = .Genre
    End With
        valid = True
    Exit Sub
err:
valid = False
Artist = "none"
Title = "none"
End Sub

Public Function GetFile(file As String) As String
Dim x As Integer
Dim i As Integer
For i = 1 To Len(file)
If Mid(file, i, 1) = "\" Then x = i
Next i
GetFile = Mid(file, x + 1, 255)
End Function
Public Sub writeTag(MP3 As String)
    Dim ID3 As ID3v1


    With ID3
        .ValidTag = "TAG"
        .Title = Title
        .Artist = Artist
        .Album = Album
        .Comment = Comment
        .Year = Year
        .Genre = Genre
    End With
    On Error GoTo ErrMsg:
    Open MP3 For Binary As 1


    If ID3.ValidTag <> "TAG" Then
        Seek 1, LOF(1) + 1
    Else
        Seek 1, LOF(1) - 127
    End If
    Put 1, FileLen(MP3) - 127, ID3
    Close 1
    Exit Sub
ErrMsg:
    MsgBox ("File '" & MP3 & "' is marked as read-only or the file is In use." & vbCr & "Please correct and try again.")
End Sub


