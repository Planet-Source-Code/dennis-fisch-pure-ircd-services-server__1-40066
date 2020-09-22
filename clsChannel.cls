VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChannel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Base 1
Public Name As String
Public Topic As String
Public Modes As New Collection
Public TopicSetOn As Long
Public TopicSetBy As String
Public Index As Long
Public Password As String
Public UserLevels As New Collection
Public Invited As New Collection
Public Key As String
Public Limit As Integer

Public All As New Collection
Public Ops As New Collection
Public Voices As New Collection
Public NormUsers As New Collection

Public Bans As New Collection
Public Exceptions As New Collection
Public Invites As New Collection

 Public Property Get UserCount() As Long
 UserCount = NormUsers.Count + Ops.Count + Voices.Count
 End Property

Public Property Get GetAll() As String
Dim i As Long
For i = 1 To All.Count
    GetAll = GetAll & " " & All.Item(i)
Next i
GetAll = Trim(GetAll)
End Property

Public Property Get GetOps() As String
Dim i As Long
For i = 1 To Ops.Count
    GetOps = GetOps & " @" & Ops.Item(i)
Next i
GetOps = Trim(GetOps)
If GetOps <> "" Then GetOps = GetOps & " "
End Property

Public Property Get GetVoices() As String
Dim i As Long
For i = 1 To Voices.Count
    GetVoices = GetVoices & " +" & Voices.Item(i)
Next i
GetVoices = Trim(GetVoices)
If GetVoices <> "" Then GetVoices = GetVoices & " "
End Property

Public Property Get GetNorms() As String
Dim i As Long
For i = 1 To NormUsers.Count
    GetNorms = GetNorms & " " & NormUsers.Item(i)
Next i
GetNorms = Trim(GetNorms)
End Property

Public Function IsOp(Nick As String) As Boolean
Dim i As Long
For i = 1 To Ops.Count
    If Ops(i) = Nick Then
        IsOp = True
        Exit Function
    End If
Next i
End Function

Public Function IsVoice(Nick As String) As Boolean
Dim i As Long
For i = 1 To Voices.Count
    If Voices(i) = Nick Then
        IsVoice = True
        Exit Function
    End If
Next i
End Function

Public Function IsNorm(Nick As String) As Boolean
Dim i As Long
For i = 1 To NormUsers.Count
    If NormUsers(i) = Nick Then
        IsNorm = True
        Exit Function
    End If
Next i
End Function

Public Function IsMode(Mode As String) As Boolean
Dim i As Long
For i = 1 To Modes.Count
    If Modes(i) = Mode Then
        IsMode = True
        Exit Function
    End If
Next i
End Function

Public Function IsOnChan(NickName As String) As Boolean
Dim i As Long
For i = 1 To All.Count
    If All(i) = NickName Then
        IsOnChan = True
        Exit Function
    End If
Next i
End Function

Public Function IsBanned(User As clsUser) As Boolean
Dim i As Long
For i = 1 To Bans.Count
    If (User.Nick & "!" & User.Ident & "@" & User.DNS) Like Bans(i) Then
        IsBanned = True
        Exit Function
    End If
Next i
End Function

Public Function IsBanned2(ID As String) As Boolean
Dim i As Long
For i = 1 To Bans.Count
    If (ID Like Bans(i)) Then
        IsBanned2 = True
        Exit Function
    End If
Next i
End Function

Public Property Get GetModesForFile()
Dim i As Long
For i = 1 To Modes.Count
    GetModesForFile = GetModesForFile + Modes(i)
Next i
End Property

Public Property Get GetModes() As String
Dim i As Long
For i = 1 To Modes.Count
    GetModes = GetModes + Modes(i)
Next i
If Key <> "" And Limit <> 0 Then
    GetModes = GetModes & "lk " & Limit & " " & Key
ElseIf Key <> "" And Limit = 0 Then
    GetModes = GetModes & "k " & Key
ElseIf Key = "" And Limit <> 0 Then
    GetModes = GetModes & "l " & Limit
End If
End Property

Public Sub AddToUserList(NickName As String, level As Long)
On Error Resume Next
UserLevels.Add NickName & " " & level, NickName
End Sub

Public Sub RemoveFromUserList(NickName As String)
UserLevels.Remove NickName
End Sub

Public Sub GetUserListItem(Index, ByRef NickName As String, ByRef level As Long)
NickName = UserLevels(Index)
level = CLng(Mid(NickName, InStr(1, NickName, " ")))
NickName = Replace(NickName, " " & level, "")
End Sub

Public Sub AddModes(NewModes As String)
Dim i As Long
For i = 1 To Len(NewModes)
    Modes.Add Mid(NewModes, i, 1), Mid(NewModes, i, 1)
Next i
End Sub

Public Function ULOp(Nick As String) As Boolean
Dim User As clsUser, level As Long
Set User = NickToObject(Nick)
Dim i As Long
For i = 1 To UserLevels.Count
    GetUserListItem i, Nick, level
    If (Nick = User.Nick Or Nick = User.IdentifiedAs) And (level >= 100 And User.Identified) Then
        ULOp = True
    End If
Next i
End Function

Public Function ULVoice(Nick As String) As Boolean
Dim User As clsUser, level As Long
Set User = NickToObject(Nick)
Dim i As Long
For i = 1 To UserLevels.Count
    GetUserListItem i, Nick, level
    If (Nick = User.Nick Or Nick = User.IdentifiedAs) And (level = 50 And User.Identified) Then
        ULVoice = True
    End If
Next i
End Function

Public Function IsInvited(Nick As String) As Boolean
Dim i As Long
For i = 1 To Invited.Count
    If Nick = Invited(i) Then
        IsInvited = True
        Invited.Remove i
        Exit Function
    End If
Next i
End Function

Public Function IsInvited2(User As clsUser) As Boolean
Dim i As Long
For i = 1 To Invites.Count
    If (User.Nick & "!" & User.Ident & "@" & User.DNS) Like Invites(i) Then
        IsInvited2 = True
        Exit Function
    End If
Next i
End Function

Public Function IsException(User As clsUser) As Boolean
Dim i As Long
For i = 1 To Exceptions.Count
    If (User.Nick & "!" & User.Ident & "@" & User.DNS) Like Exceptions(i) Then
        IsException = True
        Exit Function
    End If
Next i
End Function

Public Function IsException2(ID As String) As Boolean
Dim i As Long
For i = 1 To Exceptions.Count
    If (ID Like Exceptions(i)) Then
        IsException2 = True
        Exit Function
    End If
Next i
End Function

Public Function IsInvited3(ID As String) As Boolean
Dim i As Long
For i = 1 To Invites.Count
    If (ID Like Invites(i)) Then
        IsInvited3 = True
        Exit Function
    End If
Next i
End Function

Private Sub Class_Initialize()
ChanCount = ChanCount + 1
MaxChans = MaxChans + 1
End Sub

Private Sub Class_Terminate()
ChanCount = ChanCount - 1
End Sub
