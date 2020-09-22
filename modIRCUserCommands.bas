Attribute VB_Name = "modIRCUserCommands"
Option Explicit
Dim NickName As String, level As Long

Public Function ChangeNick(Index As Long, NewNick As String, Optional SendLink As Boolean = True) As Boolean
On Error Resume Next
Dim i As Long, X As Long, CurChan As clsChannel, CurNick As clsUser
For i = 1 To Users(Index).Onchannels.Count
    Set CurChan = ChanToObject(Users(Index).Onchannels(i))
    If CurChan.IsNorm(Users(Index).Nick) Then
        CurChan.NormUsers.Remove Users(Index).Nick
        CurChan.NormUsers.Add NewNick, NewNick
    ElseIf CurChan.IsVoice(Users(Index).Nick) Then
        CurChan.Voices.Remove Users(Index).Nick
        CurChan.Voices.Add NewNick, NewNick
    ElseIf CurChan.IsOp(Users(Index).Nick) Then
        CurChan.Ops.Remove Users(Index).Nick
        CurChan.Ops.Add NewNick, NewNick
    End If
    CurChan.All.Remove Users(Index).Nick
    CurChan.All.Add NewNick, NewNick
Next i
'1 = Command, 2 = Nick, 3 = NewNick
If SendLink Then SendLinks "Nick" & vbLf & Users(Index).Nick & vbLf & NewNick
Users(Index).Nick = NewNick
ChangeNick = True
End Function

Public Sub SendQuit(Index As Long, QuitMsg As String, Optional Kill As Boolean = False, Optional SendLink As Boolean = True)
On Error Resume Next
Dim i As Long, X As Long, CurChan As clsChannel
If QuitMsg = "" Then QuitMsg = DefQuit
Users(Index).SentQuit = True
For i = 1 To Users(Index).Onchannels.Count
    Set CurChan = ChanToObject(Users(Index).Onchannels(i))
    CurChan.All.Remove Users(Index).Nick
    If CurChan.IsNorm(Users(Index).Nick) Then
        CurChan.NormUsers.Remove Users(Index).Nick
    ElseIf CurChan.IsVoice(Users(Index).Nick) Then
        CurChan.Voices.Remove Users(Index).Nick
    ElseIf CurChan.IsOp(Users(Index).Nick) Then
        CurChan.Ops.Remove Users(Index).Nick
    End If
Next i
If Kill Then
    '1 = Command, 2 = Nick, 3 = Reason
    If SendLink Then SendLinks "KillUser" & vbLf & Users(Index).Nick & vbLf & QuitMsg
Else
    '1 = Command, 2 = Nick, 3 = QuitMsg
    If SendLink Then SendLinks "QuitUser" & vbLf & Users(Index).Nick & vbLf & QuitMsg
End If
End Sub

Public Sub SendPart(Index As Long, Chan As String, Reason As String, Optional SendLink As Boolean = True)
On Error Resume Next
Dim i As Long, X As Long, CurChan As clsChannel, Found As Boolean
Users(Index).Idle = UnixTime
For i = 1 To Users(Index).Onchannels.Count
    If Chan = Users(Index).Onchannels(i) Then Found = True
Next i
If Found = False Then Exit Sub
'1 = Command, 2 = Nick, 3 = Channel, 4 = Reason
If SendLink Then SendLinks "PartUser" & vbLf & Users(Index).Nick & vbLf & Chan & vbLf & Reason
Set CurChan = ChanToObject(Chan)
CurChan.All.Remove Users(Index).Nick
If CurChan.IsNorm(Users(Index).Nick) Then
    CurChan.NormUsers.Remove Users(Index).Nick
ElseIf CurChan.IsVoice(Users(Index).Nick) Then
    CurChan.Voices.Remove Users(Index).Nick
ElseIf CurChan.IsOp(Users(Index).Nick) Then
    CurChan.Ops.Remove Users(Index).Nick
End If
Users(Index).Onchannels.Remove Chan
If CurChan.All.Count = 0 And Not CurChan.IsMode("r") Then Set Channels(CurChan.Index) = Nothing
End Sub

Public Sub NotifyJoin(Index As Long, Chan As String, Optional SendLink As Boolean = True)
'1 = Command, 2 = Nick, 3 = Channel
If SendLink Then SendLinks "JoinChan" & vbLf & Users(Index).Nick & vbLf & Chan
End Sub

Public Sub SendNotice(Target As String, Message As String, User As String, Optional ToChannel As Boolean = False, Optional Index As Integer, Optional SendLink As Boolean = True)
On Error Resume Next
Dim TargetIndex As Long
If Index <> 0 Then
    TargetIndex = Index
Else
    TargetIndex = NickToObject(Target).Index
End If
SendLinks "NoticeUser" & vbLf & User & vbLf & Users(TargetIndex).Nick & vbLf & Message
End Sub

Public Sub KickUser(Source As String, Chan As String, Target As String, Optional Reason As String, Optional Reasoning As Boolean = False, Optional SendLink As Boolean = True)
On Error Resume Next
Dim i As Long, Channel As clsChannel, KickMsg As String
Set Channel = ChanToObject(Chan)
If Reasoning = True Then
    KickMsg = ":" & Source & " KICK " & Chan & " " & Target & " :" & Reason
Else
    KickMsg = ":" & Source & " KICK " & Chan & " " & Target
End If
Channel.All.Remove Target
If Channel.IsNorm(Target) Then
    Channel.NormUsers.Remove Target
ElseIf Channel.IsVoice(Target) Then
    Channel.Voices.Remove Target
ElseIf Channel.IsOp(Target) Then
    Channel.Ops.Remove Target
End If
NickToObject(Target).Onchannels.Remove Chan
'1 = Command, 2 = Nick, 3 = Channel, 4 = Reason, 5 = Target
If SendLink Then SendLinks "KickUser" & vbLf & Source & vbLf & Channel.Name & vbLf & Reason & vbLf & Target
End Sub

Public Sub SetTopic(Chan As String, NewTopic As String, User As String, Optional SendLink As Boolean = True)
Dim Channel As clsChannel
Set Channel = ChanToObject(Chan)
Channel.Topic = NewTopic
Channel.TopicSetOn = UnixTime
Channel.TopicSetBy = User
If Channel.IsMode("r") Then
    Dim dab As New clsDatabase
    With dab
        .FileName = App.Path & "\Channels\" & Channel.Name & ".dat"
        .WriteEntry "General", "Topic", NewTopic
        .WriteEntry "General", "TopicSetOn", Channel.TopicSetOn
        .WriteEntry "General", "TopicSetBy", User
    End With
End If
'1 = Command, 2 = Nick, 3 = Channel, 4 = Target
If SendLink Then SendLinks "SetTopic" & vbLf & User & vbLf & Channel.Name & vbLf & "" & vbLf & NewTopic
End Sub

Public Sub OpUser(Channel As clsChannel, Target As String, User As String, Optional OpAnyway As Boolean = False, Optional SendLink As Boolean = True)
On Error Resume Next
Dim Chan As String
Chan = Channel.Name
If Channel.IsOp(Target) And OpAnyway = False Then Exit Sub
If Channel.IsNorm(Target) Then Channel.NormUsers.Remove Target
Channel.Ops.Add Target, Target
'1 = Command, 2 = Nick, 3 = Channel, 4 = Target
If SendLink Then SendLinks "OpUser" & vbLf & User & vbLf & Channel.Name & vbLf & "" & vbLf & Target
End Sub

Public Sub DeOpUser(Channel As clsChannel, Target As String, User As String, Optional SendLink As Boolean = True)
On Error Resume Next
Dim Chan As String
If Not Channel.IsOp(Target) Then Exit Sub
Chan = Channel.Name
Channel.Ops.Remove Target
If Channel.IsVoice(Target) Then
Else
    Channel.NormUsers.Add Target, Target
End If
'1 = Command, 2 = Nick, 3 = Channel, 4 = Target
If SendLink Then SendLinks "DeOpUser" & vbLf & User & vbLf & Channel.Name & vbLf & "" & vbLf & Target
End Sub
Public Sub VoiceUser(Channel As clsChannel, Target As String, User As String, Optional VoiceAnyway As Boolean = False, Optional SendLink As Boolean = True)
On Error Resume Next
Dim Chan As String
Chan = Channel.Name
If Channel.IsVoice(Target) And VoiceAnyway = False Then Exit Sub
Channel.Voices.Add Target, Target
If Channel.IsNorm(Target) Then Channel.NormUsers.Remove Target
'1 = Command, 2 = Nick, 3 = Channel, 4 = Target
If SendLink Then SendLinks "VoiceUser" & vbLf & User & vbLf & Channel.Name & vbLf & "" & vbLf & Target
End Sub
Public Sub DeVoiceUser(Channel As clsChannel, Target As String, User As String, Optional SendLink As Boolean = True)
On Error Resume Next
Dim Chan As String
If Not Channel.IsVoice(Target) Then Exit Sub
Chan = Channel.Name
Channel.Voices.Remove Target
If Channel.IsOp(Target) Then
Else
    Channel.NormUsers.Add Target, Target
End If
'1 = Command, 2 = Nick, 3 = Channel, 4 = Target
If SendLink Then SendLinks "DeVoiceUser" & vbLf & User & vbLf & Channel.Name & vbLf & "" & vbLf & Target
End Sub

Public Sub BanUser(Channel As clsChannel, Target As String, User As String, Optional SendLink As Boolean = True)
Dim i As Long
If Not Channel.IsBanned2(Target) Then
    Channel.Bans.Add Target, Target
End If
If Channel.IsMode("r") Then
    Dim dab As New clsDatabase
    With dab
        .FileName = App.Path & "\Channels\" & Channel.Name & ".dat"
        .WriteEntry "Bans", "Count", Channel.Bans.Count
        For i = 1 To Channel.Bans.Count
            .WriteEntry "Ban " & CStr(i), "Mask", Channel.Bans(i)
        Next i
    End With
End If
'1 = Command, 2 = Nick, 3 = Channel, 4 = Target
If SendLink Then SendLinks "BanUser" & vbLf & User & vbLf & Channel.Name & vbLf & "" & vbLf & Target
End Sub

Public Sub UnBanUser(Channel As clsChannel, Target As String, User As String, Optional SendLink As Boolean = True)
Dim i As Long
Channel.Bans.Remove Target
If Channel.IsMode("r") Then
    Dim dab As New clsDatabase
    With dab
        .FileName = App.Path & "\Channels\" & Channel.Name & ".dat"
        .WriteEntry "Bans", "Count", Channel.Bans.Count
        For i = 1 To Channel.Bans.Count
            .WriteEntry "Ban " & CStr(i), "Mask", Channel.Bans(i)
        Next i
    End With
End If
'1 = Command, 2 = Nick, 3 = Channel, 4 = Target
If SendLink Then SendLinks "UnBanUser" & vbLf & User & vbLf & Channel.Name & vbLf & "" & vbLf & Target
End Sub

Public Sub RemoveChanModes(NewModes As String, Chan As String, User As clsUser, Optional SendLink As Boolean = True)
Dim Found As Boolean, X As Long, Channel As clsChannel, Modes As String
Set Channel = ChanToObject(Chan)
If Not (Mid(NewModes, 1, 1) = "k" Or Mid(NewModes, 1, 1) = "l") Then
    For X = 1 To Len(NewModes)
        If Channel.IsMode(Mid(NewModes, X, 1)) Then
            Channel.Modes.Remove Mid(NewModes, X, 1)
            Modes = Modes & Mid(NewModes, X, 1)
        End If
    Next X
    If Modes = "" Then Exit Sub
    Dim i As Long
ElseIf Mid(NewModes, 1, 2) = "lk" Then
    Dim Key As String, Limit As String
    Limit = Replace(NewModes, "lk ", "")
    If Channel.Key = Limit Then Channel.Key = ""
    Channel.Limit = 0
ElseIf Mid(NewModes, 1, 1) = "k" Then
    If Channel.Key = Replace(NewModes, "k ", "") Then
        Channel.Key = ""
    End If
ElseIf Mid(NewModes, 1, 1) = "l" Then
    Channel.Limit = 0
End If
If Channel.IsMode("r") Then
    Dim dab As New clsDatabase
    With dab
        .FileName = App.Path & "\Channels\" & Channel.Name & ".dat"
        .WriteEntry "General", "Modes", Channel.GetModesForFile
        .WriteEntry "General", "Key", Channel.Key
        .WriteEntry "General", "Limit", Channel.Limit
    End With
End If
'1 = Command, 2 = Nick, 3 = +/-, 4 = Modes, 5 = Channel
If SendLink Then SendLinks "ChanMode" & vbLf & User.Nick & vbLf & "-" & vbLf & Modes & vbLf & Channel.Name
End Sub

Public Sub AddChanModes(NewModes As String, Chan As String, User As clsUser, Optional SendLink As Boolean = True)
Dim Found As Boolean, X As Long, Channel As clsChannel, Modes As String
On Error Resume Next
Set Channel = ChanToObject(Chan)
If Not (Mid(NewModes, 1, 1) = "k" Or Mid(NewModes, 1, 1) = "l") Then
    For X = 1 To Len(NewModes)
        If Not Channel.IsMode(Mid(NewModes, X, 1)) Then
            Channel.Modes.Add Mid(NewModes, X, 1), Mid(NewModes, X, 1)
            Modes = Modes & Mid(NewModes, X, 1)
        End If
    Next X
    If Modes = "" Then Exit Sub
    Dim i As Long
ElseIf Mid(NewModes, 1, 2) = "lk" Then
    Dim Key As String, Limit As String
    Limit = Replace(NewModes, "lk ", "")
    Key = Mid(Limit, 1, InStr(1, Limit, " ") - 1)
    Limit = Replace(Limit & " ", Key, "")
    Limit = Trim(Limit)
    Channel.Key = Limit
    Channel.Limit = Key
ElseIf Mid(NewModes, 1, 1) = "k" Then
    Channel.Key = Replace(NewModes, "k ", "")
ElseIf Mid(NewModes, 1, 1) = "l" Then
    Channel.Limit = Replace(NewModes, "l ", "")
ElseIf Mid(NewModes, 1, 1) = "C" Then
    Channel.Limit = Replace(NewModes, "C ", "")
ElseIf Mid(NewModes, 1, 1) = "c" Then
    Channel.Limit = Replace(NewModes, "c ", "")
End If
If Channel.IsMode("r") Then
    Dim dab As New clsDatabase
    With dab
        .FileName = App.Path & "\Channels\" & Channel.Name & ".dat"
        .WriteEntry "General", "Modes", Channel.GetModesForFile
        .WriteEntry "General", "Key", Channel.Key
        .WriteEntry "General", "Limit", Channel.Limit
    End With
End If
'1 = Command, 2 = Nick, 3 = +/-, 4 = Modes, 5 = Channel
If SendLink Then SendLinks "ChanMode" & vbLf & User.Nick & vbLf & "+" & vbLf & Modes & vbLf & Channel.Name
End Sub

Public Function GetChanList(User As String)
Dim i As Long, Chan As clsChannel
For i = 1 To UBound(Channels)
    If Not Channels(i) Is Nothing Then
        GetChanList = GetChanList & ":" & ServerName & " 322 " & User & " " & Channels(i).Name & " " & Channels(i).All.Count & " :[+" & Channels(i).GetModes & "] " & Channels(i).Topic & vbCrLf
    End If
Next i
End Function

Public Sub InviteUser(Channel As clsChannel, Target As String, User As String, Optional SendLink As Boolean = True)
Dim i As Long
If Not Channel.IsInvited3(Replace(Target, "*!", "")) Then
    Channel.Invites.Add Replace(Target, "*!", ""), Replace(Target, "*!", "")
End If
If Channel.IsMode("r") Then
    Dim dab As New clsDatabase
    With dab
        .FileName = App.Path & "\Channels\" & Channel.Name & ".dat"
        .WriteEntry "Invites", "Count", Channel.Invites.Count
        For i = 1 To Channel.Invites.Count
            .WriteEntry "Invite " & CStr(i), "Mask", Channel.Invites(i)
        Next i
    End With
End If
'1 = Command, 2 = Nick, 3 = Channel, 4 = Target
If SendLink Then SendLinks "InviteUser" & vbLf & User & vbLf & Channel.Name & vbLf & "" & vbLf & Target
End Sub

Public Sub UnInviteUser(Channel As clsChannel, Target As String, User As String, Optional SendLink As Boolean = True)
Dim i As Long
If Channel.IsInvited3(Replace(Target, "*!", "")) Then
    Channel.Invites.Remove Replace(Target, "*!", "")
End If
If Channel.IsMode("r") Then
    Dim dab As New clsDatabase
    With dab
        .FileName = App.Path & "\Channels\" & Channel.Name & ".dat"
        .WriteEntry "Invites", "Count", Channel.Invites.Count
        For i = 1 To Channel.Invites.Count
            .WriteEntry "Invite " & CStr(i), "Mask", Channel.Invites(i)
        Next i
    End With
End If
'1 = Command, 2 = Nick, 3 = Channel, 4 = Target
If SendLink Then SendLinks "UnInviteUser" & vbLf & User & vbLf & Channel.Name & vbLf & "" & vbLf & Target
End Sub

Public Sub ExceptionUser(Channel As clsChannel, Target As String, User As String, Optional SendLink As Boolean = True)
Dim i As Long
If Not Channel.IsException2(Replace(Target, "*!", "")) Then
    Channel.Exceptions.Add Replace(Target, "*!", ""), Replace(Target, "*!", "")
End If
If Channel.IsMode("r") Then
    Dim dab As New clsDatabase
    With dab
        .FileName = App.Path & "\Channels\" & Channel.Name & ".dat"
        .WriteEntry "Exceptions", "Count", Channel.Exceptions.Count
        For i = 1 To Channel.Exceptions.Count
            .WriteEntry "Exception " & CStr(i), "Mask", Channel.Exceptions(i)
        Next i
    End With
End If
'1 = Command, 2 = Nick, 3 = Channel, 4 = Target
If SendLink Then SendLinks "ExceptUser" & vbLf & User & vbLf & Channel.Name & vbLf & "" & vbLf & Target
End Sub

Public Sub UnExceptionUser(Channel As clsChannel, Target As String, User As String, Optional SendLink As Boolean = True)
Dim i As Long
If Channel.IsException2(Replace(Target, "*!", "")) Then
    Channel.Exceptions.Remove Replace(Target, "*!", "")
End If
If Channel.IsMode("r") Then
    Dim dab As New clsDatabase
    With dab
        .FileName = App.Path & "\Channels\" & Channel.Name & ".dat"
        .WriteEntry "Exceptions", "Count", Channel.Exceptions.Count
        For i = 1 To Channel.Exceptions.Count
            .WriteEntry "Exception " & CStr(i), "Mask", Channel.Exceptions(i)
        Next i
    End With
End If
'1 = Command, 2 = Nick, 3 = Channel, 4 = Target
If SendLink Then SendLinks "UnExceptUser" & vbLf & User & vbLf & Channel.Name & vbLf & "" & vbLf & Target
End Sub

Public Sub AddUserMode(Index As Long, Modes As String, Optional Silent As Boolean = False, Optional SendLink As Boolean = True)
Dim NewModes As String
Modes = LCase(Modes)
Dim i As Long
For i = 1 To Len(Modes)
    Select Case Mid(Modes, i, 1)
        Case "s"
            If Not Users(Index).IsMode("s") Then
                NewModes = NewModes & "s"
                Users(Index).AddModes "s"
            End If
        Case "w"
            If Not Users(Index).IsMode("w") Then
                NewModes = NewModes & "w"
                Users(Index).AddModes "w"
            End If
    End Select
Next i
'1 = Command, 2 = Nick, 3 = +/-, 4 = Modes
If SendLink Then SendLinks "ModeUser" & vbLf & Users(Index).Nick & vbLf & "+" & vbLf & Modes
End Sub

Public Sub RemoveUsermode(Index As Long, Modes As String, Optional Silent As Boolean = False, Optional SendLink As Boolean = True)
Dim NewModes As String
Modes = LCase(Modes)
Dim i As Long
For i = 1 To Len(Modes)
    Select Case Mid(Modes, i, 1)
        Case "s"
            If Users(Index).IsMode("s") Then
                NewModes = NewModes & "s"
                Users(Index).Modes.Remove "s"
            End If
        Case "w"
            If Users(Index).IsMode("w") Then
                NewModes = NewModes & "w"
                Users(Index).Modes.Remove "w"
            End If
        Case "a"
            If Users(Index).IsMode("a") Then
                NewModes = NewModes & "a"
                Users(Index).Modes.Remove "a"
                Users(Index).Away = False
                Users(Index).AwayMsg = ""
            End If
        Case "o"
            If Users(Index).IRCOp Then
                Users(Index).MsgsSent = 0
                SendLinks "ServerMsg" & vbLf & Users(Index).Nick & " gave up his Operator status", True, ServerName
                NewModes = NewModes & "o"
                Users(Index).Modes.Remove "o"
                Users(Index).IRCOp = False
                Operators = Operators - 1
                Users(Index).DNS = Users(Index).RealDNS
                Users(Index).RealDNS = ""
                SendNotice "", "You are not an operator anymore", ServerName, , CInt(Index)
            End If
            Users(Index).IRCOp = False
    End Select
Next i
'1 = Command, 2 = Nick, 3 = +/-, 4 = Modes
If SendLink Then SendLinks "ModeUser" & vbLf & Users(Index).Nick & vbLf & "-" & vbLf & Modes
End Sub
