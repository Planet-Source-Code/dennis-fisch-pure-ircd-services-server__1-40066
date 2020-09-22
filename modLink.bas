Attribute VB_Name = "modLink"
Option Explicit
Option Base 1

Public Sub SendChanList(Index As Long)
Dim i As Long, x As Long
Dim ChanName As String
Dim ChanModes As String
Dim Limit As String
Dim Key As String
Dim Users As String
For i = 1 To UBound(Channels)
    If Not Channels(i) Is Nothing Then
        ChanName = Channels(i).Name & "°"
        For x = 1 To Channels(i).Modes.Count
            ChanModes = ChanModes + Channels(i).Modes(x)
        Next x
        ChanModes = ChanModes & "°"
        Limit = Channels(i).Limit & "°"
        Key = Channels(i).Key & "°"
        For x = 1 To Channels(i).All.Count
            Users = Users & Channels(i).All(x) & "|"
        Next x
        Users = Users & "°"
        SendLink "LICH" & ChanName & ChanModes & Limit & Key & Users, Index
    End If
Next i
End Sub

Public Sub SendNickList(Index As Long)
Dim i As Long, x As Long
Dim NickName As String
Dim NickMode As String
Dim Name As String
Dim DNS As String
For i = 1 To UBound(Users)
    If Not Users(i) Is Nothing Then
        NickName = Users(i).Nick & "°"
        For x = 1 To Users(i).Modes.Count
            NickMode = NickMode & Users(i).Modes(x)
        Next x
        NickMode = NickMode & "°"
        Name = Users(i).Name & "°"
        DNS = Users(i).DNS & "°"
        SendLink "LINI" & NickName & NickMode & Name & DNS, Index
    End If
Next i
End Sub

Public Sub ParseNickList(sNickList As String)
Dim i As Long, x As Long, y As Long
Dim NickList() As String, User() As String, NU As clsUser
'1 = Nick
'2 = Modes
'3 = Name
'4 = DNS
NickList = Split(sNickList, "LINI")
For i = LBound(NickList) To UBound(NickList)
    User = Split(NickList(i), "°")
    Set NU = GetFreeSlot
    NU.Nick = User(1)
    For x = 1 To Len(User(2))
        NU.Modes.Add Mid(User(2), x, 1)
        If Mid(User(2), x, 1) = "o" Then NU.IRCOp = True: SendSvrMsg NU.Nick & " is now operator"
    Next x
    NU.Name = User(3)
    NU.DNS = User(4)
Next i
End Sub

Public Sub ParseChanList(sChanList As String)
Dim i As Long, x As Long, y As Long
Dim ChanList() As String, chan() As String, Channel As clsChannel
'1 = Chan Name
'2 = Modes
'3 = Limit
'4 = Key
Dim ChanName As String
Dim ChanModes As String
Dim Limit As Long
Dim Key As String
Dim Users As String
ChanList = Split(sChanList, "LICH")
For i = LBound(ChanList) To UBound(ChanList)
    chan = Split(ChanList(i), "°")
    Set Channel = GetFreeChan
    Channel.Name = chan(1)
    Channel.AddModes chan(2)
    If Not IsNumeric(chan(3)) Then chan(3) = 0
    Channel.Limit = chan(3)
    If Limit <> 0 Then AddChanModes "l " & Limit, Channel.Name, modMain.Users(1)
    Channel.Key = chan(4)
    If Key <> "" Then AddChanModes "k " & Key, Channel.Name, modMain.Users(1)
Next i
End Sub

Public Sub NewUser(UserInfo As String)
Dim User As clsUser, UI() As String, x As Long
'1 = Nick
'2 = Modes
'3 = Name
'4 = DNS
Set User = GetFreeSlot
UI = Split(UserInfo, "°")
User.Nick = UI(1)
For x = 1 To Len(UI(2))
    User.Modes.Add Mid(UI(2), x, 1)
Next x
User.Name = UI(3)
User.DNS = UI(4)
End Sub

Public Sub RemUser(QuitLine As String)
Dim QLine() As String
'1 = Nick
'2 = Reason
QLine = Split(QuitLine, "°")
SendQuit NickToObject(QLine(1)).Index, QLine(2), , False
Set NickToObject(QLine(1)) = Nothing
'With modIRCUserCommands.
End Sub

Public Sub GlobalMsg(GlobalLine As String)
Dim GLine() As String, i As Long
'1 = Message
'2 = Sender
GLine = Split(GlobalLine, "°")
For i = 1 To UBound(Users)
    If Not Users(i) Is Nothing Then SendNotice Users(i).Nick, GLine(1), GLine(2), , , False
Next i
End Sub

Public Sub ChanMode(ModeLine As String)
Dim MLine() As String
'1 = Mode/Topic
'2 = Nick
'3 = Channel
'4 = Topic/Modes
'5 (only for mode) = +/-
MLine = Split(ModeLine, "°")
If MLine(1) = "TOPIC" Then
    SetTopic MLine(3), MLine(4), MLine(2)
ElseIf MLine(1) = "MODE" Then
    Select Case MLine(5)
        Case "+"
            AddChanModes MLine(4), MLine(3), NickToObject(MLine(2)), False
        Case "-"
            RemoveChanModes MLine(4), MLine(3), NickToObject(MLine(2)), False
    End Select
End If
End Sub

Public Sub NickMode(ModeLine As String)
Dim MLine() As String
'1 = Nick
'2 = Mode
'3 = +/-
MLine = Split(ModeLine, "°")
Select Case MLine(3)
    Case "+"
        AddUserMode NickToObject(MLine(1)).Index, MLine(2), , False
    Case "-"
        RemoveUsermode NickToObject(MLine(1)).Index, MLine(2), , False
End Select
End Sub

Public Sub Msg(MessageLine As String)
Dim MLine() As String
'1 = Nick/Chan
'2 = Target
'3 = Message
'4 = Sender
MLine = Split(MessageLine, "°")
Select Case MLine(1)
    Case "NICK"
        SendMsg MLine(2), MLine(3), MLine(4), False, False
    Case "CHAN"
        SendMsg MLine(2), MLine(3), MLine(4), True, False
End Select
End Sub

Public Sub Notice(NoticeLine As String)
Dim NLine() As String
'1 = Nick/Chan
'2 = Target
'3 = Message
'4 = Sender
NLine = Split(NoticeLine, "°")
Select Case NLine(1)
    Case "NICK"
        SendNotice NLine(2), NLine(3), NLine(4), False, , False
    Case "CHAN"
        SendNotice NLine(2), NLine(3), NLine(4), True, , False
End Select
End Sub

Public Sub PartChan(PartLine As String)
Dim PLine() As String
'1 = User
'2 = Channel
PLine = Split(PartLine, "°")
SendPart NickToObject(PLine(1)).Index, PLine(2), "", False
End Sub

Public Sub JoinChan(JoinLine As String)
Dim JLine() As String
'1 = Channel to Join
'2 = Channel Key
'3 = Nick
Dim chan As String
Dim ck As String
Dim Index As Long
JLine = Split(JoinLine, "°")
chan = JLine(1)
ck = JLine(2)
Index = NickToObject(JLine(3)).Index
If Not Users(Index).IsOnChan(chan) Then
    If Not ChanExists(chan) Then
        Dim NewChannel As clsChannel
        Set NewChannel = GetFreeChan
        NewChannel.Name = chan
        NewChannel.Modes.Add "t", "t"
        NewChannel.Modes.Add "n", "n"
        NewChannel.Topic = DefTopic
        NewChannel.Ops.Add Users(Index).Nick, Users(Index).Nick
        NewChannel.All.Add Users(Index).Nick, Users(Index).Nick
        Users(Index).Onchannels.Add chan, chan
        SendWsock Index, ":" & Users(Index).Nick & " JOIN " & chan, True
        SendWsock Index, ":" & ServerName & " 353 " & Users(Index).Nick & " = " & chan & " :" & Replace(NewChannel.GetOps & " " & NewChannel.GetVoices & " " & NewChannel.GetNorms, "  ", " "), True
        SendWsock Index, ":" & ServerName & " 366 " & Users(Index).Nick & " " & chan & " :End of /NAMES list.", True
    Else
        On Local Error Resume Next
        Dim JoinChan As clsChannel
        Set JoinChan = ChanToObject(chan)
        If (Not JoinChan.Key = "") Then
            If Not JoinChan.Key = ck And Not Users(Index).IRCOp And (Not Users(Index).IsOwner(JoinChan.Name)) Then
                SendWsock Index, ":" & ServerName & " 475 " & Users(Index).Nick & " " & chan & " :Cannot join channel (+b)"
                Exit Sub
            End If
        End If
        If (JoinChan.All.Count >= JoinChan.Limit And JoinChan.Limit <> 0) And Not Users(Index).IRCOp And (Not Users(Index).IsOwner(JoinChan.Name)) Then
            SendWsock Index, ":" & ServerName & " 471 " & Users(Index).Nick & " " & chan & " :Cannot join channel (+l)"
            Exit Sub
        End If
        If JoinChan.IsBanned(Users(Index)) And (Users(Index).IRCOp = False) And (JoinChan.IsException(Users(Index)) = False) And (Not Users(Index).IsOwner(JoinChan.Name)) Then
            SendWsock Index, ":" & ServerName & " 474 " & Users(Index).Nick & " " & chan & " :Cannot join channel (+b)"
            Exit Sub
        End If
        If JoinChan.IsMode("i") And (Users(Index).IRCOp = False) And (JoinChan.IsInvited2(Users(Index)) = False) And (JoinChan.IsInvited(Users(Index).Nick) = False) And (Not Users(Index).IsOwner(JoinChan.Name)) Then
            SendWsock Index, ":" & ServerName & " 473 " & Users(Index).Nick & " " & chan & " :Cannot join channel (+i)"
            Exit Sub
        End If
        If Not Users(Index).IRCOp And (JoinChan.ULOp(Users(Index).Nick) = False) And (JoinChan.ULVoice(Users(Index).Nick) = False) And (Not Users(Index).IsOwner(JoinChan.Name)) Then
            JoinChan.NormUsers.Add Users(Index).Nick, Users(Index).Nick
        ElseIf JoinChan.ULVoice(Users(Index).Nick) Then
            JoinChan.Voices.Add Users(Index).Nick, Users(Index).Nick
        Else
            JoinChan.Ops.Add Users(Index).Nick, Users(Index).Nick
        End If
        JoinChan.All.Add Users(Index).Nick, Users(Index).Nick
        Users(Index).Onchannels.Add chan, chan
        SendWsock Index, ":" & Users(Index).Nick & " JOIN " & chan, True
        SendWsock Index, ":" & ServerName & " 353 " & Users(Index).Nick & " = " & chan & " :" & Trim(Replace(JoinChan.GetOps & " " & JoinChan.GetVoices & " " & JoinChan.GetNorms, "  ", " ")), True
        SendWsock Index, ":" & ServerName & " 366 " & Users(Index).Nick & " " & chan & " :End of /NAMES list.", True
        SendWsock Index, ":" & ServerName & " 332 " & Users(Index).Nick & " " & chan & " :" & JoinChan.Topic, True
        SendWsock Index, ":" & ServerName & " 333 " & JoinChan.TopicSetBy & " " & chan & " " & JoinChan.TopicSetBy & " " & JoinChan.TopicSetOn, True
        NotifyJoin CLng(Index), chan, False
        If Users(Index).IRCOp Or JoinChan.ULOp(Users(Index).Nick) Or (Users(Index).IsOwner(JoinChan.Name)) Then
            OpUser JoinChan, Users(Index).Nick, "ChanServ", True
        ElseIf JoinChan.ULVoice(Users(Index).Nick) Then
            VoiceUser JoinChan, Users(Index).Nick, "ChanServ", True
        End If
    End If
End If
End Sub

Public Sub Renick(NickLine As String)
Dim NLine() As String
'1 = previous Nick
'2 = New nick
NLine = Split(NickLine, "°")
ChangeNick NickToObject(NLine(1)).Index, NLine(2), False
End Sub

Public Function ServerToIndex(Server As String) As Long
Dim i As Long
For i = 1 To Links.Count
    If UCase(Server) = Links(i) Then
        ServerToIndex = i
        Exit Function
    End If
Next i
End Function

Public Sub SendLink(Message As String, Optional ToCertainServer As Long)
Dim i As Long
If ToCertainServer <> 0 Then
    ServerTraffic = ServerTraffic + Len(Message)
    frmMain.wsockLink(ToCertainServer).SendData Message & vbLf
    Exit Sub
End If
For i = 0 To frmMain.wsockLink.UBound
    ServerTraffic = ServerTraffic + Len(Message)
    On Local Error Resume Next
    If Not i = 0 Then frmMain.wsockLink(i).SendData Message & vbLf
Next i
End Sub
