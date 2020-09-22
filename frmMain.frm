VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pure-IRCd Services"
   ClientHeight    =   210
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   2865
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   210
   ScaleWidth      =   2865
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin PURE_IRCd_Services.fbTrayIcon fbTrayIcon1 
      Height          =   1155
      Left            =   2460
      TabIndex        =   0
      Top             =   1740
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   2037
   End
   Begin VB.Timer tmrLinkPing 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   65535
      Left            =   540
      Top             =   1380
   End
   Begin MSWinsockLib.Winsock Link 
      Index           =   0
      Left            =   120
      Top             =   1380
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "dennis"
      RemotePort      =   6669
      LocalPort       =   6668
   End
   Begin VB.Timer tmrNS 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   2000
      Left            =   120
      Top             =   1800
   End
   Begin VB.Menu mnuTray 
      Caption         =   "mnuTray"
      Visible         =   0   'False
      Begin VB.Menu mnuTrayStartServer 
         Caption         =   "(Re)Start Server"
      End
      Begin VB.Menu mnuTrayCloseServer 
         Caption         =   "Close Server"
      End
      Begin VB.Menu mnuTrayShow 
         Caption         =   "Show/Hide"
      End
      Begin VB.Menu mnuTrayExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1

Private Sub fbTrayIcon1_MouseClick(ByVal FBButton As EnumFBButtonConstants)
Select Case FBButton
    Case &H203
        frmMain.Visible = Not frmMain.Visible
    Case &H205
        frmMain.PopupMenu mnuTray, , , , mnuTrayShow
End Select
End Sub

Private Sub Form_Load()
Dim i As Long, FS As New FileSystemObject
ReDim Users(4)
Rehash
Link(0).LocalPort = LinkPort
Link(0).Listen
For i = LBound(Users) To UBound(Users): Set Users(i) = Nothing: Next i
For i = LBound(Channels) To UBound(Channels): Set Channels(i) = Nothing: Next i
fbTrayIcon1.AddTrayIcon App.Path & "\Tray.ico", "Pure-IRCd Services"
Started = Now
If FS.FolderExists(App.Path & "\Users") = False Then FS.CreateFolder App.Path & "\Users"
If FS.FolderExists(App.Path & "\Channels") = False Then FS.CreateFolder App.Path & "\Channels"
If FS.FolderExists(App.Path & "\Memos") = False Then FS.CreateFolder App.Path & "\Memos"
CreateServices
LoadChans
LoadMemos
If LogLevel > 3 Then WriteHeader
MaxGlobalUsers = 0
CurGlobalUsers = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
If LogLevel <> 0 Then WriteFooter
fbTrayIcon1.RemoveTrayIcon
End
End Sub

Private Sub Link_Close(Index As Integer)
SendLinks "DeadLink" & vbLf & ServerName & vbLf & Link(Index).Tag
If LogLevel = 1 Or LogLevel = 3 Then
    If LogFormat = 0 Then
        LogText "[LINK]<" & Now & " (LINK CLOSED " & ServerName & " -- " & Link(Index).Tag & ")> "
    Else
        LogHTML ServerName, "LINK CLOSED " & ServerName & " -- " & Link(Index).Tag
    End If
End If
Dim i As Long
For i = 5 To UBound(Users)
    DoEvents
    If Not Users(i) Is Nothing Then
        If Users(i).IsOnLink(Link(Index).Tag) Then
            SendQuit i, ServerName & " -- " & Link(Index).Tag
            Set Users(i) = Nothing
        End If
    End If
Next i
CurLinkCount = CurLinkCount - 1
Unload Link(Index)
On Local Error Resume Next
Unload tmrLinkPing(Index)
End Sub

Private Sub Link_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim LinkCount As Long
LinkCount = Link.Count + 1
CurLinkCount = CurLinkCount + 1
MaxLinkCount = MaxLinkCount + 1
Index = LinkCount
Load Link(LinkCount)
Link(LinkCount).Close
Link(LinkCount).LocalPort = 30000 + LinkCount
Link(LinkCount).Accept requestID
Wait 100
Dim i As Long, X As Long
For i = 1 To UBound(Users)
    If Not Users(i) Is Nothing Then
        SendLinks "NewUser" & vbLf & Users(i).Nick & vbLf & Users(i).Name & vbLf & Users(i).DNS & vbLf & Users(i).Ident & vbLf & Users(i).Server & vbLf & Users(i).ServerDescritption & vbLf & Users(i).SignOn & vbLf & Users(i).GID & vbLf & Users(i).GetModes & vbLf & ServerName & " ", , Index
        For X = 1 To Users(i).Onchannels.Count
            SendLinks "JoinChan" & vbLf & Users(i).Nick & vbLf & Users(i).Onchannels(X), , Index '1 = Command, 2 = Nick, 3 = Channel
        Next X
    End If
Next i
SendLinks "Info" & vbLf & ServerName
For i = 1 To UBound(Channels)
    If Not Channels(i) Is Nothing Then
        For X = 1 To Channels(i).Ops.Count
            On Local Error Resume Next
            DeOpUser Channels(i), Channels(i).Ops(X), "ChanServ", True
        Next X
    End If
Next i
Load tmrLinkPing(Index)
tmrLinkPing(Index).Enabled = True
tmrLinkPing(Index).Tag = 1
End Sub

Private Sub Link_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
Dim strData As String, strcmd() As String, cmdArray() As String, i As Long, User As clsUser, DontSendLink As Boolean, X As Long, NewUser As clsUser, strRoute() As String
Link(Index).GetData strData, 8
If LogLevel = 1 Then
    If LogFormat = 0 Then
        LogText "[LINK]<" & Now & " (INCOMING " & Link(Index).Tag & ")> " & strData
    Else
        LogHTML Link(Index).Tag & "(Link) INCOMING", strData
    End If
End If
cmdArray = Split(strData, vbCrLf)
For X = LBound(cmdArray) To UBound(cmdArray)
    If cmdArray(X) = "" Then GoTo NextCmd
    strcmd = Split(cmdArray(X), vbLf)
    Select Case strcmd(0)
        Case "Info"
            Link(Index).Tag = strcmd(1)
            If LogLevel = 1 Or LogLevel = 3 Then
                If LogFormat = 0 Then
                    LogText "[LINK]<" & Now & " (LINKED " & ServerName & " -- " & strcmd(1) & ")> " & strData
                Else
                    LogHTML ServerName, "LINKED " & ServerName & " -- " & strcmd(1)
                End If
            End If
        Case "NewUser"
            If strcmd(8) = "" Then GoTo NextCmd
            '1 = Command, 2 = Nick, 3 = Name, 4 = DNS, 5 = Ident, 6 = Server, 7 = ServerDescription, 8 = SignOn
            Set NewUser = NickToObject(strcmd(1))
            If NewUser Is Nothing Then
                Set User = GetFreeSlot
                UserCount = UserCount - 1
                User.DNS = strcmd(3)
                User.Nick = strcmd(1)
                User.Name = strcmd(2)
                User.Ident = strcmd(4)
                User.Server = strcmd(5)
                User.ServerDescritption = strcmd(6)
                User.SignOn = strcmd(7)
                User.GID = strcmd(8)
                If InStr(1, strcmd(10), " ") = 0 Then strcmd(10) = strcmd(10) & " "
                User.Route = strcmd(10)
                User.Hops = CountSpaces(strcmd(10)) - 1
                User.AddModes strcmd(9)
            Else
                If strcmd(8) = NewUser.GID Then GoTo NextCmd
                If NewUser.SignOn < strcmd(7) Then
                    SendLinks "KillUser" & vbLf & strcmd(1) & vbLf & "Nick Collision, other nick signed on earlier"
                Else
                    Set User = GetFreeSlot
                    UserCount = UserCount - 1
                    User.DNS = strcmd(3)
                    User.Nick = strcmd(1)
                    User.Name = strcmd(2)
                    User.Ident = strcmd(4)
                    User.Server = strcmd(5)
                    User.ServerDescritption = strcmd(6)
                    User.SignOn = strcmd(7)
                    User.GID = strcmd(8)
                    If InStr(1, strcmd(10), " ") = 0 Then strcmd(10) = strcmd(10) & " "
                    User.Route = strcmd(10)
                    User.Hops = CountSpaces(strcmd(10)) - 1
                    User.AddModes strcmd(9)
                End If
            End If
            SendLinks cmdArray(X) & " " & ServerName & " ", CLng(Index)
            DontSendLink = False
        Case "QuitUser"
            '1 = Command, 2 = Nick, 3 = QuitMsg
            Set User = NickToObject(strcmd(1))
            SendQuit User.Index, strcmd(2), , False
            Set Users(User.Index) = Nothing
        Case "KillUser"
            '1 = Command, 2 = Nick, 3 = Reason
            Set User = NickToObject(strcmd(1))
            SendQuit User.Index, strcmd(2), True, False
            Set Users(User.Index) = Nothing
        Case "JoinChan"
            '1 = Command, 2 = Nick, 3 = Channel
            Dim NewChannel As clsChannel
            Set User = NickToObject(strcmd(1))
            If Not ChanExists(strcmd(2)) Then
                Set NewChannel = GetFreeChan
                NewChannel.Name = strcmd(2)
                NewChannel.Modes.Add "t", "t"
                NewChannel.Modes.Add "n", "n"
                NewChannel.Topic = DefTopic
                NewChannel.Ops.Add User.Nick, User.Nick
                NewChannel.All.Add User.Nick, User.Nick
                Users(Index).Onchannels.Add strcmd(2), strcmd(2)
            Else
                Set NewChannel = ChanToObject(strcmd(2))
                NewChannel.All.Add User.Nick, User.Nick
                If User.IRCOp Or NewChannel.ULOp(User.Nick) Or (User.IsOwner(NewChannel.Name)) Then
                    OpUser NewChannel, User.Nick, "ChanServ", True
                ElseIf NewChannel.ULVoice(User.Nick) Then
                    VoiceUser NewChannel, User.Nick, "ChanServ", True
                Else
                    NewChannel.NormUsers.Add User.Nick, User.Nick
                End If
                If NewChannel.All.Count = 1 Then
                    SendLinks "JoinChan" & vbLf & "ChanServ" & vbLf & NewChannel.Name
                    SendLinks "OpUser" & vbLf & "ChanServ" & vbLf & NewChannel.Name & vbLf & "" & vbLf & "ChanServ"
                    SendLinks "ChanMode" & vbLf & "ChanServ" & vbLf & "+" & vbLf & NewChannel.GetModesForFile & vbLf & strcmd(2)
                    If NewChannel.Key <> "" Then SendLinks "Key" & vbLf & "ChanServ" & vbLf & NewChannel.Name & vbLf & NewChannel.Key
                    If NewChannel.Limit <> 0 Then SendLinks "Limit" & vbLf & "ChanServ" & vbLf & NewChannel.Name & vbLf & NewChannel.Limit
                    Dim Y As Long
                    SendLinks "SetTopic" & vbLf & "ChanServ" & vbLf & strcmd(2) & vbLf & "" & vbLf & NewChannel.Topic
                    For Y = 1 To NewChannel.Bans.Count
                        SendLinks "BanUser" & vbLf & "ChanServ" & vbLf & NewChannel.Name & vbLf & "" & vbLf & NewChannel.Bans(Y)
                    Next Y
                    For Y = 1 To NewChannel.Invites.Count
                        SendLinks "InviteUser" & vbLf & "ChanServ" & vbLf & NewChannel.Name & vbLf & "" & vbLf & NewChannel.Invites(Y)
                    Next Y
                    For Y = 1 To NewChannel.Exceptions.Count
                        SendLinks "ExceptUser" & vbLf & "ChanServ" & vbLf & NewChannel.Name & vbLf & "" & vbLf & NewChannel.Exceptions(Y)
                    Next Y
                End If
            End If
            NotifyJoin User.Index, strcmd(2), False
            User.Onchannels.Add strcmd(2), strcmd(2)
        Case "PartUser"
            '1 = Command, 2 = Nick, 3 = Channel, 4 = Reason
            Set User = NickToObject(strcmd(1))
            SendPart User.Index, strcmd(2), strcmd(3), False
        Case "ModeUser"
            '1 = Command, 2 = Nick, 3 = +/-, 4 = Modes
            Set User = NickToObject(strcmd(1))
            Select Case strcmd(2)
                Case "+"
                    AddUserMode User.Index, strcmd(3), , False
                Case "-"
                    RemoveUsermode User.Index, strcmd(3), , False
            End Select
        Case "ChanMode"
            '1 = Command, 2 = Nick, 3 = +/-, 4 = Modes, 5 = Channel
            Set User = NickToObject(strcmd(1))
            Select Case strcmd(2)
                Case "+"
                    AddChanModes strcmd(3), strcmd(4), User, False
                Case "-"
                    RemoveChanModes strcmd(3), strcmd(4), User, False
            End Select
        Case "KickUser"
            '1 = Command, 2 = Nick, 3 = Channel, 4 = Reason, 5 = Target
            Set User = NickToObject(strcmd(1))
            If strcmd(3) = "" Then strcmd(3) = strcmd(1)
            KickUser User.Nick, strcmd(2), strcmd(4), strcmd(3), True, False
        Case "KLine"
            '1 = Command, 2 = Mask
            Klines.Add strcmd(1), strcmd(1)
        Case "ServerMsg"
            '1 = Command, 2 = Msg
        Case "Global"
            '1 = Command, 2 = Msg
            For i = LBound(Users) To UBound(Users)
                If Not Users(i) Is Nothing Then SendNotice "", "*** Global -- " & strcmd(1), ServerName, , CInt(i), False
            Next i
        Case "PrivMsgChan"
            '1 = Command, 2 = Nick, 3 = Channel, 4 = Msg
            Dim Msg As String, Target As String
            Msg = strcmd(3)
            If Mid(Msg, 1, 1) = "." Then
                Set NewChannel = ChanToObject(strcmd(2))
                If NewChannel.IsMode("r") Then
                    Set User = NickToObject(strcmd(1))
                    If (NewChannel.IsOp(User.Nick) Or User.IsMode("o")) Then
                        Dim cmd As String
                        cmd = Msg
                        If InStr(1, cmd, " ") <> 0 Then cmd = Mid(Msg, 1, InStr(1, Msg, " ") - 1)
                        Target = Replace(Msg, cmd & " ", "")
                        Select Case cmd
                            Case ".op"
                                If Not NewChannel.IsOnChan(Target) Then GoTo NextCmd
                                OpUser NewChannel, Target, "ChanServ"
                            Case ".deop"
                                If Not NewChannel.IsOnChan(Target) Then GoTo NextCmd
                                If Target = "ChanServ" Then GoTo NextCmd
                                DeOpUser NewChannel, Target, "ChanServ"
                            Case ".voice"
                                If Not NewChannel.IsOnChan(Target) Then GoTo NextCmd
                                VoiceUser NewChannel, Target, "ChanServ"
                            Case ".devoice"
                                If Not NewChannel.IsOnChan(Target) Then GoTo NextCmd
                                DeVoiceUser NewChannel, Target, "ChanServ"
                            Case ".kick"
                                If Not NewChannel.IsOnChan(Target) Then GoTo NextCmd
                                If Target = "ChanServ" Then GoTo NextCmd
                                KickUser "ChanServ", NewChannel.Name, Target, "Kick Requested by " & User.Nick, True
                            Case ".ban"
                                BanUser NewChannel, Target, "ChanServ"
                            Case ".invite"
                                InviteUser NewChannel, Target, "ChanServ"
                            Case ".uninvite"
                                UnInviteUser NewChannel, Target, "ChanServ"
                            Case ".except"
                                ExceptionUser NewChannel, Target, "ChanServ"
                            Case ".unexcept"
                                UnExceptionUser NewChannel, Target, "ChanServ"
                        End Select
                    Else
                        SendNotice User.Nick, "You're not Channel Operator", "ChanServ"
                    End If
                End If
            End If
        Case "PrivMsgUser"
            '1 = Command, 2 = Nick, 3 = Target, 4 = Msg
            Select Case strcmd(2)
                Case "ChanServ"
                    ParseCSCmd strcmd(3), NickToObject(strcmd(1)).Index
                Case "NickServ"
                    ParseNSCmd strcmd(3), NickToObject(strcmd(1)).Index
                Case "MemoServ"
                    ParseMSCmd strcmd(3), NickToObject(strcmd(1)).Index
            End Select
        Case "NoticeUser"
            '1 = Command, 2 = Nick, 3 = Target, 4 = Msg
        Case "NoticeChan"
            '1 = Command, 2 = Nick, 3 = Channel, 4 = Msg
        Case "Nick"
            '1 = Command, 2 = Nick, 3 = NewNick
            Set User = NickToObject(strcmd(1))
            ChangeNick User.Index, strcmd(2), False
        Case "OpUser"
            '1 = Command, 2 = Nick, 3 = Channel, 4 = Target
            OpUser ChanToObject(strcmd(2)), strcmd(4), strcmd(1), True, False
        Case "DeOpUser"
            '1 = Command, 2 = Nick, 4 = Channel, 4 = Target
            DeOpUser ChanToObject(strcmd(2)), strcmd(4), strcmd(1), False
        Case "VoiceUser"
            '1 = Command, 2 = Nick, 4 = Channel, 4 = Target
            VoiceUser ChanToObject(strcmd(2)), strcmd(4), strcmd(1), True, False
        Case "DeVoiceUser"
            '1 = Command, 2 = Nick, 4 = Channel, 4 = Target
            DeVoiceUser ChanToObject(strcmd(2)), strcmd(4), strcmd(1), False
        Case "BanUser"
            '1 = Command, 2 = Nick, 4 = Channel, 4 = Mask
            BanUser ChanToObject(strcmd(2)), strcmd(4), strcmd(1), False
        Case "UnBanUser"
            '1 = Command, 2 = Nick, 4 = Channel, 4 = Mask
            UnBanUser ChanToObject(strcmd(2)), strcmd(4), strcmd(1), False
        Case "ExceptUser"
            '1 = Command, 2 = Nick, 4 = Channel, 4 = Mask
            ExceptionUser ChanToObject(strcmd(2)), strcmd(4), strcmd(1), False
        Case "UnExceptUser"
            '1 = Command, 2 = Nick, 4 = Channel, 4 = Mask
            UnExceptionUser ChanToObject(strcmd(2)), strcmd(4), strcmd(1), False
        Case "InviteUser"
            '1 = Command, 2 = Nick, 4 = Channel, 4 = Mask
            InviteUser ChanToObject(strcmd(2)), strcmd(4), strcmd(1), False
        Case "UnInviteUser"
            '1 = Command, 2 = Nick, 4 = Channel, 4 = Mask
            UnInviteUser ChanToObject(strcmd(2)), strcmd(4), strcmd(1), False
        Case "Limit"
            '1 = Command, 2 = Nick, 4 = Channel, 4 = Mask
            AddChanModes "l " & strcmd(3), strcmd(2), NickToObject(strcmd(1)), False
        Case "Key"
            '1 = Command, 2 = Nick, 4 = Channel, 4 = Mask
            AddChanModes "k " & strcmd(3), strcmd(2), NickToObject(strcmd(1)), False
        Case "UnLimit"
            '1 = Command, 2 = Nick, 4 = Channel, 4 = Mask
            RemoveChanModes "l", strcmd(2), NickToObject(strcmd(1)), False
        Case "UnKey"
            '1 = Command, 2 = Nick, 4 = Channel, 4 = Mask
            RemoveChanModes "k " & strcmd(3), strcmd(2), NickToObject(strcmd(1)), False
        Case "AddInvite"
            '1 = Command, 2 = Nick, 4 = Channel, 4 = Target
            ChanToObject(strcmd(2)).Invited.Add strcmd(4), strcmd(4)
        Case "SetTopic"
            '1 = Command, 2 = Nick, 4 = Channel, 4 = Target
            SetTopic strcmd(2), strcmd(4), strcmd(1), False
        Case "DeadLink"
            '1 = Command, 2 = Server1, 3 = Server2
            For i = 5 To UBound(Users)
                If Users(i).IsOnLink(strcmd(2)) Then
                    SendQuit i, strcmd(1) & " -- " & strcmd(2)
                    Set Users(i) = Nothing
                End If
            Next i
            SendLinks cmdArray(X), CLng(Index)
            DontSendLink = False
        Case "PING"
            SendLinks "PONG" & vbLf, , Index
            DontSendLink = True
        Case "PONG"
            tmrLinkPing(Index).Tag = 1
            DontSendLink = True
        Case Else
            DontSendLink = True
    End Select
    If Not DontSendLink Then SendLinks cmdArray(X), CLng(Index)
    DontSendLink = False
NextCmd:
Next X
End Sub

Private Sub Link_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If LogLevel = 1 Or LogLevel = 3 Then
    If LogFormat = 0 Then
        LogText "[LINK]<" & Now & " (LINK CLOSED " & ServerName & " -- " & Link(Index).Tag & ")> "
    Else
        LogHTML ServerName, "LINK CLOSED " & ServerName & " -- " & Link(Index).Tag
    End If
End If
Link_Close (Index)
End Sub

Private Sub mnuTrayExit_Click()
Unload Me
End Sub

Private Sub mnuTrayShow_Click()
frmMain.Visible = Not frmMain.Visible
End Sub

Private Sub tmrLinkPing_Timer(Index As Integer)
If Not CLng(tmrLinkPing(Index).Tag) = 1 Then
    SendQuit CLng(Index), "Ping Timeout"
    Link_Close (Index)
    Exit Sub
End If
tmrLinkPing(Index).Tag = 0
Link(Index).SendData "PING" & vbLf
End Sub

Private Sub tmrNS_Timer(Index As Integer)
On Error Resume Next
If tmrNS(Index).Interval = 60 And ((Not Users(tmrNS(Index).Tag).Identified = False) And IsRegistered(Users(tmrNS(Index).Tag).Nick)) Then
    SendNotice Users(Index).Nick, "This nickname does not belong to you.", "NickServ"
    ChangeNick CLng(Index), "Guest" & GetRand
    Unload tmrNS(Index)
End If
If Users(tmrNS(Index).Tag).NR Then
    SendNotice Users(tmrNS(Index).Tag).Nick, "You have 60 seconds to identify or change nicknames.", "NickServ"
    SendNotice Users(tmrNS(Index).Tag).Nick, "To authenticate your identity: /msg NickServ identify [password]", "NickServ"
    tmrNS(Index).Interval = 60
End If
Unload tmrNS(Index)
End Sub
