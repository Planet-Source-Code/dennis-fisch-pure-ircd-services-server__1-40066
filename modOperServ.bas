Attribute VB_Name = "modOperServ"
 Option Explicit
Option Compare Text
Private DB As New clsDatabase

Public Sub ParseOSCmd(Cmd As String, Index As Long)
If Not Users(Index).IRCOp And Not (LCase(Cmd) = "os stats") Then
    SendNotice "", "Permission Denied- You're not an IRC operator", "OperServ", , CInt(Index)
    Exit Sub
End If
On Error GoTo parseerr
Dim Msg As String, CMDStr As String, lcmd As Integer, arg1 As String, arg2 As String, cmd2 As String
Dim User As clsUser
Set User = Users(Index)
Msg = Replace(Cmd, "OS ", "")
If Not InStr(1, Msg, " ") <> 0 Then
    CMDStr = Msg
Else
    CMDStr = (Mid(Msg, 1, InStr(1, Msg, " ") - 1))
End If
Msg = Replace(Msg, CMDStr & " ", "")
Select Case LCase(CMDStr)
    Case "stats"
        Dim i As Long
        lcmd = 1
        SendNotice "", "STATISTICS FOR " & ServerName, "OperServ", , CInt(Index)
        SendNotice "", "Links", "OperServ", , CInt(Index)
        For i = frmMain.Link.LBound To frmMain.Link.UBound
            On Local Error Resume Next
            If frmMain.Link(i).Tag <> "" Then SendNotice "", "Link " & (i - 1) & "     " & ServerName & " -- " & frmMain.Link(i).Tag, "OperServ", , CInt(Index)
        Next i
        SendNotice "", "---------------------------------------------------------------", "OperServ", , CInt(Index)
        SendNotice "", "Servertraffic: " & FN(ServerTraffic, 3, ".") & " bytes", "OperServ", , CInt(Index)
        SendNotice "", "Up Since: " & Started, "OperServ", , CInt(Index)
    Case "kill"
        lcmd = 4
        Set User = NickToObject(Msg)
        If Not User Is Nothing Then
            SendLinks "ServerMsg" & vbLf & "Recieved KILL message for " & User.Nick & "!" & User.Ident & "@" & User.DNS & " Path: " & Users(Index).Nick & " (OperServ)"
            SendQuit User.Index, "Killed by " & Users(Index).Nick & " (OperServ)", True
            Set Users(User.Index) = Nothing
'            frmMain.wsock_Close (User.Index)
        End If
    Case "akill"
        lcmd = 5
        Set User = NickToObject(Msg)
        If Not User Is Nothing Then
            SendLinks "ServerMsg" & vbLf & "Recieved AKILL message for " & User.Nick & "!" & User.Ident & "@" & User.DNS & " Path: " & Users(Index).Nick & " (OperServ)"
            On Local Error Resume Next
            Klines.Add User.DNS
            SendQuit User.Index, "AKilled by " & Users(Index).Nick & " (OperServ)", True
            Set Users(User.Index) = Nothing
        End If
    Case "clear"
        lcmd = 6
        modChanServ.ParseCSCmd "CS clear " & Msg, Index
    Case "global"
        lcmd = 7
        For i = LBound(Users) To UBound(Users)
            If Not Users(i) Is Nothing Then SendNotice "", "*** Global -- " & Msg, "OperServ", , CInt(i)
        Next i
    Case "logonnews"
        lcmd = 8
        FS.CreateTextFile(App.Path & "\logon.conf", True).WriteLine Msg
    Case "help"
        If Msg = "" Or Msg = "help" Then
            SendNotice "", "OperServ Commands", "OperServ", , CInt(Index)
            SendNotice "", "STATS (stats)", "OperServ", , CInt(Index)
            SendNotice "", "KILL (kill [nick] [reason] )", "OperServ", , CInt(Index)
            SendNotice "", "AKILL (akill [nick] )", "OperServ", , CInt(Index)
            SendNotice "", "CLEAR (clear [channel] )", "OperServ", , CInt(Index)
            SendNotice "", "GLOBAL (global [message] )", "OperServ", , CInt(Index)
            SendNotice "", "LOGONNEWS (logonnews [news] )", "OperServ", , CInt(Index)
        Else
            Select Case LCase(Msg)
                Case "stats"
                    SendNotice User.Nick, "Stats (stats)", "OperServ", , CInt(Index)
                Case "kill"
                    SendNotice User.Nick, "Kill (kill [nick] <Reason> )", "OperServ", , CInt(Index)
                Case "akill"
                    SendNotice User.Nick, "Akill (akill [nick] <Reason> )", "OperServ", , CInt(Index)
                Case "clear"
                    SendNotice User.Nick, "Clear (clear [channel] )", "OperServ", , CInt(Index)
                Case "global"
                    SendNotice User.Nick, "Global (global [message] )", "OperServ", , CInt(Index)
                Case "logonnews"
                    SendNotice User.Nick, "LogonNews (logonnews [news] )", "OperServ", , CInt(Index)
            End Select
        End If
    Case Else
        SendNotice User.Nick, "Command Unknown", "OperServ"
End Select
Exit Sub
parseerr:
Select Case lcmd
    Case 1
        SendNotice User.Nick, "Identify (identify [password] )", "OperServ"
    Case 2
        SendNotice User.Nick, "Drop (drop [password] )", "OperServ"
    Case 3
        SendNotice User.Nick, "Register (register [password] [email] )", "OperServ"
    Case 4
        SendNotice User.Nick, "Kill (kill [nick] [password] )", "OperServ"
    Case 5
        SendNotice User.Nick, "Info (info [nick] )", "OperServ"
    Case 6
        SendNotice User.Nick, "ChangeInfo (changeinfo [newpass] [newemail] )", "OperServ"
    Case Else
        SendNotice User.Nick, "Unknown Command or missing parameters", "OperServ"
End Select
End Sub

Public Sub SendLogonNews(Nick)
SendNotice "", "***Logon News -- " & FS.OpenTextFile(App.Path & "\logon.conf").ReadLine, "OperServ", , CInt(Nick)
End Sub
