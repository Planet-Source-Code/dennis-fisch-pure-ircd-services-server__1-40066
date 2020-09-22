Attribute VB_Name = "modNickServ"
 Option Explicit
Private DB As New clsDatabase

Public Sub ParseNSCmd(Cmd As String, Index As Long)
Cmd = "ns " & Cmd
On Error GoTo parseerr
Dim Msg As String, CMDStr As String, lcmd As Integer, arg1 As String, arg2 As String, cmd2 As String
Dim User As clsUser
Set User = Users(Index)
Msg = Replace(Cmd, "NS ", "", , , vbTextCompare)
If Not InStr(1, Msg, " ") <> 0 Then
    CMDStr = Msg
Else
    CMDStr = (Mid(Msg, 1, InStr(1, Msg, " ") - 1))
End If
Msg = Replace(Msg, CMDStr & " ", "")
Select Case LCase(CMDStr)
    Case "identify"
        lcmd = 1
        If IdentifyUser(User.Nick, Msg, User) Then
            SendNotice User.Nick, "Password accepted - you have been identified", "NickServ"
            Dim i As Long, memocount As Long
            For i = 1 To Memos.Count
                If Memos(i).Target = User.Nick And Not Memos(i).Read Then
                    memocount = memocount + 1
                End If
            Next i
            If memocount > 0 Then SendNotice User.Nick, "You have " & memocount & " new Memo(s), type '/msg MemoServ list'", "MemoServ"
            User.Identified = True
            
            User.NR = False
        Else
            If IsRegistered(User.Nick) Then
                SendNotice User.Nick, "Invalid Password", "NickServ"
            Else
                SendNotice User.Nick, User.Nick & " is not a registered nickname", "NickServ"
            End If
        End If
    Case "drop"
        lcmd = 2
        arg1 = Mid(Cmd, InStr(1, Cmd, " ") + 1)
        If DropUser(User.Nick, arg1) Then
            SendNotice User.Nick, "NickName " & User.Nick & " has been dropped!", "NickServ"
            SendLinks "ServerMsg" & vbLf & User.Nick & "!" & User.Ident & "@" & User.DNS & " dropped his nick"
        Else
            SendNotice User.Nick, "Invalid Password or Username not found!", "NickServ"
        End If
    Case "register"
        lcmd = 3
        If FS.GetFolder(App.Path & "\Users").Files.Count = MaxNickRegs Then
            SendNotice "", "Global Nicknames Limit reached, you cannot register another Nickname", ServerName, , User.Index
            SendNotice "", "Please consult the Server Admin", ServerName, , User.Index
            SendLinks "ServerMsg" & vbLf & "Nickname limit reached, no new nicknames can be registered"
            Exit Sub
        End If
        Cmd = Replace(Cmd, "ns register ", "", , , vbTextCompare)
        arg1 = Mid(Cmd, InStr(1, Cmd, " ") + 1)
        arg2 = Replace(Cmd, " " & arg1, "")
        If RegisterUser(User.Nick, arg2, arg1) Then
            SendNotice User.Nick, "NickName " & User.Nick & " has been registered!", "NickServ"
            SendNotice User.Nick, "Password is: " & arg2, "NickServ"
            SendLinks "ServerMsg" & vbLf & User.Nick & "!" & User.Ident & "@" & User.DNS & " registered " & User.Nick
        Else
            SendNotice User.Nick, "NickName " & User.Nick & " is alrerady registered!", "NickServ"
        End If
    Case "kill"
        lcmd = 4
        Cmd = Replace(Cmd, "ns kill ", "", , , vbTextCompare)
        arg1 = Mid(Cmd, InStr(1, Cmd, " ") + 1)
        arg2 = Replace(Cmd, " " & arg1, "")
        Dim KN As clsUser
        Set KN = NickToObject(arg2)
        If KillUser(arg2, arg1) And Not KN Is Nothing Then
            SendLinks "KillUser" & vbLf & KN.Nick & vbLf & "NickServ (Kill requested by " & User.Nick & ")"
            SendNotice User.Nick, "NickName " & arg2 & " has been killed as requested!", "NickServ"
            Set Users(KN.Index) = Nothing
        Else
            SendNotice User.Nick, "Invalid Password or User not Found!", "NickServ"
        End If
    Case "info"
        lcmd = 5
        Dim NI As String
        NI = GetNickInfo(Msg)
        If NI <> "" Then
            SendNotice User.Nick, "" & NI, "NickServ"
        Else
            SendNotice User.Nick, "NickName not registered", "NickServ"
        End If
    Case "changeinfo"
        lcmd = 6
        Msg = Mid(Msg, 1, InStr(1, Msg, " ") - 1)
        arg1 = Replace(Cmd, "ns changeinfo " & Msg & " ", "", , , vbTextCompare)
        NI = ChangeNickInfo(User.Nick, arg1, Msg)
        If NI <> "" Then
            SendNotice User.Nick, NI, "NickServ"
        Else
            SendNotice User.Nick, "NickName not registered!", "NickServ"
            SendNotice User.Nick, "ChangeInfo (changeinfo [newpass] [newemail] )", "NickServ"
        End If
    Case "help"
        If Msg = "" Or Msg = "help" Then
            SendNotice User.Nick, "NickServ Commands", "NickServ"
            SendNotice User.Nick, "Identify (identify [password] )", "NickServ"
            SendNotice User.Nick, "Drop (drop [password] )", "NickServ"
            SendNotice User.Nick, "Register (register [password] [email] )", "NickServ"
            SendNotice User.Nick, "Kill (kill [nick] [password] )", "NickServ"
            SendNotice User.Nick, "Info (info [nick] )", "NickServ"
            SendNotice User.Nick, "ChangeInfo (changeinfo [newpass] [newemail] )", "NickServ"
        Else
            Select Case LCase(Msg)
                Case "identify"
                    SendNotice User.Nick, "IDENITFY (identify [password] )", "NickServ"
                Case "drop"
                    SendNotice User.Nick, "DROP (drop [password] )", "NickServ"
                Case "register"
                    SendNotice User.Nick, "REGISTER (register [password] [email] )", "NickServ"
                Case "kill"
                    SendNotice User.Nick, "KILL (kill [nick] [password] )", "NickServ"
                Case "info"
                    SendNotice User.Nick, "INFO (info [nick] )", "NickServ"
                Case "changeinfo"
                    SendNotice User.Nick, "CHANGEINFO (changeinfo [newpass] [newemail] )", "NickServ"
            End Select
        End If
    Case Else
        SendNotice User.Nick, "Command Unknown", "NickServ"
End Select
Exit Sub
parseerr:
Select Case lcmd
    Case 1
        SendNotice User.Nick, "Identify (identify [password] )", "NickServ"
    Case 2
        SendNotice User.Nick, "Drop (drop [password] )", "NickServ"
    Case 3
        SendNotice User.Nick, "Register (register [password] [email] )", "NickServ"
    Case 4
        SendNotice User.Nick, "Kill (kill [nick] [password] )", "NickServ"
    Case 5
        SendNotice User.Nick, "Info (info [nick] )", "NickServ"
    Case 6
        SendNotice User.Nick, "ChangeInfo (changeinfo [newpass] [newemail] )", "NickServ"
    Case Else
        SendNotice User.Nick, "Unknown Command or missing parameters", "NickServ"
End Select
End Sub

Public Function RegisterUser(Nick As String, Password As String, Email As String) As Boolean
If FS.FileExists(App.Path & "\Users\" & Nick & ".usr") Then
    RegisterUser = False
Else
    DB.FileName = App.Path & "\Users\" & Nick & ".usr"
    DB.WriteEntry Nick, "Email", Email
    DB.WriteEntry Nick, "Password", Password
    RegisterUser = True
End If
End Function

Public Function DropUser(Nick As String, Password As String) As Boolean
If Not FS.FileExists(App.Path & "\Users\" & Nick & ".usr") Then
    DropUser = False
Else
    DB.FileName = App.Path & "\Users\" & Nick & ".usr"
    If UCase(DB.ReadEntry(Nick, "Password", Password)) = UCase(Password) Then
        FS.DeleteFile App.Path & "\Users\" & Nick & ".usr"
        DropUser = True
    End If
End If
End Function

Public Function KillUser(Nick As String, Password As String) As Boolean
If Not FS.FileExists(App.Path & "\Users\" & Nick & ".usr") Then
    KillUser = False
Else
    DB.FileName = App.Path & "\Users\" & Nick & ".usr"
    If UCase(DB.ReadEntry(Nick, "Password", Password)) = UCase(Password) Then
        KillUser = True
    End If
End If
End Function

Public Function IdentifyUser(Nick As String, Password As String, User As clsUser) As Boolean
If InStr(1, Password, " ") Then
    Nick = Mid(Password, 1, InStr(1, Password, " ") - 1)
    Password = Replace(Password, Nick & " ", "")
End If
If Not FS.FileExists(App.Path & "\Users\" & Nick & ".usr") Then
    IdentifyUser = False
Else
    DB.FileName = App.Path & "\Users\" & Nick & ".usr"
    If UCase(DB.ReadEntry(Nick, "Password", Password)) = UCase(Password) Then
        IdentifyUser = True
        User.IdentifiedAs = Nick
    End If
End If
End Function

Public Function GetNickInfo(Nick As String) As String
If Not FS.FileExists(App.Path & "\Users\" & Nick & ".usr") Then
    GetNickInfo = ""
Else
    DB.FileName = App.Path & "\Users\" & Nick & ".usr"
    GetNickInfo = "Email: " & DB.ReadEntry(Nick, "Email", "")
End If
End Function

Public Function ChangeNickInfo(Nick As String, Email As String, Password As String) As String
If Not FS.FileExists(App.Path & "\Users\" & Nick & ".usr") Then
    ChangeNickInfo = ""
Else
    DB.FileName = App.Path & "\Users\" & Nick & ".usr"
    DB.WriteEntry Nick, "Email", Email
    DB.WriteEntry Nick, "Password", Password
    ChangeNickInfo = "New Password: " & Password & "         New Email: " & Email
End If
End Function

Public Function IsRegistered(Nick As String) As Boolean
IsRegistered = FS.FileExists(App.Path & "\Users\" & Nick & ".usr")
End Function
