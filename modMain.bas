Attribute VB_Name = "modMain"
Option Explicit
Option Base 1
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetTickCount& Lib "kernel32" ()
Private Declare Function CoCreateGuid Lib "ole32" (ID As Any) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Global ChanCount As Long
Global UserCount As Long
Global MaxUser As Long
Global MaxChans As Long
Global MaximumUsers As Long
Global MaxNickRegs As Long
Global MaxChanRegs As Long
Public CurLinkCount As Long
Public MaxLinkCount As Long
Global MaxChunkSize As Long
Global Users() As clsUser
Global Channels() As clsChannel
Global Memos As New MemoCol
Global Olines(100) As Oline
Global FS As New FileSystemObject
Global DB As New clsDatabase
Public ServerName As String
Public Started As Date
Public Klines As New Collection
Public CloneControl As New Collection
Public ServerTraffic As Double
Public OverAllMax As Long
Public DefTopic As String
Public DefUserModes As String
Public DefQuit As String
Public MaxChannels As String
Public ServerDesc As String
Public AdminName As String
Public AdminEmail As String
Public SessionLimit As Long
Public Nicklen As Integer
Public MaxJoinChannels As Integer
Public TopicLen As Integer
Public KickLen As Integer
Public Msglen As Integer
Public AwayLen As Integer
Public Operators As Integer
Public LinkPort As Long
Public LogFile As String
Public LogLevel As Integer
Public LogFormat As Integer
Public LogStatusHandle As Long
Dim StatusInterval As Long
Dim StatusFile As String

Public CurGlobalUsers As Long
Public MaxGlobalUsers As Long

Public Type Oline
    UserName As String
    Password As String
    Mask As String
    InUse As Boolean
End Type

Public Function UnixTime() As Long
UnixTime = DateDiff("s", DateValue("1/1/1970"), Now)
End Function

Public Function ChanToObject(ChanName As String) As clsChannel
On Error Resume Next
Dim i As Long
Set ChanToObject = Nothing
For i = 1 To UBound(Channels)
    If Not Channels(i) Is Nothing Then
        If UCase(ChanName) = UCase(Channels(i).Name) Then
            Set ChanToObject = Channels(i)
            Exit Function
        End If
    End If
Next i
End Function

Public Function NickToObject(NickName As String, Optional StartAt As Long = 1, Optional LocalsOnly As Boolean = False) As clsUser
On Error Resume Next
Dim i As Long, UB As Long
UB = UBound(Users)
For i = 1 To UB
    If Not Users(i) Is Nothing Then
        'If ((Users(i).LocalUser = True) And (LocalsOnly)) Then
        If UCase(NickName) = UCase(Users(i).Nick) Then
            Set NickToObject = Users(i)
            Exit Function
        End If
        'End If
    End If
Next i
End Function

Public Function GetFreeSlot() As clsUser
Dim i As Long
If Not UBound(Users) >= MaximumUsers Then
    ReDim Preserve Users(UBound(Users) + 1)
    For i = 1 To UBound(Users)
        If (Users(i) Is Nothing) Then
            Set Users(i) = New clsUser
            Users(i).Index = i
            Set GetFreeSlot = Users(i)
            Exit Function
        End If
    Next i
End If
Set GetFreeSlot = Nothing
End Function

Public Function GetRandom() As Long
Randomize
Dim MyValue As Long, i As Long, r As Long
For i = 1 To 8
    MyValue = Int((9 * Rnd) + 0)
    r = CLng(CStr(r) & CStr(MyValue))
Next i
GetRandom = r
End Function

Public Function ChanExists(ChannelName As String) As Boolean
If Not ChanToObject(ChannelName) Is Nothing Then ChanExists = True
End Function

Public Function GetFreeChan() As clsChannel
Dim i As Long
For i = 1 To UBound(Channels)
    If (Channels(i) Is Nothing) Then
        Set Channels(i) = New clsChannel
        Channels(i).Index = i
        Set GetFreeChan = Channels(i)
        Exit Function
    End If
Next i
Set GetFreeChan = Nothing
End Function

Public Function CountSpaces(strCount As String) As Long
Dim i As Long
For i = 1 To Len(strCount)
    If (Mid(strCount, i, 1) = " ") Then CountSpaces = CountSpaces + 1
Next i
CountSpaces = CountSpaces + 1
End Function

Public Sub ParseModeNicks(Nicks As String, ByRef Nickarr() As String)
If InStr(1, Nicks, " ") <> 0 Then
    Nickarr = Split(Nicks, " ")
Else
    ReDim Nickarr(1)
    Nickarr(1) = Nicks
End If
End Sub

Public Function GetRand() As Long
Randomize
Dim MyValue As Long, i As Long, r As Long
For i = 1 To 4
    MyValue = Int((9 * Rnd) + 0)
    r = CLng(CStr(r) & CStr(MyValue))
Next i
GetRand = r
End Function

Public Function IsKlined(IP As String) As String
Dim i As Long
For i = 1 To Klines.Count
    DoEvents
    If IP Like Klines(i) Then
        IsKlined = Klines(i)
        Exit Function
    End If
Next i
End Function

Public Sub LoadChans()
On Error Resume Next
Dim dab As New clsDatabase, ChanFile As File
Dim Chan As clsChannel
For Each ChanFile In FS.GetFolder(App.Path & "\Channels").Files
    dab.FileName = ChanFile.Path
    If ChanToObject(dab.ReadEntry("General", "Name", "")) Is Nothing Then
        Set Chan = GetFreeChan
        Chan.Name = dab.ReadEntry("General", "Name", "")
        Chan.Topic = dab.ReadEntry("General", "Topic", "")
        Chan.TopicSetOn = CLng(dab.ReadEntry("General", "TopicSetOn", CStr(UnixTime)))
        Chan.TopicSetBy = dab.ReadEntry("General", "TopicSetBy", "")
        Chan.AddModes dab.ReadEntry("General", "Modes", "r")
        Chan.Password = dab.ReadEntry("General", "Password", "")
        Chan.Key = dab.ReadEntry("General", "Key", "")
        Chan.Limit = dab.ReadEntry("General", "Limit", "0")
        Dim i As Long
        For i = 1 To dab.ReadEntry("UserLevels", "Count", "0")
            Chan.AddToUserList dab.ReadEntry("User " & CStr(i), "Nickname", ""), CLng(dab.ReadEntry("User " & CStr(i), "Level", "0"))
        Next i
        For i = 1 To dab.ReadEntry("Bans", "Count", "0")
            Chan.Bans.Add dab.ReadEntry("Ban " & CStr(i), "Mask", ""), dab.ReadEntry("Ban " & CStr(i), "Mask", "")
        Next i
        For i = 1 To dab.ReadEntry("Exceptions", "Count", "0")
            Chan.Exceptions.Add dab.ReadEntry("Exception " & CStr(i), "Mask", ""), dab.ReadEntry("Exception " & CStr(i), "Mask", "")
        Next i
        For i = 1 To dab.ReadEntry("Invites", "Count", "0")
            Chan.Invites.Add dab.ReadEntry("Invite " & CStr(i), "Mask", ""), dab.ReadEntry("Invite " & CStr(i), "Mask", "")
        Next i
    End If
Next
End Sub

Public Sub LogText(LogStr As String)
If LogLevel = 0 Then Exit Sub
FS.OpenTextFile(LogFile, ForAppending, True).WriteLine LogStr
End Sub

Public Sub LogHTML(Originator As String, LogStr)
If LogLevel = 0 Then Exit Sub
With FS.OpenTextFile(LogFile, ForAppending, True)
    .WriteLine "<tr>"
    .WriteLine "<td width=20%>" & Now & "</td>"
    .WriteLine "<td width=20%>" & Originator & "</td>"
    .WriteLine "<td width=60%>" & LogStr & "</td>"
    .WriteLine "</tr>"
End With
End Sub

Public Function SizeString(strData As String, Size As Long) As String
If Size <= Len(strData) Then
    SizeString = strData
    Exit Function
End If
strData = strData & Space(Size - Len(strData))
SizeString = strData
End Function

Public Sub LoadMemos()
On Error Resume Next
Dim MemoFile As File
For Each MemoFile In FS.GetFolder(App.Path & "\Memos").Files
    With MemoFile.OpenAsTextStream
        Memos.Add .ReadLine, .ReadLine, .ReadLine
        Memos(Memos.Count).Read = False
        Memos(Memos.Count).MemoID = MemoFile.Name
    End With
Next
End Sub

Public Sub Rehash(Optional Nick As String = "Dill.mine.nu")
Dim DB As New clsDatabase, i As Long, Kline As String
DB.FileName = App.Path & "\.conf"
'Server Settings
'Servername
ServerName = DB.ReadEntry("General Settings", "Servername", "dill.mine.nu")
'Server Description, appears on /Whois
ServerDesc = DB.ReadEntry("General Settings", "Description", "KillTheDill")
'Maximum Amount of Clients before Server is "full".
MaximumUsers = DB.ReadEntry("General Settings", "MaxUsers", "100") + 4
'Maximum Amount of Nickname Registrations.
MaxNickRegs = DB.ReadEntry("General Settings", "MaxNickRegs", "100")
'Maximum Amount of Channel Registrations.
MaxChanRegs = DB.ReadEntry("General Settings", "MaxChanRegs", "100")
'Maximum Amount of Channels that can exist on Server.
MaxChannels = DB.ReadEntry("General Settings", "MaxChannels", "100")
'Maximum Amount of Connections accepted from one IP.
SessionLimit = DB.ReadEntry("General Settings", "Session Limit", "3")
'Maximum length of Clients Nickname
Nicklen = DB.ReadEntry("General Settings", "MaxNickLength", "25")
'Maximum of Channels a Client can join.
MaxJoinChannels = DB.ReadEntry("General Settings", "MaxJoinChannels", "7")
'maximum Topic Length.
TopicLen = DB.ReadEntry("General Settings", "TopicLen", "128")
'maximum length of Kick Reason.
KickLen = DB.ReadEntry("General Settings", "KickLen", "64")
'maximum length of notice and privmsg messages.
Msglen = DB.ReadEntry("General Settings", "MsgLen", "512")
'maximum amount of data sent in one packet.
MaxChunkSize = DB.ReadEntry("General Settings", "MaxChunkSize", "512")
'The port PURE will accept incoming Link connection request on.
LinkPort = DB.ReadEntry("General Settings", "LinkPort", "8000")
'Log Level, 0 = none, 1 = debug,2 = ALL Client Traffic, 3 = Only Important Status Messages (recommended)
LogLevel = DB.ReadEntry("General Settings", "LogLevel", "3")
'LogFile location
LogFile = DB.ReadEntry("General Settings", "LogFilename", "pure.log")
If InStr(1, LogFile, "\") = 0 Then
    ChDir App.Path
    LogFile = CurDir & "\" & LogFile
End If
Open LogFile For Output As #1: Close #1
StatusFile = FS.GetFile(LogFile).ParentFolder.Path & "\status.htm"
If LogLevel = 0 Then FS.DeleteFile (LogFile)
'LogFormat, 0 = Text, 1 = HTML/PHP
LogFormat = DB.ReadEntry("General Settings", "LogFormat", "0")
'StatusInterval, seconds until a new status file is generated.
StatusInterval = DB.ReadEntry("General Settings", "StatusInterval", "0")
KillTimer frmMain.hWnd, 0
If Not StatusInterval = 0 Then
    LogStatusHandle = SetTimer(frmMain.hWnd, 0, (StatusInterval * 1000), AddressOf LogStatus)
    If LogStatusHandle = 0 Then LogHTML ServerName, "Unable to start StatusTimer"
End If
ReDim Preserve Channels(MaxChannels)
'Admin
AdminName = DB.ReadEntry("Admin", "Name", "")
AdminEmail = DB.ReadEntry("Admin", "Email", "")
'Channel Defaults
DefTopic = DB.ReadEntry("Channel Defaults", "Topic", "Unregistered Channel")
'Standard Usermodes during log in
DefUserModes = DB.ReadEntry("Default User Settings", "UserModes", "w")
DefQuit = DB.ReadEntry("Default User Settings", "Default Quit Msg", "Fox Dilligents IRCd")
For i = 1 To DB.ReadEntry("K-lines", "Count", "0")
    Kline = DB.ReadEntry("K-lines", CStr(i), "")
    Klines.Add Kline, Kline
Next i
'Olines
Dim OLineCount As Long
OLineCount = DB.ReadEntry("O-Lines", "Count", "0")
For i = 1 To OLineCount
    'ReDim Preserve Olines(i)
    Olines(i).UserName = DB.ReadEntry("O-Line " & i, "UserName", "")
    Olines(i).Password = DB.ReadEntry("O-Line " & i, "Password", "")
    Olines(i).Mask = DB.ReadEntry("O-Line " & i, "Mask", "")
    Olines(i).InUse = True
Next i
End Sub

Public Function Wait(ByVal TimeToWait As Long)
Dim EndTime As Long
EndTime = GetTickCount + TimeToWait
Do Until GetTickCount > EndTime
    Sleep 10
    DoEvents
Loop
End Function

Public Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

Public Sub LogStatus(ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
Dim i As Long
With FS.CreateTextFile(StatusFile, True)
    .WriteLine "<body text=#00FF00 bgcolor=#000000>"
    .WriteLine "<p align=center><b>PURE IRCd LOG FILE</b></p>"
    .WriteLine "<p>Current Local Users:<i><b> " & UserCount - 4 & "</b></i><br>"
    .WriteLine "Max Local Users:<i><b> " & MaxUser & "</b></i></p>"
    .WriteLine "<p>Current Global Users:<i><b> " & CurGlobalUsers & "</b></i><br>"
    .WriteLine "Max Global Users:<i><b> " & MaxGlobalUsers & "</b></i></p>"
    .WriteLine "<p>Current Links:<i><b> " & CurLinkCount & "</b></i></br>"
    .WriteLine "Max Links:<i><b> " & MaxLinkCount & "</b></i></p>"
    .WriteLine "<hr>"
    For i = frmMain.Link.LBound To frmMain.Link.UBound
        On Local Error Resume Next
        If frmMain.Link(i).Tag <> "" Then .WriteLine "<p>Link " & (i - 1) & ":<i><b> " & ServerName & " -- " & frmMain.Link(i).Tag & "</b></i></p>"
    Next i
    .WriteLine "<hr>"
    .WriteLine "<p>Current Channels:<i><b> " & ChanCount & "</b></i><br>"
    .WriteLine "Max Channels:<i><b> " & MaxChans & "</b></i></p>"
    .WriteLine "<p>Traffic:<i><b> " & FN(ServerTraffic, 3, ".") & " bytes" & "</b></i></p>"
    .WriteLine "Server port:<i><b> " & LinkPort & "</b></i></p>"
End With
End Sub

Public Function FN(Number, Optional MaxGroupLength As Long = 3, Optional Delimeter As String = ".")
On Error Resume Next
Dim i As Long, Num As String
Num = StrReverse(Number)
For i = 1 To Len(Num) Step MaxGroupLength
    FN = FN & Mid(Num, i, MaxGroupLength) & Delimeter
Next i
FN = StrReverse(Mid(FN, 1, Len(FN) - 1))
End Function

Public Sub CreateServices()
Dim GFS As clsUser
Set GFS = GetFreeSlot
GFS.Nick = "ChanServ"
GFS.ID = "ChanServ@" & ServerName & ""
GFS.Ident = "Services"
GFS.DNS = ServerName
GFS.Email = AdminEmail
GFS.IRCOp = True
GFS.Name = "Service"
GFS.SignOn = UnixTime
GFS.Server = ServerName
GFS.ServerDescritption = ServerDesc
GFS.Modes.Add "o", "o"
GFS.GID = CreateGUID
Set GFS = GetFreeSlot
GFS.Nick = "NickServ"
GFS.ID = "NickServ@" & ServerName & ""
GFS.Ident = "Services"
GFS.DNS = ServerName
GFS.Email = AdminEmail
GFS.IRCOp = True
GFS.Name = "Service"
GFS.SignOn = UnixTime
GFS.Server = ServerName
GFS.ServerDescritption = ServerDesc
GFS.Modes.Add "o", "o"
GFS.GID = CreateGUID
Set GFS = GetFreeSlot
GFS.Nick = "MemoServ"
GFS.ID = "MemoServ@" & ServerName & ""
GFS.Ident = "Services"
GFS.DNS = ServerName
GFS.Email = AdminEmail
GFS.IRCOp = True
GFS.Name = "Service"
GFS.SignOn = UnixTime
GFS.Server = ServerName
GFS.ServerDescritption = ServerDesc
GFS.Modes.Add "o", "o"
GFS.GID = CreateGUID
Set GFS = GetFreeSlot
GFS.Nick = "OperServ"
GFS.ID = "OperServ@" & ServerName & ""
GFS.Ident = "Services"
GFS.DNS = ServerName
GFS.Email = AdminEmail
GFS.IRCOp = True
GFS.Name = "Service"
GFS.SignOn = UnixTime
GFS.Server = ServerName
GFS.ServerDescritption = ServerDesc
GFS.Modes.Add "o", "o"
GFS.GID = CreateGUID
End Sub

Public Sub WriteHeader()
With FS.OpenTextFile(LogFile, ForAppending, True)
    .WriteLine "<html>"
    .WriteLine "<head>"
    '.WriteLine "<meta http-equiv=REFRESH content=2>"
    .WriteLine "<title>PURE IRCd Status File</title>"
    .WriteLine "</head>"
    .WriteLine "<body text=#00FF00 bgcolor=#000000>"
    .WriteLine "<p align=center><b>PURE IRCd LOG FILE</b></p>"
    .WriteLine "<table border=1 cellpadding=0 cellspacing=0 style=border-collapse: collapse bordercolor=#111111 width=100% id=AutoNumber1>"
    .WriteLine "    <tr>"
    .WriteLine "        <td width=20% align=center bgcolor=#C0C0C0><font color=#000000><b>Time</b></font></td>"
    .WriteLine "        <td width=20% align=center bgcolor=#C0C0C0><font color=#000000><b>Originator</b></font></td>"
    .WriteLine "        <td width=60% align=center bgcolor=#C0C0C0><font color=#000000><b>Message</b></font></td>"
    .WriteLine "    </tr>"
End With
End Sub

Public Sub WriteFooter()
With FS.OpenTextFile(LogFile, ForAppending, True)
    .WriteLine "</table>"
    .WriteLine "</body>"
    .WriteLine "</html>"
End With
KillTimer frmMain.hWnd, 0
End Sub

Public Function GetFreeOLine() As Long
Dim i As Long
For i = 1 To UBound(Olines)
    If Not Olines(i).InUse Then
        GetFreeOLine = i
        Exit Function
    End If
Next i
End Function

Public Function CreateGUID() As String  'Used to create unieqe ID numbers to identify users Netwide.
    Dim ID(0 To 15) As Byte
    Dim Cnt As Long, GUID As String
    CoCreateGuid ID(0)
        For Cnt = 0 To 15
            CreateGUID = CreateGUID + IIf(ID(Cnt) < 16, "0", "") + Hex$(ID(Cnt))
        Next Cnt
        CreateGUID = Left$(CreateGUID, 8) + "-" + Mid$(CreateGUID, 9, 4) + "-" + Mid$(CreateGUID, 13, 4) + "-" + Mid$(CreateGUID, 17, 4) + "-" + Right$(CreateGUID, 12)
End Function

Public Sub SendLinks(Msg As String, Optional Index As Long, Optional OnlySendToLink)
On Error Resume Next
Debug.Print Msg
ServerTraffic = ServerTraffic + Len(Msg & vbCrLf)
If LogLevel = 1 Then
    If LogFormat = 0 Then
        LogText "[LINK]<" & Now & " (OUTGOING " & Msg & ")> " & Msg
    Else
        LogHTML ServerName & "(Link) OUTGOING", Msg
    End If
End If
If IsMissing(OnlySendToLink) Then
    Dim i As Long
    For i = frmMain.Link.LBound + 1 To frmMain.Link.UBound
        If Not i = Index Then frmMain.Link(i).SendData Msg & vbCrLf
    Next i
Else
    frmMain.Link(OnlySendToLink).SendData Msg & vbCrLf
End If
End Sub
