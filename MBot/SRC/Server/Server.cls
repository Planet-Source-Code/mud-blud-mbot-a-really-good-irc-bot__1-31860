VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private WithEvents Sck As CSocket
Attribute Sck.VB_VarHelpID = -1
Private WithEvents MasterTimer As vbTimer.MasterTimer
Attribute MasterTimer.VB_VarHelpID = -1
Private Tmr As vbTimer.Timer
Private Const SERVER_VERSION As String = "1.0"
Private m_ConnectedTick As Long
Private m_Nick As String

Private Declare Function GetTickCount Lib "kernel32.dll" () As Long

Public Event StateChange(State As Integer)
Public Event RecvRaw(Dta As String)
Public Event ServerReady(Svr As String, YourNick As String)
Public Event ServerPing(Svr As String)
Public Event ChannelMessage(Channel As String, from As String, Message As String, Action As Boolean)
Public Event ChannelNotice(Channel As String, from As String, Message As String)
Public Event ServerPrivateNotice(Svr As String, Message As String)
Public Event PrivateMessage(from As String, Message As String, Action As Boolean)
Public Event PrivateNotice(from As String, Message As String)
Public Event JoinChannel(Channel As String, from As String)
Public Event PartChannel(Channel As String, from As String)
Public Event NickChanged(OldNick As String, NewNick As String)
Public Event ModeChanged(Channel As String, Op As String, Modes As String)
Public Event TopicChanged(Channel As String, Op As String, Topic As String)
Public Event ChannelCTCP(Channel As String, from As String, Message As String, Params As String)
Public Event PrivateCTCP(from As String, Message As String, Params As String)
Public Event AddStatus(Message As String, Origin As String)

Public Property Get Version() As String
Version = SERVER_VERSION
End Property

Public Function Nick() As String
Nick = m_Nick
End Function

Private Sub Class_Initialize()
Set Sck = New CSocket
Set MasterTimer = New vbTimer.MasterTimer
Set Tmr = MasterTimer.Timers.Add("State", 1, ttInfinite, tsMilliseconds)
Tmr.Enabled = True
End Sub

Private Sub Class_Terminate()
Set Sck = Nothing
Set Tmr = Nothing
Set MasterTimer = Nothing
End Sub

Private Sub MasterTimer_TimerFired(ByVal tmrTimer As vbTimer.Timer, ByVal lngTickCount As Long)
Select Case tmrTimer.Key
Case "State"
Static OldState As Integer
If OldState = Sck.State Then Exit Sub
RaiseEvent StateChange(Sck.State)
OldState = Sck.State
End Select
End Sub
'==============START OF PUBLIC FUNCTIONS==============
Public Sub Connect(Host As String, Port As Integer)
Sck.Connect Host, Port
End Sub

Public Sub Disconnect()
SendRaw "QUIT :MBot by MudBlud - Please vote http://www.pscode.com/vb"
DoEvents
Sck.CloseSocket
End Sub

Public Sub SendRaw(Data As String)
If Sck.State <> SckConnected Then Exit Sub
Sck.SendData Data & vbCrLf
End Sub

Public Sub SendNotice(Target As String, Message As String)
If Sck.State <> SckConnected Then Exit Sub
Sck.SendData "NOTICE " & Target & " :" & Message & vbCrLf
End Sub

Public Sub SendMessage(Target As String, Message As String, Optional Action As Boolean = False)
If Sck.State <> SckConnected Then Exit Sub
If Action = True Then Message = Chr$(1) & "ACTION " & Message & Chr$(1)
Sck.SendData "PRIVMSG " & Target & " :" & Message & vbCrLf
End Sub

Public Sub ChangeMode(Target As String, Modes As String)
If Sck.State <> SckConnected Then Exit Sub
Sck.SendData "MODE " & Target & " " & Modes & vbCrLf
End Sub

Public Sub ChangeNick(NewNick As String)
If Sck.State <> SckConnected Then Exit Sub
Sck.SendData "NICK " & NewNick & vbCrLf
End Sub

Public Sub JoinChannel(Channel As String)
If Sck.State <> SckConnected Then Exit Sub
Sck.SendData "JOIN " & Channel & vbCrLf
End Sub

Public Sub PartChannel(Channel As String)
If Sck.State <> SckConnected Then Exit Sub
Sck.SendData "PART " & Channel & vbCrLf
End Sub

Public Sub ChangeTopic(Channel As String, Topic As String)
If Sck.State <> SckConnected Then Exit Sub
Sck.SendData "TOPIC " & Channel & " :" & Topic & vbCrLf
End Sub

Public Function UpTime() As String
UpTime = FormatCount(GetTickCount - m_ConnectedTick)
End Function

Private Sub Sck_OnClose()
m_ConnectedTick = 0
End Sub

Private Sub Sck_OnConnect()
m_ConnectedTick = GetTickCount
End Sub

Private Sub Sck_OnDataArrival(ByVal bytesTotal As Long)
Dim Dta As String
Sck.GetData Dta, vbString
DoLine Dta
End Sub

Private Sub DoLine(strLine As String)
Dim i As Integer
Dim strWord As String
Static RestLine As String
strLine = RestLine & strLine
For i = 1 To Len(strLine)
If Mid(strLine, i, 1) = Chr(13) Or Mid(strLine, i, 1) = Chr(10) Then
strWord = Mid(strLine, 1, i - 1)
If Mid(strWord, 1, 1) = Chr(13) Or Mid(strWord, 1, 1) = Chr(10) Then strWord = Trim$(Mid(strWord, 2))
If strWord <> "" Then Parse strWord
strLine = Mid(strLine, i + 1)
i = 1
End If
Next i
RestLine = strLine
End Sub

Private Sub Parse(Dta As String)
On Error Resume Next
Dim msg As String, prm As String
If Mid$(Dta, 1, 4) = "PING" Then SendRaw "PONG " & Mid$(Dta, 6, Len(Dta) - 6): RaiseEvent ServerPing(Mid$(Dta, 7, Len(Dta) - 7)): Exit Sub
If Left$(Dta, 1) = ":" Then Dta = Mid$(Dta, 2) 'take off the :
Dim Word: Word = Split(Dta, " ")
If IsNumeric(Word(1)) Then 'its a number..
    Select Case Word(1)
    Case "001"
        m_Nick = CStr(Word(2))
        RaiseEvent ServerReady(CStr(Word(0)), CStr(Word(2)))
    End Select
Else 'its text..
    Select Case UCase(Word(1))
    Case "NOTICE"
        If Mid$(Word(2), 1, 1) = "#" Then 'its channel
            RaiseEvent ChannelNotice(CStr(Word(2)), CStr(Word(0)), Mid$(Dta, InStr(Dta, ":") + 1))
        ElseIf InStr(Word(0), ".") > 0 Then 'its a server
            RaiseEvent ServerPrivateNotice(CStr(Word(0)), Mid$(Dta, InStr(Dta, ":") + 1))
        Else 'its a notice 4 u :)
            RaiseEvent PrivateNotice(CStr(Word(0)), Mid$(Dta, InStr(Dta, ":") + 1))
        End If
    Case "PRIVMSG"
        If Mid$(Word(2), 1, 1) = "#" Then ':PING 1013885324
            If Right$(Mid$(Dta, InStr(Dta, ":") + 1), 1) = Chr(1) And Left$(Mid$(Dta, InStr(Dta, ":") + 1), 1) = Chr(1) Then
                msg = Mid$(Dta, InStr(Dta, ":") + 2)
                msg = Mid$(msg, 1, InStr(msg, " ") - 1)
                prm = Mid$(Dta, InStr(Dta, ":") + 2)
                prm = Mid$(prm, InStr(prm, " ") + 1)
                prm = Left$(prm, Len(prm) - 1)
                If msg = "ACTION" Then
                    RaiseEvent ChannelMessage(CStr(Word(2)), CStr(Word(0)), prm, True)
                Else
                    RaiseEvent ChannelCTCP(CStr(Word(2)), CStr(Word(0)), msg, prm)
                End If
            Else
                RaiseEvent ChannelMessage(CStr(Word(2)), CStr(Word(0)), Mid$(Dta, InStr(Dta, ":") + 1), False)
            End If
        Else 'its a pm 4 u :)
            If Right$(Mid$(Dta, InStr(Dta, ":") + 1), 1) = Chr(1) And Left$(Mid$(Dta, InStr(Dta, ":") + 1), 1) = Chr(1) Then
                msg = Mid$(Dta, InStr(Dta, ":") + 2)
                msg = Mid$(msg, 1, InStr(msg, " ") - 1)
                prm = Mid$(Dta, InStr(Dta, ":") + 2)
                prm = Mid$(prm, InStr(prm, " ") + 1)
                prm = Left$(prm, Len(prm) - 1)
                If msg = "ACTION" Then
                    RaiseEvent PrivateMessage(CStr(Word(0)), prm, True)
                Else
                    RaiseEvent PrivateCTCP(CStr(Word(0)), msg, prm)
                End If
            Else
                RaiseEvent PrivateMessage(CStr(Word(0)), Mid$(Dta, InStr(Dta, ":") + 1), False)
            End If
        End If
    Case "JOIN" ':nick!user@host JOIN :target
        RaiseEvent JoinChannel(Mid$(Word(2), 2), CStr(Word(0)))
    Case "PART" ':nick!user@host PART :target
        RaiseEvent PartChannel(CStr(Word(2)), CStr(Word(0)))
    Case "NICK" ':oldnick!u@h NICK :new nick
        RaiseEvent NickChanged(Mid$(Word(0), 1, InStr(Word(0), "!") - 1), Mid$(Word(2), 2))
        If LCase(Mid$(Word(0), 1, InStr(Word(0), "!") - 1)) = LCase(m_Nick) Then m_Nick = Mid$(Word(2), 2)
    Case "MODE" ':nick!user@host MODE chan +gfh other
        tmp = Mid$(Dta, InStr(Dta, " ") + 1)
        Debug.Print tmp
        tmp = Mid$(tmp, InStr(tmp, " ") + 1)
        Debug.Print tmp
        tmp = Mid$(tmp, InStr(tmp, " ") + 1)
        RaiseEvent ModeChanged(CStr(Word(2)), Mid$(Word(0), 1, InStr(Word(0), "!") - 1), CStr(tmp))
    Case "TOPIC" ':n!u@h TOPIC channel :newtopic
        RaiseEvent TopicChanged(CStr(Word(2)), CStr(Word(0)), Mid$(Dta, InStr(Dta, ":") + 1))
    End Select
End If
RaiseEvent RecvRaw(Dta)
End Sub

Private Function FormatCount(Count As Long) As String
Dim Weeks As Long, Days As Long, Hours As Long, Minutes As Long, Seconds As Long, txt As String

Count = Count \ 1000
Weeks = Count \ ((24& * 3600&) \ 7)
If Weeks > 0 Then Count = Count - (((24& * 3600&) \ 7) * Weeks)
Days = Count \ (24& * 3600&)
If Days > 0 Then Count = Count - (24& * 3600& * Days)
Hours = Count \ 3600&
If Hours > 0 Then Count = Count - (3600& * Hours)
Minutes = Count \ 60
Seconds = Count Mod 60
If Weeks > 0 Then txt = txt & Weeks & "weeks "
If Days > 0 Then txt = txt & Days & "days "
If Hours > 0 Then txt = txt & Hours & "hours "
If Minutes > 0 Then txt = txt & Minutes & "minutes "
If Seconds > 0 Then txt = txt & Seconds & "seconds"
FormatCount = Trim$(txt) & "."
End Function
