VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form fMain 
   Caption         =   "MBot 1.0"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10185
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   10185
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD1 
      Left            =   7920
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar Stb 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   4170
      Width           =   10185
      _ExtentX        =   17965
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14790
            MinWidth        =   2646
            Text            =   "-"
            TextSave        =   "-"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "Not Connected"
            TextSave        =   "Not Connected"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Status"
      Height          =   4095
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8055
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1440
         Top             =   2160
      End
      Begin RichTextLib.RichTextBox RTB 
         Height          =   3735
         Left            =   120
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   6588
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"fMain.frx":0E42
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Plugins"
      Height          =   4095
      Left            =   8160
      TabIndex        =   0
      Top             =   0
      Width           =   1935
      Begin VB.CommandButton Command3 
         Caption         =   "Settings"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   3720
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Unload"
         Height          =   255
         Left            =   960
         TabIndex        =   3
         Top             =   3480
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Load"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   3480
         Width           =   855
      End
      Begin VB.ListBox lstPlugins 
         Height          =   3105
         IntegralHeight  =   0   'False
         ItemData        =   "fMain.frx":0EBD
         Left            =   120
         List            =   "fMain.frx":0EBF
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Menu mConnection 
      Caption         =   "Server"
      Begin VB.Menu mConnect 
         Caption         =   "Connect"
      End
      Begin VB.Menu mDisconnect 
         Caption         =   "Disconnect"
         Enabled         =   0   'False
      End
      Begin VB.Menu mLine0 
         Caption         =   "-"
      End
      Begin VB.Menu mSendCommand 
         Caption         =   "Send Command"
      End
      Begin VB.Menu mOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents Server As MBot.Server
Attribute Server.VB_VarHelpID = -1
Dim WithEvents Tray As RealCoding.TrayIcon
Attribute Tray.VB_VarHelpID = -1
Private cPlugins As New Collection

Private Sub Command1_Click()
Dim Tryed2Reg As Boolean
On Error Resume Next
CD1.InitDir = App.Path
CD1.Filter = "DLL Plugins (*.dll)|*.dll|"
CD1.ShowOpen
If CD1.FileName = "" Then Exit Sub 'if they clicked cancel dont try and do ne thing
'---------
Dim FN As String
FN = Mid$(CD1.FileName, InStrRev(CD1.FileName, "\") + 1) 'gets the filename from the path e.g. test.dll
FN = Mid$(FN, 1, Len(FN) - 4) 'removes the .dll part
'---------
TryCreate:
Dim tmpObject
Set tmpObject = CreateObject(FN & ".Plugin") 'try and create object
If tmpObject Is Nothing Then 'didn't create
If Tryed2Reg = True Then ' its been tryed b4 but it still wont work, my guess is it aint a plugin.
AddStatus Mid$(CD1.FileName, InStrRev(CD1.FileName, "\") + 1) & " could not be loaded.", vbRed, True
Else 'ok, lets try and register it.
ret = RegisterServer(Me.hWnd, CD1.FileName, True)
Tryed2Reg = True
GoTo TryCreate 'go back and try create the object again
End If
Else ' Plugin was loaded.
If CDbl(Server.Version) < CDbl(tmpObject.ExpectedServer) Then
    MsgBox "This plugin was not meant for this version of MBot, Please goto www.woot.com and update.", vbCritical + vbOKOnly, "MBot"
Else
    tmpObject.SetServer Server
    lstPlugins.AddItem FN
    AddStatus Mid$(CD1.FileName, InStrRev(CD1.FileName, "\") + 1) & " was loaded.", vbRed, True
    cPlugins.Add tmpObject, FN
End If
End If
Set tmpObject = Nothing
End Sub

Private Sub Command2_Click()
If lstPlugins.ListIndex < 0 Then Exit Sub
Dim Bleh As Object
Set Bleh = cPlugins(lstPlugins.List(lstPlugins.ListIndex))
Bleh.Unload
AddStatus lstPlugins.List(lstPlugins.ListIndex) & ".dll was unloaded.", vbRed, True
Set Bleh = Nothing
cPlugins.Remove lstPlugins.ListIndex + 1
lstPlugins.RemoveItem lstPlugins.ListIndex
End Sub

Private Sub Command3_Click()
If lstPlugins.ListIndex < 0 Then Exit Sub
Dim Bleh As Object
Set Bleh = cPlugins(lstPlugins.List(lstPlugins.ListIndex))
Bleh.Settings
Set Bleh = Nothing
End Sub

Private Sub Form_Load()
Hook Me.hWnd
On Error Resume Next
Set Server = New MBot.Server
Set Tray = New RealCoding.TrayIcon
If Server Is Nothing Then
MsgBox "Error: A DLL couldn't be loaded, please reinstall.", vbCritical + vbOKOnly
End
End If
If Tray Is Nothing Then
MsgBox "Error: A DLL couldn't be loaded, please reinstall.", vbCritical + vbOKOnly
End
ElseIf Not Server Is Nothing Then
AddStatus "Server DLL Loaded [Version " & Server.Version & "]" & vbCrLf, vbRed, True
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Server.Disconnect
End Sub

Private Sub Form_Resize()
On Error Resume Next
Frame2.Move 3, 0, Me.ScaleWidth - (Frame1.Width + 12), Me.ScaleHeight - (Stb.Height + 4)
Frame1.Move Me.ScaleWidth - (Frame1.Width + 3), 0, Frame1.Width, Me.ScaleHeight - (Stb.Height + 4)
RTB.Move 100, 230, Frame2.Width - 230, Frame2.Height - 320
Command3.Move 100, Frame1.Height - 380, Command1.Width + Command2.Width, Command3.Height
Command1.Move 100, Command3.Top - (Command3.Height), Command1.Width, Command1.Height
Command2.Move (Command1.Left + Command1.Width), Command3.Top - (Command3.Height), Command2.Width, Command2.Height
lstPlugins.Move 100, 230, Frame1.Width - 230, (Command1.Top - Command1.Height) - 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
Tray.RemoveIcon
Unhook Me.hWnd
For x = 1 To cPlugins.Count: cPlugins.Remove x: Next x
Set Tray = Nothing
Set Server = Nothing
End Sub

Private Sub mConnect_Click()
Server.Connect GetSetting("MBot", "Settings", "Server", "irc.dal.net"), GetSetting("MBot", "Settings", "Port", "6667")
End Sub

Private Sub mDisconnect_Click()
Server.Disconnect
End Sub

Private Sub mExit_Click()
Unload Me
End
End Sub

Private Sub mOptions_Click()
fOptions.Show vbModal, Me
End Sub

Private Sub Server_AddStatus(Message As String, Origin As String)
AddStatus Origin & ">" & Message, vbBlue, True, True, True
End Sub

Private Sub Server_ChannelCTCP(Channel As String, From As String, Message As String, Params As String)
Select Case UCase(Message)
Case "PING"
Server.SendNotice Mid$(From, 1, InStr(From, "!") - 1), Chr$(1) & "PING " & Params & Chr$(1)
AddStatus "Ping? Pong! from " & From & " in " & Channel, RGB(0, 128, 0), True, True
End Select
End Sub

Private Sub Server_JoinChannel(Channel As String, From As String)
If Mid$(From, 1, InStr(From, "!") - 1) = Server.Nick Then
AddStatus "Joined: " & Channel, RGB(150, 12, 150)
End If
End Sub

Private Sub Server_PrivateCTCP(From As String, Message As String, Params As String)
Select Case UCase(Message)
Case "PING"
Server.SendNotice Mid$(From, 1, InStr(From, "!") - 1), Chr$(1) & "PING " & Params & Chr$(1)
AddStatus "Ping? Pong! from " & From, RGB(0, 128, 0), True, True
End Select
End Sub

Private Sub Server_RecvRaw(Dta As String)
Dim Word: Word = Split(Dta, " ")
If UCase(Mid$(Dta, 1, 20)) = "ERROR :CLOSING LINK:" Then AddStatus "Disconnected " & Mid$(Dta, InStr(Dta, "(")) & vbCrLf, RGB(0, 0, 128), True: Exit Sub
If UCase(Word(1)) = "NOTICE" And UCase(Word(2)) = "AUTH" Then AddStatus Mid$(Dta, InStr(Dta, ":") + 1), RGB(128, 0, 0), True: Exit Sub
End Sub

Private Sub Server_ServerPing(Svr As String)
AddStatus "Ping? Pong!", RGB(0, 128, 0), True, True
End Sub

Private Sub Server_ServerReady(Svr As String, Nick As String)
AddStatus "Connected to " & Svr & " as " & Nick, RGB(0, 0, 128), True

End Sub

Private Sub Server_StateChange(State As Integer)
Select Case State
Case sckConnecting
    AddStatus "Connecting to " & GetSetting("MBot", "Settings", "Server", "irc.dal.net") & ":" & GetSetting("MBot", "Settings", "Port", "6667"), RGB(0, 0, 128), True
    Stb.Panels(2).Text = "Connecting..."
    mConnect.Enabled = False
    mDisconnect.Enabled = True
    mOptions.Enabled = False
Case sckConnected
    Timer1.Enabled = True
    Stb.Panels(2).Text = "Connected"
    mConnect.Enabled = False
    mDisconnect.Enabled = True
    mOptions.Enabled = False
    Server.SendRaw "NICK " & GetSetting("MBot", "Settings", "Nick", "MBot")
    Server.SendRaw "USER MBot host server :a bot? with a real name? suurreeeee"
Case sckClosed
    Timer1.Enabled = False
    AddStatus "Sucessfully disconnected.", RGB(0, 128, 0), True
    Stb.Panels(1).Text = "-"
    Stb.Panels(2).Text = "Not Connected"
    mConnect.Enabled = True
    mDisconnect.Enabled = False
    mOptions.Enabled = True
End Select
End Sub

Private Sub mSendCommand_Click()
Dim InputBx As String
InputBx = InputBox("Enter raw command to send to server:", "MBot", "")
If InputBx = "" Then Exit Sub
Server.SendRaw InputBx
End Sub

Private Sub Timer1_Timer()
Stb.Panels(1).Text = Server.UpTime
End Sub

Private Sub Tray_LeftButton(ClickType As RealCoding.ClickConstants)
If ClickType = MouseUp Then Me.Show: RemoveIcon
End Sub

Public Sub AddStatus(Text As String, Optional Color As Long = vbBlack, Optional Bold As Boolean = False, Optional Italic As Boolean = False, Optional Underline As Boolean = False)
RTB.Visible = False
RTB.SelStart = Len(RTB.Text)
RTB.SelText = Text & vbCrLf
RTB.SelStart = Len(RTB.Text) - Len(Text) - 2
RTB.SelLength = Len(Text)
RTB.SelItalic = Italic
RTB.SelBold = Bold
RTB.SelUnderline = Underline
RTB.SelColor = Color
RTB.SelStart = Len(RTB.Text)
RTB.Visible = True
End Sub

Public Sub AddIcon(Icon As Long, ToolTip As String)
Tray.AddIcon Icon, ToolTip
End Sub

Public Sub ModifyIcon(Icon As Long, ToolTip As String)
Tray.ModifyIcon Icon, ToolTip
End Sub

Public Sub RemoveIcon()
Tray.RemoveIcon
End Sub
