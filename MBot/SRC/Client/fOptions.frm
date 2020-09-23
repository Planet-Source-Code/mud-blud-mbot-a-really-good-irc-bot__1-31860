VERSION 5.00
Begin VB.Form fOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   73
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   393
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox tNick 
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Text            =   "MBot"
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox tPort 
      Height          =   285
      Left            =   4560
      TabIndex        =   3
      Text            =   "irc.dal.net"
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox tServer 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Text            =   "irc.dal.net"
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Nick:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   630
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Port:"
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   150
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "IP/Host:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   975
   End
End
Attribute VB_Name = "fOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
SaveSetting "MBot", "Settings", "Server", tServer.Text
SaveSetting "MBot", "Settings", "Port", tPort.Text
SaveSetting "MBot", "Settings", "Nick", tNick.Text
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
tServer.Text = GetSetting("MBot", "Settings", "Server", "irc.dal.net")
tPort.Text = GetSetting("MBot", "Settings", "Port", "6667")
tNick.Text = GetSetting("MBot", "Settings", "Nick", "MBot")
End Sub
