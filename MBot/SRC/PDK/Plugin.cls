VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Plugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private WithEvents Server As MBot.Server
Attribute Server.VB_VarHelpID = -1
Private Const PLUGIN_EXPECTEDSERVER As String = "1.0" 'What version does this plugin use?

Public Sub SetServer(ByRef Svr As MBot.Server)
Set Server = Svr
End Sub

Public Sub Unload()
Set Server = Nothing
End Sub

Public Sub Settings()
fSettings.Show vbModal
End Sub

Public Function ExpectedServer() As String
ExpectedServer = PLUGIN_EXPECTEDSERVER
End Function

'############################################
'### DO NOT EDIT ANYTHING ABOVE THIS LINE ###
'############################################


