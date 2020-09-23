Attribute VB_Name = "modRegSvr"
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Any, ByVal wParam As Any, ByVal lParam As Any) As Long
Private Const ERROR_SUCCESS = &H0
Private Const ERROR_AHHHHHH = &HF

Public Function RegisterServer(hWnd As Long, DllServerPath As String, bRegister As Boolean)
On Error Resume Next
Dim lb As Long, pa As Long
lb = LoadLibrary(DllServerPath)
If bRegister Then
pa = GetProcAddress(lb, "DllRegisterServer")
Else
pa = GetProcAddress(lb, "DllUnregisterServer")
End If
If CallWindowProc(pa, hWnd, ByVal 0&, ByVal 0&, ByVal 0&) = ERROR_SUCCESS Then
RegisterServer = ERROR_SUCCESS
Else
RegisterServer = ERROR_AHHHHHH
End If
FreeLibrary lb
End Function

