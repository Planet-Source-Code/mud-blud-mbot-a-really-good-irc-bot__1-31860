Attribute VB_Name = "modSubClass"
Private Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long)
Private Declare Function CallWindowProc& Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long)
Private Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)

Private Const GWL_WNDPROC = (-4&)

Public Const WM_SYSCOMMAND = &H112
Public Const WM_GETMINMAXINFO = &H24

Type POINTAPI
     x As Long
     y As Long
End Type

Type MINMAXINFO
        ptReserved As POINTAPI
        ptMaxSize As POINTAPI
        ptMaxPosition As POINTAPI
        ptMinTrackSize As POINTAPI
        ptMaxTrackSize As POINTAPI
End Type
Private lPrevWndProc As Long

Public Function WindowProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim PassOn As Boolean
PassOn = True
Select Case Msg
Case WM_SYSCOMMAND
If wParam = 61472 Then
    fMain.AddIcon fMain.Icon, "MBot"
    fMain.Hide
    PassOn = False
End If
Case WM_GETMINMAXINFO
Dim MinMax As MINMAXINFO
CopyMemory MinMax, ByVal lParam, Len(MinMax)
MinMax.ptMinTrackSize.x = 400
MinMax.ptMinTrackSize.y = 200
MinMax.ptMaxTrackSize.x = 50000
MinMax.ptMaxTrackSize.y = 50000
CopyMemory ByVal lParam, MinMax, Len(MinMax)
PassOn = False
End Select

If PassOn = True Then
    WindowProc = CallWindowProc(lPrevWndProc, hWnd, Msg, wParam, lParam)
End If
End Function


Public Sub Hook(hWnd As Long)
lPrevWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub Unhook(hWnd As Long)
Call SetWindowLong(hWnd, GWL_WNDPROC, lPrevWndProc)
End Sub
