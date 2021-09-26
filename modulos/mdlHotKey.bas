Attribute VB_Name = "mdlHotKey"
Option Explicit
Private Const WM_HOTKEY = &H312

Public Declare Function RegisterHotKey Lib _
"user32" (ByVal hWnd As Long, ByVal _
id As Long, ByVal fsModifiers As _
Long, ByVal vk As Long) As Long
'
' Modificadores
'
Public Const MOD_ALT = &H1
Public Const MOD_CONTROL = &H2
Public Const MOD_SHIFT = &H4

' Tecla que será nossa hotkey
Public Const VK_F2 = &H71

#If UNICODE Then
Public Declare Function SetWindowLong Lib _
"user32" Alias "SetWindowLongW" _
(ByVal hWnd As Long, ByVal nIndex _
As Long, ByVal dwNewLong As Any) _
As Long
#Else
Public Declare Function SetWindowLong Lib _
"user32" Alias "SetWindowLongA" _
(ByVal hWnd As Long, ByVal nIndex _
As Long, ByVal dwNewLong As Any) _
As Long
#End If
Private Declare Function CallWindowProc Lib _
"user32" Alias "CallWindowProcA" _
(ByVal wndrpcPrev As Long, ByVal _
hWnd As Long, ByVal uMsg As Long, _
ByVal wParam As Long, lParam As _
Any) As Long
Public Const GWL_USERDATA = (-21)
Public Const GWL_WNDPROC = -4

Public Function WindowProc(ByVal hWnd As _
Long, ByVal uMsg As Long, ByVal _
wParam As Long, ByVal lParam As _
Long) As Long
On Error Resume Next
If uMsg = WM_HOTKEY And wParam = 1 Then
'wParam informa o ID da hotkey
frmPrincipal.BahMetodo
WindowProc = 1
Exit Function
End If
If frmPrincipal.OldWndProc <> 0 Then
WindowProc = CallWindowProc(frmPrincipal.OldWndProc, _
hWnd, uMsg, wParam, ByVal lParam)
End If
End Function



