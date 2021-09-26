Attribute VB_Name = "mdlPrincipal"
'Verifica se aplicação já foi aberta e se sim ele finaliza o sistema
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'----------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, _
                                                                        ByVal lpOperation As String, _
                                                                        ByVal lpFile As String, _
                                                                        ByVal lpParameters As String, _
                                                                        ByVal lpDirectory As String, _
                                                                        ByVal nShowCmd As Long) As Long
'Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long

Public Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Type NOTIFYICONDATA
  cbSize As Long
  hWnd As Long
  uId As Long
  uFlags As Long
  uCallBackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type

Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200

Public Const WM_LBUTTONDOWN = &H201     'Button down
Public Const WM_LBUTTONUP = &H202       'Button up
Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
Public Const WM_RBUTTONDOWN = &H204     'Button down
Public Const WM_RBUTTONUP = &H205       'Button up
Public Const WM_RBUTTONDBLCLK = &H206   'Double-click

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
'Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias " Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public nid As NOTIFYICONDATA
Sub Main()
    
    If App.PrevInstance Then
        MsgBox "Aplicação já está aberta", vbCritical
        End
    Else
        frmPrincipal.Show
    End If

End Sub
Sub TOChild(Hwnd_Parent As Long, Hwnd_Child As Long)
SetParent Hwnd_Child, Hwnd_Parent
End Sub
Sub NotChild(Hwnd_Child As Long)
SetParent Hwnd_Child, 0&
End Sub
Function AlwaysOnTop(FrmID As Form, ByVal OnTop As Boolean) As Boolean

    Rem --- Deixa o form sempre no topo das demais janelas abertas no Windows ---
    Const SWP_NOMOVE = 2
    Const SWP_NOSIZE = 1
    Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2

    If OnTop = True Then
        AlwaysOnTop = SetWindowPos(FrmID.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Else
        AlwaysOnTop = SetWindowPos(FrmID.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
    End If

End Function
Public Function MontaHora(ByVal Seg As Long) As String
    Dim Hor As Long, Min As Long
    
    If Seg < 0 Then
        MontaHora = "- "
        Seg = Abs(Seg)
    Else
        MontaHora = ""
    End If
    
    Hor = Seg \ 3600&
    Seg = Seg - (Hor * 3600&)
    
    Min = Seg \ 60&
    Seg = Seg - (Min * 60&)
    
    MontaHora = MontaHora & Hor & ":" & _
        Format(Min, "0#") & ":" & _
        Format(Seg, "0#")
End Function

