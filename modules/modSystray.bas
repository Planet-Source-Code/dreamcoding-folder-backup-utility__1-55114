Attribute VB_Name = "modSystray"
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const REG_SZ = 1
Private Const KEY_WRITE = 131078

Global glIsConnected As Long
Global glWasConnected As Boolean
Global curSecond As Double

Public Const NIM_ADD = &H0 'Add to Tray
Public Const NIM_MODIFY = &H1 'Modify Details
Public Const NIM_DELETE = &H2 'Remove From Tray
Public Const NIF_MESSAGE = &H1 'Message
Public Const NIF_ICON = &H2 'Icon
Public Const NIF_TIP = &H4 'TooTipText
Public Const WM_MOUSEMOVE = &H200 'On Mousemove
Public Const WM_LBUTTONDOWN = &H201 'Left Button Down
Public Const WM_LBUTTONUP = &H202 'Left Button Up
Public Const WM_LBUTTONDBLCLK = &H203 'Left Double Click
Public Const WM_RBUTTONDOWN = &H204 'Right Button Down
Public Const WM_RBUTTONUP = &H205 'Right Button Up
Public Const WM_RBUTTONDBLCLK = &H206 'Right Double Click

Public nid As NOTIFYICONDATA

Public Type NOTIFYICONDATA
 cbSize As Long
 hwnd As Long
 uId As Long
 uFlags As Long
 uCallBackMessage As Long
 hIcon As Long
 szTip As String * 64
End Type

Public Function DoStartUp(Filename As String, Discription As String)
Dim hKey As Long
 RegOpenKeyEx HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run-", 0, KEY_WRITE, hKey
 RegDeleteValue hKey, Discription
 RegCloseKey hKey
 RegOpenKeyEx HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", 0, KEY_WRITE, hKey
 RegSetValueEx hKey, Discription, 0, REG_SZ, Filename, Len(Filename)
 RegCloseKey hKey
End Function

Public Function DoNotStartUp(Filename As String, Discription As String)
Dim hKey As Long
 RegOpenKeyEx HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", 0, KEY_WRITE, hKey
 RegDeleteValue hKey, Discription
 RegCloseKey hKey
 RegOpenKeyEx HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run-", 0, KEY_WRITE, hKey
 RegSetValueEx hKey, Discription, 0, REG_SZ, Filename, Len(Filename)
 RegCloseKey hKey
End Function

