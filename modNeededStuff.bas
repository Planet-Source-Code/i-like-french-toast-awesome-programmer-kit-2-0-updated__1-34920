Attribute VB_Name = "modNeededStuff"
'*** You dont need any of this,
'*** this is just what I need to
'*** make my other functions good! ***

Option Explicit

Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Public Const SND_FILENAME = &H2000
Public FormHandle As Long
Public mvarbRunningInTray As Boolean
Public SysIcon As NOTIFYICONDATA
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Const SPI_SCREENSAVERRUNNING = 97

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
Public Const REG_SZ = 1
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4
Public Const ERROR_SUCCESS = 0&

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" _
(ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
(ByVal hKey As Long, ByVal lpSubKey As String) As Long

Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
(ByVal hKey As Long, ByVal lpValueName As String) As Long

Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" _
(ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
(ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
lpType As Long, lpData As Any, lpcbData As Long) As Long

Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
(ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal _
dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" _
(ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName _
As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, _
lpftLastWriteTime As Any) As Long

Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" _
(ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As _
Byte, lpcbData As Long) As Long

Public Function GetSettingString(ByVal hKey As Long, _
ByVal strPath As String, ByVal strValue As String, Optional _
Default As String) As String
Dim hCurKey As Long
Dim lResult As Long
Dim lValueType As Long
Dim strBuffer As String
Dim lDataBufferSize As Long
Dim intZeroPos As Integer
Dim lRegResult As Long

If Not IsEmpty(Default) Then
GetSettingString = Default
Else
GetSettingString = ""
End If

lRegResult = RegOpenKey(hKey, strPath, hCurKey)
lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, _
lValueType, ByVal 0&, lDataBufferSize)

If lRegResult = ERROR_SUCCESS Then

If lValueType = REG_SZ Then

strBuffer = String(lDataBufferSize, " ")
lResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, _
ByVal strBuffer, lDataBufferSize)

intZeroPos = InStr(strBuffer, Chr$(0))
If intZeroPos > 0 Then
GetSettingString = Left$(strBuffer, intZeroPos - 1)
Else
GetSettingString = strBuffer
End If

End If

Else
' there is a problem
End If

lRegResult = RegCloseKey(hCurKey)
End Function
