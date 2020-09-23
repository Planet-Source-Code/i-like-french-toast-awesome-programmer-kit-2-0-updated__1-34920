Attribute VB_Name = "modSubs"
'************************************
'Programmer Kit made by Jacob Grice.
'You are free to use the programmer kit,
'but please post feedback and/or rate
'it. Thanks, -Jake
'************************************
'Here are the categories:
'Files
'Registry
'Encryption
'Windows
'System Tray
'Sound

'$$$ API STUFF FIRST $$$'
Option Explicit

Public Enum WindowsPaths
    WindowsDir = 0
    SystemDir = 1
    DesktopDir = 2
    CacheDir = 3
    StartupDir = 4
    StartPrograms = 5
    StartMenu = 6
End Enum

'$$$ FILES $$$'

Public Function FileExists(file As String) As Boolean
    Dim fs
    Set fs = CreateObject("Scripting.FileSystemObject")
    FileExists = IIf(fs.FileExists(file) = True, True, False)
End Function
 
Public Function CreateFile(file As String) As Boolean
    On Error GoTo e
    Open file For Output As #1
    Close #1
    CreateFile = True
    Exit Function
e:
    CreateFile = False
End Function

Public Function ReadFileToString(file As String, Optional Include_vbCrLf As Boolean = True) As String
    On Error GoTo e
    ReadFileToString = vbNullString
    Open file For Input As #1
    Dim tmpstr As String
    Do While Not EOF(1)
    Line Input #1, tmpstr
    ReadFileToString = ReadFileToString & tmpstr & IIf(Include_vbCrLf = True, vbCrLf, vbNullString)
    Loop
    Close #1
    If Not Len(ReadFileToString) = 0 Then ReadFileToString = Left$(ReadFileToString, Len(ReadFileToString) - 1)
    Exit Function
e:
    ReadFileToString = "[File doesn't exist or is protected.]"
    MsgBox ReadFileToString
    ReadFileToString = vbNullString
End Function

Public Function FileDel(file As String) As Boolean
    On Error GoTo e
    Kill file
    FileDel = True
    Exit Function
e:
    FileDel = False
End Function

'$$$ REGISTRY $$$'
Public Function KeyInReg(AppName As String, Section As String, Key As String) As Boolean
    KeyInReg = IIf(Len(GetSetting(AppName, Section, Key)) = 0, False, True)
End Function

Public Function GetKeyValue(AppName As String, Section As String, Key As String) As String
    GetKeyValue = GetSetting(AppName, Section, Key)
End Function

Public Function SetKeyValue(AppName As String, Section As String, Key As String, KeyValue As String) As Boolean
    On Error GoTo e
    SaveSetting AppName, Section, Key, KeyValue
    SetKeyValue = True
    Exit Function
e:
    SetKeyValue = False
End Function

Public Function DeleteKey(AppName As String, Section As String, Key As String) As Boolean
    On Error GoTo e
    DeleteSetting AppName, Section, Key
    DeleteKey = True
    Exit Function
e:
    DeleteKey = False
End Function

Public Function DeleteSection(AppName As String, Section As String) As Boolean
    On Error GoTo e
    DeleteSetting AppName, Section
    DeleteSection = True
    Exit Function
e:
    DeleteSection = False
End Function

Public Function DeleteAppRegEntries(AppName As String) As Boolean
    On Error GoTo e
    DeleteSetting AppName
    DeleteAppRegEntries = True
    Exit Function
e:
    DeleteAppRegEntries = False
End Function

'$$$$ ENCRYPTION $$$$'
Public Function Encrypt(ByVal text As String) As String
Dim i As Long, temp As Long, tempText As String
Encrypt = ""
i = 0
temp = 0
tempText = ""
If Trim(text) = "" Then
Encrypt = ""
Exit Function
End If
For i = 1 To Len(text)
temp = Asc(Mid(text, i, 1))
If temp + 50 > 255 Then
temp = (temp + 50) - 255
Else
temp = temp + 50
End If
tempText = Chr(temp)
Encrypt = Encrypt + tempText
Next i
End Function

Public Function Decrypt(ByVal dText As String) As String
Dim i As Long, temp As Long, tempText As String
Decrypt = ""
i = 0
temp = 0
tempText = ""
If Trim(dText) = "" Then
Decrypt = ""
Exit Function
End If
For i = 1 To Len(dText)
temp = Asc(Mid(dText, i, 1))
If temp - 50 < 0 Then
temp = 255 - (-1 * (temp - 50))
Else
temp = temp - 50
End If
tempText = Chr(temp)
Decrypt = Decrypt + tempText
Next i
End Function

'$$$$ MICROSOFT WINDOWS $$$$'

Public Function GetDir(dir As WindowsPaths) As String
    GetDir = Empty

    If dir = CacheDir Then
    GetDir = GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", _
    "Cache")
    
    ElseIf dir = DesktopDir Then
    GetDir = GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", _
    "Desktop")
    
    ElseIf dir = StartPrograms Then
    GetDir = GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", _
    "Programs")
    
    ElseIf dir = StartMenu Then
    GetDir = GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", _
    "Start Menu")
    
    ElseIf dir = StartupDir Then
    GetDir = GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", _
    "Startup")
    
    ElseIf dir = SystemDir Then
    GetDir = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Setup", _
    "SysDir")
    
    ElseIf dir = WindowsDir Then
    GetDir = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Setup", _
    "WinDir")
    End If
End Function

Public Sub DisableCtrlAltDel()
'Disable Ctrl + Alt + Del
On Error GoTo error
Dim ret As Integer
Dim pOld As Boolean
ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub EnableCtrlAltDel()
'Enable Ctrl + Alt + Del
On Error GoTo error
Dim ret As Integer
Dim pOld As Boolean
ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub AlwaysOnTop(F As Form)
'sets the given form On TopMost
Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2

    SetWindowPos F.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Public Sub NotAlwaysOnTop(F As Form)
'sets the given form Off TopMost
Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2

    SetWindowPos F.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Public Sub MoveMouse(x As Long, y As Long)
    SetCursorPos x, y
End Sub

Public Function MousePos() As POINTAPI
    GetCursorPos MousePos
End Function

'$$$$ SYSTRAY $$$$'
Public Sub ShowIcon(ByRef Systrayform As Form)
    SysIcon.cbSize = Len(SysIcon)
    SysIcon.hwnd = Systrayform.hwnd
    SysIcon.uId = vbNull
    SysIcon.uFlags = 7
    SysIcon.ucallbackMessage = 512
    SysIcon.hIcon = Systrayform.Icon
    SysIcon.szTip = Systrayform.Caption + Chr(0)
    Shell_NotifyIcon 0, SysIcon
    mvarbRunningInTray = True
End Sub

Public Sub RemoveIcon(Systrayform As Form)
    SysIcon.cbSize = Len(SysIcon)
    SysIcon.hwnd = Systrayform.hwnd
    SysIcon.uId = vbNull
    SysIcon.uFlags = 7
    SysIcon.ucallbackMessage = vbNull
    SysIcon.hIcon = Systrayform.Icon
    SysIcon.szTip = Chr(0)
    Shell_NotifyIcon 2, SysIcon
    If Systrayform.Visible = False Then Systrayform.Show    'Incase user can't see form
    mvarbRunningInTray = False
End Sub

Public Sub ChangeIcon(Systrayform As Form, picNewIcon As PictureBox)
    If mvarbRunningInTray = True Then   'If running in the tray
        SysIcon.cbSize = Len(SysIcon)
        SysIcon.hwnd = Systrayform.hwnd
        'SysIcon.uId = vbNull
        'SysIcon.uFlags = 7
        'SysIcon.ucallbackMessage = 512
        SysIcon.hIcon = picNewIcon.Picture
        'SysIcon.szTip = sysTrayForm.Caption + Chr(0)
        Shell_NotifyIcon 1, SysIcon
    End If
End Sub

Public Sub ChangeToolTip(Systrayform As Form, strNewTip As String)
    If mvarbRunningInTray = True Then   'If running in the tray
        SysIcon.cbSize = Len(SysIcon)
        SysIcon.hwnd = Systrayform.hwnd
        SysIcon.szTip = strNewTip & Chr(0)
        Shell_NotifyIcon 1, SysIcon
    End If
End Sub

'$$$$ SOUND $$$$'
Public Function sndPlay(sfile As String)
    PlaySound sfile, ByVal 0&, SND_FILENAME
End Function

'$$$ INTERNET $$$'
Public Function IsConnected() As Boolean
    If InternetGetConnectedState(0&, 0&) = 1 Then
        IsConnected = True
    Else
        IsConnected = False
    End If
End Function

Public Function DownloadFile(URL As String, LocalFilename As String) As Boolean
    Dim lngRetVal As Long
    lngRetVal = URLDownloadToFile(0, URL, LocalFilename, 0, 0)
    If lngRetVal = 0 Then DownloadFile = True
End Function
