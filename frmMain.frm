VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Programmer Kit 2.0"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10830
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   10830
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picStar 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3095
      Picture         =   "frmMain.frx":08E2
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   10
      Top             =   100
      Width           =   240
   End
   Begin VB.PictureBox picLogo 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   120
      Picture         =   "frmMain.frx":0C6C
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   9
      Top             =   120
      Width           =   480
   End
   Begin VB.CommandButton btnAbout 
      Caption         =   "About"
      Height          =   375
      Left            =   5400
      TabIndex        =   8
      Top             =   4440
      Width           =   1215
   End
   Begin MSComctlLib.ImageList imlTreeSubs 
      Left            =   5640
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":154E
            Key             =   "Useful"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18EA
            Key             =   "sound"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C86
            Key             =   "inet"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2222
            Key             =   "stray"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27BE
            Key             =   "Registry"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B62
            Key             =   "Files"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F16
            Key             =   "MSWin"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":32B2
            Key             =   "Code"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":364E
            Key             =   "Encryption"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtExample 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2640
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2640
      Width           =   8055
   End
   Begin VB.TextBox txtDescription 
      Height          =   1215
      Left            =   2640
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Width           =   8055
   End
   Begin MSComctlLib.TreeView treeSubs 
      Height          =   5655
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   9975
      _Version        =   393217
      Indentation     =   0
      LineStyle       =   1
      Style           =   7
      SingleSel       =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "This is constantly updated, so please check back to where you downloaded it from to see if there's an update!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   11
      Top             =   4920
      Width           =   8055
   End
   Begin VB.Label Label2 
      Caption         =   "Programmer Kit made by Jacob Grice."
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   4440
      Width           =   2775
   End
   Begin VB.Label lblNote 
      Caption         =   "*Note: Any sub/function with a           for an icon means it's very useful code."
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   120
      Width           =   6255
   End
   Begin VB.Label lblExample 
      Caption         =   "Example:"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   2400
      Width           =   3495
   End
   Begin VB.Label lblDescription 
      Caption         =   "Description:"
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label lblSubsFunctions 
      Caption         =   "Subs & Functions:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      UseMnemonic     =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnAbout_Click()
Load frmAbout
frmAbout.Show
End Sub

Private Sub Form_Load()
    modSubs.ShowIcon Me
    treeSubs.ImageList = imlTreeSubs
    treeSubs.Nodes.Add , , "files", "Files", "Files", "Files"
        treeSubs.Nodes.Add "files", tvwChild, "FileExists", "FileExists", "Useful", "Useful"
        treeSubs.Nodes.Add "files", tvwChild, "CreateFile", "CreateFile", "Files", "Files"
        treeSubs.Nodes.Add "files", tvwChild, "ReadFileToString", "ReadFileToString", "Files", "Files"
        treeSubs.Nodes.Add "files", tvwChild, "DelFile", "DelFile", "Files", "Files"
        
    treeSubs.Nodes.Add , , "registry", "Registry", "Registry", "Registry"
        treeSubs.Nodes.Add "registry", tvwChild, "KeyInReg", "KeyInReg", "Useful", "Useful"
        treeSubs.Nodes.Add "registry", tvwChild, "GetKeyValue", "GetKeyValue", "Registry", "Registry"
        treeSubs.Nodes.Add "registry", tvwChild, "SetKeyValue", "SetKeyValue", "Registry", "Registry"
        treeSubs.Nodes.Add "registry", tvwChild, "DeleteKey", "DeleteKey", "Registry", "Registry"
        treeSubs.Nodes.Add "registry", tvwChild, "DeleteSection", "DeleteSection", "Registry", "Registry"
        treeSubs.Nodes.Add "registry", tvwChild, "DeleteAppRegEntries", "DeleteAppRegEntries", "Registry", "Registry"
        
    treeSubs.Nodes.Add , , "encryption", "Encryption", "Encryption", "Encryption"
        treeSubs.Nodes.Add "encryption", tvwChild, "Encrypt", "Encrypt", "Encryption", "Encryption"
        treeSubs.Nodes.Add "encryption", tvwChild, "Decrypt", "Decrypt", "Encryption", "Encryption"
        
    treeSubs.Nodes.Add , , "mswindows", "Windows", "MSWin", "MSWin"
        treeSubs.Nodes.Add "mswindows", tvwChild, "GetDir", "GetDir", "Useful", "Useful"
        treeSubs.Nodes.Add "mswindows", tvwChild, "DisableCtrlAltDel", "EnableCtrlAltDel", "MSWin", "MSWin"
        treeSubs.Nodes.Add "mswindows", tvwChild, "EnableCtrlAltDel", "EnableCtrlAltDel", "MSWin", "MSWin"
        treeSubs.Nodes.Add "mswindows", tvwChild, "AlwaysOnTop", "AlwaysOnTop", "MSWin", "MSWin"
        treeSubs.Nodes.Add "mswindows", tvwChild, "NotAlwaysOnTop", "NotAlwaysOnTop", "MSWin", "MSWin"
        treeSubs.Nodes.Add "mswindows", tvwChild, "MoveMouse", "MoveMouse", "MSWin", "MSWin"
        treeSubs.Nodes.Add "mswindows", tvwChild, "MousePos", "MousePos", "MSWin", "MSWin"
        
    treeSubs.Nodes.Add , , "systray", "System Tray", "stray", "stray"
        treeSubs.Nodes.Add "systray", tvwChild, "ShowIcon", "ShowIcon", "Useful", "Useful"
        treeSubs.Nodes.Add "systray", tvwChild, "RemoveIcon", "RemoveIcon", "Useful", "Useful"
        treeSubs.Nodes.Add "systray", tvwChild, "ChangeIcon", "ChangeIcon", "stray", "stray"
        treeSubs.Nodes.Add "systray", tvwChild, "ChangeToolTip", "ChangeToolTip", "stray", "stray"
    
    treeSubs.Nodes.Add , , "sound", "Sound", "sound", "sound"
        treeSubs.Nodes.Add "sound", tvwChild, "sndPlay", "sndPlay", "sound", "sound"
        
    treeSubs.Nodes.Add , , "inet", "Internet", "inet", "inet"
        treeSubs.Nodes.Add "inet", tvwChild, "IsConnected", "IsConnected", "inet", "inet"
        treeSubs.Nodes.Add "inet", tvwChild, "DownloadFile", "DownloadFile", "inet", "inet"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    modSubs.RemoveIcon Me
End Sub
Private Sub treeSubs_Click()
    txtDescription.text = vbNullString
    txtExample.text = vbNullString

    Select Case treeSubs.SelectedItem.Key
        Case "FileExists"
            txtDescription.text = "Checks if a file exists. Returns true if it does, false if it doesn't."
            txtExample.text = "If FileExists(""C:\File.txt"") =  True Then" & vbCrLf & _
            "   [Do Something]" & vbCrLf & "Else" & vbCrLf & _
            "   [Do Something Else]" & vbCrLf & "End If"
            
        Case "CreateFile"
            txtDescription.text = "Creates a blank file. Overwrites it if it exists and it is not read-only. Returns true if successful, false if not."
            txtExample.text = "If CreateFile(""C:\File.txt"") = True Then" & vbCrLf & _
            "   MsgBox ""File created!""" & vbCrLf & _
            "Else" & vbCrLf & _
            "   MsgBox ""File could not be created!""" & vbCrLf & "End If"
            
        Case "ReadFileToString"
            txtDescription.text = "Reads an entire file into a string. Returns whatever the string is."
            txtExample.text = "Dim str As String" & vbCrLf & "str = ReadFileToString(""C:\File.txt"")" & vbCrLf & "MsgBox str"
        Case "DelFile"
            txtDescription.text = "Attempts to delete a file. If there is an error (e.g. File doesn't exist), it returns false. If it deletes successfully, it returns true."
            txtExample.text = "Dim fileDeleted As String" & vbCrLf & "fileDeleted = DelFile(""C:\File.txt"")"
            
        Case "KeyInReg"
            txtDescription.text = "Checks if a key is in the registry settings for your application. Returns false if not, and true if it exists."
            txtExample.text = "If KeyInReg(App.EXEName, ""Options"", ""Sound On"") = True Then " & vbNewLine & _
            "   MsgBox ""  'Sound On' is a key in this app's registry settings!""" & vbNewLine & "End If"
        
        Case "GetKeyValue"
            txtDescription.text = "Gets the value of a key and returns it as a string. The value will be nothing if it doesn't exist."
            txtExample.text = "Dim UserName As String" & vbNewLine & _
            "UserName = GetKeyValue(App.EXEName, ""User Info"", ""Name"")" & vbNewLine & _
            "MsgBox ""Your name is "" & UserName"
        
        Case "SetKeyValue"
            txtDescription.text = "Sets the value of a key in the registry. If it does not exist, it creates it. If it does exist, it overwrites it. Returns true if successful."
            txtExample.text = "SetKeyValue App.EXEName, ""Last Names"", ""Jacob"", ""Grice"""
        
        Case "DeleteKey"
            txtDescription.text = "Deletes a key from the registry. Returns true if successful."
            txtExample.text = "Dim deleted As Boolean" & vbNewLine & "Deleted = " & _
            "DeleteKey(App.EXEName, ""Animals"", ""Cat"")"
        
        Case "DeleteSection"
            txtDescription.text = "Deletes a section from the registry. Returns true if successful."
            txtExample.text = "Dim deleted As Boolean" & vbNewLine & "Deleted = " & _
            "DeleteSection(App.EXEName, ""Animals"")"
        
        Case "DeleteAppRegEntries"
            txtDescription.text = "Deletes all the registry entries for an application. Returns true if successful."
            txtExample.text = "Dim deleted As Boolean" & vbNewLine & "Deleted = " & _
            "DeleteAppRegEntries(App.EXEName)"
        
        Case "Encrypt"
            txtDescription.text = "Encrypts a string of text. Returns the encrypted text."
            txtExample.text = "MsgBox Encrypt(""Jake is cool!"")"
        
        Case "Decrypt"
            txtDescription.text = "Decrypts a string of text. Returns the decrypted text."
            txtExample.text = "MsgBox Decrypt(""|ìùóRõ•Rï°°ûS"")"
        
        Case "GetDir"
            txtDescription.text = "Gets a windows directory. Directories are:     WindowsDir, SystemDir, DesktopDir, CacheDir, StartupDir, StartPrograms, and StartMenu. Returns the path to that directory."
            txtExample.text = "Dim WinDir As String" & vbNewLine & _
            "WinDir = GetDir(WindowsDir)" & vbNewLine & "MsgBox ""The windows directory is: "" & WinDir"
            
        Case "DisableCtrlAltDel"
            txtDescription.text = "Disables the Ctrl+Alt+Delete menu."
            txtExample.text = "DisableCtrlAltDel"
            
        Case "EnableCtrlAltDel"
            txtDescription.text = "Enables the Ctrl+Alt+Delete menu."
            txtExample.text = "EnableCtrlAltDel"
            
        Case "AlwaysOnTop"
            txtDescription.text = "Makes the form stay on top of everything, even when it loses focus."
            txtExample.text = "AlwaysOnTop Me"
            
        Case "NotAlwaysOnTop"
            txtDescription.text = "Makes the form normal. When it loses focus, it's not on top."
            txtExample.text = "NotAlwaysOnTop Me"
            
        Case "MoveMouse"
            txtDescription.text = "Moves the mouse to a certain place (depending on the X and Y)."
            txtExample.text = "MoveMouse 100, 325"
            
        Case "MousePos"
            txtDescription.text = "Gets the mouses position."
            txtExample.text = "Dim X As Long, Y As Long" & vbNewLine & _
            "X = MousePos().X" & vbNewLine & "Y = MousePos().Y"
            
        Case "ShowIcon"
            txtDescription.text = "Shows the icon of a form in the system tray."
            txtExample.text = "ShowIcon Me"
            
        Case "RemoveIcon"
            txtDescription.text = "Removes the icon of a form from the system tray."
            txtExample.text = "RemoveIcon Me"
            
        Case "ChangeIcon"
            txtDescription.text = "Changes the icon placed in the system tray with ShowIcon."
            txtExample.text = "ChangeIcon Me, picIcon"
            
        Case "ChangeToolTip"
            txtDescription.text = "Changes the tooltip text of an icon placed in the system tray with ShowIcon."
            txtExample.text = "ChangeToolTip Me, ""Click for a menu."""
        
        Case "sndPlay"
            txtDescription.text = "Plays a sound from a file. If the file does not exist, plays a ""ding"" sound."
            txtExample.text = "sndPlay ""C:\hello.wav"""

        Case "IsConnected"
            txtDescription.text = "Checks if the local machine is connected to the internet. Returns true if it is, false if not."
            txtExample.text = "Dim connected As Boolean" & vbNewLine & _
            "connected = IsConnected"
            
        Case "DownloadFile"
            txtDescription.text = "Downloads a file from the internet and saves it to your hard disk. Returns true if successful, false if not."
            txtExample.text = "DownloadFile ""http://www.microsoft.com/ms.htm"", ""C:\msSite.htm"""
    End Select
End Sub
