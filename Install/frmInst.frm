VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "DM Tiny Downloader Manager - Setup"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdfolder 
      Caption         =   "...."
      Height          =   345
      Left            =   4800
      TabIndex        =   6
      Top             =   1380
      Width           =   555
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   390
      Left            =   1620
      TabIndex        =   5
      Top             =   1890
      Width           =   1215
   End
   Begin VB.CommandButton cmdinstall 
      Caption         =   "Install"
      Height          =   390
      Left            =   240
      TabIndex        =   4
      Top             =   1890
      Width           =   1215
   End
   Begin VB.TextBox txtinstall 
      Height          =   345
      Left            =   255
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1380
      Width           =   4470
   End
   Begin VB.PictureBox pictop 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   0
      ScaleHeight     =   780
      ScaleWidth      =   5460
      TabIndex        =   0
      Top             =   0
      Width           =   5460
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DM Tiny Downloader Manager Installer"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   135
         TabIndex        =   1
         Top             =   135
         Width           =   5100
      End
   End
   Begin VB.Label lbltitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Were do you want to install DM Tiny Downloader Manager"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   2
      Top             =   1035
      Width           =   4725
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   -15
      X2              =   420
      Y1              =   780
      Y2              =   780
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_NEWDIALOGSTYLE As Long = &H40
Private Const def_install_location = "C:\DownloadMan\"
'
Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Function FixPath(lzPath As String) As String
    If Right(lzPath, 1) = "\" Then FixPath = lzPath Else FixPath = lzPath & "\"
End Function

Private Function GetFolder(ByVal hWndOwner As Long, ByVal sTitle As String) As String
Dim bInf As BROWSEINFO
Dim RetVal As Long
Dim PathID As Long
Dim RetPath As String
Dim Offset As Integer
    bInf.hOwner = hWndOwner
    bInf.lpszTitle = sTitle

    bInf.ulFlags = BIF_RETURNONLYFSDIRS Or BIF_NEWDIALOGSTYLE
    PathID = SHBrowseForFolder(bInf)
    
    RetPath = Space$(512)
    RetVal = SHGetPathFromIDList(ByVal PathID, ByVal RetPath)
    If RetVal Then
        Offset = InStr(RetPath, Chr$(0))
        GetFolder = Left$(RetPath, Offset - 1)
    End If
End Function

Private Sub cmdcancel_Click()
    Unload Form1 ' unload the for,
End Sub

Private Sub cmdfolder_Click()
Dim FolName As String
    ' Promts the user to select a folder
    FolName = GetFolder(hWnd, "Please select a folder:")
    If Len(FolName) <= 0 Then
        txtinstall.Text = def_install_location
    Else
        txtinstall.Text = FixPath(FolName)
    End If
    
    FolName = ""
    
End Sub
Function FindFile(lzFile As String) As Boolean
    If Dir(lzFile) <> "" Then FindFile = True Else FindFile = False
End Function
Private Sub cmdinstall_Click()
Dim ComDll As String, HtmlPage As String
Dim MenuItemName As String, InstallDir As String
    
    InstallDir = txtinstall.Text
    
    ComDll = FixPath(App.Path) & "Download.dll" ' path to the download.dll make sure it in the install path
    HtmlPage = FixPath(App.Path) & "dload.html"
    ' path to the dload.html you must have this unless the program will not work
    ' make sure it in the install path
    
    If Not FindFile(HtmlPage) Then ' html page was not found
        MsgBox "Unable to locate the file needed for the install" _
        & vbCrLf & vbCrLf & HtmlPage & vbCrLf & vbCrLf & "The installer will now be stopped.", vbCritical, Caption
        End
        Exit Sub
    ElseIf Not FindFile(ComDll) Then ' download dll not found
        MsgBox "Unable to locate the COM object " & vbCrLf & vbCrLf & ComDll _
        & vbCrLf & vbCrLf & "The installer will now be stopped.", vbCritical, Caption
        End
    Else
        If Not Dir(InstallDir, vbDirectory) <> "" Then MkDir InstallDir
        
        
        FileCopy ComDll, txtinstall.Text & "Download.dll" ' copy the dll to the install folder
        FileCopy HtmlPage, txtinstall.Text & "dload.html" ' copy the html page to the install folder
    End If
    
    If Not RegisterActiveX(InstallDir & "Download.dll", Register) Then
        ' the download.dll failed to register
        MsgBox "Unable to regsiter the COM object " & vbCrLf & vbCrLf & ComDll _
        & vbCrLf & vbCrLf & "The installer will now be stopped.", vbCritical, Caption
        End
    Else
        MenuItemName = "&Download Links with DM Tiny Downloader Manager" ' add the menu item to IE
        SaveString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\MenuExt\" & MenuItemName, "", txtinstall.Text & "dload.html" ' add the location of the webpage
        SaveDword HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\MenuExt\" & MenuItemName, "Contexts", 243 ' add the Dword for the menu item
    End If
    ' clean up used vairables
    MenuItemName = ""
    ComDll = ""
    HtmlPage = ""
    InstallDir = ""
    MsgBox "DM Tiny Downloader Manager has now been installed to " & vbCrLf & vbCrLf & txtinstall.Text, vbInformation
    
    txtinstall.Text = ""
    End
End Sub

Private Sub Form_Load()
    Line1.X2 = Form1.Width
    txtinstall.Text = def_install_location ' default install path
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Form1 = Nothing
End Sub
