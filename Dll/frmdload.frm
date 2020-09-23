VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmdload 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DM Tiny Downloader Manager"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkall 
      Caption         =   "Check all Links for Download"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1785
      TabIndex        =   8
      Top             =   1665
      Width           =   5010
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   360
      Left            =   45
      TabIndex        =   7
      Top             =   6045
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   1
   End
   Begin Download.DmDownload DmDownload1 
      Left            =   585
      Top             =   6870
      _ExtentX        =   423
      _ExtentY        =   423
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6480
      TabIndex        =   6
      Top             =   6750
      Width           =   1020
   End
   Begin VB.CommandButton cmdownload 
      Caption         =   "Download Now"
      Height          =   375
      Left            =   4995
      TabIndex        =   5
      Top             =   6750
      Width           =   1335
   End
   Begin VB.CommandButton cmdfolder 
      Caption         =   "...."
      Height          =   360
      Left            =   6315
      TabIndex        =   4
      Top             =   4830
      Width           =   435
   End
   Begin MSComctlLib.ListView lstV 
      Height          =   2535
      Left            =   45
      TabIndex        =   3
      Top             =   1935
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2670
      Top             =   6630
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdload.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtsave 
      Height          =   330
      Left            =   45
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   4830
      Width           =   6135
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1425
      Left            =   0
      ScaleHeight     =   1425
      ScaleWidth      =   7635
      TabIndex        =   11
      Top             =   0
      Width           =   7635
      Begin VB.Image Image1 
         Height          =   1440
         Left            =   4740
         Picture         =   "frmdload.frx":0352
         Top             =   0
         Width           =   2625
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Website Page Title:"
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
         Left            =   105
         TabIndex        =   19
         Top             =   75
         Width           =   1575
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TILE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   1995
         TabIndex        =   18
         Top             =   60
         Width           =   345
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Website Page URL:"
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
         Left            =   90
         TabIndex        =   17
         Top             =   405
         Width           =   1530
      End
      Begin VB.Label lblURL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "URL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   1995
         TabIndex        =   16
         Top             =   405
         Width           =   315
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Links Found:"
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
         Left            =   105
         TabIndex        =   15
         Top             =   705
         Width           =   1065
      End
      Begin VB.Label lblFondUrls 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1995
         TabIndex        =   14
         Top             =   705
         Width           =   90
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Page Referrer:"
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
         Index           =   2
         Left            =   105
         TabIndex        =   13
         Top             =   990
         Width           =   1200
      End
      Begin VB.Label lblRef 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "URL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   1995
         TabIndex        =   12
         Top             =   990
         Width           =   315
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   6870
         Picture         =   "frmdload.frx":2476
         Top             =   45
         Width           =   720
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   7
      X1              =   0
      X2              =   6735
      Y1              =   1425
      Y2              =   1425
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   6
      X1              =   15
      X2              =   6750
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   5
      X1              =   15
      X2              =   6750
      Y1              =   6570
      Y2              =   6570
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   4
      X1              =   0
      X2              =   6735
      Y1              =   6555
      Y2              =   6555
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   780
      TabIndex        =   10
      Top             =   5505
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      Height          =   420
      Left            =   105
      TabIndex        =   9
      Top             =   5520
      Width           =   6525
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   3
      X1              =   0
      X2              =   6735
      Y1              =   5415
      Y2              =   5415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   15
      X2              =   6750
      Y1              =   5430
      Y2              =   5430
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save to Location:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   105
      TabIndex        =   1
      Top             =   4530
      Width           =   1260
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Downloads Found:"
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
      Index           =   1
      Left            =   105
      TabIndex        =   0
      Top             =   1650
      Width           =   1515
   End
End
Attribute VB_Name = "frmdload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_NEWDIALOGSTYLE As Long = &H40
Private Const def_save_location = "C:\MyDownloads\"
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

Private Sub chkall_Click()
Dim I As Integer
    For I = 1 To lstV.ListItems.Count
        lstV.ListItems(I).Checked = True
    Next
    I = 0
End Sub

Private Sub cmdcancel_Click()
    Unload frmdload
End Sub

Private Sub cmdfolder_Click()
On Error Resume Next
Dim FolName As String
    FolName = GetFolder(hWnd, "Pick your save location:")
    If Len(FolName) = 0 Then
        txtsave.Text = def_save_location
    Else
        txtsave.Text = FixPath(FolName)
    End If
    FolName = ""
End Sub

Private Sub cmdownload_Click()
Dim iCount As Long, AnythingSelected As Boolean
Dim WebCls As New main, Filename As String, FileUrl As String

    iCount = lstV.ListItems.Count ' get the list count
    ' this jusr checks to see if anything in the list is selected
    For I = 1 To iCount
        If lstV.ListItems(I).Checked Then AnythingSelected = True
    Next
    I = 0
    
    If Not AnythingSelected Then
        ' nothing was selected in the list so nothing to do
        MsgBox "There was nothing selected to download from the list.", vbInformation, Caption
        Exit Sub
    End If
        
    If Not Dir(txtsave.Text, vbDirectory) <> "" Then
        If MsgBox("The folder:" & vbCrLf & vbCrLf & txtsave.Text & vbCrLf & vbCrLf _
        & "Does not exist would you like to create it now?", vbYesNo Or vbQuestion) = vbNo Then
            Exit Sub
        Else
            MkDir txtsave.Text
        End If
    End If
    
       ' Start downloading any selected items
    For I = 1 To iCount
        If lstV.ListItems(I).Checked Then
            Filename = WebCls.GetWebFileNameTitle(lstV.ListItems(I).Key)
            FileUrl = lstV.ListItems(I).Key
            DmDownload1.DownloadFile FileUrl, txtsave.Text & Filename, vbAsyncTypeByteArray
        End If
    Next
    I = 0
    Filename = ""
    FileUrl = ""
    
End Sub

Private Sub Command1_Click()
 DmDownload1.DownloadFile "http://localhost/demo/backup2.reg", "C:\ben.reg"
 
End Sub

Private Sub DmDownload1_DownloadComplete(mCurBytes As Long, mMaxBytes As Long, LocalFile As String)
    ProgressBar1.Value = 0
    lblStatus.Caption = "Finished..."
End Sub

Private Sub DmDownload1_DownloadProgress(mCurBytes As Long, mMaxBytes As Long)
On Error Resume Next
    ProgressBar1.Max = mMaxBytes
    ProgressBar1.Value = mCurBytes
    If ProgressBar1.Value = mMaxBytes Then ProgressBar1.Value = mCurBytes
    
End Sub

Private Sub DmDownload1_Status(StatusText As String)
    lblStatus.Caption = StatusText
End Sub

Private Sub Form_Load()
    txtsave.Text = def_save_location
    Me.Icon = Nothing
    
    Line1(2).X2 = frmdload.Width
    Line1(3).X2 = frmdload.Width
    Line1(4).X2 = frmdload.Width
    Line1(5).X2 = frmdload.Width
    Line1(6).X2 = frmdload.Width
    Line1(7).X2 = frmdload.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmdload = Nothing
End Sub

