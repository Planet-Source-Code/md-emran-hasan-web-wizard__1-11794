VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form f1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Web Wizard 2000"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   Icon            =   "frmWebWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "&Preview"
      Height          =   375
      Left            =   1320
      TabIndex        =   72
      Top             =   4680
      Width           =   975
   End
   Begin VB.Frame FrameAbout 
      Height          =   3735
      Left            =   240
      TabIndex        =   60
      Top             =   600
      Visible         =   0   'False
      Width           =   4215
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Shaon, Javed, Rubayet && Andalib."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   71
         Top             =   2280
         Width           =   3015
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Md Emran Hasan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   70
         Top             =   1920
         Width           =   2655
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Web : http;//www.emran.koolhost.com"
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   3240
         Width           =   2895
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail : ehasan@citechco.net"
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   3000
         Width           =   2295
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright © 2000 Emran Hasan"
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Beta Tester : "
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Programmer : "
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Freeware "
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   1480
         Width           =   1935
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Version 2.1.1"
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "An easy HTML Page Creator"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   62
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Web Wizard 2000"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   120
         TabIndex        =   61
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame FrameSave 
      Caption         =   "Save As"
      Height          =   1575
      Left            =   240
      TabIndex        =   51
      Top             =   600
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CommandButton BrowseSave 
         Caption         =   "..."
         Height          =   255
         Left            =   3720
         TabIndex        =   56
         Top             =   1080
         Width           =   255
      End
      Begin VB.TextBox txtLocation 
         Height          =   285
         Left            =   240
         TabIndex        =   55
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox txtFileName 
         Height          =   285
         Left            =   240
         TabIndex        =   53
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
         Height          =   255
         Left            =   240
         TabIndex        =   54
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "File Name"
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame FrameMETA 
      Caption         =   "META Tag"
      Height          =   2055
      Left            =   240
      TabIndex        =   46
      Top             =   2280
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox txtMetaDescription 
         Height          =   495
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   50
         Top             =   1320
         Width           =   3975
      End
      Begin VB.TextBox txtMetaKeywords 
         Height          =   495
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   48
         Top             =   480
         Width           =   3975
      End
      Begin VB.Label Label17 
         Caption         =   "Description"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label16 
         Caption         =   "Keywords"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame FrameInfo 
      Caption         =   "Page Information"
      Height          =   1575
      Left            =   240
      TabIndex        =   41
      Top             =   600
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox txtHead 
         Height          =   285
         Left            =   120
         TabIndex        =   45
         Top             =   1080
         Width           =   3855
      End
      Begin VB.TextBox txtTitle 
         Height          =   285
         Left            =   120
         TabIndex        =   43
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Title"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame FrameContact 
      Caption         =   "Contact Information"
      Height          =   3495
      Left            =   240
      TabIndex        =   32
      Top             =   600
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   120
         TabIndex        =   40
         Top             =   2880
         Width           =   3975
      End
      Begin VB.TextBox txtFax 
         Height          =   285
         Left            =   120
         TabIndex        =   38
         Top             =   2160
         Width           =   3975
      End
      Begin VB.TextBox txtPhone 
         Height          =   285
         Left            =   120
         TabIndex        =   36
         Top             =   1560
         Width           =   3975
      End
      Begin VB.TextBox txtStreet 
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Top             =   480
         Width           =   3975
      End
      Begin VB.Label Label12 
         Caption         =   "E-mail"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Faximile"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Street Address"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame FrameText 
      Caption         =   "Text of page"
      Height          =   3735
      Left            =   240
      TabIndex        =   29
      Top             =   600
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox txtText 
         Height          =   2895
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter the text you want to display in your web page below :"
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Frame FrameImage 
      Caption         =   "Image"
      Height          =   1575
      Left            =   240
      TabIndex        =   24
      Top             =   2760
      Width           =   4215
      Begin VB.TextBox txtBgImage 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   360
         TabIndex        =   27
         Top             =   720
         Width           =   3015
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton BrowseImg 
         BackColor       =   &H00C0C0C0&
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   3480
         TabIndex        =   25
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Backgound Image"
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   395
         Width           =   1335
      End
   End
   Begin VB.Frame FrameColor 
      Caption         =   "Colors"
      Height          =   2055
      Left            =   240
      TabIndex        =   15
      Top             =   600
      Width           =   4215
      Begin VB.CommandButton BrowseBg 
         Caption         =   "..."
         Height          =   255
         Left            =   2520
         TabIndex        =   19
         Top             =   220
         Width           =   255
      End
      Begin VB.CommandButton BrowseTxt 
         Caption         =   "..."
         Height          =   255
         Left            =   2520
         TabIndex        =   18
         Top             =   700
         Width           =   255
      End
      Begin VB.CommandButton BrowseLnk 
         Caption         =   "..."
         Height          =   255
         Left            =   2520
         TabIndex        =   17
         Top             =   1180
         Width           =   255
      End
      Begin VB.CommandButton BrowseVis 
         Caption         =   "..."
         Height          =   255
         Left            =   2520
         TabIndex        =   16
         Top             =   1660
         Width           =   255
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Background Color :"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   3735
      End
      Begin VB.Shape shpBg 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   1560
         Top             =   220
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Text Color : "
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   1455
      End
      Begin VB.Shape shpTxt 
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   1560
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Link Color : "
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Shape shpLnk 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   1560
         Top             =   1180
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Visited Link : "
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Shape shpVis 
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   1560
         Top             =   1660
         Width           =   855
      End
   End
   Begin VB.Frame FrameHead 
      Caption         =   "Style Elements"
      Height          =   3735
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CheckBox chkDontUnder 
         Caption         =   "Don't Underline the links"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   3095
         Width           =   2295
      End
      Begin VB.CommandButton BrowseHoverHover 
         Caption         =   "..."
         Height          =   255
         Left            =   2160
         TabIndex        =   58
         Top             =   2640
         Width           =   255
      End
      Begin VB.CheckBox chkH1 
         Caption         =   "H1 && H2"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkH5Bold 
         Caption         =   "Bold"
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   2160
         Width           =   615
      End
      Begin VB.CommandButton BrowseH5 
         Caption         =   "..."
         Height          =   255
         Left            =   1320
         TabIndex        =   12
         Top             =   2160
         Width           =   255
      End
      Begin VB.CheckBox chkH5 
         Caption         =   "H5 && H6"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkH3Bold 
         Caption         =   "Bold"
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton BrowseH3 
         Caption         =   "..."
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         Top             =   1440
         Width           =   255
      End
      Begin VB.CheckBox chkH3 
         Caption         =   "H3 && H4"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkH1Bold 
         Caption         =   "Bold"
         Height          =   255
         Left            =   1680
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton BrowseH1 
         Caption         =   "..."
         Height          =   255
         Left            =   1320
         TabIndex        =   6
         Top             =   720
         Width           =   255
      End
      Begin VB.Shape shpHoverHover 
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   1200
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Hover Color  :"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   2660
         Width           =   1095
      End
      Begin VB.Shape shpH5 
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   360
         Top             =   2160
         Width           =   855
      End
      Begin VB.Shape shpH3 
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   360
         Top             =   1440
         Width           =   855
      End
      Begin VB.Shape shpH1 
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   360
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.TextBox Text1 
      Height          =   4935
      Left            =   4800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   120
      Width           =   4815
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   0
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&About"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Build"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   4680
      Width           =   975
   End
   Begin MSComctlLib.TabStrip tb1 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   8070
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Body"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Style Sheet"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Text"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Contact"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Info"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Save As"
            ImageVarType    =   2
         EndProperty
      EndProperty
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "f1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BrowseBg_Click()
cd1.DialogTitle = "Choose background color"
cd1.ShowColor
shpBg.FillColor = cd1.Color
End Sub

Private Sub BrowseH1_Click()
cd1.DialogTitle = "Choose text color"
cd1.ShowColor
shpH1.FillColor = cd1.Color
End Sub

Private Sub BrowseH3_Click()
cd1.DialogTitle = "Choose text color"
cd1.ShowColor
shpH3.FillColor = cd1.Color
End Sub

Private Sub BrowseH5_Click()
cd1.DialogTitle = "Choose text color"
cd1.ShowColor
shpH5.FillColor = cd1.Color
End Sub

Private Sub BrowseImg_Click()
cd1.DialogTitle = "Choose background image"
cd1.DefaultExt = ".gif"
cd1.Filter = "GIF Image (*.gif)|*.gif|JPEG Image (*.jpg)|*.jpg|Bitmap Image (*.bmp)|*.bmp|"
cd1.ShowOpen
txtBgImage.Text = cd1.FileName
End Sub

Private Sub BrowseLnk_Click()
cd1.DialogTitle = "Choose link color"
cd1.ShowColor
shpLnk.FillColor = cd1.Color
End Sub

Private Sub BrowseSnd_Click()
cd1.DialogTitle = "Choose background sound file"
cd1.DefaultExt = ".wav"
cd1.Filter = "WAV Audio (*.wav)|*.wav|MIDI File (*.mid)|*.mid|"
cd1.ShowOpen
txtBgSound.Text = cd1.FileName
End Sub

Private Sub BrowseSave_Click()
Dim lpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo
'Replace 'This Is My Title' with the title you want to put on the 'Browse For Folders' dialog.
szTitle = "Select the directory..."
With tBrowseInfo
.hWndOwner = Me.hWnd
.lpszTitle = lstrcat(szTitle, "")
.ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
End With
lpIDList = SHBrowseForFolder(tBrowseInfo)
If (lpIDList) Then
sBuffer = Space(MAX_PATH)
SHGetPathFromIDList lpIDList, sBuffer
'sBuffer value is the directory that the user choose from the dialog.
sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
txtLocation.Text = sBuffer
End If
End Sub

Private Sub BrowseTxt_Click()
cd1.DialogTitle = "Choose text color"
cd1.ShowColor
shpTxt.FillColor = cd1.Color
End Sub

Private Sub BrowseVis_Click()
cd1.DialogTitle = "Choose visited link color"
cd1.ShowColor
shpVis.FillColor = cd1.Color
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
    txtBgImage.Enabled = True
    txtBgImage.BackColor = vbWhite
    BrowseImg.Enabled = True
Else
    txtBgImage.Enabled = False
    txtBgImage.BackColor = &HC0C0C0
    BrowseImg.Enabled = False
End If
End Sub



Private Sub chkH1_Click()
If chkH1.Value = 1 Then
 shpH1.FillColor = vbBlack
 BrowseH1.Enabled = True
Else
 shpH1.FillColor = &H8000000F
 BrowseH1.Enabled = False
End If
End Sub

Private Sub chkH3_Click()
If chkH3.Value = 1 Then
 shpH3.FillColor = vbBlack
 BrowseH3.Enabled = True
Else
 shpH3.FillColor = &H8000000F
 BrowseH3.Enabled = False
End If
End Sub

Private Sub chkH5_Click()
If chkH5.Value = 1 Then
 shpH5.FillColor = vbBlack
 BrowseH5.Enabled = True
Else
 shpH5.FillColor = &H8000000F
 BrowseH5.Enabled = False
End If
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Dim FileName
Dim isName
FileName = txtFileName.Text

If FileName = "" Then
 isName = 0
Else
 isName = 1
End If

Call Body
Call Text
Call Contact

Text1.Text = Text1.Text & "</body></html>"

If isName = 0 Then
 MsgBox "Please enter the filename in which all the information will be saved in the Save As section", vbInformation, "Web Wizard 2000"
Else
tempdir = txtLocation.Text
If Right$(tempdir, 1) <> "\" Then
    mainname = txtLocation.Text & "\" & FileName
Else
    mainname = txtLocation.Text & FileName
End If
Open mainname For Output As #1
Print #1, Text1.Text
Close #1
End If
End Sub


Private Sub Command3_Click()
    FrameColor.Visible = False
    FrameImage.Visible = False
    FrameHead.Visible = False
    FrameText.Visible = False
    FrameMETA.Visible = False
    FrameInfo.Visible = False
    FrameSave.Visible = False
    FrameContact.Visible = False
    FrameAbout.Visible = True
End Sub

Private Sub Command4_Click()
Dim prevFileName

prevFileName = "c:\windows\temp\tem007.htm"

Call Body
Call Text
Call Contact

Text1.Text = Text1.Text & "</body></html>"

Open prevFileName For Output As #1
Print #1, Text1.Text
Close #1

pview = Shell("c:\program files\internet explorer\iexplore.exe " & prevFileName, vbMaximizedFocus)
End Sub

Private Sub Label5_Click()
If Check1.Value = 1 Then
    Check1.Value = 0
Else
    Check1.Value = 1
End If
End Sub
Private Sub tb1_Click()
If tb1.SelectedItem.Index = 1 Then
    FrameColor.Visible = True
    FrameImage.Visible = True
    FrameHead.Visible = False
    FrameText.Visible = False
    FrameMETA.Visible = False
    FrameInfo.Visible = False
    FrameSave.Visible = False
    FrameContact.Visible = False
    FrameAbout.Visible = False
ElseIf tb1.SelectedItem.Index = 2 Then
    FrameColor.Visible = False
    FrameImage.Visible = False
    FrameHead.Visible = True
    FrameText.Visible = False
    FrameMETA.Visible = False
    FrameInfo.Visible = False
    FrameSave.Visible = False
    FrameContact.Visible = False
    FrameAbout.Visible = False
ElseIf tb1.SelectedItem.Index = 3 Then
    FrameColor.Visible = False
    FrameImage.Visible = False
    FrameHead.Visible = False
    FrameText.Visible = True
    FrameMETA.Visible = False
    FrameInfo.Visible = False
    FrameSave.Visible = False
    FrameContact.Visible = False
    FrameAbout.Visible = False
    txtText.SetFocus
ElseIf tb1.SelectedItem.Index = 4 Then
    FrameColor.Visible = False
    FrameImage.Visible = False
    FrameHead.Visible = False
    FrameText.Visible = False
    FrameMETA.Visible = False
    FrameInfo.Visible = False
    FrameSave.Visible = False
    FrameContact.Visible = True
    FrameAbout.Visible = False
ElseIf tb1.SelectedItem.Index = 5 Then
    FrameColor.Visible = False
    FrameImage.Visible = False
    FrameHead.Visible = False
    FrameText.Visible = False
    FrameMETA.Visible = True
    FrameInfo.Visible = True
    FrameSave.Visible = False
    FrameContact.Visible = False
    FrameAbout.Visible = False
ElseIf tb1.SelectedItem.Index = 6 Then
    FrameColor.Visible = False
    FrameImage.Visible = False
    FrameHead.Visible = False
    FrameText.Visible = False
    FrameMETA.Visible = False
    FrameInfo.Visible = False
    FrameSave.Visible = True
    FrameContact.Visible = False
    FrameAbout.Visible = False
End If
End Sub
