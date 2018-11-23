VERSION 5.00
Begin VB.Form About 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "About Smart Viewer Sample App"
   ClientHeight    =   4440
   ClientLeft      =   2265
   ClientTop       =   1800
   ClientWidth     =   4830
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4440
   ScaleWidth      =   4830
   Begin VB.CommandButton Done 
      Caption         =   "Done"
      Height          =   495
      Left            =   1680
      TabIndex        =   9
      Top             =   3840
      Width           =   1455
   End
   Begin VB.PictureBox SupportFrame 
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   4515
      TabIndex        =   4
      Top             =   2280
      Width           =   4575
      Begin VB.Label PhoneLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Phone:  (604) 669-8379"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1080
         TabIndex        =   8
         Top             =   735
         Width           =   2055
      End
      Begin VB.Label SupportLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "For Technical Support on Crystal Reports"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   4335
      End
      Begin VB.Label FaxLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fax:      (604) 681-7163"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   465
         Width           =   2085
      End
      Begin VB.Label CISLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "CIS:      Go Reports"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1080
         TabIndex        =   5
         Top             =   1005
         Width           =   2055
      End
   End
   Begin VB.PictureBox AuthorFrame 
      ForeColor       =   &H00000000&
      Height          =   2055
      Left            =   120
      ScaleHeight     =   1995
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.PictureBox SSIMGPicture 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   690
         Left            =   480
         ScaleHeight     =   660
         ScaleWidth      =   3585
         TabIndex        =   10
         Top             =   120
         Width           =   3615
      End
      Begin VB.Label DescriptionLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sample Crystal Smart Viewer Application"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   960
         Width           =   3645
      End
      Begin VB.Label VersionLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Version 1.0"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1680
         TabIndex        =   2
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label AuthorLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Written By:  Chris Murray - Seagate Software "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   1545
         Width           =   4005
      End
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Done_Click()
    Unload Me
End Sub
