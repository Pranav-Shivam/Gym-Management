VERSION 5.00
Begin VB.Form Form5 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GYM Management System"
   ClientHeight    =   8370
   ClientLeft      =   2610
   ClientTop       =   1425
   ClientWidth     =   16200
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   16200
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   8415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      Begin VB.CommandButton CmdBuy 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Buy Supplements"
         BeginProperty Font 
            Name            =   "Harrington"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   6840
         Width           =   2655
      End
      Begin VB.CommandButton CmdEquipment 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Equipments"
         BeginProperty Font 
            Name            =   "Harrington"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5520
         Width           =   2655
      End
      Begin VB.CommandButton CmdTrainer 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Trainers"
         BeginProperty Font 
            Name            =   "Harrington"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   4200
         Width           =   2655
      End
      Begin VB.CommandButton CmdMember 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Members"
         BeginProperty Font 
            Name            =   "Harrington"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Image ImgSmall 
         Height          =   2415
         Left            =   0
         Picture         =   "MainMenu.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3015
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gym Management System"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   720
      Left            =   5520
      TabIndex        =   5
      Top             =   720
      Width           =   8085
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H0080C0FF&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   4920
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   9375
   End
   Begin VB.Image ImgBig 
      Height          =   8415
      Left            =   3000
      Picture         =   "MainMenu.frx":333B7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13215
   End
   Begin VB.Menu mnuReport 
      Caption         =   "&Reports"
      Begin VB.Menu mnuReportEquipment 
         Caption         =   "Equipment Report"
      End
      Begin VB.Menu mnuReportMember 
         Caption         =   "Member Report"
      End
      Begin VB.Menu mnuReportTrainer 
         Caption         =   "Trainer Report"
      End
      Begin VB.Menu mnuReportFee 
         Caption         =   "Fee Report"
      End
   End
   Begin VB.Menu mnuFees 
      Caption         =   "Fees"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBuy_Click()
Form9.Show
End Sub

Private Sub CmdEquipment_Click()
Form11.Show
End Sub

Private Sub CmdMember_Click()
Form6.Show
End Sub

Private Sub CmdTrainer_Click()
Form3.Show
End Sub

Private Sub mnuAbout_Click()
Form7.Show
End Sub

Private Sub mnuExit_Click()
Form13.Show
Unload Me
End Sub

Private Sub mnuFees_Click()
Form15.Show
End Sub

Private Sub mnuReportEquipment_Click()
Form12.Show
End Sub

Private Sub mnuReportFee_Click()
Form17.Show
End Sub

Private Sub mnuReportMember_Click()
Form8.Show
End Sub

Private Sub mnuReportTrainer_Click()
Form14.Show
End Sub
