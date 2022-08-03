VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5130
   ClientLeft      =   5130
   ClientTop       =   2745
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   9300
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   6405
      Left            =   -240
      Picture         =   "FormSplash.frx":0000
      ScaleHeight     =   6405
      ScaleWidth      =   9675
      TabIndex        =   0
      Top             =   -240
      Width           =   9675
      Begin VB.Timer Timer1 
         Interval        =   180
         Left            =   840
         Top             =   960
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Loading..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   4
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label percentage 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   7800
         TabIndex        =   3
         Top             =   3720
         Width           =   495
      End
      Begin VB.Shape ShpProgress 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FFFF&
         Height          =   495
         Left            =   960
         Shape           =   4  'Rounded Rectangle
         Top             =   4080
         Width           =   15
      End
      Begin VB.Shape Shape1 
         Height          =   495
         Left            =   960
         Shape           =   4  'Rounded Rectangle
         Top             =   4080
         Width           =   7275
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Version 6.0"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   8640
         TabIndex        =   2
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         Caption         =   "GYM MANAGEMENT      APPLICATION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   1095
         Left            =   4440
         TabIndex        =   1
         Top             =   1800
         Width           =   5055
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ctr, ctr2, r As Double
Dim ctr3 As String

Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
End Sub

Private Sub Timer1_Timer()
ctr = 0
If ctr2 <= 100 Then
Randomize
    r = Int((200 - 100 + 1) * Rnd + 250)
    ctr = r / 72
    ctr = Round(ctr, 0)
    ctr = ctr2 + ctr
    ctr3 = str(ctr)
        If ctr >= 100 Then
        percentage.Caption = "100%"
        ShpProgress.Width = 7700
        Form2.Show
        Unload Me
        Else
        ShpProgress.Width = ShpProgress.Width + r
        percentage.Caption = (ctr3) + "%"
        ctr2 = Int(ctr3)
        End If
    End If
End Sub
