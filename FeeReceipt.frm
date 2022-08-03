VERSION 5.00
Begin VB.Form Form16 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receipt"
   ClientHeight    =   7980
   ClientLeft      =   4065
   ClientTop       =   1125
   ClientWidth     =   11610
   LinkTopic       =   "Form16"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   11610
   Begin VB.CommandButton CmdPrint 
      BackColor       =   &H0080FF80&
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6720
      Width           =   2055
   End
   Begin VB.Label LblHide4 
      BackColor       =   &H0080FFFF&
      Height          =   855
      Left            =   7080
      TabIndex        =   17
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label LblHide3 
      BackColor       =   &H0080FFFF&
      Height          =   855
      Left            =   10560
      TabIndex        =   16
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label LblHide2 
      BackColor       =   &H0080FFFF&
      Height          =   615
      Left            =   7080
      TabIndex        =   15
      Top             =   3960
      Width           =   3975
   End
   Begin VB.Label LblHide1 
      BackColor       =   &H0080FFFF&
      Height          =   735
      Left            =   6840
      TabIndex        =   14
      Top             =   5280
      Width           =   4215
   End
   Begin VB.Image Image3 
      Height          =   2055
      Left            =   7320
      Picture         =   "FeeReceipt.frx":0000
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   3495
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Left            =   360
      Picture         =   "FeeReceipt.frx":14E9F
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   3615
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Signature of Manager"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   600
      TabIndex        =   13
      Top             =   7440
      Width           =   3165
   End
   Begin VB.Label LblDate 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   840
      TabIndex        =   12
      Top             =   3840
      Width           =   2535
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "worth Rupees"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6720
      TabIndex        =   11
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label LblPaidfor 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   3240
      TabIndex        =   10
      Top             =   3240
      Width           =   3255
   End
   Begin VB.Label LblFee 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   9000
      TabIndex        =   9
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Label LblId 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   8400
      TabIndex        =   8
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Label LblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   1560
      TabIndex        =   7
      Top             =   2520
      Width           =   3855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   1065
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID Number"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6120
      TabIndex        =   5
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "has paid the fee of"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   4
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "on"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   3
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The amount has been paid by the customer."
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   2
      Top             =   4560
      Width           =   6600
   End
   Begin VB.Image Image1 
      Height          =   1935
      Left            =   9240
      Picture         =   "FeeReceipt.frx":214D6
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label LblSubhead 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gold's Gym"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   27.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   645
      Left            =   3600
      TabIndex        =   1
      Top             =   1200
      Width           =   3240
   End
   Begin VB.Label LblHead 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GYM Management System"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   27.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   630
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   7200
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdPrint_Click()
CmdPrint.Visible = False
PrintForm
CmdPrint.Visible = True
End Sub

Private Sub Form_Load()
LblName.Caption = Form15.CmbMembername.Text
LblId.Caption = Form15.CmbMemberid.Text
LblPaidfor.Caption = Form15.TxtPaid.Text
LblFee.Caption = Form15.TxtAmount.Text
LblDate.Caption = Form15.DTPicker1.Value
End Sub
