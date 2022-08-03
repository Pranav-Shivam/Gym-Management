VERSION 5.00
Begin VB.Form Form9 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GYM Management System"
   ClientHeight    =   8790
   ClientLeft      =   2610
   ClientTop       =   1305
   ClientWidth     =   15780
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   15780
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdCalculate 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6720
      Width           =   2415
   End
   Begin VB.CommandButton CmdBuy 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "Buy"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox TxtQuantity 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   11400
      TabIndex        =   20
      Top             =   5760
      Width           =   1935
   End
   Begin VB.TextBox TxtTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   11400
      TabIndex        =   18
      Top             =   4320
      Width           =   1935
   End
   Begin VB.ComboBox CmbVitamin 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7800
      TabIndex        =   15
      Text            =   "Quantity"
      Top             =   7200
      Width           =   2055
   End
   Begin VB.CheckBox ChkVitamin 
      BackColor       =   &H008080FF&
      Caption         =   "MultiVitamin(60 Capsule)"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7800
      TabIndex        =   14
      Top             =   5760
      Width           =   2055
   End
   Begin VB.ComboBox CmbCreatine 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7800
      TabIndex        =   12
      Text            =   "Quantity"
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CheckBox ChkCreatine 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "Creatine Monohydrate (300 Gram)"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   7800
      TabIndex        =   11
      Top             =   3000
      Width           =   2055
   End
   Begin VB.ComboBox CmbGainer 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2760
      TabIndex        =   9
      Text            =   "Quantity"
      Top             =   7320
      Width           =   1935
   End
   Begin VB.CheckBox ChkGainer 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "Mass Gainer (5Kg)"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   2760
      TabIndex        =   8
      Top             =   5760
      Width           =   1935
   End
   Begin VB.ComboBox CmbWhey 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2760
      TabIndex        =   6
      Text            =   "Quantity"
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CheckBox ChkWhey 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "Whey Protein (1Kg)"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   2760
      TabIndex        =   5
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox TxtId 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11040
      TabIndex        =   4
      Top             =   1800
      Width           =   3495
   End
   Begin VB.TextBox TxtName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   2
      Top             =   1800
      Width           =   3495
   End
   Begin VB.Label Label9 
      BackColor       =   &H000080FF&
      Caption         =   "Total Quantity"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   11160
      TabIndex        =   19
      Top             =   5280
      Width           =   2415
   End
   Begin VB.Label Label8 
      BackColor       =   &H000080FF&
      Caption         =   "Total Amount"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   11160
      TabIndex        =   17
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackColor       =   &H008080FF&
      Caption         =   "Cost:1299"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   16
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Image Image5 
      Height          =   2535
      Left            =   5640
      Picture         =   "BuySupplement.frx":0000
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000FFFF&
      Caption         =   "Cost:899"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   13
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Image Image4 
      Height          =   2535
      Left            =   5640
      Picture         =   "BuySupplement.frx":7BB2
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000FFFF&
      Caption         =   "Cost:2599"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   10
      Top             =   7920
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H008080FF&
      Caption         =   "Cost:2999"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Image Image3 
      Height          =   2535
      Left            =   600
      Picture         =   "BuySupplement.frx":A634
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Image Image2 
      Height          =   2535
      Left            =   600
      Picture         =   "BuySupplement.frx":1239B
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Member ID"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8880
      TabIndex        =   3
      Top             =   1920
      Width           =   1590
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   8400
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   6375
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Member Name"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1560
      TabIndex        =   1
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   6375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Buy Supplements"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   24
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   540
      Left            =   5400
      TabIndex        =   0
      Top             =   360
      Width           =   4410
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   3960
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   7575
   End
   Begin VB.Image Image1 
      Height          =   8775
      Left            =   0
      Picture         =   "BuySupplement.frx":15D73
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15855
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim a As Integer, b As Integer, c As Integer, d As Integer, p As Integer, q As Integer, r As Integer, s As Integer

Private Sub CmdBuy_Click()
If TxtName.Text = "" Then
MsgBox "Enter Member Name", vbCritical, "Name Missing"
Else
Form10.Show
End If
End Sub

Private Sub CmdCalculate_Click()
On Error Resume Next
If ChkWhey.Value = 1 Then
a = CmbWhey.Text * 2999
p = CmbWhey.Text
Else
a = 0
p = 0
End If

If ChkGainer.Value = 1 Then
b = CmbGainer.Text * 2599
q = CmbGainer.Text
Else
b = 0
q = 0
End If

If ChkCreatine.Value = 1 Then
c = CmbCreatine.Text * 899
r = CmbCreatine.Text
Else
c = 0
r = 0
End If

If ChkVitamin.Value = 1 Then
d = CmbVitamin.Text * 1299
s = CmbVitamin.Text
Else
d = 0
s = 0
End If

TxtTotal.Text = a + b + c + d

TxtQuantity.Text = p + q + r + s

End Sub

Private Sub Form_Load()
CmbWhey.AddItem "1"
CmbWhey.AddItem "2"
CmbWhey.AddItem "3"
CmbWhey.AddItem "4"
CmbWhey.AddItem "5"
CmbGainer.AddItem "1"
CmbGainer.AddItem "2"
CmbGainer.AddItem "3"
CmbGainer.AddItem "4"
CmbGainer.AddItem "5"
CmbCreatine.AddItem "1"
CmbCreatine.AddItem "2"
CmbCreatine.AddItem "3"
CmbCreatine.AddItem "4"
CmbCreatine.AddItem "5"
CmbVitamin.AddItem "1"
CmbVitamin.AddItem "2"
CmbVitamin.AddItem "3"
CmbVitamin.AddItem "4"
CmbVitamin.AddItem "5"
End Sub

Private Sub TxtId_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If TxtId.Text = "" Then
            MsgBox "Please enter ID", vbCritical, "ID Missing !!!"
        Else
            ChkWhey.SetFocus
        End If
End If
End Sub

Private Sub TxtName_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If TxtName.Text = "" Then
            MsgBox "Please enter Name", vbCritical, "Name Missing !!!"
        Else
            TxtId.SetFocus
        End If
End If
End Sub
