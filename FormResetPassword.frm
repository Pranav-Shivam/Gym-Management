VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form4 
   BackColor       =   &H0080FFFF&
   Caption         =   "Reset"
   ClientHeight    =   6555
   ClientLeft      =   4515
   ClientTop       =   2295
   ClientWidth     =   11310
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   17.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   ScaleHeight     =   6555
   ScaleWidth      =   11310
   Begin VB.CommandButton CmdChange 
      BackColor       =   &H008080FF&
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5160
      Width           =   3375
   End
   Begin VB.TextBox TxtConfirmnew 
      Appearance      =   0  'Flat
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   6000
      PasswordChar    =   "*"
      TabIndex        =   11
      Top             =   4200
      Width           =   3735
   End
   Begin VB.TextBox TxtNew 
      Appearance      =   0  'Flat
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   6000
      PasswordChar    =   "*"
      TabIndex        =   10
      Top             =   3360
      Width           =   3735
   End
   Begin VB.CommandButton CmdVerify 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Verify"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox TxtDate 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   4
      Top             =   1800
      Width           =   3615
   End
   Begin VB.CommandButton CmdCheck 
      BackColor       =   &H0080FF80&
      Caption         =   "Check"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox TxtUserid 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   1
      Top             =   240
      Width           =   3615
   End
   Begin MSAdodcLib.Adodc AdodcReset 
      Height          =   375
      Left            =   8280
      Top             =   6000
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=dsngym"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "dsngym"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * from Login"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label LblConfirm 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm New Password"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   4320
      Width           =   3855
   End
   Begin VB.Label LblNew 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter New Password"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   3480
      Width           =   3855
   End
   Begin VB.Label LblMsg1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4200
      TabIndex        =   7
      Top             =   2640
      Width           =   195
   End
   Begin VB.Label LblDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Date of Birth"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Label LblMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4140
      TabIndex        =   3
      Top             =   1080
      Width           =   15
   End
   Begin VB.Label LblUserid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter User ID"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdChange_Click()
If TxtNew.Text = TxtConfirmnew.Text Then
AdodcReset.Recordset.Fields("Password") = TxtConfirmnew.Text
AdodcReset.Recordset.Update
MsgBox "Password Changed Successfully", vbInformation, "Success"
Form4.Hide
Form2.Show
Else
MsgBox "Password Did Not Match,Please Enter Correct Details", vbExclamation, "Failure"
TxtNew.Text = ""
TxtConfirmnew = ""
End If
End Sub

Private Sub CmdCheck_Click()
AdodcReset.RecordSource = "Select * from Login where Username='" + TxtUserid.Text + " ' "
AdodcReset.Refresh
If AdodcReset.Recordset.EOF Then
LblMsg.Caption = "User ID Not Found ...Sorry Cannot Reset Password !!!"
LblMsg.ForeColor = &HFF&
Else
LblMsg.Caption = "User ID found in the Gym's Database"
LblMsg.ForeColor = &H8000&
TxtDate.SetFocus
End If
End Sub



Private Sub CmdVerify_Click()
On Error Resume Next
Dim str As String
str = StrComp(AdodcReset.Recordset.Fields("DOB").Value, TxtDate.Text, vbTextCompare)
If str = True Then
LblMsg1.Caption = "Account not verified,Can't reset the password"
LblMsg.Caption = "Sorry...Date of Birth Did Not Match !!!"
LblMsg.ForeColor = &HFF&
LblMsg1.ForeColor = &HFF&
Else
LblMsg.ForeColor = &H8000&
LblMsg1.ForeColor = &H8000&
LblMsg.Caption = "Congratulations !!!"
LblMsg1.Caption = "Account is verified , Set your new Password"
TxtNew.Visible = True
TxtConfirmnew.Visible = True
CmdChange.Visible = True
LblNew.Visible = True
LblConfirm.Visible = True
TxtNew.SetFocus
End If

End Sub

Private Sub Form_Load()
TxtNew.Visible = False
TxtConfirmnew.Visible = False
CmdChange.Visible = False
LblNew.Visible = False
LblConfirm.Visible = False
End Sub

Private Sub TxtConfirmnew_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If TxtConfirmnew.Text = "" Then
            MsgBox "Please enter Password.", vbCritical, "Password please !!!"
        Else
            CmdChange.SetFocus
        End If
End If
End Sub

Private Sub TxtDate_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If TxtDate.Text = "" Then
            MsgBox "Please enter Date of Birth", vbCritical, "DOB Please !!!"
        Else
            CmdVerify.SetFocus
        End If
End If
End Sub

Private Sub TxtNew_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If TxtNew.Text = "" Then
            MsgBox "Please enter Password.", vbCritical, "Password please !!!"
        Else
            TxtConfirmnew.SetFocus
        End If
End If
End Sub

Private Sub TxtUserid_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If TxtUserid.Text = "" Then
            MsgBox "Please enter UserID.", vbCritical, "UserID please !!!"
        Else
            CmdCheck.SetFocus
        End If
End If
End Sub
