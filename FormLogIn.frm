VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   8235
   ClientLeft      =   3480
   ClientTop       =   1650
   ClientWidth     =   13095
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   13095
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      Caption         =   "Show Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   315
      Left            =   9120
      TabIndex        =   8
      Top             =   5880
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc AdodcLogin 
      Height          =   495
      Left            =   5280
      Top             =   1200
      Visible         =   0   'False
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   873
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
      RecordSource    =   "select * from Login"
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
   Begin VB.CommandButton CmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Cancel"
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6360
      Width           =   1445
   End
   Begin VB.CommandButton CmdReset 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Reset Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7200
      Width           =   3120
   End
   Begin VB.CommandButton CmdLogin 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Login"
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
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6360
      Width           =   1445
   End
   Begin VB.TextBox TxtPassword 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   5520
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   5280
      Width           =   5295
   End
   Begin VB.TextBox TxtUsername 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      DataSource      =   "AdodcLogin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   1
      Top             =   3840
      Width           =   5295
   End
   Begin VB.Image Image2 
      Height          =   3135
      Left            =   1920
      Picture         =   "FormLogIn.frx":0000
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   3375
   End
   Begin VB.Label LblLogin 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Gym Management Login"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   810
      Left            =   2040
      TabIndex        =   7
      Top             =   2040
      Width           =   8955
   End
   Begin VB.Label LblPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5520
      TabIndex        =   2
      Top             =   4560
      Width           =   2535
   End
   Begin VB.Label LblUsername 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   5520
      TabIndex        =   0
      Top             =   3120
      Width           =   2505
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BackStyle       =   1  'Opaque
      Height          =   5055
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   3000
      Width           =   9735
   End
   Begin VB.Image Image1 
      Height          =   10815
      Left            =   -1320
      Picture         =   "FormLogIn.frx":A3A9
      Stretch         =   -1  'True
      Top             =   -2520
      Width           =   14415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
If Check1.Value = 0 Then
    TxtPassword.PasswordChar = "*"
    Else
    TxtPassword.PasswordChar = ""
End If
End Sub

Private Sub CmdCancel_Click()
End
End Sub

Private Sub CmdLogin_Click()
On Error Resume Next
AdodcLogin.RecordSource = "select * from Login where Username='" + TxtUsername.Text + "'and Password='" + TxtPassword.Text + "'"
AdodcLogin.Refresh
If AdodcLogin.Recordset.EOF Then
MsgBox "Login Failed,try again...!!!", vbCritical, "Incorrect Username or Password"
Else
MsgBox "Login Successful", vbInformation, "Success"
Form5.Show
Unload Me
End If
End Sub

Private Sub CmdReset_Click()
Form4.Show
End Sub

Private Sub TxtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If TxtPassword.Text = "" Then
            MsgBox "Please enter Password.", vbCritical, "Password please !!!"
        Else
            CmdLogin.SetFocus
        End If
End If
End Sub

Private Sub TxtUsername_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If TxtUsername.Text = "" Then
            MsgBox "Please enter Username.", vbCritical, "Username please !!!"
        Else
            TxtPassword.SetFocus
        End If
End If
End Sub
