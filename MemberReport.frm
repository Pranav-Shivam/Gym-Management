VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form8 
   BackColor       =   &H00C0E0FF&
   Caption         =   "GYM Management System-Member Report"
   ClientHeight    =   8655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15840
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form8"
   ScaleHeight     =   8655
   ScaleWidth      =   15840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdDelete 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Delete Record"
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2760
      Width           =   3375
   End
   Begin VB.CommandButton CmdAll 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "View All Records"
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
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2760
      Width           =   3375
   End
   Begin VB.TextBox TxtSearch 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12000
      TabIndex        =   6
      Top             =   2760
      Width           =   2415
   End
   Begin VB.CommandButton CmdGo 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Caption         =   "Go"
      Height          =   615
      Left            =   14640
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton CmdFemale 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Caption         =   "Female"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   1770
   End
   Begin VB.CommandButton CmdMale 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Caption         =   "Male"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "MemberReport.frx":0000
      Height          =   4575
      Left            =   360
      TabIndex        =   1
      Top             =   3600
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   8070
      _Version        =   393216
      BackColor       =   8454143
      ForeColor       =   16711680
      HeadLines       =   1
      RowHeight       =   24
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdodcGrid 
      Height          =   375
      Left            =   360
      Top             =   240
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      RecordSource    =   "Select * from Member"
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
   Begin VB.Image Image1 
      Height          =   1935
      Left            =   13080
      Picture         =   "MemberReport.frx":0018
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search by Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   12000
      TabIndex        =   7
      Top             =   2280
      Width           =   2505
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF80FF&
      BackStyle       =   0  'Transparent
      Caption         =   "  Gender Wise Report"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   2280
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "    Member Report"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   4680
      TabIndex        =   0
      Top             =   600
      Width           =   6255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   4560
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   6735
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim confirm As Integer

Private Sub CmdAll_Click()
AdodcGrid.RecordSource = "select * from Member"
AdodcGrid.Refresh
AdodcGrid.Caption = AdodcGrid.RecordSource
End Sub

Private Sub CmdDelete_Click()
confirm = MsgBox("Do you want to delete the Record", vbYesNo + vbExclamation, "Warning Message")
If confirm = vbYes Then
AdodcGrid.Recordset.Delete
MsgBox "Record Deleted Successfully", vbInformation, "Success"
Else
MsgBox "Record not Deleted", vbInformation, "Failure"
End If
End Sub

Private Sub CmdFemale_Click()
AdodcGrid.RecordSource = "select * from Member where Gender='Female'"
AdodcGrid.Refresh
AdodcGrid.Caption = AdodcGrid.RecordSource
End Sub

Private Sub CmdGo_Click()
On Error Resume Next
AdodcGrid.RecordSource = "Select * from Member where MemberName='" + TxtSearch.Text + "'"
AdodcGrid.Refresh
If AdodcGrid.Recordset.EOF Then
MsgBox "Data Not Found", vbCritical, "Message"
Else
AdodcGrid.Caption = AdodcGrid.RecordSource
End If
End Sub


Private Sub CmdMale_Click()
AdodcGrid.RecordSource = "select * from Member where Gender='Male'"
AdodcGrid.Refresh
AdodcGrid.Caption = AdodcGrid.RecordSource
End Sub

Private Sub TxtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
CmdGo.SetFocus
End If
End Sub
