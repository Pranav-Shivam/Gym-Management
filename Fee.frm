VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form15 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gym Management System"
   ClientHeight    =   8955
   ClientLeft      =   4800
   ClientTop       =   945
   ClientWidth     =   10725
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   10725
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   8160
      Top             =   2160
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
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
      RecordSource    =   "Select * from Fee"
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   600
      Top             =   2040
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin VB.CommandButton CmdReceipt 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Receipt"
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7800
      Width           =   2175
   End
   Begin VB.CommandButton CmdDelete 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Delete"
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
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7800
      Width           =   2175
   End
   Begin VB.CommandButton CmdSave 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Save"
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
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7800
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   6600
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   163250177
      CurrentDate     =   43267
   End
   Begin VB.TextBox TxtPaid 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4440
      TabIndex        =   10
      Top             =   5880
      Width           =   3615
   End
   Begin VB.TextBox TxtAmount 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4440
      TabIndex        =   9
      Top             =   5160
      Width           =   3615
   End
   Begin VB.ComboBox CmbMembername 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4440
      TabIndex        =   8
      Text            =   "Member Name"
      Top             =   4440
      Width           =   3615
   End
   Begin VB.ComboBox CmbMemberid 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4440
      TabIndex        =   7
      Text            =   "Member Id"
      Top             =   3720
      Width           =   3615
   End
   Begin VB.TextBox TxtFeeid 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4440
      TabIndex        =   6
      Top             =   3000
      Width           =   3615
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   600
      TabIndex        =   5
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Fee Amount"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   5160
      Width           =   2295
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Member Name"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   4440
      Width           =   3015
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Member ID"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Fee Paid"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   600
      TabIndex        =   1
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Fee ID"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Image Image2 
      Height          =   735
      Left            =   3600
      Picture         =   "Fee.frx":0000
      Stretch         =   -1  'True
      Top             =   720
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FFFF&
      Height          =   1215
      Left            =   3000
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   4695
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   9015
      Left            =   0
      Picture         =   "Fee.frx":47C5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14055
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmbMemberid_Click()
CmbMembername.SetFocus

End Sub

Private Sub CmbMembername_Click()
TxtAmount.SetFocus
End Sub

Private Sub CmdDelete_Click()
Adodc2.Recordset.Delete
Adodc2.Refresh
TxtFeeid.Text = ""
ResetData
MsgBox "Record deleted successfully!!", vbInformation + vbOKOnly, "Deleted"
TxtFeeid.SetFocus
End Sub
Private Sub CmdReceipt_Click()
Form16.Show
End Sub

Private Sub CmdSave_Click()
On Error Resume Next
With Adodc2.Recordset
        .Fields(0).Value = TxtFeeid.Text
        .Fields(1).Value = CmbMemberid.Text
        .Fields(2).Value = CmbMembername.Text
        .Fields(3).Value = TxtAmount.Text
        .Fields(4).Value = TxtPaid.Text
        .Fields(5).Value = DTPicker1.Value
End With
    Adodc2.Recordset.Update
    Adodc2.Refresh
    MsgBox "Record Saved Successfully!!", vbInformation + vbOKOnly, "Saved Successfully"
    TxtFeeid.Text = ""
    ResetData
    TxtFeeid.SetFocus
End Sub
Public Sub ResetData()
    CmbMemberid.Text = "Member Id"
    CmbMembername.Text = "Member Name"
    TxtAmount.Text = ""
    TxtPaid.Text = ""
    DTPicker1.Value = Date
End Sub

Private Sub Form_Load()
Adodc1.RecordSource = "Select MemberId from Member"
Adodc1.Refresh
If Adodc1.Recordset.RecordCount > 0 Then
While Adodc1.Recordset.EOF = False
CmbMemberid.AddItem Adodc1.Recordset.Fields(0).Value
Adodc1.Recordset.MoveNext
Wend
End If

Adodc1.RecordSource = "Select MemberName from Member"
Adodc1.Refresh
If Adodc1.Recordset.RecordCount > 0 Then
While Adodc1.Recordset.EOF = False
CmbMembername.AddItem Adodc1.Recordset.Fields(0).Value
Adodc1.Recordset.MoveNext
Wend
End If
End Sub



Private Sub TxtAmount_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If TxtAmount = "" Then
    MsgBox "Plz enter Fee Amount", vbCritical, "Amount"
    TxtAmount.SetFocus
    Else
    TxtPaid.SetFocus
    End If
End If
End Sub

Private Sub TxtFeeid_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If Val(TxtFeeid.Text) = 0 Then
            MsgBox "Please Enter ID ", vbCritical + vbOKOnly, "Enter ID"
            TxtFeeid.SetFocus
            Exit Sub
        Else
            CheckData
            CmbMemberid.SetFocus
        End If
    End If
End Sub

Public Sub CheckData()
    Adodc2.RecordSource = "select * from Fee where FeeId =" & Val(TxtFeeid.Text)
    Adodc2.Refresh
    If Adodc2.Recordset.RecordCount = 1 Then
        loadData
    Else
        ResetData
        Adodc2.Recordset.AddNew
    End If
End Sub

Private Sub TxtFeeid_LostFocus()
If Val(TxtFeeid.Text) = 0 Then
            MsgBox "Plz. Enter ID ", vbCritical + vbOKOnly, "Enter ID"
            TxtFeeid.SetFocus
            Exit Sub
        Else
            CheckData
            CmbMemberid.SetFocus
        End If
End Sub

Public Sub loadData()
    CmbMemberid.Text = Adodc2.Recordset.Fields(1).Value
    CmbMembername.Text = Adodc2.Recordset.Fields(2).Value
    TxtAmount.Text = Adodc2.Recordset.Fields(3).Value
    TxtPaid.Text = Adodc2.Recordset.Fields(4).Value
    DTPicker1.Value = Adodc2.Recordset.Fields(5).Value
End Sub

Private Sub TxtPaid_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If TxtPaid = "" Then
    MsgBox "Plz enter Paid for Field", vbCritical, "Missing Data"
    TxtPaid.SetFocus
    Else
    DTPicker1.SetFocus
    End If
End If
End Sub
