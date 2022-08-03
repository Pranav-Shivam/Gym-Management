VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form6 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Members"
   ClientHeight    =   8460
   ClientLeft      =   2055
   ClientTop       =   1485
   ClientWidth     =   15780
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   15780
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3000
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   720
      Top             =   960
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      RecordSource    =   "select * from Member"
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
   Begin VB.TextBox Text1 
      DataField       =   "Age"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3480
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton CmdDelete 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6960
      Width           =   2055
   End
   Begin VB.CommandButton CmdSave 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6960
      Width           =   2055
   End
   Begin VB.CommandButton CmdUpload 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Upload Photo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5760
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   465
      Left            =   4080
      TabIndex        =   16
      Top             =   7200
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   820
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   12640511
      CalendarTitleBackColor=   16744576
      Format          =   161415169
      CurrentDate     =   43265
   End
   Begin VB.TextBox TxtFees 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4080
      TabIndex        =   15
      Top             =   6480
      Width           =   4575
   End
   Begin VB.TextBox TxtAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4080
      TabIndex        =   14
      Top             =   5760
      Width           =   4575
   End
   Begin VB.TextBox TxtMobile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4080
      TabIndex        =   13
      Top             =   5040
      Width           =   4575
   End
   Begin VB.OptionButton OptFemale 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Female"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   6480
      TabIndex        =   12
      Top             =   4320
      Width           =   2175
   End
   Begin VB.OptionButton OptMale 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Male"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   4080
      TabIndex        =   11
      Top             =   4320
      Width           =   2175
   End
   Begin VB.TextBox TxtAge 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4080
      TabIndex        =   10
      Top             =   3600
      Width           =   4575
   End
   Begin VB.TextBox TxtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4080
      TabIndex        =   9
      Top             =   2880
      Width           =   4575
   End
   Begin VB.TextBox TxtId 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4080
      TabIndex        =   8
      Top             =   2160
      Width           =   4575
   End
   Begin VB.Image ImgPhoto 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   3375
      Left            =   10440
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   3855
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Add Members"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   840
      Left            =   5400
      TabIndex        =   17
      Top             =   600
      Width           =   4650
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   1215
      Left            =   4800
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   6255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Joining Date"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   360
      TabIndex        =   7
      Top             =   7200
      Width           =   3225
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Fees"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   360
      TabIndex        =   6
      Top             =   6480
      Width           =   3225
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   360
      TabIndex        =   5
      Top             =   5760
      Width           =   3225
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Mobile No."
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   360
      TabIndex        =   4
      Top             =   5040
      Width           =   3225
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   360
      TabIndex        =   3
      Top             =   4320
      Width           =   3225
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Age"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   360
      TabIndex        =   2
      Top             =   3600
      Width           =   3225
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Member Name"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   360
      TabIndex        =   1
      Top             =   2880
      Width           =   3225
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Member ID"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   360
      TabIndex        =   0
      Top             =   2160
      Width           =   3225
   End
   Begin VB.Image Image1 
      Height          =   8535
      Left            =   0
      Picture         =   "Member.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15825
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CmdDelete_Click()
Adodc1.Recordset.Delete
Adodc1.Refresh
TxtId.Text = ""
ResetData
MsgBox "Record deleted successfully!!", vbInformation + vbOKOnly, "Deleted"
TxtId.SetFocus
End Sub

Private Sub CmdSave_Click()
With Adodc1.Recordset
        .Fields(0).Value = TxtId.Text
        .Fields(1).Value = TxtName.Text
        .Fields(2).Value = TxtAge.Text
        If OptMale.Value = True Then
            .Fields(3).Value = "Male"
        Else
            .Fields(3).Value = "Female"
        End If
        .Fields(4).Value = TxtMobile.Text
        .Fields(5).Value = TxtAddress.Text
        .Fields(6).Value = TxtFees.Text
        .Fields(7).Value = DTPicker1.Value
    End With
    SavePicture
    Adodc1.Recordset.Update
    Adodc1.Refresh
    MsgBox "Record Saved Successfully!!", vbInformation + vbOKOnly, "Saved Successfully"
     TxtId.Text = ""
    ResetData
    TxtId.SetFocus
End Sub
Public Sub SavePicture()
    Dim st As ADODB.Stream
    Set st = New ADODB.Stream
    st.Type = adTypeBinary
    st.Open
    st.LoadFromFile (CommonDialog1.FileName)
    Adodc1.Recordset.Fields(8).Value = st.Read
    st.Close
End Sub
Public Sub ResetData()
    TxtName.Text = ""
    TxtAge.Text = ""
    OptMale.Value = True
    TxtMobile.Text = ""
    TxtAddress.Text = ""
    TxtFees.Text = ""
    DTPicker1.Value = Date
    ImgPhoto.Picture = LoadPicture("")
End Sub

Private Sub CmdUpload_Click()
CommonDialog1.Filter = "jpeg|*.jpg"
CommonDialog1.ShowOpen
ImgPhoto.Picture = LoadPicture(CommonDialog1.FileName)
End Sub



Private Sub TxtAddress_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If TxtAddress.Text = "" Then
            MsgBox "Please enter Address", vbCritical, "Address please !!!"
        Else
            TxtFees.SetFocus
        End If
End If
End Sub

Private Sub TxtAge_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If TxtAge.Text = "" Then
            MsgBox "Please enter Age", vbCritical, "Age please !!!"
        Else
            OptMale.SetFocus
        End If
End If
End Sub

Private Sub TxtFees_Change()
If KeyAscii = vbKeyReturn Then
        If TxtFees.Text = "" Then
            MsgBox "Please enter Fees", vbCritical, "Fees Missing !!!"
        Else
            DTPicker1.SetFocus
        End If
End If
End Sub

Private Sub TxtId_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Val(TxtId.Text) = 0 Then
            MsgBox "Please Enter ID ", vbCritical + vbOKOnly, "Enter ID"
            TxtId.SetFocus
            Exit Sub
        Else
            CheckData
            TxtName.SetFocus
        End If
    End If
End Sub

Public Sub CheckData()
    Adodc1.RecordSource = "select * from Member where MemberId =" & Val(TxtId.Text)
    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount = 1 Then
        loadData
    Else
        ResetData
        Adodc1.Recordset.AddNew
    End If
    
End Sub

Private Sub TxtId_LostFocus()
If Val(TxtId.Text) = 0 Then
            MsgBox "Plz. Enter ID ", vbCritical + vbOKOnly, "Enter ID"
            TxtId.SetFocus
            Exit Sub
        Else
            CheckData
            TxtName.SetFocus
        End If
End Sub
Public Sub loadData()
    TxtName.Text = Adodc1.Recordset.Fields(1).Value
    TxtAge.Text = Adodc1.Recordset.Fields(2).Value
    If Adodc1.Recordset.Fields(3).Value = "Male" Then
        OptMale.Value = True
    Else
        OptFemale.Value = True
    End If
    TxtMobile.Text = Adodc1.Recordset.Fields(4).Value
    TxtAddress.Text = Adodc1.Recordset.Fields(5).Value
    TxtFees.Text = Adodc1.Recordset.Fields(6).Value
    DTPicker1.Value = Adodc1.Recordset.Fields(7).Value
    LoadImageFile
    
End Sub
Public Sub LoadImageFile()
    Dim st As ADODB.Stream
    Set st = New ADODB.Stream
    st.Type = adTypeBinary
    st.Open
    If IsNull(Adodc1.Recordset.Fields(8).Value) = False Then
        st.Write (Adodc1.Recordset.Fields(8).Value)
        st.SaveToFile App.Path & "\ks.jpg", adSaveCreateOverWrite
        ImgPhoto.Picture = LoadPicture(App.Path & "\ks.jpg")
    End If
    st.Close
End Sub



Private Sub TxtMobile_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyReturn Then
        If TxtMobile.Text = "" Then
            MsgBox "Please enter Mobile No.", vbCritical, "Number please !!!"
        Else
            TxtAddress.SetFocus
        End If
End If
End Sub

Private Sub TxtName_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If TxtName.Text = "" Then
            MsgBox "Please enter Name.", vbCritical, "Name please !!!"
        Else
            TxtAge.SetFocus
        End If
End If
End Sub
