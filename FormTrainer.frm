VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gym Management System"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15705
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   15705
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      DataField       =   "Gender"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   6240
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   7680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   4080
      Top             =   7680
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      RecordSource    =   "Select * from Trainer"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3000
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.OptionButton OptFemale 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Female"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6480
      TabIndex        =   16
      Top             =   4680
      Width           =   1935
   End
   Begin VB.OptionButton OptMale 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Male"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4200
      TabIndex        =   15
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CommandButton CmdDelete 
      BackColor       =   &H008080FF&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6960
      Width           =   1935
   End
   Begin VB.CommandButton CmdSave 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6960
      Width           =   1935
   End
   Begin VB.CommandButton CmdUpload 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "Upload"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5640
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   4200
      TabIndex        =   11
      Top             =   5520
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Fax"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   161218561
      CurrentDate     =   43267
   End
   Begin VB.TextBox TxtSalary 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   10
      Top             =   6360
      Width           =   4215
   End
   Begin VB.TextBox TxtAge 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   9
      Top             =   3840
      Width           =   4215
   End
   Begin VB.TextBox TxtName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   8
      Top             =   3000
      Width           =   4215
   End
   Begin VB.TextBox TxtId 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   2160
      Width           =   4215
   End
   Begin VB.Image ImgPhoto 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   3255
      Left            =   10920
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080C0FF&
      Caption         =   "Trainer Name"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   20.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   3000
      Width           =   3255
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "Age"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   20.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   3840
      Width           =   3255
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080C0FF&
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   20.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   4680
      Width           =   3255
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080C0FF&
      Caption         =   "Joining Date"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   20.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   5520
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Salary"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   20.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   6360
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Trainer ID"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   20.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   2160
      Width           =   3255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add Trainers"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   36
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   810
      Left            =   5760
      TabIndex        =   0
      Top             =   360
      Width           =   4800
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   8775
      Left            =   0
      Picture         =   "FormTrainer.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15735
   End
End
Attribute VB_Name = "Form3"
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
        .Fields(4).Value = DTPicker1.Value
        .Fields(5).Value = TxtSalary.Text
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
    Adodc1.Recordset.Fields(6).Value = st.Read
    st.Close
End Sub
Public Sub ResetData()
    TxtName.Text = ""
    TxtAge.Text = ""
    OptMale.Value = True
    TxtSalary.Text = ""
    DTPicker1.Value = Date
    ImgPhoto.Picture = LoadPicture("")
End Sub

Private Sub CmdUpload_Click()
CommonDialog1.Filter = "jpeg|*.jpg"
CommonDialog1.ShowOpen
ImgPhoto.Picture = LoadPicture(CommonDialog1.FileName)
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
    Adodc1.RecordSource = "select * from Trainer where TrainerId =" & Val(TxtId.Text)
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
    DTPicker1.Value = Adodc1.Recordset.Fields(4).Value
    TxtSalary.Text = Adodc1.Recordset.Fields(5).Value
    LoadImageFile
End Sub

Public Sub LoadImageFile()
    Dim st As ADODB.Stream
    Set st = New ADODB.Stream
    st.Type = adTypeBinary
    st.Open
    If IsNull(Adodc1.Recordset.Fields(6).Value) = False Then
        st.Write (Adodc1.Recordset.Fields(6).Value)
        st.SaveToFile App.Path & "\aa.jpg", adSaveCreateOverWrite
        ImgPhoto.Picture = LoadPicture(App.Path & "\aa.jpg")
    End If
    st.Close
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
