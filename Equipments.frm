VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form11 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Equipments"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15750
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   15750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdUpdate 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton CmdPrev 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton CmdNext 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton Clear 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4200
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   10680
      Top             =   7920
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      RecordSource    =   "Equipment"
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
   Begin VB.CommandButton CmdDelete 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Height          =   855
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton CmdAdd 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox TxtDate 
      DataField       =   "PurchaseDate"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   15
      Top             =   7320
      Width           =   2535
   End
   Begin VB.TextBox TxtUsed 
      DataField       =   "UsedFor"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   13
      Top             =   5880
      Width           =   4935
   End
   Begin VB.TextBox TxtQuantity 
      DataField       =   "Quantity"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   12
      Top             =   3720
      Width           =   4935
   End
   Begin VB.TextBox TxtWeight 
      DataField       =   "Weight"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   11
      Top             =   5160
      Width           =   4935
   End
   Begin VB.TextBox TxtType 
      DataField       =   "EquipmentType"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   10
      Top             =   4440
      Width           =   4935
   End
   Begin VB.TextBox TxtPrice 
      DataField       =   "Price"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   9
      Top             =   6600
      Width           =   4935
   End
   Begin VB.TextBox TxtName 
      DataField       =   "EquipmentName"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   8
      Top             =   3000
      Width           =   4935
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(dd-mm-yyyy)"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   360
      Left            =   7080
      TabIndex        =   16
      Top             =   7440
      Width           =   2280
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Date of Purchase"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   14
      Top             =   7320
      Width           =   3495
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Price"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   6600
      Width           =   3495
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Used For"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   5880
      Width           =   3495
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Quantity "
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   3720
      Width           =   3495
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Weight"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   5160
      Width           =   3495
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Equipment Type"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   4440
      Width           =   3495
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   " Equipment Name"
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
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   3000
      Width           =   3495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   " Equipments "
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   26.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   6000
      TabIndex        =   1
      Top             =   1440
      Width           =   3480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      Caption         =   "Gym Management System"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   26.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4560
      TabIndex        =   0
      Top             =   480
      Width           =   6960
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      Height          =   1095
      Left            =   4320
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   7335
   End
   Begin VB.Image Image1 
      Height          =   8775
      Left            =   0
      Picture         =   "Equipments.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15735
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Clear_Click()
TxtName.Text = ""
TxtQuantity = ""
TxtType.Text = ""
TxtUsed.Text = ""
TxtWeight.Text = ""
TxtPrice.Text = ""
TxtDate.Text = ""
End Sub

Private Sub CmdAdd_Click()
On Error Resume Next
Adodc1.Recordset.AddNew
TxtName.SetFocus
End Sub

Private Sub CmdDelete_Click()
On Error Resume Next
confirm = MsgBox("Do you want to delete this record", vbYesNo + vbCritical, "Confirmation")
If confirm = vbYes Then
Adodc1.Recordset.Delete
MsgBox "Record has been deleted Successfully", vbInformation, "Message"
Else
MsgBox "Record not deleted !!!", vbInformation, "Failure"
End If
Adodc1.Recordset.MoveFirst
End Sub

Private Sub CmdNext_Click()
On Error Resume Next
If Not Adodc1.Recordset.EOF Then
Adodc1.Recordset.MoveNext
Else
Adodc1.Recordset.MoveFirst
End If
End Sub

Private Sub CmdPrev_Click()
On Error Resume Next
If Adodc1.Recordset.BOF Then
Adodc1.Recordset.MoveLast
Else
Adodc1.Recordset.MovePrevious
End If
Adodc1.Recordset.MovePrevious
End Sub

Private Sub CmdUpdate_Click()
Adodc1.Recordset.Update
MsgBox "Records Successfuly Updated", vbInformation, "Updated"
End Sub



Private Sub TxtDate_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If TxtDate.Text = "" Then
            MsgBox "Please enter Date of Purchase", vbCritical, "Date please !!!"
        End If
End If
End Sub

Private Sub TxtName_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If TxtName.Text = "" Then
            MsgBox "Please enter Name.", vbCritical, "Name please !!!"
        Else
            TxtQuantity.SetFocus
        End If
End If
End Sub



Private Sub TxtPrice_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If TxtPrice.Text = "" Then
            MsgBox "Please enter Price", vbCritical, "Price please !!!"
        Else
            TxtDate.SetFocus
        End If
End If
End Sub

Private Sub TxtQuantity_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If TxtQuantity.Text = "" Then
            MsgBox "Please enter Quantity", vbCritical, "Quantity please !!!"
        Else
            TxtType.SetFocus
        End If
End If
End Sub

Private Sub TxtType_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If TxtType.Text = "" Then
            MsgBox "Please enter Type", vbCritical, "Type please !!!"
        Else
            TxtWeight.SetFocus
        End If
End If
End Sub

Private Sub TxtUsed_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If TxtUsed.Text = "" Then
            MsgBox "Please enter Price", vbCritical, "Price please !!!"
        Else
            TxtPrice.SetFocus
        End If
End If
End Sub

Private Sub TxtWeight_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If TxtWeight.Text = "" Then
            MsgBox "Please enter Weight", vbCritical, "Weight please !!!"
        Else
            TxtUsed.SetFocus
        End If
End If
End Sub
