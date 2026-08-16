VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmControl 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control Information"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmControl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.UpDown UpDownNo 
      Height          =   375
      Index           =   7
      Left            =   4320
      TabIndex        =   38
      Top             =   2400
      Width           =   240
      _ExtentX        =   344
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownNo 
      Height          =   375
      Index           =   2
      Left            =   2400
      TabIndex        =   37
      Top             =   2400
      Width           =   240
      _ExtentX        =   344
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownNo 
      Height          =   375
      Index           =   6
      Left            =   6360
      TabIndex        =   36
      Top             =   3960
      Width           =   240
      _ExtentX        =   344
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownNo 
      Height          =   375
      Index           =   3
      Left            =   6480
      TabIndex        =   35
      Top             =   2400
      Width           =   240
      _ExtentX        =   344
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownNo 
      Height          =   375
      Index           =   5
      Left            =   4320
      TabIndex        =   34
      Top             =   3960
      Width           =   240
      _ExtentX        =   344
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownNo 
      Height          =   375
      Index           =   4
      Left            =   2520
      TabIndex        =   33
      Top             =   3960
      Width           =   240
      _ExtentX        =   344
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownNo 
      Height          =   375
      Index           =   1
      Left            =   3600
      TabIndex        =   32
      Top             =   1200
      Width           =   240
      _ExtentX        =   344
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownNo 
      Height          =   375
      Index           =   0
      Left            =   3600
      TabIndex        =   31
      Top             =   600
      Width           =   240
      _ExtentX        =   344
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.Frame Frame1 
      Caption         =   "Record Operation"
      Height          =   855
      Left            =   0
      TabIndex        =   23
      Top             =   5160
      Width           =   7095
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Add"
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Edit"
         Default         =   -1  'True
         Height          =   495
         Index           =   2
         Left            =   1200
         TabIndex        =   29
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Save"
         Height          =   495
         Index           =   3
         Left            =   2160
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Undo"
         Height          =   495
         Index           =   4
         Left            =   3120
         TabIndex        =   27
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Delete"
         Height          =   495
         Index           =   5
         Left            =   4080
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Search"
         Height          =   495
         Index           =   6
         Left            =   5160
         TabIndex        =   25
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdRecord 
         Cancel          =   -1  'True
         Caption         =   "Retur&n"
         Height          =   495
         Index           =   7
         Left            =   6120
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox txtNo 
      DataField       =   "CANC_DEDUC_3"
      DataSource      =   "MyAdodc"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   7
      Left            =   5160
      MaxLength       =   8
      TabIndex        =   15
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox txtNo 
      DataField       =   "CANC_DEDUC_6"
      DataSource      =   "MyAdodc"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   6
      Left            =   3120
      MaxLength       =   8
      TabIndex        =   13
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox txtNo 
      DataField       =   "CANC_DEDUC_12"
      DataSource      =   "MyAdodc"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   5
      Left            =   1200
      MaxLength       =   8
      TabIndex        =   11
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox txtNo 
      DataField       =   "ECO_BAG_LIMIT"
      DataSource      =   "MyAdodc"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   4
      Left            =   5160
      MaxLength       =   8
      TabIndex        =   9
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox txtNo 
      DataField       =   "BUS_BAG_LIMIT"
      DataSource      =   "MyAdodc"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   3
      Left            =   3120
      MaxLength       =   8
      TabIndex        =   7
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtNo 
      DataField       =   "FIRST_BAG_LIMIT"
      DataSource      =   "MyAdodc"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   2
      Left            =   1200
      MaxLength       =   8
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtNo 
      DataField       =   "EXCESS_BAG_CHG"
      DataSource      =   "MyAdodc"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   2400
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtNo 
      DataField       =   "AIR_TAX"
      DataSource      =   "MyAdodc"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   2400
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc MyAdodc 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      Top             =   5610
      Visible         =   0   'False
      Width           =   7065
      _ExtentX        =   12462
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
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\AirlineApp\AirReservation.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\AirlineApp\AirReservation.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM Control"
      Caption         =   "Control Info"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Image Image3 
      Height          =   585
      Left            =   0
      Picture         =   "frmControl.frx":0442
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image Image2 
      Height          =   600
      Left            =   6360
      Picture         =   "frmControl.frx":0884
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image imgMoveRecord 
      Height          =   315
      Index           =   0
      Left            =   1080
      Picture         =   "frmControl.frx":0CC6
      ToolTipText     =   "First Record"
      Top             =   4560
      Width           =   315
   End
   Begin VB.Label lblMoveRecord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Home"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   1380
      TabIndex        =   16
      ToolTipText     =   "First Record"
      Top             =   4710
      Width           =   885
   End
   Begin VB.Image imgMoveRecord 
      Height          =   315
      Index           =   1
      Left            =   2520
      Picture         =   "frmControl.frx":0D64
      ToolTipText     =   "Previous Record"
      Top             =   4560
      Width           =   315
   End
   Begin VB.Label lblMoveRecord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Up"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   17
      ToolTipText     =   "Previous Record"
      Top             =   4740
      Width           =   885
   End
   Begin VB.Image imgMoveRecord 
      Height          =   315
      Index           =   2
      Left            =   4080
      Picture         =   "frmControl.frx":0E02
      ToolTipText     =   "Next Record"
      Top             =   4560
      Width           =   315
   End
   Begin VB.Label lblMoveRecord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Next"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   4380
      TabIndex        =   18
      ToolTipText     =   "Next Record"
      Top             =   4710
      Width           =   675
   End
   Begin VB.Image imgMoveRecord 
      Height          =   315
      Index           =   3
      Left            =   5520
      Picture         =   "frmControl.frx":0EA0
      ToolTipText     =   "Last Record"
      Top             =   4560
      Width           =   315
   End
   Begin VB.Label lblMoveRecord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&End"
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   3
      Left            =   5820
      TabIndex        =   19
      ToolTipText     =   "Last Record"
      Top             =   4710
      Width           =   825
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Cancellation Deduction"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   21
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   975
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   3480
      Width           =   6375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Bag Limit "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   840
      TabIndex        =   20
      Top             =   1800
      Width           =   1665
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   1095
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   6375
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3 Days"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Index           =   7
      Left            =   5130
      TabIndex        =   14
      Top             =   3720
      Width           =   525
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "6 Days"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Index           =   6
      Left            =   3450
      TabIndex        =   12
      Top             =   3720
      Width           =   525
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12 Days"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Index           =   5
      Left            =   1200
      TabIndex        =   10
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Economy Class"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   4
      Left            =   5160
      TabIndex        =   8
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Business Class"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Index           =   3
      Left            =   3120
      TabIndex        =   6
      Top             =   2160
      Width           =   1410
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First Class"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Index           =   2
      Left            =   1200
      TabIndex        =   4
      Top             =   2160
      Width           =   885
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Excess Bag Charges:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Index           =   1
      Left            =   720
      TabIndex        =   2
      Top             =   1200
      Width           =   1725
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Air Tax:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      Top             =   600
      Width           =   615
   End
   Begin VB.Shape shpMoveRecord 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   345
      Index           =   3
      Left            =   5520
      Shape           =   2  'Oval
      Top             =   4710
      Width           =   1155
   End
   Begin VB.Shape shpMoveRecord 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   345
      Index           =   2
      Left            =   4080
      Shape           =   2  'Oval
      Top             =   4710
      Width           =   1155
   End
   Begin VB.Shape shpMoveRecord 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   345
      Index           =   1
      Left            =   2520
      Shape           =   2  'Oval
      Top             =   4710
      Width           =   1155
   End
   Begin VB.Shape shpMoveRecord 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   345
      Index           =   0
      Left            =   1110
      Shape           =   2  'Oval
      Top             =   4710
      Width           =   1155
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Control Information"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   495
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   7995
   End
End
Attribute VB_Name = "frmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRecord_Click(Index As Integer)
    Record Index
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " on " & Date
End Sub

'Private Sub Form_Resize()
'    Image1.Height = Me.Height
'    Image1.Width = Me.Width
'End Sub
Private Sub Form_Unload(Cancel As Integer)
    CloseForm
End Sub

Private Sub imgMoveRecord_Click(Index As Integer)
    MoveRecord Index, MyAdodc
End Sub

Private Sub lblMoveRecord_Click(Index As Integer)
    MoveRecord Index, MyAdodc
End Sub

Private Sub MyAdodc_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    lblTitle.Caption = "Control (" & MyAdodc.Recordset.AbsolutePosition & "/" _
        & MyAdodc.Recordset.RecordCount & ")'s Information"
End Sub

Private Sub txtNo_KeyPress(Index As Integer, KeyAscii As Integer)
    If (Chr(KeyAscii) >= "0" And Chr(KeyAscii) <= "9") _
        Or Chr(KeyAscii) = vbBack Then Exit Sub
    Beep
    KeyAscii = 0
End Sub

Private Sub UpDownNo_DownClick(Index As Integer)
    If Val(txtNo(Index).Text) = 0 Then Exit Sub
    txtNo(Index).Text = Val(txtNo(Index).Text) - 1
End Sub

Private Sub UpDownNo_UpClick(Index As Integer)
    txtNo(Index).Text = Val(txtNo(Index).Text) + 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyChecker KeyCode
End Sub

