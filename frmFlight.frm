VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFlight 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Flight Information"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFlight.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.UpDown UpDownBooking 
      Height          =   375
      Index           =   2
      Left            =   5880
      TabIndex        =   26
      Top             =   3000
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownBooking 
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   25
      Top             =   3000
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownBooking 
      Height          =   375
      Index           =   0
      Left            =   2040
      TabIndex        =   24
      Top             =   3000
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      DataField       =   "FLIGHTDATE"
      DataSource      =   "MyAdodc"
      Height          =   375
      Left            =   1920
      TabIndex        =   23
      Top             =   1320
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   58523648
      CurrentDate     =   37553
   End
   Begin VB.Frame Frame1 
      Caption         =   "Record Operation"
      Height          =   855
      Left            =   360
      TabIndex        =   15
      Top             =   4320
      Width           =   5655
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Add"
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Edit"
         Default         =   -1  'True
         Height          =   495
         Index           =   2
         Left            =   840
         TabIndex        =   21
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Save"
         Height          =   495
         Index           =   3
         Left            =   1680
         TabIndex        =   20
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Undo"
         Height          =   495
         Index           =   4
         Left            =   2400
         TabIndex        =   19
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Delete"
         Height          =   495
         Index           =   5
         Left            =   3120
         TabIndex        =   18
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Search"
         Height          =   495
         Index           =   6
         Left            =   3960
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdRecord 
         Cancel          =   -1  'True
         Caption         =   "Retur&n"
         Height          =   495
         Index           =   7
         Left            =   4800
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox txtBooking 
      DataField       =   "ECO_SEATS_BK"
      DataMember      =   "Command2"
      DataSource      =   "MyAdodc"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   2
      Left            =   4440
      MaxLength       =   8
      TabIndex        =   8
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox txtBooking 
      DataField       =   "BUS_SEATS_BK"
      DataMember      =   "Command2"
      DataSource      =   "MyAdodc"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   2520
      MaxLength       =   8
      TabIndex        =   6
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox txtBooking 
      DataField       =   "FIRST_SEATS_BK"
      DataMember      =   "Command2"
      DataSource      =   "MyAdodc"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   600
      MaxLength       =   8
      TabIndex        =   4
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox txtFLIGHTNO 
      DataField       =   "FLIGHTNO"
      DataSource      =   "MyAdodc"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1815
      MaxLength       =   4
      TabIndex        =   1
      Top             =   825
      Width           =   1020
   End
   Begin MSAdodcLib.Adodc MyAdodc 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      ToolTipText     =   "Move To Specified Record"
      Top             =   4710
      Visible         =   0   'False
      Width           =   6345
      _ExtentX        =   11192
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
      RecordSource    =   "SELECT * FROM FLIGHT"
      Caption         =   "Fare Information"
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
      Picture         =   "frmFlight.frx":0E42
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image Image2 
      Height          =   600
      Left            =   5760
      Picture         =   "frmFlight.frx":1284
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image imgMoveRecord 
      Height          =   315
      Index           =   1
      Left            =   1800
      Picture         =   "frmFlight.frx":16C6
      ToolTipText     =   "Previous Record"
      Top             =   3720
      Width           =   315
   End
   Begin VB.Image imgMoveRecord 
      Height          =   315
      Index           =   2
      Left            =   3360
      Picture         =   "frmFlight.frx":1764
      ToolTipText     =   "Next Record"
      Top             =   3720
      Width           =   315
   End
   Begin VB.Image imgMoveRecord 
      Height          =   315
      Index           =   3
      Left            =   4800
      Picture         =   "frmFlight.frx":1802
      ToolTipText     =   "Last Record"
      Top             =   3720
      Width           =   315
   End
   Begin VB.Image imgMoveRecord 
      Height          =   315
      Index           =   0
      Left            =   360
      Picture         =   "frmFlight.frx":18A0
      ToolTipText     =   "First Record"
      Top             =   3720
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "&Seats Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   13
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Economy Class"
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
      Index           =   4
      Left            =   4440
      TabIndex        =   7
      Top             =   2760
      Width           =   1260
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Business Class:"
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
      Left            =   2520
      TabIndex        =   5
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&First Class:"
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
      Left            =   480
      TabIndex        =   3
      Top             =   2760
      Width           =   930
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   1215
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   2400
      Width           =   5895
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Flight &Date:"
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
      Left            =   885
      TabIndex        =   2
      Top             =   1440
      Width           =   900
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Flight &No:"
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
      Left            =   1035
      TabIndex        =   0
      Top             =   870
      Width           =   750
   End
   Begin VB.Label lblMoveRecord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Home"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   660
      TabIndex        =   9
      ToolTipText     =   "First Record"
      Top             =   3870
      Width           =   885
   End
   Begin VB.Label lblMoveRecord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Up"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   10
      ToolTipText     =   "Previous Record"
      Top             =   3900
      Width           =   885
   End
   Begin VB.Label lblMoveRecord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Next"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   3660
      TabIndex        =   11
      ToolTipText     =   "Next Record"
      Top             =   3870
      Width           =   675
   End
   Begin VB.Label lblMoveRecord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&End"
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   3
      Left            =   5100
      TabIndex        =   12
      ToolTipText     =   "Last Record"
      Top             =   3870
      Width           =   825
   End
   Begin VB.Shape shpMoveRecord 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   345
      Index           =   3
      Left            =   4800
      Shape           =   2  'Oval
      Top             =   3870
      Width           =   1155
   End
   Begin VB.Shape shpMoveRecord 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   345
      Index           =   2
      Left            =   3360
      Shape           =   2  'Oval
      Top             =   3870
      Width           =   1155
   End
   Begin VB.Shape shpMoveRecord 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   345
      Index           =   1
      Left            =   1800
      Shape           =   2  'Oval
      Top             =   3870
      Width           =   1155
   End
   Begin VB.Shape shpMoveRecord 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   345
      Index           =   0
      Left            =   390
      Shape           =   2  'Oval
      Top             =   3870
      Width           =   1155
   End
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   -120
      Stretch         =   -1  'True
      Top             =   480
      Width           =   6660
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Flight Information"
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
      Height          =   615
      Left            =   -120
      TabIndex        =   14
      Top             =   0
      Width           =   6915
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7800
   End
End
Attribute VB_Name = "frmFlight"
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
    lblTitle.Caption = "Flight (" & MyAdodc.Recordset.AbsolutePosition & "/" _
        & MyAdodc.Recordset.RecordCount & ")'s Information"
End Sub

Private Sub txtBooking_KeyPress(Index As Integer, KeyAscii As Integer)
    If (Chr(KeyAscii) >= "0" And Chr(KeyAscii) <= "9") _
        And Chr(KeyAscii) = vbBack Then Exit Sub
    Beep
    KeyAscii = 0
End Sub



Private Sub UpDownBooking_DownClick(Index As Integer)
    If Val(txtBooking(Index).Text) = 0 Then Exit Sub
    txtBooking(Index).Text = Val(txtBooking(Index).Text) - 1
End Sub

Private Sub UpDownBooking_UpClick(Index As Integer)
    txtBooking(Index).Text = Val(txtBooking(Index).Text) + 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyChecker KeyCode
End Sub

