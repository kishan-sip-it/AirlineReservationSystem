VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAirBus 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Air Bus Information"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAirBus.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   5670
   Begin MSComCtl2.UpDown UpDownClass 
      Height          =   375
      Index           =   5
      Left            =   5040
      TabIndex        =   34
      Top             =   3360
      Width           =   240
      _ExtentX        =   344
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownClass 
      Height          =   375
      Index           =   4
      Left            =   3360
      TabIndex        =   33
      Top             =   3360
      Width           =   240
      _ExtentX        =   344
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownClass 
      Height          =   375
      Index           =   3
      Left            =   1920
      TabIndex        =   32
      Top             =   3360
      Width           =   240
      _ExtentX        =   344
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownClass 
      Height          =   375
      Index           =   2
      Left            =   5040
      TabIndex        =   31
      Top             =   2160
      Width           =   240
      _ExtentX        =   344
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownClass 
      Height          =   375
      Index           =   1
      Left            =   3360
      TabIndex        =   30
      Top             =   2160
      Width           =   240
      _ExtentX        =   344
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownClass 
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   29
      Top             =   2160
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
      TabIndex        =   28
      Top             =   4680
      Width           =   5655
      Begin VB.CommandButton cmdRecord 
         Cancel          =   -1  'True
         Caption         =   "Retur&n"
         Height          =   495
         Index           =   7
         Left            =   4800
         TabIndex        =   27
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Search"
         Height          =   495
         Index           =   6
         Left            =   3960
         TabIndex        =   26
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Delete"
         Height          =   495
         Index           =   5
         Left            =   3120
         TabIndex        =   25
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Undo"
         Height          =   495
         Index           =   4
         Left            =   2400
         TabIndex        =   24
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Save"
         Height          =   495
         Index           =   3
         Left            =   1680
         TabIndex        =   23
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Edit"
         Default         =   -1  'True
         Height          =   495
         Index           =   2
         Left            =   840
         TabIndex        =   22
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Add"
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.TextBox txtClass 
      DataField       =   "ECO_WL_CAP"
      DataSource      =   "MyAdodc"
      Height          =   375
      Index           =   5
      Left            =   4200
      TabIndex        =   13
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox txtClass 
      DataField       =   "BUS_WL_CAP"
      DataSource      =   "MyAdodc"
      Height          =   375
      Index           =   4
      Left            =   2520
      TabIndex        =   11
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox txtClass 
      DataField       =   "FIRST_WL_CAP"
      DataSource      =   "MyAdodc"
      Height          =   375
      Index           =   3
      Left            =   1080
      TabIndex        =   9
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox txtClass 
      DataField       =   "ECO_CAP"
      DataSource      =   "MyAdodc"
      Height          =   375
      Index           =   2
      Left            =   4200
      TabIndex        =   7
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox txtClass 
      DataField       =   "BUS_CAP"
      DataSource      =   "MyAdodc"
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   5
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox txtClass 
      DataField       =   "FIRST_CAP"
      DataSource      =   "MyAdodc"
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   3
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox txtAIRBUSNO 
      DataField       =   "AIRBUSNO"
      DataSource      =   "MyAdodc"
      Height          =   315
      Left            =   2475
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   1
      Top             =   1080
      Width           =   825
   End
   Begin MSAdodcLib.Adodc MyAdodc 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      Top             =   5100
      Visible         =   0   'False
      Width           =   5670
      _ExtentX        =   10001
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
      RecordSource    =   "SELECT * FROM AirBus"
      Caption         =   "AirBus"
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
   Begin VB.Image imgMoveRecord 
      Height          =   315
      Index           =   0
      Left            =   120
      Picture         =   "frmAirBus.frx":0E42
      ToolTipText     =   "First Record"
      Top             =   4020
      Width           =   315
   End
   Begin VB.Label lblMoveRecord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&End"
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   3
      Left            =   4860
      TabIndex        =   20
      ToolTipText     =   "Last Record"
      Top             =   4170
      Width           =   825
   End
   Begin VB.Image imgMoveRecord 
      Height          =   315
      Index           =   3
      Left            =   4560
      Picture         =   "frmAirBus.frx":0EE0
      ToolTipText     =   "Last Record"
      Top             =   4020
      Width           =   315
   End
   Begin VB.Shape shpMoveRecord 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   345
      Index           =   3
      Left            =   4560
      Shape           =   2  'Oval
      Top             =   4170
      Width           =   1155
   End
   Begin VB.Label lblMoveRecord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Next"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   3420
      TabIndex        =   19
      ToolTipText     =   "Next Record"
      Top             =   4170
      Width           =   675
   End
   Begin VB.Image imgMoveRecord 
      Height          =   315
      Index           =   2
      Left            =   3120
      Picture         =   "frmAirBus.frx":0F7E
      ToolTipText     =   "Next Record"
      Top             =   4020
      Width           =   315
   End
   Begin VB.Label lblMoveRecord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Up"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   18
      ToolTipText     =   "Previous Record"
      Top             =   4200
      Width           =   885
   End
   Begin VB.Image imgMoveRecord 
      Height          =   315
      Index           =   1
      Left            =   1560
      Picture         =   "frmAirBus.frx":101C
      ToolTipText     =   "Previous Record"
      Top             =   4020
      Width           =   315
   End
   Begin VB.Shape shpMoveRecord 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   345
      Index           =   1
      Left            =   1560
      Shape           =   2  'Oval
      Top             =   4170
      Width           =   1155
   End
   Begin VB.Label lblMoveRecord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Home"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   420
      TabIndex        =   17
      ToolTipText     =   "First Record"
      Top             =   4170
      Width           =   885
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Air Bus Information"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   495
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   5700
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Waiting List Capacity:"
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
      Left            =   360
      TabIndex        =   15
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   975
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   2880
      Width           =   5370
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Booking Capacity:"
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
      Left            =   360
      TabIndex        =   14
      Top             =   1560
      Width           =   2190
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   975
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   5370
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E&conomy Class"
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
      Left            =   4200
      TabIndex        =   12
      Top             =   3120
      Width           =   1260
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bu&siness Class"
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
      Left            =   2565
      TabIndex        =   10
      Top             =   3120
      Width           =   1290
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F&irst Class"
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
      Left            =   1125
      TabIndex        =   8
      Top             =   3120
      Width           =   885
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
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
      Index           =   3
      Left            =   4245
      TabIndex        =   6
      Top             =   1920
      Width           =   1260
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Business Class"
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
      Left            =   2520
      TabIndex        =   4
      Top             =   1920
      Width           =   1290
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&First Class"
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
      Left            =   1080
      TabIndex        =   2
      Top             =   1920
      Width           =   885
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Air Bus No:"
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
      Left            =   1560
      TabIndex        =   0
      Top             =   1080
      Width           =   900
   End
   Begin VB.Shape shpMoveRecord 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   345
      Index           =   0
      Left            =   150
      Shape           =   2  'Oval
      Top             =   4170
      Width           =   1155
   End
   Begin VB.Shape shpMoveRecord 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   345
      Index           =   2
      Left            =   3120
      Shape           =   2  'Oval
      Top             =   4170
      Width           =   1155
   End
   Begin VB.Image Image1 
      Height          =   4935
      Left            =   -240
      Stretch         =   -1  'True
      Top             =   480
      Width           =   7215
   End
End
Attribute VB_Name = "frmAirBus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRecord_Click(Index As Integer)
    Record Index
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyChecker KeyCode
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
    lblTitle.Caption = "Air Bus (" & MyAdodc.Recordset.AbsolutePosition & "/" & MyAdodc.Recordset.RecordCount _
       & ") Information"
End Sub
Private Sub txtClass_KeyPress(Index As Integer, KeyAscii As Integer)
    If (Chr(KeyAscii) >= "0" And Chr(KeyAscii) <= "9") _
        Or Chr(KeyAscii) = vbBack Then Exit Sub
    Beep
    KeyAscii = 0
    
End Sub
Private Sub UpDownClass_DownClick(Index As Integer)
    If Val(txtClass(Index).Text) = 0 Then Exit Sub
    txtClass.Item(Index).Text = Val(txtClass.Item(Index).Text) - 1
End Sub
Private Sub UpDownClass_UpClick(Index As Integer)
    txtClass.Item(Index).Text = Val(txtClass.Item(Index).Text) + 1
End Sub


