VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmService 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Service Information"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmService.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Record Operation"
      Height          =   855
      Left            =   0
      TabIndex        =   17
      Top             =   3480
      Width           =   5655
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Add"
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Edit"
         Default         =   -1  'True
         Height          =   495
         Index           =   2
         Left            =   840
         TabIndex        =   23
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Save"
         Height          =   495
         Index           =   3
         Left            =   1680
         TabIndex        =   22
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Undo"
         Height          =   495
         Index           =   4
         Left            =   2400
         TabIndex        =   21
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Delete"
         Height          =   495
         Index           =   5
         Left            =   3120
         TabIndex        =   20
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Search"
         Height          =   495
         Index           =   6
         Left            =   3960
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdRecord 
         Cancel          =   -1  'True
         Caption         =   "Retur&n"
         Height          =   495
         Index           =   7
         Left            =   4800
         TabIndex        =   18
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox txtSS_CODE 
      DataField       =   "SS_CODE"
      DataSource      =   "MyAdodc"
      Height          =   315
      Left            =   1455
      MaxLength       =   2
      TabIndex        =   1
      Top             =   1320
      Width           =   690
   End
   Begin VB.TextBox txtSS_DESC 
      DataField       =   "SS_DESC"
      DataSource      =   "MyAdodc"
      Height          =   315
      Left            =   1455
      MaxLength       =   15
      TabIndex        =   3
      Top             =   1800
      Width           =   4035
   End
   Begin VB.TextBox txtSS_FARE 
      DataField       =   "SS_FARE"
      DataSource      =   "MyAdodc"
      Height          =   315
      Left            =   1455
      MaxLength       =   5
      TabIndex        =   5
      Top             =   2280
      Width           =   825
   End
   Begin VB.CommandButton cmdDeleteREcord 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   3360
      TabIndex        =   15
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdSaveRecord 
      Caption         =   "&Save"
      Height          =   495
      Left            =   3360
      TabIndex        =   14
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "&Return"
      Height          =   495
      Left            =   2040
      TabIndex        =   13
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdUndoRecord 
      Caption         =   "&Undo"
      Height          =   495
      Left            =   720
      TabIndex        =   12
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdEditRecord 
      Caption         =   "&Edit"
      Height          =   495
      Left            =   2040
      TabIndex        =   11
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdNewRecord 
      Caption         =   "&New"
      Height          =   495
      Left            =   720
      TabIndex        =   10
      Top             =   4200
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc MyAdodc 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   4005
      Visible         =   0   'False
      Width           =   6120
      _ExtentX        =   10795
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
      RecordSource    =   "SELECT * FROM SERVICE"
      Caption         =   "Service Info"
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
      Height          =   600
      Left            =   0
      Picture         =   "frmService.frx":0442
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   5280
      Picture         =   "frmService.frx":0884
      Stretch         =   -1  'True
      Top             =   0
      Width           =   720
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Service Information "
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
      TabIndex        =   25
      Top             =   0
      Width           =   6075
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Service Information"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   16
      Top             =   840
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   1695
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   5295
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Code"
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
      Left            =   1020
      TabIndex        =   0
      Top             =   1365
      Width           =   435
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Description:"
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
      Left            =   480
      TabIndex        =   2
      Top             =   1800
      Width           =   990
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Fare:"
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
      Left            =   1050
      TabIndex        =   4
      Top             =   2280
      Width           =   405
   End
   Begin VB.Label lblMoveRecord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Home"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   420
      TabIndex        =   6
      ToolTipText     =   "First Record"
      Top             =   3150
      Width           =   885
   End
   Begin VB.Image imgMoveRecord 
      Height          =   315
      Index           =   1
      Left            =   1560
      Picture         =   "frmService.frx":0CC6
      ToolTipText     =   "Previous Record"
      Top             =   3000
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
      TabIndex        =   7
      ToolTipText     =   "Previous Record"
      Top             =   3180
      Width           =   885
   End
   Begin VB.Image imgMoveRecord 
      Height          =   315
      Index           =   2
      Left            =   3120
      Picture         =   "frmService.frx":0D64
      ToolTipText     =   "Next Record"
      Top             =   3000
      Width           =   315
   End
   Begin VB.Label lblMoveRecord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Next"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   3420
      TabIndex        =   8
      ToolTipText     =   "Next Record"
      Top             =   3150
      Width           =   675
   End
   Begin VB.Image imgMoveRecord 
      Height          =   315
      Index           =   3
      Left            =   4560
      Picture         =   "frmService.frx":0E02
      ToolTipText     =   "Last Record"
      Top             =   3000
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
      TabIndex        =   9
      ToolTipText     =   "Last Record"
      Top             =   3150
      Width           =   825
   End
   Begin VB.Image imgMoveRecord 
      Height          =   315
      Index           =   0
      Left            =   120
      Picture         =   "frmService.frx":0EA0
      ToolTipText     =   "First Record"
      Top             =   3000
      Width           =   315
   End
   Begin VB.Shape shpMoveRecord 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   345
      Index           =   3
      Left            =   4560
      Shape           =   2  'Oval
      Top             =   3150
      Width           =   1155
   End
   Begin VB.Shape shpMoveRecord 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   345
      Index           =   2
      Left            =   3120
      Shape           =   2  'Oval
      Top             =   3150
      Width           =   1155
   End
   Begin VB.Shape shpMoveRecord 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   345
      Index           =   1
      Left            =   1560
      Shape           =   2  'Oval
      Top             =   3150
      Width           =   1155
   End
   Begin VB.Shape shpMoveRecord 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   345
      Index           =   0
      Left            =   150
      Shape           =   2  'Oval
      Top             =   3150
      Width           =   1155
   End
End
Attribute VB_Name = "frmService"
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
Private Sub MyAdodc_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    lblTitle.Caption = "Service Schedule (" & MyAdodc.Recordset.AbsolutePosition & "/" _
        & MyAdodc.Recordset.RecordCount & ")'s Information"
End Sub

Private Sub imgMoveRecord_Click(Index As Integer)
    MoveRecord Index, MyAdodc
End Sub

Private Sub lblMoveRecord_Click(Index As Integer)
    MoveRecord Index, MyAdodc
End Sub

Private Sub txtSS_CODE_KeyPress(KeyAscii As Integer)
    UpperControl KeyAscii
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyChecker KeyCode
End Sub
