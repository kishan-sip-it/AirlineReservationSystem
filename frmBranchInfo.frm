VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmBranchInfo 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Branch"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   Icon            =   "frmBranchInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Record Operation"
      Height          =   1455
      Left            =   0
      TabIndex        =   11
      Top             =   3960
      Width           =   4815
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Add"
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Edit"
         Default         =   -1  'True
         Height          =   495
         Index           =   2
         Left            =   1320
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Save"
         Height          =   495
         Index           =   3
         Left            =   2520
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Undo"
         Height          =   495
         Index           =   4
         Left            =   3720
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Delete"
         Height          =   495
         Index           =   5
         Left            =   480
         TabIndex        =   14
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Search"
         Height          =   495
         Index           =   6
         Left            =   1920
         TabIndex        =   13
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdRecord 
         Cancel          =   -1  'True
         Caption         =   "Retur&n"
         Height          =   495
         Index           =   7
         Left            =   3360
         TabIndex        =   12
         Top             =   840
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc MyAdodc 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   5070
      Visible         =   0   'False
      Width           =   4830
      _ExtentX        =   8520
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
      RecordSource    =   "SELECT * FROM BRANCH"
      Caption         =   "Branch"
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
   Begin VB.TextBox txtTELEPHONE 
      DataField       =   "TELEPHONE"
      DataSource      =   "MyAdodc"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   1200
      MaxLength       =   15
      TabIndex        =   10
      Top             =   3000
      Width           =   2475
   End
   Begin VB.TextBox txtCITY 
      DataField       =   "CITY"
      DataSource      =   "MyAdodc"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   1200
      MaxLength       =   25
      TabIndex        =   8
      Top             =   2595
      Width           =   3495
   End
   Begin VB.TextBox txtADDRESS2 
      DataField       =   "ADDRESS2"
      DataSource      =   "MyAdodc"
      ForeColor       =   &H00FF0000&
      Height          =   555
      Left            =   1200
      MaxLength       =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1980
      Width           =   3495
   End
   Begin VB.TextBox txtADDRESS1 
      DataField       =   "ADDRESS1"
      DataSource      =   "MyAdodc"
      ForeColor       =   &H00FF0000&
      Height          =   555
      Left            =   1200
      MaxLength       =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1350
      Width           =   3495
   End
   Begin VB.TextBox txtBRANCH_CODE 
      DataField       =   "BRANCH_CODE"
      DataSource      =   "MyAdodc"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   2
      Top             =   975
      Width           =   660
   End
   Begin VB.Label lblMoveRecord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Home"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   300
      TabIndex        =   22
      ToolTipText     =   "First Record"
      Top             =   3510
      Width           =   885
   End
   Begin VB.Image imgMoveRecord 
      Height          =   315
      Index           =   1
      Left            =   1200
      Picture         =   "frmBranchInfo.frx":0E42
      ToolTipText     =   "Previous Record"
      Top             =   3360
      Width           =   315
   End
   Begin VB.Label lblMoveRecord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Up"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   21
      ToolTipText     =   "Previous Record"
      Top             =   3540
      Width           =   885
   End
   Begin VB.Image imgMoveRecord 
      Height          =   315
      Index           =   2
      Left            =   2520
      Picture         =   "frmBranchInfo.frx":0EE0
      ToolTipText     =   "Next Record"
      Top             =   3360
      Width           =   315
   End
   Begin VB.Label lblMoveRecord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Next"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   2820
      TabIndex        =   20
      ToolTipText     =   "Next Record"
      Top             =   3510
      Width           =   675
   End
   Begin VB.Image imgMoveRecord 
      Height          =   315
      Index           =   3
      Left            =   3720
      Picture         =   "frmBranchInfo.frx":0F7E
      ToolTipText     =   "Last Record"
      Top             =   3360
      Width           =   315
   End
   Begin VB.Label lblMoveRecord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&End"
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   3
      Left            =   4020
      TabIndex        =   19
      ToolTipText     =   "Last Record"
      Top             =   3510
      Width           =   825
   End
   Begin VB.Image imgMoveRecord 
      Height          =   315
      Index           =   0
      Left            =   0
      Picture         =   "frmBranchInfo.frx":101C
      ToolTipText     =   "First Record"
      Top             =   3360
      Width           =   315
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Phone:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   4
      Left            =   525
      TabIndex        =   9
      Top             =   3015
      Width           =   660
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&City:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   3
      Left            =   750
      TabIndex        =   7
      Top             =   2640
      Width           =   435
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address&2:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   2
      Left            =   270
      TabIndex        =   5
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address&1:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   195
      TabIndex        =   3
      Top             =   1395
      Width           =   975
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Branch Code:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   1
      Top             =   1020
      Width           =   1275
   End
   Begin VB.Image Image2 
      Height          =   600
      Left            =   4200
      Picture         =   "frmBranchInfo.frx":10BA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image Image3 
      Height          =   585
      Left            =   0
      Picture         =   "frmBranchInfo.frx":14FC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Branch Information"
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
      TabIndex        =   0
      Top             =   0
      Width           =   4875
   End
   Begin VB.Shape shpMoveRecord 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   345
      Index           =   0
      Left            =   30
      Shape           =   2  'Oval
      Top             =   3510
      Width           =   1155
   End
   Begin VB.Shape shpMoveRecord 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   345
      Index           =   1
      Left            =   1200
      Shape           =   2  'Oval
      Top             =   3510
      Width           =   1155
   End
   Begin VB.Shape shpMoveRecord 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   345
      Index           =   2
      Left            =   2520
      Shape           =   2  'Oval
      Top             =   3510
      Width           =   1155
   End
   Begin VB.Shape shpMoveRecord 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   345
      Index           =   3
      Left            =   3720
      Shape           =   2  'Oval
      Top             =   3510
      Width           =   1155
   End
End
Attribute VB_Name = "frmBranchInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRecord_Click(Index As Integer)
    Record Index
End Sub

Private Sub cmdReport_Click()
    'Unload envReport
    
    envReport.Branch_Grouping txtBRANCH_CODE.Text & ""
    'ReportOfBranch.Show
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " on " & Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CloseForm
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyChecker KeyCode
End Sub
Private Sub imgMoveRecord_Click(Index As Integer)
    MoveRecord Index, MyAdodc
End Sub
Private Sub lblMoveRecord_Click(Index As Integer)
    MoveRecord Index, MyAdodc
End Sub

Private Sub MyAdodc_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    lblTitle.Caption = "Branch " & MyAdodc.Recordset.AbsolutePosition & "/" & MyAdodc.Recordset.RecordCount _
       & "Information"
End Sub

