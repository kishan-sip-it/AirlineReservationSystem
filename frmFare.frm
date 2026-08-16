VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFare 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fare information"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFare.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.UpDown UpDownClass 
      Height          =   375
      Index           =   2
      Left            =   6240
      TabIndex        =   30
      Top             =   3840
      Width           =   240
      _ExtentX        =   344
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownClass 
      Height          =   375
      Index           =   1
      Left            =   4200
      TabIndex        =   29
      Top             =   3840
      Width           =   240
      _ExtentX        =   344
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownClass 
      Height          =   375
      Index           =   0
      Left            =   2040
      TabIndex        =   28
      Top             =   3840
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
      TabIndex        =   20
      Top             =   5040
      Width           =   6735
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Add"
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Edit"
         Default         =   -1  'True
         Height          =   495
         Index           =   2
         Left            =   1080
         TabIndex        =   26
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Save"
         Height          =   495
         Index           =   3
         Left            =   2040
         TabIndex        =   25
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Undo"
         Height          =   495
         Index           =   4
         Left            =   3000
         TabIndex        =   24
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Delete"
         Height          =   495
         Index           =   5
         Left            =   3960
         TabIndex        =   23
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Search"
         Height          =   495
         Index           =   6
         Left            =   4920
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdRecord 
         Cancel          =   -1  'True
         Caption         =   "Retur&n"
         Height          =   495
         Index           =   7
         Left            =   5760
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "frmFare.frx":0442
      DataField       =   "DESTINATION"
      DataSource      =   "MyAdodc"
      Height          =   330
      Left            =   2040
      TabIndex        =   7
      Top             =   2640
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   16711680
      ListField       =   "DESTINATION"
      Text            =   ""
      Object.DataMember      =   "OriginDestination"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frmFare.frx":045E
      DataField       =   "ORIGIN"
      DataSource      =   "MyAdodc"
      Height          =   330
      Left            =   2040
      TabIndex        =   5
      Top             =   2040
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   16711680
      ListField       =   "ORIGIN"
      Text            =   ""
      Object.DataMember      =   "OriginDestination"
   End
   Begin VB.TextBox txtFare 
      DataField       =   "FIRST_FARE"
      DataSource      =   "MyAdodc"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   360
      MaxLength       =   15
      TabIndex        =   9
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox txtFare 
      DataField       =   "BUS_FARE"
      DataSource      =   "MyAdodc"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   2520
      MaxLength       =   15
      TabIndex        =   11
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox txtFare 
      DataField       =   "ECO_FARE"
      DataSource      =   "MyAdodc"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   2
      Left            =   4560
      MaxLength       =   15
      TabIndex        =   13
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox txtROUTE_DESC 
      DataField       =   "ROUTE_DESC"
      DataSource      =   "MyAdodc"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2040
      MaxLength       =   25
      TabIndex        =   3
      Top             =   1440
      Width           =   3855
   End
   Begin VB.TextBox txtROUTE_CODE 
      DataField       =   "ROUTE_CODE"
      DataSource      =   "MyAdodc"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2040
      MaxLength       =   7
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc MyAdodc 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      ToolTipText     =   "Move To Specified Record"
      Top             =   5490
      Visible         =   0   'False
      Width           =   6750
      _ExtentX        =   11906
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
      RecordSource    =   "SELECT * FROM FARE"
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
   Begin VB.Image Image2 
      Height          =   600
      Left            =   6000
      Picture         =   "frmFare.frx":047A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image Image3 
      Height          =   600
      Left            =   0
      Picture         =   "frmFare.frx":08BC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
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
      Index           =   4
      Left            =   360
      TabIndex        =   8
      Top             =   3600
      Width           =   930
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
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
      Index           =   5
      Left            =   2520
      TabIndex        =   10
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Economy Class:"
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
      Left            =   4560
      TabIndex        =   12
      Top             =   3600
      Width           =   1305
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "Fare Information"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label lblMoveRecord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Home"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   1020
      TabIndex        =   14
      ToolTipText     =   "First Record"
      Top             =   4590
      Width           =   885
   End
   Begin VB.Image imgMoveRecord 
      Height          =   315
      Index           =   1
      Left            =   2160
      Picture         =   "frmFare.frx":0CFE
      ToolTipText     =   "Previous Record"
      Top             =   4440
      Width           =   315
   End
   Begin VB.Label lblMoveRecord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Up"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   15
      ToolTipText     =   "Previous Record"
      Top             =   4620
      Width           =   885
   End
   Begin VB.Image imgMoveRecord 
      Height          =   315
      Index           =   2
      Left            =   3720
      Picture         =   "frmFare.frx":0D9C
      ToolTipText     =   "Next Record"
      Top             =   4440
      Width           =   315
   End
   Begin VB.Label lblMoveRecord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Next"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   4020
      TabIndex        =   16
      ToolTipText     =   "Next Record"
      Top             =   4590
      Width           =   675
   End
   Begin VB.Image imgMoveRecord 
      Height          =   315
      Index           =   3
      Left            =   5160
      Picture         =   "frmFare.frx":0E3A
      ToolTipText     =   "Last Record"
      Top             =   4440
      Width           =   315
   End
   Begin VB.Label lblMoveRecord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&End"
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   3
      Left            =   5460
      TabIndex        =   17
      ToolTipText     =   "Last Record"
      Top             =   4590
      Width           =   825
   End
   Begin VB.Image imgMoveRecord 
      Height          =   315
      Index           =   0
      Left            =   720
      Picture         =   "frmFare.frx":0ED8
      ToolTipText     =   "First Record"
      Top             =   4440
      Width           =   315
   End
   Begin VB.Shape shpMoveRecord 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   345
      Index           =   3
      Left            =   5160
      Shape           =   2  'Oval
      Top             =   4590
      Width           =   1155
   End
   Begin VB.Shape shpMoveRecord 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   345
      Index           =   2
      Left            =   3720
      Shape           =   2  'Oval
      Top             =   4590
      Width           =   1155
   End
   Begin VB.Shape shpMoveRecord 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   345
      Index           =   1
      Left            =   2160
      Shape           =   2  'Oval
      Top             =   4590
      Width           =   1155
   End
   Begin VB.Shape shpMoveRecord 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   345
      Index           =   0
      Left            =   750
      Shape           =   2  'Oval
      Top             =   4590
      Width           =   1155
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   975
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   6495
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Destination:"
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
      Left            =   1065
      TabIndex        =   6
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Origin:"
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
      Left            =   1485
      TabIndex        =   4
      Top             =   2040
      Width           =   540
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
      Left            =   1065
      TabIndex        =   2
      Top             =   1440
      Width           =   990
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Route &Code:"
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
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   1005
   End
   Begin VB.Image Image1 
      Height          =   5340
      Left            =   -240
      Stretch         =   -1  'True
      Top             =   600
      Width           =   7020
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Fare Information"
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
      Left            =   -600
      TabIndex        =   19
      Top             =   0
      Width           =   7275
   End
   Begin VB.Image Image5 
      Height          =   600
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6720
   End
End
Attribute VB_Name = "frmFare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRecord_Click(Index As Integer)
    Record Index
End Sub

Private Sub Form_Resize()
    Image1.Height = Me.Height
    Image1.Width = Me.Width
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
    lblTitle.Caption = "Fare(" & MyAdodc.Recordset.AbsolutePosition _
        & "/" & MyAdodc.Recordset.RecordCount & ")'s Information "
End Sub


Private Sub txtFare_GotFocus(Index As Integer)
  If (Val(txtFare(1).Text) > Val(txtFare(0).Text)) Then
MsgBox "Fare of BusinessClass should be less then FirstClass", vbInformation, "STOP...."
txtFare(1).Text = " "
End If
If (Val(txtFare(2).Text) > Val(txtFare(0).Text) Or Val(txtFare(2).Text) > Val(txtFare(1).Text)) Then
MsgBox "Fare of EconpomyClass should be less then FirstClass and BusinessClass", vbInformation, "STOP...."
txtFare(2).Text = " "
End If
End Sub

Private Sub txtFare_KeyPress(Index As Integer, KeyAscii As Integer)
    If (Chr(KeyAscii) >= "0" And Chr(KeyAscii) <= "9") _
        Or Chr(KeyAscii) = vbBack Then Exit Sub
    Beep
    KeyAscii = 0
 
End Sub

Private Sub UpDownClass_DownClick(Index As Integer)
    If Val(txtFare(Index).Text) = 0 Then Exit Sub
    txtFare(Index).Text = Val(txtFare(Index).Text) - 1
End Sub

Private Sub UpDownClass_UpClick(Index As Integer)
    txtFare(Index).Text = Val(txtFare(Index).Text) + 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyChecker KeyCode
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " on " & Date
End Sub

