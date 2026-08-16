VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmControlInfo 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control Information"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmControlInfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmControlInfo.frx":0442
   ScaleHeight     =   5370
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.UpDown UpDownNum 
      Height          =   375
      Index           =   2
      Left            =   2160
      TabIndex        =   34
      Top             =   2280
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownNum 
      Height          =   375
      Index           =   6
      Left            =   4800
      TabIndex        =   33
      Top             =   2760
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownNum 
      Height          =   375
      Index           =   1
      Left            =   3360
      TabIndex        =   32
      Top             =   1320
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownNum 
      Height          =   375
      Index           =   7
      Left            =   4800
      TabIndex        =   31
      Top             =   3240
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownNum 
      Height          =   375
      Index           =   5
      Left            =   4800
      TabIndex        =   30
      Top             =   2280
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownNum 
      Height          =   375
      Index           =   3
      Left            =   2160
      TabIndex        =   29
      Top             =   2760
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownNum 
      Height          =   375
      Index           =   4
      Left            =   2160
      TabIndex        =   28
      Top             =   3240
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownNum 
      Height          =   375
      Index           =   0
      Left            =   3360
      TabIndex        =   27
      Top             =   960
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtNum 
      DataField       =   "CANC_DEDUC_3"
      DataSource      =   "MyAdodc"
      Height          =   315
      Index           =   7
      Left            =   4200
      MaxLength       =   2
      TabIndex        =   22
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox txtNum 
      DataField       =   "CANC_DEDUC_6"
      DataSource      =   "MyAdodc"
      Height          =   315
      Index           =   6
      Left            =   4245
      MaxLength       =   2
      TabIndex        =   21
      Top             =   2760
      Width           =   540
   End
   Begin VB.TextBox txtNum 
      DataField       =   "CANC_DEDUC_12"
      DataSource      =   "MyAdodc"
      Height          =   315
      Index           =   5
      Left            =   4245
      MaxLength       =   2
      TabIndex        =   20
      Top             =   2280
      Width           =   540
   End
   Begin VB.TextBox txtNum 
      DataField       =   "FIRST_BAG_LIMIT"
      DataSource      =   "MyAdodc"
      Height          =   315
      Index           =   2
      Left            =   1545
      MaxLength       =   3
      TabIndex        =   15
      Top             =   2280
      Width           =   660
   End
   Begin VB.TextBox txtNum 
      DataField       =   "BUS_BAG_LIMIT"
      DataSource      =   "MyAdodc"
      Height          =   315
      Index           =   3
      Left            =   1545
      MaxLength       =   3
      TabIndex        =   14
      Top             =   2775
      Width           =   660
   End
   Begin VB.TextBox txtNum 
      DataField       =   "ECO_BAG_LIMIT"
      DataSource      =   "MyAdodc"
      Height          =   315
      Index           =   4
      Left            =   1545
      MaxLength       =   3
      TabIndex        =   13
      Top             =   3285
      Width           =   660
   End
   Begin VB.Frame frameRecordOperation 
      Caption         =   "Record Operation"
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   5415
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&New"
         Height          =   495
         Index           =   1
         Left            =   480
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Edit"
         Default         =   -1  'True
         Height          =   495
         Index           =   2
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Save"
         Height          =   495
         Index           =   3
         Left            =   2640
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Undo"
         Height          =   495
         Index           =   4
         Left            =   3720
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Delete"
         Height          =   495
         Index           =   5
         Left            =   480
         TabIndex        =   8
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Search"
         Height          =   495
         Index           =   7
         Left            =   1560
         TabIndex        =   7
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "Retur&n"
         Height          =   495
         Index           =   8
         Left            =   2640
         TabIndex        =   6
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.TextBox txtNum 
      DataField       =   "AIR_TAX"
      DataSource      =   "MyAdodc"
      Height          =   315
      Index           =   0
      Left            =   2625
      MaxLength       =   4
      TabIndex        =   1
      Top             =   975
      Width           =   735
   End
   Begin VB.TextBox txtNum 
      DataField       =   "EXCESS_BAG_CHG"
      DataSource      =   "MyAdodc"
      Height          =   315
      Index           =   1
      Left            =   2625
      MaxLength       =   3
      TabIndex        =   0
      Top             =   1350
      Width           =   720
   End
   Begin MSAdodcLib.Adodc MyAdodc 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      ToolTipText     =   "Move To Specified Record"
      Top             =   4995
      Width           =   5595
      _ExtentX        =   9869
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
      UserName        =   "SCOTT"
      Password        =   "TIGER"
      RecordSource    =   "SELECT * FROM CONTROL"
      Caption         =   "Control Information"
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Cancllation Deduction(%)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2880
      TabIndex        =   26
      Top             =   1920
      Width           =   2370
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   1695
      Left            =   2760
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Before &3 Days:"
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
      Left            =   3045
      TabIndex        =   25
      Top             =   3240
      Width           =   1170
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Before &6 Days:"
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
      Left            =   3045
      TabIndex        =   24
      Top             =   2760
      Width           =   1170
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Before &12 Days:"
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
      Left            =   2925
      TabIndex        =   23
      Top             =   2325
      Width           =   1260
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "&Bag Limit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   1695
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   2415
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
      Left            =   600
      TabIndex        =   18
      Top             =   2325
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
      Index           =   3
      Left            =   195
      TabIndex        =   17
      Top             =   2820
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
      Index           =   4
      Left            =   225
      TabIndex        =   16
      Top             =   3330
      Width           =   1305
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Control Information"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5775
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Air Tax:"
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
      Left            =   1995
      TabIndex        =   3
      Top             =   1020
      Width           =   615
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Excess Bag Charge:"
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
      Left            =   975
      TabIndex        =   2
      Top             =   1395
      Width           =   1620
   End
End
Attribute VB_Name = "frmControlInfo"
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
Private Sub Form_Unload(Cancel As Integer)
    CloseForm
End Sub

Private Sub MyAdodc_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  lblTitle.Caption = "Control (" & MyAdodc.Recordset.AbsolutePosition & "/" _
        & MyAdodc.Recordset.RecordCount & ")'s Information"
End Sub

Private Sub txtNum_KeyPress(Index As Integer, KeyAscii As Integer)
    If Chr(KeyAscii) = vbBack Or _
        (Chr(KeyAscii) >= "0" And Chr(KeyAscii) <= "9") Then Exit Sub
    Beep
    KeyAscii = 0
End Sub
Private Sub UpDownNum_DownClick(Index As Integer)
    txtNum(Index).Text = Val(txtNum(Index).Text) - 1
    If Val(txtNum(Index).Text) <= 0 Then
        Beep
        MsgBox "Bag Limit can't be Zero", vbExclamation, "Input Error"
        txtNum(Index).Text = 1
    End If
End Sub

Private Sub UpDownNum_UpClick(Index As Integer)
    txtNum(Index).Text = Val(txtNum(Index).Text) + 1
End Sub
Private Sub Form_Load()
    Me.Caption = Me.Caption & " on " & Date
End Sub

