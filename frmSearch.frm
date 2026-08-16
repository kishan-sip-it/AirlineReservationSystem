VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSearch 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.ComboBox cboCode 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   2520
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin MSAdodcLib.Adodc MyAdodc 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   5850
      _ExtentX        =   10319
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
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSearch_Click()
    If cboCode.Text = "" Then
        MsgBox "Please, select Code", vbCritical, "Select Error"
        Exit Sub
    End If
    
    If CancelReport = True Then
        mCPNR = Val(cboCode.Text)
        mBranchCode = cboCode.Text
        CancelReport = False
    Else
        With SearchForm.MyAdodc.Recordset
            .Find .Fields(0).Name & "='" & cboCode.Text & "'", 0, adSearchForward, 1
            
            If .EOF = True Or .BOF = True Then _
                MsgBox cboCode.Text & " not found", vbCritical, "Not found"
        End With
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
    With MyAdodc.Recordset
        cboCode.Clear
        Do Until .EOF
            cboCode.AddItem .Fields(0)
            .MoveNext
        Loop
    End With
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " on " & Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set SearchForm = Nothing
End Sub

