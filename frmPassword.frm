VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Password"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3255
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmPassword.frx":0442
   ScaleHeight     =   6495
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Appearance      =   0  'Flat
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox txtUser 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   -30
      TabIndex        =   3
      Top             =   1560
      Width           =   1125
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PleaseEnterUserName,Password"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "&User:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   405
      TabIndex        =   0
      Top             =   1080
      Width           =   570
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If txtUser.Text <> "kishan" Or txtPassword.Text <> "kishan@2026" Then
        MsgBox "Sorry,Invalid Username or Password" _
            & vbNewLine & "Please, try again", vbExclamation, _
            "Invalid Username or Password"
            txtUser.SetFocus
        ValidPassWord = False
    Else
        ValidPassWord = True
    End If
    
    Unload Me
    If ValidPassWord = True Then AirReservationMDI.Show
End Sub

