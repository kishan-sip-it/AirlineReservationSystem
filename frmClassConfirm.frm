VERSION 5.00
Begin VB.Form frmClassConfirm 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Class Confirm"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3960
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   735
      Left            =   2040
      TabIndex        =   5
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Ticket Reservation Choice"
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.OptionButton optChoice 
         BackColor       =   &H00FFC0FF&
         Caption         =   "&Cancel"
         Height          =   735
         Index           =   2
         Left            =   240
         MouseIcon       =   "frmClassConfirm.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   1680
         Width           =   3375
      End
      Begin VB.OptionButton optChoice 
         BackColor       =   &H00FFC0FF&
         Caption         =   "&Waiting"
         Height          =   615
         Index           =   1
         Left            =   240
         MouseIcon       =   "frmClassConfirm.frx":0442
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   1080
         Width           =   3375
      End
      Begin VB.OptionButton optChoice 
         BackColor       =   &H00FFC0FF&
         Caption         =   "&Confirm"
         Height          =   735
         Index           =   0
         Left            =   240
         MouseIcon       =   "frmClassConfirm.frx":0884
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   360
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmClassConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Choice = "Cancel"
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim I As Integer
    For I = optChoice.LBound To optChoice.UBound
        optChoice(I).Value = False
    Next I
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Choice = "" Then
        MsgBox "Please, select Confirmation", vbInformation, "Confirmation"
        Cancel = True
    End If
End Sub

Private Sub optChoice_Click(Index As Integer)
    If optChoice.Item(0).Value = True Then
        Choice = "Confirm"
    ElseIf optChoice.Item(1).Value = True Then
        Choice = "Waiting"
    ElseIf optChoice.Item(2).Value = True Then
        Choice = "Cancel"
    End If
    Unload Me
End Sub
