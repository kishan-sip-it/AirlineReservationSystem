VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm AirReservationMDI 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Air-Line Reservation"
   ClientHeight    =   5265
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7590
   Icon            =   "AirReservationMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   Picture         =   "AirReservationMDI.frx":0E42
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   4770
      Width           =   7590
      _ExtentX        =   13388
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   2831
            Picture         =   "AirReservationMDI.frx":DD9F
            TextSave        =   "12:10 AM"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "12/28/02"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            AutoSize        =   2
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   2
            Enabled         =   0   'False
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Bevel           =   2
            Picture         =   "AirReservationMDI.frx":E1F1
            Text            =   "Developed by KISHAN"
            TextSave        =   "Developed by KISHAN"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuEntryScreen 
      Caption         =   "&Screen"
      Begin VB.Menu mnuScreen 
         Caption         =   "&Reservation"
         Index           =   1
      End
      Begin VB.Menu mnuScreen 
         Caption         =   "Fa&re"
         Index           =   2
      End
      Begin VB.Menu mnuScreen 
         Caption         =   "Flight &Schedule"
         Index           =   3
      End
      Begin VB.Menu mnuScreen 
         Caption         =   "&Flight"
         Index           =   4
      End
      Begin VB.Menu mnuScreen 
         Caption         =   "&Air Bus"
         Index           =   5
      End
      Begin VB.Menu mnuScreen 
         Caption         =   "&Branch Information"
         Index           =   6
      End
      Begin VB.Menu mnuScreen 
         Caption         =   "&Control Information"
         Index           =   7
      End
      Begin VB.Menu mnuScreen 
         Caption         =   "&Service"
         Index           =   8
      End
   End
   Begin VB.Menu mnuRecord 
      Caption         =   "&Record"
      Begin VB.Menu mnuNewRecord 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuEditRecord 
         Caption         =   "&Edit"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuRecordSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveRecord 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuUndoRecord 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRecordSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteRecord 
         Caption         =   "&Delete"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuRecordSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecordSearch 
         Caption         =   "&Search"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuRecordSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecordReturn 
         Caption         =   "&Return"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "&Report"
      Begin VB.Menu mnuCancellation 
         Caption         =   "Cancellation"
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
      Begin VB.Menu mnuExitNo 
         Caption         =   "&No"
      End
      Begin VB.Menu mnuExitYes 
         Caption         =   "&Yes"
      End
   End
End
Attribute VB_Name = "AirReservationMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SelectScreen As Integer

Public Sub MDIForm_Load()
    ' (removed leftover dev-machine debug line referencing a local game exe)
    mnuRecord.Visible = False
    mnuEntryScreen.Visible = True
    
    Code(1) = "Reservation Code"
    Code(2) = "Route Code"
    Code(3) = "Flight Code"
    Code(4) = "Flight Code"
    Code(5) = "Air Bus Code"
    Code(6) = "Branch Code"
    Code(8) = "Service Code"
    
    
End Sub

Private Sub mnuCancellation_Click()
    Dim SQL As String
    SQL = "SELECT FLIGHT_SCH.*, FARE.*, BRANCH.*, RESERVATION.*, SERVICE.SS_DESC AS EXPR1 FROM RESERVATION, BRANCH, FLIGHT_SCH, FARE, SERVICE WHERE RESERVATION.BRANCH_CODE = BRANCH.BRANCH_CODE AND RESERVATION.FLIGHTNO = FLIGHT_SCH.FLIGHTNO AND FLIGHT_SCH.ROUTE_CODE = FARE.ROUTE_CODE AND RESERVATION.SS_CODE = SERVICE.SS_CODE"
    Unload envReport
    envReport.rsReservationRoute.Open SQL, envReport.cnReport, adOpenDynamic, adLockPessimistic
    'ReservationReport.Show
    
    
    SQL = "SELECT BRANCH.*, ROUTEINFO.*, FLIGHT_SCH.*, CANCELLATION.*, CANCELLATION.TOTAL_FARE - CANCELLATION.CancellationCharge AS Refund FROM CANCELLATION, BRANCH, FLIGHT_SCH, ROUTEINFO, RESERVATION WHERE CANCELLATION.BRANCH_CODE = BRANCH.BRANCH_CODE AND CANCELLATION.FLIGHTNO = FLIGHT_SCH.FLIGHTNO AND FLIGHT_SCH.ROUTE_CODE = ROUTEINFO.`Route Code`"
    
    MsgBox SQL
    
    Unload envReport
    envReport.rsCancelReservation.Open SQL, envReport.cnReport, adOpenDynamic, adLockPessimistic
    CancelReservationReport.Show
End Sub

Private Sub mnuAllFlightFare_Click()
    Unload envReport
    AllFlightFareReport.Show
End Sub

Private Sub mnuAllReservation_Click()
    Dim SQL As String
    SQL = "SELECT FLIGHT_SCH.*, FARE.*, BRANCH.*, RESERVATION.*, SERVICE.SS_DESC AS EXPR1 FROM RESERVATION, BRANCH, FLIGHT_SCH, FARE, SERVICE WHERE RESERVATION.BRANCH_CODE = BRANCH.BRANCH_CODE AND RESERVATION.FLIGHTNO = FLIGHT_SCH.FLIGHTNO AND FLIGHT_SCH.ROUTE_CODE = FARE.ROUTE_CODE AND RESERVATION.SS_CODE = SERVICE.SS_CODE"
    Unload envReport
    envReport.rsReservationRoute.Open SQL, envReport.cnReport, adOpenDynamic, adLockPessimistic
    ReservationReport.Show
End Sub

Private Sub mnuBranchReport_Click()
    Unload envReport
    ReportOfBranch.Show
End Sub

Private Sub mnuCancelReservation_Click()
    
    Dim SQL As String
    SQL = "SELECT BRANCH.*, ROUTEINFO.*, FLIGHT_SCH.*, CANCELLATION.*, CANCELLATION.TOTAL_FARE - CANCELLATION.CancellationCharge AS Refund FROM CANCELLATION, BRANCH, FLIGHT_SCH, ROUTEINFO, RESERVATION WHERE CANCELLATION.BRANCH_CODE = BRANCH.BRANCH_CODE AND CANCELLATION.FLIGHTNO = FLIGHT_SCH.FLIGHTNO AND FLIGHT_SCH.ROUTE_CODE = ROUTEINFO.`Route Code`"
    
    
    Unload envReport
    envReport.rsCancelReservation.Open SQL, envReport.cnReport, adOpenDynamic, adLockPessimistic
    CancelReservationReport.Show
    
    Exit Sub
    With frmSearch
        .MyAdodc.RecordSource = "SELECT * FROM CANCELLATION"
        .MyAdodc.Refresh
        
        .lblTitle = Code(1)
        CancelReport = True
        Set SearchForm = Me.ActiveForm
        .Show vbModal
    End With
End Sub

Public Sub mnuDeleteRecord_Click()
On Error GoTo DeleteRecordError
    If MsgBox("Do you want to delete this Record" _
        , vbQuestion + vbYesNo + vbDefaultButton2 _
        , "Delete Confirmation") = vbYes Then
        AirReservationMDI.ActiveForm.MyAdodc.Recordset.Delete
        
        AirReservationMDI.ActiveForm.MyAdodc.Recordset.MoveNext
        AirReservationMDI.ActiveForm.MyAdodc.Recordset.MovePrevious
        AirReservationMDI.ActiveForm.MyAdodc.Refresh
    End If
    Exit Sub
DeleteRecordError:
    MsgBox Err.Description, vbCritical, "New Record Error"
End Sub

Public Sub mnuEditRecord_Click()
On Error GoTo EditError
    LockMenu
    'EditControl Me.ActiveForm
    Exit Sub
EditError:
    MsgBox Err.Description, vbCritical, "New Record Error"
End Sub

Private Sub mnuExitYes_Click()
    Unload AirReservationMDI
End Sub

Public Sub mnuNewRecord_Click()
On Error GoTo NewRecordError
    Me.ActiveForm.MyAdodc.Recordset.AddNew
    LockMenu
    'EditControl Me.ActiveForm
    Exit Sub
NewRecordError:
    MsgBox Err.Description, vbCritical, "New Record Error"
End Sub

Public Sub mnuScreenAirBus_Click()
    OpenForm frmAirBus
    frmAirBus.Show
End Sub

Public Sub mnuReservationScreen_Click()
    OpenForm frmReservation
    frmReservation.Show
End Sub

Public Sub mnuRecordReturn_Click()
    If MsgBox("Close(Yes/No:- ", vbYesNo + vbDefaultButton2 + vbQuestion, "Close confirmation") = vbYes Then _
        Unload AirReservationMDI.ActiveForm
End Sub

Public Sub mnuRecordSearch_Click()
On Error GoTo SearchError
    With frmSearch
        .MyAdodc.RecordSource = AirReservationMDI.ActiveForm.MyAdodc.RecordSource
        .MyAdodc.Refresh
        
        .lblTitle = Code(SelectScreen)
        
        Set SearchForm = Me.ActiveForm
        .Show vbModal
    End With
    Exit Sub
SearchError:
    MsgBox Err.Description, vbCritical, "New Record Error"
End Sub

Public Sub mnuSaveRecord_Click()
On Error GoTo SaveError
    If MsgBox("Do you want to save this Record" _
        , vbQuestion + vbYesNo, "Save confirmation") = vbYes Then
        AirReservationMDI.ActiveForm.MyAdodc.Recordset.Update
        MsgBox "Record has been saved", vbInformation, "Save Record"
    Else
        AirReservationMDI.ActiveForm.MyAdodc.Refresh
    End If
    UnLockMenu
   ' NoEditControl Me.ActiveForm
    Exit Sub
SaveError:
    MsgBox Err.Description, vbCritical, "New Record Error"
End Sub

Public Sub mnuScreen_Click(Index As Integer)
    SelectScreen = Index
    Select Case Index
        Case 1
           OpenForm frmReservation
        Case 2
            OpenForm frmFare
        Case 3
            OpenForm frmFlightSchedule
        Case 4
            OpenForm frmFlight
        Case 5
            OpenForm frmAirBus
        Case 6
            OpenForm frmBranchInfo
        Case 7
            OpenForm frmControl
        Case 8
            OpenForm frmService
    End Select
End Sub
Private Sub mnuSpecificFlight_Click()
    FlightFareReport.Show
End Sub

Public Sub mnuSpecificReservation_Click()
    If mPNR = 0 Then
        MsgBox "Please, Open Reservation Screen", vbExclamation, "No Reservation No"
        Exit Sub
    End If
    
    Dim SQL As String
    SQL = "SELECT FLIGHT_SCH.*, FARE.*, BRANCH.*, RESERVATION.*, SERVICE.SS_DESC AS EXPR1 FROM RESERVATION, BRANCH, FLIGHT_SCH, FARE, SERVICE WHERE RESERVATION.BRANCH_CODE = BRANCH.BRANCH_CODE AND RESERVATION.FLIGHTNO = FLIGHT_SCH.FLIGHTNO AND FLIGHT_SCH.ROUTE_CODE = FARE.ROUTE_CODE AND RESERVATION.SS_CODE = SERVICE.SS_CODE AND RESERVATION.PNR=" & mPNR
    
    Unload envReport
    envReport.rsReservationRoute.Open SQL, envReport.cnReport, adOpenDynamic, adLockPessimistic
    
    'ReservationReport.Show
'    ReservationReport.PrintReport
    
    mPNR = 0
End Sub

Public Sub mnuUndoRecord_Click()
On Error GoTo UndoError
    UnLockMenu
    'NoEditControl Me.ActiveForm
    AirReservationMDI.ActiveForm.MyAdodc.Refresh
    Exit Sub
UndoError:
    MsgBox Err.Description, vbCritical, "New Record Error"
End Sub
