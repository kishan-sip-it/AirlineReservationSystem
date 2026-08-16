VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFlightSchedule 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Flight Schedule"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9765
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFlightSchedule.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   9765
   Begin MSComCtl2.UpDown UpDownDay 
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   30
      Top             =   3960
      Width           =   240
      _ExtentX        =   344
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownDay 
      Height          =   375
      Index           =   0
      Left            =   2040
      TabIndex        =   29
      Top             =   3480
      Width           =   240
      _ExtentX        =   344
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.Frame Frame1 
      Caption         =   "Record Operation"
      Height          =   855
      Left            =   840
      TabIndex        =   21
      Top             =   5280
      Width           =   8415
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Add"
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   28
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Edit"
         Default         =   -1  'True
         Height          =   495
         Index           =   2
         Left            =   1200
         TabIndex        =   27
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Save"
         Height          =   495
         Index           =   3
         Left            =   2400
         TabIndex        =   26
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Undo"
         Height          =   495
         Index           =   4
         Left            =   3600
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Delete"
         Height          =   495
         Index           =   5
         Left            =   4800
         TabIndex        =   24
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Search"
         Height          =   495
         Index           =   6
         Left            =   6000
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdRecord 
         Cancel          =   -1  'True
         Caption         =   "Retur&n"
         Height          =   495
         Index           =   7
         Left            =   7200
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid 
      Height          =   1335
      Left            =   3480
      TabIndex        =   20
      Top             =   2040
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   2355
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Route Information"
      ColumnCount     =   23
      BeginProperty Column00 
         DataField       =   "Route Code"
         Caption         =   "Route Code"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Route"
         Caption         =   "Route"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Origin"
         Caption         =   "Origin"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Destination"
         Caption         =   "Destination"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "First Class"
         Caption         =   "First Class"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Business Class"
         Caption         =   "Business Class"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Economic Class"
         Caption         =   "Economic Class"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Departure Time"
         Caption         =   "Departure Time"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "HOURS"
         Caption         =   "HOURS"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "Flight Day1"
         Caption         =   "Flight Day1"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "Flight Day2"
         Caption         =   "Flight Day2"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "Flight No"
         Caption         =   "Flight No"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "Flight DAte"
         Caption         =   "Flight DAte"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "FirstClass Booking"
         Caption         =   "FirstClass Booking"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "Business Class Booking"
         Caption         =   "Business Class Booking"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column15 
         DataField       =   "Economy Class Booking"
         Caption         =   "Economy Class Booking"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column16 
         DataField       =   "AirBus No"
         Caption         =   "AirBus No"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column17 
         DataField       =   "First Class Capacity"
         Caption         =   "First Class Capacity"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column18 
         DataField       =   "Business Class Capacity"
         Caption         =   "Business Class Capacity"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column19 
         DataField       =   "Economic Class Capacity"
         Caption         =   "Economic Class Capacity"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column20 
         DataField       =   "First Class Waiting"
         Caption         =   "First Class Waiting"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column21 
         DataField       =   "Business Class Waiting"
         Caption         =   "Business Class Waiting"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column22 
         DataField       =   "Economic Class Waiting"
         Caption         =   "Economic Class Waiting"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   689.953
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1349.858
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1755.213
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   1830.047
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   1349.858
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   1709.858
         EndProperty
         BeginProperty Column22 
            ColumnWidth     =   1725.165
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtDay 
      DataField       =   "FLIGHT_DAY2"
      DataSource      =   "MyAdodc"
      Height          =   375
      Index           =   1
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   13
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox txtDay 
      DataField       =   "FLIGHT_DAY1"
      DataSource      =   "MyAdodc"
      Height          =   375
      Index           =   0
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   11
      Top             =   3480
      Width           =   495
   End
   Begin MSDataListLib.DataCombo RouteCode 
      Bindings        =   "frmFlightSchedule.frx":0E42
      DataField       =   "ROUTE_CODE"
      DataSource      =   "MyAdodc"
      Height          =   330
      Left            =   1560
      TabIndex        =   5
      Top             =   2040
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Style           =   2
      ListField       =   "Route Code"
      Text            =   ""
      Object.DataMember      =   "RouteInformation"
   End
   Begin MSDataListLib.DataCombo AirBusNo 
      Bindings        =   "frmFlightSchedule.frx":0E5E
      DataField       =   "AIRBUSNO"
      DataSource      =   "MyAdodc"
      Height          =   330
      Left            =   1560
      TabIndex        =   3
      Top             =   1440
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Style           =   2
      ListField       =   "AIRBUSNO"
      Text            =   ""
      Object.DataMember      =   "AirBusInfo"
   End
   Begin VB.TextBox txtJOURNEY_HRS 
      DataField       =   "JOURNEY_HRS"
      DataSource      =   "MyAdodc"
      Height          =   315
      Left            =   1560
      MaxLength       =   5
      TabIndex        =   9
      Top             =   2970
      Width           =   825
   End
   Begin VB.TextBox txtDEPRT_TIME 
      DataField       =   "DEPRT_TIME"
      DataSource      =   "MyAdodc"
      Height          =   315
      Left            =   1530
      MaxLength       =   5
      TabIndex        =   7
      Top             =   2475
      Width           =   825
   End
   Begin VB.TextBox txtFLIGHTNO 
      DataField       =   "FLIGHTNO"
      DataSource      =   "MyAdodc"
      Height          =   315
      Left            =   1530
      MaxLength       =   4
      TabIndex        =   1
      Top             =   975
      Width           =   660
   End
   Begin MSAdodcLib.Adodc MyAdodc 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      Top             =   5355
      Visible         =   0   'False
      Width           =   9765
      _ExtentX        =   17224
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
      RecordSource    =   "SELECT * FROM Flight_Sch"
      Caption         =   "FlightSchedule"
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
   Begin MSAdodcLib.Adodc Temp 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      Top             =   5730
      Visible         =   0   'False
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      RecordSource    =   "SELECT * FROM RouteInfo"
      Caption         =   "Temp"
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
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Flight Schedule Information"
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
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   9795
   End
   Begin VB.Image imgMoveRecord 
      Height          =   315
      Index           =   1
      Left            =   1800
      Picture         =   "frmFlightSchedule.frx":0E7A
      ToolTipText     =   "Previous Record"
      Top             =   4800
      Width           =   315
   End
   Begin VB.Image imgMoveRecord 
      Height          =   315
      Index           =   2
      Left            =   3360
      Picture         =   "frmFlightSchedule.frx":0F18
      ToolTipText     =   "Next Record"
      Top             =   4800
      Width           =   315
   End
   Begin VB.Image imgMoveRecord 
      Height          =   315
      Index           =   3
      Left            =   4800
      Picture         =   "frmFlightSchedule.frx":0FB6
      ToolTipText     =   "Last Record"
      Top             =   4800
      Width           =   315
   End
   Begin VB.Image imgMoveRecord 
      Height          =   315
      Index           =   0
      Left            =   360
      Picture         =   "frmFlightSchedule.frx":1054
      ToolTipText     =   "First Record"
      Top             =   4800
      Width           =   315
   End
   Begin VB.Image Image2 
      Height          =   600
      Left            =   9120
      Picture         =   "frmFlightSchedule.frx":10F2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image Image3 
      Height          =   600
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Flight Day&2:"
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
      Left            =   585
      TabIndex        =   12
      Top             =   4020
      Width           =   915
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Flight Day&1:"
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
      Left            =   585
      TabIndex        =   10
      Top             =   3525
      Width           =   915
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Journey Hours:"
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
      Left            =   255
      TabIndex        =   8
      Top             =   3015
      Width           =   1260
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Departure &Time:"
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
      Left            =   165
      TabIndex        =   6
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Route Code:"
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
      Left            =   495
      TabIndex        =   4
      Top             =   2025
      Width           =   1005
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
      Index           =   1
      Left            =   600
      TabIndex        =   2
      Top             =   1515
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
      Left            =   840
      TabIndex        =   0
      Top             =   1020
      Width           =   750
   End
   Begin VB.Label lblMoveRecord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Home"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   14
      ToolTipText     =   "First Record"
      Top             =   5040
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
      TabIndex        =   15
      ToolTipText     =   "Previous Record"
      Top             =   4980
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
      TabIndex        =   16
      ToolTipText     =   "Next Record"
      Top             =   4950
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
      MousePointer    =   99  'Custom
      TabIndex        =   17
      ToolTipText     =   "Last Record"
      Top             =   4950
      Width           =   825
   End
   Begin VB.Label lblDay 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   1
      Left            =   2400
      TabIndex        =   19
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label lblDay 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   0
      Left            =   2400
      TabIndex        =   18
      Top             =   3480
      Width           =   2775
   End
   Begin VB.Shape shpMoveRecord 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   345
      Index           =   3
      Left            =   4800
      Shape           =   2  'Oval
      Top             =   4920
      Width           =   1155
   End
   Begin VB.Shape shpMoveRecord 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   345
      Index           =   2
      Left            =   3360
      Shape           =   2  'Oval
      Top             =   4950
      Width           =   1155
   End
   Begin VB.Shape shpMoveRecord 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   345
      Index           =   1
      Left            =   1800
      Shape           =   2  'Oval
      Top             =   4950
      Width           =   1155
   End
   Begin VB.Shape shpMoveRecord 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   345
      Index           =   0
      Left            =   390
      Shape           =   2  'Oval
      Top             =   4950
      Width           =   1155
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9825
   End
   Begin VB.Image Image1 
      Height          =   5580
      Left            =   -120
      Stretch         =   -1  'True
      Top             =   480
      Width           =   10260
   End
End
Attribute VB_Name = "frmFlightSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CDays(1 To 7) As String
Private Sub AirBusNo_LostFocus()
    DataGrid.Visible = False
End Sub

Private Sub cmdRecord_Click(Index As Integer)
    Record Index
End Sub

Private Sub Form_Load()
    CDays(1) = "Sunday"
    CDays(2) = "Monday"
    CDays(3) = "Tuesday"
    CDays(4) = "Wednesday"
    CDays(5) = "Thursday"
    CDays(6) = "Friday"
    CDays(7) = "Saturday"
    DataGrid.Visible = False
    MyAdodc.Recordset.MoveNext
    MyAdodc.Recordset.MoveFirst
    
    Me.Caption = Me.Caption & " on " & Date
End Sub

Private Sub Form_Resize()
    Image1.Width = Me.Width
    Image1.Height = Me.Height
    DataGrid.Width = Me.Width - 4000
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
    lblTitle.Caption = "Flight Schedule (" & MyAdodc.Recordset.AbsolutePosition & "/" _
        & MyAdodc.Recordset.RecordCount & ")'s Information "

End Sub
Private Sub RouteCode_Click(Area As Integer)
    Temp.RecordSource = "SELECT * From RouteInfo WHERE ([Route Code] = '" & RouteCode.Text & "')"
    Temp.Refresh
    
'    DataGrid.Top = RouteCode.Top
'
'    DataGrid.ReBind
'    DataGrid.Visible = True
'    DataGrid.Caption = "Route Information"
End Sub

Private Sub RouteCode_LostFocus()
    DataGrid.Visible = False
End Sub
Private Sub txtDay_Change(Index As Integer)
    On Error Resume Next
    If MyAdodc.Recordset.BOF = True Or MyAdodc.Recordset.EOF = True Then Exit Sub
    If Val(txtDay(Index).Text) < 0 Or Val(txtDay(Index).Text) > 7 Then Exit Sub
    lblDay(Index).Caption = CDays(Val(txtDay(Index).Text))
End Sub

Private Sub txtDay_KeyPress(Index As Integer, KeyAscii As Integer)
    If (Chr(KeyAscii) >= "0" And Chr(KeyAscii) <= "9") _
        Or Chr(KeyAscii) = vbBack Then Exit Sub
    Beep
    KeyAscii = 0
End Sub

Private Sub txtDEPRT_TIME_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) >= "0" And Chr(KeyAscii) <= "9") _
       Or Chr(KeyAscii) = ":" Or Chr(KeyAscii) = vbBack Then Exit Sub
    Beep
    KeyAscii = 0
End Sub

Private Sub txtJOURNEY_HRS_KeyPress(KeyAscii As Integer)
   If (Chr(KeyAscii) >= "0" And Chr(KeyAscii) <= "9") _
       Or Chr(KeyAscii) = ":" Or Chr(KeyAscii) = vbBack Then Exit Sub
    Beep
    KeyAscii = 0
End Sub

Private Sub UpDownDay_DownClick(Index As Integer)
    txtDay(Index).Text = Val(txtDay(Index).Text) - 1
    If Val(txtDay(Index).Text) = 0 Then
        Beep
        txtDay(Index).Text = 1
    End If
End Sub

Private Sub UpDownDay_UpClick(Index As Integer)
    txtDay(Index).Text = Val(txtDay(Index).Text) + 1
    If Val(txtDay(Index).Text) > 7 Then
        Beep
        txtDay(Index).Text = 7
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyChecker KeyCode
End Sub

