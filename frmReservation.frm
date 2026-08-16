VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmReservation 
   Caption         =   "Reservation Information"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11550
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReservation.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   11550
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      DataField       =   "SEAT_NO"
      DataSource      =   "MyAdodc"
      Height          =   315
      Left            =   3600
      TabIndex        =   50
      Top             =   2760
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid RouteGrid 
      Bindings        =   "frmReservation.frx":0442
      Height          =   2040
      Left            =   5640
      TabIndex        =   29
      Top             =   1080
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   3598
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   3
      RowHeight       =   19
      TabAcrossSplits =   -1  'True
      TabAction       =   2
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
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
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1379.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1425.26
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1049.953
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1049.953
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   824.882
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1665.071
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   2115.213
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   2160
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   1755.213
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   2145.26
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   2220.094
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   1679.811
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   2069.858
         EndProperty
         BeginProperty Column22 
            ColumnWidth     =   2145.26
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox chkConfirm 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Confirm Reservation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   720
      Left            =   7560
      Picture         =   "frmReservation.frx":045B
      TabIndex        =   24
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancelReservation 
      Caption         =   "&Cancel Reservation"
      Height          =   855
      Left            =   9480
      Picture         =   "frmReservation.frx":114DE
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   6720
      Width           =   1815
   End
   Begin VB.TextBox txtCREDIT_CARD_NO 
      DataField       =   "CREDIT_CARD_NO"
      DataSource      =   "MyAdodc"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   6885
      MaxLength       =   10
      TabIndex        =   21
      Top             =   5520
      Width           =   3135
   End
   Begin VB.TextBox txtPASSPORT_NO 
      DataField       =   "PASSPORT_NO"
      DataSource      =   "MyAdodc"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   6840
      MaxLength       =   8
      TabIndex        =   17
      Top             =   4560
      Width           =   3975
   End
   Begin VB.TextBox txtPASS_ADD3 
      DataField       =   "PASS_ADD3"
      DataSource      =   "MyAdodc"
      ForeColor       =   &H00FF0000&
      Height          =   675
      Left            =   6795
      MaxLength       =   40
      TabIndex        =   15
      Top             =   3840
      Width           =   3975
   End
   Begin VB.TextBox txtPASS_ADD2 
      DataField       =   "PASS_ADD2"
      DataSource      =   "MyAdodc"
      ForeColor       =   &H00FF0000&
      Height          =   555
      Left            =   6795
      MaxLength       =   40
      TabIndex        =   13
      Top             =   3240
      Width           =   3975
   End
   Begin VB.TextBox txtPASS_ADD1 
      DataField       =   "PASS_ADD1"
      DataSource      =   "MyAdodc"
      ForeColor       =   &H00FF0000&
      Height          =   555
      Left            =   6795
      MaxLength       =   40
      TabIndex        =   11
      Top             =   2640
      Width           =   3975
   End
   Begin VB.TextBox txtPASS_NAME 
      BackColor       =   &H00000000&
      DataField       =   "PASS_NAME"
      DataSource      =   "MyAdodc"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6840
      MaxLength       =   20
      TabIndex        =   9
      Top             =   2280
      Width           =   3975
   End
   Begin VB.TextBox txtBags 
      DataSource      =   "MyAdodc"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1680
      TabIndex        =   22
      Top             =   4800
      Width           =   735
   End
   Begin VB.ComboBox cboBranchCode 
      DataField       =   "BRANCH_CODE"
      DataSource      =   "MyAdodc"
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   1680
      TabIndex        =   23
      Top             =   5280
      Width           =   2895
   End
   Begin VB.OptionButton optClass 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Economy"
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   5
      Top             =   3960
      Width           =   975
   End
   Begin VB.OptionButton optClass 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Business"
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   6
      Top             =   3960
      Width           =   975
   End
   Begin VB.OptionButton optClass 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&First Class"
      Height          =   255
      Index           =   2
      Left            =   4080
      TabIndex        =   7
      Top             =   3960
      Width           =   1095
   End
   Begin VB.ComboBox cboFlightNo 
      DataField       =   "FLIGHTNO"
      DataSource      =   "MyAdodc"
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox txtPNR 
      BackColor       =   &H00000000&
      DataField       =   "PNR"
      DataSource      =   "MyAdodc"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      MaxLength       =   5
      TabIndex        =   37
      Top             =   2280
      Width           =   840
   End
   Begin VB.TextBox txtTOTAL_FARE 
      DataField       =   "TOTAL_FARE"
      DataSource      =   "MyAdodc"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   6000
      Width           =   2055
   End
   Begin VB.CommandButton cmdShowRoute 
      Caption         =   "&Show Route"
      Height          =   735
      Left            =   7200
      Picture         =   "frmReservation.frx":11920
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   840
      Width           =   1815
   End
   Begin VB.ComboBox cboDestination 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   960
      Width           =   4335
   End
   Begin MSAdodcLib.Adodc MyAdodc 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   5490
      Visible         =   0   'False
      Width           =   11550
      _ExtentX        =   20373
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
      RecordSource    =   "SELECT * FROM RESERVATION ORDER BY PNR"
      Caption         =   "Reservation Info"
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
   Begin MSAdodcLib.Adodc FlightInfo 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   5820
      Visible         =   0   'False
      Width           =   11550
      _ExtentX        =   20373
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
      RecordSource    =   "SELECT * FROM RouteInfo"
      Caption         =   "Flight Information"
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
   Begin MSDataListLib.DataCombo ServiceCode 
      Bindings        =   "frmReservation.frx":11D62
      DataField       =   "SS_CODE"
      DataSource      =   "MyAdodc"
      Height          =   330
      Left            =   6840
      TabIndex        =   19
      Top             =   5055
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   16711680
      ListField       =   "SS_CODE"
      Text            =   "DataCombo1"
      Object.DataMember      =   "ServiceInfo"
   End
   Begin VB.ComboBox cboOrigin 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   480
      Width           =   4335
   End
   Begin MSAdodcLib.Adodc Temp 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   5160
      Visible         =   0   'False
      Width           =   11550
      _ExtentX        =   20373
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
      RecordSource    =   "SELECT *  FROM RouteInfo"
      Caption         =   "Temp Information"
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Seat No:"
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
      Height          =   255
      Left            =   3720
      TabIndex        =   51
      Top             =   2400
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   3960
      X2              =   7560
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      DataField       =   "Route"
      DataSource      =   "Temp"
      Height          =   255
      Left            =   5880
      TabIndex        =   49
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      DataField       =   "Route Code"
      DataSource      =   "Temp"
      Height          =   255
      Left            =   6480
      TabIndex        =   48
      Top             =   7200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgMoveRecord 
      Height          =   315
      Index           =   0
      Left            =   4920
      Picture         =   "frmReservation.frx":11D7E
      ToolTipText     =   "First Record"
      Top             =   6000
      Width           =   315
   End
   Begin VB.Label lblMoveRecord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&End"
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   3
      Left            =   9660
      TabIndex        =   47
      ToolTipText     =   "Last Record"
      Top             =   6150
      Width           =   825
   End
   Begin VB.Image imgMoveRecord 
      Height          =   315
      Index           =   3
      Left            =   9360
      Picture         =   "frmReservation.frx":11E1C
      ToolTipText     =   "Last Record"
      Top             =   6000
      Width           =   315
   End
   Begin VB.Label lblMoveRecord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Next"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   8220
      TabIndex        =   46
      ToolTipText     =   "Next Record"
      Top             =   6150
      Width           =   675
   End
   Begin VB.Image imgMoveRecord 
      Height          =   315
      Index           =   2
      Left            =   7920
      Picture         =   "frmReservation.frx":11EBA
      ToolTipText     =   "Next Record"
      Top             =   6000
      Width           =   315
   End
   Begin VB.Label lblMoveRecord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Up"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   6600
      TabIndex        =   45
      ToolTipText     =   "Previous Record"
      Top             =   6180
      Width           =   885
   End
   Begin VB.Image imgMoveRecord 
      Height          =   315
      Index           =   1
      Left            =   6360
      Picture         =   "frmReservation.frx":11F58
      ToolTipText     =   "Previous Record"
      Top             =   6000
      Width           =   315
   End
   Begin VB.Label lblMoveRecord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Home"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   5220
      TabIndex        =   44
      ToolTipText     =   "First Record"
      Top             =   6150
      Width           =   885
   End
   Begin VB.Label lblExtraBagCharges 
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2520
      TabIndex        =   43
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label lblGrid 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Reservation Information"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   3960
      TabIndex        =   42
      Top             =   1680
      Width           =   3645
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Flight Reservation"
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
      Left            =   -120
      TabIndex        =   41
      Top             =   0
      Width           =   12255
   End
   Begin VB.Shape Shape6 
      BorderWidth     =   2
      Height          =   615
      Left            =   240
      Top             =   5880
      Width           =   3615
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   360
      Picture         =   "frmReservation.frx":11FF6
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   11295
   End
   Begin VB.Shape Shape5 
      BorderWidth     =   2
      Height          =   3735
      Left            =   5520
      Shape           =   4  'Rounded Rectangle
      Top             =   2160
      Width           =   5415
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Card No:"
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
      Index           =   11
      Left            =   5610
      TabIndex        =   20
      Top             =   5580
      Width           =   1245
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Code:"
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
      Index           =   10
      Left            =   5715
      TabIndex        =   18
      Top             =   5085
      Width           =   1140
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Passport &No:"
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
      Index           =   9
      Left            =   5805
      TabIndex        =   16
      Top             =   4575
      Width           =   1050
   End
   Begin VB.Label lblService 
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      DataField       =   "SS_DESC"
      DataMember      =   "ServiceInfo"
      DataSource      =   "envAirBusInfo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8280
      TabIndex        =   39
      Top             =   5040
      Width           =   2415
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address&3:"
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
      Index           =   8
      Left            =   5910
      TabIndex        =   14
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address&2:"
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
      Left            =   5910
      TabIndex        =   12
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address&1:"
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
      Left            =   5910
      TabIndex        =   10
      Top             =   2700
      Width           =   855
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
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
      Left            =   6180
      TabIndex        =   8
      Top             =   2325
      Width           =   510
   End
   Begin VB.Shape Shape4 
      BorderWidth     =   2
      Height          =   1335
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   4560
      Width           =   5175
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Branch Code:"
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
      Index           =   14
      Left            =   600
      TabIndex        =   26
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "&Class:"
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
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   3840
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   495
      Left            =   1680
      Top             =   3840
      Width           =   3615
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
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
      Left            =   840
      TabIndex        =   2
      Top             =   3360
      Width           =   795
   End
   Begin VB.Label lblFlightDate 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      DataField       =   "FLIGHT_DATE"
      DataSource      =   "MyAdodc"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   3360
      Width           =   3135
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Flight No:"
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
      Left            =   870
      TabIndex        =   0
      Top             =   2760
      Width           =   750
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reservation &No:"
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
      Left            =   345
      TabIndex        =   38
      Top             =   2325
      Width           =   1290
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   2295
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   2160
      Width           =   5175
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   0
      Picture         =   "frmReservation.frx":120EF
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   11535
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      DataField       =   "PNR"
      DataSource      =   "CancelReservation"
      Height          =   375
      Left            =   720
      TabIndex        =   36
      Top             =   8040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      DataField       =   "Route Code"
      DataSource      =   "FlightInfo"
      Height          =   375
      Left            =   2760
      TabIndex        =   35
      Top             =   2280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Extra Bags"
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
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total &Fare($):"
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
      Index           =   13
      Left            =   300
      TabIndex        =   34
      Top             =   6000
      Width           =   1065
   End
   Begin VB.Label lblDestination 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Destination:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   0
      TabIndex        =   28
      Top             =   960
      Width           =   1395
   End
   Begin VB.Label lblOrigin 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Origin:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   360
      TabIndex        =   27
      Top             =   480
      Width           =   1020
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   0
      Picture         =   "frmReservation.frx":121E8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12000
   End
   Begin VB.Shape shpMoveRecord 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   345
      Index           =   3
      Left            =   9360
      Shape           =   2  'Oval
      Top             =   6150
      Width           =   1155
   End
   Begin VB.Shape shpMoveRecord 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   345
      Index           =   2
      Left            =   7920
      Shape           =   2  'Oval
      Top             =   6150
      Width           =   1155
   End
   Begin VB.Shape shpMoveRecord 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   345
      Index           =   1
      Left            =   6360
      Shape           =   2  'Oval
      Top             =   6150
      Width           =   1155
   End
   Begin VB.Shape shpMoveRecord 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   345
      Index           =   0
      Left            =   4950
      Shape           =   2  'Oval
      Top             =   6150
      Width           =   1155
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   0
      Picture         =   "frmReservation.frx":12CF8
      Stretch         =   -1  'True
      Top             =   480
      Width           =   420
   End
End
Attribute VB_Name = "frmReservation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim RS As New ADODB.Recordset, CN As New ADODB.Connection
Dim Route As String, FlightNo As String, TotalCharge As Currency, ServiceCharge As Currency
Dim FlightClass As String, FlightDAte As Date



Private Sub cboFlightNo_Validate(Cancel As Boolean)
    If cboFlightNo.Text = "" Then
        Beep
        MsgBox "Please, Select Flight", vbExclamation, "Select Error"
        Cancel = True
    Else
        lblFlightDate.Caption = CDate(DisplayVal("Flight", cboFlightNo))
    End If
End Sub



Private Sub chkConfirm_Click()
    If cboFlightNo.Text = "" Or lblFlightDate.Caption = "" Then Exit Sub
    
    txtPASS_NAME.ForeColor = IIf(chkConfirm.Value = 1, vbGreen, vbYellow)
    txtPNR.ForeColor = IIf(chkConfirm.Value = 1, vbGreen, vbYellow)
    
    MyAdodc.Recordset.Fields("Pass_Status") = IIf(chkConfirm.Value = 1, "C", "W")
    
    Dim rsAirBus As New ADODB.Recordset
    
    rsAirBus.Open "select * from AllFlightInfo where Flight.FlightNo='" & cboFlightNo.Text & "'", MyAdodc.ConnectionString, adOpenDynamic, adLockPessimistic
        
    Dim SQL As String
    SQL = Temp.RecordSource
    
    Temp.RecordSource = "SELECT * From FLIGHT where FLIGHT.FLIGHTNO='" & cboFlightNo.Text & "' AND FLIGHT.FLIGHTDATE=#" & lblFlightDate.Caption & "#;"
    Temp.Refresh
    
    With Temp.Recordset
         If optClass(0).Value = True Then 'Economy
                If .RecordCount <> 0 Then
                 If .Fields(4) <= rsAirBus.Fields(3) Then
                     frmClassConfirm.Show vbModal
                     If Choice = "Cancel" Then
                         MyAdodc.Recordset.Delete
                         MyAdodc.Refresh
                         Exit Sub
                 End If
                 End If
              End If
         ElseIf optClass(1).Value = True Then 'Business
                If .RecordCount <> 0 Then
                 If .Fields(3) <= rsAirBus.Fields(2) Then
                     frmClassConfirm.Show vbModal
                     If Choice = "Cancel" Then
                         MyAdodc.Recordset.Delete
                         MyAdodc.Refresh
                         Exit Sub
                     End If
                End If
             End If
         ElseIf optClass(2).Value = True Then 'First Class
            If .RecordCount <> 0 Then
                 If .Fields(2) <= rsAirBus.Fields(1) Then
                     frmClassConfirm.Show vbModal
                     If Choice = "Cancel" Then
                         MyAdodc.Recordset.Delete
                         MyAdodc.Refresh
                         Exit Sub
                     End If
                End If
             End If
         End If
        rsAirBus.Close
        Set rsAirBus = Nothing
        Choice = ""

         If Choice = "Waiting" Then
             chkConfirm.Value = False
             Exit Sub
         End If
         
         If .RecordCount = 0 Then
            .AddNew
         
             .Fields(0) = cboFlightNo.Text
             .Fields(1) = lblFlightDate.Caption
         End If
         
         If optClass(0).Value = True Then
             .Fields(4) = Val(.Fields(4) & "") + 1
         ElseIf optClass(1).Value = True Then
             .Fields(3) = Val(.Fields(3) & "") + 1
         ElseIf optClass(2).Value = True Then
             .Fields(2) = Val(.Fields(2) & "") + 1
         End If
         
'         .Update
    End With
    Temp.RecordSource = SQL
    Temp.Refresh
    
    MsgBox "Ticket Confirmed", vbInformation, "Confirmed Ticket"
    
    SQL = "SELECT FLIGHT_SCH.*, FARE.*, BRANCH.*, RESERVATION.*, SERVICE.SS_DESC AS EXPR1 FROM RESERVATION, BRANCH, FLIGHT_SCH, FARE, SERVICE WHERE RESERVATION.BRANCH_CODE = BRANCH.BRANCH_CODE AND RESERVATION.FLIGHTNO = FLIGHT_SCH.FLIGHTNO AND FLIGHT_SCH.ROUTE_CODE = FARE.ROUTE_CODE AND RESERVATION.SS_CODE = SERVICE.SS_CODE AND RESERVATION.PNR=" & MyAdodc.Recordset.Fields(0)
    
    Unload envReport
   ' envReport.rsReservationRoute.Open SQL, envReport.cnReport, adOpenDynamic, adLockPessimistic
    
    'ReservationReport.Show
    'ReservationReport.PrintReport
    
    mPNR = 0
'    myadodc.Recordset.Update
End Sub
Private Sub cmdCancelReservation_Click()
On Error GoTo CancelReservationError
    frmPassword.Show vbModal
    If ValidPassWord = False Then Exit Sub

    If Val(txtPNR.Text) = 0 Then Exit Sub
    If MsgBox("Cancel this Reservation  for " & MyAdodc.Recordset.Fields("pass_name") _
        & vbNewLine & "Reservation No:- " & txtPNR.Text _
        & vbNewLine & vbTab & "Name:- " & txtPASS_NAME.Text _
        , vbYesNo + vbDefaultButton2 + vbQuestion, "Cancellation Confirmation") = vbNo Then Exit Sub
    
    Dim RsControl As New ADODB.Recordset
    RsControl.Open "select * from Control", MyAdodc.ConnectionString
        
    Dim DateDiff As Long, CancelCharge As Currency
    
    DateDiff = MyAdodc.Recordset.Fields(2) - Date
    If DateDiff <= 3 Then
        CancelCharge = DateDiff * (RsControl.Fields(7) / 100)
    ElseIf DateDiff <= 6 Then
        CancelCharge = DateDiff * (RsControl.Fields(6) / 100)
    ElseIf DateDiff > 6 Then
        CancelCharge = DateDiff * (RsControl.Fields(5) / 100)
    End If
    
    Dim SQL As String
    SQL = "INSERT INTO CANCELLATION " _
          & "  (PNR, FLIGHTNO, FLIGHT_DATE, CLASS, RESERV_DATE, " _
          & "   PASS_NAME, PASS_ADD1, PASS_ADD2, PASS_ADD3,  " _
           & " PASSPORT_NO, SS_CODE, CREDIT_CARD_NO, TOTAL_FARE, " _
          & " BRANCH_CODE,Cancel_date ,CancellationCharge) " _
        & " SELECT PNR, FLIGHTNO, FLIGHT_DATE, CLASS, RESERV_DATE, " _
           & " PASS_NAME, PASS_ADD1, PASS_ADD2, PASS_ADD3, " _
          & " PASSPORT_NO, SS_CODE, CREDIT_CARD_NO, " _
          & " TOTAL_FARE , BRANCH_CODE, Date()," & CancelCharge _
        & " From RESERVATION " _
            & " WHERE PNR =" & txtPNR.Text
            
    Dim CN As New ADODB.Connection
    CN.Open MyAdodc.ConnectionString
    CN.Execute SQL
    MyAdodc.Recordset.Delete
    MyAdodc.Refresh
    MyAdodc.Recordset.MoveFirst
    MsgBox "Reservation Cancelled for " & MyAdodc.Recordset.Fields("pass_name"), vbExclamation, "Cancel Reservation"
    
    
    SQL = "SELECT BRANCH.*, ROUTEINFO.*, FLIGHT_SCH.*, CANCELLATION.*, CANCELLATION.TOTAL_FARE - CANCELLATION.CancellationCharge AS Refund FROM CANCELLATION, BRANCH, FLIGHT_SCH, ROUTEINFO, RESERVATION WHERE CANCELLATION.BRANCH_CODE = BRANCH.BRANCH_CODE AND CANCELLATION.FLIGHTNO = FLIGHT_SCH.FLIGHTNO AND FLIGHT_SCH.ROUTE_CODE = ROUTEINFO.`Route Code`"
    
    
    Unload envReport
    envReport.rsCancelReservation.Open SQL, envReport.cnReport, adOpenDynamic, adLockPessimistic
    CancelReservationReport.Show
    CancelReservationReport.PrintReport
    
    Exit Sub
CancelReservationError:
    MsgBox "Your Reservation is Cancelled", vbCritical, "Cancel Reservation Error"
End Sub

Private Sub cmdRecord_Click(Index As Integer)
  Record Index
End Sub

Private Sub cmdShowRoute_Click()
On Error GoTo RouteError
    If cboOrigin.Text = "" Then
        MsgBox "Please select Source City", vbExclamation, "Select Error"
        Exit Sub
    ElseIf cboDestination.Text = "" Then
        MsgBox "Please select Destination City", vbExclamation, "Select Error"
        Exit Sub
    End If
    
    Dim SQL As String
    SQL = "SELECT * From ROUTEINFO " _
        & " WHERE (""Origin"" = '" & cboOrigin.Text & "') AND " _
        & " (""Destination"" = '" & cboDestination.Text & "') "
        
    If Not IsNull(FlightDAte.Value) Then _
        SQL = SQL & " AND " _
              & "(""Flight DAte"" = '" & FlightDAte.Value & "')"
    Temp.RecordSource = SQL
    Temp.Refresh
    
    If Temp.Recordset.RecordCount = 0 Then
        MsgBox "No Route found " & vbNewLine & "'Origin' " & vbTab & cboOrigin.Text _
            & vbNewLine & "'Destination' " & vbTab & cboDestination.Text, vbExclamation, "No Records"
    Else
        GridShow
    End If
    Exit Sub
RouteError:
    MsgBox Err.Description, vbCritical, "Route Error"
End Sub


Private Function FillValue(Table As String, cbo As ComboBox) As String
    RS.Open "SELECT * FROM " & Table, MyAdodc.ConnectionString, adOpenDynamic, adLockPessimistic
    Do Until RS.EOF
        cbo.AddItem RS.Fields(0)
        RS.MoveNext
    Loop
    cbo.Text = MyAdodc.Recordset.Fields(RS.Fields(0).Name)
    RS.Close
End Function

Public Function DisplayVal(Table As String, cbo As ComboBox) As Variant
    If cbo.Text = "" Then Exit Function
    
    RS.Open "select * from " & Table & " where  " & cbo.DataField & "='" & cbo.Text & "'", MyAdodc.ConnectionString, adOpenDynamic, adLockBatchOptimistic
    
    If Not RS.BOF Or Not RS.EOF Then DisplayVal = RS.Fields(1)
    RS.Close
End Function



Private Sub Form_Initialize()
    mPNR = MyAdodc.Recordset.Fields(0)
    CN.Open MyAdodc.ConnectionString
    
    FillValue "Branch", cboBranchCode
    
    FillValue "Flight", cboFlightNo
    
    FillCity "Origin", cboOrigin
    FillCity "Destination", cboDestination
    
    If MyAdodc.Recordset.AbsolutePosition = 1 Then
        MyAdodc.Recordset.MoveNext
        MyAdodc.Recordset.MovePrevious
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyChecker KeyCode
End Sub

Private Sub Form_Load()
    'FlightDAte.Value = ""
    Me.Caption = Me.Caption & " on " & Date
    cmdShowRoute.Visible = False
  lblFlightDate.Caption = Date
    
End Sub

Private Sub Form_Resize()
    Image3.Height = Me.Height
    Image3.Width = Me.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RS = Nothing
    Set CN = Nothing
    CloseForm
End Sub

Private Sub imgMoveRecord_Click(Index As Integer)
    MoveRecord Index, MyAdodc
End Sub





Private Sub lblMoveRecord_Click(Index As Integer)
    MoveRecord Index, MyAdodc
End Sub

Private Sub optClass_Click(Index As Integer)
    With MyAdodc.Recordset
        Select Case Index
            Case 0
                .Fields("CLASS") = "E"
            Case 1
                .Fields("CLASS") = "B"
            Case 2
                .Fields("CLASS") = "F"
        End Select
    End With
    FareCalc
End Sub
Private Sub MyAdodc_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo MoveError
    mPNR = pRecordset.Fields(0)
    With MyAdodc.Recordset
        MyAdodc.Caption = "Records:- " & .AbsolutePosition & "/" & .RecordCount
        If .EOF = True Or .BOF = True Then Exit Sub
        Select Case .Fields("class")
            Case "E"
                optClass.Item(0).Value = True
            Case "B"
                optClass.Item(1).Value = True
            Case "F"
                optClass.Item(2).Value = True
        End Select
    
       ' chkConfirm.Value = IIf(.Fields("Pass_Status") = "C", 1, 0)
    End With
    lblFlightDate.Caption = CDate(DisplayVal("Flight", cboFlightNo))
    txtPASS_NAME.ForeColor = IIf(chkConfirm.Value = 1, vbGreen, vbYellow)
    txtPNR.ForeColor = IIf(chkConfirm.Value = 1, vbGreen, vbYellow)
    
    lblTitle.Caption = "Reservation (" & MyAdodc.Recordset.AbsolutePosition & "/" _
        & MyAdodc.Recordset.RecordCount & ")'s Information"
    Exit Sub
MoveError:
    MsgBox "EOF OR BOF", vbCritical, "Record Move Error"
End Sub

Private Sub RouteGrid_Click()
    If MsgBox("Do you want to Select this flight", vbYesNo + vbQuestion, "Flight Selection") = vbYes Then
        RouteGrid.Col = 11
        cboFlightNo.Text = RouteGrid.Text
        NoGridShow
    End If
End Sub

Private Sub RouteGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    RouteGrid.Col = 0
    Route = RouteGrid.Text
    
    RouteGrid.Col = 11
    FlightNo = RouteGrid.Text
End Sub

Public Sub ServiceShow()
    On Error Resume Next
envAirBusInfo.rsServiceInfo.Find "ss_code='" & ServiceCode.Text & "'" _
       , 0, adSearchForward, 1
    lblService.Caption = envAirBusInfo.rsServiceInfo.Fields(1)
    
End Sub
Public Sub FillCity(CityType As String, cbo As ComboBox)
    RS.Open "SELECT DISTINCT " & CityType & " FROM Fare order by " & CityType, MyAdodc.ConnectionString, adOpenDynamic, adLockPessimistic
    Do Until RS.EOF
        cbo.AddItem RS.Fields(0)
        RS.MoveNext
    Loop
    RS.Close
End Sub
Public Sub GridShow()
    RouteGrid.Visible = True
    RouteGrid.Top = Shape2.Top
    RouteGrid.Height = Image4.Top - Image2.Top - 500
    MyAdodc.Visible = False
    chkConfirm.Visible = False
    cmdCancelReservation.Visible = False
    lblGrid.Caption = "Route Information"
End Sub

Public Sub NoGridShow()
    RouteGrid.Visible = False
    RouteGrid.Top = Shape2.Top
    
    MyAdodc.Visible = True
    chkConfirm.Visible = True
    cmdCancelReservation.Visible = True
    
    lblGrid.Caption = "Reservation Information"
End Sub

Public Sub NewReservation()
    AutoCode MyAdodc.Recordset, txtPNR, "P", MyAdodc.Recordset.Fields(0).ActualSize
End Sub

Private Sub ServiceCode_Change()
    If ServiceCode.Text = "" Then Exit Sub
    RS.Open "SELECT * FROM SERVICE WHERE SS_CODE='" & ServiceCode.Text & "'", MyAdodc.ConnectionString, adOpenDynamic, adLockReadOnly
    
    lblService.Caption = RS.Fields(1)
    
    RS.Close
End Sub

Private Sub Text1_LostFocus()
If (Text1.Text = "E%") Then
    optClass(2).Enabled = True
    ElseIf (Text1.Text = "B%") Then
        optClass(1).Enabled = True
        Else
        optClass(0).Enabled = True
        End If
  End If
  
End Sub

Private Sub txtBags_Change()
    ExtraBagCharge
End Sub
Public Sub ExtraBagCharge()
    Dim ExtraBag As Long
    With RS
        .Open "select * from CONTROL", MyAdodc.ConnectionString, adOpenDynamic, adLockReadOnly
        If optClass(2).Value = True Then
                If Val(txtBags.Text) > RS.Fields(2) Then ExtraBag = Val(txtBags.Text) - RS.Fields(2)
        ElseIf optClass(1).Value = True Then
                If Val(txtBags.Text) > RS.Fields(3) Then ExtraBag = Val(txtBags.Text) - RS.Fields(3)
        ElseIf optClass(0).Value = True Then
                If Val(txtBags.Text) > RS.Fields(4) Then ExtraBag = Val(txtBags.Text) - RS.Fields(4)
        End If
'        Dim Charge As Currency
'        Charge = ExtraBag * .Fields(1)
'        'lblExtraBagCharges.Caption = _
'        'IIf(Val(lblExtraBagCharges.Caption) > Charge, Val(lblExtraBagCharges.Caption) - Charge, _
'            Charge - Val(lblExtraBagCharges))
    lblExtraBagCharges.Caption = Val(Val(txtBags) * RS.Fields(1))
        txtTOTAL_FARE.Text = RS.Fields(0) + Val(ExtraBag) + Val(lblExtraBagCharges.Caption)
        .Close
    End With
    FareCalc
End Sub

Public Sub FareCalc()
    If Route = "" Then Exit Sub
    Dim FareRate As Currency
    
    Unload envAirBusInfo
    
    envAirBusInfo.FARE Route
    
    If Not envAirBusInfo.rsFare.EOF Then
        If optClass(2).Value = True Then
            FareRate = envAirBusInfo.rsFare.Fields(4)
        ElseIf optClass(1).Value = True Then
            FareRate = envAirBusInfo.rsFare.Fields(5)
        ElseIf optClass(0).Value = True Then
            FareRate = envAirBusInfo.rsFare.Fields(6)
        End If
    End If
    
    FareRate = Val(lblExtraBagCharges.Caption) + FareRate + ServiceCharge
    txtTOTAL_FARE.Text = FareRate
End Sub

Public Sub ServiceCheck()
    Dim RSService As New ADODB.Recordset
    RSService.Open "SELECT * FROM SERVICE", Temp.ConnectionString, adOpenDynamic, adLockBatchOptimistic
    
    RSService.Find "SS_CODE='" & ServiceCode.Text & "'", 0, adSearchForward, 1
    Dim Service As Currency
    Service = Val(RSService.Fields(RSService.Fields.Count - 1))
            
    txtTOTAL_FARE.Text = Val(txtTOTAL_FARE.Text) + Service - ServiceCharge
    ServiceCharge = Service
End Sub



Private Sub txtCREDIT_CARD_NO_KeyPress(KeyAscii As Integer)
If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) Then
MsgBox " Sorry!!.. you have to enter digits only!!..", vbQuestion, "Wait.."
KeyAscii = 0
End If

End Sub

