VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form patientFrm 
   Caption         =   "Form1"
   ClientHeight    =   12915
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   23760
   LinkTopic       =   "Form1"
   ScaleHeight     =   12915
   ScaleWidth      =   23760
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   12600
      Left            =   5040
      TabIndex        =   34
      Top             =   240
      Width           =   18855
      _ExtentX        =   33258
      _ExtentY        =   22225
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Nirmala UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "REGISTRATION"
      TabPicture(0)   =   "patientFrm.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "SET APPOINTMENT"
      TabPicture(1)   =   "patientFrm.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame9(1)"
      Tab(1).Control(1)=   "Frame18"
      Tab(1).Control(2)=   "Frame17"
      Tab(1).Control(3)=   "Frame16"
      Tab(1).Control(4)=   "Adodc3"
      Tab(1).Control(5)=   "Frame10"
      Tab(1).ControlCount=   6
      Begin VB.Frame Frame9 
         Caption         =   "VACCINE STOCK"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5655
         Index           =   1
         Left            =   -65040
         TabIndex        =   94
         Top             =   6600
         Width           =   8640
         Begin MSChart20Lib.MSChart vaccineChart 
            Height          =   4695
            Left            =   240
            OleObjectBlob   =   "patientFrm.frx":0038
            TabIndex        =   95
            Top             =   720
            Width           =   7935
         End
      End
      Begin VB.Frame Frame18 
         Caption         =   "BEDS AVAILABLE "
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   -74520
         TabIndex        =   93
         Top             =   8160
         Width           =   9375
         Begin MSChart20Lib.MSChart hospitalBedsgraph 
            Height          =   2895
            Left            =   360
            OleObjectBlob   =   "patientFrm.frx":1C8C
            TabIndex        =   96
            Top             =   480
            Width           =   8535
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "SEARCH"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -65160
         TabIndex        =   91
         Top             =   720
         Width           =   8775
         Begin VB.CommandButton Command10 
            Caption         =   "SEARCH"
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   78
            Top             =   1320
            Width           =   2415
         End
         Begin VB.TextBox psearchfname 
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   75
            Top             =   840
            Width           =   2835
         End
         Begin VB.TextBox psearchmname 
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3000
            TabIndex        =   76
            Top             =   840
            Width           =   2835
         End
         Begin VB.TextBox psearchlname 
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5880
            TabIndex        =   77
            Top             =   840
            Width           =   2835
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "PATIENT NAME [FIRST] - [MIDDLE] - [LAST]"
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   92
            Top             =   360
            Width           =   4485
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "APPOINTMENTS"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   -65160
         TabIndex        =   88
         Top             =   2880
         Width           =   8775
         Begin VB.CommandButton Command9 
            Caption         =   "TODAY'S APPOINTMENT"
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   90
            Top             =   600
            Width           =   2655
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "patientFrm.frx":350D
            Height          =   2055
            Left            =   0
            TabIndex        =   89
            Top             =   1320
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   3625
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   19
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16393
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16393
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
      End
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   615
         Left            =   -60480
         Top             =   5640
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\VB Covid - 19 Tracker\covid.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\VB Covid - 19 Tracker\covid.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from appointment"
         Caption         =   "Adodc3"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Frame Frame10 
         Caption         =   "SET APPOINTMENT"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7095
         Left            =   -74520
         TabIndex        =   62
         Top             =   720
         Width           =   9255
         Begin VB.CommandButton Command8 
            Caption         =   "SET APPOINTMENT"
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   480
            TabIndex        =   73
            Top             =   6120
            Width           =   2415
         End
         Begin VB.CommandButton Command7 
            Caption         =   "UPDATE APPOINTMENT"
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   5400
            TabIndex        =   74
            Top             =   6120
            Width           =   2415
         End
         Begin VB.Frame Frame13 
            Caption         =   "STATUS AND RESULT"
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1695
            Left            =   3360
            TabIndex        =   85
            Top             =   4200
            Width           =   5655
            Begin VB.Frame Frame15 
               Caption         =   "RESULT"
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1215
               Left            =   2520
               TabIndex        =   87
               Top             =   360
               Width           =   3015
               Begin VB.ComboBox result 
                  BeginProperty Font 
                     Name            =   "Nirmala UI"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   120
                  TabIndex        =   72
                  Text            =   "Combo1"
                  Top             =   480
                  Width           =   2775
               End
            End
            Begin VB.Frame Frame14 
               Caption         =   "STATUS"
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1215
               Left            =   240
               TabIndex        =   86
               Top             =   360
               Width           =   2055
               Begin VB.OptionButton scompleted 
                  Caption         =   "COMPLETED"
                  BeginProperty Font 
                     Name            =   "Nirmala UI"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   120
                  TabIndex        =   71
                  Top             =   720
                  Width           =   1455
               End
               Begin VB.OptionButton spending 
                  Caption         =   "PENDING"
                  BeginProperty Font 
                     Name            =   "Nirmala UI"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   70
                  Top             =   360
                  Width           =   1215
               End
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "APPOINTMENT DATE"
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1695
            Left            =   360
            TabIndex        =   84
            Top             =   4200
            Width           =   2775
            Begin MSComCtl2.DTPicker appointmentdate 
               Height          =   375
               Left            =   120
               TabIndex        =   69
               Top             =   600
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   661
               _Version        =   393216
               Format          =   125894657
               CurrentDate     =   44302
            End
         End
         Begin MSAdodcLib.Adodc hospitalData 
            Height          =   375
            Left            =   840
            Top             =   3960
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
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
            Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\VB Covid - 19 Tracker\covid.mdb;Persist Security Info=False"
            OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\VB Covid - 19 Tracker\covid.mdb;Persist Security Info=False"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "select * from hospital"
            Caption         =   "Adodc3"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin VB.Frame Frame11 
            Caption         =   "APPOINTMENT FOR?"
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   360
            TabIndex        =   81
            Top             =   2640
            Width           =   8655
            Begin VB.OptionButton covidTest 
               Caption         =   "COVID - 19 TEST"
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   1560
               TabIndex        =   67
               Top             =   600
               Width           =   2175
            End
            Begin VB.OptionButton covidvaccine 
               Caption         =   "COVID - 19 VACCINATION"
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   5040
               TabIndex        =   68
               Top             =   600
               Width           =   2655
            End
         End
         Begin VB.TextBox lastName 
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   6120
            TabIndex        =   66
            Top             =   1920
            Width           =   2835
         End
         Begin VB.TextBox middleName 
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3240
            TabIndex        =   65
            Top             =   1920
            Width           =   2835
         End
         Begin VB.TextBox firstName 
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   360
            TabIndex        =   64
            Top             =   1920
            Width           =   2835
         End
         Begin VB.ComboBox hospitalList 
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   1680
            TabIndex        =   63
            Text            =   "SELECT HOSPITAL"
            Top             =   840
            Width           =   3375
         End
         Begin VB.Label bedCount 
            Alignment       =   2  'Center
            BackColor       =   &H8000000A&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   7320
            TabIndex        =   83
            Top             =   720
            Width           =   945
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "BEDS AVAILABLE"
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5400
            TabIndex        =   82
            Top             =   840
            Width           =   1770
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "PATIENT NAME [FIRST] - [MIDDLE] - [LAST]"
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   360
            TabIndex        =   80
            Top             =   1440
            Width           =   4485
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "HOSPITAL"
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   360
            TabIndex        =   79
            Top             =   840
            Width           =   1065
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "PATIENT REGISTRATION"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   10815
         Left            =   360
         TabIndex        =   35
         Top             =   720
         Width           =   18135
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   16800
            Top             =   240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MSAdodcLib.Adodc Adodc2 
            Height          =   495
            Left            =   1200
            Top             =   7560
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   873
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   2
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
            Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\VB Covid - 19 Tracker\covid.mdb;Persist Security Info=False"
            OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\VB Covid - 19 Tracker\covid.mdb;Persist Security Info=False"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "patient"
            Caption         =   "Adodc2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin VB.CommandButton upploadbtn 
            Caption         =   "UPLOAD"
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   14400
            TabIndex        =   19
            Top             =   4200
            Width           =   3375
         End
         Begin VB.Frame Frame8 
            Height          =   2535
            Left            =   240
            TabIndex        =   57
            Top             =   8040
            Width           =   13215
            Begin VB.CommandButton Command6 
               Caption         =   "CLEAR ALL"
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   3120
               TabIndex        =   61
               Top             =   1560
               Width           =   2340
            End
            Begin VB.Frame Frame9 
               Caption         =   "SEARCH PATIENT"
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2175
               Index           =   0
               Left            =   6720
               TabIndex        =   58
               Top             =   240
               Width           =   6255
               Begin VB.CommandButton Command3 
                  Caption         =   "PRINT PATIENT DETAILS"
                  BeginProperty Font 
                     Name            =   "Nirmala UI"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Left            =   3360
                  TabIndex        =   60
                  Top             =   1560
                  Width           =   2460
               End
               Begin VB.CommandButton Command5 
                  Caption         =   "SEARCH"
                  BeginProperty Font 
                     Name            =   "Nirmala UI"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Left            =   480
                  TabIndex        =   26
                  Top             =   1560
                  Width           =   2655
               End
               Begin VB.TextBox searchLname 
                  Height          =   375
                  Left            =   4200
                  TabIndex        =   25
                  Top             =   960
                  Width           =   1950
               End
               Begin VB.TextBox searchMname 
                  Height          =   375
                  Left            =   2160
                  TabIndex        =   24
                  Top             =   960
                  Width           =   1950
               End
               Begin VB.TextBox searchFname 
                  Height          =   375
                  Left            =   120
                  TabIndex        =   23
                  Top             =   960
                  Width           =   1950
               End
               Begin VB.Label Label20 
                  AutoSize        =   -1  'True
                  Caption         =   "NAME [FIRST] - [MIDDLE] - [LAST]"
                  BeginProperty Font 
                     Name            =   "Nirmala UI"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   59
                  Top             =   480
                  Width           =   2940
               End
            End
            Begin VB.CommandButton Command4 
               Caption         =   "UPDATE CURRENT PATIENT"
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   3120
               TabIndex        =   22
               Top             =   480
               Width           =   2340
            End
            Begin VB.CommandButton Command2 
               Caption         =   "DELETE CURRENT PATIENT"
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   240
               TabIndex        =   21
               Top             =   1560
               Width           =   2340
            End
            Begin VB.CommandButton Command1 
               Caption         =   "ADD PATIENT"
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   240
               TabIndex        =   20
               Top             =   480
               Width           =   2340
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "CONTACT NUMBERS"
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2415
            Left            =   6960
            TabIndex        =   53
            Top             =   5040
            Width           =   6495
            Begin VB.TextBox alternateMob 
               Height          =   375
               Left            =   2400
               TabIndex        =   17
               Top             =   960
               Width           =   3735
            End
            Begin VB.TextBox email 
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2400
               TabIndex        =   18
               Top             =   1680
               Width           =   3735
            End
            Begin VB.TextBox mobNumber 
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2400
               TabIndex        =   16
               Top             =   360
               Width           =   3735
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "ALTERNATE MOBILE"
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   540
               TabIndex        =   56
               Top             =   960
               Width           =   1755
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "E-MAIL"
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1680
               TabIndex        =   55
               Top             =   1680
               Width           =   615
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "PHONE"
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1650
               TabIndex        =   54
               Top             =   360
               Width           =   645
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "ADDRESS"
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2415
            Left            =   240
            TabIndex        =   48
            Top             =   5040
            Width           =   5655
            Begin VB.TextBox pinCode 
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2160
               TabIndex        =   15
               Top             =   1800
               Width           =   2895
            End
            Begin VB.TextBox city 
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2160
               TabIndex        =   14
               Top             =   1320
               Width           =   2895
            End
            Begin VB.TextBox street 
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2160
               TabIndex        =   13
               Top             =   840
               Width           =   2895
            End
            Begin VB.TextBox houseNo 
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2160
               TabIndex        =   12
               Top             =   360
               Width           =   2895
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "PIN CODE"
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1080
               TabIndex        =   52
               Top             =   1815
               Width           =   870
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "CITY"
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1560
               TabIndex        =   51
               Top             =   1335
               Width           =   375
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "STREET"
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1320
               TabIndex        =   50
               Top             =   855
               Width           =   645
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "HOUSE NO."
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   960
               TabIndex        =   49
               Top             =   380
               Width           =   1035
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "NEXT OF KIN"
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2655
            Left            =   6960
            TabIndex        =   43
            Top             =   1680
            Width           =   6495
            Begin VB.TextBox kinEmail 
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2160
               TabIndex        =   11
               Top             =   2040
               Width           =   3855
            End
            Begin VB.TextBox kinPhone 
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2160
               TabIndex        =   10
               Top             =   1560
               Width           =   3855
            End
            Begin VB.TextBox kinRelation 
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2160
               TabIndex        =   9
               Top             =   1080
               Width           =   3855
            End
            Begin VB.TextBox kinName 
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2160
               TabIndex        =   8
               Top             =   600
               Width           =   3855
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "E - MAIL"
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1200
               TabIndex        =   47
               Top             =   2040
               Width           =   735
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "PHONE"
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1320
               TabIndex        =   46
               Top             =   1560
               Width           =   645
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "RELATION"
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1080
               TabIndex        =   45
               Top             =   1080
               Width           =   885
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "NAME"
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1440
               TabIndex        =   44
               Top             =   600
               Width           =   555
            End
         End
         Begin VB.Frame Frame4 
            Height          =   2535
            Left            =   240
            TabIndex        =   38
            Top             =   1800
            Width           =   5655
            Begin MSComCtl2.DTPicker dob 
               Height          =   375
               Left            =   2520
               TabIndex        =   7
               Top             =   1800
               Width           =   2895
               _ExtentX        =   5106
               _ExtentY        =   661
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   125894657
               CurrentDate     =   44302
            End
            Begin VB.ComboBox bloodList 
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2520
               TabIndex        =   6
               Text            =   "BLOOD GROUP"
               Top             =   1320
               Width           =   2895
            End
            Begin VB.ComboBox maritialList 
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2520
               TabIndex        =   5
               Text            =   "MARITAL STATUS"
               Top             =   840
               Width           =   2895
            End
            Begin VB.ComboBox genderList 
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2520
               TabIndex        =   4
               Text            =   "SELECT GENDER"
               Top             =   360
               Width           =   2895
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "DATE OF BIRTH"
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   42
               Top             =   1800
               Width           =   1335
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "BLOOD GROUP"
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   41
               Top             =   1320
               Width           =   1335
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "MARITAL STATUS"
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   40
               Top             =   840
               Width           =   1515
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "GENDER"
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   39
               Top             =   360
               Width           =   750
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   735
            Left            =   240
            TabIndex        =   36
            Top             =   720
            Width           =   13215
            Begin VB.TextBox lname 
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   9960
               TabIndex        =   3
               Top             =   120
               Width           =   2775
            End
            Begin VB.TextBox mname 
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   6960
               TabIndex        =   2
               Top             =   120
               Width           =   2775
            End
            Begin VB.TextBox fname 
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   3960
               TabIndex        =   1
               Top             =   120
               Width           =   2775
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "NAME [FIRST] - [MIDDLE] - [LAST]"
               BeginProperty Font 
                  Name            =   "Nirmala UI"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   240
               TabIndex        =   37
               Top             =   120
               Width           =   3540
            End
         End
         Begin VB.Image pImage 
            BorderStyle     =   1  'Fixed Single
            Height          =   3135
            Left            =   14400
            Stretch         =   -1  'True
            Top             =   840
            Width           =   3375
         End
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   5520
      Top             =   3720
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\VB Covid - 19 Tracker\covid.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\VB Covid - 19 Tracker\covid.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "loggedInUser"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000011&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   13455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H80000011&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HOSPITAL"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   855
         Left            =   0
         TabIndex        =   33
         Top             =   4320
         Width           =   4935
      End
      Begin VB.Label logoutLbl 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "LOGOUT"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   855
         Left            =   0
         TabIndex        =   32
         Top             =   11640
         Width           =   4935
      End
      Begin VB.Image Image1 
         Height          =   1635
         Left            =   240
         Picture         =   "patientFrm.frx":3522
         Stretch         =   -1  'True
         Top             =   720
         Width           =   1635
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WELCOME"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   675
         Left            =   2160
         TabIndex        =   31
         Top             =   840
         Width           =   2355
      End
      Begin VB.Label loggedInUserLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         DataField       =   "userName"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   450
         Left            =   2520
         TabIndex        =   30
         Top             =   1560
         Width           =   90
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000B&
         X1              =   0
         X2              =   5400
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Label dashboardMenu 
         Alignment       =   2  'Center
         BackColor       =   &H80000011&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DASHBOARD"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   0
         TabIndex        =   29
         Top             =   2640
         Width           =   4935
      End
      Begin VB.Label vaccineMenu 
         Alignment       =   2  'Center
         BackColor       =   &H80000011&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "VACCINE STOCK"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   855
         Left            =   0
         TabIndex        =   28
         Top             =   3480
         Width           =   4935
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PATIENTS DETAILS"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   855
         Left            =   0
         TabIndex        =   27
         Top             =   5160
         Width           =   4935
      End
   End
End
Attribute VB_Name = "patientFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String
Function loadVaccineGraph()
rs.Close


rs.Open "select * from supply"
Dim i As Integer
i = 1
Dim j As Integer
j = 1
With rs
    .MoveFirst
    Do While Not .EOF
    
        vaccineChart.RowCount = j
        vaccineChart.Row = i
        vaccineChart.RowLabel = rs!hospital_name
        

        vaccineChart.Column = 1
        vaccineChart.ColumnLabel = "COVAXIN"
        
        vaccineChart.Data = rs!covaxin
        vaccineChart.Column = 2
        vaccineChart.ColumnLabel = "COVISHEILD"
        vaccineChart.Data = rs!covisheild
        i = i + 1
        j = j + 1
        .MoveNext
    Loop
End With
rs.Close
rs.Open "select * from patient"
End Function
Function loadBedsGraph()
rs.Close
rs.Open "select * from hospital"
Dim i As Integer
i = 1
Dim j As Integer
j = 1
With rs
    .MoveFirst
    Do While Not .EOF
    
        hospitalBedsgraph.RowCount = j
        hospitalBedsgraph.Row = i
        hospitalBedsgraph.RowLabel = rs!hospital_name
        hospitalBedsgraph.Column = 1
        hospitalBedsgraph.ColumnLabel = rs!hospital_name
        hospitalBedsgraph.Data = rs!hospital_beds
    
        i = i + 1
        j = j + 1
        .MoveNext
    Loop
End With
rs.Close
rs.Open "select * from patient"
End Function

Function saveData()
rs.Fields(1) = fname.Text
rs.Fields(2) = mname.Text
rs.Fields(3) = lname.Text
rs.Fields(4) = genderList.Text
rs.Fields(5) = maritialList.Text
rs.Fields(6) = bloodList.Text
rs.Fields(7) = dob.Value
rs.Fields(8) = kinName.Text
rs.Fields(9) = kinRelation.Text
rs.Fields(10) = kinPhone.Text
rs.Fields(11) = kinEmail.Text
rs.Fields(12) = houseNo.Text
rs.Fields(13) = street.Text
rs.Fields(14) = city.Text
rs.Fields(15) = pinCode.Text
rs.Fields(16) = mobNumber.Text
rs.Fields(17) = alternateMob.Text
rs.Fields(18) = email.Text
rs.Fields(19) = str
End Function
Function clearPatient()
    fname.Text = ""
    mname.Text = ""
    lname.Text = ""
    genderList.Text = ""
    maritialList.Text = ""
    kinName.Text = ""
    kinPhone.Text = ""
    kinEmail.Text = ""
    kinRelation.Text = ""
    houseNo.Text = ""
    city.Text = ""
    street.Text = ""
    pinCode.Text = ""
    mobNumber.Text = ""
    alternateMob.Text = ""
    email.Text = ""
    str = ""
    bloodList.Text = "SELECT BLOOD GROUP"
    dob.Value = Date
    
    pImage.Picture = Nothing
    
    
    
    
End Function

Private Sub Combo1_Change()

End Sub

Private Sub Command1_Click()
If Not (fname.Text = "" Or mname.Text = "" Or lname.Text = "" Or genderList.Text = "SELECT GENDER" Or maritialList.Text = "MARITAL STATUS" Or genderList.Text = "" Or maritialList.Text = "" Or kinName.Text = "" Or kinPhone.Text = "" Or kinEmail.Text = "" Or kinRelation.Text = "" Or houseNo.Text = "" Or city.Text = "" Or street.Text = "" Or pinCode.Text = "" Or mobNumber.Text = "" Or alternateMob.Text = "" Or email.Text = "") Then
    rs.addNew
    Call saveData
    rs.Update
    Call clearPatient
    fname.SetFocus
    MsgBox "Patient Added", vbInformation, "Success"
Else
    MsgBox "Please fill al details", vbCritical, "Error"
End If

End Sub

Private Sub Command10_Click()
If psearchfname.Text = "" Or psearchmname.Text = "" Or psearchlname.Text = "" Then
    MsgBox "Please enter name for searching patient", vbInformation, "Error"
Else
    rs.Close
    rs.Open "select * from appointment where patient_name ='" & psearchfname.Text & " " & psearchmname.Text & " " & psearchlname.Text & "'"
    If Not rs.EOF Then
        rs.Close
        rs.Open "select * from appointment"
        hospitalList.Text = rs!hospital_name
        
        If rs!appointment_description = "COVID-19 TEST" Then
            covidTest.Value = True
        Else
            covidvaccine.Value = True
        End If
        Dim adate As String
        adate = CStr(rs!aapointment_date)
        appointmentdate.Value = adate
        If rs!appointment_description = "PENDING" Then
            spending.Value = True
        Else
            scompleted.Value = True
        End If
        result.Text = rs!covid_result
        rs.Close
        rs.Open "select * from patient where fname='" & psearchfname.Text & "' and Mname = '" & psearchmname.Text & "' and Lname = '" & psearchlname.Text & "'"
        firstName.Text = rs!fname
        middleName.Text = rs!mname
        lastName.Text = rs!lname
        rs.Close
        rs.Open "select * from hospital where hospital_name = '" & hospitalList.Text & "'"
        bedCount.Caption = rs!hospital_beds
        rs.Close
        rs.Open "select * from patient", cn, adOpenDynamic, adLockPessimistic
        
    End If
    
End If
End Sub



Private Sub Command11_Click()
DataReport1.Show

End Sub

Private Sub Command2_Click()
If Not (fname.Text = "" Or mname.Text = "" Or lname.Text = "" Or genderList.Text = "SELECT GENDER" Or maritialList.Text = "MARITAL STATUS" Or genderList.Text = "" Or maritialList.Text = "" Or kinName.Text = "" Or kinPhone.Text = "" Or kinEmail.Text = "" Or kinRelation.Text = "" Or houseNo.Text = "" Or city.Text = "" Or street.Text = "" Or pinCode.Text = "" Or mobNumber.Text = "" Or alternateMob.Text = "" Or email.Text = "") Then
    confirm = MsgBox("Do you want to Delete the Patient?", vbYesNo + vbCritical, "Delete Confirmation")
    
    If confirm = vbYes Then
        rs.Close
        rs.Open "select * from patient where fname='" & fname.Text & "' and Mname = '" & mname.Text & "' and Lname = '" & lname.Text & "'"
        rs.Delete adAffectCurrent
        MsgBox "Record has been Deleted successfully", vbInformation, "Message"
        Call clearPatient
        
        
        rs.Update
        rs.Close
        rs.Open "select * from patient"
        
    
    Else
        MsgBox "Profile not Deleted", vbInformation, "Message"
    End If
Else
    MsgBox "Select Profile for deleting", vbInformation, "Message"

    
End If
End Sub

Private Sub Command4_Click()
If Not (fname.Text = "" Or mname.Text = "" Or lname.Text = "" Or genderList.Text = "SELECT GENDER" Or maritialList.Text = "MARITAL STATUS" Or genderList.Text = "" Or maritialList.Text = "" Or kinName.Text = "" Or kinPhone.Text = "" Or kinEmail.Text = "" Or kinRelation.Text = "" Or houseNo.Text = "" Or city.Text = "" Or street.Text = "" Or pinCode.Text = "" Or mobNumber.Text = "" Or alternateMob.Text = "" Or email.Text = "") Then
    rs.Close
    rs.Open "select * from patient where fname='" & fname.Text & "' and Mname = '" & mname.Text & "' and Lname = '" & lname.Text & "'"
    Call saveData
    rs.Update
    Call clearPatient
    fname.SetFocus
    MsgBox "Patient Updated", vbInformation, "Success"
    rs.Close
    rs.Open "select * from patient"
Else
    MsgBox "Please fill al details", vbCritical, "Error"
End If
End Sub

Private Sub Command5_Click()
If searchFname.Text = "" Or searchMname.Text = "" Or searchLname.Text = "" Then
    MsgBox "Please fill all fields for searching the Patient!!", vbInformation, "Error"
Else
    rs.Close
    rs.Open "select * from patient where fname='" & searchFname.Text & "' and Mname = '" & searchMname.Text & "' and Lname = '" & searchLname.Text & "'"
    If Not rs.EOF Then
        fname.Text = rs!fname
        mname.Text = rs!mname
        lname.Text = rs!lname
        genderList.Text = rs!gender
        maritialList.Text = rs!mStatus
        bloodList.Text = rs!bloodgroup
        dob.Value = rs!dob
        kinName.Text = rs!Kname
        kinRelation.Text = rs!krelation
        kinPhone.Text = rs!kPhone
        kinEmail.Text = rs!kemail
        houseNo.Text = rs!house_no
        street.Text = rs!street
        city.Text = rs!city
        pinCode.Text = rs!pin
        mobNumber.Text = rs!phone
        alternateMob.Text = rs!aphone
        email.Text = rs!email
        str = rs!photo
        pImage.Picture = LoadPicture(str)
        searchFname.Text = ""
        searchMname.Text = ""
        searchLname.Text = ""
        rs.Close
        rs.Open "select * from patient"
    Else
        MsgBox "Patient not found", vbInformation, "Error"
        rs.Close
        rs.Open "select * from patient"
    End If
End If

    
End Sub

Private Sub Command6_Click()
Call clearPatient

End Sub

Private Sub Command7_Click()
If hospitalList.Text = "SELECT HOSPITAL" Or hospitalList.Text = "" Or firstName.Text = "" Or middleName.Text = "" Or lastName.Text = "" Or (covidTest.Value = False And covidvaccine.Value = False) Or (spending.Value = False And scompleted.Value = False) Or result.Text = "" Or result.Text = "SELECT RESULT" Then
    MsgBox "Please enter all fields", vbInformation, "Message"
ElseIf appointmentdate.Value <= Date Then
    MsgBox "Appointment Date Can't be today's or previous Date", vbInformation, "Error"
ElseIf appointmentdate.Value < Date And spending.Value = True Then
    MsgBox "Can't set status to pending after the appointment is done", vbInformation, "Message"
Else
    rs.Close
    rs.Open "select * from patient where fname='" & firstName.Text & "' and Mname='" & middleName.Text & "' and Lname='" & lastName.Text & "'"
    If rs.EOF Then
        MsgBox "Please register the patient first", vbInformation, "Patient Not Found"
        rs.Close
        rs.Open "select * from patient", cn, adOpenDynamic, adLockPessimistic
    Else
        If CInt(bedCount.Caption) <= 0 Then
            MsgBox "No bed available in " + hospitalList.Text + " hospital", vbInformation, "Message"
        Else
            Dim hID As Integer
            'get Hospital id
            rs.Close
            rs.Open "select * from hospital where hospital_name = '" & hospitalList.Text & "'"
            hID = rs!hospital_id
            hID = CInt(hID)
            rs.Close
        
            'Get patient id
            Dim pID As Integer
            Dim pname As String
            rs.Open "select * from patient where fname='" & firstName.Text & "' and Mname='" & middleName.Text & "' and Lname='" & lastName.Text & "'"
            pID = rs!id
            pname = rs!fname + " " + rs!mname + " " + rs!lname
            rs.Close
            
            rs.Open "select * from appointment", cn, adOpenDynamic, adLockPessimistic
            rs.Fields(1) = hID
            rs.Fields(2) = hospitalList.Text
            rs.Fields(3) = pID
            rs.Fields(4) = pname
            If covidTest.Value = True Then
                rs.Fields(5) = covidTest.Caption
            Else
                rs.Fields(5) = covidvaccine.Caption
            End If
            rs.Fields(6) = appointmentdate.Value
            If spending.Value = True Then
                rs.Fields(7) = spending.Caption
            Else
                rs.Fields(7) = scompleted.Caption
            End If
            rs.Fields(8) = result.Text
            rs.Update
            
            MsgBox "APPOINTMENT UPDATED !!!", vbInformation, "Success"
            rs.Close
            Adodc3.RecordSource = "select * from appointment"
            Adodc3.Refresh
            DataGrid1.Refresh
            firstName.Text = ""
            middleName.Text = ""
            lastName.Text = ""
            hospitalList.Text = "SELECT HOSPITAL"
            bedCount.Caption = ""
            result.Text = "SELECT RESULT"
            spending.Value = False
            scompleted.Value = False
            covidvaccine.Value = False
            covidTest.Value = False
            rs.Open "select * from patient", cn, adOpenDynamic, adLockPessimistic
            Call loadBedsGraph
            End If
        End If
End If



End Sub

Private Sub Command8_Click()
If hospitalList.Text = "SELECT HOSPITAL" Or hospitalList.Text = "" Or firstName.Text = "" Or middleName.Text = "" Or lastName.Text = "" Or (covidTest.Value = False And covidvaccine.Value = False) Or (spending.Value = False And scompleted.Value = False) Or result.Text = "" Or result.Text = "SELECT RESULT" Then
    MsgBox "Please enter all fields", vbInformation, "Message"
ElseIf appointmentdate.Value <= Date Then
    MsgBox "Appointment Date Can't be today's or previous Date", vbInformation, "Error"
ElseIf appointmentdate.Value < Date And spending.Value = True Then
    MsgBox "Can't set status to pending after the appointment is done", vbInformation, "Message"
Else
    rs.Close
    rs.Open "select * from patient where fname='" & firstName.Text & "' and Mname='" & middleName.Text & "' and Lname='" & lastName.Text & "'"
    If rs.EOF Then
        MsgBox "Please register the patient first", vbInformation, "Patient Not Found"
        rs.Close
        rs.Open "select * from patient", cn, adOpenDynamic, adLockPessimistic
    Else
        If CInt(bedCount.Caption) <= 0 Then
            MsgBox "No bed available in " + hospitalList.Text + " hospital", vbInformation, "Message"
        Else
            Dim hID As Integer
            'get Hospital id
            rs.Close
            rs.Open "select * from hospital where hospital_name = '" & hospitalList.Text & "'"
            hID = rs!hospital_id
            hID = CInt(hID)
            rs.Close
        
            'Get patient id
            Dim pID As Integer
            Dim pname As String
            rs.Open "select * from patient where fname='" & firstName.Text & "' and Mname='" & middleName.Text & "' and Lname='" & lastName.Text & "'"
            pID = rs!id
            pname = rs!fname + " " + rs!mname + " " + rs!lname
            rs.Close
            
            rs.Open "select * from appointment", cn, adOpenDynamic, adLockPessimistic
            rs.addNew
            rs.Fields(1) = hID
            rs.Fields(2) = hospitalList.Text
            rs.Fields(3) = pID
            rs.Fields(4) = pname
            If covidTest.Value = True Then
                rs.Fields(5) = covidTest.Caption
            Else
                rs.Fields(5) = covidvaccine.Caption
            End If
            rs.Fields(6) = appointmentdate.Value
            If spending.Value = True Then
                rs.Fields(7) = spending.Caption
            Else
                rs.Fields(7) = scompleted.Caption
            End If
            rs.Fields(8) = result.Text
            rs.Update
            
            MsgBox "Appointment Confirmed", vbInformation, "Success"
            rs.Close
            Adodc3.RecordSource = "select * from appointment"
            Adodc3.Refresh
            DataGrid1.Refresh
            Dim hcount As String
            
            rs.Open "select * from hospital where hospital_name = '" & hospitalList.Text & "'", cn, adOpenDynamic, adLockPessimistic
            hcount = rs!hospital_beds
            hcount = CInt(hcount) - 1
            rs.Fields(8) = hcount
            rs.Update
            rs.Close
            firstName.Text = ""
            middleName.Text = ""
            lastName.Text = ""
            hospitalList.Text = "SELECT HOSPITAL"
            bedCount.Caption = ""
            result.Text = "SELECT RESULT"
            spending.Value = False
            scompleted.Value = False
            covidvaccine.Value = False
            covidTest.Value = False
            rs.Open "select * from patient", cn, adOpenDynamic, adLockPessimistic
            Call loadBedsGraph
            End If
        End If
End If
End Sub

Private Sub Command9_Click()
Adodc3.Refresh
Adodc3.RecordSource = "select * from appointment where aapointment_date = '" & Date & "'"
Adodc3.Refresh
End Sub

Private Sub dashboardMenu_Click()
cn.Close
dashboardFrm.Show
Unload Me
End Sub

Private Sub Form_Load()



genderList.Text = "SELECT GENDER"
genderList.AddItem "M"
genderList.AddItem "F"
genderList.AddItem "Trans"
genderList.AddItem "LSB"
genderList.AddItem "GAY"
genderList.AddItem "Other"

maritialList.Text = "MARITAL STATUS"
maritialList.AddItem "M"
maritialList.AddItem "UM"

bloodList.AddItem "A+"
bloodList.AddItem "A-"
bloodList.AddItem "B+"
bloodList.AddItem "B-"
bloodList.AddItem "O+"
bloodList.AddItem "O-"
bloodList.AddItem "AB+"
bloodList.AddItem "AB-"

result.Text = "SELECT RESULT"
result.AddItem "PENDING"
result.AddItem "COVID-19 TEST NOT TAKEN"
result.AddItem "COVID-19 POSITIVE"
result.AddItem "COVID-19 NEGATIVE"
result.AddItem "RECOVERED FROM COVID-19"
result.AddItem "DEATH DUE TO COVID-19"

Call clearPatient
hospitalData.Refresh
With hospitalData.Recordset
    Do Until hospitalData.Recordset.EOF
        hospitalList.AddItem ![hospital_name]
        .MoveNext
    Loop
End With

cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\VB Covid - 19 Tracker\covid.mdb;Persist Security Info=False"
rs.Open "select * from patient", cn, adOpenDynamic, adLockPessimistic
Call loadVaccineGraph
Call loadBedsGraph
End Sub

Private Sub hospitalList_Click()
rs.Close
rs.Open "select * from hospital where hospital_name='" & hospitalList.Text & "'", cn, adOpenDynamic, adLockPessimistic
bedCount.Caption = rs!hospital_beds
rs.Close
rs.Open "select * from patient", cn, adOpenDynamic, adLockPessimistic
End Sub

Private Sub Label9_Click()
cn.Close
hospitalFrm.Show
Unload Me

End Sub

Private Sub logoutLbl_Click()
cn.Close
loginFrm.Show
Unload Me

End Sub

Private Sub upploadbtn_Click()
CommonDialog1.ShowOpen
CommonDialog1.Filter = "jpeg | *.jpg"
str = CommonDialog1.FileName
pImage.Picture = LoadPicture(str)

End Sub

Private Sub vaccineMenu_Click()
cn.Close
vaccineStock.Show
Unload Me
End Sub
