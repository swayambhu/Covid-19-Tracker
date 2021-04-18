VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form hospitalFrm 
   Caption         =   "Form1"
   ClientHeight    =   12915
   ClientLeft      =   -135
   ClientTop       =   210
   ClientWidth     =   23760
   LinkTopic       =   "Form1"
   ScaleHeight     =   12915
   ScaleWidth      =   23760
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   13320
      Top             =   720
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
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H80000011&
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
         TabIndex        =   46
         Top             =   5160
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
         TabIndex        =   20
         Top             =   3480
         Width           =   4935
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
         TabIndex        =   19
         Top             =   2640
         Width           =   4935
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000B&
         X1              =   0
         X2              =   5400
         Y1              =   2640
         Y2              =   2640
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
         TabIndex        =   18
         Top             =   1560
         Width           =   90
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
         TabIndex        =   17
         Top             =   840
         Width           =   2355
      End
      Begin VB.Image Image1 
         Height          =   1635
         Left            =   240
         Picture         =   "hospitalFrm.frx":0000
         Stretch         =   -1  'True
         Top             =   720
         Width           =   1635
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
         TabIndex        =   16
         Top             =   11640
         Width           =   4935
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
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
         TabIndex        =   15
         Top             =   4320
         Width           =   4935
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   11160
      Left            =   5040
      TabIndex        =   24
      Top             =   1320
      Width           =   18855
      _ExtentX        =   33258
      _ExtentY        =   19685
      _Version        =   393216
      Tabs            =   1
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
      TabCaption(0)   =   "ADD HOSPITAL"
      TabPicture(0)   =   "hospitalFrm.frx":27A3
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "hospitalval"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Image2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "hospital"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "DataGrid1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Adodc1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "hospitalSearch"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command4"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command3"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Command2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Command1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Frame3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Frame2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "hospitalName"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      Begin VB.TextBox hospitalName 
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3720
         TabIndex        =   1
         Top             =   1080
         Width           =   6975
      End
      Begin VB.Frame Frame2 
         Caption         =   "ADDRESS"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   960
         TabIndex        =   33
         Top             =   1800
         Width           =   9735
         Begin VB.TextBox plot 
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1200
            TabIndex        =   2
            Top             =   720
            Width           =   3255
         End
         Begin VB.TextBox street 
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6120
            TabIndex        =   3
            Top             =   720
            Width           =   3255
         End
         Begin VB.TextBox city 
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1200
            TabIndex        =   4
            Top             =   1440
            Width           =   3255
         End
         Begin VB.TextBox state 
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6120
            TabIndex        =   5
            Top             =   1440
            Width           =   3255
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PLOT"
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   42
            Top             =   720
            Width           =   660
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "STREET"
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5040
            TabIndex        =   41
            Top             =   720
            Width           =   915
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "STATE"
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   40
            Top             =   3600
            Width           =   780
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CITY"
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   39
            Top             =   1440
            Width           =   570
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "STATE"
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5160
            TabIndex        =   38
            Top             =   1440
            Width           =   780
         End
         Begin VB.Label plotval 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*Required"
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   3480
            TabIndex        =   37
            Top             =   1200
            Width           =   915
         End
         Begin VB.Label streetval 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*Required"
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   8460
            TabIndex        =   36
            Top             =   1200
            Width           =   915
         End
         Begin VB.Label cityval 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "*Required"
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   3480
            TabIndex        =   35
            Top             =   1920
            Width           =   915
         End
         Begin VB.Label stateval 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*Required"
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   8400
            TabIndex        =   34
            Top             =   1920
            Width           =   915
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "CONTACT"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   960
         TabIndex        =   26
         Top             =   4560
         Width           =   9735
         Begin VB.TextBox phoneNumber 
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2520
            TabIndex        =   6
            Top             =   720
            Width           =   3375
         End
         Begin VB.TextBox email 
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   2520
            TabIndex        =   8
            Top             =   1560
            Width           =   3375
         End
         Begin VB.TextBox beds 
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   8040
            TabIndex        =   7
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL BEDS"
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6240
            TabIndex        =   32
            Top             =   720
            Width           =   1560
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CONTACT NO."
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   31
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "EMAIL ID"
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   30
            Top             =   1560
            Width           =   1140
         End
         Begin VB.Label phonenumberval 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*Required"
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   4920
            TabIndex        =   29
            Top             =   1200
            Width           =   915
         End
         Begin VB.Label bedsval 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "*Required"
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   8520
            TabIndex        =   28
            Top             =   1200
            Width           =   915
         End
         Begin VB.Label emailval 
            AutoSize        =   -1  'True
            Caption         =   "*Required"
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   5880
            TabIndex        =   27
            Top             =   1800
            Width           =   930
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ADD"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         TabIndex        =   9
         Top             =   7080
         Width           =   1415
      End
      Begin VB.CommandButton Command2 
         Caption         =   "UPDATE"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   10
         Top             =   7080
         Width           =   1415
      End
      Begin VB.CommandButton Command3 
         Caption         =   "DELETE"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6240
         TabIndex        =   11
         Top             =   7080
         Width           =   1415
      End
      Begin VB.CommandButton Command4 
         Caption         =   "CLEAR"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8160
         TabIndex        =   12
         Top             =   7080
         Width           =   1415
      End
      Begin VB.TextBox hospitalSearch 
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         TabIndex        =   13
         Top             =   8040
         Width           =   6975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "SEARCH"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   960
         TabIndex        =   14
         Top             =   7920
         Width           =   2295
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   720
         Top             =   7080
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "hospitalFrm.frx":27BF
         Height          =   2055
         Left            =   0
         TabIndex        =   25
         Top             =   9120
         Width           =   18735
         _ExtentX        =   33046
         _ExtentY        =   3625
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   23
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
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
      Begin MSAdodcLib.Adodc hospital 
         Height          =   375
         Left            =   960
         Top             =   4080
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HOSPITAL NAME"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   45
         Top             =   1080
         Width           =   2130
      End
      Begin VB.Image Image2 
         Height          =   8085
         Left            =   11280
         Picture         =   "hospitalFrm.frx":27D6
         Stretch         =   -1  'True
         Top             =   360
         Width           =   7560
      End
      Begin VB.Label hospitalval 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "*Required"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   9720
         TabIndex        =   44
         Top             =   1560
         Width           =   915
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "*SEARCH BY HOSPITAL NAME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   7800
         TabIndex        =   43
         Top             =   8520
         Width           =   2700
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "HOSPTIAL MANAGEMENT"
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5160
      TabIndex        =   23
      Top             =   120
      Width           =   6735
   End
   Begin VB.Label dateLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "date"
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   21360
      TabIndex        =   22
      Top             =   240
      Width           =   645
   End
   Begin VB.Label timelbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "time"
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   21360
      TabIndex        =   21
      Top             =   720
      Width           =   660
   End
End
Attribute VB_Name = "hospitalFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Function saveHospital()
rs.Fields(1) = hospitalName.Text
rs.Fields(2) = plot.Text
rs.Fields(3) = street.Text
rs.Fields(4) = city.Text
rs.Fields(5) = state.Text
rs.Fields(6) = email.Text
rs.Fields(7) = phoneNumber.Text
rs.Fields(8) = beds.Text
End Function
Function clearHospital()
hospitalName.Text = ""
plot.Text = ""
street.Text = ""
city.Text = ""
state.Text = ""
email.Text = ""
phoneNumber.Text = ""
beds.Text = ""

End Function

Function hosSearch()
If hospitalSearch.Text = "" Then
    MsgBox "Please enter Hospital Name fetching Information", vbCritical, "Empty Paramters"
Else
    rs.Close
    rs.Open "select * from hospital where hospital_name='" & hospitalSearch.Text & "'", cn, adOpenDynamic, adLockPessimistic
    If Not rs.EOF Then
        hospitalName.Text = rs!hospital_name
        plot.Text = rs!plot
        street.Text = rs!street
        city.Text = rs!city
        state.Text = rs!state
        email.Text = rs!Hospital_email
        phoneNumber.Text = rs!hospital_contact
        beds.Text = rs!hospital_beds
    Else
        MsgBox "Hospital Not found", vbCritical, "Error"
    End If
    
End If

End Function

Private Sub beds_Change()
If Not IsNumeric(beds.Text) Then
    beds.Text = ""
    bedsval.Caption = "*Only numbers are allowed"
    bedsval.Visible = True
ElseIf beds.Text = "" Then
    bedsval.Caption = "*Required"
    bedsval.Visible = True
Else
    bedsval.Visible = False
End If

End Sub

Private Sub city_Change()
If city.Text = "" Then
    cityval.Caption = "*Required"
    cityval.Visible = True
Else
    cityval.Visible = False
End If
    
End Sub

Private Sub Command1_Click()
If hospitalName.Text = "" Or plot.Text = "" Or street.Text = "" Or city.Text = "" Or state.Text = "" Or email.Text = "" Or beds.Text = "" Then
    MsgBox "Please enter all fields", vbExclamation, "Message"
ElseIf Not isEmail(email.Text) Then
    MsgBox "Please enter valid email id", vbCritical, "Invalid Email"
ElseIf Len(phoneNumber.Text) < 8 Or Len(phoneNumber.Text) > 11 Then
    MsgBox "Contact Number shall be between 7 to 11 digits", vbCritical, "Invalid Mobile"
Else
    rs.Close
    rs.Open "select * from hospital where hospital_name ='" & hospitalName.Text & "'"
    If rs.EOF Then
        rs.addNew
        Call saveHospital
        
        rs.Update
        Call clearHospital
        
        hospital.Refresh
        MsgBox "New Hospital added Successfully...!!", vbInformation, "Success"
        rs.Close
        rs.Open "select * from hospital"
        hospitalName.SetFocus
    Else
        hospitalval.Visible = True
        hospitalval.Caption = "*Hospial already exists"
        rs.Close
        rs.Open "select * from hospital"
    End If
    
End If

End Sub


Private Sub Command2_Click()
If hospitalName.Text = "" Or plot.Text = "" Or street.Text = "" Or city.Text = "" Or state.Text = "" Or email.Text = "" Or beds.Text = "" Then
    MsgBox "Please enter all fields", vbExclamation, "Message"
Else
        rs.Fields(1) = hospitalName.Text
        rs.Fields(2) = plot.Text
        rs.Fields(3) = street.Text
        rs.Fields(4) = city.Text
        rs.Fields(5) = state.Text
        rs.Fields(6) = email.Text
        rs.Fields(7) = phoneNumber.Text
        rs.Fields(8) = beds.Text
        rs.Update
        hospital.Refresh
        Call clearHospital
        
        MsgBox "Data updated"
        rs.Close
End If

End Sub

Private Sub Command3_Click()
If Not (hospitalName.Text = "" Or plot.Text = "" Or street.Text = "" Or city.Text = "" Or state.Text = "" Or email.Text = "" Or beds.Text = "") Then
    confirm = MsgBox("Do you want to Delete the Dealer?", vbYesNo + vbCritical, "Delete Confirmation")
    If confirm = vbYes Then
        rs.Delete adAffectCurrent
        MsgBox "Record has been Deleted successfully", vbInformation, "Message"
        Call clearHospital
        
        rs.Update
    Else
        MsgBox "Profile not Deleted", vbInformation, "Message"
    End If
Else
    MsgBox "Select Profile for deleting", vbInformation, "Message"

    
End If
End Sub

Private Sub Command4_Click()
Call clearHospital

End Sub

Private Sub Command5_Click()
Call hosSearch
hospitalSearch.Text = ""
End Sub

Private Sub dashboardMenu_Click()
cn.Close
Unload Me
dashboardFrm.Show

End Sub

Private Sub email_Change()
If email.Text = "" Then
    emailval.Visible = True
    emailval.Caption = "*Required"
ElseIf Not isEmail(email.Text) Then
    emailval.Visible = True
    emailval.Caption = "*Invalid Email Id"
Else
    emailval.Visible = False
End If
End Sub

Public Function isEmail(email As String) As Boolean
Dim myAt As Integer
Dim myDot As Integer
Dim myDotDot As Integer

isEmail = True
myAt = InStr(1, email, "@", vbTextCompare)
myDot = InStr(myAt + 2, email, ".", vbTextCompare)
myDotDot = InStr(myAt + 2, email, "..", vbTextCompare)
If myAt = 0 Or myDot = 0 Or Not myDotDot = 0 Or Right(email, 1) = "." Then isEmail = False
End Function



Private Sub Form_Load()

cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\VB Covid - 19 Tracker\covid.mdb;Persist Security Info=False"
rs.ActiveConnection = cn
rs.CursorLocation = adUseClient
rs.LockType = adLockOptimistic
rs.CursorType = adOpenDynamic
rs.Open "select * from hospital"
End Sub

Private Sub hospitalName_Change()
If hospitalName.Text = "" Then
    hospitalval.Caption = "*Required"
    hospitalval.Visible = True
Else
    hospitalval.Visible = False
End If
End Sub

Private Sub Label14_Click()
patientFrm.Show
cn.Close
Unload Me

End Sub

Private Sub logoutLbl_Click()
Unload Me
loginFrm.Show

End Sub

Private Sub phoneNumber_Change()
If Not IsNumeric(phoneNumber.Text) Then
    phoneNumber.Text = ""
    phonenumberval.Caption = "*Only Numbers are allowed"
ElseIf phoneNumber.Text = "" Then
    phonenumberval.Caption = "*Required"
    phonenumberval.Visible = True
Else
    phonenumberval.Visible = False
End If

End Sub

Private Sub plot_Change()
If plot.Text = "" Then
    plotval.Caption = "*Required"
    plotval.Visible = True
Else
    plotval.Visible = False
End If

End Sub

Private Sub state_Change()
If state.Text = "" Then
    stateval.Caption = "*Required"
    stateval.Visible = True
Else
    stateval.Visible = False
End If

End Sub

Private Sub street_Change()
If street.Text = "" Then
    streetval.Caption = "*Required"
    streetval.Visible = True
Else
    streetval.Visible = False
End If
End Sub

Private Sub Timer2_Timer()
dateLbl.Caption = Date
timelbl1.Caption = Time
End Sub

Private Sub vaccineMenu_Click()
cn.Close
Unload Me
vaccineStock.Show

End Sub
