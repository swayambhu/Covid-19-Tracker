VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form dashboardFrm 
   Caption         =   "COVID - 19 Dashboard"
   ClientHeight    =   12915
   ClientLeft      =   120
   ClientTop       =   270
   ClientWidth     =   23760
   Icon            =   "dashboardFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   12915
   ScaleWidth      =   23760
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H80000011&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   13455
      Left            =   0
      TabIndex        =   8
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
         TabIndex        =   17
         Top             =   5160
         Width           =   4935
      End
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
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   0
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   11760
         Width           =   4935
      End
      Begin VB.Image Image1 
         Height          =   1635
         Left            =   240
         Picture         =   "dashboardFrm.frx":10CA
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
         TabIndex        =   12
         Top             =   840
         Width           =   2355
      End
      Begin VB.Label loggedInUserLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         DataField       =   "userName"
         DataSource      =   "Adodc2"
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
         TabIndex        =   11
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
         BackColor       =   &H0000FF00&
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   3480
         Width           =   4935
      End
   End
   Begin VB.Frame dashboardFrame 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   12855
      Left            =   4920
      TabIndex        =   0
      Top             =   0
      Width           =   18855
      Begin VB.Frame drFrame 
         Height          =   7575
         Left            =   3840
         TabIndex        =   19
         Top             =   2280
         Visible         =   0   'False
         Width           =   9855
         Begin VB.CommandButton Command3 
            Caption         =   "HOSPITAL REPORT"
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   3240
            TabIndex        =   22
            Top             =   2520
            Width           =   3375
         End
         Begin VB.CommandButton Command2 
            Caption         =   "PATIENTS REPORT"
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   3240
            TabIndex        =   21
            Top             =   1200
            Width           =   3375
         End
         Begin VB.CommandButton Command1 
            Caption         =   "BACK"
            BeginProperty Font 
               Name            =   "Nirmala UI"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   3840
            TabIndex        =   20
            Top             =   6120
            Width           =   2415
         End
      End
      Begin VB.CommandButton dReport 
         Caption         =   "REPORTS"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   14520
         TabIndex        =   18
         Top             =   1440
         Width           =   2775
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   7320
         Top             =   600
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
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
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   1200
         Top             =   7080
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
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
         RecordSource    =   "covid_count"
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
      Begin MSChart20Lib.MSChart maharashtraBar 
         Height          =   5175
         Left            =   240
         OleObjectBlob   =   "dashboardFrm.frx":386D
         TabIndex        =   7
         Top             =   7440
         Width           =   18255
      End
      Begin MSChart20Lib.MSChart nagpurPie 
         Height          =   4215
         Left            =   12240
         OleObjectBlob   =   "dashboardFrm.frx":5A9F
         TabIndex        =   6
         Top             =   2400
         Width           =   6015
      End
      Begin MSChart20Lib.MSChart punePie 
         DragIcon        =   "dashboardFrm.frx":795E
         Height          =   4215
         Left            =   6240
         OleObjectBlob   =   "dashboardFrm.frx":8A28
         TabIndex        =   3
         Top             =   2400
         Width           =   6015
      End
      Begin MSChart20Lib.MSChart mumbaiPie 
         Height          =   4215
         Left            =   240
         OleObjectBlob   =   "dashboardFrm.frx":12F63
         TabIndex        =   5
         Top             =   2400
         Width           =   6015
      End
      Begin VB.Timer dateTmr 
         Interval        =   10
         Left            =   240
         Top             =   1920
      End
      Begin VB.Label dateLbl1 
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
         Left            =   16320
         TabIndex        =   16
         Top             =   360
         Width           =   645
      End
      Begin VB.Label timeLbl 
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
         Left            =   16320
         TabIndex        =   15
         Top             =   840
         Width           =   660
      End
      Begin VB.Line Line2 
         BorderColor     =   &H0000FF00&
         X1              =   480
         X2              =   2640
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label dsLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DASHBOARD"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   33.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   915
         Left            =   360
         TabIndex        =   4
         Top             =   120
         Width           =   4215
      End
      Begin VB.Label citiesLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOP 3 CITIES AFFECTED AS ON  :"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   480
         TabIndex        =   2
         Top             =   1320
         Width           =   6960
      End
      Begin VB.Label dateLbl 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   7800
         TabIndex        =   1
         Top             =   1320
         Width           =   4335
      End
   End
End
Attribute VB_Name = "dashboardFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset


Dim userName As String

Private Sub Command1_Click()
drFrame.Visible = False
End Sub

Private Sub Command2_Click()
DataReport1.Show
End Sub

Private Sub Command3_Click()
DataReport2.Show
End Sub

Private Sub dateTmr_Timer()
dateLbl.Caption = Date
dateLbl1.Caption = Date
timeLbl.Caption = Time

End Sub

Public Function loadData()

Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset
cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\VB Covid - 19 Tracker\covid.mdb;Persist Security Info=False"
cn.Open

rs.Open ("SELECT * from covid_count where city_name='Pune'"), cn, adOpenStatic, adLockReadOnly

With rs
    punePie.Column = 1
    punePie.Data = .Fields("death_count")
    punePie.ColumnLabel = "DEATH - " + punePie.Data
    
    punePie.Column = 2
    punePie.Data = .Fields("recover_count")
    punePie.ColumnLabel = "RECOVERED - " + punePie.Data
    
    
    punePie.Column = 3
    
    punePie.Data = .Fields("active_count")
    punePie.ColumnLabel = "ACTIVE - " + punePie.Data

End With
rs.Close

rs.Open ("SELECT * from covid_count where city_name='Mumbai'"), cn, adOpenStatic, adLockReadOnly

With rs
    mumbaiPie.Column = 1
    mumbaiPie.Data = .Fields("death_count")
    mumbaiPie.ColumnLabel = "DEATH - " + mumbaiPie.Data
    
    mumbaiPie.Column = 2
    mumbaiPie.Data = .Fields("recover_count")
    mumbaiPie.ColumnLabel = "RECOVERED - " + mumbaiPie.Data
    
    
    mumbaiPie.Column = 3
    
    mumbaiPie.Data = .Fields("active_count")
    mumbaiPie.ColumnLabel = "ACTIVE - " + mumbaiPie.Data

End With
rs.Close

rs.Open ("SELECT * from covid_count where city_name='Nagpur'"), cn, adOpenStatic, adLockReadOnly

With rs
    nagpurPie.Column = 1
    nagpurPie.Data = .Fields("death_count")
    nagpurPie.ColumnLabel = "DEATH - " + nagpurPie.Data
    
    nagpurPie.Column = 2
    nagpurPie.Data = .Fields("recover_count")
    nagpurPie.ColumnLabel = "RECOVERED - " + nagpurPie.Data
    
    
    nagpurPie.Column = 3
    
    nagpurPie.Data = .Fields("active_count")
    nagpurPie.ColumnLabel = "ACTIVE - " + nagpurPie.Data

End With
rs.Close

            


rs.Open ("select death_count from covid_count")

With rs
    maharashtraBar.Column = 1
    maharashtraBar.Data = 0
    .MoveFirst
    Do While Not .EOF
        
        maharashtraBar.Data = maharashtraBar.Data + .Fields("death_count")
        i = i + 1
        .MoveNext
    Loop

maharashtraBar.ColumnLabel = "DEATH - " + maharashtraBar.Data
End With
rs.Close

rs.Open ("select recover_count from covid_count")


With rs
    maharashtraBar.Column = 2
    maharashtraBar.Data = 0
    .MoveFirst
    Do While Not .EOF
        
        maharashtraBar.Data = maharashtraBar.Data + .Fields("recover_count")
        i = i + 1
        .MoveNext
    Loop

maharashtraBar.ColumnLabel = "Recover - " + maharashtraBar.Data
End With
rs.Close


rs.Open ("select active_count from covid_count")


With rs
    maharashtraBar.Column = 3
    maharashtraBar.Data = 0
    .MoveFirst
    Do While Not .EOF
        maharashtraBar.Data = maharashtraBar.Data + .Fields("active_count")
        i = i + 1
        .MoveNext
    Loop

maharashtraBar.ColumnLabel = "Active - " + maharashtraBar.Data
End With
rs.Close

rs.Open ("select total_count from covid_count")
With rs
    maharashtraBar.Column = 4
    maharashtraBar.Data = 0
    .MoveFirst
    Do While Not .EOF
        maharashtraBar.Data = maharashtraBar.Data + .Fields("total_count")
        i = i + 1
        .MoveNext
    Loop

maharashtraBar.ColumnLabel = "TOTAL - " + maharashtraBar.Data
End With
rs.Close

End Function



Private Sub dReport_Click()
drFrame.Visible = True
End Sub

Private Sub Form_Load()
loadData


End Sub

Private Sub Label2_Click()

cn.Close
Unload Me
loginFrm.Show

End Sub

Private Sub Label14_Click()
cn.Close
patientFrm.Show
Unload Me

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

Private Sub vaccineMenu_Click()
cn.Close

vaccineStock.Show
Unload Me


End Sub


