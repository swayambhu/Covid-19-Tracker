VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form loginFrm 
   Caption         =   "Admin Login"
   ClientHeight    =   6690
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12000
   Icon            =   "loginFrm.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "loginFrm.frx":10CA
   ScaleHeight     =   6690
   ScaleWidth      =   12000
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   495
      Left            =   2040
      Top             =   5400
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      Left            =   9360
      Top             =   5280
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
      RecordSource    =   "select * from authentication"
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
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "ADMIN LOGIN"
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   3360
      TabIndex        =   0
      Top             =   1560
      Width           =   6735
      Begin VB.TextBox passwordTxt 
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
         IMEMode         =   3  'DISABLE
         Left            =   3360
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox usernameTxt 
         DataField       =   "userName"
         DataSource      =   "Adodc2"
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
         Left            =   3360
         TabIndex        =   1
         Top             =   1080
         Width           =   3015
      End
      Begin VB.CommandButton loginBtn 
         Appearance      =   0  'Flat
         Caption         =   "LOGIN"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2160
         TabIndex        =   3
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ADMIN LOGIN"
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
         Left            =   1800
         TabIndex        =   8
         Top             =   240
         Width           =   3120
      End
      Begin VB.Label passwordLbl 
         AutoSize        =   -1  'True
         Caption         =   "PASSWORD"
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
         Left            =   1320
         TabIndex        =   5
         Top             =   1680
         Width           =   1485
      End
      Begin VB.Label usernameLbl 
         AutoSize        =   -1  'True
         Caption         =   "USERNAME"
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
         Left            =   1320
         TabIndex        =   4
         Top             =   1080
         Width           =   1455
      End
   End
   Begin VB.Label loggedInUserLbl 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label dateLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   10440
      TabIndex        =   7
      Top             =   5640
      Width           =   60
   End
   Begin VB.Label timeLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   10440
      TabIndex        =   6
      Top             =   6120
      Width           =   60
   End
End
Attribute VB_Name = "loginFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
usernameTxt.Text = ""
End Sub

Public Function validateLogin()
Adodc1.RecordSource = "select * from authentication where username='" + usernameTxt.Text + "' and password='" + passwordTxt.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
    usernameTxt.Text = ""
    passwordTxt.Text = ""
    MsgBox "INVALID CREDENTIALS", vbCritical
    
Else
    MsgBox "Welcome " + usernameTxt.Text, vbOKOnly
    loginFrm.Hide
    dashboardFrm.Show
    Adodc2.Recordset.Fields("userName") = StrConv(usernameTxt.Text, vbUpperCase)
    Adodc2.Recordset.Update
    
End If
End Function


Private Sub loginBtn_Click()
validateLogin

End Sub

Private Sub passwordTxt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then  ' The ENTER key.
    validateLogin
End If
End Sub

Private Sub Timer1_Timer()

dateLbl.Caption = Date
timeLbl.Caption = Time
End Sub


