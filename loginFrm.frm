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
      Left            =   960
      Top             =   5520
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
      Height          =   4815
      Left            =   3000
      TabIndex        =   0
      Top             =   1560
      Width           =   7095
      Begin VB.CommandButton addNew 
         Caption         =   "ADD NEW"
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
         Left            =   2640
         TabIndex        =   13
         Top             =   3840
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox confirmPwdTxt 
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
         TabIndex        =   11
         Top             =   2880
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.CommandButton addUserBtn 
         Caption         =   "ADD USER"
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
         Left            =   1680
         TabIndex        =   10
         Top             =   2880
         Width           =   2175
      End
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
         Top             =   1920
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
         Left            =   4200
         TabIndex        =   3
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label incorrectConfirmPwd 
         AutoSize        =   -1  'True
         Caption         =   "*Password does not match"
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
         Left            =   4440
         TabIndex        =   16
         Top             =   3360
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.Label incorrectPwd 
         AutoSize        =   -1  'True
         Caption         =   "*Incorrect Password"
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
         Left            =   4440
         TabIndex        =   15
         Top             =   2400
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.Label incorrectUsrName 
         AutoSize        =   -1  'True
         Caption         =   "*Inccorrect Username"
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
         Left            =   4440
         TabIndex        =   14
         Top             =   1560
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.Label confirmPwdLbl 
         AutoSize        =   -1  'True
         Caption         =   "CONFIRM PASSWORD"
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
         TabIndex        =   12
         Top             =   2880
         Visible         =   0   'False
         Width           =   2790
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
         Left            =   1560
         TabIndex        =   5
         Top             =   1920
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
         Left            =   1560
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
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Function addNewUser()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\VB Covid - 19 Tracker\covid.mdb;Persist Security Info=False"
rs.ActiveConnection = con
rs.CursorLocation = adUseClient
rs.LockType = adLockOptimistic
rs.CursorType = adOpenDynamic
rs.Source = "select * from authentication where username ='" & usernameTxt.Text & "'"
rs.Open
If rs.EOF Then
    dashboardFrm.loggedInUserLbl.Caption = StrConv(usernameTxt.Text, vbUpperCase)
    Adodc2.Recordset.Fields("userName") = StrConv(usernameTxt.Text, vbUpperCase)
    Adodc2.Recordset.Update
    rs.addNew
    Call saveData
    rs.Update
    MsgBox "New User Added", vbOKOnly, "Add User"
    con.Close
    Unload Me
    dashboardFrm.Show

Else
    incorrectUsrName.Caption = "Username already exists"
    incorrectUsrName.Visible = True
    con.Close
End If
End Function

Function saveData()
rs.Fields(0) = usernameTxt.Text
rs.Fields(1) = passwordTxt.Text
End Function

Private Sub addNew_Click()
If usernameTxt.Text = "" And passwordTxt.Text = "" And confirmPwdTxt.Text = "" Then
    incorrectUsrName.Visible = True
    incorrectUsrName.Caption = "*Please Enter Username"
    incorrectPwd.Visible = True
    incorrectPwd.Caption = "*Please Enter Password"
    incorrectConfirmPwd.Visible = True
    incorrectConfirmPwd.Caption = "*Please Confirm Password"
ElseIf usernameTxt.Text = "" And passwordTxt.Text = "" Then
    incorrectUsrName.Visible = True
    incorrectUsrName.Caption = "*Please Enter Username"
    incorrectPwd.Visible = True
    incorrectPwd.Caption = "*Please Enter Password"
ElseIf usernameTxt.Text = "" And confirmPwdTxt.Text = "" Then
    incorrectUsrName.Visible = True
    incorrectUsrName.Caption = "*Please Enter Username"
    incorrectConfirmPwd.Visible = True
    incorrectConfirmPwd.Caption = "*Please Confirm Password"
ElseIf passwordTxt.Text = "" And confirmPwdTxt.Text = "" Then
    incorrectPwd.Visible = True
    incorrectPwd.Caption = "*Please Enter Password"
    incorrectConfirmPwd.Visible = True
    incorrectConfirmPwd.Caption = "*Please Confirm Password"
ElseIf usernameTxt.Text = "" Then
    incorrectUsrName.Visible = True
    incorrectUsrName.Caption = "*Please Enter Username"
ElseIf passwordTxt.Text = "" Then
    incorrectPwd.Visible = True
    incorrectPwd.Caption = "*Please Enter Password"
ElseIf confirmPwdTxt.Text = "" Then
    incorrectConfirmPwd.Visible = True
    incorrectConfirmPwd.Caption = "*Please Confirm Password"
ElseIf passwordTxt.Text <> confirmPwdTxt.Text Then
    incorrectConfirmPwd.Visible = True
    incorrectConfirmPwd.Caption = "*Password Does not Match"
Else
    Dim ans As String
    ans = InputBox("Please Enter Unique Key", "Add New User")
    If ans = "abcd1234g" Then
        Call addNewUser
    Else
        MsgBox "Invalid Key", vbCritical, "Access Denied"
    End If
End If
    
End Sub

Private Sub addUserBtn_Click()
addUserBtn.Visible = False
loginBtn.Visible = False
Frame1.Height = 4815
confirmPwdLbl.Visible = True
confirmPwdTxt.Visible = True
addNew.Visible = True

End Sub



Private Sub confirmPwdTxt_Change()
incorrectConfirmPwd.Visible = False
End Sub

Private Sub Form_Load()
Frame1.Height = 3855
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
    Unload Me
    dashboardFrm.Show
    dashboardFrm.loggedInUserLbl.Caption = StrConv(usernameTxt.Text, vbUpperCase)
    Adodc2.Recordset.Fields("userName") = StrConv(usernameTxt.Text, vbUpperCase)
    Adodc2.Recordset.Update
    
End If
End Function


Private Sub loginBtn_Click()
validateLogin

End Sub

Private Sub passwordTxt_Change()
incorrectPwd.Visible = False
End Sub

Private Sub passwordTxt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And confirmPwdTxt.Visible = False Then  ' The ENTER key.
    validateLogin
End If
End Sub

Private Sub Timer1_Timer()

dateLbl.Caption = Date
timeLbl.Caption = Time
End Sub


Private Sub usernameTxt_Change()
incorrectUsrName.Visible = False
End Sub

