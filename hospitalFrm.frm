VERSION 5.00
Begin VB.Form hospitalFrm 
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
   Begin VB.Frame Frame1 
      BackColor       =   &H80000011&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   13455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin VB.Label timeLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   2640
         TabIndex        =   8
         Top             =   12360
         Width           =   90
      End
      Begin VB.Label dateLbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   600
         TabIndex        =   7
         Top             =   12360
         Width           =   90
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   3
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
         TabIndex        =   2
         Top             =   11760
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
         TabIndex        =   1
         Top             =   4320
         Width           =   4935
      End
   End
End
Attribute VB_Name = "hospitalFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub dashboardMenu_Click()
Unload Me
dashboardFrm.Show

End Sub

Private Sub logoutLbl_Click()
Unload Me
loginFrm.Show

End Sub

Private Sub vaccineMenu_Click()
Unload Me
vaccineStock.Show

End Sub
