VERSION 5.00
Object = "{5A775647-818E-4061-92FE-2C097A4D7E15}#1.0#0"; "vkusercontrolsxp.ocx"
Begin VB.Form frmDBPass 
   Caption         =   "Database PassWord"
   ClientHeight    =   3270
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4665
   MaxButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4665
   StartUpPosition =   2  'CenterScreen
   Begin vkUserContolsXP.vkLabel lblRestart 
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   2640
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   "Restart The Software"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   33023
   End
   Begin vkUserContolsXP.vkTimer vkTimer1 
      Left            =   3960
      Top             =   2640
      _ExtentX        =   926
      _ExtentY        =   926
      Interval        =   1000
   End
   Begin vkUserContolsXP.vkFrame vkFrame1 
      Height          =   3285
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   5794
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleColor1     =   12640511
      TitleColor2     =   33023
      TitleGradient   =   2
      BorderColor     =   33023
      BorderWidth     =   2
      Begin vkUserContolsXP.vkLabel lblMessage 
         Height          =   615
         Left            =   360
         TabIndex        =   6
         Top             =   2040
         Visible         =   0   'False
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   1085
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Password Registered"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   33023
         Alignment       =   2
      End
      Begin vkUserContolsXP.vkCommand cmdOK 
         Height          =   495
         Left            =   480
         TabIndex        =   2
         Top             =   2160
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   873
         Caption         =   "OK"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   33023
         BorderColor     =   33023
         CustomStyle     =   0
      End
      Begin vkUserContolsXP.vkTextBox txtPassWord 
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   1200
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   33023
         PassWordChar    =   "*"
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkLabel vkLabel1 
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "User Name:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vkUserContolsXP.vkTextBox txtUser 
         Height          =   375
         Left            =   1440
         TabIndex        =   0
         Top             =   600
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   33023
         PassWordChar    =   "*"
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkLabel vkLabel2 
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1320
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Password:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmDBPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DBUser As String
Dim DBPassword As String
Private Sub cmdOK_Click()
    DBUser = txtUser.Text
    DBPassword = txtPassword.Text
    SaveSetting App.CompanyName, "DataBase", "DBUser", DBUser  'get the user name from the registry
    SaveSetting App.CompanyName, "DataBase", "DBPassword", DBPassword 'get the db password from the registry
    cmdOK.Visible = False
    lblMessage.Visible = True
    lblRestart.Visible = True
    vkTimer1.Enabled = True
End Sub

Private Sub Form_Load()
    Me.Width = 4905
    Me.Height = 3825
    Call frmColor(frmDBPass)
    DeleteSetting App.CompanyName, "DataBase", "LoginTable"
    DeleteSetting App.CompanyName, "DataBase", "StuddetailsTable"
    DeleteSetting App.CompanyName, "DataBase", "SubjTable"
    DeleteSetting App.CompanyName, "DataBase", "StudmarksTable"
End Sub
Private Sub vkTimer1_Timer()
    Unload Me
    End
End Sub

