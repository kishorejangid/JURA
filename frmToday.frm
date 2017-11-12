VERSION 5.00
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkusercontrolsxp.ocx"
Begin VB.Form frmToday 
   BorderStyle     =   0  'None
   ClientHeight    =   1335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2010
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   2010
   ShowInTaskbar   =   0   'False
   Begin vkUserContolsXP.vkTimer vkTimer1 
      Left            =   0
      Top             =   1680
      _ExtentX        =   926
      _ExtentY        =   926
      Interval        =   10
      Enabled         =   -1  'True
   End
   Begin vkUserContolsXP.vkLabel vkLabel1 
      Height          =   255
      Left            =   315
      TabIndex        =   5
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   "Kishore Jangid"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   33023
   End
   Begin vkUserContolsXP.vkTimer today_timer1 
      Left            =   2400
      Top             =   1680
      _ExtentX        =   926
      _ExtentY        =   926
      Interval        =   1000
   End
   Begin vkUserContolsXP.vkLabel today_label4 
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   720
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      BackColor       =   16777215
      BackStyle       =   0
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
   End
   Begin vkUserContolsXP.vkLabel today_label3 
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   "Time:"
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
   Begin vkUserContolsXP.vkLabel today_label2 
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BackColor       =   16777215
      BackStyle       =   0
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
   End
   Begin vkUserContolsXP.vkLabel today_label1 
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   "Date:"
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
   Begin vkUserContolsXP.vkFrame fToday 
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2000
      _ExtentX        =   3519
      _ExtentY        =   2355
      Caption         =   "TODAY"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TitleColor1     =   12640511
      TitleColor2     =   33023
      TitleGradient   =   2
      BorderColor     =   33023
   End
End
Attribute VB_Name = "frmToday"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim l As Integer
Private Sub Form_Load()
    CreateRoundRectFromWindow Me, 7, 7
    Call frmColor(frmToday)
    l = 1080
    today_timer1.Enabled = True
    today_label2.Caption = Date
    Me.BackColor = mdiMain.BackColor
    Me.Top = (Screen.Height - 3 * Me.Height) - 200
    Me.Left = Screen.Width - Me.Width - 200
End Sub

Private Sub fToday_MouseDown(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub today_timer1_Timer()
    today_label4.Caption = Time
End Sub
Private Sub vkTimer1_Timer()
    vkLabel1.Left = l
    l = l - 25
    If l < -1500 Then l = 2000
End Sub
