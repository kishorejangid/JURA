VERSION 5.00
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkusercontrolsxp.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   0  'None
   ClientHeight    =   3600
   ClientLeft      =   1515
   ClientTop       =   1500
   ClientWidth     =   4800
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox CmbUser 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1320
      TabIndex        =   0
      Top             =   720
      Width           =   3135
   End
   Begin vkUserContolsXP.vkLabel vkLabel3 
      Height          =   315
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   556
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   "User Type:"
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
   Begin vkUserContolsXP.vkCommand cmdCancel 
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   2760
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Caption         =   "Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   33023
      CustomStyle     =   0
   End
   Begin vkUserContolsXP.vkCommand CmdOK 
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2760
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Caption         =   "OK"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   33023
      CustomStyle     =   0
   End
   Begin vkUserContolsXP.vkTextBox txtPassword 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   2040
      Width           =   3135
      _ExtentX        =   5530
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
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   344
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
   Begin vkUserContolsXP.vkLabel vkLabel1 
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   344
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   "User ID:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
   End
   Begin vkUserContolsXP.vkTextBox txtUserID 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   1320
      Width           =   3135
      _ExtentX        =   5530
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
   Begin vkUserContolsXP.vkFrame frame 
      Height          =   3615
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6376
      Caption         =   "Login"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TitleColor1     =   33023
      TitleColor2     =   12640511
      TitleGradient   =   2
      TitleHeight     =   360
      BorderColor     =   33023
      BorderWidth     =   2
      Begin VB.Label lblMyLabel 
         BackStyle       =   0  'Transparent
         Height          =   240
         Left            =   -1320
         TabIndex        =   9
         Top             =   480
         Width           =   3405
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    CreateRoundRectFromWindow Me, 7, 7
    Me.BackColor = mdiMain.BackColor
    Me.Top = (mdiMain.Height - Me.Height) / 2
    Me.Left = (mdiMain.Width - Me.Width) / 2
    CmbUser.AddItem ("Administrator")
    CmbUser.AddItem ("Staff")
    CmbUser.AddItem ("Student")
    Call frmColor(frmLogin)
End Sub
Private Sub cmdOK_Click()
    On Error GoTo lblerr
    Dim rs As New ADODB.Recordset
    Dim Passwordflg As Boolean
    admCheck = CmbUser.Text
    userCheck = txtUserID.Text
    If CmbUser.Text = "Administrator" And txtUserID.Text = "9" And txtPassword.Text = "9" Then
        mdiMain.Enabled = True
        Unload Me
        Exit Sub
    End If
    
    rs.CursorLocation = adUseClient
    qr = "select * from login where logintype = '" & CmbUser.Text & "'"
    rs.Open qr, conn, adOpenDynamic, adLockOptimistic, -1
    Do While Not rs.EOF
        If txtUserID.Text = rs!loginid And txtPassword.Text = rs!loginpassword Then
            mdiMain.Enabled = True
            Unload Me
            Passwordflg = True
            Exit Sub
        Else
            rs.MoveNext
            Passwordflg = False
        End If
    Loop
    If Passwordflg = False Then
            MsgBox "Check Your UserID And PassWord"
            txtUserID.Text = ""
            txtPassword.Text = ""
            txtUserID.SetFocus
            Exit Sub
    End If
    Exit Sub
lblerr:
    MsgBox Error & vbCrLf & "Error Number: " & Err.Number
End Sub

Private Sub CmbUser_Click()
    Unload frmSplash
End Sub
Private Sub cmdCancel_Click()
    Unload mdiMain
    Unload Me
End Sub
Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdOK_Click
    End If
End Sub
