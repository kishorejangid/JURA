VERSION 5.00
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkusercontrolsxp.ocx"
Begin VB.Form frmEmailSettings 
   BorderStyle     =   0  'None
   Caption         =   "Settings"
   ClientHeight    =   3240
   ClientLeft      =   0
   ClientTop       =   -30
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEmailSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin vkUserContolsXP.vkFrame fEmailSetting 
      Height          =   3255
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5741
      Caption         =   "Settings"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleColor1     =   8438015
      TitleColor2     =   33023
      TitleGradient   =   2
      TitleHeight     =   360
      BorderColor     =   33023
      BorderWidth     =   2
      Begin JURA.StylerButton cmdClose 
         Height          =   255
         Left            =   4320
         TabIndex        =   10
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         Caption         =   "X"
         CaptionDisableColor=   12236471
         CaptionEffectColor=   16777215
         FocusDottedRect =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedValue    =   1
      End
      Begin vkUserContolsXP.vkCommand btnUpdate 
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   2640
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   661
         Caption         =   "Update"
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
         Left            =   1200
         TabIndex        =   3
         Top             =   2040
         Width           =   3375
         _ExtentX        =   5953
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkLabel lblPassword 
         Height          =   255
         Left            =   240
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2040
         Width           =   855
         _ExtentX        =   1508
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
      Begin vkUserContolsXP.vkTextBox txtUserName 
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   1560
         Width           =   3375
         _ExtentX        =   5953
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkLabel lblUserName 
         Height          =   255
         Left            =   240
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1560
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "User Name;"
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
      Begin vkUserContolsXP.vkTextBox txtPort 
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   960
         Width           =   3375
         _ExtentX        =   5953
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkLabel lblPort 
         Height          =   255
         Left            =   240
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   960
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Port:"
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
      Begin vkUserContolsXP.vkTextBox txtServer 
         Height          =   375
         Left            =   1200
         TabIndex        =   0
         Top             =   480
         Width           =   3375
         _ExtentX        =   5953
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkLabel lblServer 
         Height          =   255
         Left            =   240
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   480
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Server:"
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
Attribute VB_Name = "frmEmailSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnUpdate_Click()
    SaveSetting App.CompanyName, "Email", "Server", txtServer.Text
    SaveSetting App.CompanyName, "Email", "Port", txtPort.Text
    SaveSetting App.CompanyName, "Email", "UserName", txtUserName.Text
    SaveSetting App.CompanyName, "Email", "Password", txtPassword.Text
    Unload Me
End Sub

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub fEmailSetting_KeyDown(KeyCode As Integer, Shift As Integer)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_Load()
    CreateRoundRectFromWindow Me, 7, 7
    Me.Top = (mdiMain.Height - Me.Height) / 5
    Me.Left = (mdiMain.Width - Me.Width) / 2
    txtServer.Text = GetSetting(App.CompanyName, "Email", "Server", "")
    txtPort.Text = GetSetting(App.CompanyName, "Email", "Port", "")
    txtUserName.Text = GetSetting(App.CompanyName, "Email", "UserName", "")
    txtPassword.Text = GetSetting(App.CompanyName, "Email", "Password", "")
    Call frmColor(frmEmailSettings)
    Me.BackColor = mdiMain.BackColor
End Sub

Private Sub vkFrame1_KeyDown(KeyCode As Integer, Shift As Integer)

End Sub
