VERSION 5.00
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkusercontrolsxp.ocx"
Begin VB.Form frmMsgBox 
   BorderStyle     =   0  'None
   ClientHeight    =   2760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   MDIChild        =   -1  'True
   ScaleHeight     =   2760
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   Begin vkUserContolsXP.vkCommand cmdOK 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   4200
      _ExtentX        =   7408
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
   End
   Begin vkUserContolsXP.vkFrame fMsgBox 
      Height          =   2760
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   4868
      Caption         =   "Jura"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TitleGradient   =   2
      TitleHeight     =   350
      BorderWidth     =   2
      Begin vkUserContolsXP.vkLabel lblMsg 
         Height          =   1275
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   600
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   2249
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
         Alignment       =   2
      End
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CreateRoundRectFromWindow Me, 7, 7
    Me.BackColor = mdiMain.BackColor
    Me.Top = (mdiMain.Height - Me.Height) / 5
    Me.Left = (mdiMain.Width - Me.Width) / 2
    Call frmColor(frmMsgBox)
End Sub
