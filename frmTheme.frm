VERSION 5.00
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkusercontrolsxp.ocx"
Begin VB.Form frmTheme 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3120
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2880
   ScaleWidth      =   3120
   ShowInTaskbar   =   0   'False
   Begin vkUserContolsXP.vkFrame fTheme 
      Height          =   2895
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   5106
      Caption         =   "Color"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TitleColor1     =   12640511
      TitleColor2     =   33023
      TitleGradient   =   2
      TitleHeight     =   350
      BorderColor     =   33023
      Begin vkUserContolsXP.vkOptionButton optGreen 
         Height          =   255
         Left            =   1920
         TabIndex        =   8
         Top             =   1320
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Green"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Group           =   1
      End
      Begin vkUserContolsXP.vkOptionButton optBlack 
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   1320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Black"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Group           =   1
      End
      Begin vkUserContolsXP.vkOptionButton optBlue 
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Blue"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Group           =   1
      End
      Begin vkUserContolsXP.vkOptionButton optOrange 
         Height          =   255
         Left            =   360
         TabIndex        =   0
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Orange"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   1
         Group           =   1
      End
      Begin vkUserContolsXP.vkOptionButton optViolet 
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   1680
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Violet"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Group           =   1
      End
      Begin vkUserContolsXP.vkOptionButton optBrown 
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         Top             =   960
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Brown"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Group           =   1
      End
      Begin vkUserContolsXP.vkOptionButton optCyan 
         Height          =   255
         Left            =   1920
         TabIndex        =   4
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Cyan"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Group           =   1
      End
      Begin vkUserContolsXP.vkCommand cmdClose 
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   2160
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   873
         Caption         =   "Close"
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
   End
End
Attribute VB_Name = "frmTheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CreateRoundRectFromWindow Me, 7, 7
    Me.Top = (mdiMain.Height - Me.Height) / 5
    Me.Left = (mdiMain.Width - Me.Width) / 2
    Me.BackColor = mdiMain.BackColor
    Call frmColor(frmTheme)
    If Color = Black Then
        frmTheme.optBlack.Value = vbChecked
    ElseIf Color = Blue Then
        frmTheme.optBlue.Value = vbChecked
    ElseIf Color = Cyan Then
        frmTheme.optCyan.Value = vbChecked
    ElseIf Color = Red Then
        frmTheme.optBrown.Value = vbChecked
    ElseIf Color = Violet Then
        frmTheme.optViolet.Value = vbChecked
    ElseIf Color = Orange Then
        frmTheme.optOrange.Value = vbChecked
    ElseIf Color = Green Then
        frmTheme.optGreen.Value = vbChecked
    End If
End Sub




Private Sub fTheme_MouseDown(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub optBlack_Change(Value As CheckBoxConstants)
    If optBlack.Value = vbChecked Then
        Color = Black
        lColor = LBlack
        Call ActiveFormClr(Color, lColor)
    End If
    SaveSetting App.CompanyName, "Theme", "frmColor", Black
    SaveSetting App.CompanyName, "Theme", "frmlColor", LBlack
End Sub

Private Sub optBlue_Change(Value As CheckBoxConstants)
    If optBlue.Value = vbChecked Then
        Color = Blue
        lColor = LBlue
        Call ActiveFormClr(Color, lColor)
    End If
    SaveSetting App.CompanyName, "Theme", "frmColor", Blue
    SaveSetting App.CompanyName, "Theme", "frmlColor", LBlue
End Sub

Private Sub optBrown_Change(Value As CheckBoxConstants)
    If optBrown.Value = vbChecked Then
        Color = Brown
        lColor = LBrown
        Call ActiveFormClr(Color, lColor)
    End If
    SaveSetting App.CompanyName, "Theme", "frmColor", Brown
    SaveSetting App.CompanyName, "Theme", "frmlColor", LBrown
End Sub

Private Sub optCyan_Change(Value As CheckBoxConstants)
    If optCyan.Value = vbChecked Then
        Color = Cyan
        lColor = LCyan
        Call ActiveFormClr(Color, lColor)
    End If
    SaveSetting App.CompanyName, "Theme", "frmColor", Cyan
    SaveSetting App.CompanyName, "Theme", "frmlColor", LCyan
End Sub

Private Sub optGreen_Change(Value As CheckBoxConstants)
 If optGreen.Value = vbChecked Then
        Color = Green
        lColor = LGreen
        Call ActiveFormClr(Color, lColor)
    End If
    SaveSetting App.CompanyName, "Theme", "frmColor", Green
    SaveSetting App.CompanyName, "Theme", "frmlColor", LGreen
End Sub

Private Sub optOrange_Change(Value As CheckBoxConstants)
    If optOrange.Value = vbChecked Then
        Color = Orange
        lColor = LOrange
        Call ActiveFormClr(Color, lColor)
    End If
    SaveSetting App.CompanyName, "Theme", "frmColor", Orange
    SaveSetting App.CompanyName, "Theme", "frmlColor", LOrange
End Sub



Private Sub optViolet_Change(Value As CheckBoxConstants)
    If optViolet.Value = vbChecked Then
        Color = Violet
        lColor = LViolet
        Call ActiveFormClr(Color, lColor)
    End If
    SaveSetting App.CompanyName, "Theme", "frmColor", Violet
    SaveSetting App.CompanyName, "Theme", "frmlColor", LViolet
End Sub
Private Sub ActiveFormClr(clr As String, lclr As String)
        Me.fTheme.TitleColor1 = clr
        Me.fTheme.TitleColor2 = lclr
        Me.fTheme.BorderColor = clr
        Me.cmdClose.BorderColor = clr
        Me.cmdClose.ForeColor = clr
        frmToday.fToday.TitleColor1 = clr
        frmToday.fToday.TitleColor2 = lclr
        frmToday.fToday.BorderColor = clr
        frmToday.vkLabel1.ForeColor = clr
        frmToday.today_label1.ForeColor = clr
        frmToday.today_label2.ForeColor = clr
        frmToday.today_label3.ForeColor = clr
        frmToday.today_label4.ForeColor = clr
End Sub

