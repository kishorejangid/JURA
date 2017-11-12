VERSION 5.00
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkusercontrolsxp.ocx"
Begin VB.Form frmSettings 
   BorderStyle     =   0  'None
   Caption         =   "Application Settings"
   ClientHeight    =   6705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9810
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   9810
   ShowInTaskbar   =   0   'False
   Begin vkUserContolsXP.vkCheck cbRandomTheme 
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   1320
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   661
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   "Use Random Theme"
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
   Begin JURA.StylerButton cmdClose 
      Height          =   255
      Left            =   9240
      TabIndex        =   2
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
   Begin JURA.StylerButton cmdMin 
      Height          =   255
      Left            =   9000
      TabIndex        =   1
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Caption         =   "-"
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
   Begin vkUserContolsXP.vkFrame fSettings 
      Height          =   6705
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9810
      _ExtentX        =   17304
      _ExtentY        =   11827
      Caption         =   "Application Settings"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleColor1     =   33023
      TitleColor2     =   8438015
      TitleGradient   =   2
      TitleHeight     =   300
      BorderColor     =   33023
      RoundAngle      =   5
      BorderWidth     =   2
      Begin vkUserContolsXP.vkTextBox txtCollegeName 
         Height          =   375
         Left            =   1440
         TabIndex        =   10
         Top             =   600
         Width           =   8055
         _ExtentX        =   14208
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
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkLabel lblCollegeName 
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "College Name:"
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
      Begin vkUserContolsXP.vkFrame vkFrame1 
         Height          =   1935
         Left            =   240
         TabIndex        =   4
         Top             =   3360
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   3413
         Caption         =   "Examination Month && Year"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox cmbExamYear 
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
            Left            =   240
            TabIndex        =   8
            Top             =   1320
            Width           =   1935
         End
         Begin vkUserContolsXP.vkLabel lblYear 
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   1080
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Year:"
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
         Begin vkUserContolsXP.vkLabel lblMonth 
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Month:"
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
         Begin VB.ComboBox cmbExamMonth 
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
            Left            =   240
            TabIndex        =   5
            Top             =   600
            Width           =   1935
         End
      End
      Begin vkUserContolsXP.vkCommand cmdSave 
         Height          =   495
         Left            =   7320
         TabIndex        =   3
         Top             =   5880
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         Caption         =   "Save"
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
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SettingsTop As Integer

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub cmdMin_Click()
    Dim i As Long
    If State = 1 Then
        State = 0
        cmdMin.Caption = "+"
        For i = 6705 To 310 Step -150
            Me.Height = i
            DoEvents
        Next i
        Me.Height = 310
        fSettings.Height = 310
        fSettings.BorderWidth = 0
        Me.Top = 100
    Else
        State = 1
        cmdMin.Caption = "--"
        For i = 310 To 6705 Step 150
            Me.Height = i
            DoEvents
        Next i
        Me.Height = 6705
        fSettings.Height = 6705
        fSettings.BorderWidth = 2
        Me.Top = SettingsTop
    End If

End Sub

Private Sub cmdSave_Click()
    On Error Resume Next
       
    If cbRandomTheme.Value = vbChecked Then
        SaveSetting App.CompanyName, "Settings", "flgRandomTheme", True              'Odd or Even Semester
        flgRandomTheme = True
    Else
        SaveSetting App.CompanyName, "Settings", "flgRandomTheme", False              'Odd or Even Semester
        flgRandomTheme = False
    End If
    
    SaveSetting App.CompanyName, "Settings", "ExamMonth", Trim(cmbExamMonth.Text)
    strExamMonth = cmbExamMonth.Text
    
    SaveSetting App.CompanyName, "Settings", "ExamYear", Trim(cmbExamYear.Text)
    strExamYear = cmbExamYear.Text
    
    SaveSetting App.CompanyName, "Settings", "CollegeName", Trim(txtCollegeName.Text)
    strCollegeName = txtCollegeName.Text
    
    If Err.Number = 0 Then
        MsgBox "Settings Saved Succesfully"
    End If
End Sub

Private Sub Form_Load()
    CreateRoundRectFromWindow Me, 7, 7
    Me.Top = (mdiMain.Height - Me.Height) / 5
    Me.Left = (mdiMain.Width - Me.Width) / 2
    Me.BackColor = mdiMain.BackColor
    Call frmColor(frmSettings)
    Call cmbExamMonthYearData
    cmbExamMonth.Text = strExamMonth
    cmbExamYear.Text = strExamYear
    txtCollegeName.Text = strCollegeName
    If flgRandomTheme = True Then
        cbRandomTheme.Value = vbChecked
    Else
        cbRandomTheme.Value = vbUnchecked
    End If
End Sub

Private Sub cmbExamMonthYearData()
    cmbExamMonth.AddItem "JAN/FEB"
    cmbExamMonth.AddItem "FEB/MAR"
    cmbExamMonth.AddItem "MAR/APR"
    cmbExamMonth.AddItem "APR/MAY"
    cmbExamMonth.AddItem "MAY/JUN"
    cmbExamMonth.AddItem "JUN/JUL"
    cmbExamMonth.AddItem "JUL/AUG"
    cmbExamMonth.AddItem "AUG/SEP"
    cmbExamMonth.AddItem "SEP/OCT"
    cmbExamMonth.AddItem "OCT/NOV"
    cmbExamMonth.AddItem "NOV/DEC"
    cmbExamMonth.AddItem "DEC/JAN"
    cmbExamMonth.Text = cmbExamMonth.List(0)
    
    Dim i As Integer
    For i = 0 To 10
        cmbExamYear.AddItem "2006" + i
    Next
    cmbExamYear.Text = cmbExamYear.List(0)
End Sub


