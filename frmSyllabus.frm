VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkusercontrolsxp.ocx"
Begin VB.Form frmSyllabus 
   BorderStyle     =   0  'None
   Caption         =   "Syllabus"
   ClientHeight    =   8415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13815
   Icon            =   "frmSyllabus.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8415
   ScaleMode       =   0  'User
   ScaleWidth      =   13815
   ShowInTaskbar   =   0   'False
   Begin vkUserContolsXP.vkFrame fSyllabus 
      Height          =   8415
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   14843
      Caption         =   "Syllabus"
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
         Left            =   13080
         TabIndex        =   5
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
      Begin vkUserContolsXP.vkFrame fBorder 
         Height          =   7215
         Left            =   360
         TabIndex        =   4
         Top             =   960
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   12726
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
         ShowTitle       =   0   'False
         BorderColor     =   33023
         BreakCorner     =   0   'False
         BorderWidth     =   3
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser 
         Height          =   7215
         Left            =   360
         TabIndex        =   3
         Top             =   960
         Width           =   13095
         ExtentX         =   23098
         ExtentY         =   12726
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin VB.ComboBox syllabus_cmbDept 
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
         Left            =   1440
         TabIndex        =   2
         Top             =   480
         Width           =   12015
      End
      Begin vkUserContolsXP.vkLabel lblDept 
         Height          =   255
         Left            =   360
         TabIndex        =   0
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Department:"
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
Attribute VB_Name = "frmSyllabus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strFileName As String

Private Sub CmdClose_Click()
    Unload Me
End Sub



Private Sub Form_Load()
    Me.Width = Screen.Width * 0.9
    fSyllabus.Width = Me.Width
    cmdClose.Left = Me.Width - cmdClose.Width - 240
    syllabus_cmbDept.Width = Me.Width - syllabus_cmbDept.Left - 360
    fBorder.Width = Me.Width - fBorder.Left - 360
    WebBrowser.Width = Me.Width - WebBrowser.Left - 400
    Me.Top = 250
    Me.Left = (mdiMain.Width - Me.Width) / 2
    CreateRoundRectFromWindow Me, 7, 7
    Me.BackColor = mdiMain.BackColor
    Call frmColor(frmSyllabus)
    syllabus_cmbDept.AddItem "Computer Science && Engineering"
    syllabus_cmbDept.AddItem "Information Technology"
    syllabus_cmbDept.AddItem "Electronics && Communication Engineering"
    syllabus_cmbDept.AddItem "Electrical && Electronics Engineering"
    syllabus_cmbDept.AddItem "Mechanical Engineering"
    syllabus_cmbDept.Text = syllabus_cmbDept.List(0)
    Call syllabus_cmbDept_Click
End Sub


Private Sub fSyllabus_MouseDown(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub syllabus_cmbDept_Click()
    If syllabus_cmbDept.Text = "Computer Science && Engineering" Then
        strFileName = "" & App.Path & "\Syllabus\CSE.pdf"
        WebBrowser.Navigate strFileName
    ElseIf syllabus_cmbDept.Text = "Information Technology" Then
        strFileName = "" & App.Path & "\Syllabus\IT.pdf"
        WebBrowser.Navigate strFileName
    ElseIf syllabus_cmbDept.Text = "Electronics && Communication Engineering" Then
        strFileName = "" & App.Path & "\Syllabus\ECE.pdf"
        WebBrowser.Navigate strFileName
    ElseIf syllabus_cmbDept.Text = "Electrical && Electronics Engineering" Then
        strFileName = "" & App.Path & "\Syllabus\EEE.pdf"
        WebBrowser.Navigate strFileName
    ElseIf syllabus_cmbDept.Text = "Mechanical Engineering" Then
        strFileName = "" & App.Path & "\Syllabus\CSE.doc"
        WebBrowser.Navigate strFileName
    End If
End Sub
