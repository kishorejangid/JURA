VERSION 5.00
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkusercontrolsxp.ocx"
Begin VB.Form frmNewSubject 
   BorderStyle     =   0  'None
   Caption         =   "New Subject"
   ClientHeight    =   5055
   ClientLeft      =   0
   ClientTop       =   -30
   ClientWidth     =   6615
   Icon            =   "frmNewSubject.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5055
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Begin vkUserContolsXP.vkFrame fNewSubject 
      Height          =   5055
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   8916
      BackColor1      =   16777215
      Caption         =   "New Subject"
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
      Begin vkUserContolsXP.vkLabel lblCredit 
         Height          =   375
         Left            =   480
         TabIndex        =   18
         Top             =   2160
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Credit:"
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
      Begin vkUserContolsXP.vkFrame vkFrame4 
         Height          =   375
         Left            =   1680
         TabIndex        =   17
         Top             =   1920
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   661
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
         TitleGradient   =   2
         BorderColor     =   33023
         BreakCorner     =   0   'False
         BorderWidth     =   2
         DisplayPicture  =   0   'False
      End
      Begin VB.ComboBox cmbCredit 
         Appearance      =   0  'Flat
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
         Left            =   1680
         TabIndex        =   2
         Top             =   1920
         Width           =   4455
      End
      Begin vkUserContolsXP.vkFrame vkFrame1 
         Height          =   375
         Left            =   1680
         TabIndex        =   16
         Top             =   3720
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   661
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
         BorderWidth     =   2
      End
      Begin VB.ComboBox cmbBatch 
         Appearance      =   0  'Flat
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
         Left            =   1680
         TabIndex        =   5
         Top             =   3720
         Width           =   4455
      End
      Begin vkUserContolsXP.vkLabel vkLabel5 
         Height          =   375
         Left            =   480
         TabIndex        =   15
         Top             =   3840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Batch:"
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
         Left            =   6120
         TabIndex        =   14
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
      Begin vkUserContolsXP.vkCommand cmdInsert 
         Height          =   510
         Left            =   480
         TabIndex        =   6
         Top             =   4320
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   900
         Caption         =   "Insert"
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
      Begin vkUserContolsXP.vkTextBox txtSubjName 
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   1320
         Width           =   4455
         _ExtentX        =   7858
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
      Begin vkUserContolsXP.vkLabel vkLabel3 
         Height          =   375
         Left            =   480
         TabIndex        =   13
         Top             =   2760
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Semester:"
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
      Begin vkUserContolsXP.vkLabel vkLabel4 
         Height          =   375
         Left            =   480
         TabIndex        =   12
         Top             =   3240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
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
      Begin vkUserContolsXP.vkFrame vkFrame2 
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Top             =   2520
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   661
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
         TitleGradient   =   2
         BorderColor     =   33023
         BreakCorner     =   0   'False
         BorderWidth     =   2
         DisplayPicture  =   0   'False
      End
      Begin vkUserContolsXP.vkFrame vkFrame3 
         Height          =   375
         Left            =   1680
         TabIndex        =   10
         Top             =   3120
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   661
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
         BorderWidth     =   2
      End
      Begin VB.ComboBox cmbSem 
         Appearance      =   0  'Flat
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
         Left            =   1680
         TabIndex        =   3
         Top             =   2520
         Width           =   4455
      End
      Begin VB.ComboBox cmbDept 
         Appearance      =   0  'Flat
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
         Left            =   1680
         TabIndex        =   4
         Top             =   3120
         Width           =   4455
      End
      Begin vkUserContolsXP.vkLabel vkLabel2 
         Height          =   375
         Left            =   480
         TabIndex        =   9
         Top             =   1440
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Subject Name:"
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
      Begin vkUserContolsXP.vkTextBox txtSubjCode 
         Height          =   375
         Left            =   1680
         TabIndex        =   0
         Top             =   720
         Width           =   4455
         _ExtentX        =   7858
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
      Begin vkUserContolsXP.vkLabel vkLabel1 
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Subject Code:"
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
Attribute VB_Name = "frmNewSubject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmbDept_Click()
    iDept = Department(cmbDept)
End Sub

Private Sub cmbSem_Click()
    iSem = cmbSem.Text
End Sub

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub cmdInsert_Click()
    On Error Resume Next
    Dim rs As New ADODB.Recordset
    Dim qr As String
    Dim strSubjCode As String
    Dim strSubjName As String
    If txtSubjcode.Text = "" Or txtSubjName.Text = "" Then
        MsgBox "Enter Some Data"
        Exit Sub
    End If
    Dim iCredit As String
    If cmbCredit.Text = "" Then
        iCredit = vbNullString
    Else
        iCredit = cmbCredit.Text
    End If
    strSubjCode = Trim(txtSubjcode.Text)
    strSubjName = Trim(txtSubjName.Text)
    rs.CursorLocation = adUseClient
    qr = "insert into subj (subjcode,subjname,semno,dept,batch,credit) values('" & strSubjCode & "','" & strSubjName & "'," & iSem & "," & iDept & "," & Mid(cmbBatch.Text, 3, 2) & ",'" & iCredit & "')"
    rs.Open qr, conn, adOpenDynamic, adLockOptimistic, 1
    If Err.Number <> 0 Then
        MsgBox Error & vbCrLf & "Error Number: " & Err.Number
    ElseIf Err.Number = -2147217873 Then
        MsgBox "Subject Code all ready exist in the database for the same dept"
    Else
        MsgBox "Inserted"
    End If
    txtSubjcode.Text = ""
    txtSubjName.Text = ""
    cmbCredit.Text = ""
    txtSubjcode.SetFocus
End Sub
Private Sub fNewSubject_MouseDown(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
Private Sub Form_GotFocus()
    txtSubjcode.SetFocus
End Sub

Private Sub Form_Load()
    CreateRoundRectFromWindow Me, 7, 7
    Me.Top = (mdiMain.Height - Me.Height) / 5
    Me.Left = (mdiMain.Width - Me.Width) / 2
    Me.BackColor = mdiMain.BackColor
    Call frmColor(frmNewSubject)
    Call cmbDept_Load(cmbDept)
    Call cmbSem_Load(cmbSem)
    Call cmbBatch_Load(cmbBatch)
    cmbCredit.AddItem ("1")
    cmbCredit.AddItem ("2")
    cmbCredit.AddItem ("3")
    cmbCredit.AddItem ("4")
    cmbCredit.AddItem ("5")
End Sub
Private Sub txtSubjCode_LostFocus()
    txtSubjcode.Text = UCase(txtSubjcode.Text)
End Sub
Private Sub txtSubjName_LostFocus()
    txtSubjName.Text = JangidFormat(txtSubjName.Text)
End Sub
