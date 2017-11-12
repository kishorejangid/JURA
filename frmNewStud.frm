VERSION 5.00
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkusercontrolsxp.ocx"
Begin VB.Form frmNewStud 
   BorderStyle     =   0  'None
   Caption         =   "New Student"
   ClientHeight    =   3285
   ClientLeft      =   0
   ClientTop       =   -30
   ClientWidth     =   4665
   Icon            =   "frmNewStud.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3285
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbSec 
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
      Top             =   1800
      Width           =   2895
   End
   Begin vkUserContolsXP.vkLabel lblSec 
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1800
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   "Section:"
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
   Begin vkUserContolsXP.vkFrame fNewStud 
      Height          =   3285
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   5794
      Caption         =   "New Student"
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
      TitleColor2     =   12640511
      TitleGradient   =   2
      TitleHeight     =   360
      BorderColor     =   33023
      BorderWidth     =   2
      Begin JURA.StylerButton cmdClose 
         Height          =   255
         Left            =   4080
         TabIndex        =   7
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
         Height          =   495
         Left            =   480
         TabIndex        =   3
         Top             =   2520
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   873
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
      Begin vkUserContolsXP.vkTextBox txtName 
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
      Begin vkUserContolsXP.vkLabel vkLabel2 
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   1320
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Name:"
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
      Begin vkUserContolsXP.vkTextBox txtRegNo 
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   33023
         MaxLength       =   11
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
         Caption         =   "Register No:"
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
Attribute VB_Name = "frmNewStud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub cmdInsert_Click()
    On Error Resume Next
    Dim rs As New ADODB.Recordset
    Dim qr As String
    Dim regno As String
    Dim name As String
    If txtRegNo.Text = "" Or txtName.Text = "" Then
        JuraMsgBox "Enter Some Data"
        Exit Sub
    End If
    regno = txtRegNo.Text
    name = txtName.Text
    rs.CursorLocation = adUseClient
    qr = "insert into studdetails (regno,studname,sec) values('" & regno & "','" & name & "','" & cmbSec.Text & "')"
    rs.Open qr, conn, adOpenDynamic, adLockOptimistic
    If Err.Number <> 0 Then
        MsgBox Error & vbCrLf & "Error Number: " & Err.Number
    Else
        MsgBox "Inserted"
    End If
    txtRegNo.Text = ""
    txtName.Text = ""
End Sub

Private Sub fNewStud_MouseDown(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_Load()
    CreateRoundRectFromWindow Me, 7, 7
    Me.Top = (mdiMain.Height - Me.Height) / 5
    Me.Left = (mdiMain.Width - Me.Width) / 2
    Me.BackColor = mdiMain.BackColor
    Call frmColor(frmNewStud)
    Call cmbSec_Load(cmbSec)
End Sub
Private Sub txtName_LostFocus()
    txtName.Text = JangidFormat(txtName.Text)
End Sub

Private Sub vkFrame1_KeyDown(KeyCode As Integer, Shift As Integer)

End Sub
