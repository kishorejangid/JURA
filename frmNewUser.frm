VERSION 5.00
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkusercontrolsxp.ocx"
Begin VB.Form frmNewUser 
   BorderStyle     =   0  'None
   Caption         =   "User"
   ClientHeight    =   3630
   ClientLeft      =   0
   ClientTop       =   -30
   ClientWidth     =   4665
   Icon            =   "frmNewUser.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   Begin vkUserContolsXP.vkFrame fNewUser 
      Height          =   3645
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   6429
      Caption         =   "New User"
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
         Left            =   4200
         TabIndex        =   11
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
      Begin vkUserContolsXP.vkCommand cmdDelete 
         Height          =   495
         Left            =   480
         TabIndex        =   10
         Top             =   2880
         Visible         =   0   'False
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   873
         Caption         =   "Delete User"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   33023
         BorderColor     =   33023
         CustomStyle     =   0
      End
      Begin vkUserContolsXP.vkCommand cmdChange 
         Height          =   495
         Left            =   480
         TabIndex        =   9
         Top             =   2880
         Visible         =   0   'False
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   873
         Caption         =   "Change Password"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   33023
         BorderColor     =   33023
         CustomStyle     =   0
      End
      Begin vkUserContolsXP.vkLabel lblLine 
         Height          =   30
         Left            =   0
         TabIndex        =   8
         Top             =   1080
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   53
         BorderStyle     =   1
         BorderColor     =   33023
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
      Begin vkUserContolsXP.vkLabel lblUserType 
         Height          =   255
         Left            =   360
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   720
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
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
      Begin VB.ComboBox cmbUserType 
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
         TabIndex        =   1
         Top             =   600
         Width           =   2895
      End
      Begin vkUserContolsXP.vkLabel lblUserID 
         Height          =   255
         Left            =   360
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1440
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
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
      End
      Begin vkUserContolsXP.vkTextBox txtUserID 
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   1320
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
      Begin vkUserContolsXP.vkLabel lblPassword 
         Height          =   255
         Left            =   360
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2040
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
      Begin vkUserContolsXP.vkTextBox txtPassword 
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   1920
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
      Begin vkUserContolsXP.vkCommand cmdCreate 
         Height          =   495
         Left            =   480
         TabIndex        =   4
         Top             =   2880
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   873
         Caption         =   "Create"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   33023
         BorderColor     =   33023
         CustomStyle     =   0
      End
   End
End
Attribute VB_Name = "frmNewUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdChange_Click()
    On Error Resume Next
    Dim rs As New ADODB.Recordset
    Dim qr As String
    Dim UserType As String
    Dim UserID As String
    Dim Password As String
    If txtUserID.Text = "" Or txtPassword.Text = "" Then
        JuraMsgBox ("Enter Some Data")
        Exit Sub
    End If
    UserType = cmbUserType.Text
    UserID = txtUserID.Text
    Password = txtPassword.Text
    rs.CursorLocation = adUseClient
    qr = "update login set loginpassword = '" & Password & "' where loginid='" & UserID & "'"
    rs.Open qr, conn, adOpenDynamic, adLockOptimistic, 1
    If Err.Number <> 0 Then
        MsgBox Error & vbCrLf & "Error Number: " & Err.Number
    Else
        JuraMsgBox ("Password Changed")
    End If
    rs.Close
    Unload Me
End Sub

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCreate_Click()
    On Error Resume Next
    Dim rs As New ADODB.Recordset
    Dim qr As String
    Dim UserType As String
    Dim UserID As String
    Dim Password As String
    If txtUserID.Text = "" Or txtPassword.Text = "" Then
        'MsgBox "Enter Some Data"
        JuraMsgBox ("Enter Some Data")
        Exit Sub
    End If
    UserType = cmbUserType.Text
    UserID = txtUserID.Text
    Password = txtPassword.Text
    rs.CursorLocation = adUseClient
    qr = "insert into login values('" & UserType & "','" & UserID & "','" & Password & "')"
    rs.Open qr, conn, adOpenDynamic, adLockOptimistic, 1
    If Err.Number <> 0 Then
        If Err.Number = -2147217873 Then
            MsgBox "A User With This UserID Alredy Exists" & vbCrLf & "Reenter With a Different ID"
        Else
            MsgBox Error & vbCrLf & "Error Number: " & Err.Number
        End If
    Else
        JuraMsgBox ("Created")
        'MsgBox "Created"
    End If
    rs.Close
    txtUserID.Text = ""
    txtPassword.Text = ""
End Sub

Private Sub cmdDelete_Click()
    On Error Resume Next
    Dim rs As New ADODB.Recordset
    Dim qr As String
    Dim UserType As String
    Dim UserID As String
    Dim Password As String
    If txtUserID.Text = "" Or txtPassword.Text = "" Then
        JuraMsgBox ("Enter Some Data")
        Exit Sub
    End If
    UserType = cmbUserType.Text
    UserID = txtUserID.Text
    Password = txtPassword.Text
    rs.CursorLocation = adUseClient
    qr = "delete from login where logintype = '" & UserType & "' and loginid = '" & UserID & "' "
    rs.Open qr, conn, adOpenDynamic, adLockOptimistic, 1
    If Err.Number <> 0 Then
        MsgBox Error & vbCrLf & "Error Number: " & Err.Number
    Else
        'MsgBox "User Deleted"
        JuraMsgBox ("User Deleted")
    End If
    rs.Close
    txtUserID.Text = ""
    txtPassword.Text = ""
End Sub

Private Sub fNewUser_MouseDown(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_Load()
    CreateRoundRectFromWindow Me, 7, 7
    Me.Top = (mdiMain.Height - Me.Height) / 5
    Me.Left = (mdiMain.Width - Me.Width) / 2
    Me.BackColor = mdiMain.BackColor
    cmbUserType.AddItem ("Administrator")
    cmbUserType.AddItem ("Staff")
    cmbUserType.AddItem ("Student")
    cmbUserType.Text = cmbUserType.List(2)
    Call frmColor(frmNewUser)
End Sub

