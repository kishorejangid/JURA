VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkusercontrolsxp.ocx"
Begin VB.Form frmEmail 
   BorderStyle     =   0  'None
   Caption         =   "Email"
   ClientHeight    =   5880
   ClientLeft      =   105
   ClientTop       =   0
   ClientWidth     =   8640
   Icon            =   "frmEmail.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   8640
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lboxEmail 
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
      Height          =   1350
      IntegralHeight  =   0   'False
      Left            =   1200
      TabIndex        =   16
      Top             =   1475
      Visible         =   0   'False
      Width           =   5175
   End
   Begin vkUserContolsXP.vkFrame fEmail 
      Height          =   5895
      Left            =   0
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   10398
      Caption         =   "Email"
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
      Begin JURA.StylerButton cmdMin 
         Height          =   255
         Left            =   7920
         TabIndex        =   18
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
      Begin JURA.StylerButton cmdClose 
         Height          =   255
         Left            =   8160
         TabIndex        =   17
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
      Begin MSComDlg.CommonDialog Dialog1 
         Left            =   120
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin vkUserContolsXP.vkCommand btnAttach 
         Height          =   495
         Left            =   6600
         TabIndex        =   5
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         Caption         =   "Attachment"
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
      Begin vkUserContolsXP.vkCommand btnSend 
         Height          =   495
         Left            =   6600
         TabIndex        =   6
         Top             =   1200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         Caption         =   "Send"
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
      Begin vkUserContolsXP.vkCommand btnSettings 
         Height          =   495
         Left            =   6600
         TabIndex        =   7
         Top             =   1800
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         Caption         =   "Settings"
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
      Begin vkUserContolsXP.vkTextBox txtAttach 
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   2040
         Width           =   5175
         _ExtentX        =   9128
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkBar prgProgress 
         Height          =   495
         Left            =   240
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   5160
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   873
         BorderColor     =   33023
         LeftColor       =   33023
         RightColor      =   33023
         Value           =   1
         BackPicture     =   "frmEmail.frx":57E2
         FrontPicture    =   "frmEmail.frx":57FE
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
      Begin vkUserContolsXP.vkLabel lblProgress 
         Height          =   255
         Left            =   240
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   4800
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Progress:"
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
      Begin vkUserContolsXP.vkLabel lblAttach 
         Height          =   255
         Left            =   240
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2040
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Attachment:"
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
      Begin vkUserContolsXP.vkLabel lblSubject 
         Height          =   375
         Left            =   240
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1560
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Subject:"
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
      Begin vkUserContolsXP.vkLabel lblTo 
         Height          =   375
         Left            =   240
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1080
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "To:"
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
      Begin vkUserContolsXP.vkLabel lblFrom 
         Height          =   255
         Left            =   240
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   600
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "From:"
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
      Begin vkUserContolsXP.vkLabel lblMessage 
         Height          =   375
         Left            =   240
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2520
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Message:"
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
      Begin vkUserContolsXP.vkTextBox txtMessage 
         Height          =   2295
         Left            =   1200
         TabIndex        =   4
         Top             =   2520
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   4048
         BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   33023
         MultiLine       =   -1  'True
         ScrollBars      =   2
         LegendText      =   "MESSAGE"
         LegendAlignmentY=   2
         LegendBackColor1=   8438015
         LegendBackColor2=   33023
         LegendGradient  =   1
         LegendForeColor =   16777215
         LegendType      =   1
         LegendWidth     =   300
      End
      Begin vkUserContolsXP.vkTextBox txtSubject 
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   1560
         Width           =   5175
         _ExtentX        =   9128
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox txtTo 
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   1080
         Width           =   5175
         _ExtentX        =   9128
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox txtFrom 
         Height          =   375
         Left            =   1200
         TabIndex        =   0
         Top             =   600
         Width           =   5175
         _ExtentX        =   9128
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
   End
End
Attribute VB_Name = "frmEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Server As String
Dim Port As String
Dim UserName As String
Dim Password As String
Dim From As String
Dim Reciever As String
Dim Subject As String
Dim Message As String
Dim Attach As String
Dim email As CDO.Message
Dim EmailTop As Integer
Dim State As Integer

Private Sub btnAttach_Click()
    With Dialog1
       .InitDir = "" & App.Path & "\Reports"
       .Filter = "HTML|*.html|Text|*.txt|PDF|*.pdf|"
       .ShowOpen
          If .FileName <> "" Then
             Attach = .FileName
          End If
    End With
    txtAttach.Text = Attach
    txtAttach.Refresh
End Sub

Private Sub btnSend_Click()
    Reciever = txtTo.Text
    From = txtFrom.Text
    Subject = txtSubject.Text
    Message = txtMessage.Text
    SendMail (Message)
End Sub

Private Sub btnSettings_Click()
    frmEmailSettings.Show
End Sub

Private Sub CmdClose_Click()
    Unload Me
End Sub
Private Sub cmdMin_Click()
    Dim i As Long
    If State = 1 Then
        State = 0
        cmdMin.Caption = "+"
        For i = 5880 To 310 Step -150
            Me.Height = i
            DoEvents
        Next i
        Me.Height = 310
        fEmail.Height = 310
        fEmail.BorderWidth = 0
        Me.Top = 300
    Else
        State = 1
        cmdMin.Caption = "--"
        For i = 310 To 5880 Step 150
            Me.Height = i
            DoEvents
        Next i
        Me.Height = 5880
        fEmail.Height = 5880
        fEmail.BorderWidth = 2
        Me.Top = EmailTop
    End If
End Sub

Private Sub fEmail_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub fEmail_MouseDown(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub



Private Sub Form_Load()
    CreateRoundRectFromWindow Me, 7, 7
    Me.Top = (mdiMain.Height - Me.Height) / 5
    Me.Left = (mdiMain.Width - Me.Width) / 2
    Me.BackColor = mdiMain.BackColor
    Call frmColor(frmEmail)
    Server = GetSetting(App.CompanyName, "Email", "Server", "smtp.gmail.com")
    Port = GetSetting(App.CompanyName, "Email", "Port", "465")
    UserName = GetSetting(App.CompanyName, "Email", "UserName", "jangidsoft@gmail.com")
    Password = GetSetting(App.CompanyName, "Email", "Password", "9994580345")
    txtFrom.Text = UserName
    txtFrom.Refresh
    State = 1
    EmailTop = Me.Top
End Sub
Public Function SendMail(msgBody As String)
    On Error Resume Next
    Set email = New CDO.Message
    email.Configuration.Fields(cdoSMTPServer) = Server
    prgProgress.Value = 5
    email.Configuration.Fields(cdoSMTPServerPort) = Port
    prgProgress.Value = 10
    email.Configuration.Fields(cdoSMTPUseSSL) = True
    prgProgress.Value = 15
    email.Configuration.Fields(cdoSMTPAuthenticate) = 1
    prgProgress.Value = 20
    email.Configuration.Fields(cdoSendUserName) = UserName
    prgProgress.Value = 25
    email.Configuration.Fields(cdoSendPassword) = Password
    prgProgress.Value = 30
    email.Configuration.Fields(cdoSMTPConnectionTimeout) = 30
    prgProgress.Value = 35
    email.Configuration.Fields(cdoSendUsingMethod) = 2
    prgProgress.Value = 40
    email.Configuration.Fields.Update
    prgProgress.Value = 45
    email.To = Reciever
    prgProgress.Value = 50
    email.From = From
    prgProgress.Value = 60
    email.Subject = Subject
    prgProgress.Value = 70
    email.TextBody = msgBody & vbCrLf & "Sent From Jura(Jangid's University Result Analysis"
    prgProgress.Value = 80
    email.AddAttachment (Attach)
    prgProgress.Value = 90
    email.Send
    prgProgress.Value = 100
    Set email = Nothing
    If Err.Number <> 0 Then
        MsgBox Error & vbCrLf & "Error Number: " & Err.Number
    Else
        MsgBox "Mail Sent Sucessfully"
    End If
End Function
Private Sub Form_Unload(Cancel As Integer)
    Set email = Nothing
End Sub



Private Sub lboxEmail_Click()
    txtTo.Text = lboxEmail.Text
    lboxEmail.Visible = False
End Sub
Private Sub txtTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        lboxEmail.SetFocus
    End If
End Sub

Private Sub txtTo_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    lboxEmail.Visible = True
    lboxEmail.Clear
    Dim rs As New ADODB.Recordset
    Dim sql As String
    Dim E As String
    E = Trim(txtTo.Text & "%")
    sql = "select email from studdetails where email like '" & E & "'"
    rs.Open sql, conn, adOpenDynamic, adLockOptimistic
    Do Until rs.EOF
            lboxEmail.AddItem (rs.Fields(0))
            rs.MoveNext
    Loop
End Sub


