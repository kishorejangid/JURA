VERSION 5.00
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkusercontrolsxp.ocx"
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H000080FF&
   BorderStyle     =   0  'None
   ClientHeight    =   4230
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin vkUserContolsXP.vkTimer vkTimer3 
      Left            =   1440
      Top             =   360
      _ExtentX        =   926
      _ExtentY        =   926
      Interval        =   40
   End
   Begin vkUserContolsXP.vkTimer vkTimer2 
      Left            =   840
      Top             =   360
      _ExtentX        =   926
      _ExtentY        =   926
      Interval        =   40
   End
   Begin vkUserContolsXP.vkTimer vkTimer1 
      Left            =   240
      Top             =   360
      _ExtentX        =   926
      _ExtentY        =   926
      Interval        =   40
      Enabled         =   -1  'True
   End
   Begin VB.Frame frame 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4125
      Left            =   50
      TabIndex        =   0
      Top             =   50
      Width           =   7275
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   120
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   1035
         Width           =   2655
      End
      Begin VB.Label lblCopyright 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   4
         Top             =   3420
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         BackStyle       =   0  'Transparent
         Caption         =   "Mahadev Stationeries"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   3
         Top             =   3660
         Width           =   2415
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "Warning: This Product is copyright Protected"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   255
         TabIndex        =   2
         Top             =   3735
         Width           =   6855
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6405
         TabIndex        =   5
         Top             =   2700
         Width           =   330
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Windows"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5115
         TabIndex        =   6
         Top             =   2340
         Width           =   1500
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2880
         TabIndex        =   8
         Top             =   1200
         Width           =   2430
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "LicenseTo: Kishore Jangid"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jangid Corporation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3675
         TabIndex        =   7
         Top             =   585
         Width           =   3345
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim a As Integer
Dim l_val As Integer
Private m_Trans As transperant
Private Sub Form_Load()
    CreateRoundRectFromWindow Me, 7, 7
    frame.Width = Me.Width - 100
    frame.Height = Me.Height - 100
    l_val = 0
    Set m_Trans = New transperant
    m_Trans.hWnd = Me.hWnd
    m_Trans.Alpha = l_val
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = "JURA™"
    lblCompanyProduct.Caption = "Jangid Corporation"
    frmSplash.BackColor = Color
End Sub
Private Sub Form_Unload(Cancel As Integer)
    vkTimer1.Enabled = False
    vkTimer2.Enabled = False
    vkTimer3.Enabled = False
End Sub
Private Sub vkTimer1_Timer()
    If l_val >= 255 Then
        vkTimer1.Enabled = False
        vkTimer2.Enabled = True
        Exit Sub
    Else
        m_Trans.Alpha = l_val
    End If
    l_val = l_val + 5
End Sub

Private Sub vkTimer2_Timer()
    Static i As Integer
    If i >= 2 Then
     frmSplash.Show Modal, mdiMain
     mdiMain.Show
     frmLogin.Show Modal, mdiMain
     vkTimer2.Interval = 0
     vkTimer3.Enabled = True
    End If
    i = i + 1
End Sub

Private Sub vkTimer3_Timer()
    If l_val <= 0 Then
        vkTimer3.Enabled = False
        Exit Sub
    Else
        m_Trans.Alpha = l_val
    End If
    l_val = l_val - 5
End Sub

