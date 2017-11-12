VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkusercontrolsxp.ocx"
Begin VB.Form frmDownloadResults 
   BorderStyle     =   0  'None
   ClientHeight    =   8790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13230
   MDIChild        =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   13230
   ShowInTaskbar   =   0   'False
   Begin vkUserContolsXP.vkCommand cmdNext 
      Height          =   375
      Left            =   6840
      TabIndex        =   17
      Top             =   8160
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      Caption         =   "Skip"
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
   Begin vkUserContolsXP.vkCommand cmdInsertMarks 
      Height          =   375
      Left            =   9960
      TabIndex        =   16
      Top             =   7680
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      Caption         =   "Insert Marks Of"
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
   Begin VB.ComboBox cmbURL 
      Height          =   315
      Left            =   600
      TabIndex        =   13
      Top             =   480
      Width           =   12495
   End
   Begin vkUserContolsXP.vkLabel vkLabel7 
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   "URL:"
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
   Begin JURA.StylerButton btnClose 
      Height          =   255
      Left            =   12720
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
      RoundedValue    =   3
   End
   Begin vkUserContolsXP.vkFrame vkFrame1 
      Height          =   2895
      Left            =   6720
      TabIndex        =   2
      Top             =   5760
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5106
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
      Begin vkUserContolsXP.vkLabel lblSec 
         Height          =   270
         Left            =   4440
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   600
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   476
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
      Begin VB.ComboBox cmbMarksSec 
         Height          =   315
         Left            =   4440
         TabIndex        =   34
         Top             =   840
         Width           =   975
      End
      Begin vkUserContolsXP.vkCommand cmdStudName 
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   1920
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         Caption         =   "Insert Names"
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
      Begin vkUserContolsXP.vkCommand cmdAuto 
         Height          =   375
         Left            =   3240
         TabIndex        =   23
         Top             =   2400
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         Caption         =   "Auto Insert Marks"
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
      Begin VB.ComboBox cmbBatch 
         Height          =   315
         Left            =   2040
         TabIndex        =   22
         Top             =   840
         Width           =   1095
      End
      Begin vkUserContolsXP.vkCommand cmdGo 
         Height          =   375
         Left            =   5640
         TabIndex        =   20
         Top             =   840
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         Caption         =   "Go"
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
      Begin VB.ComboBox cmbSem 
         Height          =   315
         Left            =   3240
         TabIndex        =   19
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox cmbDept 
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   1815
      End
      Begin VB.ComboBox cmbEndRegNo 
         Height          =   315
         Left            =   3240
         TabIndex        =   15
         Top             =   1440
         Width           =   3015
      End
      Begin VB.ComboBox cmbStartRegNo 
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   3015
      End
      Begin vkUserContolsXP.vkLabel lblRegNo 
         Height          =   270
         Left            =   960
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   476
         BorderStyle     =   1
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
      Begin vkUserContolsXP.vkLabel vkLabel6 
         Height          =   270
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   600
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   476
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Dept:"
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
      Begin vkUserContolsXP.vkLabel vkLabel5 
         Height          =   270
         Left            =   3240
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   600
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   476
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
         Height          =   270
         Left            =   3240
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   476
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Ending Reg No:"
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
      Begin vkUserContolsXP.vkLabel vkLabel3 
         Height          =   270
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   476
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Starting Reg No:"
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
      Begin vkUserContolsXP.vkLabel vkLabel2 
         Height          =   270
         Left            =   2640
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   476
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
      Begin vkUserContolsXP.vkLabel vkLabel1 
         Height          =   270
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   476
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Reg No:"
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
      Begin vkUserContolsXP.vkLabel lblName 
         Height          =   270
         Left            =   3360
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   476
         BorderStyle     =   1
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Batch:"
         Height          =   255
         Left            =   2160
         TabIndex        =   21
         Top             =   600
         Width           =   495
      End
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4695
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   960
      Width           =   12975
      ExtentX         =   22886
      ExtentY         =   8281
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
   Begin vkUserContolsXP.vkFrame fDownloadResult 
      Height          =   8790
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13230
      _ExtentX        =   23336
      _ExtentY        =   15505
      Caption         =   "Download Results"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleGradient   =   2
      TitleHeight     =   360
      BorderWidth     =   2
      Begin vkUserContolsXP.vkFrame vkFrame2 
         Height          =   2895
         Left            =   120
         TabIndex        =   25
         Top             =   5760
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   5106
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
         Begin vkUserContolsXP.vkFrame fStudName 
            Height          =   2895
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Visible         =   0   'False
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   5106
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
               Left            =   4080
               TabIndex        =   33
               Top             =   600
               Width           =   1215
            End
            Begin vkUserContolsXP.vkCommand cmdClose 
               Height          =   495
               Left            =   240
               TabIndex        =   32
               Top             =   2160
               Width           =   6015
               _ExtentX        =   10610
               _ExtentY        =   873
               Caption         =   "Close Menu"
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
            Begin vkUserContolsXP.vkCommand cmdInsert 
               Height          =   495
               Left            =   240
               TabIndex        =   31
               Top             =   1080
               Width           =   6015
               _ExtentX        =   10610
               _ExtentY        =   873
               Caption         =   "Insert Name Of"
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
            Begin vkUserContolsXP.vkCommand cmdStudGo 
               Height          =   375
               Left            =   5520
               TabIndex        =   30
               Top             =   600
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
               Caption         =   "Go"
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
            Begin vkUserContolsXP.vkTextBox txtStart 
               Height          =   375
               Left            =   240
               TabIndex        =   29
               Top             =   600
               Width           =   3615
               _ExtentX        =   6376
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
               LegendForeColor =   16750899
            End
            Begin vkUserContolsXP.vkLabel lblStart 
               Height          =   270
               Left            =   240
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   360
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   476
               BackColor       =   16777215
               BackStyle       =   0
               Caption         =   "Starting Reg No:"
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   2535
            Left            =   120
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   120
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   4471
            _Version        =   393216
            Rows            =   1
            Cols            =   6
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   -2147483634
            ScrollBars      =   2
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   6
         End
      End
   End
End
Attribute VB_Name = "frmDownloadResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Dim reg As Double
Dim icount As Integer
Dim Index As Integer

Dim document As MSHTML.HTMLDocument
Dim btn As MSHTML.HTMLButtonElement
Dim element As MSHTML.HTMLBaseElement
Dim frm As MSHTML.HTMLFormElement
Dim Tbl As MSHTML.HTMLTable
Dim tr As MSHTML.HTMLTableRow
Dim tc As MSHTML.HTMLTableCell
Dim inp As MSHTML.HTMLInputElement
Dim spa As MSHTML.HTMLSpanElement
Dim fDC As Boolean

Private Sub btnClose_Click()
    Unload Me
End Sub
Private Sub cmbBatch_Change()
    Call cmbBatch_Click
End Sub
Private Sub cmbBatch_Click()
    iBatch = cmbBatch.Text
    Call cmbRegNo_Load(cmbStartRegNo)
    Call cmbRegNo_Load(cmbEndRegNo)
    cmbStartRegNo.Text = cmbStartRegNo.List(Index)
    cmbEndRegNo.Text = cmbEndRegNo.List(cmbEndRegNo.ListCount - 1)
End Sub
Private Sub cmbDept_Change()
    Call cmbDept_Click
End Sub
Private Sub cmbDept_Click()
    iDept = Department(cmbDept)
    Call cmbRegNo_Load(cmbStartRegNo)
    Call cmbRegNo_Load(cmbEndRegNo)
    cmbStartRegNo.Text = cmbStartRegNo.List(Index)
    cmbEndRegNo.Text = cmbEndRegNo.List(cmbEndRegNo.ListCount - 1)
End Sub

Private Sub cmbMarksSec_Change()
    strSec = cmbMarksSec.Text
    Call cmbRegNo_Load(cmbStartRegNo)
    Call cmbRegNo_Load(cmbEndRegNo)
    cmbStartRegNo.Text = cmbStartRegNo.List(Index)
    cmbEndRegNo.Text = cmbEndRegNo.List(cmbEndRegNo.ListCount - 1)
End Sub

Private Sub cmbMarksSec_Click()
    strSec = cmbMarksSec.Text
    Call cmbRegNo_Load(cmbStartRegNo)
    Call cmbRegNo_Load(cmbEndRegNo)
    cmbStartRegNo.Text = cmbStartRegNo.List(Index)
    cmbEndRegNo.Text = cmbEndRegNo.List(cmbEndRegNo.ListCount - 1)
End Sub

Private Sub cmbSem_Change()
    Call cmbSem_Click
End Sub
Private Sub cmbSem_Click()
    iSem = cmbSem.Text
End Sub
Private Sub cmbStartRegNo_Click()
    Index = cmbStartRegNo.ListIndex
End Sub
Private Sub cmbURL_Change()
'    Reload
End Sub
Private Sub cmbURL_Click()
 '   Reload
End Sub
Private Sub CmdClose_Click()
    fStudName.Visible = False
End Sub
Private Sub cmdGo_Click()
    On Error Resume Next
    reg = cmbStartRegNo.Text
    Reload
    cmdInsertMarks.Caption = "Insert Marks Of " & reg
    cmdNext.Caption = "Skip " & reg
    cmdNext.Visible = True
    cmdInsertMarks.Visible = True
End Sub

Private Sub cmdInsert_Click()
    On Error Resume Next
    Dim rs As New ADODB.Recordset
    Dim sql As String
        
    rs.CursorLocation = adUseClient
    
    sql = "insert into studdetails (regno,studname,sec) values('" & txtStart.Text & "','" & lblName.Caption & "','" & cmbSec.Text & "')"
    rs.Open sql, conn, adOpenDynamic, adLockOptimistic, 1
    
    If Err.Number <> 0 Then
        MsgBox Error & vbCrLf & "Error Number: " & Err.Number
    End If
    reg = reg + 1
    txtStart.Text = reg
    cmdInsert.Caption = "Insert Name Of " & reg
    Reload
End Sub

Private Sub cmdInsertName_Click()
    On Error Resume Next
    Dim rs As New ADODB.Recordset
    Dim qr As String
    Dim regno As String
    Dim name As String
    
    'If reg > cmbEndRegNo.Text Then
        'MsgBox "Last Reg No Reached"
        'Exit Sub
    'End If
    
    If lblRegNo.Caption = "" Then
        Exit Sub
    End If
    
    regno = lblRegNo.Caption
    name = lblName.Caption
    
    rs.CursorLocation = adUseClient
    
    qr = "insert into studdetails (regno,studname) values('" & regno & "','" & name & "')"
    rs.Open qr, conn, adOpenDynamic, adLockOptimistic, 1
    
    If Err.Number <> 0 Then
        MsgBox Error & vbCrLf & "Error Number: " & Err.Number
    Else
        'MsgBox "Inserted"
    End If
    
    
    reg = reg + 1
    cmdInsertMarks.Caption = "Insert Marks Of " & reg
    cmdInsertName.Caption = "Insert Name Of " & reg
    cmdNext.Caption = "Skip " & reg
    Reload
End Sub
Private Sub cmdNext_Click()
    Index = Index + 1
    reg = cmbStartRegNo.List(Index)
    cmdNext.Caption = "Skip " & reg
    cmdInsertMarks.Caption = "Insert Marks Of " & reg
    Reload
End Sub
Private Sub cmdInsertMarks_Click()
    On Error Resume Next
    Dim rs As New ADODB.Recordset
    Dim qr As String
    Dim iRows As Integer
    Dim strGrade As String
    Dim iGrade As Integer
    iRows = 0
    rs.CursorLocation = adUseClient
    
    If reg > cmbEndRegNo.Text Then
        MsgBox "Last Reg No Reached"
        Exit Sub
    End If
    
    For i = 0 To MSHFlexGrid1.rows - 1
        If MSHFlexGrid1.TextMatrix(i, 2) = "AB" Then
            MSHFlexGrid1.TextMatrix(i, 2) = "NULL"
        End If
        If MSHFlexGrid1.TextMatrix(i, 1) = "--" Then
            MSHFlexGrid1.TextMatrix(i, 1) = "0"
        End If
        If cmbBatch.Text > 2007 Then
            strGrade = MSHFlexGrid1.TextMatrix(i, 3)
            Select Case strGrade
                 Case "S"
                     iGrade = 10
                 Case "A"
                     iGrade = 9
                 Case "B"
                     iGrade = 8
                 Case "C"
                     iGrade = 7
                 Case "D"
                     iGrade = 6
                 Case "E"
                     iGrade = 5
                 Case "I"
                     iGrade = 0
                 Case "W"
                     iGrade = 0
                 Case "U"
                     iGrade = 0
                 Case "AB"
                    iGrade = 0
            End Select
            qr = "insert into studmarks (regno,semno,dept,batch,subjcode,internals,grade,value,result) values('" & lblRegNo.Caption & "'," & GetSubjSem(MSHFlexGrid1.TextMatrix(i, 0), iDept, Mid(iBatch, 3, 2)) & ", " & iDept & "," & Mid(lblRegNo.Caption, 4, 2) & ",'" & MSHFlexGrid1.TextMatrix(i, 0) & "'," & CInt(Trim(MSHFlexGrid1.TextMatrix(i, 2))) & ",'" & Trim(MSHFlexGrid1.TextMatrix(i, 3)) & "'," & iGrade & ",'" & Trim(MSHFlexGrid1.TextMatrix(i, 4)) & "')"
        Else
            qr = "insert into studmarks (regno,semno,dept,batch,subjcode,internals,externals,result) values('" & lblRegNo.Caption & "'," & GetSubjSem(MSHFlexGrid1.TextMatrix(i, 0), iDept, Mid(iBatch, 3, 2)) & ", " & iDept & "," & Mid(lblRegNo.Caption, 4, 2) & ",'" & MSHFlexGrid1.TextMatrix(i, 0) & "'," & MSHFlexGrid1.TextMatrix(i, 2) & "," & MSHFlexGrid1.TextMatrix(i, 3) & ",'" & Trim(MSHFlexGrid1.TextMatrix(i, 5)) & "')"
            'MsgBox qr
            
        End If
        rs.Open qr, conn, adOpenDynamic, adLockOptimistic, 1
        iRows = iRows + 1
    Next i
    
    If Err.Number <> 0 Then
        MsgBox Error & vbCrLf & "Error Number: " & Err.Number
    Else
       'MsgBox "Inserted"
    End If
    Index = Index + 1
    reg = cmbStartRegNo.List(Index)
    cmdInsertMarks.Caption = "Insert Marks Of " & reg
    cmdNext.Caption = "Skip " & reg
    MSHFlexGrid1.Clear
    cmdInsertMarks.Enabled = False
    Reload
End Sub

Private Sub cmdStudGo_Click()
    On Error Resume Next
    reg = txtStart.Text
    cmdInsert.Caption = "Insert Name Of " & reg
    Reload
End Sub



Private Sub cmdStudName_Click()
    fStudName.Visible = True
End Sub

Private Sub fDownloadResult_MouseDown(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_Load()
    CreateRoundRectFromWindow Me, 7, 7
    Me.Top = 100
    Me.Left = (mdiMain.Width - Me.Width) / 2
    Me.BackColor = mdiMain.BackColor
    Index = 0
    Call frmColor(frmDownloadResults)
    Call cmbDept_Load(cmbDept)
    Call cmbSem_Load(cmbSem)
    Call cmbBatch_Load(cmbBatch)
    Call cmbSec_Load(cmbSec)
    Call cmbSec_Load(cmbMarksSec)
    strSec = cmbMarksSec.Text
    Call cmbRegNo_Load(cmbStartRegNo)
    Call cmbRegNo_Load(cmbEndRegNo)
    cmbStartRegNo.Text = cmbStartRegNo.List(Index)
    cmbEndRegNo.Text = cmbEndRegNo.List(cmbEndRegNo.ListCount - 1)
    cmbURL.AddItem ("http://218.248.20.139/ug2008/UG20082011.aspx") 'http://218.248.20.131/UG20072011/UG20072011.aspx")
    
    cmbURL.Text = cmbURL.List(0)
    Reload
    cmdNext.Visible = True
    cmdInsertMarks.Visible = True
End Sub



Private Sub Reload()
    WebBrowser1.Navigate cmbURL.Text
    
    MSHFlexGrid1.ColWidth(0) = 900
    MSHFlexGrid1.ColWidth(1) = 3500
    MSHFlexGrid1.ColWidth(2) = 500
    MSHFlexGrid1.ColWidth(3) = 500
    MSHFlexGrid1.ColWidth(4) = 500
    MSHFlexGrid1.ColWidth(5) = 500
    
    
    MSHFlexGrid1.ColAlignment(0) = flexAlignCenterCenter
    MSHFlexGrid1.ColAlignment(2) = flexAlignCenterCenter
    MSHFlexGrid1.ColAlignment(3) = flexAlignCenterCenter
    MSHFlexGrid1.ColAlignment(4) = flexAlignCenterCenter
    MSHFlexGrid1.ColAlignment(5) = flexAlignCenterCenter
End Sub


Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
On Error Resume Next
        
    Dim x As Integer
    Dim y As Integer
    
    lblRegNo.Caption = reg
    
    Set document = WebBrowser1.document
    Set inp = document.getElementById("TextReg")
    Set spa = document.getElementById("StName")
    Set btn = document.getElementById("Button3")
    
    If inp Is Nothing = False Then
        inp.Value = reg
    End If
    
    If spa Is Nothing = False Then
        lblName.Caption = spa.innerText
    End If
    
    If btn Is Nothing = False Then
        btn.Click
    End If
    
    For Each element In document.All
        If InStr(1, element.ID, "GridView1") > 0 Then
            Set Tbl = element
            For x = 1 To Tbl.rows.Length - 1
                MSHFlexGrid1.rows = Tbl.rows.Length - 1
                icount = Tbl.rows.Length
                Set tr = Tbl.rows(x)
                For y = 0 To tr.cells.Length - 1
                    Set tc = tr.cells(y)
                    MSHFlexGrid1.TextMatrix(x - 1, y) = tc.innerText
                Next y
            Next x
        End If
    Next
    cmdInsertMarks.Enabled = True
End Sub

