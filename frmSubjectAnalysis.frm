VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSChrt20.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkusercontrolsxp.ocx"
Begin VB.Form frmSubjectAnalysis 
   BorderStyle     =   0  'None
   ClientHeight    =   8670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10575
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   10575
   ShowInTaskbar   =   0   'False
   Begin vkUserContolsXP.vkFrame fMSChart 
      Height          =   3255
      Left            =   240
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4680
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   5741
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
      RoundAngle      =   0
      BorderWidth     =   3
   End
   Begin MSChart20Lib.MSChart MSChart 
      Height          =   3255
      Left            =   240
      OleObjectBlob   =   "frmSubjectAnalysis.frx":0000
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4680
      Width           =   10095
   End
   Begin vkUserContolsXP.vkFrame fGrid 
      Height          =   3375
      Left            =   240
      TabIndex        =   10
      Top             =   1200
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   5953
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
      RoundAngle      =   0
      BorderWidth     =   3
   End
   Begin JURA.StylerButton btnClose 
      Height          =   255
      Left            =   10080
      TabIndex        =   1
      Top             =   0
      Width           =   390
      _ExtentX        =   688
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
   Begin vkUserContolsXP.vkFrame fSubjectAnalysis 
      Height          =   8655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   15266
      Caption         =   "Subject Wise Analysis"
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
      Begin vkUserContolsXP.vkLabel lblSec 
         Height          =   255
         Left            =   6480
         TabIndex        =   14
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.ComboBox cmbSec 
         Height          =   315
         Left            =   6480
         TabIndex        =   13
         Top             =   720
         Width           =   1335
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   3375
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   5953
         _Version        =   393216
         Rows            =   11
         Cols            =   9
         FixedCols       =   0
         RowHeightMin    =   300
         _NumberOfBands  =   1
         _Band(0).Cols   =   9
      End
      Begin vkUserContolsXP.vkCommand cmdPrint 
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   8040
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   661
         Caption         =   "Print"
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
         Left            =   5040
         TabIndex        =   7
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cmbDept 
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   3255
      End
      Begin vkUserContolsXP.vkLabel lblSemester 
         Height          =   255
         Left            =   5040
         TabIndex        =   5
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
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
      Begin vkUserContolsXP.vkLabel lblDept 
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
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
      Begin vkUserContolsXP.vkLabel lblBatch 
         Height          =   255
         Left            =   3600
         TabIndex        =   3
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
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
      Begin VB.ComboBox cmbBatch 
         Height          =   315
         Left            =   3600
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmSubjectAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MSHFlexLoad()
    MSHFlexGrid1.Clear
    MSHFlexGrid1.ColWidth(0) = 905
    MSHFlexGrid1.ColWidth(1) = 3500
    MSHFlexGrid1.RowHeightMin = 300
    For iColumn = 2 To 8
        MSHFlexGrid1.ColAlignment(iColumn) = flexAlignCenterCenter
        MSHFlexGrid1.ColAlignmentFixed(iColumn) = flexAlignCenterCenter
        MSHFlexGrid1.ColWidth(iColumn) = 802.5
    Next
    MSHFlexGrid1.TextMatrix(0, 0) = "Subj Code"
    MSHFlexGrid1.ColAlignmentFixed(0) = flexAlignCenterCenter
    MSHFlexGrid1.ColAlignment(0) = flexAlignCenterCenter
    MSHFlexGrid1.ColAlignmentFixed(1) = flexAlignCenterCenter
    MSHFlexGrid1.TextMatrix(0, 1) = "Subject Name"
    MSHFlexGrid1.TextMatrix(0, 2) = "Appeared"
    MSHFlexGrid1.TextMatrix(0, 3) = "Passed"
    MSHFlexGrid1.TextMatrix(0, 4) = "Pass %"
    MSHFlexGrid1.TextMatrix(0, 5) = "Failed"
    MSHFlexGrid1.TextMatrix(0, 6) = "Maximum"
    MSHFlexGrid1.TextMatrix(0, 7) = "Minimum"
    MSHFlexGrid1.TextMatrix(0, 8) = "Average"
    fGrid.Width = MSHFlexGrid1.Width - 20
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub cmbBatch_Click()
    iBatch = cmbBatch.Text
    Call SubjLoad
End Sub

Private Sub cmbDept_Click()
    iDept = Department(cmbDept)
    Call SubjLoad
End Sub

Private Sub cmbSec_Change()
    strSec = cmbSec.Text
End Sub

Private Sub cmbSec_Click()
    strSec = cmbSec.Text
End Sub

Private Sub cmbSem_Click()
    iSem = cmbSem.Text
    Call SubjLoad
End Sub

Private Sub cmdPrint_Click()
    If iBatch > 2007 Then
        Call rptGradePrint
    Else
        Call rptPrint
    End If
End Sub
Private Sub rptPrint()
    On Error Resume Next
    SavePicture CaptureGraph(Me), App.Path & "/Images/graph.jpg"
    Dim PDF As New clsPDF
    Dim i, j As Integer
    PDF.PDFTitle = "Subject Report"
    PDF.PDFFileName = App.Path & "\Reports\" & "Subject" & " Report " & ".pdf"
    PDF.PDFLoadAfm = App.Path & "\Fonts"
    PDF.PDFView = True
    
    PDF.PDFSetBorder = BORDER_ALL
    PDF.PDFSetTopMargin = 1
    PDF.PDFSetBottomMargin = 1
    PDF.PDFSetRightMargin = 1
    PDF.PDFSetLeftMargin = 1
    PDF.PDFBeginDoc
        
        PDF.PDFDrawRectangle 1, 1, 19, 27
        
        PDF.PDFSetTextColor = vbBlack
        PDF.PDFSetFont FONT_TIMES, 18, FONT_BOLD
        PDF.PDFTextOut "FRANCIS XAVIER ENGINEERING COLLEGE", 2.8, 2
        PDF.PDFSetFont FONT_TIMES, 16, FONT_BOLD
        PDF.PDFTextOut "Tirunelveli-627003", 7.5, 2.75
        PDF.PDFSetFont FONT_TIMES, 14, FONT_BOLD
        PDF.PDFTextOut "RESULT ANALYSIS OF EVEN SEMESTER (2010-2011) - AU TIRUNELVELI", 1.15, 3.65
        PDF.PDFTextOut "DEPARTMENT OF " & UCase(cmbDept.Text), 3.8, 4.5
        PDF.PDFTextOut "SUBJECT WISE ANALYSIS", 6.25, 5.15
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 1, 5.35, 20, 5.35
        PDF.PDFSetLineWidth = 0.025
        PDF.PDFDrawLine 1, 5.4, 20, 5.4
        
        PDF.PDFSetFont 2, 10, FONT_BOLD
        PDF.PDFSetFont 2, 12, FONT_BOLD
        PDF.PDFTextOut "Batch:", 2.5, 6
        PDF.PDFTextOut cmbBatch.Text, 4, 6
        PDF.PDFTextOut "Semester:", 13, 6
        PDF.PDFTextOut cmbSem.Text, 15.25, 6
        
                      
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 1, 6.25, 20, 6.25
        PDF.PDFSetLineWidth = 0.02
        PDF.PDFDrawLine 1, 6.3, 20, 6.3
        
        
        PDF.PDFSetFont 2, 10, FONT_NORMAL
        PDF.PDFTextOut "Subject", 0.5, 6.85
        PDF.PDFTextOut "Subject Name", 2.5, 6.85
        PDF.PDFTextOut "Appeared", 9.25, 6.85
        PDF.PDFTextOut "Passed", 11, 6.85
        PDF.PDFTextOut "Pass %", 12.4, 6.85
        PDF.PDFTextOut "Failed", 13.7, 6.85
        PDF.PDFTextOut "Max", 14.95, 6.85
        PDF.PDFTextOut "Min", 16.15, 6.85
        PDF.PDFTextOut "Avg", 17.65, 6.85
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 1, 7, 20, 7
        PDF.PDFSetLineWidth = 0.025
        PDF.PDFDrawLine 1, 7.05, 20, 7.05
        
        PDF.PDFSetTextColor = vbBlack
        PDF.PDFSetFont 2, 10, FONT_NORMAL
        For i = 1 To MSHFlexGrid1.rows
            PDF.PDFTextOut MSHFlexGrid1.TextMatrix(i, 0), 0.5, 7.3 + i * 0.6    'Subj Code
            If Len(MSHFlexGrid1.TextMatrix(i, 1)) > 40 Then
                PDF.PDFSetFont 2, 8, FONT_NORMAL
                PDF.PDFTextOut MSHFlexGrid1.TextMatrix(i, 1), 2.5, 7.3 + i * 0.6   'Subj Name
            Else
                PDF.PDFSetFont 2, 10, FONT_NORMAL
                PDF.PDFTextOut MSHFlexGrid1.TextMatrix(i, 1), 2.5, 7.3 + i * 0.6   'Subj Name
            End If
            PDF.PDFTextOut MSHFlexGrid1.TextMatrix(i, 2), 9.75, 7.3 + i * 0.6   'Appeared
            PDF.PDFTextOut MSHFlexGrid1.TextMatrix(i, 3), 11.35, 7.3 + i * 0.6  'Passed
            PDF.PDFTextOut MSHFlexGrid1.TextMatrix(i, 4), 12.5, 7.3 + i * 0.6   'Pass %
            PDF.PDFTextOut MSHFlexGrid1.TextMatrix(i, 5), 14, 7.3 + i * 0.6     'Failed
            PDF.PDFTextOut MSHFlexGrid1.TextMatrix(i, 6), 15.1, 7.3 + i * 0.6     'Max
            PDF.PDFTextOut MSHFlexGrid1.TextMatrix(i, 7), 16.3, 7.3 + i * 0.6  'Min
            PDF.PDFTextOut MSHFlexGrid1.TextMatrix(i, 8), 17.55, 7.3 + i * 0.6  'Avg
        Next
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 3.25, 6.3, 3.25, 6.7 + i * 0.6
        
        PDF.PDFImage App.Path & "/Images/graph.jpg", 2, 5, 5, 5
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 10.2, 6.3, 10.2, 6.7 + i * 0.6
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 11.8, 6.3, 11.8, 6.7 + i * 0.6
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 13.25, 6.3, 13.25, 6.7 + i * 0.6
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 14.65, 6.3, 14.65, 6.7 + i * 0.6
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 15.75, 6.3, 15.75, 6.7 + i * 0.6
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 16.85, 6.3, 16.85, 6.7 + i * 0.6
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 18.15, 6.3, 18.15, 6.7 + i * 0.6
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 1, 6.7 + i * 0.6, 20, 6.7 + i * 0.6
        PDF.PDFSetLineWidth = 0.025
        PDF.PDFDrawLine 1, 6.75 + i * 0.6, 20, 6.75 + i * 0.6
        
        PDF.PDFSetFont FONT_TIMES, 12, FONT_BOLD
        PDF.PDFTextOut "Staff Incharge", 2.5, 24.75
        PDF.PDFTextOut "HOD", 9, 24.75
        PDF.PDFTextOut "Principal", 15, 24.75
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 1, 27.25, 20, 27.25
        PDF.PDFSetLineWidth = 0.025
        PDF.PDFDrawLine 1, 27.3, 20, 27.3
        PDF.PDFTextOut "Report Generated By JURA", 7.5, 27.75
    PDF.PDFEndDoc
End Sub
Private Sub rptGradePrint()
    On Error Resume Next
    'SavePicture CaptureGraph(Me), App.Path & "/Images/graph.jpg"
    Dim PDF As New clsPDF
    Dim i, j As Integer
    PDF.PDFTitle = "Subject Report"
    PDF.PDFFileName = App.Path & "\Reports\" & "Subject" & " Report " & ".pdf"
    PDF.PDFLoadAfm = App.Path & "\Fonts"
    PDF.PDFView = True
    
    PDF.PDFSetBorder = BORDER_ALL
    PDF.PDFSetTopMargin = 1
    PDF.PDFSetBottomMargin = 1
    PDF.PDFSetRightMargin = 1
    PDF.PDFSetLeftMargin = 1
    PDF.PDFBeginDoc
        
        PDF.PDFDrawRectangle 1, 1, 19, 27
        
        PDF.PDFSetTextColor = vbBlack
        PDF.PDFSetFont FONT_TIMES, 18, FONT_BOLD
        PDF.PDFTextOut "FRANCIS XAVIER ENGINEERING COLLEGE", 2.8, 2
        PDF.PDFSetFont FONT_TIMES, 16, FONT_BOLD
        PDF.PDFTextOut "Tirunelveli-627003", 7.5, 2.75
        PDF.PDFSetFont FONT_TIMES, 14, FONT_BOLD
        PDF.PDFTextOut "RESULT ANALYSIS OF EVEN SEMESTER (2010-2011) - AU TIRUNELVELI", 1.15, 3.65
        PDF.PDFTextOut "DEPARTMENT OF " & UCase(cmbDept.Text), 3.8, 4.5
        PDF.PDFTextOut "SUBJECT WISE ANALYSIS", 6.25, 5.15
        
        PDF.PDFImage App.Path & "/Images/graph.jpg", 3, 15, 15, 6
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 1, 5.35, 20, 5.35
        PDF.PDFSetLineWidth = 0.025
        PDF.PDFDrawLine 1, 5.4, 20, 5.4
        
        PDF.PDFSetFont 2, 10, FONT_BOLD
        PDF.PDFSetFont 2, 12, FONT_BOLD
        PDF.PDFTextOut "Batch:", 2.5, 6
        PDF.PDFTextOut cmbBatch.Text, 4, 6
        PDF.PDFTextOut "Semester:", 13, 6
        PDF.PDFTextOut cmbSem.Text, 15.25, 6
        
                      
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 1, 6.25, 20, 6.25
        PDF.PDFSetLineWidth = 0.02
        PDF.PDFDrawLine 1, 6.3, 20, 6.3
        
        
        PDF.PDFSetFont 2, 10, FONT_NORMAL
        PDF.PDFTextOut "Subject", 0.5, 6.85
        PDF.PDFTextOut "Subject Name", 2.5, 6.85
        PDF.PDFTextOut "Appeared", 9.25, 6.85
        PDF.PDFTextOut "Passed", 11, 6.85
        PDF.PDFTextOut "Pass %", 12.4, 6.85
        PDF.PDFTextOut "Failed", 13.7, 6.85
        
        'PDF.PDFTextOut "Max", 14.95, 6.85
        'PDF.PDFTextOut "Min", 16.15, 6.85
        'PDF.PDFTextOut "Avg", 17.65, 6.85
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 1, 7, 20, 7
        PDF.PDFSetLineWidth = 0.025
        PDF.PDFDrawLine 1, 7.05, 20, 7.05
        
        PDF.PDFSetTextColor = vbBlack
        PDF.PDFSetFont 2, 10, FONT_NORMAL
        For i = 1 To MSHFlexGrid1.rows
            PDF.PDFTextOut MSHFlexGrid1.TextMatrix(i, 0), 0.5, 7.3 + i * 0.6    'Subj Code
            If Len(MSHFlexGrid1.TextMatrix(i, 1)) > 40 Then
                PDF.PDFSetFont 2, 8, FONT_NORMAL
                PDF.PDFTextOut MSHFlexGrid1.TextMatrix(i, 1), 2.5, 7.3 + i * 0.6   'Subj Name
                PDF.PDFSetFont 2, 10, FONT_NORMAL
            Else
                PDF.PDFSetFont 2, 10, FONT_NORMAL
                PDF.PDFTextOut MSHFlexGrid1.TextMatrix(i, 1), 2.5, 7.3 + i * 0.6   'Subj Name
            End If
            PDF.PDFTextOut MSHFlexGrid1.TextMatrix(i, 2), 9.75, 7.3 + i * 0.6   'Appeared
            PDF.PDFTextOut MSHFlexGrid1.TextMatrix(i, 3), 11.35, 7.3 + i * 0.6  'Passed
            PDF.PDFTextOut MSHFlexGrid1.TextMatrix(i, 4), 12.5, 7.3 + i * 0.6   'Pass %
            PDF.PDFTextOut MSHFlexGrid1.TextMatrix(i, 5), 14, 7.3 + i * 0.6     'Failed
            'PDF.PDFTextOut MSHFlexGrid1.TextMatrix(i, 6), 15.1, 7.3 + i * 0.6     'Max
            'PDF.PDFTextOut MSHFlexGrid1.TextMatrix(i, 7), 16.3, 7.3 + i * 0.6  'Min
            'PDF.PDFTextOut MSHFlexGrid1.TextMatrix(i, 8), 17.55, 7.3 + i * 0.6  'Avg
        Next
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 3.25, 6.3, 3.25, 6.7 + i * 0.6
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 10.2, 6.3, 10.2, 6.7 + i * 0.6
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 11.8, 6.3, 11.8, 6.7 + i * 0.6
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 13.25, 6.3, 13.25, 6.7 + i * 0.6
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 14.65, 6.3, 14.65, 6.7 + i * 0.6
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 15.75, 6.3, 15.75, 6.7 + i * 0.6
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 16.85, 6.3, 16.85, 6.7 + i * 0.6
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 18.15, 6.3, 18.15, 6.7 + i * 0.6
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 1, 6.7 + i * 0.6, 20, 6.7 + i * 0.6
        PDF.PDFSetLineWidth = 0.025
        PDF.PDFDrawLine 1, 6.75 + i * 0.6, 20, 6.75 + i * 0.6
        
        PDF.PDFSetFont FONT_TIMES, 12, FONT_BOLD
        PDF.PDFTextOut "Staff Incharge", 2.5, 24.75
        PDF.PDFTextOut "HOD", 9, 24.75
        PDF.PDFTextOut "Principal", 15, 24.75
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 1, 27.25, 20, 27.25
        PDF.PDFSetLineWidth = 0.025
        PDF.PDFDrawLine 1, 27.3, 20, 27.3
        PDF.PDFTextOut "Report Generated By JURA", 7.5, 27.75
    PDF.PDFEndDoc
End Sub




Private Sub Form_Load()
    CreateRoundRectFromWindow Me, 7, 7
    Me.Left = (mdiMain.Width - Me.Width) / 2
    Me.Top = 100
    Me.BackColor = mdiMain.BackColor
    Call frmColor(frmSubjectAnalysis)
    Call cmbDept_Load(cmbDept)
    Call cmbSem_Load(cmbSem)
    Call cmbSec_Load(cmbSec)
    Call cmbBatch_Load(cmbBatch)
    Call MSHFlexLoad
    Call SubjLoad
End Sub
Private Sub SubjLoad()
    On Error Resume Next
    MSHFlexGrid1.Clear
    Call MSHFlexLoad
    Dim rs As New ADODB.Recordset
    sql = "select subjcode from subj where semno= " & cmbSem.Text & " and dept = " & Department(cmbDept) & " and batch=" & Mid(cmbBatch.Text, 3, 2) & ""
    rs.CursorLocation = adUseClient
    rs.Open sql, conn, adOpenDynamic, adLockOptimistic
    If iBatch > 2007 Then
        For i = 1 To rs.RecordCount
            MSHFlexGrid1.TextMatrix(i, 0) = rs.Fields(0)
            MSHFlexGrid1.TextMatrix(i, 1) = GetSubjName(rs.Fields(0))
            MSHFlexGrid1.TextMatrix(i, 2) = GetNoOfStudAppeared(iSem, iDept, iBatch, rs.Fields(0))
            MSHFlexGrid1.TextMatrix(i, 3) = GetGradePassedCount(iSem, iDept, iBatch, rs.Fields(0), "('S','A','B','C','D','E')")
            MSHFlexGrid1.TextMatrix(i, 4) = Round((MSHFlexGrid1.TextMatrix(i, 3) / MSHFlexGrid1.TextMatrix(i, 2)) * 100, 2)
            MSHFlexGrid1.TextMatrix(i, 5) = GetGradePassedCount(iSem, iDept, iBatch, rs.Fields(0), "('U','I','W')")
            MSHFlexGrid1.TextMatrix(i, 6) = ""
            MSHFlexGrid1.TextMatrix(i, 7) = ""
            MSHFlexGrid1.TextMatrix(i, 8) = ""
            rs.MoveNext
        Next
    Else
        For i = 1 To rs.RecordCount
            MSHFlexGrid1.TextMatrix(i, 0) = rs.Fields(0)
            MSHFlexGrid1.TextMatrix(i, 1) = GetSubjName(rs.Fields(0))
            MSHFlexGrid1.TextMatrix(i, 2) = GetNoOfStudAppeared(iSem, iDept, iBatch, rs.Fields(0))
            MSHFlexGrid1.TextMatrix(i, 3) = GetCount(iSem, iDept, iBatch, rs.Fields(0), 50, 100)
            MSHFlexGrid1.TextMatrix(i, 4) = Round((MSHFlexGrid1.TextMatrix(i, 3) / MSHFlexGrid1.TextMatrix(i, 2)) * 100, 2)
            MSHFlexGrid1.TextMatrix(i, 5) = GetCount(iSem, iDept, iBatch, rs.Fields(0), 0, 49)
            MSHFlexGrid1.TextMatrix(i, 6) = GetMaxMarks(iSem, iDept, iBatch, rs.Fields(0))
            MSHFlexGrid1.TextMatrix(i, 7) = GetMinMarks(iSem, iDept, iBatch, rs.Fields(0))
            MSHFlexGrid1.TextMatrix(i, 8) = GetAvgMarks(iSem, iDept, iBatch, rs.Fields(0))
            rs.MoveNext
        Next
    End If
    MSChart_Load
End Sub


Private Sub fSubjectAnalysis_MouseDown(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
    ReleaseCapture
    SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub MSChart_Load()
    MSChart.RowCount = GetSubjCount(iSem, iDept, iBatch)
    MSChart.ColumnCount = 1
    
    For i = 1 To MSChart.RowCount
        MSChart.Row = i
        MSChart.RowLabel = MSHFlexGrid1.TextMatrix(i, 0)
        MSChart.Data = MSHFlexGrid1.TextMatrix(i, 4)
    Next
End Sub

