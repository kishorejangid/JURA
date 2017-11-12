VERSION 5.00
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkusercontrolsxp.ocx"
Begin VB.Form frmClass 
   BorderStyle     =   0  'None
   ClientHeight    =   3480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4440
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   Begin vkUserContolsXP.vkLabel lblSec 
      Height          =   255
      Left            =   3120
      TabIndex        =   12
      Top             =   1080
      Width           =   1095
      _ExtentX        =   1931
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
   Begin vkUserContolsXP.vkFrame fClass 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   6165
      Caption         =   "Class Report"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleGradient   =   2
      TitleHeight     =   300
      BorderWidth     =   2
      Begin vkUserContolsXP.vkCommand cmdPrint 
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   2400
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   661
         Caption         =   "Generate Report"
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
      Begin vkUserContolsXP.vkBar prgReport 
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   2880
         Visible         =   0   'False
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   661
         Decimals        =   0
         Value           =   1
         BackPicture     =   "frmClass.frx":0000
         FrontPicture    =   "frmClass.frx":001C
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
         Left            =   3960
         TabIndex        =   13
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
      Begin VB.ComboBox cmbSec 
         Height          =   315
         Left            =   3120
         TabIndex        =   11
         Top             =   1320
         Width           =   1095
      End
      Begin vkUserContolsXP.vkLabel vkLabel3 
         Height          =   270
         Left            =   240
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1680
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
      Begin vkUserContolsXP.vkLabel vkLabel4 
         Height          =   270
         Left            =   2400
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.ComboBox cmbStartRegNo 
         Height          =   315
         Left            =   240
         TabIndex        =   8
         Top             =   1920
         Width           =   1935
      End
      Begin VB.ComboBox cmbEndRegNo 
         Height          =   315
         Left            =   2400
         TabIndex        =   7
         Top             =   1920
         Width           =   1815
      End
      Begin VB.ComboBox cmbBatch 
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   1095
      End
      Begin vkUserContolsXP.vkLabel vkLabel1 
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1080
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
      Begin vkUserContolsXP.vkLabel lblSemester 
         Height          =   255
         Left            =   1680
         TabIndex        =   3
         Top             =   1080
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
      Begin VB.ComboBox cmbDept 
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   3975
      End
      Begin VB.ComboBox cmbSem 
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Top             =   1320
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub cmbBatch_Change()
    Call cmbBatch_Click
End Sub

Private Sub cmbBatch_Click()
    On Error Resume Next
    iBatch = cmbBatch.Text
    Call cmbRegNo_Load(cmbStartRegNo)
    Call cmbRegNo_Load(cmbEndRegNo)
    cmbStartRegNo.Text = cmbStartRegNo.List(Index)
    cmbEndRegNo.Text = cmbEndRegNo.List(cmbEndRegNo.ListCount - 1)
    If Err.Number = 13 Then
        MsgBox "Error:" & vbCrLf & Err.Description, vbCritical, Error
        cmbBatch.Text = ""
    End If
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

Private Sub cmbSec_Change()
    strSec = cmbSec.Text
    Call cmbRegNo_Load(cmbStartRegNo)
    Call cmbRegNo_Load(cmbEndRegNo)
    cmbStartRegNo.Text = cmbStartRegNo.List(Index)
    cmbEndRegNo.Text = cmbEndRegNo.List(cmbEndRegNo.ListCount - 1)
End Sub
Private Sub cmbSec_Click()
    strSec = cmbSec.Text
    Call cmbRegNo_Load(cmbStartRegNo)
    Call cmbRegNo_Load(cmbEndRegNo)
    cmbStartRegNo.Text = cmbStartRegNo.List(Index)
    cmbEndRegNo.Text = cmbEndRegNo.List(cmbEndRegNo.ListCount - 1)
End Sub

Private Sub cmbSem_Change()
    Call cmbSem_Click
End Sub

Private Sub cmbSem_Click()
    On Error Resume Next
    iSem = cmbSem.Text
    If Err.Number = 13 Then
        MsgBox "Error:" & vbLf & Err.Description, vbCritical, "Error"
        cmbSem.Text = ""
    End If
End Sub

Private Sub cmdPrint_Click()
    If iBatch > 2007 Then
        Call CreateGradePDF
    Else
        Call CreatePDF
    End If
End Sub

Private Sub Form_Load()
    CreateRoundRectFromWindow Me, 7, 7
    Me.BackColor = mdiMain.BackColor
    Call frmColor(frmClass)
    Call cmbDept_Load(cmbDept)
    Call cmbBatch_Load(cmbBatch)
    Call cmbSem_Load(cmbSem)
    Call cmbSec_Load(cmbSec)
    Call cmbRegNo_Load(cmbStartRegNo)
    Call cmbRegNo_Load(cmbEndRegNo)
    cmbStartRegNo.Text = cmbStartRegNo.List(Index)
    cmbEndRegNo.Text = cmbEndRegNo.List(cmbEndRegNo.ListCount - 1)
End Sub
Private Sub CreatePDF()
    On Error Resume Next
    
    Dim sqlRegNo As String
    Dim sqlMarks As String
    Dim strRegNo As String
    Dim iTotal  As Integer
    Dim PDF As New clsPDF
    Dim i, j, k As Integer
    Dim rsRegNo As New ADODB.Recordset
    Dim rsMarks As New ADODB.Recordset
    Dim iprgValue As Integer
    
    rsRegNo.CursorLocation = adUseClient
    rsMarks.CursorLocation = adUseClient
    
    prgReport.Visible = True
    prgReport.Value = 0
    
    PDF.PDFTitle = "MarksSheet"
    PDF.PDFFileName = App.Path & "\Reports\Class Report" & "(" & cmbSem.Text & ")" & ".pdf"
    PDF.PDFLoadAfm = App.Path & "\Fonts"
    
    PDF.PDFView = True
    
    PDF.PDFSetBorder = BORDER_ALL
    PDF.PDFSetTopMargin = 1
    PDF.PDFSetBottomMargin = 1
    PDF.PDFSetRightMargin = 1
    PDF.PDFSetLeftMargin = 1
    PDF.PDFBeginDoc
        
        sqlRegNo = "select regno from studdetails where substr(regno, 6, 3) = '" & iDept & "' and substr(regno, 4, 2) = '" & Mid(iBatch, 3, 2) & "' and sec='" & strSec & "' and regno between '" & cmbStartRegNo.Text & "' and  '" & cmbEndRegNo.Text & "'  order by regno"
        rsRegNo.Open sqlRegNo, conn, adOpenDynamic, adLockOptimistic
        iprgValue = 100 / rsRegNo.RecordCount
        For i = 1 To rsRegNo.RecordCount / 2
                    
            PDF.PDFNewPage
        
                PDF.PDFDrawRectangle 1, 1, 19, 27
                
                PDF.PDFSetFont FONT_TIMES, 24, FONT_BOLD
                PDF.PDFSetTextColor = vbRed
                PDF.PDFTextOut "Francis Xavier Engineering College", 2.9, 2
                PDF.PDFSetFont FONT_TIMES, 18, FONT_BOLD
                PDF.PDFTextOut "Marks Sheet", 8, 3.25
                
                PDF.PDFSetLineWidth = 0.03
                PDF.PDFDrawLine 1, 3.65, 20, 3.65
                PDF.PDFSetLineWidth = 0.025
                PDF.PDFDrawLine 1, 3.7, 20, 3.7
                
                PDF.PDFSetTextColor = vbBlack
                PDF.PDFSetFont FONT_TIMES, 12, FONT_NORMAL
                PDF.PDFTextOut "Department:", 1, 4.25
                PDF.PDFTextOut cmbDept.Text, 3.35, 4.25
                PDF.PDFTextOut "Batch:", 10, 4.25
                PDF.PDFTextOut cmbBatch.Text, 11.25, 4.25
                PDF.PDFTextOut "Semester:", 14.5, 4.25
                PDF.PDFTextOut cmbSem.Text, 16.25, 4.25
                
                PDF.PDFSetLineWidth = 0.03
                PDF.PDFDrawLineHor 1, 4.5, 19
                PDF.PDFSetLineWidth = 0.02
                PDF.PDFDrawLine 1, 4.54, 20, 4.54
                
                
                
                PDF.PDFTextOut "Subject Code", 0.75, 5
                PDF.PDFTextOut "Subject Name", 3.5, 5
                PDF.PDFTextOut "Internals", 10.25, 5
                PDF.PDFTextOut "Externals", 12.5, 5
                PDF.PDFTextOut "Marks", 14.85, 5
                PDF.PDFTextOut "Result", 17, 5
                
                
                PDF.PDFSetLineWidth = 0.03
                PDF.PDFDrawLine 1, 5.25, 20, 5.25
                PDF.PDFSetLineWidth = 0.025
                PDF.PDFDrawLine 1, 5.3, 20, 5.3
                
                'Student 1
                
                strRegNo = rsRegNo.Fields(0)
                sqlMarks = "select subjcode,internals,externals,result from studmarks where semno=" & iSem & " and batch=" & Mid(iBatch, 3, 2) & "and regno='" & strRegNo & "' order by subjcode"
                rsMarks.Open sqlMarks, conn, adOpenDynamic, adLockOptimistic
                
                PDF.PDFTextOut "Register No:", 1, 6
                PDF.PDFTextOut strRegNo, 3.5, 6
                PDF.PDFTextOut "Student Name:", 10, 6
                
                If Len(GetStudName(strRegNo)) > 20 Then
                    PDF.PDFSetFont FONT_TIMES, 10, FONT_NORMAL
                    PDF.PDFTextOut GetStudName(strRegNo), 13, 6
                Else
                    PDF.PDFSetFont FONT_TIMES, 12, FONT_NORMAL
                    PDF.PDFTextOut GetStudName(strRegNo), 13, 6
                End If
                                                    
                iTotal = 0
                PDF.PDFSetFont FONT_TIMES, 12, FONT_NORMAL
                
                For j = 1 To rsMarks.RecordCount
                
                    PDF.PDFTextOut rsMarks.Fields(0), 1, 6.25 + j * 0.7
                    
                    If Len(GetSubjName(rsMarks.Fields(0))) > 40 Then
                        PDF.PDFSetFont FONT_TIMES, 10, FONT_NORMAL
                        PDF.PDFTextOut GetSubjName(rsMarks.Fields(0)), 3.5, 6.25 + j * 0.7
                        PDF.PDFSetFont FONT_TIMES, 12, FONT_NORMAL
                    Else
                        PDF.PDFSetFont FONT_TIMES, 12, FONT_NORMAL
                        PDF.PDFTextOut GetSubjName(rsMarks.Fields(0)), 3.5, 6.25 + j * 0.7
                    End If
                                        
                    If rsMarks.Fields(3) = "WH2" Then
                        PDF.PDFSetTextColor = vbRed
                        PDF.PDFTextOut " ", 15.25, 6.25 + j * 0.7
                        PDF.PDFSetFont FONT_TIMES, 8, FONT_NORMAL
                        PDF.PDFTextOut "WH2", 17.5, 6.25 + j * 0.7
                        PDF.PDFSetFont FONT_TIMES, 12, FONT_NORMAL
                        PDF.PDFSetTextColor = vbBlack
                    ElseIf rsMarks.Fields(3) = "WH1" Then
                        PDF.PDFSetTextColor = vbRed
                        PDF.PDFTextOut " ", 15.25, 6.25 + j * 0.7
                        PDF.PDFSetFont FONT_TIMES, 8, FONT_NORMAL
                        PDF.PDFTextOut "WH1", 17.5, 6.25 + j * 0.7
                        PDF.PDFSetFont FONT_TIMES, 12, FONT_NORMAL
                        PDF.PDFSetTextColor = vbBlack
                    ElseIf rsMarks.Fields(3) = "WH" Then
                        PDF.PDFSetTextColor = vbRed
                        PDF.PDFTextOut " ", 15.25, 6.25 + j * 0.7
                        PDF.PDFSetFont FONT_TIMES, 8, FONT_NORMAL
                        PDF.PDFTextOut "WH", 17.5, 6.25 + j * 0.7
                        PDF.PDFSetFont FONT_TIMES, 12, FONT_NORMAL
                        PDF.PDFSetTextColor = vbBlack
                    ElseIf rsMarks.Fields(3) = "AB" Then
                        PDF.PDFTextOut rsMarks.Fields(1), 11, 6.25 + j * 0.7
                        PDF.PDFTextOut "AB", 13.2, 6.25 + j * 0.7
                        PDF.PDFSetTextColor = vbRed
                        PDF.PDFTextOut " ", 15.25, 6.25 + j * 0.7
                        PDF.PDFTextOut "AB", 17.5, 6.25 + j * 0.7
                        PDF.PDFSetTextColor = vbBlack
                    ElseIf rsMarks.Fields(3) = "SA" Then
                        PDF.PDFTextOut " ", 11, 6.25 + j * 0.7
                        PDF.PDFTextOut " ", 13.2, 6.25 + j * 0.7
                        PDF.PDFSetTextColor = vbRed
                        PDF.PDFTextOut " ", 15.25, 6.25 + j * 0.7
                        PDF.PDFTextOut "SA", 17.5, 6.25 + j * 0.7
                        PDF.PDFSetTextColor = vbBlack
                    Else
                        PDF.PDFTextOut rsMarks.Fields(1), 11, 6.25 + j * 0.7
                        PDF.PDFTextOut rsMarks.Fields(2), 13.2, 6.25 + j * 0.7
                        
                        If (rsMarks.Fields(1) + rsMarks.Fields(2)) < 50 Then
                            PDF.PDFSetTextColor = vbRed
                            PDF.PDFTextOut rsMarks.Fields(1) + rsMarks.Fields(2), 15.25, 6.25 + j * 0.7
                            PDF.PDFTextOut "F", 17.5, 6.25 + j * 0.7
                            PDF.PDFSetTextColor = vbBlack
                            iTotal = iTotal + (rsMarks.Fields(1) + rsMarks.Fields(2))
                        Else
                            PDF.PDFTextOut rsMarks.Fields(1) + rsMarks.Fields(2), 15.25, 6.25 + j * 0.7
                            PDF.PDFTextOut "P", 17.5, 6.25 + j * 0.7
                            iTotal = iTotal + (rsMarks.Fields(1) + rsMarks.Fields(2))
                        End If
                    End If
                    rsMarks.MoveNext
                Next
                PDF.PDFSetTextColor = vbBlack
                PDF.PDFSetLineWidth = 0.03
                PDF.PDFDrawLine 1, 6.25 + (j + 0.75) * 0.7, 20, 6.25 + (j + 0.75) * 0.7
                
                                
                PDF.PDFTextOut "Total:", 13.5, 6.25 + (j + 1.45) * 0.7
                PDF.PDFTextOut CStr(iTotal), 15.15, 6.25 + (j + 1.45) * 0.7
                
                PDF.PDFTextOut "Percentage:", 1, 6.25 + (j + 1.45) * 0.7
                PDF.PDFTextOut Round(iTotal / GetSubjCount(iSem, iDept, iBatch), 2), 3.5, 6.25 + (j + 1.45) * 0.7
            
                PDF.PDFSetLineWidth = 0.03
                PDF.PDFDrawLine 1, 6.25 + (j + 1.75) * 0.7, 20, 6.25 + (j + 1.75) * 0.7
                PDF.PDFSetLineWidth = 0.025
                PDF.PDFDrawLine 1, 6.3 + (j + 1.75) * 0.7, 20, 6.3 + (j + 1.75) * 0.7
                
                rsMarks.Close
                prgReport.Value = prgReport.Value + iprgValue
                Me.Refresh
                
                'Student 2
                
                rsRegNo.MoveNext
                strRegNo = rsRegNo.Fields(0)
                sqlMarks = "select subjcode,internals,externals,result from studmarks where semno=" & iSem & " and batch=" & Mid(iBatch, 3, 2) & "and regno='" & strRegNo & "' order by subjcode"
                rsMarks.Open sqlMarks, conn, adOpenDynamic, adLockOptimistic
                k = j
                iTotal = 0
                PDF.PDFTextOut "Register No:", 1, 6 + (k + 3.25) * 0.7
                PDF.PDFTextOut strRegNo, 3.5, 6 + (k + 3.25) * 0.7
                PDF.PDFTextOut "Student Name:", 10, 6 + (k + 3.25) * 0.7
                
                If Len(GetStudName(strRegNo)) > 20 Then
                    PDF.PDFSetFont FONT_TIMES, 10, FONT_NORMAL
                    PDF.PDFTextOut GetStudName(strRegNo), 13, 6 + (k + 3.25) * 0.7
                Else
                    PDF.PDFSetFont FONT_TIMES, 12, FONT_NORMAL
                    PDF.PDFTextOut GetStudName(strRegNo), 13, 6 + (k + 3.25) * 0.7
                End If
        
                PDF.PDFSetFont FONT_TIMES, 12, FONT_NORMAL
                
                For j = 1 To rsMarks.RecordCount
                
                    PDF.PDFTextOut rsMarks.Fields(0), 1, 6 + (k + 3.5 + j) * 0.7
                    
                    If Len(GetSubjName(rsMarks.Fields(0))) > 40 Then
                        PDF.PDFSetFont FONT_TIMES, 10, FONT_NORMAL
                        PDF.PDFTextOut GetSubjName(rsMarks.Fields(0)), 3.5, 6 + (k + 3.5 + j) * 0.7
                        PDF.PDFSetFont FONT_TIMES, 12, FONT_NORMAL
                    Else
                        PDF.PDFSetFont FONT_TIMES, 12, FONT_NORMAL
                        PDF.PDFTextOut GetSubjName(rsMarks.Fields(0)), 3.5, 6 + (k + 3.5 + j) * 0.7
                    End If
                    
                    If rsMarks.Fields(3) = "WH2" Then
                        PDF.PDFSetTextColor = vbRed
                        PDF.PDFTextOut " ", 15.25, 6 + (k + 3.5 + j) * 0.7
                        PDF.PDFSetFont FONT_TIMES, 8, FONT_NORMAL
                        PDF.PDFTextOut "WH2", 17.5, 6 + (k + 3.5 + j) * 0.7
                        PDF.PDFSetFont FONT_TIMES, 12, FONT_NORMAL
                        PDF.PDFSetTextColor = vbBlack
                    ElseIf rsMarks.Fields(3) = "WH1" Then
                        PDF.PDFSetTextColor = vbRed
                        PDF.PDFTextOut " ", 15.25, 6 + (k + 3.5 + j) * 0.7
                        PDF.PDFSetFont FONT_TIMES, 8, FONT_NORMAL
                        PDF.PDFTextOut "WH1", 17.5, 6 + (k + 3.5 + j) * 0.7
                        PDF.PDFSetFont FONT_TIMES, 12, FONT_NORMAL
                        PDF.PDFSetTextColor = vbBlack
                    ElseIf rsMarks.Fields(3) = "WH" Then
                        PDF.PDFSetTextColor = vbRed
                        PDF.PDFTextOut " ", 15.25, 6 + (k + 3.5 + j) * 0.7
                        PDF.PDFSetFont FONT_TIMES, 8, FONT_NORMAL
                        PDF.PDFTextOut "WH", 17.5, 6 + (k + 3.5 + j) * 0.7
                        PDF.PDFSetFont FONT_TIMES, 12, FONT_NORMAL
                        PDF.PDFSetTextColor = vbBlack
                    ElseIf rsMarks.Fields(3) = "AB" Then
                        PDF.PDFTextOut rsMarks.Fields(1), 11, 6 + (k + 3.5 + j) * 0.7
                        PDF.PDFTextOut "AB", 13.2, 6 + (k + 3.5 + j) * 0.7
                        PDF.PDFSetTextColor = vbRed
                        PDF.PDFTextOut " ", 15.25, 6 + (k + 3.5 + j) * 0.7
                        PDF.PDFTextOut "AB", 17.5, 6 + (k + 3.5 + j) * 0.7
                        PDF.PDFSetTextColor = vbBlack
                    ElseIf rsMarks.Fields(3) = "SA" Then
                        PDF.PDFTextOut " ", 11, 6 + (k + 3.5 + j) * 0.7
                        PDF.PDFTextOut " ", 13.2, 6 + (k + 3.5 + j) * 0.7
                        PDF.PDFSetTextColor = vbRed
                        PDF.PDFTextOut " ", 15.25, 6 + (k + 3.5 + j) * 0.7
                        PDF.PDFTextOut "SA", 17.5, 6 + (k + 3.5 + j) * 0.7
                        PDF.PDFSetTextColor = vbBlack
                    Else
                        PDF.PDFTextOut rsMarks.Fields(1), 11, 6 + (k + 3.5 + j) * 0.7
                        PDF.PDFTextOut rsMarks.Fields(2), 13.2, 6 + (k + 3.5 + j) * 0.7
                        
                        If (rsMarks.Fields(1) + rsMarks.Fields(2)) < 50 Then
                            PDF.PDFSetTextColor = vbRed
                            PDF.PDFTextOut (rsMarks.Fields(1) + rsMarks.Fields(2)), 15.25, 6 + (k + 3.5 + j) * 0.7
                            PDF.PDFTextOut "F", 17.5, 6 + (k + 3.5 + j) * 0.7
                            PDF.PDFSetTextColor = vbBlack
                            iTotal = iTotal + (rsMarks.Fields(1) + rsMarks.Fields(2))
                        Else
                            PDF.PDFTextOut (rsMarks.Fields(1) + rsMarks.Fields(2)), 15.25, 6 + (k + 3.5 + j) * 0.7
                            PDF.PDFTextOut "P", 17.5, 6 + (k + 3.5 + j) * 0.7
                            iTotal = iTotal + (rsMarks.Fields(1) + rsMarks.Fields(2))
                        End If
                    End If
                    rsMarks.MoveNext
                Next
                
                PDF.PDFSetTextColor = vbBlack
                PDF.PDFSetLineWidth = 0.03
                PDF.PDFDrawLine 1, 6 + (k + 3.5 + j + 0.75) * 0.7, 20, 6 + (k + 3.5 + j + 0.75) * 0.7
                                
                
                PDF.PDFTextOut "Total:", 13.5, 6 + (k + 3.5 + j + 1.45) * 0.7
                PDF.PDFTextOut CStr(iTotal), 15.15, 6 + (k + 3.5 + j + 1.45) * 0.7
                
                PDF.PDFTextOut "Percentage:", 1, 6 + (k + 3.5 + j + 1.45) * 0.7
                PDF.PDFTextOut Round(iTotal / GetSubjCount(iSem, iDept, iBatch), 2), 3.5, 6 + (k + 3.5 + j + 1.45) * 0.7
                                
            
                PDF.PDFSetLineWidth = 0.03
                PDF.PDFDrawLine 1, 6 + (k + 3.5 + j + 1.75) * 0.7, 20, 6 + (k + 3.5 + j + 1.75) * 0.7
                PDF.PDFSetLineWidth = 0.025
                PDF.PDFDrawLine 1, 6.05 + (k + 3.5 + j + 1.75) * 0.7, 20, 6.05 + (k + 3.5 + j + 1.75) * 0.7
                
                
            
                PDF.PDFSetLineWidth = 0.03
                PDF.PDFDrawLine 1, 27.25, 20, 27.25
                PDF.PDFSetLineWidth = 0.025
                PDF.PDFDrawLine 1, 27.3, 20, 27.3
                PDF.PDFTextOut "Report Generated By JURA", 7.5, 27.75
                
                PDF.PDFEndPage
                rsMarks.Close
                rsRegNo.MoveNext
                prgReport.Value = prgReport.Value + iprgValue
                Me.Refresh
            Next
        PDF.PDFEndDoc
        rsRegNo.Close
        prgReport.Value = 100
End Sub


Private Sub CreateGradePDF()
    On Error Resume Next
    
    Dim sqlRegNo As String
    Dim sqlMarks As String
    Dim strRegNo As String
    Dim PDF As New clsPDF
    Dim i, j, k As Integer
    Dim rsRegNo As New ADODB.Recordset
    Dim rsMarks As New ADODB.Recordset
    Dim iprgValue As Integer
    
    rsRegNo.CursorLocation = adUseClient
    rsMarks.CursorLocation = adUseClient
    
    prgReport.Visible = True
    prgReport.Value = 0
    
    PDF.PDFTitle = "MarksSheet"
    PDF.PDFFileName = App.Path & "\Reports\Class Report" & "(" & cmbSem.Text & ")" & ".pdf"
    PDF.PDFLoadAfm = App.Path & "\Fonts"
    
    PDF.PDFView = True
    
    PDF.PDFSetBorder = BORDER_ALL
    PDF.PDFSetTopMargin = 1
    PDF.PDFSetBottomMargin = 1
    PDF.PDFSetRightMargin = 1
    PDF.PDFSetLeftMargin = 1
    PDF.PDFBeginDoc
        
        sqlRegNo = "select regno from studdetails where substr(regno, 6, 3) = '" & iDept & "' and substr(regno, 4, 2) = '" & Mid(iBatch, 3, 2) & "' and sec='" & strSec & "' and regno between '" & cmbStartRegNo.Text & "' and  '" & cmbEndRegNo.Text & "'  order by regno"
        rsRegNo.Open sqlRegNo, conn, adOpenDynamic, adLockOptimistic
        
        iprgValue = 100 / rsRegNo.RecordCount
        
        For i = 1 To rsRegNo.RecordCount / 2
                    
            PDF.PDFNewPage
        
                PDF.PDFDrawRectangle 1, 1, 19, 27
                
                PDF.PDFSetFont FONT_TIMES, 24, FONT_BOLD
                PDF.PDFSetTextColor = vbRed
                PDF.PDFTextOut "Francis Xavier Engineering College", 2.9, 2
                PDF.PDFSetFont FONT_TIMES, 18, FONT_BOLD
                PDF.PDFTextOut "Marks Sheet", 8, 3.25
                
                PDF.PDFSetLineWidth = 0.03
                PDF.PDFDrawLine 1, 3.65, 20, 3.65
                PDF.PDFSetLineWidth = 0.025
                PDF.PDFDrawLine 1, 3.7, 20, 3.7
                
                PDF.PDFSetTextColor = vbBlack
                PDF.PDFSetFont FONT_TIMES, 12, FONT_NORMAL
                PDF.PDFTextOut "Department:", 1, 4.25
                PDF.PDFTextOut cmbDept.Text, 3.35, 4.25
                PDF.PDFTextOut "Batch:", 10, 4.25
                PDF.PDFTextOut cmbBatch.Text, 11.25, 4.25
                PDF.PDFTextOut "Semester:", 14.5, 4.25
                PDF.PDFTextOut cmbSem.Text, 16.25, 4.25
                
                PDF.PDFSetLineWidth = 0.03
                PDF.PDFDrawLineHor 1, 4.5, 19
                PDF.PDFSetLineWidth = 0.02
                PDF.PDFDrawLine 1, 4.54, 20, 4.54
                
                
                
                PDF.PDFTextOut "Subject Code", 0.75, 5
                PDF.PDFTextOut "Subject Name", 3.5, 5
                PDF.PDFTextOut "Internals", 11.5, 5
                PDF.PDFTextOut "Grade", 14.35, 5
                PDF.PDFTextOut "Result", 17, 5
                
                
                PDF.PDFSetLineWidth = 0.03
                PDF.PDFDrawLine 1, 5.25, 20, 5.25
                PDF.PDFSetLineWidth = 0.025
                PDF.PDFDrawLine 1, 5.3, 20, 5.3
                
                'Student 1
                
                strRegNo = rsRegNo.Fields(0)
                sqlMarks = "select subjcode,internals,grade,result from studmarks where semno=" & iSem & " and batch=" & Mid(iBatch, 3, 2) & "and regno='" & strRegNo & "' order by subjcode"
                rsMarks.Open sqlMarks, conn, adOpenDynamic, adLockOptimistic
                
                PDF.PDFTextOut "Register No:", 1, 6
                PDF.PDFTextOut strRegNo, 3.5, 6
                PDF.PDFTextOut "Student Name:", 10, 6
                
                If Len(GetStudName(strRegNo)) > 20 Then
                    PDF.PDFSetFont FONT_TIMES, 10, FONT_NORMAL
                    PDF.PDFTextOut GetStudName(strRegNo), 13, 6
                Else
                    PDF.PDFSetFont FONT_TIMES, 12, FONT_NORMAL
                    PDF.PDFTextOut GetStudName(strRegNo), 13, 6
                End If
                                                                    
                PDF.PDFSetFont FONT_TIMES, 12, FONT_NORMAL
                
                For j = 1 To rsMarks.RecordCount
                
                    PDF.PDFTextOut rsMarks.Fields(0), 1, 6.25 + j * 0.7
                    
                    If Len(GetSubjName(rsMarks.Fields(0))) > 40 Then
                        PDF.PDFSetFont FONT_TIMES, 10, FONT_NORMAL
                        PDF.PDFTextOut GetSubjName(rsMarks.Fields(0)), 3.5, 6.25 + j * 0.7
                        PDF.PDFSetFont FONT_TIMES, 12, FONT_NORMAL
                    Else
                        PDF.PDFSetFont FONT_TIMES, 12, FONT_NORMAL
                        PDF.PDFTextOut GetSubjName(rsMarks.Fields(0)), 3.5, 6.25 + j * 0.7
                    End If
                                        
                    If rsMarks.Fields(3) = "WH2" Then
                        PDF.PDFSetTextColor = vbRed
                        PDF.PDFSetFont FONT_TIMES, 8, FONT_NORMAL
                        PDF.PDFTextOut "WH2", 17.5, 6.25 + j * 0.7
                        PDF.PDFSetFont FONT_TIMES, 12, FONT_NORMAL
                        PDF.PDFSetTextColor = vbBlack
                    ElseIf rsMarks.Fields(3) = "WH1" Then
                        PDF.PDFSetTextColor = vbRed
                        PDF.PDFSetFont FONT_TIMES, 8, FONT_NORMAL
                        PDF.PDFTextOut "WH1", 17.5, 6.25 + j * 0.7
                        PDF.PDFSetFont FONT_TIMES, 12, FONT_NORMAL
                        PDF.PDFSetTextColor = vbBlack
                    ElseIf rsMarks.Fields(3) = "WH" Then
                        PDF.PDFSetTextColor = vbRed
                        PDF.PDFSetFont FONT_TIMES, 10, FONT_NORMAL
                        PDF.PDFTextOut "WH", 17.5, 6.25 + j * 0.7
                        PDF.PDFSetFont FONT_TIMES, 12, FONT_NORMAL
                        PDF.PDFSetTextColor = vbBlack
                    ElseIf rsMarks.Fields(3) = "AB" Then
                        PDF.PDFTextOut rsMarks.Fields(1), 12, 6.25 + j * 0.7
                        PDF.PDFTextOut "AB", 14.75, 6.25 + j * 0.7
                        PDF.PDFSetTextColor = vbRed
                        PDF.PDFTextOut "AB", 17.5, 6.25 + j * 0.7
                        PDF.PDFSetTextColor = vbBlack
                    ElseIf rsMarks.Fields(3) = "SA" Then
                        PDF.PDFTextOut " ", 12, 6.25 + j * 0.7
                        PDF.PDFTextOut " ", 14.75, 6.25 + j * 0.7
                        PDF.PDFSetTextColor = vbRed
                        PDF.PDFTextOut "SA", 17.5, 6.25 + j * 0.7
                        PDF.PDFSetTextColor = vbBlack
                    Else
                        PDF.PDFTextOut rsMarks.Fields(1), 12, 6.25 + j * 0.7
                        
                        If rsMarks.Fields(2) = "U" Then
                            PDF.PDFTextOut rsMarks.Fields(2), 14.75, 6.25 + j * 0.7
                            PDF.PDFSetTextColor = vbRed
                            PDF.PDFTextOut "RA", 17.5, 6.25 + j * 0.7
                            PDF.PDFSetTextColor = vbBlack
                        ElseIf rsMarks.Fields(2) = "W" Then
                            PDF.PDFTextOut rsMarks.Fields(2), 14.75, 6.25 + j * 0.7
                            PDF.PDFSetTextColor = vbRed
                            PDF.PDFTextOut "RA", 17.5, 6.25 + j * 0.7
                            PDF.PDFSetTextColor = vbBlack
                        ElseIf rsMarks.Fields(2) = "I" Then
                            PDF.PDFTextOut rsMarks.Fields(2), 14.75, 6.25 + j * 0.7
                            PDF.PDFSetTextColor = vbRed
                            PDF.PDFTextOut "RA", 17.5, 6.25 + j * 0.7
                            PDF.PDFSetTextColor = vbBlack
                        Else
                            PDF.PDFTextOut rsMarks.Fields(2), 14.75, 6.25 + j * 0.7
                            PDF.PDFTextOut "P", 17.5, 6.25 + j * 0.7
                        End If
                    End If
                    rsMarks.MoveNext
                Next
                PDF.PDFSetTextColor = vbBlack
                PDF.PDFSetLineWidth = 0.03
                PDF.PDFDrawLine 1, 6.25 + (j + 0.75) * 0.7, 20, 6.25 + (j + 0.75) * 0.7
                
                                
                PDF.PDFTextOut "GPA:", 13.5, 6.25 + (j + 1.45) * 0.7
                PDF.PDFTextOut CalcGPA(strRegNo, iSem, iDept, iBatch), 15.15, 6.25 + (j + 1.45) * 0.7
                                
            
                PDF.PDFSetLineWidth = 0.03
                PDF.PDFDrawLine 1, 6.25 + (j + 1.75) * 0.7, 20, 6.25 + (j + 1.75) * 0.7
                PDF.PDFSetLineWidth = 0.025
                PDF.PDFDrawLine 1, 6.3 + (j + 1.75) * 0.7, 20, 6.3 + (j + 1.75) * 0.7
                
                rsMarks.Close
                prgReport.Value = prgReport.Value + iprgValue
                Me.Refresh
                
                'Student 2
                
                rsRegNo.MoveNext
                strRegNo = rsRegNo.Fields(0)
                sqlMarks = "select subjcode,internals,grade,result from studmarks where semno=" & iSem & " and batch=" & Mid(iBatch, 3, 2) & "and regno='" & strRegNo & "' order by subjcode"
                rsMarks.Open sqlMarks, conn, adOpenDynamic, adLockOptimistic
                k = j
                PDF.PDFTextOut "Register No:", 1, 6 + (k + 3.25) * 0.7
                PDF.PDFTextOut strRegNo, 3.5, 6 + (k + 3.25) * 0.7
                PDF.PDFTextOut "Student Name:", 10, 6 + (k + 3.25) * 0.7
                
                If Len(GetStudName(strRegNo)) > 20 Then
                    PDF.PDFSetFont FONT_TIMES, 10, FONT_NORMAL
                    PDF.PDFTextOut GetStudName(strRegNo), 13, 6 + (k + 3.25) * 0.7
                Else
                    PDF.PDFSetFont FONT_TIMES, 12, FONT_NORMAL
                    PDF.PDFTextOut GetStudName(strRegNo), 13, 6 + (k + 3.25) * 0.7
                End If
        
                PDF.PDFSetFont FONT_TIMES, 12, FONT_NORMAL
                
                For j = 1 To rsMarks.RecordCount
                
                    PDF.PDFTextOut rsMarks.Fields(0), 1, 6 + (k + 3.5 + j) * 0.7
                    
                    If Len(GetSubjName(rsMarks.Fields(0))) > 40 Then
                        PDF.PDFSetFont FONT_TIMES, 10, FONT_NORMAL
                        PDF.PDFTextOut GetSubjName(rsMarks.Fields(0)), 3.5, 6 + (k + 3.5 + j) * 0.7
                        PDF.PDFSetFont FONT_TIMES, 12, FONT_NORMAL
                    Else
                        PDF.PDFSetFont FONT_TIMES, 12, FONT_NORMAL
                        PDF.PDFTextOut GetSubjName(rsMarks.Fields(0)), 3.5, 6 + (k + 3.5 + j) * 0.7
                    End If
                    
                    If rsMarks.Fields(3) = "WH2" Then
                        PDF.PDFSetTextColor = vbRed
                        PDF.PDFSetFont FONT_TIMES, 8, FONT_NORMAL
                        PDF.PDFTextOut "WH2", 17.5, 6 + (k + 3.5 + j) * 0.7
                        PDF.PDFSetFont FONT_TIMES, 12, FONT_NORMAL
                        PDF.PDFSetTextColor = vbBlack
                    ElseIf rsMarks.Fields(3) = "WH1" Then
                        PDF.PDFSetTextColor = vbRed
                        PDF.PDFSetFont FONT_TIMES, 8, FONT_NORMAL
                        PDF.PDFTextOut "WH1", 17.5, 6 + (k + 3.5 + j) * 0.7
                        PDF.PDFSetFont FONT_TIMES, 12, FONT_NORMAL
                        PDF.PDFSetTextColor = vbBlack
                    ElseIf rsMarks.Fields(3) = "WH" Then
                        PDF.PDFSetTextColor = vbRed
                        PDF.PDFSetFont FONT_TIMES, 8, FONT_NORMAL
                        PDF.PDFTextOut "WH", 17.5, 6 + (k + 3.5 + j) * 0.7
                        PDF.PDFSetFont FONT_TIMES, 12, FONT_NORMAL
                        PDF.PDFSetTextColor = vbBlack
                    ElseIf rsMarks.Fields(3) = "AB" Then
                        PDF.PDFTextOut rsMarks.Fields(1), 12, 6 + (k + 3.5 + j) * 0.7
                        PDF.PDFTextOut "AB", 14.75, 6 + (k + 3.5 + j) * 0.7
                        PDF.PDFSetTextColor = vbRed
                        PDF.PDFTextOut "AB", 17.5, 6 + (k + 3.5 + j) * 0.7
                        PDF.PDFSetTextColor = vbBlack
                    ElseIf rsMarks.Fields(3) = "SA" Then
                        PDF.PDFTextOut " ", 12, 6 + (k + 3.5 + j) * 0.7
                        PDF.PDFTextOut " ", 14.75, 6 + (k + 3.5 + j) * 0.7
                        PDF.PDFSetTextColor = vbRed
                        PDF.PDFTextOut "SA", 17.5, 6 + (k + 3.5 + j) * 0.7
                        PDF.PDFSetTextColor = vbBlack
                    Else
                        PDF.PDFTextOut rsMarks.Fields(1), 12, 6 + (k + 3.5 + j) * 0.7
                        
                        If rsMarks.Fields(2) = "U" Then
                            PDF.PDFTextOut rsMarks.Fields(2), 14.75, 6 + (k + 3.5 + j) * 0.7
                            PDF.PDFSetTextColor = vbRed
                            PDF.PDFTextOut "RA", 17.5, 6 + (k + 3.5 + j) * 0.7
                            PDF.PDFSetTextColor = vbBlack
                        ElseIf rsMarks.Fields(2) = "W" Then
                            PDF.PDFTextOut rsMarks.Fields(2), 14.75, 6 + (k + 3.5 + j) * 0.7
                            PDF.PDFSetTextColor = vbRed
                            PDF.PDFTextOut "RA", 17.5, 6 + (k + 3.5 + j) * 0.7
                            PDF.PDFSetTextColor = vbBlack
                        ElseIf rsMarks.Fields(2) = "I" Then
                            PDF.PDFTextOut rsMarks.Fields(2), 14.75, 6 + (k + 3.5 + j) * 0.7
                            PDF.PDFSetTextColor = vbRed
                            PDF.PDFTextOut "RA", 17.5, 6 + (k + 3.5 + j) * 0.7
                            PDF.PDFSetTextColor = vbBlack
                        Else
                            PDF.PDFTextOut rsMarks.Fields(2), 14.75, 6 + (k + 3.5 + j) * 0.7
                            PDF.PDFTextOut "P", 17.5, 6 + (k + 3.5 + j) * 0.7
                        End If
                        
                    End If
                    rsMarks.MoveNext
                Next
                
                PDF.PDFSetTextColor = vbBlack
                PDF.PDFSetLineWidth = 0.03
                PDF.PDFDrawLine 1, 6 + (k + 3.5 + j + 0.75) * 0.7, 20, 6 + (k + 3.5 + j + 0.75) * 0.7
                                
                
                PDF.PDFTextOut "GPA:", 13.5, 6 + (k + 3.5 + j + 1.45) * 0.7
                PDF.PDFTextOut CalcGPA(strRegNo, iSem, iDept, iBatch), 15.15, 6 + (k + 3.5 + j + 1.45) * 0.7
                                                
            
                PDF.PDFSetLineWidth = 0.03
                PDF.PDFDrawLine 1, 6 + (k + 3.5 + j + 1.75) * 0.7, 20, 6 + (k + 3.5 + j + 1.75) * 0.7
                PDF.PDFSetLineWidth = 0.025
                PDF.PDFDrawLine 1, 6.05 + (k + 3.5 + j + 1.75) * 0.7, 20, 6.05 + (k + 3.5 + j + 1.75) * 0.7
                
                
            
                PDF.PDFSetLineWidth = 0.03
                PDF.PDFDrawLine 1, 27.25, 20, 27.25
                PDF.PDFSetLineWidth = 0.025
                PDF.PDFDrawLine 1, 27.3, 20, 27.3
                PDF.PDFTextOut "Report Generated By JURA", 7.5, 27.75
                
                PDF.PDFEndPage
                rsMarks.Close
                rsRegNo.MoveNext
                prgReport.Value = prgReport.Value + iprgValue
                Me.Refresh
            Next
        PDF.PDFEndDoc
        rsRegNo.Close
        prgReport.Value = 100
End Sub

