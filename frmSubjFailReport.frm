VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkusercontrolsxp.ocx"
Begin VB.Form frmSubjFailReport 
   BorderStyle     =   0  'None
   Caption         =   "Subject Report"
   ClientHeight    =   5040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8520
   Icon            =   "frmSubjFailReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   Begin vkUserContolsXP.vkLabel lblSubjName 
      Height          =   255
      Left            =   3240
      TabIndex        =   13
      Top             =   480
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   450
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
   End
   Begin vkUserContolsXP.vkFrame fSubjFailReport 
      Height          =   5055
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   8916
      BackColor1      =   16777215
      BackColor2      =   14737632
      Caption         =   "Subject Report"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TitleColor1     =   8388608
      TitleColor2     =   16744576
      TitleGradient   =   2
      TitleHeight     =   300
      BorderColor     =   12582912
      BorderWidth     =   2
      Begin JURA.StylerButton cmdPrint 
         Height          =   375
         Left            =   3000
         TabIndex        =   31
         Top             =   4440
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         Caption         =   "Print Report"
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
      End
      Begin VB.ComboBox cmbBatch 
         Height          =   315
         Left            =   1680
         TabIndex        =   30
         Top             =   840
         Width           =   1335
      End
      Begin vkUserContolsXP.vkLabel vkLabel1 
         Height          =   255
         Left            =   1680
         TabIndex        =   29
         Top             =   600
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
      Begin JURA.StylerButton cmdFirstLast 
         Height          =   375
         Left            =   5760
         TabIndex        =   28
         Top             =   4440
         Width           =   2535
         _ExtentX        =   7011
         _ExtentY        =   661
         Caption         =   "Best 5 & Last 5"
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
      End
      Begin JURA.StylerButton cmdOK 
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   4440
         Width           =   2535
         _ExtentX        =   14208
         _ExtentY        =   661
         Caption         =   "Go"
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
      End
      Begin vkUserContolsXP.vkTextBox txtCount 
         Height          =   300
         Left            =   2280
         TabIndex        =   26
         Top             =   2520
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   529
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
         Alignment       =   2
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkLabel lblCount 
         Height          =   255
         Left            =   1680
         TabIndex        =   25
         Top             =   2520
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Count:"
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
      Begin vkUserContolsXP.vkCheck cbFilter 
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   2520
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Filter Marks"
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
      Begin vkUserContolsXP.vkLabel lblExternalsEnd 
         Height          =   255
         Left            =   1680
         TabIndex        =   23
         Top             =   3600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Externals End:"
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
      Begin vkUserContolsXP.vkLabel lblExternalsStart 
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   3600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Externals Start:"
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
      Begin vkUserContolsXP.vkLabel lblInternalsEnd 
         Height          =   255
         Left            =   1680
         TabIndex        =   21
         Top             =   2880
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Internals End:"
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
      Begin vkUserContolsXP.vkLabel lblInternalsStart 
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   2880
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Internals Start:"
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
      Begin VB.ComboBox cmbExternalsEnd 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         TabIndex        =   19
         Top             =   3840
         Width           =   1335
      End
      Begin VB.ComboBox cmbExternalsStart 
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   18
         Top             =   3840
         Width           =   1335
      End
      Begin VB.ComboBox cmbInternalsEnd 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         TabIndex        =   17
         Top             =   3120
         Width           =   1335
      End
      Begin VB.ComboBox cmbInternalsStart 
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   16
         Top             =   3120
         Width           =   1335
      End
      Begin JURA.StylerButton cmdClose 
         Height          =   255
         Left            =   8040
         TabIndex        =   15
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
         Left            =   7800
         TabIndex        =   14
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
      Begin vkUserContolsXP.vkFrame frameStudName 
         Height          =   375
         Left            =   3360
         TabIndex        =   12
         Top             =   2040
         Visible         =   0   'False
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   661
         BackGradient    =   0
         Caption         =   "StudName"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TitleGradient   =   2
         TitleHeight     =   350
         BorderColor     =   16711680
      End
      Begin vkUserContolsXP.vkLabel lblDept 
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   600
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
         Left            =   240
         TabIndex        =   11
         Top             =   1200
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
      Begin vkUserContolsXP.vkLabel lblSubj 
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   1200
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
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
      Begin vkUserContolsXP.vkLabel lblStart 
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Marks Start:"
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
      Begin vkUserContolsXP.vkLabel lblEnd 
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Top             =   1800
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Marks End:"
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
      Begin VB.ComboBox cmbSubj 
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   1440
         Width           =   1335
      End
      Begin VB.ComboBox cmbStart 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   2040
         Width           =   1335
      End
      Begin VB.ComboBox cmbEnd 
         Height          =   315
         Left            =   1680
         TabIndex        =   5
         Top             =   2040
         Width           =   1335
      End
      Begin VB.ComboBox cmbDept 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox cmbSem 
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   1440
         Width           =   1335
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   3375
         Left            =   3240
         TabIndex        =   7
         Top             =   840
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   5953
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorBkg    =   -2147483634
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
         _Band(0).GridLinesBand=   0
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
End
Attribute VB_Name = "frmSubjFailReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SubjFailReportTop As Integer
Dim State As Integer
Dim strSubj As String
Private Sub cmbSubj_Load()
    On Error Resume Next
    cmbSubj.Clear
    MSHFlexGrid1.Clear
    Dim rs As New ADODB.Recordset
    sql = "select subjcode from subj where semno= " & cmbSem.Text & " and dept = " & Department(cmbDept) & " and batch=" & Mid(cmbBatch.Text, 3, 2) & ""
    rs.Open sql, conn, adOpenDynamic, adLockOptimistic
    Set cmbSubj.DataSource = rs
    Do While Not rs.EOF
        cmbSubj.AddItem (rs.Fields(0))
        rs.MoveNext
    Loop
    cmbSubj.Text = cmbSubj.List(0)
    lblSubjName.Caption = GetSubjName(cmbSubj.Text)
    strSubj = cmbSubj.Text
End Sub

Private Sub cbFilter_Click()
    If cbFilter.Value = vbChecked Then
        cmbStart.Enabled = False
        cmbEnd.Enabled = False
        cmbInternalsStart.Enabled = True
        cmbInternalsEnd.Enabled = True
        cmbExternalsStart.Enabled = True
        cmbExternalsEnd.Enabled = True
    Else
        cmbStart.Enabled = True
        cmbEnd.Enabled = True
        cmbInternalsStart.Enabled = False
        cmbInternalsEnd.Enabled = False
        cmbExternalsStart.Enabled = False
        cmbExternalsEnd.Enabled = False
    End If
End Sub



Private Sub cmbBatch_Change()
    iBatch = cmbBatch.Text
    If iBatch > 2007 Then
        lblStart.Visible = False
        lblEnd.Visible = False
        cmbStart.Visible = False
        cmbEnd.Visible = False
        cmbInternalsStart.Visible = False
        cmbInternalsEnd.Visible = False
        cmbExternalsStart.Visible = False
        cmbExternalsEnd.Visible = False
        lblInternalsStart.Visible = False
        lblInternalsEnd.Visible = False
        lblExternalsStart.Visible = False
        lblExternalsEnd.Visible = False
        cbFilter.Visible = False
    Else
        lblStart.Visible = True
        lblEnd.Visible = True
        cmbStart.Visible = True
        cmbEnd.Visible = True
        cmbInternalsStart.Visible = True
        cmbInternalsEnd.Visible = True
        cmbExternalsStart.Visible = True
        cmbExternalsEnd.Visible = True
        lblInternalsStart.Visible = True
        lblInternalsEnd.Visible = True
        lblExternalsStart.Visible = True
        lblExternalsEnd.Visible = True
        cbFilter.Visible = True
    End If
End Sub

Private Sub cmbBatch_Click()
    iBatch = cmbBatch.Text
    If iBatch > 2007 Then
        lblStart.Visible = False
        lblEnd.Visible = False
        cmbStart.Visible = False
        cmbEnd.Visible = False
        cmbInternalsStart.Visible = False
        cmbInternalsEnd.Visible = False
        cmbExternalsStart.Visible = False
        cmbExternalsEnd.Visible = False
        lblInternalsStart.Visible = False
        lblInternalsEnd.Visible = False
        lblExternalsStart.Visible = False
        lblExternalsEnd.Visible = False
        cbFilter.Visible = False
    Else
        lblStart.Visible = True
        lblEnd.Visible = True
        cmbStart.Visible = True
        cmbEnd.Visible = True
        cmbInternalsStart.Visible = True
        cmbInternalsEnd.Visible = True
        cmbExternalsStart.Visible = True
        cmbExternalsEnd.Visible = True
        lblInternalsStart.Visible = True
        lblInternalsEnd.Visible = True
        lblExternalsStart.Visible = True
        lblExternalsEnd.Visible = True
        cbFilter.Visible = True
    End If
End Sub

Private Sub cmbDept_Click()
   Call cmbSubj_Load
   iSem = cmbSem.Text
   strSubj = cmbSubj.Text
End Sub



Private Sub cmbSem_Change()
   Call cmbSubj_Load
   iSem = cmbSem.Text
   strSubj = cmbSubj.Text
End Sub

Private Sub cmbSem_Click()
   Call cmbSubj_Load
   iSem = cmbSem.Text
   strSubj = cmbSubj.Text
End Sub
Private Sub cmbSubj_Click()
    lblSubjName.Caption = GetSubjName(cmbSubj.Text)
    strSubj = cmbSubj.Text
End Sub

Private Sub CmdClose_Click()
    Unload Me
End Sub




Private Sub cmdMin_Click()
    Dim i As Long
    If State = 1 Then
        State = 0
        cmdMin.Caption = "+"
        For i = 4920 To 310 Step -150
            Me.Height = i
            DoEvents
        Next i
        Me.Height = 310
        fSubjFailReport.BorderWidth = 0
        fSubjFailReport.Height = 310
        Me.Top = 100
    Else
        State = 1
        cmdMin.Caption = "--"
        For i = 310 To 4920 Step 150
            Me.Height = i
            DoEvents
        Next i
        Me.Height = 4920
        fSubjFailReport.BorderWidth = 2
        fSubjFailReport.Height = 4920
        Me.Top = SubjFailReportTop
    End If
End Sub



Private Sub cmdOK_Click()
    Dim rs As New ADODB.Recordset
    Dim strSql As String
    If iBatch > 2007 Then
        If cbFilter.Value = vbChecked Then
            strSql = "select regno,internals,grade from studmarks where semno=" & iSem & " and substr(regno,6,3)='" & Department(cmbDept) & "' and batch=" & Mid(cmbBatch.Text, 3, 2) & " and internals between " & cmbInternalsStart.Text & " and " & cmbInternalsEnd.Text & " and grade between '" & cmbExternalsStart.Text & "' and '" & cmbExternalsEnd.Text & "' and subjcode = '" & strSubj & "' order by regno"
        Else
            strSql = "select regno,internals,grade from studmarks where semno=" & iSem & " and substr(regno,6,3)='" & Department(cmbDept) & "' and batch=" & Mid(cmbBatch.Text, 3, 2) & "  and subjcode = '" & strSubj & "' order by regno"
        End If
        rs.Open strSql, conn, adOpenDynamic, adLockOptimistic
        Set MSHFlexGrid1.DataSource = rs
        Call MSHFlexGrid1_Load
        txtCount.Text = MSHFlexGrid1.rows - 1
        rs.Close
    Else
        If cbFilter.Value = vbChecked Then
            strSql = "select regno,internals,externals,(internals+externals) as Sum from studmarks where semno=" & iSem & " and substr(regno,6,3)='" & Department(cmbDept) & "' and batch=" & Mid(cmbBatch.Text, 3, 2) & " and internals between " & cmbInternalsStart.Text & " and " & cmbInternalsEnd.Text & " and externals between " & cmbExternalsStart.Text & " and " & cmbExternalsEnd.Text & " and subjcode = '" & strSubj & "' order by regno"
        Else
            strSql = "select regno,internals,externals,(internals+externals) as Sum from studmarks where semno=" & iSem & " and substr(regno,6,3)='" & Department(cmbDept) & "' and batch=" & Mid(cmbBatch.Text, 3, 2) & " and (internals+externals) between " & cmbStart.Text & " and " & cmbEnd.Text & " and subjcode = '" & strSubj & "' order by regno"
        End If
        rs.Open strSql, conn, adOpenDynamic, adLockOptimistic
        Set MSHFlexGrid1.DataSource = rs
        Call MSHFlexGrid1_Load
        txtCount.Text = MSHFlexGrid1.rows - 1
        rs.Close
    End If
End Sub

Private Sub cmdFirstLast_Click()
    If iBatch > 2007 Then
    Else
        Call PDFReport
    End If
End Sub

Private Sub PDFReport()
    On Error Resume Next
    Dim l As Integer
    Dim rsBest As New ADODB.Recordset
    Dim strSqlBest As String
    rsBest.CursorLocation = adUseClient
    strSqlBest = "SELECT regno,internals,externals,(internals+externals) as Sum FROM (SELECT regno,semno,dept,batch,subjcode,internals,externals,DENSE_RANK() OVER (ORDER BY (internals+externals) DESC) best FROM studmarks where semno=" & iSem & " and batch=" & Mid(iBatch, 3, 2) & " and dept='" & Department(cmbDept) & "' and subjcode='" & cmbSubj.Text & "' and externals is not null ) WHERE  best <= 5"
    rsBest.Open strSqlBest, conn, adOpenDynamic, adLockOptimistic
    
    Dim rsLast As New ADODB.Recordset
    Dim strSqlLast As String
    rsLast.CursorLocation = adUseClient
    strSqlLast = "SELECT regno,internals,externals,(internals+externals) as Sum FROM (SELECT regno,semno,dept,batch,subjcode,internals,externals,DENSE_RANK() OVER (ORDER BY (internals+externals) ASC) last FROM studmarks where semno=" & iSem & " and batch=" & Mid(iBatch, 3, 2) & " and dept='" & Department(cmbDept) & "' and subjcode='" & cmbSubj.Text & "' and externals is not null ) WHERE  last <= 5"
    rsLast.Open strSqlLast, conn, adOpenDynamic, adLockOptimistic
    
    Dim PDF As New clsPDF                                                  'Calls the Pdf Class
    Dim i, j As Integer
    Dim strLeft As Double
    PDF.PDFTitle = "Subject Report"                                           'Pdf Title
    PDF.PDFFileName = App.Path & "\Reports\" & "Subject" & " Report " & cmbDept.Text & "(" & cmbSubj.Text & ")" & ".pdf"  'Saves The Pdf In the filename as Students Regno and Semester In the Folder Report at application Folder
    PDF.PDFLoadAfm = App.Path & "\Fonts"                                  'Font used in Pdf
    
    PDF.PDFView = True
    PDF.PDFCreator = "Kishore Jangid"
    PDF.PDFSetBorder = BORDER_ALL
    PDF.PDFSetTopMargin = 1
    PDF.PDFSetBottomMargin = 1
    PDF.PDFSetRightMargin = 1
    PDF.PDFSetLeftMargin = 1
    PDF.PDFBeginDoc
        
        PDF.PDFDrawRectangle 1, 1, 19, 27                                 'Page Border
        
        PDF.PDFSetTextColor = vbBlack
        PDF.PDFSetFont FONT_TIMES, 18, FONT_BOLD
        strCollegeName = "FRANCIS XAVIER ENGINEERING COLLEGE"
        l = PDF.PDFGetStringWidth(strCollegeName, "Times-Bold", 18)
        strLeft = (19 - (l * 2.54) / 72) / 2
        PDF.PDFTextOut strCollegeName, strLeft, 2
        
        PDF.PDFSetFont FONT_TIMES, 16, FONT_BOLD
        l = PDF.PDFGetStringWidth("Tirunelveli-627003", "Times-Bold", 16)
        strLeft = (19 - (l * 2.54) / 72) / 2
        PDF.PDFTextOut "Tirunelveli-627003", strLeft, 2.75
        
        PDF.PDFSetFont FONT_TIMES, 14, FONT_BOLD
        l = PDF.PDFGetStringWidth("RESULT ANALYSIS OF EVEN SEMESTER (2010-2011) - AU TIRUNELVELI", "Times-Bold", 14)
        strLeft = (19 - (l * 2.54) / 72) / 2
        PDF.PDFTextOut "RESULT ANALYSIS OF EVEN SEMESTER (2010-2011) - AU TIRUNELVELI", strLeft, 3.65
        
        l = PDF.PDFGetStringWidth("DEPARTMENT OF " & UCase(cmbDept.Text), "Times-Bold", 14)
        strLeft = (19 - (l * 2.54) / 72) / 2
        PDF.PDFTextOut "DEPARTMENT OF " & UCase(cmbDept.Text), strLeft, 4.4
        
        l = PDF.PDFGetStringWidth("SUBJECT WISE LIST OF TOPPERS & SLOW LEARNERS", "Times-Bold", 14)
        strLeft = (19 - (l * 2.54) / 72) / 2
        PDF.PDFTextOut "SUBJECT WISE LIST OF TOPPERS & SLOW LEARNERS", strLeft, 5.15
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 1, 5.35, 20, 5.35
        PDF.PDFSetLineWidth = 0.025
        PDF.PDFDrawLine 1, 5.4, 20, 5.4
        
        PDF.PDFSetTextColor = vbBlack
        PDF.PDFSetFont FONT_TIMES, 12, FONT_BOLD
        PDF.PDFTextOut "Subject:", 1, 6.25
        PDF.PDFSetFont FONT_TIMES, 12, FONT_NORMAL
        PDF.PDFTextOut GetSubjName(cmbSubj.Text) & " (" & cmbSubj.Text & ")", 3, 6.25
        PDF.PDFSetFont FONT_TIMES, 12, FONT_BOLD
        PDF.PDFTextOut "Semester:", 14.5, 6.25
        PDF.PDFSetFont FONT_TIMES, 12, FONT_NORMAL
        PDF.PDFTextOut cmbSem.Text, 16.75, 6.25
        PDF.PDFSetFont FONT_TIMES, 12, FONT_BOLD
        PDF.PDFTextOut "Faculty:", 1, 7
        
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 1, 7.5, 20, 7.5
        PDF.PDFSetLineWidth = 0.02
        PDF.PDFDrawLine 1, 7.55, 20, 7.55
        
        
        PDF.PDFSetFont FONT_TIMES, 12, FONT_BOLD
        PDF.PDFTextOut "Register No", 1, 8
        PDF.PDFTextOut "Student Name", 4, 8
        PDF.PDFTextOut "Internals", 12.25, 8
        PDF.PDFTextOut "Externals", 14.6, 8
        PDF.PDFTextOut "Marks", 17, 8
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 1, 8.15, 20, 8.15
        PDF.PDFSetLineWidth = 0.025
        PDF.PDFDrawLine 1, 8.2, 20, 8.2
        
        PDF.PDFSetFont FONT_TIMES, 12, FONT_BOLD
        PDF.PDFTextOut "Best 5:-", 0.75, 9
        
        PDF.PDFSetTextColor = vbBlack
        PDF.PDFSetFont FONT_TIMES, 10, FONT_NORMAL
        
        For i = 1 To rsBest.RecordCount
                PDF.PDFTextOut rsBest.Fields(0), 1, 9.1 + i * 0.6
                PDF.PDFTextOut GetStudName(rsBest.Fields(0)), 4, 9.1 + i * 0.6
                PDF.PDFTextOut rsBest.Fields(1), 13, 9.1 + i * 0.6
                PDF.PDFTextOut rsBest.Fields(2), 15.2, 9.1 + i * 0.6
                PDF.PDFTextOut rsBest.Fields(3), 17.25, 9.1 + i * 0.6
                rsBest.MoveNext
        Next
        
        PDF.PDFSetFont FONT_TIMES, 12, FONT_BOLD
        PDF.PDFTextOut "Last 5:-", 0.75, 16
        
        PDF.PDFSetFont FONT_TIMES, 10, FONT_NORMAL
        For i = 1 To rsLast.RecordCount
                PDF.PDFTextOut rsLast.Fields(0), 1, 16.1 + i * 0.6
                PDF.PDFTextOut GetStudName(rsLast.Fields(0)), 4, 16.1 + i * 0.6
                PDF.PDFTextOut rsLast.Fields(1), 13, 16.1 + i * 0.6
                PDF.PDFTextOut rsLast.Fields(2), 15.2, 16.1 + i * 0.6
                PDF.PDFTextOut rsLast.Fields(3), 17.25, 16.1 + i * 0.6
                rsLast.MoveNext
        Next
        
        PDF.PDFSetFont FONT_TIMES, 12, FONT_BOLD
        PDF.PDFTextOut "Staff Incharge", 2.5, 24.75
        PDF.PDFTextOut "HOD", 9, 24.75
        PDF.PDFTextOut "Principal", 15, 24.75
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 1, 27.25, 20, 27.25
        PDF.PDFSetLineWidth = 0.025
        PDF.PDFDrawLine 1, 27.3, 20, 27.3
        PDF.PDFSetFont FONT_TIMES, 10, FONT_NORMAL
        PDF.PDFTextOut "Report Generated By JURA", 7.5, 27.75
    PDF.PDFEndDoc
End Sub


Private Sub cmdPrint_Click()
    Dim rs As New ADODB.Recordset
    rs.CursorLocation = adUseClient
    Dim strSql As String
    If cbFilter.Value = vbChecked Then
        strSql = "select s1.regno as RegNo,s2.studname as StudName,s1.internals as Int,s1.externals as Ext,(s1.internals+s1.externals) as Sum from studmarks s1,studdetails s2 where s1.regno=s2.regno and s1.semno=" & iSem & " and substr(s1.regno,6,3)='" & Department(cmbDept) & "' and s1.batch=" & Mid(cmbBatch.Text, 3, 2) & " and s1.internals between " & cmbInternalsStart.Text & " and " & cmbInternalsEnd.Text & " and s1.externals between " & cmbExternalsStart.Text & " and " & cmbExternalsEnd.Text & " and s1.subjcode = '" & strSubj & "' order by s1.regno"
    Else
        strSql = "select s1.regno as RegNo,s2.studname as StudName,s1.internals as Int,s1.externals as Ext,(s1.internals+s1.externals) as Sum from studmarks s1,studdetails s2 where s1.regno=s2.regno and s1.semno=" & iSem & " and substr(s1.regno,6,3)='" & Department(cmbDept) & "' and s1.batch=" & Mid(cmbBatch.Text, 3, 2) & "and (s1.internals+s1.externals) between " & cmbStart.Text & " and " & cmbEnd.Text & " and subjcode = '" & strSubj & "' order by s1.regno"
    End If
    rs.Open strSql, conn, adOpenDynamic, adLockOptimistic
    Set SubjectReport.DataSource = rs
    
    With SubjectReport
        With .Sections("Section1")
            .Controls("Text1").DataField = "RegNo"
            .Controls("Text2").DataField = "StudName"
            .Controls("Text3").DataField = "Int"
            .Controls("Text4").DataField = "Ext"
            .Controls("Text5").DataField = "Sum"
        End With
        With .Sections("Section2")
            .Controls("lblDept").Caption = cmbDept.Text
            .Controls("lblSem").Caption = cmbSem.Text
            .Controls("lblBatch").Caption = cmbBatch.Text
            .Controls("lblIntStart").Caption = cmbInternalsStart.Text
            .Controls("lblIntEnds").Caption = cmbInternalsEnd.Text
            .Controls("lblExtStart").Caption = cmbExternalsStart.Text
            .Controls("lblExtEnds").Caption = cmbExternalsEnd.Text
            .Controls("lblSubjCode").Caption = cmbSubj.Text
            .Controls("lblSubjName").Caption = GetSubjName(cmbSubj.Text)
            .Controls("lblCount").Caption = txtCount.Text
        End With
        .Sections("Section3").Controls("lblDate").Caption = Date
        
    .LeftMargin = 100
    .RightMargin = 100
    .Show
    End With
    
End Sub

Private Sub Form_Load()
    CreateRoundRectFromWindow Me, 7, 7
    Me.Top = (mdiMain.Height - Me.Height) / 5
    Me.Left = (mdiMain.Width - Me.Width) / 2
    Me.BackColor = mdiMain.BackColor
    Call frmColor(frmSubjFailReport)
    
    'cmbDept.AddItem ("C.S.E.")
    'cmbDept.AddItem ("I.T.")
    'cmbDept.AddItem ("E.C.E.")
    'cmbDept.AddItem ("E.E.E.")
    'cmbDept.AddItem ("Mech.")
    'cmbDept.Text = cmbDept.List(0)
    Call cmbDept_Load(cmbDept)
    
    Call cmbBatch_Load(cmbBatch)
    Call MSHFlexGrid1_Load
    Call cmbSem_Load(cmbSem)
    Call cmbSubj_Load
    lblSubjName.Caption = GetSubjName(cmbSubj.Text)
        
    Dim i As Integer
    For i = 0 To 100 Step 5
        cmbStart.AddItem (i)
        cmbEnd.AddItem (i)
    Next i
    For i = 0 To 20 Step 2
        cmbInternalsStart.AddItem (i)
        cmbInternalsEnd.AddItem (i)
    Next
    For i = 0 To 80 Step 2
        cmbExternalsStart.AddItem (i)
        cmbExternalsEnd.AddItem (i)
    Next
    
    cmbStart.Text = cmbStart.List(0)
    cmbEnd.Text = cmbEnd.List(20)
    cmbInternalsStart.Text = cmbInternalsStart.List(0)
    cmbInternalsEnd.Text = cmbInternalsEnd.List(10)
    cmbExternalsStart.Text = cmbExternalsStart.List(0)
    cmbExternalsEnd.Text = cmbExternalsEnd.List(40)
    
    State = 1
    SubjFailReportTop = Me.Top
End Sub

Private Sub MSHFlexGrid1_Load()
    If iBatch > 2007 Then
        MSHFlexGrid1.RowHeightMin = 350
        MSHFlexGrid1.ColWidth(0) = 1950
        MSHFlexGrid1.ColWidth(1) = 1400
        MSHFlexGrid1.ColWidth(2) = 1400
        MSHFlexGrid1.ColAlignment(0) = flexAlignCenterCenter
        MSHFlexGrid1.ColAlignment(1) = flexAlignCenterCenter
        MSHFlexGrid1.ColAlignment(2) = flexAlignCenterCenter
        MSHFlexGrid1.ColAlignmentFixed(0) = flexAlignCenterCenter
        MSHFlexGrid1.ColAlignmentFixed(1) = flexAlignCenterCenter
        MSHFlexGrid1.ColAlignmentFixed(2) = flexAlignCenterCenter
        MSHFlexGrid1.TextMatrix(0, 0) = "Regno"
        MSHFlexGrid1.TextMatrix(0, 1) = "Internals"
        MSHFlexGrid1.TextMatrix(0, 2) = "Grade"
    Else
        MSHFlexGrid1.RowHeightMin = 350
        MSHFlexGrid1.ColWidth(0) = 1600
        MSHFlexGrid1.ColWidth(1) = 1050
        MSHFlexGrid1.ColWidth(2) = 1050
        MSHFlexGrid1.ColWidth(3) = 1050
        MSHFlexGrid1.ColAlignment(0) = flexAlignCenterCenter
        MSHFlexGrid1.ColAlignment(1) = flexAlignCenterCenter
        MSHFlexGrid1.ColAlignment(2) = flexAlignCenterCenter
        MSHFlexGrid1.ColAlignment(3) = flexAlignCenterCenter
        MSHFlexGrid1.ColAlignmentFixed(0) = flexAlignCenterCenter
        MSHFlexGrid1.ColAlignmentFixed(1) = flexAlignCenterCenter
        MSHFlexGrid1.ColAlignmentFixed(2) = flexAlignCenterCenter
        MSHFlexGrid1.ColAlignmentFixed(3) = flexAlignCenterCenter
        MSHFlexGrid1.TextMatrix(0, 0) = "Regno"
        MSHFlexGrid1.TextMatrix(0, 1) = "Internals"
        MSHFlexGrid1.TextMatrix(0, 2) = "Externals"
        MSHFlexGrid1.TextMatrix(0, 3) = "Marks"
    End If
End Sub
Private Sub frameStudName_MouseMove(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
    frameStudName.Visible = False
End Sub
Private Sub fSubjFailReport_MouseDown(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
Private Sub MSHFlexGrid1_DblClick()
    Dim rsStudName As New ADODB.Recordset
    Dim sql As String
    Dim StudRegno As String
    Dim Cursor As PointAPI
    Dim xPos As Long
    Dim yPos As Long

    GetCursorPos Cursor
    ScreenToClient Me.hWnd, Cursor
    
    xPos = Me.ScaleX(Cursor.x, vbPixels, vbTwips)
    yPos = Me.ScaleY(Cursor.y, vbPixels, vbTwips)
    
    With MSHFlexGrid1
        StudRegno = .TextMatrix(.MouseRow, 0)
    End With
    
    sql = "select studname from studdetails where regno = '" & StudRegno & "'"
    rsStudName.Open sql, conn, adOpenDynamic, adLockOptimistic
    
    frameStudName.Caption = "(" & StudRank(StudRegno, Val(cmbSem.Text), iDept, iBatch, strSec) & ")-" & rsStudName.Fields("studname")
    If xPos > Me.Width - 3200 Then
        frameStudName.Move xPos - 3500, yPos
    Else
        frameStudName.Move xPos + 500, yPos
    End If
    frameStudName.Width = 50
    frameStudName.Height = 10
    Me.Refresh
    frameStudName.Visible = True
    
    Do Until frameStudName.Width > 3000
        frameStudName.Width = frameStudName.Width + 40
    Loop
    Do Until frameStudName.Height > 375
        frameStudName.Height = frameStudName.Height + 1
    Loop
End Sub

