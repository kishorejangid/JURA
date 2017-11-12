Attribute VB_Name = "General"
Dim DBUser As String
Dim DBPassword As String
Private LoginTable As Boolean
Private StuddetailsTable As Boolean
Private StudmarksTable As Boolean
Private SubjTable As Boolean
Private DeptTable As Boolean
Sub Main()
    
    'Fetch Theme from Registry
    Color = GetSetting(App.CompanyName, "Theme", "frmColor", Orange)
    lColor = GetSetting(App.CompanyName, "Theme", "frmlColor", LOrange)
    
    'Fetch Application Settings from Registry
    flgSem = GetSetting(App.CompanyName, "Settings", "flgSem", "Odd")
    strExamMonth = GetSetting(App.CompanyName, "Settings", "ExamMonth")
    strExamYear = GetSetting(App.CompanyName, "Settings", "ExamYear", DateTime.Year(DateTime.Date))
    strCollegeName = GetSetting(App.CompanyName, "Settings", "CollegeName", "Francis Xavier Engineering College")
    flgRandomTheme = GetSetting(App.CompanyName, "Settings", "flgRandomTheme", False)
    
    Call OpenDatabase
    
    'Fetch DataBase Settings from Registry
    LoginTable = GetSetting(App.CompanyName, "DataBase", "LoginTable", False)
    StuddetailsTable = GetSetting(App.CompanyName, "DataBase", "StuddetailsTable", False)
    SubjTable = GetSetting(App.CompanyName, "DataBase", "SubjTable", False)
    DeptTable = GetSetting(App.CompanyName, "DataBase", "DeptTable", False)
    StudmarksTable = GetSetting(App.CompanyName, "DataBase", "StudmarksTable", False)
    
    If LoginTable = False Then
        Call CreateLoginTable
    End If
    If StuddetailsTable = False Then
        Call CreateStuddetailsTable
    End If
    If SubjTable = False Then
        Call CreateSubjTable
    End If
    If DeptTable = False Then
        Call CreateDeptTable
    End If
    If StudmarksTable = False Then
        Call CreateStudmarksTable
    End If
     
    If Dir(App.Path & "\Reports", vbDirectory) = vbNullString Then
        MkDir App.Path & "\Reports"
    End If
    If Dir(App.Path & "\Images", vbDirectory) = vbNullString Then
        MkDir App.Path & "\Images"
    End If
    If Dir(App.Path & "\Syllabus", vbDirectory) = vbNullString Then
        MkDir App.Path & "\Syllabus"
    End If
    iTop = 50
End Sub
Public Sub OpenDatabase()
    On Error Resume Next
    Set conn = New ADODB.Connection
    DBUser = GetSetting(App.CompanyName, "DataBase", "DBUser")
    DBPassword = GetSetting(App.CompanyName, "DataBase", "DBPassword")
    With conn
        .Provider = "MSDAORA.1"
        .ConnectionString = "User ID=' " & DBUser & " ';Persist Security Info=False;User ID=' " & DBUser & " ';Password = ' " & DBPassword & " '"
        .Open
    End With
    If Err.Number <> 0 Then
        If Err.Number = -2147217843 Then
            MsgBox "Error Connecting Oracle:" & vbCrLf & vbCrLf & "        * Please Check Whether Oracle 10g Is Installed On Your System" & vbCrLf & "           If Installed Provide The DBUserName And Password" & vbCrLf & vbCrLf & "Error Number: " & Err.Number & vbCrLf & vbCrLf & "************************************** JURA® *************************************"
            frmDBPass.Show
        Else
            MsgBox Error & vbCrLf & "Error Number: " & Err.Number
        End If
    Else
        'Hide Login form
        'frmSplash.Show
        mdiMain.Show
        mdiMain.Enabled = True
        admCheck = "Administrator"
    End If
End Sub

Private Sub CreateLoginTable()
    On Error Resume Next
    Dim sqlLogin As String
    Dim CreateLogin As New ADODB.Recordset
    CreateLogin.CursorLocation = adUseClient
    sqlLogin = "CREATE TABLE LOGIN(LOGINTYPE VARCHAR2(20) NOT NULL ENABLE,LOGINID VARCHAR2(20) NOT NULL ENABLE,LOGINPASSWORD VARCHAR2(20) NOT NULL ENABLE,CONSTRAINT LOGIN_CON UNIQUE (LOGINID) ENABLE)"
    CreateLogin.Open sqlLogin, conn, adOpenDynamic, adLockOptimistic, -1
    
    If Err.Number <> 0 Then
        If Err.Number = -2147217900 Then
            SaveSetting App.CompanyName, "DataBase", "LoginTable", True
        Else
            MsgBox "Error:" & Err.Description & vbCrLf & "Error Number:" & Err.Number
        End If
    Else
        SaveSetting App.CompanyName, "DataBase", "LoginTable", True
        MsgBox "Table Login Created Sucessfully"
    End If
End Sub
Private Sub CreateStuddetailsTable()
    On Error Resume Next
    Dim sql As String
    Dim Create As New ADODB.Recordset
    Create.CursorLocation = adUseClient
    sql = "CREATE TABLE  STUDDETAILS(REGNO VARCHAR2(11) NOT NULL ENABLE,STUDNAME VARCHAR2(50) NOT NULL ENABLE,SEC VARCHAR2(1) NOT NULL ENABLE,DOB DATE,GENDER VARCHAR2(6),FATHER VARCHAR2(50),MOTHER VARCHAR2(50),OCCUPATION VARCHAR2(50),ADDRESS VARCHAR2(100),CITY VARCHAR2(25),PINCODE NUMBER,STATE VARCHAR2(25),EMAIL VARCHAR2(50),LANDLINE NUMBER,MOBILE NUMBER,IMAGE VARCHAR2(100),CONSTRAINT STUDDETAILS_CON PRIMARY KEY (REGNO) ENABLE)"
    Create.Open sql, conn, adOpenDynamic, adLockOptimistic, -1
    
    If Err.Number <> 0 Then
        If Err.Number = -2147217900 Then
            SaveSetting App.CompanyName, "DataBase", "StuddetailsTable", True
        Else
            MsgBox "Error:" & Err.Description & vbCrLf & "Error Number:" & Err.Number
        End If
    Else
        SaveSetting App.CompanyName, "DataBase", "StuddetailsTable", True
        MsgBox "Table Studdetails Created Sucessfully"
    End If
End Sub
Private Sub CreateStudmarksTable()
    On Error Resume Next
    Dim sql As String
    Dim Create As New ADODB.Recordset
    Create.CursorLocation = adUseClient
    sql = "CREATE TABLE STUDMARKS(REGNO VARCHAR2(11) NOT NULL ENABLE,SEMNO NUMBER(2,0),DEPT NUMBER,BATCH NUMBER,SUBJCODE VARCHAR2(10),INTERNALS NUMBER(2,0),EXTERNALS NUMBER(2,0),CONSTRAINT STUDMARKS_CON FOREIGN KEY (REGNO) REFERENCES  STUDDETAILS (REGNO) ENABLE,CONSTRAINT STUDMARKS_CONSUBJ FOREIGN KEY (SUBJCODE,SEMNO,DEPT,BATCH) REFERENCES  SUBJ (SUBJCODE,SEMNO,DEPT,BATCH) ENABLE )"
    Create.Open sql, conn, adOpenDynamic, adLockOptimistic, -1
    
    If Err.Number <> 0 Then
        If Err.Number = -2147217900 Then
            SaveSetting App.CompanyName, "DataBase", "StudmarksTable", True
        Else
            MsgBox "Error:" & Err.Description & vbCrLf & "Error Number:" & Err.Number
        End If
    Else
        SaveSetting App.CompanyName, "DataBase", "StudmarksTable", True
        MsgBox "Table Studmarks Created Sucessfully"
    End If
End Sub
Private Sub CreateSubjTable()
    On Error Resume Next
    Dim sql As String
    Dim Create As New ADODB.Recordset
    Create.CursorLocation = adUseClient
    sql = "CREATE TABLE  SUBJ(SUBJCODE VARCHAR2(10),SUBJNAME VARCHAR2(50),SEMNO NUMBER,DEPT NUMBER,BATCH NUMBER,CREDIT NUMBER,CONSTRAINT SUBJ_CON PRIMARY KEY (SUBJCODE,SEMNO,DEPT,BATCH) ENABLE)"
    Create.Open sql, conn, adOpenDynamic, adLockOptimistic, -1
    
    If Err.Number <> 0 Then
        If Err.Number = -2147217900 Then
            SaveSetting App.CompanyName, "DataBase", "SubjTable", True
        Else
            MsgBox "Error:" & Err.Description & vbCrLf & "Error Number:" & Err.Number
        End If
    Else
        SaveSetting App.CompanyName, "DataBase", "SubjTable", True
        MsgBox "Table Subj Created Sucessfully"
    End If
End Sub
Private Sub CreateDeptTable()
    On Error Resume Next
    Dim sql As String
    Dim Create As New ADODB.Recordset
    Create.CursorLocation = adUseClient
    sql = "CREATE TABLE DEPT(DEPTCODE NUMBER NOT NULL ENABLE,DEPTNAME VARCHAR2(100) NOT NULL ENABLE,DEPTSHORT VARCHAR2(10) NOT NULL ENABLE,CONSTRAINT DEPT_CON UNIQUE (DEPTNAME) ENABLE,PRIMARY KEY (DEPTCODE) ENABLE )"
    Create.Open sql, conn, adOpenDynamic, adLockOptimistic, -1
    
    If Err.Number <> 0 Then
        If Err.Number = -2147217900 Then
            SaveSetting App.CompanyName, "DataBase", "DeptTable", True
        Else
            MsgBox "Error:" & Err.Description & vbCrLf & "Error Number:" & Err.Number
        End If
    Else
        SaveSetting App.CompanyName, "DataBase", "DeptTable", True
        MsgBox "Table Dept Created Sucessfully"
    End If
End Sub

Public Sub frmColor(frm As Form)
    On Error Resume Next
    Dim ctrl As Control
    For Each ctrl In frm.Controls
        If TypeOf ctrl Is vkTextBox Then
            ctrl.BorderColor = Color
            ctrl.LegendBackColor1 = lColor
            ctrl.LegendBackColor2 = Color
        ElseIf TypeOf ctrl Is vkLabel Then
            ctrl.BorderColor = Color
            ctrl.ForeColor = Color
        ElseIf TypeOf ctrl Is vkFrame Then
            ctrl.BorderColor = Color
            ctrl.TitleColor1 = Color
            ctrl.TitleColor2 = lColor
        ElseIf TypeOf ctrl Is vkCommand Then
            ctrl.BorderColor = Color
            ctrl.ForeColor = Color
        ElseIf TypeOf ctrl Is vkBar Then
            ctrl.BorderColor = Color
            ctrl.LeftColor = Color
            ctrl.RightColor = Color
        ElseIf TypeOf ctrl Is MSHFlexGrid Then
            ctrl.BackColorFixed = lColor
            ctrl.BackColorSel = lColor
            ctrl.GridColor = Color
            ctrl.ForeColorFixed = Color
        End If
    Next
End Sub
Public Sub cmbRegNo_Load(ComboBox As ComboBox)
    On Error GoTo ErrHnd
    Dim rs As New ADODB.Recordset
    Dim qr As String
    ComboBox.Clear
    qr = "select regno from studdetails where substr(regno, 6, 3) =  '" & iDept & "' and substr(regno,4,2)='" & Mid(iBatch, 3, 2) & "' and sec='" & strSec & "' order by regno"
    rs.Open qr, conn, adOpenDynamic, adLockOptimistic, adCmdText
    Set ComboBox.DataSource = rs
    Do While Not rs.EOF
        ComboBox.AddItem (rs.Fields("regno"))
        rs.MoveNext
    Loop
    ComboBox.Text = ComboBox.List(0)
    ComboBox.Refresh
    rs.Close
    Exit Sub
ErrHnd:
    MsgBox Error & vbCrLf & "Error Number: " & Err.Number, vbCritical, "Error"
End Sub
Public Sub cmbSec_Load(ComboBox As ComboBox)
    On Error GoTo ErrHnd
    Dim rs As New ADODB.Recordset
    Dim qr As String
    ComboBox.Clear
    ComboBox.AddItem ("A")
    ComboBox.AddItem ("B")
    ComboBox.AddItem ("C")
    ComboBox.AddItem ("D")
    ComboBox.AddItem ("E")
    ComboBox.AddItem ("F")
    ComboBox.AddItem ("G")
    ComboBox.AddItem ("H")
    ComboBox.Text = ComboBox.List(0)
    strSec = ComboBox.Text
    Exit Sub
ErrHnd:
    MsgBox Error & vbCrLf & "Error Number: " & Err.Number, vbCritical, "Error"
End Sub
Public Function Department(ComboBox As ComboBox) As Integer
    On Error Resume Next
    Dim rsDeptCode As New ADODB.Recordset
    Dim strSql As String
    strSql = "select deptcode from dept where deptname='" & UCase(ComboBox.Text) & "'"
    rsDeptCode.Open strSql, conn, adOpenDynamic, adLockOptimistic, -1
    Department = rsDeptCode.Fields(0)
End Function
Public Sub cmbDept_Load(ComboBox As ComboBox)
    On Error GoTo errHan
    ComboBox.Clear
    Dim rsDept As New ADODB.Recordset
    Dim strSql As String
    strSql = "select deptname from dept order by deptshort"
    rsDept.Open strSql, conn, adOpenDynamic, adLockOptimistic, -1
    rsDept.MoveFirst
    ComboBox.FontSize = 8
    While Not rsDept.EOF
        ComboBox.AddItem rsDept.Fields(0)
        rsDept.MoveNext
    Wend
    ComboBox.Text = ComboBox.List(0)
    iDept = Department(ComboBox)
    Exit Sub
errHan:
    If Err.Number = 3021 Then
        MsgBox "No Departments are Created yet.", vbInformation, "feedback"
    End If
End Sub
Public Sub cmbBatch_Load(ComboBox As ComboBox)
    For i = 2006 To 2010
        ComboBox.AddItem (i)
    Next
    ComboBox.Text = ComboBox.List(0)
    iBatch = Val(Trim(ComboBox.Text))
End Sub
Public Sub cmbSem_Load(ComboBox As ComboBox)
    Dim S As Integer
    For S = 1 To 8
        ComboBox.AddItem (S)
    Next S
    ComboBox.Text = ComboBox.List(0)
    iSem = Val(ComboBox.Text)
End Sub
Public Function Sem2Word(iSem As Integer) As String
    Select Case iSem
    Case 1:
        Sem2Word = "FIRST"
    Case 2:
        Sem2Word = "SECOND"
    Case 3:
        Sem2Word = "THIRD"
    Case 4:
        Sem2Word = "FOURTH"
    Case 5:
        Sem2Word = "FIFTH"
    Case 6:
        Sem2Word = "SIXTH"
    Case 7:
        Sem2Word = "SEVENTH"
    Case 8:
        Sem2Word = "EIGHT"
    End Select
End Function
Public Function StudRank(sRegNo As String, iSem As Integer, iDept As Integer, iBatch As Integer, strSec As String) As Integer
    Dim rsStudCount As New ADODB.Recordset
    Dim rsTotal As New ADODB.Recordset
    Dim sqlStudCount As String
    Dim sqlTotal As String
    Dim iStudCount As Integer
    Dim i As Integer
    Dim arr(500) As String
    Dim iRank As Integer
    sqlStudCount = "select count(distinct s1.regno) from studmarks s1,studdetails s2 where s1.regno=s2.regno and s1.semno=" & iSem & " and s1.dept=" & iDept & " and s1.batch= " & Mid(iBatch, 3, 2) & " and s2.sec='" & strSec & "' "
    If iBatch > 2007 Then
        sqlTotal = "SELECT s1.regno,round(sum(s1.value*s2.credit)/(select sum(credit) FROM subj WHERE batch=" & Mid(iBatch, 3, 2) & " AND semno=" & iSem & " AND dept=" & iDept & "),2) AS GPA FROM studmarks s1,subj s2 WHERE s1.batch=" & Mid(iBatch, 3, 2) & " AND s1.semno=" & iSem & " AND s1.subjcode=s2.subjcode and s1.semno=s2.semno and s1.dept=s2.dept and s1.batch=s2.batch and s1.regno='" & sRegNo & "'  GROUP BY s1.regno"
    Else
        sqlTotal = "select s1.regno,sum(s1.internals+s1.externals) as Total from studmarks s1,studdetails s2 where s1.regno=s2.regno and s2.sec='" & strSec & "' and s1.semno= " & iSem & " and s1.dept=" & iDept & " and s1.batch= " & Mid(iBatch, 3, 2) & " group by s1.regno order by Total desc"
    End If
    rsStudCount.Open sqlStudCount, conn, adOpenDynamic, adLockOptimistic
    rsTotal.Open sqlTotal, conn, adOpenDynamic, adLockOptimistic
    iStudCount = rsStudCount.Fields(0)
    For i = 0 To iStudCount
        If Not rsTotal.EOF Then
            arr(i) = rsTotal.Fields(0)
            rsTotal.MoveNext
        End If
    Next
    For i = 0 To iStudCount
        If arr(i) = sRegNo Then
            StudRank = i + 1
        End If
    Next
End Function
Public Function ArrearCount(Sem As Integer, Batch As Integer, Dept As Integer, Condition As String) As Integer
    Dim rs As New ADODB.Recordset
    Dim sql As String
    If iBatch > 2007 Then
        sql = "select count(regno) from (select regno,count(subjcode) as count from studmarks where semno= " & Sem & " and batch = " & Mid(Batch, 3, 2) & " and grade in ('U','I','W','RA') and dept= " & Dept & " group by regno) where count  " & Condition & ""
    Else
        sql = "select count(regno) from (select regno,count(subjcode) as count from studmarks where semno= " & Sem & " and batch = " & Mid(Batch, 3, 2) & " and (internals+externals)<50 and dept= '" & Dept & "' group by regno) where count  " & Condition & ""
    End If
    rs.Open sql, conn, adOpenDynamic, adLockOptimistic
    ArrearCount = rs.Fields(0)
End Function
Public Function GetSubjCount(Sem As Integer, Dept As Integer, Batch As Integer) As Integer
    Dim rs As New ADODB.Recordset
    sql = "select count(subjcode) from subj where semno= " & Sem & " and dept = " & Dept & " and batch= " & Mid(Batch, 3, 2) & ""
    rs.Open sql, conn, adOpenDynamic, adLockOptimistic
    GetSubjCount = rs.Fields(0)
End Function
Public Function GetSubjName(strSubjCode As String) As String
    On Error Resume Next
    Dim rs As New ADODB.Recordset
    sql = "select subjname from subj where subjcode= '" & strSubjCode & "'"
    rs.Open sql, conn, adOpenDynamic, adLockOptimistic
    GetSubjName = rs.Fields(0)
End Function
Public Function GetSubjSem(strSubjCode As String, Dept As Integer, Batch As String) As String
    On Error Resume Next
    Dim rs As New ADODB.Recordset
    sql = "select semno from subj where subjcode= '" & strSubjCode & "' and dept=" & Dept & " and batch = " & Batch & ""
    rs.Open sql, conn, adOpenDynamic, adLockOptimistic
    GetSubjSem = rs.Fields(0)
    If Err.Number <> 0 Then
        MsgBox (Err.Number & vbCrLf & Err.Description & vbCrLf & "SubjCode: " & strSubjCode)
    End If
End Function
Public Function GetStudName(strRegNo As String) As String
    On Error Resume Next
    Dim rs As New ADODB.Recordset
    sql = "select studname from studdetails where regno= '" & strRegNo & "'"
    rs.Open sql, conn, adOpenDynamic, adLockOptimistic
    GetStudName = rs.Fields(0)
End Function

Public Function GetCount(Sem As Integer, Dept As Integer, Batch As Integer, Subj As String, ConditionStart As String, ConditionEnd As String) As Integer
    On Error Resume Next
    Dim rs As New ADODB.Recordset
    Dim sql As String
    sql = "Select Count(s1.regno) from studmarks s1,studdetails s2 where s1.regno=s2.regno and s2.sec='" & strSec & "' and s1.semno=" & Sem & " and s1.dept=" & Dept & " and s1.batch=" & Mid(Batch, 3, 2) & " and s1.subjcode='" & Subj & "' and (s1.internals+s1.externals) between " & ConditionStart & " and " & ConditionEnd & ""
    rs.Open sql, conn, adOpenDynamic, adLockOptimistic
    GetCount = rs.Fields(0)
End Function
'For Grade System for counting no of passed students in a subject
Public Function GetGradePassedCount(Sem As Integer, Dept As Integer, Batch As Integer, Subj As String, Condition As String) As Integer
    On Error Resume Next
    Dim rs As New ADODB.Recordset
    Dim sql As String
    sql = "Select Count(regno) from studmarks where semno=" & Sem & " and dept=" & Dept & " and batch=" & Mid(Batch, 3, 2) & " and subjcode='" & Subj & "' and grade in " & Condition & " "
    rs.Open sql, conn, adOpenDynamic, adLockOptimistic
    GetGradePassedCount = rs.Fields(0)
End Function
Public Function GetMaxMarks(Sem As Integer, Dept As Integer, Batch As Integer, Subj As String) As Integer
    On Error Resume Next
    Dim rs As New ADODB.Recordset
    Dim sql As String
    sql = "Select max(internals+externals) from studmarks where semno=" & Sem & " and dept=" & Dept & " and batch=" & Mid(Batch, 3, 2) & " and subjcode='" & Subj & "'"
    rs.Open sql, conn, adOpenDynamic, adLockOptimistic
    GetMaxMarks = rs.Fields(0)
End Function
Public Function GetMinMarks(Sem As Integer, Dept As Integer, Batch As Integer, Subj As String) As Integer
    On Error Resume Next
    Dim rs As New ADODB.Recordset
    Dim sql As String
    sql = "Select min(internals+externals) from studmarks where semno=" & Sem & " and dept=" & Dept & " and batch=" & Mid(Batch, 3, 2) & " and subjcode='" & Subj & "'"
    rs.Open sql, conn, adOpenDynamic, adLockOptimistic
    GetMinMarks = rs.Fields(0)
End Function
Public Function GetAvgMarks(Sem As Integer, Dept As Integer, Batch As Integer, Subj As String) As Double
    On Error Resume Next
    Dim rs As New ADODB.Recordset
    Dim sql As String
    sql = "Select Round(Avg(internals+externals),2) from studmarks where semno=" & Sem & " and dept=" & Dept & " and batch=" & Mid(Batch, 3, 2) & " and subjcode='" & Subj & "'"
    rs.Open sql, conn, adOpenDynamic, adLockOptimistic
    GetAvgMarks = rs.Fields(0)
End Function
Public Function GetNoOfStudAppeared(Sem As Integer, Dept As Integer, Batch As Integer, Subj As String) As Integer
    Dim rs As New ADODB.Recordset
    Dim strSql As String
    If iBatch > 2007 Then
        strSql = "Select Count(s1.regno) From studmarks s1,studdetails s2 Where s1.regno=s2.regno and s2.sec='" & strSec & "' and s1.semno=" & Sem & " and s1.dept=" & Dept & " and s1.batch=" & Mid(Batch, 3, 2) & " and s1.subjcode='" & Subj & "' and s1.grade in ('S','A','B','C','D','E','U','I','W')"
    Else
        strSql = "Select Count(s1.regno) From studmarks s1,studdetails s2 Where s1.regno=s2.regno and s2.sec='" & strSec & "' and s1.semno=" & Sem & " and s1.dept=" & Dept & " and s1.batch=" & Mid(Batch, 3, 2) & " and s1.subjcode='" & Subj & "' and s1.externals is not null"
    End If
    rs.Open strSql, conn, adOpenDynamic, adLockOptimistic
    GetNoOfStudAppeared = rs.Fields(0)
End Function
Public Function PassPercentage(sSubjcode As String, iSem As Integer, iDept As Integer, iBatch As Integer) As String
    If iBatch > 2007 Then
        PassPercentage = Round(((GetCount(iSem, CInt(iDept), iBatch, sSubjcode, 50, 100) / GetNoOfStudAppeared(iSem, CInt(iDept), iBatch, sSubjcode)) * 100), 2)
    Else
        PassPercentage = Round(((GetCount(iSem, CInt(iDept), iBatch, sSubjcode, 50, 100) / GetNoOfStudAppeared(iSem, CInt(iDept), iBatch, sSubjcode)) * 100), 2)
    End If
End Function
'Related to Grade System
Public Function GetGradeCount(Sem As Integer, Dept As Integer, Batch As Integer, Subj As String, Grade As String) As Integer
    On Error Resume Next
    Dim rs As New ADODB.Recordset
    Dim sql As String
    sql = "Select Count(regno) from studmarks where semno=" & Sem & " and dept=" & Dept & " and batch=" & Mid(Batch, 3, 2) & " and subjcode='" & Subj & "' and grade = " & Grade & ""
    rs.Open sql, conn, adOpenDynamic, adLockOptimistic
    GetGradeCount = rs.Fields(0)
End Function
'Gets the sum of credit in the given semester for calculating gpa
Public Function GetSumCredit(Sem As String, Dept As Integer, Batch As String) As Integer
    Dim rs As New ADODB.Recordset
    sql = "select sum(credit) from subj where semno= " & Sem & " and dept = " & Dept & " and batch= " & Mid(Batch, 3, 2) & ""
    rs.Open sql, conn, adOpenDynamic, adLockOptimistic
    GetSumCredit = rs.Fields(0)
End Function
'Get Subject Credit
Public Function GetSubjCredit(strSubjCode As String, iDept As Integer, iBatch As Integer) As Integer
    On Error Resume Next
    Dim rs As New ADODB.Recordset
    sql = "select credit from subj where subjcode= '" & strSubjCode & "' and dept=" & iDept & " and batch= " & Mid(iBatch, 3, 2) & " "
    rs.Open sql, conn, adOpenDynamic, adLockOptimistic
    GetSubjCredit = rs.Fields(0)
End Function
'Calculate GPA of a student
Public Function CalcGPA(sRegNo As String, iSem As Integer, iDept As Integer, iBatch As Integer) As Double
    On Error Resume Next
    Dim rs As New ADODB.Recordset
    Dim strSql As String
    strSql = "SELECT s1.regno,round(sum(s1.value*s2.credit)/(select sum(credit) FROM subj WHERE batch=" & Mid(iBatch, 3, 2) & " AND semno=" & iSem & " AND dept=" & iDept & "),2) AS GPA FROM studmarks s1,subj s2 WHERE s1.batch=" & Mid(iBatch, 3, 2) & " AND s1.semno=" & iSem & " AND s1.subjcode=s2.subjcode and s1.semno=s2.semno and s1.dept=s2.dept and s1.batch=s2.batch and s1.regno='" & sRegNo & "'  GROUP BY s1.regno"
    rs.Open strSql, conn, adOpenDynamic, adLockOptimistic
    CalcGPA = rs.Fields(1)
End Function
Public Function JangidFormat(str As String) As String
    Dim i As Integer
    Dim newstr As String
    newstr = UCase(Mid$(str, 1, 1))
    For i = 2 To Len(str)
        newstr = newstr & Mid$(str, i, 1)
        If Mid$(str, i, 1) = " " Or Mid$(str, i, 1) = "," Then
            newstr = newstr & UCase(Mid$(str, i + 1, 1))
            i = i + 1
        End If
    Next
    JangidFormat = newstr
End Function
Public Sub JuraMsgBox(strMsg As String)
    frmMsgBox.lblMsg.Caption = strMsg
    If Len(strMsg) > 42 Then
        frmMsgBox.lblMsg.Width = (Len(strMsg) * 75) - 250
        frmMsgBox.Width = Len(strMsg) * 75
        frmMsgBox.fMsgBox.Width = Len(strMsg) * 75
        frmMsgBox.cmdOK.Width = (Len(strMsg) * 75) - 250
    End If
    frmMsgBox.Refresh
    frmMsgBox.Show
End Sub
Public Function getGradeValue(strGrade As String) As Integer
    Dim iValue As Integer
    Select Case strGrade
        Case "S"
            iValue = 10
        Case "A"
            iValue = 9
        Case "B"
            iValue = 8
        Case "C"
            iValue = 7
        Case "D"
            iValue = 6
        Case "E"
            iValue = 5
        Case "U"
            iValue = 0
        Case "I"
            iValue = 0
        Case "W"
            iValue = 0
    End Select
    getGradeValue = iValue
End Function



  
  

