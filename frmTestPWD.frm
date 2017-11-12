VERSION 5.00
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form frmTestPWD 
   BorderStyle     =   0  'None
   ClientHeight    =   3615
   ClientLeft      =   1515
   ClientTop       =   1500
   ClientWidth     =   4815
   Icon            =   "frmTestPWD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   Begin vkUserContolsXP.vkCommand cmdCancel 
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   2400
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Caption         =   "Cancel"
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
   Begin vkUserContolsXP.vkCommand CmdOK 
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Caption         =   "OK"
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
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Top             =   1560
      Width           =   3135
      _ExtentX        =   5530
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
   Begin vkUserContolsXP.vkLabel vkLabel2 
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   344
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
   Begin vkUserContolsXP.vkLabel vkLabel1 
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   344
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
      Alignment       =   2
   End
   Begin vkUserContolsXP.vkTextBox txtPassword 
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   840
      Width           =   3135
      _ExtentX        =   5530
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
   Begin vkUserContolsXP.vkFrame vkFrame1 
      Height          =   3615
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6376
      Caption         =   "Password"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Label lblMyLabel 
         BackStyle       =   0  'Transparent
         Height          =   240
         Left            =   -1320
         TabIndex        =   7
         Top             =   480
         Width           =   3405
      End
   End
End
Attribute VB_Name = "frmTestPWD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Private arUserID()   As Byte
  Private arPWord()    As Byte
  Private m_strUserID  As String
  Private m_strPWord   As String
Private Sub ClearVariables()

' ---------------------------------------------------------------------------
' clear variables
' ---------------------------------------------------------------------------
  Erase arUserID()
  Erase arPWord()
  
  m_strUserID = String$(250, 0)
  m_strPWord = String$(250, 0)
  
  m_strUserID = ""
  m_strPWord = ""
  
End Sub




Private Sub cmdCancel_Click()
 Reset_frmTestPWD
Unload frmTestPWD
Unload mdiMain
End Sub

Private Sub CmdOK_Click()
    Dim strTmp1  As String
    Dim strTmp2  As String
    If Len(m_strUserID) = 0 Then
                  MsgBox "A user ID must be entered.", _
                         vbInformation Or vbOKOnly, "User ID missing"
                  txtPassword(0).SetFocus
                  Exit Sub
              Else
                  arUserID = ConvertToArray(m_strUserID)
              End If
                  
              ' Test length of password
              If Len(m_strPWord) = 0 Then
                  MsgBox "A password / passphrase must be entered.", _
                         vbInformation Or vbOKOnly, "Password / Passphrase missing"
                  Erase arPWord()
                  m_strPWord = ""
                  txtPassword(1).SetFocus
                  Exit Sub
              Else
                  arPWord = ConvertToArray(m_strPWord)
              End If
                  
              ' Test length of password data entered
              If Not Correct_Password_Length(arPWord()) Then
                  Erase arPWord()
                  m_strPWord = ""
                  txtPassword(1).Text = ""
                  txtPassword(1).SetFocus
                  Exit Sub
              End If
             
              ' Is this user on file?
              If Not Query_User(arUserID(), strTmp1, strTmp2) Then
                  MsgBox "User [ " & m_strUserID & _
                         " ] cannot be found in the database.", _
                         vbInformation Or vbOKOnly, "Invalid User ID"
                  
                  ' get rid of any data returned
                  strTmp1 = String$(250, 0)
                  strTmp2 = String$(250, 0)
                  Reset_frmTestPWD
                  Exit Sub
              Else
                  ' get rid of any data returned
                  strTmp1 = String$(250, 0)
                  strTmp2 = String$(250, 0)
              End If
              
              ' Compare with the data entered with the hashed results
              ' in the database.
              If Not Validate_Password(arUserID(), arPWord()) Then
                  Erase arPWord()
                  m_strPWord = ""
                  txtPassword(1).Text = ""
                  txtPassword(1).SetFocus
                  Exit Sub
              End If
              
              ' We were successful
              MsgBox "Successfully identified user [ " & _
                     m_strUserID & " ] in the database.", _
                     vbInformation Or vbOKOnly, "Success"
              Reset_frmTestPWD
              
End Sub

Private Sub Form_Initialize()
  Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
End Sub

Private Sub Form_Load()

  Me.Caption = g_strVersion
  CenterCaption frmTestPWD
  frmTestPWD.Hide
  g_intHashType = 1

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)


  ClearVariables
 
  Select Case UnloadMode
         
         Case 0: cmdCancel_Click
         Case 1: Exit Sub
         Case 2: TerminateApplication
         Case 3: TerminateApplication
         Case 4: TerminateApplication
  End Select
  
End Sub

Public Sub Reset_frmTestPWD()

' ---------------------------------------------------------------------------
' Empty variables
' ---------------------------------------------------------------------------
  ClearVariables
  
' ---------------------------------------------------------------------------
' Display the form
' ---------------------------------------------------------------------------
  With frmTestPWD
       .txtPassword(0).Text = ""
       .txtPassword(1).Text = ""
       .lblMyLabel = MYNAME
       .Show vbModeless
       .Refresh
  End With
  
' ---------------------------------------------------------------------------
' place cursor in first text box
' ---------------------------------------------------------------------------
  txtPassword(0).SetFocus
  
End Sub


Private Sub txtPassword_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim CtrlDown    As Integer
  Dim PressedKey  As Integer
  
' ---------------------------------------------------------------------------
' Initialize  variables
' ---------------------------------------------------------------------------
  CtrlDown = (Shift And vbCtrlMask) > 0   ' Define control key
  
  If Len(Trim$(KeyCode)) > 0 Then
      ' Convert to uppercase
      PressedKey = CInt(Asc(StrConv(Chr$(KeyCode), vbUpperCase)))
  End If
    
' ---------------------------------------------------------------------------
' Check to see if it is okay to make changes
' ---------------------------------------------------------------------------
  If CtrlDown And PressedKey = vbKeyX Then
      Edit_Cut            ' Ctrl + X was pressed
  ElseIf CtrlDown And PressedKey = vbKeyA Then
      SendKeys "{Home}{End}"
  ElseIf CtrlDown And PressedKey = vbKeyC Then
      Edit_Copy           ' Ctrl + C was pressed
  ElseIf CtrlDown And PressedKey = vbKeyV Then
      Edit_Paste          ' Ctrl + V was pressed
  ElseIf PressedKey = vbKeyDelete Then
      Edit_Delete         ' Delete key was pressed
  End If

End Sub

Private Sub txtPassword_KeyPress(Index As Integer, KeyAscii As Integer)

' ---------------------------------------------------------------------------
' If ENTER is pressed then nullify keystroke and press the OK button
' ---------------------------------------------------------------------------
  If KeyAscii = 13 Then
      KeyAscii = 0                        ' Nullify keystroke
      txtPassword_Validate Index, False   ' force validate event to fire
      CmdOK_Click
      Exit Sub
  End If
  
' ---------------------------------------------------------------------------
' If TAB is pressed then nullify keystroke and TAB to the next tabstop
' ---------------------------------------------------------------------------
  If KeyAscii = 9 Then
      KeyAscii = 0
      SendKeys "{TAB}"
  End If
  
' ---------------------------------------------------------------------------
' Accept on valid characters
' ---------------------------------------------------------------------------
  Select Case KeyAscii
  
         ' backspace and other printable keyboard characters
         Case 8, 32 To 126:
              Exit Sub      ' Good input
              
         ' Bad input
         Case Else:
              KeyAscii = 0  ' Nullify keystroke
  End Select
  
End Sub

Private Sub txtPassword_Validate(Index As Integer, Cancel As Boolean)


  Cancel = False
  txtPassword(Index).Text = Trim$(txtPassword(Index).Text)
  
  Select Case Index
  
         ' User ID
         Case 0:
              ' something may have changed
              ClearVariables
              txtPassword(1).Text = ""
              
              If Len(txtPassword(0).Text) > 0 Then
                                  
                  m_strUserID = txtPassword(0).Text
              End If
              
         ' Password
         Case Else:

              If Len(txtPassword(1).Text) > 0 Then
                                   
                  m_strPWord = txtPassword(1).Text
                  txtPassword(1).Text = String$(30, "*")
              Else
                  ' else empty the holding areas
                  txtPassword(1).Text = ""
                  Erase arPWord()
                  m_strPWord = ""
              End If
  End Select
  
End Sub
