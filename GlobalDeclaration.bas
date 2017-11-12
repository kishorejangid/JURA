Attribute VB_Name = "GlobalDeclaration"
Public Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetCursorPos Lib "User32" (lpPoint As PointAPI) As Long
Public Declare Function ScreenToClient Lib "User32" (ByVal hWnd As Long, lpPoint As PointAPI) As Long
Public Declare Function ReleaseCapture Lib "User32" () As Long
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function CreateThread Lib "Kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Public Declare Function TerminateThread Lib "Kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long

Public Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Public Const WS_CAPTION = &HC00000
Public Const WS_SYSMENU = &H80000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const GWL_STYLE = (-16)

Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_SYSCOMMAND = &H112

Public Type PointAPI
    X As Long
    Y As Long
End Type

Public conn As ADODB.Connection 'Oracle Connection String
Public admCheck As String       'Checks for Administrator
Public userCheck As String

Public iDept As Integer
Public iSem As Integer
Public iBatch As Integer
Public strSec As String

'Settings
Public flgRandomTheme  As Boolean
Public strExamMonth As String
Public strExamYear As String
Public strCollegeName As String


'Theme
Public Color As String
Public lColor As String

Public Const Orange = &H80FF&
Public Const LOrange = &H80C0FF
Public Const Blue = &HFF0000
Public Const LBlue = &HFFC0C0
Public Const Black = &H808080
Public Const LBlack = &HC0C0C0
Public Const Violet = &HC000C0
Public Const LViolet = &HFFC0FF
Public Const Cyan = &HC0C000
Public Const LCyan = &HFFFF80
Public Const Brown = &H4080&
Public Const LBrown = &H80C0FF
Public Const LGreen = &HFF00&
Public Const Green = &H8000&


'Constant For Grade

Public Const S = 10
Public Const a = 9
Public Const B = 8
Public Const C = 7
Public Const D = 6
Public Const E = 5
Public Const U = 0
Public Const W = 0
