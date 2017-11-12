VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkusercontrolsxp.ocx"
Begin VB.Form frmMarks 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Marks"
   ClientHeight    =   8160
   ClientLeft      =   0
   ClientTop       =   -30
   ClientWidth     =   12120
   Icon            =   "frmMarks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   12120
   ShowInTaskbar   =   0   'False
   Begin vkUserContolsXP.vkFrame fMarks 
      Height          =   8175
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   14420
      Caption         =   "Marks"
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
      TitleHeight     =   300
      BorderWidth     =   2
      Begin JURA.StylerButton cmdClose 
         Height          =   255
         Left            =   11640
         TabIndex        =   35
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
         Left            =   11400
         TabIndex        =   34
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
      Begin vkUserContolsXP.vkFrame vkFrame1 
         Height          =   2295
         Left            =   120
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   360
         Width           =   11910
         _ExtentX        =   21008
         _ExtentY        =   4048
         Caption         =   "Enter Your Data"
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
         TitleColor1     =   12640511
         TitleColor2     =   33023
         TitleGradient   =   2
         BorderColor     =   33023
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
            Left            =   8160
            TabIndex        =   38
            Top             =   720
            Width           =   1335
         End
         Begin vkUserContolsXP.vkLabel lblSec 
            Height          =   195
            Left            =   8160
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   480
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   344
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
         Begin VB.ComboBox cmbBatch 
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
            Left            =   4920
            TabIndex        =   36
            Top             =   720
            Width           =   1455
         End
         Begin VB.ComboBox cmbDept 
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
            Left            =   480
            TabIndex        =   1
            Top             =   720
            Width           =   4215
         End
         Begin vkUserContolsXP.vkLabel marks_lblSem 
            Height          =   255
            Left            =   6600
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   480
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
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
         Begin VB.ComboBox cmbSem 
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
            Left            =   6600
            TabIndex        =   2
            Top             =   720
            Width           =   1335
         End
         Begin vkUserContolsXP.vkLabel marks_lblBatch 
            Height          =   255
            Left            =   4920
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   480
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
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
         Begin vkUserContolsXP.vkLabel vkLabel2 
            Height          =   195
            Left            =   480
            TabIndex        =   0
            TabStop         =   0   'False
            Top             =   480
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   344
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
         Begin vkUserContolsXP.vkTextBox marks_txtName 
            Height          =   375
            Left            =   2760
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   1560
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
         Begin vkUserContolsXP.vkLabel vkLabel1 
            Height          =   255
            Left            =   2760
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   1320
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
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
         Begin VB.ComboBox cmbRegNo 
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
            Left            =   480
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   1560
            Width           =   2055
         End
         Begin vkUserContolsXP.vkLabel marks_lblReg 
            Height          =   255
            Left            =   480
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   1320
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
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
         Begin vkUserContolsXP.vkLabel lblRank 
            Height          =   255
            Left            =   9120
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   1320
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Rank:"
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
         Begin vkUserContolsXP.vkLabel marks_lblPercentagge 
            Height          =   255
            Left            =   7920
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   1320
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   450
            BackStyle       =   0
            Caption         =   "Percentage:"
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
         Begin vkUserContolsXP.vkLabel marks_lblTotal 
            Height          =   255
            Left            =   6600
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   1320
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            BackStyle       =   0
            Caption         =   "Total:"
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
         Begin vkUserContolsXP.vkTextBox marks_txtTotal 
            Height          =   375
            Left            =   6600
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   1560
            Width           =   1095
            _ExtentX        =   1931
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkTextBox marks_txtPercentage 
            Height          =   375
            Left            =   7920
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   1560
            Width           =   975
            _ExtentX        =   1720
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkTextBox txtRank 
            Height          =   375
            Left            =   9120
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   1560
            Width           =   975
            _ExtentX        =   1720
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkTextBox txtPrevRank 
            Height          =   375
            Left            =   10320
            TabIndex        =   22
            Top             =   1560
            Width           =   975
            _ExtentX        =   1720
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkLabel lblPrevRank 
            Height          =   255
            Left            =   10320
            TabIndex        =   21
            Top             =   1320
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Prev Rank:"
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
      Begin vkUserContolsXP.vkFrame marks_frmSemMarks 
         Height          =   4815
         Left            =   120
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2760
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   8493
         Caption         =   "Marks"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TitleColor1     =   12640511
         TitleColor2     =   33023
         TitleGradient   =   2
         BorderColor     =   33023
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid marks_MSHFlexGrid1 
            Height          =   3975
            Left            =   240
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   480
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   7011
            _Version        =   393216
            ForeColor       =   -2147483625
            Cols            =   5
            FixedCols       =   0
            BackColorFixed  =   31993
            ForeColorFixed  =   -2147483643
            BackColorSel    =   4629503
            ForeColorSel    =   16777215
            BackColorBkg    =   -2147483628
            GridColor       =   33023
            ScrollBars      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   5
         End
      End
      Begin vkUserContolsXP.vkFrame marks_frmRank 
         Height          =   2415
         Left            =   8040
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2760
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   4260
         Caption         =   "Rank"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TitleColor1     =   12640511
         TitleColor2     =   33023
         TitleGradient   =   2
         BorderColor     =   33023
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid marks_MSHFlexGrid2 
            Height          =   1815
            Left            =   240
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   360
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   3201
            _Version        =   393216
            ForeColor       =   -2147483625
            Cols            =   3
            BackColorFixed  =   33023
            ForeColorFixed  =   -2147483643
            BackColorSel    =   6730751
            BackColorBkg    =   -2147483634
            GridColor       =   33023
            FocusRect       =   0
            HighLight       =   0
            FillStyle       =   1
            ScrollBars      =   2
            _NumberOfBands  =   1
            _Band(0).Cols   =   3
         End
      End
      Begin vkUserContolsXP.vkFrame marks_frmTotal 
         Height          =   2295
         Left            =   8040
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   5280
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   4048
         Caption         =   "Arrears"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TitleColor1     =   12640511
         TitleColor2     =   33023
         TitleGradient   =   2
         BorderColor     =   33023
         Begin vkUserContolsXP.vkLabel vkLabel4 
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   480
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            BorderStyle     =   1
            BorderColor     =   33023
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   ""
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
         End
         Begin vkUserContolsXP.vkLabel vkLabel4 
            Height          =   375
            Index           =   1
            Left            =   2040
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   480
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            BorderStyle     =   1
            BorderColor     =   33023
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   ""
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
         End
         Begin vkUserContolsXP.vkLabel vkLabel4 
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   960
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            BorderStyle     =   1
            BorderColor     =   33023
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   ""
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
         End
         Begin vkUserContolsXP.vkLabel vkLabel4 
            Height          =   375
            Index           =   3
            Left            =   2040
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   960
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            BorderStyle     =   1
            BorderColor     =   33023
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   ""
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
         End
         Begin vkUserContolsXP.vkLabel vkLabel4 
            Height          =   375
            Index           =   4
            Left            =   240
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   1440
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            BorderStyle     =   1
            BorderColor     =   33023
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   ""
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
         End
         Begin vkUserContolsXP.vkLabel vkLabel4 
            Height          =   375
            Index           =   5
            Left            =   2040
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1440
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            BorderStyle     =   1
            BorderColor     =   33023
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   ""
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
         End
         Begin vkUserContolsXP.vkLabel lblmorearrears 
            Height          =   240
            Left            =   240
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   1920
            Visible         =   0   'False
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   423
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   " and some more arrears......."
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin vkUserContolsXP.vkTimer marks_timer1 
            Left            =   3360
            Top             =   1680
            _ExtentX        =   926
            _ExtentY        =   926
            Interval        =   10
         End
      End
      Begin vkUserContolsXP.vkCommand marks_cmdMarksheet 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   7680
         Visible         =   0   'False
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   661
         Caption         =   "Print MarkSheet"
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
      Begin vkUserContolsXP.vkCommand btnPdf 
         Height          =   375
         Left            =   6120
         TabIndex        =   5
         Top             =   7680
         Visible         =   0   'False
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   661
         Caption         =   "Save as PDF File"
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
      Begin vkUserContolsXP.vkBar marks_prgDatabase 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   7680
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   661
         BorderColor     =   33023
         LeftColor       =   33023
         RightColor      =   33023
         DisplayLabel    =   0
         BackPicture     =   "frmMarks.frx":57E2
         FrontPicture    =   "frmMarks.frx":57FE
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
End
Attribute VB_Name = "frmMarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sql As String
Public reg As String
Dim s2 As String
Dim s3 As String
Dim total As Integer
Dim arrear(10) As String
Dim State As Integer
Dim fTop As Integer

Private Sub btnPdf_Click()
    Call CreatePDF
End Sub

Private Sub cmbBatch_Click()
    iBatch = cmbBatch.Text
    Call cmbRegNo_Load(cmbRegNo)
    Call cmbRegNo_Click
    Call marks_MSHFlexGrid1_Load
    Call marks_MSHFlexGrid2_Data
End Sub

Private Sub cmbDept_Click()
    iDept = Department(cmbDept)
    Call cmbRegNo_Load(cmbRegNo)
    Call cmbRegNo_Click
    Call cmbSem_Click
    Call marks_MSHFlexGrid2_Data
End Sub

Private Sub cmbSec_Change()
    strSec = cmbSec.Text
    Call cmbRegNo_Load(cmbRegNo)
    Call cmbRegNo_Click
    Call marks_MSHFlexGrid2_Data
End Sub

Private Sub cmbSec_Click()
    strSec = cmbSec.Text
    Call cmbRegNo_Load(cmbRegNo)
    Call cmbRegNo_Click
    Call marks_MSHFlexGrid2_Data
End Sub

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub cmdMin_Click()
    Dim i As Long
    If State = 1 Then
        State = 0
        cmdMin.Caption = "+"
        For i = 8160 To 310 Step -150
            Me.Height = i
            DoEvents
        Next i
        Me.Height = 310
        fMarks.Height = 310
        fMarks.BorderWidth = 0
        Me.Top = 300
    Else
        State = 1
        cmdMin.Caption = "--"
        For i = 310 To 8160 Step 150
            Me.Height = i
            DoEvents
        Next i
        Me.Height = 8160
        fMarks.Height = 8160
        fMarks.BorderWidth = 2
        Me.Top = fTop
    End If
End Sub

Private Sub fMarks_MouseDown(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_Load()
    CreateRoundRectFromWindow Me, 7, 7
    Me.Top = (mdiMain.Height - Me.Height) / 5
    Me.Left = (mdiMain.Width - Me.Width) / 2
    Me.BackColor = mdiMain.BackColor
    Call frmColor(frmMarks)
    Call cmbDept_Load(cmbDept)
    Call cmbBatch_Load(cmbBatch)
    Call cmbSec_Load(cmbSec)
    Call cmbSem_Load(cmbSem)
    Call cmbRegNo_Load(cmbRegNo)
    Call cmbRegNo_Click
    Call marks_MSHFlexGrid1_Load
    Call marks_MSHFlexGrid2_Load
    Call cmbSem_Click
    Call Arrears
    State = 1
    fTop = Me.Top
End Sub
Private Sub cmbRegNo_Click()
    On Error GoTo ErrHndlr
    Dim rs As New ADODB.Recordset
    total = 0
    marks_txtName.Text = ""
    marks_MSHFlexGrid1.Clear
    reg = cmbRegNo.Text
    rs.CursorLocation = adUseClient
    sql = "select studname from studdetails where regno = '" & reg & "'"
    rs.Open sql, conn, adOpenDynamic, adLockOptimistic, -1
    marks_txtName.Text = rs.Fields(0)
    Call marks_MSHFlexGrid1_Data
    rs.Close
    txtRank.Text = StudRank(cmbRegNo.Text, iSem, iDept, iBatch, strSec)
    If iSem = 1 Then
        lblPrevRank.Caption = ""
        txtPrevRank.Text = ""
    Else
        lblPrevRank.Caption = "Sem " & iSem - 1 & " Rank"
        txtPrevRank.Text = StudRank(cmbRegNo.Text, iSem - 1, iDept, iBatch, strSec)
    End If
    Call Arrears
    Exit Sub
ErrHndlr:
    If Err.Number = 3021 Then
            MsgBox "No student found in the database for the given department,batch or section." & vbCrLf & vbCrLf & "Error Number: " & Err.Number
    Else
        MsgBox Error & vbCrLf & "Error Number: " & Err.Number & " cmbRegNo", vbCritical, App.Title
    End If
End Sub
Private Sub cmbRegNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmbRegNo_Click
End Sub

Private Sub cmbRegNo_LostFocus()
    Call cmbRegNo_Click
End Sub

Private Sub cmbSem_Click()
    total = 0
    txtRank.Text = ""
    iSem = cmbSem.Text
    marks_prgDatabase.Value = 0
    marks_timer1.Enabled = True
    Call marks_MSHFlexGrid2_Data
    Call marks_MSHFlexGrid1_Data
    txtRank.Text = StudRank(cmbRegNo.Text, iSem, iDept, iBatch, strSec)
    If iSem = 1 Then
        lblPrevRank.Caption = ""
        txtPrevRank.Text = ""
    Else
        lblPrevRank.Caption = "Sem " & iSem - 1 & " Rank"
        txtPrevRank.Text = StudRank(cmbRegNo.Text, iSem - 1, iDept, iBatch, strSec)
    End If
    Call Arrears
End Sub

Private Sub cmbSem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmbSem_Click
End Sub

Private Sub cmbSem_LostFocus()
    Call cmbSem_Click
End Sub

Private Sub marks_cmdMarksheet_Click()                                    'Generate A Report Which Can Be Printed And then  saves it In the html format for mailing
    On Error GoTo ErrHnd
    Dim rs As New ADODB.Recordset
    rs.CursorLocation = adUseClient
    sql = "select s2.subjcode,s2.subjname,s1.internals,nvl(s1.externals,0) as externals,(s1.internals+nvl(s1.externals,0)) as Sum  from studmarks s1,subj s2 where s1.subjcode=s2.subjcode and s1.semno=s2.semno and s1.dept=s2.dept and s1.batch=s2.batch and s1.semno= " & iSem & " and s1.regno= '" & cmbRegNo.Text & "' and s1.batch=" & Mid(iBatch, 3, 2) & ""
    rs.Open sql, conn, adOpenDynamic, adLockOptimistic, -1
    Set DataReport1.DataSource = rs
    DataReport1.DataMember = rs.DataMember
    total = 0
    Do While Not rs.EOF
        total = total + rs.Fields("Sum")
        rs.MoveNext
    Loop
    DataReport1.Sections("Section2").Controls("label11").Caption = reg
    DataReport1.Sections("Section2").Controls("label13").Caption = marks_txtName.Text
    DataReport1.Sections("Section2").Controls("label17").Caption = iSem
    DataReport1.Sections("Section5").Controls("label9").Caption = marks_txtPercentage.Text
    DataReport1.Sections("Section5").Controls("label15").Caption = total
    DataReport1.Sections("Section5").Controls("label21").Caption = arrear(0)
    DataReport1.Sections("Section5").Controls("label22").Caption = arrear(1)
    DataReport1.Sections("Section5").Controls("label23").Caption = arrear(2)
    DataReport1.Sections("Section5").Controls("label24").Caption = arrear(3)
    DataReport1.Sections("Section5").Controls("label27").Caption = arrear(4)
    DataReport1.Sections("Section5").Controls("label28").Caption = arrear(5)
    DataReport1.Sections("Section5").Controls("label29").Caption = arrear(6)
    DataReport1.Sections("Section5").Controls("label30").Caption = arrear(7)
    DataReport1.Sections("Section5").Controls("label31").Caption = arrear(8)
    DataReport1.Sections("Section5").Controls("label32").Caption = arrear(9)
    DataReport1.Sections("Section5").Controls("label26").Caption = txtRank.Text
    DataReport1.Sections("Section3").Controls("label18").Caption = Date & " " & Time
    DataReport1.LeftMargin = 100
    DataReport1.RightMargin = 100
    DataReport1.Caption = "Mark Sheet"
    DataReport1.Show
    DataReport1.ExportReport rptKeyHTML, App.Path & "\Reports\" & reg & "(" & iSem & ")" & ".html", True, False, rptRangeAllPages
    Exit Sub
ErrHnd:
    MsgBox Error & vbCrLf & "Error Number: " & Err.Number, vbCritical, "Error"
End Sub



Private Sub marks_MSHFlexGrid2_DblClick()
    With marks_MSHFlexGrid2
        If .MouseRow = 0 Then Exit Sub
        cmbRegNo.Text = .TextMatrix(.MouseRow, 1)
        Call cmbRegNo_KeyPress(13)
    End With
End Sub

Private Sub marks_timer1_Timer()
    marks_prgDatabase.Value = marks_prgDatabase.Value + 10
End Sub

Private Sub marks_MSHFlexGrid1_Load()
    marks_MSHFlexGrid1.Clear
    marks_MSHFlexGrid1.ColWidth(0) = 1175
    marks_MSHFlexGrid1.ColWidth(1) = 3500
    marks_MSHFlexGrid1.ColWidth(2) = 900
    marks_MSHFlexGrid1.ColWidth(3) = 900
    marks_MSHFlexGrid1.ColWidth(4) = 850
    marks_MSHFlexGrid1.RowHeightMin = 350
    
    marks_MSHFlexGrid1.TextMatrix(0, 0) = "Subject Code"
    marks_MSHFlexGrid1.TextMatrix(0, 1) = "Subject Name"
    
    If iBatch > 2007 Then
        marks_MSHFlexGrid1.TextMatrix(0, 2) = "Credit"
        marks_MSHFlexGrid1.TextMatrix(0, 3) = "Internals"
        marks_MSHFlexGrid1.TextMatrix(0, 4) = "Grade"
    Else
        marks_MSHFlexGrid1.TextMatrix(0, 2) = "Internals"
        marks_MSHFlexGrid1.TextMatrix(0, 3) = "Externals"
        marks_MSHFlexGrid1.TextMatrix(0, 4) = "Marks"
    End If
    
    marks_MSHFlexGrid1.ColAlignment(0) = flexAlignCenterCenter
    marks_MSHFlexGrid1.ColAlignment(2) = flexAlignCenterCenter
    marks_MSHFlexGrid1.ColAlignment(3) = flexAlignCenterCenter
    marks_MSHFlexGrid1.ColAlignment(4) = flexAlignCenterCenter
    marks_MSHFlexGrid1.ColAlignmentFixed(0) = flexAlignCenterCenter
    marks_MSHFlexGrid1.ColAlignmentFixed(1) = flexAlignCenterCenter
    marks_MSHFlexGrid1.ColAlignmentFixed(2) = flexAlignCenterCenter
    marks_MSHFlexGrid1.ColAlignmentFixed(3) = flexAlignCenterCenter
    marks_MSHFlexGrid1.ColAlignmentFixed(4) = flexAlignCenterCenter
End Sub

Private Sub marks_MSHFlexGrid1_Data()
    'On Error GoTo ErrHnd
    Dim rs As New ADODB.Recordset
    rs.CursorLocation = adUseClient
    If iBatch > 2007 Then
        sql = "select s2.subjcode,s2.subjname,s2.credit,s1.internals,s1.grade from studmarks s1,subj s2,studdetails s3 where s1.subjcode=s2.subjcode and s1.semno=s2.semno and s1.dept=s2.dept and s1.batch=s2.batch and s1.regno=s3.regno and s3.sec='" & strSec & "' and s1.semno= " & iSem & " and s1.regno= '" & cmbRegNo.Text & "' and s1.batch=" & Mid(iBatch, 3, 2) & ""
        rs.Open sql, conn, adOpenDynamic, adLockOptimistic, -1
        Set marks_MSHFlexGrid1.DataSource = rs
        marks_MSHFlexGrid1.Refresh
        rs.Close
'        marks_txtTotal.Text = CalcGPA(cmbRegNo.Text, cmbSem.Text, cmbDept.Text, cmbBatch.Text)
        marks_txtPercentage.Visible = False
        marks_lblPercentagge.Visible = False
        marks_lblTotal.Caption = "GPA"
    Else
        sql = "select s2.subjcode,s2.subjname,s1.internals,s1.externals,(s1.internals+s1.externals) as Sum  from studmarks s1,subj s2,studdetails s3 where s1.subjcode=s2.subjcode and s1.semno=s2.semno and s1.dept=s2.dept and s1.batch=s2.batch and s1.regno=s3.regno and s3.sec='" & strSec & "' and s1.semno= " & iSem & " and s1.regno= '" & cmbRegNo.Text & "' and s1.batch=" & Mid(iBatch, 3, 2) & " order by subjcode"
        rs.Open sql, conn, adOpenDynamic, adLockOptimistic, -1
        Set marks_MSHFlexGrid1.DataSource = rs
        marks_MSHFlexGrid1.Refresh
        rs.Close
        For i = 1 To (marks_MSHFlexGrid1.rows - 1)
            If marks_MSHFlexGrid1.TextMatrix(i, 4) = "" Then
                'Do Nothing
            Else
                total = total + marks_MSHFlexGrid1.TextMatrix(i, 4)
            End If
        Next
        marks_txtTotal.Text = total & "/ " & (GetSubjCount(iSem, iDept, iBatch) * 100)
        If total = 0 Then
            marks_txtPercentage.Text = 0
        Else
            marks_txtPercentage.Text = Round(total / GetSubjCount(iSem, iDept, iBatch), 2)
        End If
    End If
    Exit Sub
'ErrHnd:
 '   MsgBox Error & vbCrLf & "Error Number: " & Err.Number, vbCritical, "Error" & "MSHFlexGridData"
End Sub
Private Sub marks_MSHFlexGrid2_Load()
    marks_MSHFlexGrid2.ColWidth(0) = 700
    marks_MSHFlexGrid2.ColWidth(1) = 1400
    marks_MSHFlexGrid2.ColWidth(2) = 950
    marks_MSHFlexGrid2.TextMatrix(0, 0) = "Rank"
    marks_MSHFlexGrid2.TextMatrix(0, 1) = "Student"
    marks_MSHFlexGrid2.TextMatrix(0, 2) = "Percentage"
    marks_MSHFlexGrid2.ColAlignment(0) = flexAlignCenterCenter
    marks_MSHFlexGrid2.ColAlignment(1) = flexAlignCenterCenter
    marks_MSHFlexGrid2.ColAlignment(2) = flexAlignCenterCenter
    marks_MSHFlexGrid2.ColAlignmentFixed(0) = flexAlignCenterCenter
    marks_MSHFlexGrid2.ColAlignmentFixed(1) = flexAlignCenterCenter
    marks_MSHFlexGrid2.ColAlignmentFixed(2) = flexAlignCenterCenter
End Sub
Private Sub marks_MSHFlexGrid2_Data()                                     'Populates The ranking of the class in a grid
    On Error GoTo ErrHnd
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    Dim iSubjCount As Integer
    iSubjCount = GetSubjCount(iSem, iDept, iBatch)
    rs.CursorLocation = adUseClient
    sql = "select s1.regno,round((sum(s1.internals+s1.externals)/" & iSubjCount & "),2) as Percent from studmarks s1,studdetails s2 where s1.dept='" & iDept & "' and s1.batch='" & Mid(iBatch, 3, 2) & "' and s1.semno = '" & iSem & "' and s2.sec='" & strSec & "' and s1.regno=s2.regno group by s1.regno order by 2 desc"
    rs.Open sql, conn, adOpenDynamic, adLockOptimistic, -1
    Set marks_MSHFlexGrid2.DataSource = rs
    marks_MSHFlexGrid2.Refresh
    For i = 1 To rs.RecordCount
        marks_MSHFlexGrid2.TextMatrix(i, 0) = i
    Next i
    Call marks_MSHFlexGrid2_Load
    rs.Close
    Exit Sub
ErrHnd:
    MsgBox Error & vbCrLf & "Error Number: " & Err.Number, vbCritical, "Error" & "MSHFlexGridRankData"
End Sub
Private Sub Arrears()
    On Error Resume Next
    Dim r, C, z As Integer
    lblmorearrears.Visible = False
    For z = 0 To 5
        vkLabel4(z).Caption = ""
    Next
    For C = 0 To 9
        arrear(C) = ""
    Next
    C = 0
    For r = 1 To marks_MSHFlexGrid1.rows - 1
        If marks_MSHFlexGrid1.TextMatrix(r, 4) = "" Then
            arrear(C) = marks_MSHFlexGrid1.TextMatrix(r, 0)
            vkLabel4(C).Caption = marks_MSHFlexGrid1.TextMatrix(r, 0)
            C = C + 1
        ElseIf marks_MSHFlexGrid1.TextMatrix(r, 4) < 50 Then
            arrear(C) = marks_MSHFlexGrid1.TextMatrix(r, 0)
            vkLabel4(C).Caption = marks_MSHFlexGrid1.TextMatrix(r, 0)
            C = C + 1
        End If
    Next
    If Err.Number <> 0 Then
        If Err.Number = 340 Then
            lblmorearrears.Visible = True
        Else
            MsgBox "Error:" & Err.Description & vbCrLf & "Error Number:" & Err.Number & "Arrears"
        End If
    End If
End Sub

Private Sub CreatePDF()                                                   'Create a MarksSheet in Pdf Format
    On Error Resume Next
    Dim PDF As New clsPDF                                                  'Calls the Pdf Class
    Dim i, j As Integer
    Dim dLeft As Double
    
    PDF.PDFTitle = "MarksSheet"                                           'Pdf Title
    PDF.PDFFileName = App.Path & "\Reports\" & cmbRegNo.Text & "(" & cmbSem.Text & ")" & ".pdf"  'Saves The Pdf In the filename as Students Regno and Semester In the Folder Report at application Folder
    PDF.PDFLoadAfm = App.Path & "\Fonts"                                  'Font used in Pdf
    
    PDF.PDFView = True
    
    PDF.PDFSetBorder = BORDER_ALL
    PDF.PDFSetTopMargin = 1
    PDF.PDFSetBottomMargin = 1
    PDF.PDFSetRightMargin = 1
    PDF.PDFSetLeftMargin = 1
    PDF.PDFBeginDoc                                                       'Begins a new Page
        
        PDF.PDFDrawRectangle 1, 1, 19, 27                                 'Page Border
        
        PDF.PDFImage App.Path & "\Images\UniLogo.jpg", 2, 2, 2, 2.1     'Anna Univ Logo
        
        PDF.PDFSetTextColor = vbBlack
        PDF.PDFSetFont FONT_TIMES, 21, FONT_BOLD
        dLeft = PDF.PDFGetStringWidth("ANNA UNIVERSITY - TIRUNELVELI", "Times-Bold", 21)
        dLeft = (19 - (dLeft * 2.54) / 72) / 2
        PDF.PDFTextOut "ANNA UNIVERSITY - TIRUNELVELI", dLeft, 2
        
        PDF.PDFSetFont FONT_TIMES, 12, FONT_BOLD
        dLeft = PDF.PDFGetStringWidth("U.G. DEGREE EXAMINATIONS RESULT " & strExamMonth & " - " & strExamYer, "Times-Bold", 12)
        dLeft = (19 - (dLeft * 2.54) / 72) / 2
        PDF.PDFTextOut "U.G. DEGREE EXAMINATIONS RESULT " & strExamMonth & " - " & strExamYear, dLeft, 2.75
        
        dLeft = PDF.PDFGetStringWidth("MARK STATEMENT FOR " & Sem2Word(CInt(cmbSem.Text)) & " SEMESTER", "Times-Bold", 12)
        
        dLeft = (19 - (dLeft * 2.54) / 72) / 2
        PDF.PDFTextOut "MARK STATEMENT FOR " & Sem2Word(CInt(cmbSem.Text)) & "  SEMESTER", dLeft, 3.35
                                            
        PDF.PDFTextOut "Register Number:", 1, 5.25
        PDF.PDFTextOut cmbRegNo.Text, 5.5, 5.25
        PDF.PDFTextOut "Student Name:", 1, 5.85
        PDF.PDFTextOut GetStudName(cmbRegNo.Text), 5.5, 5.85
        PDF.PDFTextOut "Branch:", 1, 6.45
        
        If cmbDept.Text = "Information Technology" Then
            PDF.PDFTextOut "B.Tech - " & cmbDept.Text, 5.5, 6.45
        Else
            PDF.PDFTextOut "B.E - " & cmbDept.Text, 5.5, 6.45
        End If
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLineHor 1, 7, 19
        PDF.PDFDrawLineHor 1, 7.7, 19
        
                
        
        
        PDF.PDFSetTextColor = vbBlack
        PDF.PDFSetFont FONT_TIMES, 12, FONT_BOLD
        PDF.PDFTextOut "Subject Code", 1, 7.5
        PDF.PDFTextOut "Subject Name", 4, 7.5
        PDF.PDFTextOut "Internals", 10.25, 7.5
        PDF.PDFTextOut "Externals", 12.6, 7.5
        PDF.PDFTextOut "Marks", 15, 7.5
        PDF.PDFTextOut "Result", 16.75, 7.5
        
        
        
        PDF.PDFSetTextColor = vbBlack
        PDF.PDFSetFont FONT_TIMES, 12, FONT_NORMAL
        For i = 1 To GetSubjCount(iSem, iDept, iBatch)                                                 'Gets Marks From The MSHFlexGrid1
            PDF.PDFTextOut marks_MSHFlexGrid1.TextMatrix(i, 0), 1.5, 7.75 + i * 0.7
            PDF.PDFTextOut Mid(marks_MSHFlexGrid1.TextMatrix(i, 1), 1, 37), 4, 7.75 + i * 0.7
            PDF.PDFTextOut marks_MSHFlexGrid1.TextMatrix(i, 2), 11, 7.75 + i * 0.7
            PDF.PDFTextOut marks_MSHFlexGrid1.TextMatrix(i, 3), 13.2, 7.75 + i * 0.7
            PDF.PDFTextOut marks_MSHFlexGrid1.TextMatrix(i, 4), 15.25, 7.75 + i * 0.7
            If marks_MSHFlexGrid1.TextMatrix(i, 4) < 50 Then
                PDF.PDFTextOut "F", 17.25, 7.75 + i * 0.7
            Else
                PDF.PDFTextOut "P", 17.25, 7.75 + i * 0.7
            End If
        Next
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLineHor 1, 7.65 + i * 0.7, 19
        
        
        PDF.PDFTextOut "Total:", 1.5, 8.25 + i * 0.7
        PDF.PDFTextOut marks_txtTotal.Text, 3.65, 8.25 + i * 0.7
        PDF.PDFTextOut "Percentage:", 1.5, 8.85 + i * 0.7
        PDF.PDFTextOut marks_txtPercentage.Text, 3.65, 8.85 + i * 0.7
        
        'PDF.PDFTextOut "Faculty Advisor", 2, 27
        'PDF.PDFTextOut "H.O.D", 16, 27
                                                        
        
    PDF.PDFEndDoc                                                         'Ends The Page
    
    
End Sub




