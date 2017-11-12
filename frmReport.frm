VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkusercontrolsxp.ocx"
Begin VB.Form frmReport 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Report"
   ClientHeight    =   8475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14865
   Icon            =   "frmReport.frx":0000
   MDIChild        =   -1  'True
   ScaleHeight     =   8475
   ScaleWidth      =   14865
   ShowInTaskbar   =   0   'False
   Begin JURA.ThemedComboBox ThemedComboBox1 
      Left            =   12000
      Top             =   600
      _ExtentX        =   556
      _ExtentY        =   529
   End
   Begin vkUserContolsXP.vkCommand cmdPrint 
      Height          =   495
      Left            =   13200
      TabIndex        =   47
      Top             =   6960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
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
   Begin vkUserContolsXP.vkCommand cmdOrder 
      Height          =   495
      Left            =   13200
      TabIndex        =   46
      Top             =   6360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "Order By Total"
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
   Begin vkUserContolsXP.vkCommand cmdOK 
      Height          =   375
      Left            =   8400
      TabIndex        =   45
      Top             =   720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
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
   Begin vkUserContolsXP.vkFrame fReport 
      Height          =   8490
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   14880
      _ExtentX        =   26247
      _ExtentY        =   14975
      Caption         =   "Semester Report"
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
      TitleHeight     =   350
      Begin vkUserContolsXP.vkLabel lblSec 
         Height          =   270
         Left            =   7200
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   476
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
         Left            =   7200
         TabIndex        =   52
         Top             =   720
         Width           =   975
      End
      Begin vkUserContolsXP.vkLabel lblSem 
         Height          =   195
         Left            =   5760
         TabIndex        =   51
         Top             =   480
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   344
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
      Begin vkUserContolsXP.vkLabel lblBatch 
         Height          =   255
         Left            =   4320
         TabIndex        =   50
         Top             =   480
         Width           =   735
         _ExtentX        =   1296
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
      Begin vkUserContolsXP.vkLabel lblDept 
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
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
         ItemData        =   "frmReport.frx":57E2
         Left            =   4320
         List            =   "frmReport.frx":57E4
         TabIndex        =   44
         Top             =   720
         Width           =   1215
      End
      Begin JURA.StylerButton cmdClose 
         Height          =   255
         Left            =   14280
         TabIndex        =   41
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
         Left            =   14040
         TabIndex        =   40
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
         Left            =   360
         TabIndex        =   39
         Top             =   7080
         Visible         =   0   'False
         Width           =   4500
         _ExtentX        =   7938
         _ExtentY        =   661
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
         TitleHeight     =   350
      End
      Begin vkUserContolsXP.vkLabel lblInfo 
         Height          =   255
         Left            =   10560
         TabIndex        =   38
         Top             =   960
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   450
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
      Begin vkUserContolsXP.vkLabel lblPassPercentage 
         Height          =   255
         Left            =   10560
         TabIndex        =   37
         Top             =   720
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   450
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
         ItemData        =   "frmReport.frx":57E6
         Left            =   5760
         List            =   "frmReport.frx":57E8
         TabIndex        =   1
         Top             =   720
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   240
         TabIndex        =   35
         Top             =   7560
         Width           =   10095
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
            Height          =   350
            Left            =   120
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   240
            Width           =   8955
            _ExtentX        =   15796
            _ExtentY        =   609
            _Version        =   393216
            ForeColor       =   -2147483625
            Rows            =   11
            Cols            =   11
            FixedRows       =   0
            ForeColorFixed  =   -2147483643
            BackColorSel    =   6011383
            GridColor       =   33023
            FocusRect       =   0
            HighLight       =   0
            ScrollBars      =   0
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
            _Band(0).Cols   =   11
         End
         Begin JURA.StylerButton btnUp 
            Height          =   375
            Left            =   9120
            TabIndex        =   42
            Top             =   240
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            Caption         =   "+"
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
         Begin JURA.StylerButton btnDown 
            Height          =   375
            Left            =   9120
            TabIndex        =   43
            Top             =   240
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            Caption         =   "--"
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
         Left            =   240
         TabIndex        =   0
         Text            =   "Computer Science && Engineering"
         Top             =   720
         Width           =   3855
      End
      Begin vkUserContolsXP.vkFrame vkFrame2 
         Height          =   6975
         Left            =   10560
         TabIndex        =   4
         Top             =   1320
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   12303
         Caption         =   "Subjects"
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
         Begin vkUserContolsXP.vkCommand cmdBest 
            Height          =   495
            Left            =   2640
            TabIndex        =   48
            Top             =   6240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   873
            Caption         =   "Best && Last"
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
         Begin vkUserContolsXP.vkLabel lblArrears 
            Height          =   255
            Index           =   4
            Left            =   2040
            TabIndex        =   34
            Top             =   6480
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
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
         End
         Begin vkUserContolsXP.vkLabel lblArrears 
            Height          =   255
            Index           =   3
            Left            =   2040
            TabIndex        =   33
            Top             =   6120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
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
         End
         Begin vkUserContolsXP.vkLabel lblArrears 
            Height          =   255
            Index           =   2
            Left            =   2040
            TabIndex        =   32
            Top             =   5760
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
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
         End
         Begin vkUserContolsXP.vkLabel lblArrears 
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   31
            Top             =   5400
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
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
         End
         Begin vkUserContolsXP.vkLabel lblArrears 
            Height          =   255
            Index           =   0
            Left            =   2040
            TabIndex        =   30
            Top             =   5040
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
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
         End
         Begin vkUserContolsXP.vkLabel vkLabel1 
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   29
            Top             =   1800
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
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
         End
         Begin vkUserContolsXP.vkLabel vkLabel1 
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   28
            Top             =   2160
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
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
         End
         Begin vkUserContolsXP.vkLabel vkLabel1 
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   27
            Top             =   1440
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
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
         End
         Begin vkUserContolsXP.vkLabel vkLabel1 
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   26
            Top             =   1080
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
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
         End
         Begin vkUserContolsXP.vkLabel vkLabel1 
            Height          =   375
            Index           =   5
            Left            =   120
            TabIndex        =   25
            Top             =   2520
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
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
         End
         Begin vkUserContolsXP.vkLabel vkLabel1 
            Height          =   375
            Index           =   6
            Left            =   120
            TabIndex        =   24
            Top             =   2880
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
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
         End
         Begin vkUserContolsXP.vkLabel vkLabel1 
            Height          =   375
            Index           =   7
            Left            =   120
            TabIndex        =   23
            Top             =   3240
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
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
         End
         Begin vkUserContolsXP.vkLabel vkLabel1 
            Height          =   375
            Index           =   8
            Left            =   120
            TabIndex        =   22
            Top             =   3600
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
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
         End
         Begin vkUserContolsXP.vkLabel vkLabel1 
            Height          =   375
            Index           =   9
            Left            =   120
            TabIndex        =   21
            Top             =   3960
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
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
         End
         Begin vkUserContolsXP.vkLabel vkLabel2 
            Height          =   375
            Index           =   6
            Left            =   1080
            TabIndex        =   20
            Top             =   2880
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   661
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
         End
         Begin vkUserContolsXP.vkLabel vkLabel2 
            Height          =   375
            Index           =   7
            Left            =   1080
            TabIndex        =   19
            Top             =   3240
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   661
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
         End
         Begin vkUserContolsXP.vkLabel vkLabel2 
            Height          =   375
            Index           =   8
            Left            =   1080
            TabIndex        =   18
            Top             =   3600
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   661
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
         End
         Begin vkUserContolsXP.vkLabel vkLabel2 
            Height          =   375
            Index           =   9
            Left            =   1080
            TabIndex        =   17
            Top             =   3960
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   661
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
         End
         Begin vkUserContolsXP.vkLabel vkLabel6 
            Height          =   270
            Left            =   120
            TabIndex        =   16
            Top             =   6120
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   476
            BackStyle       =   0
            Caption         =   "With 4 Arrears:"
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
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   5760
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            BackStyle       =   0
            Caption         =   "With 3 Arrears:"
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
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   5400
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            BackStyle       =   0
            Caption         =   "With 2 Arrears:"
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
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   5040
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            BackStyle       =   0
            Caption         =   "With 1 Arrear:"
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
         Begin vkUserContolsXP.vkLabel vkLabel7 
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   6480
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            BackStyle       =   0
            Caption         =   "With More Than 4 Arrears:"
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
            Height          =   375
            Index           =   5
            Left            =   1080
            TabIndex        =   11
            Top             =   2520
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   661
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
         End
         Begin vkUserContolsXP.vkLabel vkLabel2 
            Height          =   375
            Index           =   4
            Left            =   1080
            TabIndex        =   10
            Top             =   2160
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   661
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
         End
         Begin vkUserContolsXP.vkLabel vkLabel2 
            Height          =   375
            Index           =   3
            Left            =   1080
            TabIndex        =   9
            Top             =   1800
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   661
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
         End
         Begin vkUserContolsXP.vkLabel vkLabel2 
            Height          =   375
            Index           =   2
            Left            =   1080
            TabIndex        =   8
            Top             =   1440
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   661
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
         End
         Begin vkUserContolsXP.vkLabel vkLabel2 
            Height          =   375
            Index           =   1
            Left            =   1080
            TabIndex        =   7
            Top             =   1080
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   661
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
         End
         Begin vkUserContolsXP.vkLabel vkLabel2 
            Height          =   375
            Index           =   0
            Left            =   1080
            TabIndex        =   6
            Top             =   720
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   661
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
         End
         Begin vkUserContolsXP.vkLabel vkLabel1 
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   720
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
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
         End
         Begin VB.Line Line1 
            X1              =   100
            X2              =   3930
            Y1              =   4800
            Y2              =   4800
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   6300
         Left            =   240
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1320
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   11113
         _Version        =   393216
         ForeColor       =   -2147483625
         Cols            =   12
         FixedCols       =   0
         ForeColorFixed  =   -2147483643
         BackColorSel    =   33023
         ForeColorSel    =   -2147483638
         BackColorBkg    =   16777215
         GridColor       =   33023
         GridLinesFixed  =   1
         ScrollBars      =   2
         BandDisplay     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   12
         _Band(0).GridLinesBand=   0
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim State As Integer
Dim ReportTop As Integer

Private Sub cmbSec_Click()
    Call cmbBatch_Click
End Sub

Private Sub cmdBest_Click()
    If iBatch > 2007 Then
        Call rptGradeBest
    Else
        Call rptBest
    End If
    'Call rptOverall
End Sub

Private Sub rptOverall()

    On Error Resume Next
    Dim strLeft As Double
    Dim rsRegCount As New ADODB.Recordset
    Dim sqlRegCount As String
    rsRegCount.CursorLocation = adUseClient
    sqlRegCount = "select count(regno) from studdetails where substr(regno,6,3)='" & iDept & "' and substr(regno,4,2)='" & Mid(iBatch, 3, 2) & "' and sec='" & strSec & "'"
    rsRegCount.Open sqlRegCount, conn, adOpenDynamic, adLockOptimistic
    
       
    Dim PDF As New clsPDF
    Dim i, j As Integer
    PDF.PDFTitle = "Best 5 & Last 5"
    PDF.PDFFileName = App.Path & "\Reports\" & "Best & Last" & ".pdf"
    PDF.PDFAuthor = "JURA"
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
        PDF.PDFSetFont FONT_TIMES, 14, FONT_BOLD
        l = PDF.PDFGetStringWidth("RESULT ANALYSIS OF EVEN SEMESTER (2010-2011) - AU TIRUNELVELI", "Times-Bold", 14)
        strLeft = (19 - (l * 2.54) / 72) / 2
        PDF.PDFTextOut "RESULT ANALYSIS OF EVEN SEMESTER (2010-2011) - AU TIRUNELVELI", strLeft, 3.65
        
        l = PDF.PDFGetStringWidth("DEPARTMENT OF " & UCase(cmbDept.Text), "Times-Bold", 14)
        strLeft = (19 - (l * 2.54) / 72) / 2
        PDF.PDFTextOut "DEPARTMENT OF " & UCase(cmbDept.Text), strLeft, 4.4
        
        l = PDF.PDFGetStringWidth("OVERALL REPORT", "Times-Bold", 14)
        strLeft = (19 - (l * 2.54) / 72) / 2
        PDF.PDFTextOut "OVERALL REPORT", strLeft, 5.15
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 1, 5.5, 20, 5.5
        PDF.PDFSetLineWidth = 0.025
        PDF.PDFDrawLine 1, 5.55, 20, 5.55
        
        PDF.PDFSetFont 2, 12, FONT_BOLD
        PDF.PDFTextOut "Batch:", 2, 6
        PDF.PDFTextOut cmbBatch.Text, 3.5, 6
        PDF.PDFTextOut "Semester:", 13.5, 6
        PDF.PDFTextOut cmbSem.Text, 15.5, 6
        
                      
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 1, 6.25, 20, 6.25
        PDF.PDFSetLineWidth = 0.02
        PDF.PDFDrawLine 1, 6.3, 20, 6.3
        
        
         
        PDF.PDFSetFont 2, 10, FONT_NORMAL
        PDF.PDFTextOut "No Of Students In The Class:", 2, 8
        PDF.PDFTextOut rsRegCount.Fields(0), 12, 8
        PDF.PDFTextOut "No Of Students Appeared For Examination:", 2, 8.75
        PDF.PDFTextOut "No Of Students Get Passed:", 2, 9.5
        PDF.PDFTextOut "No Of Students Failed:", 2, 10.25
        PDF.PDFTextOut "No Of Students With Shortage Of Attendance:", 2, 11
        PDF.PDFTextOut "No Of Students With WH (Want Of Clarification):", 2, 11.75
        PDF.PDFTextOut "No Of Students With WH1:", 2, 12.5
        PDF.PDFTextOut "No Of Students With WH2 (Malpractise):", 2, 13.25
               
        
                
        PDF.PDFDrawLineHor 1, 22.25, 19
        
        PDF.PDFSetFont FONT_TIMES, 12, FONT_BOLD
        PDF.PDFTextOut "Staff Incharge", 2.5, 25.75
        PDF.PDFTextOut "HOD", 9, 25.75
        PDF.PDFTextOut "Principal", 15, 25.75
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 1, 27.25, 20, 27.25
        PDF.PDFSetLineWidth = 0.025
        PDF.PDFDrawLine 1, 27.3, 20, 27.3
        PDF.PDFSetFont 2, 10, FONT_NORMAL
        PDF.PDFTextOut "Report Generated By JURA", 7.5, 27.75
    PDF.PDFEndDoc
End Sub

Private Sub rptGradeBest()
    On Error Resume Next
    Dim rsBest As New ADODB.Recordset
    Dim sqlBest As String
    rsBest.CursorLocation = adUseClient
    sqlBest = "select regno,gpa,rank from (SELECT s1.regno,round(sum(s1.value*s2.credit)/(select sum(credit) FROM subj WHERE batch=" & Mid(iBatch, 3, 2) & " AND semno=" & iSem & " AND dept=" & Department(cmbDept) & "),2) AS GPA,Dense_rank() over (order by (round(sum(s1.value*s2.credit)/(select sum(credit) FROM subj WHERE batch=" & Mid(iBatch, 3, 2) & " AND semno=" & iSem & " AND dept=" & Department(cmbDept) & "),2)) desc) as RANK FROM studmarks s1,subj s2 WHERE s1.batch=" & Mid(iBatch, 3, 2) & " AND s1.semno=" & iSem & " AND s1.subjcode=s2.subjcode and s1.semno=s2.semno and s1.dept=s2.dept and s1.batch=s2.batch GROUP BY s1.regno) where RANK<=5"
    rsBest.Open sqlBest, conn, adOpenDynamic, adLockOptimistic
    
    Dim rsLast As New ADODB.Recordset
    Dim sqlLast As String
    rsLast.CursorLocation = adUseClient
    sqlLast = "select regno,gpa,rank from (SELECT s1.regno,round(sum(s1.value*s2.credit)/(select sum(credit) FROM subj WHERE batch=" & Mid(iBatch, 3, 2) & " AND semno=" & iSem & " AND dept=" & Department(cmbDept) & "),2) AS GPA,Dense_rank() over (order by (round(sum(s1.value*s2.credit)/(select sum(credit) FROM subj WHERE batch=" & Mid(iBatch, 3, 2) & " AND semno=" & iSem & " AND dept=" & Department(cmbDept) & "),2)) asc) as RANK FROM studmarks s1,subj s2 WHERE s1.batch=" & Mid(iBatch, 3, 2) & " AND s1.semno=" & iSem & " AND s1.subjcode=s2.subjcode and s1.semno=s2.semno and s1.dept=s2.dept and s1.batch=s2.batch GROUP BY s1.regno) where RANK<=5"
    rsLast.Open sqlLast, conn, adOpenDynamic, adLockOptimistic
    
    Dim PDF As New clsPDF
    Dim i, j As Integer
    PDF.PDFTitle = "Best 5 & Last 5"
    PDF.PDFFileName = App.Path & "\Reports\" & "Best & Last" & ".pdf"
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
        PDF.PDFTextOut "DEPARTMENT OF " & UCase(cmbDept.Text), 3.6, 4.4
        PDF.PDFTextOut "OVERALL CLASS TOPPERS & SLOW LEARNERS ", 3.7, 5.15
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 1, 5.5, 20, 5.5
        PDF.PDFSetLineWidth = 0.025
        PDF.PDFDrawLine 1, 5.55, 20, 5.55
        
        PDF.PDFSetFont 2, 12, FONT_BOLD
        PDF.PDFTextOut "Batch:", 2, 6
        PDF.PDFTextOut cmbBatch.Text, 3.5, 6
        PDF.PDFTextOut "Semester:", 13.5, 6
        PDF.PDFTextOut cmbSem.Text, 15.5, 6
        
                      
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 1, 6.25, 20, 6.25
        PDF.PDFSetLineWidth = 0.02
        PDF.PDFDrawLine 1, 6.3, 20, 6.3
        
        PDF.PDFTextOut "Register No", 2, 6.9
        PDF.PDFTextOut "Student Name", 6.5, 6.9
        If iBatch > 2007 Then
            PDF.PDFTextOut "GPA", 15.5, 6.9
        Else
            PDF.PDFTextOut "Percentage", 15, 6.9
        End If
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 1, 7.1, 20, 7.1
        PDF.PDFSetLineWidth = 0.025
        PDF.PDFDrawLine 1, 7.15, 20, 7.15
        
        
        PDF.PDFTextOut "Best 5:", 1, 7.65
        PDF.PDFTextOut "Last 5:", 1, 15
        
        PDF.PDFSetFont 2, 10, FONT_NORMAL
        For i = 1 To rsBest.RecordCount
            PDF.PDFTextOut rsBest.Fields(0), 2, 7.85 + i * 0.5
            PDF.PDFTextOut GetStudName(rsBest.Fields(0)), 6.5, 7.85 + i * 0.5
            PDF.PDFTextOut rsBest.Fields(1), 15.75, 7.85 + i * 0.5
            rsBest.MoveNext
        Next
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 1, 14.5, 20, 14.5
        PDF.PDFSetLineWidth = 0.025
        PDF.PDFDrawLine 1, 14.5, 20, 14.5
        
               
        PDF.PDFSetFont 2, 10, FONT_NORMAL
        For i = 1 To rsLast.RecordCount
            PDF.PDFTextOut rsLast.Fields(0), 2, 15.35 + i * 0.5
            PDF.PDFTextOut GetStudName(rsLast.Fields(0)), 6.5, 15.35 + i * 0.5
            PDF.PDFTextOut rsLast.Fields(1), 15.75, 15.35 + i * 0.5
            rsLast.MoveNext
        Next
        
        
        'Vertical Lines
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 6.75, 6.3, 6.75, 22.25
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 14.5, 6.3, 14.5, 22.25
        
        PDF.PDFDrawLineHor 1, 22.25, 19
        
        PDF.PDFSetFont FONT_TIMES, 12, FONT_BOLD
        PDF.PDFTextOut "Staff Incharge", 2.5, 25.75
        PDF.PDFTextOut "HOD", 9, 25.75
        PDF.PDFTextOut "Principal", 15, 25.75
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 1, 27.25, 20, 27.25
        PDF.PDFSetLineWidth = 0.025
        PDF.PDFDrawLine 1, 27.3, 20, 27.3
        PDF.PDFSetFont 2, 10, FONT_NORMAL
        PDF.PDFTextOut "Report Generated By JURA", 7.5, 27.75
    PDF.PDFEndDoc
End Sub
Private Sub rptBest()
    On Error Resume Next
    Dim rsBest As New ADODB.Recordset
    Dim sqlBest As String
    rsBest.CursorLocation = adUseClient
    sqlBest = "select regno,sum from (select regno,round(avg(internals+externals),2) as sum,Dense_rank() over (order by nvl(round(avg(internals+externals),2),0) desc) b from studmarks where semno=" & iSem & " and dept=" & Department(cmbDept) & " and batch=" & Mid(iBatch, 3, 2) & " group by regno ) where b<=5"
    rsBest.Open sqlBest, conn, adOpenDynamic, adLockOptimistic
    
    Dim rsLast As New ADODB.Recordset
    Dim sqlLast As String
    rsLast.CursorLocation = adUseClient
    sqlLast = "select regno,sum from (select regno,round(avg(internals+externals),2) as sum,Dense_rank() over (order by round(avg(internals+externals),2) asc) b from studmarks where semno=" & iSem & " and dept=" & Department(cmbDept) & " and batch=" & Mid(iBatch, 3, 2) & " group by regno ) where b<=5"
    rsLast.Open sqlLast, conn, adOpenDynamic, adLockOptimistic
    
    Dim PDF As New clsPDF
    Dim i, j As Integer
    PDF.PDFTitle = "Best 5 & Last 5"
    PDF.PDFFileName = App.Path & "\Reports\" & "Best & Last" & ".pdf"
    PDF.PDFAuthor = "Jangid's University Result Analysis"
    PDF.PDFCreator = "Jangid"
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
        PDF.PDFTextOut "DEPARTMENT OF " & UCase(cmbDept.Text), 3.6, 4.4
        PDF.PDFTextOut "OVERALL CLASS TOPPERS & SLOW LEARNERS ", 3.7, 5.15
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 1, 5.5, 20, 5.5
        PDF.PDFSetLineWidth = 0.025
        PDF.PDFDrawLine 1, 5.55, 20, 5.55
        
        PDF.PDFSetFont 2, 12, FONT_BOLD
        PDF.PDFTextOut "Batch:", 2, 6
        PDF.PDFTextOut cmbBatch.Text, 3.5, 6
        PDF.PDFTextOut "Semester:", 13.5, 6
        PDF.PDFTextOut cmbSem.Text, 15.5, 6
        
                      
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 1, 6.25, 20, 6.25
        PDF.PDFSetLineWidth = 0.02
        PDF.PDFDrawLine 1, 6.3, 20, 6.3
        
        PDF.PDFTextOut "Register No", 2, 6.9
        PDF.PDFTextOut "Student Name", 6.5, 6.9
        PDF.PDFTextOut "Percentage", 15, 6.9
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 1, 7.1, 20, 7.1
        PDF.PDFSetLineWidth = 0.025
        PDF.PDFDrawLine 1, 7.15, 20, 7.15
        
        
        PDF.PDFTextOut "Best 5:", 1, 7.65
        PDF.PDFTextOut "Last 5:", 1, 15
        
        PDF.PDFSetFont 2, 10, FONT_NORMAL
        For i = 1 To rsBest.RecordCount
            PDF.PDFTextOut rsBest.Fields(0), 2, 7.85 + i * 0.5
            PDF.PDFTextOut GetStudName(rsBest.Fields(0)), 6.5, 7.85 + i * 0.5
            PDF.PDFTextOut rsBest.Fields(1), 15.75, 7.85 + i * 0.5
            rsBest.MoveNext
        Next
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 1, 14.5, 20, 14.5
        PDF.PDFSetLineWidth = 0.025
        PDF.PDFDrawLine 1, 14.5, 20, 14.5
        
               
        PDF.PDFSetFont 2, 10, FONT_NORMAL
        For i = 1 To rsLast.RecordCount
            PDF.PDFTextOut rsLast.Fields(0), 2, 15.35 + i * 0.5
            PDF.PDFTextOut GetStudName(rsLast.Fields(0)), 6.5, 15.35 + i * 0.5
            PDF.PDFTextOut rsLast.Fields(1), 15.75, 15.35 + i * 0.5
            rsLast.MoveNext
        Next
        
        
        'Vertical Lines
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 6.75, 6.3, 6.75, 22.25
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 14.5, 6.3, 14.5, 22.25
        
        PDF.PDFDrawLineHor 1, 22.25, 19
        
        PDF.PDFSetFont FONT_TIMES, 12, FONT_BOLD
        PDF.PDFTextOut "Staff Incharge", 2.5, 25.75
        PDF.PDFTextOut "HOD", 9, 25.75
        PDF.PDFTextOut "Principal", 15, 25.75
        
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 1, 27.25, 20, 27.25
        PDF.PDFSetLineWidth = 0.025
        PDF.PDFDrawLine 1, 27.3, 20, 27.3
        PDF.PDFSetFont 2, 10, FONT_NORMAL
        PDF.PDFTextOut "Report Generated By JURA", 7.5, 27.75
    PDF.PDFEndDoc
End Sub

Private Sub Form_Load()
    CreateRoundRectFromWindow Me, 7, 7
    Me.Top = 250
    Me.Left = 250
    Me.BackColor = mdiMain.BackColor
    Call MSHFlexGrid2_load
    Call MSHFlexGrid3_Load
    Call cmbDept_Load(cmbDept)
    Call cmbBatch_Load(cmbBatch)
    Call cmbSem_Load(cmbSem)
    Call cmbSec_Load(cmbSec)
    Call frmColor(frmReport)
    btnDown.Visible = False
    State = 1
    ReportTop = Me.Top
End Sub
Private Sub cmbDept_Click()
    MSHFlexGrid2.Clear
    MSHFlexGrid3.Clear
    For i = 0 To 9
        vkLabel1(i).Caption = ""
        vkLabel2(i).Caption = ""
    Next
    lblArrears(0).Caption = ""
    lblArrears(1).Caption = ""
    lblArrears(2).Caption = ""
    lblArrears(3).Caption = ""
    lblArrears(4).Caption = ""
    iDept = Department(cmbDept)
    iSem = cmbSem.Text
    iBatch = cmbBatch.Text
End Sub
Private Sub cmbDept_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmbDept_Click
End Sub
Private Sub cmbBatch_Click()
    MSHFlexGrid2.Clear
    MSHFlexGrid3.Clear
    For i = 0 To 9
        vkLabel1(i).Caption = ""
        vkLabel2(i).Caption = ""
    Next
    lblArrears(0).Caption = ""
    lblArrears(1).Caption = ""
    lblArrears(2).Caption = ""
    lblArrears(3).Caption = ""
    lblArrears(4).Caption = ""
    iBatch = cmbBatch.Text
End Sub
Private Sub cmbBatch_Change()
    Call cmbBatch_Click
End Sub
Private Sub cmbSem_Click()
    On Error Resume Next
    MSHFlexGrid2.Clear
    MSHFlexGrid3.Clear
    For i = 0 To 9
        vkLabel1(i).Caption = ""
        vkLabel2(i).Caption = ""
    Next
    lblArrears(0).Caption = ""
    lblArrears(1).Caption = ""
    lblArrears(2).Caption = ""
    lblArrears(3).Caption = ""
    lblArrears(4).Caption = ""
    iSem = cmbSem.Text
End Sub
Private Sub cmbSem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmbSem_Click
End Sub
Private Sub cmbSem_Change()
    Call cmbSem_Click
End Sub
Private Sub cmdMin_Click()
    Dim i As Long
    If State = 1 Then
        State = 0
        cmdMin.Caption = "+"
        For i = 8475 To 310 Step -150
            Me.Height = i
            DoEvents
        Next i
        Me.Height = 310
        fReport.Height = 310
        fReport.BorderWidth = 0
        Me.Top = 100
    Else
        State = 1
        cmdMin.Caption = "--"
        For i = 310 To 8475 Step 150
            Me.Height = i
            DoEvents
        Next i
        Me.Height = 8475
        fReport.Height = 8475
        fReport.BorderWidth = 1
        Me.Top = ReportTop
    End If
End Sub
Private Sub CmdClose_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    MSHFlexGrid2.Clear
    Call rptSubjCode
    MSHFlexGrid2.Refresh
    If iBatch > 2007 Then
        Call rptGradeLoad
    Else
        Call rptLoad
    End If
    Call Analyse
End Sub
Private Sub cmdOrder_Click()
    MSHFlexGrid2.Clear
    Call rptSubjCode
    If iBatch > 2007 Then
        Call rptGradeOrderedLoad
    Else
        Call rptOrderedLoad
    End If
End Sub
Private Sub Analyse()
    On Error Resume Next
       
    Call btnup_Click
    Call MSHFlexGrid3_Load
    
    If iBatch > 2007 Then
        lblArrears(0).Caption = ArrearCount(iSem, iBatch, iDept, "=1")
        lblArrears(1).Caption = ArrearCount(iSem, iBatch, iDept, "=2")
        lblArrears(2).Caption = ArrearCount(iSem, iBatch, iDept, "=3")
        lblArrears(3).Caption = ArrearCount(iSem, iBatch, iDept, "=4")
        lblArrears(4).Caption = ArrearCount(iSem, iBatch, iDept, ">4")
        
        For i = 1 To GetSubjCount(iSem, iDept, iBatch)
            MSHFlexGrid3.TextMatrix(0, i) = GetNoOfStudAppeared(iSem, CInt(iDept), iBatch, MSHFlexGrid2.TextMatrix(0, i))
            MSHFlexGrid3.TextMatrix(1, i) = GetGradeCount(iSem, CInt(iDept), iBatch, MSHFlexGrid2.TextMatrix(0, i), "any ('S','A','B','C','D','E')")
            MSHFlexGrid3.TextMatrix(2, i) = GetGradeCount(iSem, CInt(iDept), iBatch, MSHFlexGrid2.TextMatrix(0, i), "any ('I','U','W')")
            MSHFlexGrid3.TextMatrix(3, i) = GetGradeCount(iSem, CInt(iDept), iBatch, MSHFlexGrid2.TextMatrix(0, i), "'S'")
            MSHFlexGrid3.TextMatrix(4, i) = GetGradeCount(iSem, CInt(iDept), iBatch, MSHFlexGrid2.TextMatrix(0, i), "'A'")
            MSHFlexGrid3.TextMatrix(5, i) = GetGradeCount(iSem, CInt(iDept), iBatch, MSHFlexGrid2.TextMatrix(0, i), "'B'")
            MSHFlexGrid3.TextMatrix(6, i) = GetGradeCount(iSem, CInt(iDept), iBatch, MSHFlexGrid2.TextMatrix(0, i), "'C'")
            MSHFlexGrid3.TextMatrix(7, i) = GetGradeCount(iSem, CInt(iDept), iBatch, MSHFlexGrid2.TextMatrix(0, i), "'D'")
            MSHFlexGrid3.TextMatrix(8, i) = GetGradeCount(iSem, CInt(iDept), iBatch, MSHFlexGrid2.TextMatrix(0, i), "'E'")
        Next
    Else
        lblArrears(0).Caption = ArrearCount(iSem, iBatch, iDept, "=1")
        lblArrears(1).Caption = ArrearCount(iSem, iBatch, iDept, "=2")
        lblArrears(2).Caption = ArrearCount(iSem, iBatch, iDept, "=3")
        lblArrears(3).Caption = ArrearCount(iSem, iBatch, iDept, "=4")
        lblArrears(4).Caption = ArrearCount(iSem, iBatch, iDept, ">4")
        
        For i = 1 To GetSubjCount(iSem, iDept, iBatch)
            MSHFlexGrid3.TextMatrix(0, i) = GetNoOfStudAppeared(iSem, CInt(iDept), iBatch, MSHFlexGrid2.TextMatrix(0, i))
            MSHFlexGrid3.TextMatrix(1, i) = GetCount(iSem, CInt(iDept), iBatch, MSHFlexGrid2.TextMatrix(0, i), 50, 100)
            MSHFlexGrid3.TextMatrix(2, i) = GetCount(iSem, CInt(iDept), iBatch, MSHFlexGrid2.TextMatrix(0, i), 0, 49)
            MSHFlexGrid3.TextMatrix(3, i) = GetCount(iSem, CInt(iDept), iBatch, MSHFlexGrid2.TextMatrix(0, i), 50, 59)
            MSHFlexGrid3.TextMatrix(4, i) = GetCount(iSem, CInt(iDept), iBatch, MSHFlexGrid2.TextMatrix(0, i), 60, 69)
            MSHFlexGrid3.TextMatrix(5, i) = GetCount(iSem, CInt(iDept), iBatch, MSHFlexGrid2.TextMatrix(0, i), 70, 79)
            MSHFlexGrid3.TextMatrix(6, i) = GetCount(iSem, CInt(iDept), iBatch, MSHFlexGrid2.TextMatrix(0, i), 80, 89)
            MSHFlexGrid3.TextMatrix(7, i) = GetCount(iSem, CInt(iDept), iBatch, MSHFlexGrid2.TextMatrix(0, i), 90, 100)
            MSHFlexGrid3.TextMatrix(8, i) = GetMinMarks(iSem, CInt(iDept), iBatch, MSHFlexGrid2.TextMatrix(0, i))
            MSHFlexGrid3.TextMatrix(9, i) = GetMaxMarks(iSem, CInt(iDept), iBatch, MSHFlexGrid2.TextMatrix(0, i))
            MSHFlexGrid3.TextMatrix(10, i) = GetAvgMarks(iSem, CInt(iDept), iBatch, MSHFlexGrid2.TextMatrix(0, i))
        Next
    End If
End Sub
Private Sub rptSubjCode()
    Dim rs As New ADODB.Recordset
    Dim sql As String
    Dim i As Integer
    rs.CursorLocation = adUseClient
    sql = "select subjcode,subjname from subj where semno=" & iSem & " and dept='" & iDept & "' and batch=" & Mid(iBatch, 3, 2) & " order by subjcode"
    rs.Open sql, conn, adOpenDynamic, adLockOptimistic
    MSHFlexGrid2.TextMatrix(0, 0) = "REGNO"
    If iBatch > 2007 Then
        MSHFlexGrid2.TextMatrix(0, 11) = "GPA"
    Else
        MSHFlexGrid2.TextMatrix(0, 11) = "Total"
    End If
    For i = 0 To rs.RecordCount - 1
        MSHFlexGrid2.TextMatrix(0, i + 1) = rs.Fields(0)
        vkLabel1(i).Caption = rs.Fields(0)
        vkLabel2(i).Caption = rs.Fields(1)
        rs.MoveNext
    Next
End Sub
Private Sub rptLoad()
    On Error Resume Next
    Dim sqlRegNo As String
    Dim sqlMarks As String
    Dim strRegNo As String
    Dim i As Integer
    Dim j As Integer
    Dim iTotal As Integer
    Dim rsRegNo As New ADODB.Recordset
    Dim rsMarks As New ADODB.Recordset
    
    rsRegNo.CursorLocation = adUseClient
    rsMarks.CursorLocation = adUseClient
    sqlRegNo = "select regno from studdetails where substr(regno, 6, 3) = '" & iDept & "' and substr(regno,4,2)=" & Mid(iBatch, 3, 2) & " and sec= '" & cmbSec.Text & "' order by regno"
    rsRegNo.Open sqlRegNo, conn, adOpenDynamic, adLockOptimistic
    MSHFlexGrid2.rows = rsRegNo.RecordCount + 1
    
    For i = 0 To rsRegNo.RecordCount - 1
        iTotal = 0
        strRegNo = rsRegNo.Fields(0)
        MSHFlexGrid2.TextMatrix(i + 1, 0) = strRegNo
        rsRegNo.MoveNext
        For j = 1 To GetSubjCount(iSem, iDept, iBatch)
            sqlMarks = "select (internals+externals),result from studmarks where semno=" & iSem & " and batch=" & Mid(iBatch, 3, 2) & "and regno='" & strRegNo & "' and subjcode='" & MSHFlexGrid2.TextMatrix(0, j) & "'"
            rsMarks.Open sqlMarks, conn, adOpenDynamic, adLockOptimistic
            If rsMarks.Fields(1) = "AB" Then
                MSHFlexGrid2.TextMatrix(i + 1, j) = "AB"
            ElseIf rsMarks.Fields(1) = "WH2" Then
                MSHFlexGrid2.TextMatrix(i + 1, j) = "WH2"
            ElseIf rsMarks.Fields(1) = "WH1" Then
                MSHFlexGrid2.TextMatrix(i + 1, j) = "WH1"
            ElseIf rsMarks.Fields(1) = "SA" Then
                MSHFlexGrid2.TextMatrix(i + 1, j) = "SA"
            ElseIf rsMarks.Fields(1) = "WH" Then
                MSHFlexGrid2.TextMatrix(i + 1, j) = "WH"
            ElseIf rsMarks.Fields(0) < 50 Then
                MSHFlexGrid2.Row = i + 1
                MSHFlexGrid2.Col = j
                MSHFlexGrid2.TextMatrix(i + 1, j) = rsMarks.Fields(0)
                MSHFlexGrid2.CellForeColor = vbRed
            Else
                MSHFlexGrid2.TextMatrix(i + 1, j) = rsMarks.Fields(0)
            End If
            iTotal = iTotal + rsMarks.Fields(0)
            rsMarks.Close
        Next
        If iTotal = 0 Then
            MSHFlexGrid2.TextMatrix(i + 1, 11) = ""
        Else
            MSHFlexGrid2.TextMatrix(i + 1, 11) = iTotal
        End If
    Next
    MSHFlexGrid2.Refresh
End Sub
Private Sub rptGradeLoad()
    On Error Resume Next
    Dim sqlRegNo As String
    Dim sqlMarks As String
    Dim strRegNo As String
    Dim i As Integer
    Dim j As Integer
    Dim iTotal As Integer
    Dim rsRegNo As New ADODB.Recordset
    Dim rsMarks As New ADODB.Recordset
    
    rsRegNo.CursorLocation = adUseClient
    rsMarks.CursorLocation = adUseClient
    sqlRegNo = "select regno from studdetails where substr(regno, 6, 3) = '" & iDept & "' and substr(regno,4,2)=" & Mid(iBatch, 3, 2) & " and sec= '" & cmbSec.Text & "' order by regno"
    rsRegNo.Open sqlRegNo, conn, adOpenDynamic, adLockOptimistic
    MSHFlexGrid2.rows = rsRegNo.RecordCount + 1
    
    For i = 0 To rsRegNo.RecordCount - 1
        iTotal = 0
        strRegNo = rsRegNo.Fields(0)
        MSHFlexGrid2.TextMatrix(i + 1, 0) = strRegNo
        rsRegNo.MoveNext
        For j = 1 To GetSubjCount(iSem, iDept, iBatch)
            sqlMarks = "select grade,result from studmarks where semno=" & iSem & " and batch=" & Mid(iBatch, 3, 2) & "and regno='" & strRegNo & "' and subjcode='" & MSHFlexGrid2.TextMatrix(0, j) & "'"
            rsMarks.Open sqlMarks, conn, adOpenDynamic, adLockOptimistic
            If rsMarks.Fields(0) = Null Then
                MSHFlexGrid2.TextMatrix(i + 1, j) = "AB"
            ElseIf rsMarks.Fields(0) = "RA" Or rsMarks.Fields(0) = "U" Or rsMarks.Fields(0) = "W" Or rsMarks.Fields(0) = "I" Then
                MSHFlexGrid2.Row = i + 1
                MSHFlexGrid2.Col = j
                MSHFlexGrid2.TextMatrix(i + 1, j) = rsMarks.Fields(0)
                MSHFlexGrid2.CellForeColor = vbRed
            Else
            
                MSHFlexGrid2.TextMatrix(i + 1, j) = rsMarks.Fields(0)
            End If
            
            rsMarks.Close
        Next
        MSHFlexGrid2.TextMatrix(i + 1, 11) = CalcGPA(strRegNo, CInt(iSem), CInt(iDept), CInt(iBatch))
    Next
    MSHFlexGrid2.Refresh
End Sub
Private Sub rptGradeOrderedLoad()
    On Error Resume Next
    Dim sqlRegNo As String
    Dim sqlMarks As String
    Dim strRegNo As String
    Dim i As Integer
    Dim j As Integer
    Dim rsRegNo As New ADODB.Recordset
    Dim rsMarks As New ADODB.Recordset
    
    rsRegNo.CursorLocation = adUseClient
    rsMarks.CursorLocation = adUseClient
    sqlRegNo = "select regno from (SELECT s1.regno,round(sum(s1.value*s2.credit)/(select sum(credit) FROM subj WHERE batch=" & Mid(iBatch, 3, 2) & " AND semno=" & iSem & " AND dept=" & Department(cmbDept) & "),2) AS GPA,Dense_rank() over (order by (round(sum(s1.value*s2.credit)/(select sum(credit) FROM subj WHERE batch=" & Mid(iBatch, 3, 2) & " AND semno=" & iSem & " AND dept=" & Department(cmbDept) & "),2)) desc) as RANK FROM studmarks s1,subj s2 WHERE s1.batch=" & Mid(iBatch, 3, 2) & " AND s1.semno=" & iSem & " AND s1.subjcode=s2.subjcode GROUP BY s1.regno)"
    rsRegNo.Open sqlRegNo, conn, adOpenDynamic, adLockOptimistic
    MSHFlexGrid2.rows = rsRegNo.RecordCount + 1
    
    For i = 0 To rsRegNo.RecordCount - 1
        iTotal = 0
        strRegNo = rsRegNo.Fields(0)
        MSHFlexGrid2.TextMatrix(i + 1, 0) = strRegNo
        rsRegNo.MoveNext
        For j = 1 To GetSubjCount(iSem, iDept, iBatch)
            sqlMarks = "select grade from studmarks where semno=" & iSem & " and batch=" & Mid(iBatch, 3, 2) & "and regno='" & strRegNo & "' and subjcode='" & MSHFlexGrid2.TextMatrix(0, j) & "'"
            rsMarks.Open sqlMarks, conn, adOpenDynamic, adLockOptimistic
            If rsMarks.Fields(0) = Null Then
                MSHFlexGrid2.TextMatrix(i + 1, j) = "AB"
            ElseIf rsMarks.Fields(0) = "RA" Then
                MSHFlexGrid2.Row = i + 1
                MSHFlexGrid2.Col = j
                MSHFlexGrid2.TextMatrix(i + 1, j) = rsMarks.Fields(0)
                MSHFlexGrid2.CellForeColor = vbRed
            Else
            
                MSHFlexGrid2.TextMatrix(i + 1, j) = rsMarks.Fields(0)
            End If
            rsMarks.Close
        Next
        MSHFlexGrid2.TextMatrix(i + 1, 11) = CalcGPA(strRegNo, CInt(iSem), CInt(iDept), CInt(iBatch))
    Next
    MSHFlexGrid2.Refresh
End Sub
Private Sub rptOrderedLoad()
    On Error Resume Next
    Dim sqlRegNo As String
    Dim sqlMarks As String
    Dim strRegNo As String
    Dim i As Integer
    Dim j As Integer
    Dim iTotal As Integer
    Dim rsRegNo As New ADODB.Recordset
    Dim rsMarks As New ADODB.Recordset
    
    rsRegNo.CursorLocation = adUseClient
    rsMarks.CursorLocation = adUseClient
    sqlRegNo = "select regno,sum((internals+externals)) as Total from studmarks where dept='" & iDept & "' and semno=" & iSem & " and batch=" & Mid(iBatch, 3, 2) & " group by regno order by total desc"
    rsRegNo.Open sqlRegNo, conn, adOpenDynamic, adLockOptimistic
    MSHFlexGrid2.rows = rsRegNo.RecordCount + 1
    
    For i = 0 To rsRegNo.RecordCount - 1
        iTotal = 0
        strRegNo = rsRegNo.Fields(0)
        MSHFlexGrid2.TextMatrix(i + 1, 0) = strRegNo
        rsRegNo.MoveNext
        For j = 1 To GetSubjCount(iSem, iDept, iBatch)
            sqlMarks = "select (internals+externals),result from studmarks where semno=" & iSem & " and batch=" & Mid(iBatch, 3, 2) & "and regno='" & strRegNo & "' and subjcode='" & MSHFlexGrid2.TextMatrix(0, j) & "'"
            rsMarks.Open sqlMarks, conn, adOpenDynamic, adLockOptimistic
            If rsMarks.Fields(1) = "AB" Then
                MSHFlexGrid2.TextMatrix(i + 1, j) = "AB"
            ElseIf rsMarks.Fields(1) = "WH2" Then
                MSHFlexGrid2.TextMatrix(i + 1, j) = "WH2"
            ElseIf rsMarks.Fields(1) = "WH" Then
                MSHFlexGrid2.TextMatrix(i + 1, j) = "WH"
            ElseIf rsMarks.Fields(0) < 50 Then
                MSHFlexGrid2.Row = i + 1
                MSHFlexGrid2.Col = j
                MSHFlexGrid2.TextMatrix(i + 1, j) = rsMarks.Fields(0)
                MSHFlexGrid2.CellForeColor = vbRed
            Else
                MSHFlexGrid2.TextMatrix(i + 1, j) = rsMarks.Fields(0)
            End If
            iTotal = iTotal + rsMarks.Fields(0)
            rsMarks.Close
        Next
        If iTotal = 0 Then
            MSHFlexGrid2.TextMatrix(i + 1, 11) = ""
        Else
            MSHFlexGrid2.TextMatrix(i + 1, 11) = iTotal
        End If
    Next
    MSHFlexGrid2.Refresh
End Sub
Private Sub MSHFlexGrid2_load()
    MSHFlexGrid2.ColWidth(0) = 1200
    MSHFlexGrid2.RowHeightMin = 300
    For icount = 0 To 11
        MSHFlexGrid2.ColAlignment(icount) = flexAlignCenterCenter
        MSHFlexGrid2.ColAlignmentFixed(icount) = flexAlignCenterCenter
        MSHFlexGrid2.ColWidth(icount + 1) = 780
    Next
    MSHFlexGrid2.Refresh
End Sub
Private Sub MSHFlexGrid3_Load()
    MSHFlexGrid3.Clear
    MSHFlexGrid3.ColWidth(0) = 1100
    MSHFlexGrid3.RowHeightMin = 300
    For icount = 0 To 10
        MSHFlexGrid3.ColAlignment(icount) = flexAlignCenterCenter
        MSHFlexGrid3.ColAlignmentFixed(icount) = flexAlignCenterCenter
        MSHFlexGrid3.ColWidth(icount + 1) = 780
    Next
    If iBatch > 2007 Then
        MSHFlexGrid3.TextMatrix(0, 0) = "Appeared"
        MSHFlexGrid3.TextMatrix(1, 0) = "Passed"
        MSHFlexGrid3.TextMatrix(2, 0) = "Failures"
        MSHFlexGrid3.TextMatrix(3, 0) = "S"
        MSHFlexGrid3.TextMatrix(4, 0) = "A"
        MSHFlexGrid3.TextMatrix(5, 0) = "B"
        MSHFlexGrid3.TextMatrix(6, 0) = "C"
        MSHFlexGrid3.TextMatrix(7, 0) = "D"
        MSHFlexGrid3.TextMatrix(8, 0) = "E"
        'MSHFlexGrid3.TextMatrix(9, 0) = ""
        'MSHFlexGrid3.TextMatrix(10, 0) = ""
    Else
        MSHFlexGrid3.TextMatrix(0, 0) = "Appeared"
        MSHFlexGrid3.TextMatrix(1, 0) = "Passed"
        MSHFlexGrid3.TextMatrix(2, 0) = "Failures"
        MSHFlexGrid3.TextMatrix(3, 0) = "50-59"
        MSHFlexGrid3.TextMatrix(4, 0) = "60-69"
        MSHFlexGrid3.TextMatrix(5, 0) = "70-79"
        MSHFlexGrid3.TextMatrix(6, 0) = "80-89"
        MSHFlexGrid3.TextMatrix(7, 0) = ">=90"
        MSHFlexGrid3.TextMatrix(8, 0) = "Min Marks"
        MSHFlexGrid3.TextMatrix(9, 0) = "Max Marks"
        MSHFlexGrid3.TextMatrix(10, 0) = "Average"
    End If
    MSHFlexGrid3.Refresh
End Sub

Private Sub frameStudName_MouseClick(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
    frameStudName.Visible = False
End Sub
Private Sub frameStudName_MouseMove(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
    frameStudName.Visible = False
End Sub
Private Sub fReport_MouseDown(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub



Private Sub MSHFlexGrid2_DblClick()
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
    With MSHFlexGrid2
        StudRegno = .TextMatrix(.MouseRow, 0)
        If .TextMatrix(.MouseRow, 0) = "REGNO" Or .TextMatrix(.MouseRow, 0) = "" Then Exit Sub
    End With
    sql = "select studname from studdetails where regno = '" & StudRegno & "'"
    rsStudName.Open sql, conn, adOpenDynamic, adLockOptimistic
    frameStudName.Caption = "(" & StudRank(StudRegno, iSem, iDept, iBatch, strSec) & ")-" & rsStudName.Fields("studname")
    frameStudName.Move xPos + 500, yPos
    frameStudName.Width = 50
    Me.Refresh
    frameStudName.Visible = True
    Do Until frameStudName.Width > (Len(frameStudName.Caption) * 100 + 100)
        frameStudName.Width = frameStudName.Width + 40
    Loop
End Sub
'Private Sub MSHFlexGrid2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Private Sub MSHFlexGrid2_Click()
    On Error Resume Next
    Dim rsInfo As New ADODB.Recordset
    Dim sql As String
    Dim Subj As String
    lblInfo.Visible = True
    lblPassPercentage.Visible = True
    With MSHFlexGrid2
        Subj = .TextMatrix(0, .MouseCol)
        If .TextMatrix(0, .MouseCol) = "TOTAL" Or .TextMatrix(0, .MouseCol) = "REGNO" Or .TextMatrix(0, .MouseCol) = "" Then
            lblInfo.Caption = ""
            lblPassPercentage.Caption = ""
            Exit Sub
        End If
    End With
    sql = "select count(regno) from studmarks where subjcode = '" & Subj & "' and semno = " & iSem & " and batch=" & Mid(iBatch, 3, 2) & " and dept=" & iDept & " and (internals+externals)<50"
    rsInfo.Open sql, conn, adOpenDynamic, adLockOptimistic
    lblInfo.Caption = "Arrears in " & Subj & ": " & rsInfo.Fields(0)
    lblPassPercentage.Caption = "Pass% In " & Subj & ": " & PassPercentage(Subj, iSem, CInt(iDept), iBatch) & "%"
End Sub

Private Sub btnup_Click()
    btnUp.Visible = False
    btnDown.Visible = True
    Frame1.Top = 4400
    MSHFlexGrid3.Height = 3355
    MSHFlexGrid2.Height = 3240
    Frame1.Height = 3795
End Sub
Private Sub btndown_Click()
    btnUp.Visible = True
    btnDown.Visible = False
    Frame1.Top = 7560
    Frame1.Height = 735
    MSHFlexGrid3.Top = 240
    MSHFlexGrid3.Height = 350
    MSHFlexGrid2.Height = 6300
    Me.Refresh
End Sub

Private Sub Create()
    On Error GoTo Label
    Dim Create As New ADODB.Recordset
    Create.CursorLocation = adUseClient
    sql = "create table temp(regno varchar2(11),clm0 varchar2(10),clm1 varchar2(10),clm2 varchar2(10),clm3 varchar2(10),clm4 varchar2(10),clm5 varchar2(10),clm6 varchar2(10),clm7 varchar2(10),clm8 varchar2(10),clm9 varchar2(10),clm10 varchar2(10))"
    Create.Open sql, conn, adOpenDynamic, adLockOptimistic, -1
    Exit Sub
Label:
    Call Truncate
End Sub
Private Sub Insert()
    On Error Resume Next
    Dim Insert As New ADODB.Recordset
    Insert.CursorLocation = adUseClient
    For i = 1 To MSHFlexGrid2.rows - 1
       sql = "insert into temp values('" & MSHFlexGrid2.TextMatrix(i, 0) & "','" & MSHFlexGrid2.TextMatrix(i, 1) & "','" & MSHFlexGrid2.TextMatrix(i, 2) & "','" & MSHFlexGrid2.TextMatrix(i, 3) & "','" & MSHFlexGrid2.TextMatrix(i, 4) & "','" & MSHFlexGrid2.TextMatrix(i, 5) & "','" & MSHFlexGrid2.TextMatrix(i, 6) & "','" & MSHFlexGrid2.TextMatrix(i, 7) & "','" & MSHFlexGrid2.TextMatrix(i, 8) & "','" & MSHFlexGrid2.TextMatrix(i, 9) & "','" & MSHFlexGrid2.TextMatrix(i, 10) & "','" & MSHFlexGrid2.TextMatrix(i, 11) & "')"
       Insert.Open sql, conn, adOpenDynamic, adLockOptimistic, -1
    Next
    If Err.Number <> 0 Then
        MsgBox Err.Number & "  " & Err.Description
    End If
End Sub
Private Sub Truncate()
    Dim rsTruncate As New ADODB.Recordset
    Dim sql As String
    sql = "truncate table temp"
    rsTruncate.Open sql, conn, adOpenDynamic, adLockOptimistic
End Sub
Private Sub Drop()
    On Error Resume Next
    Dim Drop As New ADODB.Recordset
    Drop.CursorLocation = adUseClient
    sql = "drop table temp"
    Drop.Open sql, conn, adOpenDynamic, adLockOptimistic, -1
    Drop.Close
End Sub
Private Sub Form_Terminate()
    Call Drop
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call Drop
End Sub
Private Sub cmdPrint_Click()
    If iBatch > 2007 Then
        Call rptGradePrint
    Else
        Call rptPrint
    End If
End Sub
Private Sub rptGradePrint()
    Call Create
    Call Insert
    Dim sql As String
    Dim rsPrint As New ADODB.Recordset
    rsPrint.CursorLocation = adUseClient
    sql = "select regno,clm0,clm1,clm2,clm3,clm4,clm5,clm6,clm7,clm8,clm9,clm10 as GPA from temp"
    rsPrint.Open sql, conn, adOpenDynamic, adLockOptimistic
    Set DataReport2.DataSource = rsPrint
    With DataReport2
        With .Sections("Section1")
            .Controls("Text1").DataField = "regno"
            .Controls("Text2").DataField = "clm0"
            .Controls("Text3").DataField = "clm1"
            .Controls("Text4").DataField = "clm2"
            .Controls("Text5").DataField = "clm3"
            .Controls("Text6").DataField = "clm4"
            .Controls("Text7").DataField = "clm5"
            .Controls("Text8").DataField = "clm6"
            .Controls("Text9").DataField = "clm7"
            .Controls("Text10").DataField = "clm8"
            .Controls("Text11").DataField = "clm9"
            .Controls("Text12").DataField = "GPA"
        End With
        With .Sections("Section2")
            .Controls("label1").Caption = MSHFlexGrid2.TextMatrix(0, 0)
            .Controls("label2").Caption = MSHFlexGrid2.TextMatrix(0, 1)
            .Controls("label3").Caption = MSHFlexGrid2.TextMatrix(0, 2)
            .Controls("label4").Caption = MSHFlexGrid2.TextMatrix(0, 3)
            .Controls("label5").Caption = MSHFlexGrid2.TextMatrix(0, 4)
            .Controls("label6").Caption = MSHFlexGrid2.TextMatrix(0, 5)
            .Controls("label7").Caption = MSHFlexGrid2.TextMatrix(0, 6)
            .Controls("label8").Caption = MSHFlexGrid2.TextMatrix(0, 7)
            .Controls("label9").Caption = MSHFlexGrid2.TextMatrix(0, 8)
            .Controls("label10").Caption = MSHFlexGrid2.TextMatrix(0, 9)
            .Controls("label11").Caption = MSHFlexGrid2.TextMatrix(0, 10)
            .Controls("label13").Caption = "Total"
        End With
        
        .Sections("Section3").Controls("label48").Caption = Date & "  " & Time
        
        With .Sections("Section4")
            .Controls("label20").Caption = cmbDept.Text
            .Controls("label22").Caption = cmbSem.Text
        End With
        
        With .Sections("Section5")
            .Controls("label93").Caption = "Grade S:"
            .Controls("label94").Caption = "Grade A:"
            .Controls("label95").Caption = "Grade B:"
            .Controls("label96").Caption = "Grade C:"
            .Controls("label108").Caption = "Grade D:"
            .Controls("label114").Caption = "Grade E:"
            .Controls("label125").Caption = ""
            .Controls("label136").Caption = ""
            
            .Controls("label14").Caption = "No of Students With Arrear:"
            .Controls("label15").Caption = vkLabel3.Caption & "   " & lblArrears(0).Caption
            .Controls("label16").Caption = vkLabel4.Caption & "  " & lblArrears(1).Caption
            .Controls("label17").Caption = vkLabel5.Caption & "  " & lblArrears(2).Caption
            .Controls("label18").Caption = vkLabel6.Caption & "  " & lblArrears(3).Caption
            .Controls("label19").Caption = vkLabel7.Caption & "  " & lblArrears(4).Caption
        
            .Controls("label26").Caption = vkLabel1(0).Caption & "  " & vkLabel2(0).Caption
            .Controls("label27").Caption = vkLabel1(1).Caption & "  " & vkLabel2(1).Caption
            .Controls("label28").Caption = vkLabel1(2).Caption & "  " & vkLabel2(2).Caption
            .Controls("label29").Caption = vkLabel1(3).Caption & "  " & vkLabel2(3).Caption
            .Controls("label30").Caption = vkLabel1(4).Caption & "  " & vkLabel2(4).Caption
            .Controls("label31").Caption = vkLabel1(5).Caption & "  " & vkLabel2(5).Caption
            .Controls("label32").Caption = vkLabel1(6).Caption & "  " & vkLabel2(6).Caption
            .Controls("label33").Caption = vkLabel1(7).Caption & "  " & vkLabel2(7).Caption
            .Controls("label34").Caption = vkLabel1(8).Caption & "  " & vkLabel2(8).Caption
            .Controls("label35").Caption = vkLabel1(9).Caption & "  " & vkLabel2(9).Caption
            
            .Controls("label138").Caption = MSHFlexGrid3.TextMatrix(0, 1)
            .Controls("label139").Caption = MSHFlexGrid3.TextMatrix(0, 2)
            .Controls("label140").Caption = MSHFlexGrid3.TextMatrix(0, 3)
            .Controls("label141").Caption = MSHFlexGrid3.TextMatrix(0, 4)
            .Controls("label142").Caption = MSHFlexGrid3.TextMatrix(0, 5)
            .Controls("label143").Caption = MSHFlexGrid3.TextMatrix(0, 6)
            .Controls("label144").Caption = MSHFlexGrid3.TextMatrix(0, 7)
            .Controls("label145").Caption = MSHFlexGrid3.TextMatrix(0, 8)
            .Controls("label146").Caption = MSHFlexGrid3.TextMatrix(0, 9)
            .Controls("label147").Caption = MSHFlexGrid3.TextMatrix(0, 10)
            
            .Controls("label149").Caption = MSHFlexGrid3.TextMatrix(1, 1)
            .Controls("label150").Caption = MSHFlexGrid3.TextMatrix(1, 2)
            .Controls("label151").Caption = MSHFlexGrid3.TextMatrix(1, 3)
            .Controls("label152").Caption = MSHFlexGrid3.TextMatrix(1, 4)
            .Controls("label153").Caption = MSHFlexGrid3.TextMatrix(1, 5)
            .Controls("label154").Caption = MSHFlexGrid3.TextMatrix(1, 6)
            .Controls("label155").Caption = MSHFlexGrid3.TextMatrix(1, 7)
            .Controls("label156").Caption = MSHFlexGrid3.TextMatrix(1, 8)
            .Controls("label157").Caption = MSHFlexGrid3.TextMatrix(1, 9)
            .Controls("label158").Caption = MSHFlexGrid3.TextMatrix(1, 10)
        
            .Controls("label38").Caption = MSHFlexGrid3.TextMatrix(2, 1)
            .Controls("label39").Caption = MSHFlexGrid3.TextMatrix(2, 2)
            .Controls("label40").Caption = MSHFlexGrid3.TextMatrix(2, 3)
            .Controls("label41").Caption = MSHFlexGrid3.TextMatrix(2, 4)
            .Controls("label42").Caption = MSHFlexGrid3.TextMatrix(2, 5)
            .Controls("label43").Caption = MSHFlexGrid3.TextMatrix(2, 6)
            .Controls("label44").Caption = MSHFlexGrid3.TextMatrix(2, 7)
            .Controls("label45").Caption = MSHFlexGrid3.TextMatrix(2, 8)
            .Controls("label46").Caption = MSHFlexGrid3.TextMatrix(2, 9)
            .Controls("label47").Caption = MSHFlexGrid3.TextMatrix(2, 10)
        
            .Controls("label49").Caption = MSHFlexGrid3.TextMatrix(3, 1)
            .Controls("label50").Caption = MSHFlexGrid3.TextMatrix(3, 2)
            .Controls("label51").Caption = MSHFlexGrid3.TextMatrix(3, 3)
            .Controls("label52").Caption = MSHFlexGrid3.TextMatrix(3, 4)
            .Controls("label53").Caption = MSHFlexGrid3.TextMatrix(3, 5)
            .Controls("label54").Caption = MSHFlexGrid3.TextMatrix(3, 6)
            .Controls("label55").Caption = MSHFlexGrid3.TextMatrix(3, 7)
            .Controls("label56").Caption = MSHFlexGrid3.TextMatrix(3, 8)
            .Controls("label57").Caption = MSHFlexGrid3.TextMatrix(3, 9)
            .Controls("label58").Caption = MSHFlexGrid3.TextMatrix(3, 10)
        
            .Controls("label60").Caption = MSHFlexGrid3.TextMatrix(4, 1)
            .Controls("label61").Caption = MSHFlexGrid3.TextMatrix(4, 2)
            .Controls("label62").Caption = MSHFlexGrid3.TextMatrix(4, 3)
            .Controls("label63").Caption = MSHFlexGrid3.TextMatrix(4, 4)
            .Controls("label64").Caption = MSHFlexGrid3.TextMatrix(4, 5)
            .Controls("label65").Caption = MSHFlexGrid3.TextMatrix(4, 6)
            .Controls("label66").Caption = MSHFlexGrid3.TextMatrix(4, 7)
            .Controls("label67").Caption = MSHFlexGrid3.TextMatrix(4, 8)
            .Controls("label68").Caption = MSHFlexGrid3.TextMatrix(4, 9)
            .Controls("label69").Caption = MSHFlexGrid3.TextMatrix(4, 10)
    
            .Controls("label71").Caption = MSHFlexGrid3.TextMatrix(5, 1)
            .Controls("label72").Caption = MSHFlexGrid3.TextMatrix(5, 2)
            .Controls("label73").Caption = MSHFlexGrid3.TextMatrix(5, 3)
            .Controls("label74").Caption = MSHFlexGrid3.TextMatrix(5, 4)
            .Controls("label75").Caption = MSHFlexGrid3.TextMatrix(5, 5)
            .Controls("label76").Caption = MSHFlexGrid3.TextMatrix(5, 6)
            .Controls("label77").Caption = MSHFlexGrid3.TextMatrix(5, 7)
            .Controls("label78").Caption = MSHFlexGrid3.TextMatrix(5, 8)
            .Controls("label79").Caption = MSHFlexGrid3.TextMatrix(5, 9)
            .Controls("label80").Caption = MSHFlexGrid3.TextMatrix(5, 10)
    
            .Controls("label82").Caption = MSHFlexGrid3.TextMatrix(6, 1)
            .Controls("label83").Caption = MSHFlexGrid3.TextMatrix(6, 2)
            .Controls("label84").Caption = MSHFlexGrid3.TextMatrix(6, 3)
            .Controls("label85").Caption = MSHFlexGrid3.TextMatrix(6, 4)
            .Controls("label86").Caption = MSHFlexGrid3.TextMatrix(6, 5)
            .Controls("label87").Caption = MSHFlexGrid3.TextMatrix(6, 6)
            .Controls("label88").Caption = MSHFlexGrid3.TextMatrix(6, 7)
            .Controls("label89").Caption = MSHFlexGrid3.TextMatrix(6, 8)
            .Controls("label90").Caption = MSHFlexGrid3.TextMatrix(6, 9)
            .Controls("label91").Caption = MSHFlexGrid3.TextMatrix(6, 10)
    
            .Controls("label97").Caption = MSHFlexGrid3.TextMatrix(7, 1)
            .Controls("label98").Caption = MSHFlexGrid3.TextMatrix(7, 2)
            .Controls("label99").Caption = MSHFlexGrid3.TextMatrix(7, 3)
            .Controls("label100").Caption = MSHFlexGrid3.TextMatrix(7, 4)
            .Controls("label101").Caption = MSHFlexGrid3.TextMatrix(7, 5)
            .Controls("label102").Caption = MSHFlexGrid3.TextMatrix(7, 6)
            .Controls("label103").Caption = MSHFlexGrid3.TextMatrix(7, 7)
            .Controls("label104").Caption = MSHFlexGrid3.TextMatrix(7, 8)
            .Controls("label105").Caption = MSHFlexGrid3.TextMatrix(7, 9)
            .Controls("label106").Caption = MSHFlexGrid3.TextMatrix(7, 10)
        
            .Controls("label59").Caption = MSHFlexGrid3.TextMatrix(8, 1)
            .Controls("label70").Caption = MSHFlexGrid3.TextMatrix(8, 2)
            .Controls("label81").Caption = MSHFlexGrid3.TextMatrix(8, 3)
            .Controls("label92").Caption = MSHFlexGrid3.TextMatrix(8, 4)
            .Controls("label107").Caption = MSHFlexGrid3.TextMatrix(8, 5)
            .Controls("label109").Caption = MSHFlexGrid3.TextMatrix(8, 6)
            .Controls("label110").Caption = MSHFlexGrid3.TextMatrix(8, 7)
            .Controls("label111").Caption = MSHFlexGrid3.TextMatrix(8, 8)
            .Controls("label112").Caption = MSHFlexGrid3.TextMatrix(8, 9)
            .Controls("label113").Caption = MSHFlexGrid3.TextMatrix(8, 10)
            
            .Controls("label115").Caption = MSHFlexGrid3.TextMatrix(9, 1)
            .Controls("label116").Caption = MSHFlexGrid3.TextMatrix(9, 2)
            .Controls("label117").Caption = MSHFlexGrid3.TextMatrix(9, 3)
            .Controls("label118").Caption = MSHFlexGrid3.TextMatrix(9, 4)
            .Controls("label119").Caption = MSHFlexGrid3.TextMatrix(9, 5)
            .Controls("label120").Caption = MSHFlexGrid3.TextMatrix(9, 6)
            .Controls("label121").Caption = MSHFlexGrid3.TextMatrix(9, 7)
            .Controls("label122").Caption = MSHFlexGrid3.TextMatrix(9, 8)
            .Controls("label123").Caption = MSHFlexGrid3.TextMatrix(9, 9)
            .Controls("label124").Caption = MSHFlexGrid3.TextMatrix(9, 10)
        
            .Controls("label126").Caption = MSHFlexGrid3.TextMatrix(10, 1)
            .Controls("label127").Caption = MSHFlexGrid3.TextMatrix(10, 2)
            .Controls("label128").Caption = MSHFlexGrid3.TextMatrix(10, 3)
            .Controls("label129").Caption = MSHFlexGrid3.TextMatrix(10, 4)
            .Controls("label130").Caption = MSHFlexGrid3.TextMatrix(10, 5)
            .Controls("label131").Caption = MSHFlexGrid3.TextMatrix(10, 6)
            .Controls("label132").Caption = MSHFlexGrid3.TextMatrix(10, 7)
            .Controls("label133").Caption = MSHFlexGrid3.TextMatrix(10, 8)
            .Controls("label134").Caption = MSHFlexGrid3.TextMatrix(10, 9)
            .Controls("label135").Caption = MSHFlexGrid3.TextMatrix(10, 10)
        End With
               
    .LeftMargin = 100
    .RightMargin = 100
    .Caption = "Marks frmReport"
    .Show
    '.ExportReport rptKeyHTML, App.Path & "\Reports\" & "Report" & "(" & cmbSem.Text & ")" & ".html", True, False, rptRangeAllPages
    End With
End Sub

Private Sub rptPrint()
    On Error Resume Next
    Call Create
    Call Insert
    Dim strTotal As String
    Dim sql As String
    strTotal = "(clm0"
    For i = 1 To GetSubjCount(iSem, iDept, iBatch) - 1
        strTotal = strTotal & " + " & "clm" & i
    Next
    strTotal = strTotal & ")"
    Dim rsPrint As New ADODB.Recordset
    rsPrint.CursorLocation = adUseClient
    sql = "select regno,clm0,clm1,clm2,clm3,clm4,clm5,clm6,clm7,clm8,clm9,clm10 as Total from temp"
    rsPrint.Open sql, conn, adOpenDynamic, adLockOptimistic
    Set DataReport2.DataSource = rsPrint
    With DataReport2
        With .Sections("Section1")
            .Controls("Text1").DataField = "regno"
            .Controls("Text2").DataField = "clm0"
            .Controls("Text3").DataField = "clm1"
            .Controls("Text4").DataField = "clm2"
            .Controls("Text5").DataField = "clm3"
            .Controls("Text6").DataField = "clm4"
            .Controls("Text7").DataField = "clm5"
            .Controls("Text8").DataField = "clm6"
            .Controls("Text9").DataField = "clm7"
            .Controls("Text10").DataField = "clm8"
            .Controls("Text11").DataField = "clm9"
            .Controls("Text12").DataField = "Total"
        End With
        With .Sections("Section2")
            .Controls("label1").Caption = MSHFlexGrid2.TextMatrix(0, 0)
            .Controls("label2").Caption = MSHFlexGrid2.TextMatrix(0, 1)
            .Controls("label3").Caption = MSHFlexGrid2.TextMatrix(0, 2)
            .Controls("label4").Caption = MSHFlexGrid2.TextMatrix(0, 3)
            .Controls("label5").Caption = MSHFlexGrid2.TextMatrix(0, 4)
            .Controls("label6").Caption = MSHFlexGrid2.TextMatrix(0, 5)
            .Controls("label7").Caption = MSHFlexGrid2.TextMatrix(0, 6)
            .Controls("label8").Caption = MSHFlexGrid2.TextMatrix(0, 7)
            .Controls("label9").Caption = MSHFlexGrid2.TextMatrix(0, 8)
            .Controls("label10").Caption = MSHFlexGrid2.TextMatrix(0, 9)
            .Controls("label11").Caption = MSHFlexGrid2.TextMatrix(0, 10)
            .Controls("label13").Caption = "Total"
        End With
        
        .Sections("Section3").Controls("label48").Caption = Date & "  " & Time
        
        With .Sections("Section4")
            .Controls("label20").Caption = cmbDept.Text
            .Controls("label22").Caption = cmbSem.Text
        End With
        
        With .Sections("Section5")
            .Controls("label14").Caption = "No of Students With Arrear:"
            .Controls("label15").Caption = vkLabel3.Caption & "   " & lblArrears(0).Caption
            .Controls("label16").Caption = vkLabel4.Caption & "  " & lblArrears(1).Caption
            .Controls("label17").Caption = vkLabel5.Caption & "  " & lblArrears(2).Caption
            .Controls("label18").Caption = vkLabel6.Caption & "  " & lblArrears(3).Caption
            .Controls("label19").Caption = vkLabel7.Caption & "  " & lblArrears(4).Caption
        
            .Controls("label26").Caption = vkLabel1(0).Caption & "  " & vkLabel2(0).Caption
            .Controls("label27").Caption = vkLabel1(1).Caption & "  " & vkLabel2(1).Caption
            .Controls("label28").Caption = vkLabel1(2).Caption & "  " & vkLabel2(2).Caption
            .Controls("label29").Caption = vkLabel1(3).Caption & "  " & vkLabel2(3).Caption
            .Controls("label30").Caption = vkLabel1(4).Caption & "  " & vkLabel2(4).Caption
            .Controls("label31").Caption = vkLabel1(5).Caption & "  " & vkLabel2(5).Caption
            .Controls("label32").Caption = vkLabel1(6).Caption & "  " & vkLabel2(6).Caption
            .Controls("label33").Caption = vkLabel1(7).Caption & "  " & vkLabel2(7).Caption
            .Controls("label34").Caption = vkLabel1(8).Caption & "  " & vkLabel2(8).Caption
            .Controls("label35").Caption = vkLabel1(9).Caption & "  " & vkLabel2(9).Caption
            
            .Controls("label138").Caption = MSHFlexGrid3.TextMatrix(0, 1)
            .Controls("label139").Caption = MSHFlexGrid3.TextMatrix(0, 2)
            .Controls("label140").Caption = MSHFlexGrid3.TextMatrix(0, 3)
            .Controls("label141").Caption = MSHFlexGrid3.TextMatrix(0, 4)
            .Controls("label142").Caption = MSHFlexGrid3.TextMatrix(0, 5)
            .Controls("label143").Caption = MSHFlexGrid3.TextMatrix(0, 6)
            .Controls("label144").Caption = MSHFlexGrid3.TextMatrix(0, 7)
            .Controls("label145").Caption = MSHFlexGrid3.TextMatrix(0, 8)
            .Controls("label146").Caption = MSHFlexGrid3.TextMatrix(0, 9)
            .Controls("label147").Caption = MSHFlexGrid3.TextMatrix(0, 10)
            
            .Controls("label149").Caption = MSHFlexGrid3.TextMatrix(1, 1)
            .Controls("label150").Caption = MSHFlexGrid3.TextMatrix(1, 2)
            .Controls("label151").Caption = MSHFlexGrid3.TextMatrix(1, 3)
            .Controls("label152").Caption = MSHFlexGrid3.TextMatrix(1, 4)
            .Controls("label153").Caption = MSHFlexGrid3.TextMatrix(1, 5)
            .Controls("label154").Caption = MSHFlexGrid3.TextMatrix(1, 6)
            .Controls("label155").Caption = MSHFlexGrid3.TextMatrix(1, 7)
            .Controls("label156").Caption = MSHFlexGrid3.TextMatrix(1, 8)
            .Controls("label157").Caption = MSHFlexGrid3.TextMatrix(1, 9)
            .Controls("label158").Caption = MSHFlexGrid3.TextMatrix(1, 10)
        
            .Controls("label38").Caption = MSHFlexGrid3.TextMatrix(2, 1)
            .Controls("label39").Caption = MSHFlexGrid3.TextMatrix(2, 2)
            .Controls("label40").Caption = MSHFlexGrid3.TextMatrix(2, 3)
            .Controls("label41").Caption = MSHFlexGrid3.TextMatrix(2, 4)
            .Controls("label42").Caption = MSHFlexGrid3.TextMatrix(2, 5)
            .Controls("label43").Caption = MSHFlexGrid3.TextMatrix(2, 6)
            .Controls("label44").Caption = MSHFlexGrid3.TextMatrix(2, 7)
            .Controls("label45").Caption = MSHFlexGrid3.TextMatrix(2, 8)
            .Controls("label46").Caption = MSHFlexGrid3.TextMatrix(2, 9)
            .Controls("label47").Caption = MSHFlexGrid3.TextMatrix(2, 10)
        
            .Controls("label49").Caption = MSHFlexGrid3.TextMatrix(3, 1)
            .Controls("label50").Caption = MSHFlexGrid3.TextMatrix(3, 2)
            .Controls("label51").Caption = MSHFlexGrid3.TextMatrix(3, 3)
            .Controls("label52").Caption = MSHFlexGrid3.TextMatrix(3, 4)
            .Controls("label53").Caption = MSHFlexGrid3.TextMatrix(3, 5)
            .Controls("label54").Caption = MSHFlexGrid3.TextMatrix(3, 6)
            .Controls("label55").Caption = MSHFlexGrid3.TextMatrix(3, 7)
            .Controls("label56").Caption = MSHFlexGrid3.TextMatrix(3, 8)
            .Controls("label57").Caption = MSHFlexGrid3.TextMatrix(3, 9)
            .Controls("label58").Caption = MSHFlexGrid3.TextMatrix(3, 10)
        
            .Controls("label60").Caption = MSHFlexGrid3.TextMatrix(4, 1)
            .Controls("label61").Caption = MSHFlexGrid3.TextMatrix(4, 2)
            .Controls("label62").Caption = MSHFlexGrid3.TextMatrix(4, 3)
            .Controls("label63").Caption = MSHFlexGrid3.TextMatrix(4, 4)
            .Controls("label64").Caption = MSHFlexGrid3.TextMatrix(4, 5)
            .Controls("label65").Caption = MSHFlexGrid3.TextMatrix(4, 6)
            .Controls("label66").Caption = MSHFlexGrid3.TextMatrix(4, 7)
            .Controls("label67").Caption = MSHFlexGrid3.TextMatrix(4, 8)
            .Controls("label68").Caption = MSHFlexGrid3.TextMatrix(4, 9)
            .Controls("label69").Caption = MSHFlexGrid3.TextMatrix(4, 10)
    
            .Controls("label71").Caption = MSHFlexGrid3.TextMatrix(5, 1)
            .Controls("label72").Caption = MSHFlexGrid3.TextMatrix(5, 2)
            .Controls("label73").Caption = MSHFlexGrid3.TextMatrix(5, 3)
            .Controls("label74").Caption = MSHFlexGrid3.TextMatrix(5, 4)
            .Controls("label75").Caption = MSHFlexGrid3.TextMatrix(5, 5)
            .Controls("label76").Caption = MSHFlexGrid3.TextMatrix(5, 6)
            .Controls("label77").Caption = MSHFlexGrid3.TextMatrix(5, 7)
            .Controls("label78").Caption = MSHFlexGrid3.TextMatrix(5, 8)
            .Controls("label79").Caption = MSHFlexGrid3.TextMatrix(5, 9)
            .Controls("label80").Caption = MSHFlexGrid3.TextMatrix(5, 10)
    
            .Controls("label82").Caption = MSHFlexGrid3.TextMatrix(6, 1)
            .Controls("label83").Caption = MSHFlexGrid3.TextMatrix(6, 2)
            .Controls("label84").Caption = MSHFlexGrid3.TextMatrix(6, 3)
            .Controls("label85").Caption = MSHFlexGrid3.TextMatrix(6, 4)
            .Controls("label86").Caption = MSHFlexGrid3.TextMatrix(6, 5)
            .Controls("label87").Caption = MSHFlexGrid3.TextMatrix(6, 6)
            .Controls("label88").Caption = MSHFlexGrid3.TextMatrix(6, 7)
            .Controls("label89").Caption = MSHFlexGrid3.TextMatrix(6, 8)
            .Controls("label90").Caption = MSHFlexGrid3.TextMatrix(6, 9)
            .Controls("label91").Caption = MSHFlexGrid3.TextMatrix(6, 10)
    
            .Controls("label97").Caption = MSHFlexGrid3.TextMatrix(7, 1)
            .Controls("label98").Caption = MSHFlexGrid3.TextMatrix(7, 2)
            .Controls("label99").Caption = MSHFlexGrid3.TextMatrix(7, 3)
            .Controls("label100").Caption = MSHFlexGrid3.TextMatrix(7, 4)
            .Controls("label101").Caption = MSHFlexGrid3.TextMatrix(7, 5)
            .Controls("label102").Caption = MSHFlexGrid3.TextMatrix(7, 6)
            .Controls("label103").Caption = MSHFlexGrid3.TextMatrix(7, 7)
            .Controls("label104").Caption = MSHFlexGrid3.TextMatrix(7, 8)
            .Controls("label105").Caption = MSHFlexGrid3.TextMatrix(7, 9)
            .Controls("label106").Caption = MSHFlexGrid3.TextMatrix(7, 10)
        
            .Controls("label59").Caption = MSHFlexGrid3.TextMatrix(8, 1)
            .Controls("label70").Caption = MSHFlexGrid3.TextMatrix(8, 2)
            .Controls("label81").Caption = MSHFlexGrid3.TextMatrix(8, 3)
            .Controls("label92").Caption = MSHFlexGrid3.TextMatrix(8, 4)
            .Controls("label107").Caption = MSHFlexGrid3.TextMatrix(8, 5)
            .Controls("label109").Caption = MSHFlexGrid3.TextMatrix(8, 6)
            .Controls("label110").Caption = MSHFlexGrid3.TextMatrix(8, 7)
            .Controls("label111").Caption = MSHFlexGrid3.TextMatrix(8, 8)
            .Controls("label112").Caption = MSHFlexGrid3.TextMatrix(8, 9)
            .Controls("label113").Caption = MSHFlexGrid3.TextMatrix(8, 10)
            
            .Controls("label115").Caption = MSHFlexGrid3.TextMatrix(9, 1)
            .Controls("label116").Caption = MSHFlexGrid3.TextMatrix(9, 2)
            .Controls("label117").Caption = MSHFlexGrid3.TextMatrix(9, 3)
            .Controls("label118").Caption = MSHFlexGrid3.TextMatrix(9, 4)
            .Controls("label119").Caption = MSHFlexGrid3.TextMatrix(9, 5)
            .Controls("label120").Caption = MSHFlexGrid3.TextMatrix(9, 6)
            .Controls("label121").Caption = MSHFlexGrid3.TextMatrix(9, 7)
            .Controls("label122").Caption = MSHFlexGrid3.TextMatrix(9, 8)
            .Controls("label123").Caption = MSHFlexGrid3.TextMatrix(9, 9)
            .Controls("label124").Caption = MSHFlexGrid3.TextMatrix(9, 10)
        
            .Controls("label126").Caption = MSHFlexGrid3.TextMatrix(10, 1)
            .Controls("label127").Caption = MSHFlexGrid3.TextMatrix(10, 2)
            .Controls("label128").Caption = MSHFlexGrid3.TextMatrix(10, 3)
            .Controls("label129").Caption = MSHFlexGrid3.TextMatrix(10, 4)
            .Controls("label130").Caption = MSHFlexGrid3.TextMatrix(10, 5)
            .Controls("label131").Caption = MSHFlexGrid3.TextMatrix(10, 6)
            .Controls("label132").Caption = MSHFlexGrid3.TextMatrix(10, 7)
            .Controls("label133").Caption = MSHFlexGrid3.TextMatrix(10, 8)
            .Controls("label134").Caption = MSHFlexGrid3.TextMatrix(10, 9)
            .Controls("label135").Caption = MSHFlexGrid3.TextMatrix(10, 10)
        End With
               
    .LeftMargin = 100
    .RightMargin = 100
    .Caption = "Marks frmReport"
    .Show
    '.ExportReport rptKeyHTML, App.Path & "\Reports\" & "Report" & "(" & cmbSem.Text & ")" & ".html", True, False, rptRangeAllPages
    End With
End Sub


