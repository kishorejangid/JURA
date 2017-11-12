VERSION 5.00
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkusercontrolsxp.ocx"
Begin VB.Form frmEnterMarks 
   BorderStyle     =   0  'None
   Caption         =   "Enter Marks"
   ClientHeight    =   8325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7005
   Icon            =   "frmEnterMarks.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8325
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   Begin vkUserContolsXP.vkCommand cmdInsert 
      Height          =   495
      Left            =   720
      TabIndex        =   32
      Top             =   7560
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   873
      Caption         =   "Insert"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   33023
      CustomStyle     =   0
   End
   Begin vkUserContolsXP.vkLabel lblExternals 
      Height          =   255
      Left            =   4800
      TabIndex        =   31
      Top             =   2280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   "Externals:"
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
   Begin vkUserContolsXP.vkFrame fEnterMarks 
      Height          =   8325
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   7020
      _ExtentX        =   12383
      _ExtentY        =   14684
      Caption         =   "Enter Marks"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TitleColor1     =   33023
      TitleColor2     =   8438015
      TitleGradient   =   2
      TitleHeight     =   300
      BorderColor     =   33023
      BorderWidth     =   2
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
         Left            =   1320
         TabIndex        =   45
         Top             =   1080
         Width           =   2175
      End
      Begin vkUserContolsXP.vkLabel vkLabel8 
         Height          =   255
         Left            =   360
         TabIndex        =   44
         Top             =   1200
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
      Begin JURA.StylerButton cmdClose 
         Height          =   255
         Left            =   6480
         TabIndex        =   43
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
      Begin vkUserContolsXP.vkTextBox txtSubjcode 
         Height          =   375
         Index           =   9
         Left            =   720
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   6960
         Width           =   1335
         _ExtentX        =   2355
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox txtSubjcode 
         Height          =   375
         Index           =   8
         Left            =   720
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   6480
         Width           =   1335
         _ExtentX        =   2355
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox txtSubjcode 
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1335
         _ExtentX        =   2355
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox txtSubjcode 
         Height          =   375
         Index           =   7
         Left            =   720
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   6000
         Width           =   1335
         _ExtentX        =   2355
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox txtSubjcode 
         Height          =   375
         Index           =   6
         Left            =   720
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   5520
         Width           =   1335
         _ExtentX        =   2355
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox txtSubjcode 
         Height          =   375
         Index           =   5
         Left            =   720
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   5040
         Width           =   1335
         _ExtentX        =   2355
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox txtSubjcode 
         Height          =   375
         Index           =   4
         Left            =   720
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   4560
         Width           =   1335
         _ExtentX        =   2355
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox txtSubjcode 
         Height          =   375
         Index           =   3
         Left            =   720
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   4080
         Width           =   1335
         _ExtentX        =   2355
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox txtSubjcode 
         Height          =   375
         Index           =   2
         Left            =   720
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   3600
         Width           =   1335
         _ExtentX        =   2355
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox txtSubjcode 
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1335
         _ExtentX        =   2355
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox txtExternals 
         Height          =   375
         Index           =   9
         Left            =   4920
         TabIndex        =   23
         Top             =   6960
         Width           =   1335
         _ExtentX        =   2355
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
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox txtExternals 
         Height          =   375
         Index           =   8
         Left            =   4920
         TabIndex        =   21
         Top             =   6480
         Width           =   1335
         _ExtentX        =   2355
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
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox txtExternals 
         Height          =   375
         Index           =   7
         Left            =   4920
         TabIndex        =   19
         Top             =   6000
         Width           =   1335
         _ExtentX        =   2355
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
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox txtExternals 
         Height          =   375
         Index           =   6
         Left            =   4920
         TabIndex        =   17
         Top             =   5520
         Width           =   1335
         _ExtentX        =   2355
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
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox txtExternals 
         Height          =   375
         Index           =   5
         Left            =   4920
         TabIndex        =   15
         Top             =   5040
         Width           =   1335
         _ExtentX        =   2355
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
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox txtExternals 
         Height          =   375
         Index           =   4
         Left            =   4920
         TabIndex        =   13
         Top             =   4560
         Width           =   1335
         _ExtentX        =   2355
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
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox txtExternals 
         Height          =   375
         Index           =   3
         Left            =   4920
         TabIndex        =   11
         Top             =   4080
         Width           =   1335
         _ExtentX        =   2355
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
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox txtExternals 
         Height          =   375
         Index           =   2
         Left            =   4920
         TabIndex        =   9
         Top             =   3600
         Width           =   1335
         _ExtentX        =   2355
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
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox txtExternals 
         Height          =   375
         Index           =   1
         Left            =   4920
         TabIndex        =   7
         Top             =   3120
         Width           =   1335
         _ExtentX        =   2355
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
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox txtExternals 
         Height          =   375
         Index           =   0
         Left            =   4920
         TabIndex        =   5
         Top             =   2640
         Width           =   1335
         _ExtentX        =   2355
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
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox txtInternals 
         Height          =   375
         Index           =   9
         Left            =   2880
         TabIndex        =   22
         Top             =   6960
         Width           =   1335
         _ExtentX        =   2355
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
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox txtInternals 
         Height          =   375
         Index           =   8
         Left            =   2880
         TabIndex        =   20
         Top             =   6480
         Width           =   1335
         _ExtentX        =   2355
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
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox txtInternals 
         Height          =   375
         Index           =   7
         Left            =   2880
         TabIndex        =   18
         Top             =   6000
         Width           =   1335
         _ExtentX        =   2355
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
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox txtInternals 
         Height          =   375
         Index           =   6
         Left            =   2880
         TabIndex        =   16
         Top             =   5520
         Width           =   1335
         _ExtentX        =   2355
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
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox txtInternals 
         Height          =   375
         Index           =   5
         Left            =   2880
         TabIndex        =   14
         Top             =   5040
         Width           =   1335
         _ExtentX        =   2355
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
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox txtInternals 
         Height          =   375
         Index           =   4
         Left            =   2880
         TabIndex        =   12
         Top             =   4560
         Width           =   1335
         _ExtentX        =   2355
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
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox txtInternals 
         Height          =   375
         Index           =   3
         Left            =   2880
         TabIndex        =   10
         Top             =   4080
         Width           =   1335
         _ExtentX        =   2355
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
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox txtInternals 
         Height          =   375
         Index           =   2
         Left            =   2880
         TabIndex        =   8
         Top             =   3600
         Width           =   1335
         _ExtentX        =   2355
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
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox txtInternals 
         Height          =   375
         Index           =   1
         Left            =   2880
         TabIndex        =   6
         Top             =   3120
         Width           =   1335
         _ExtentX        =   2355
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
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox txtInternals 
         Height          =   375
         Index           =   0
         Left            =   2880
         TabIndex        =   4
         Top             =   2640
         Width           =   1335
         _ExtentX        =   2355
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
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkLabel vkLabel6 
         Height          =   255
         Left            =   2760
         TabIndex        =   30
         Top             =   2280
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Internals:"
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
         Height          =   375
         Left            =   360
         TabIndex        =   29
         Top             =   2280
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Subject Code:"
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
         Height          =   375
         Left            =   360
         TabIndex        =   28
         Top             =   1680
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Register No:"
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
      Begin vkUserContolsXP.vkTextBox txtName 
         Height          =   375
         Left            =   4440
         TabIndex        =   27
         Top             =   1560
         Width           =   2175
         _ExtentX        =   3836
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkLabel vkLabel4 
         Height          =   255
         Left            =   3600
         TabIndex        =   26
         Top             =   1680
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
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
         Left            =   1320
         TabIndex        =   3
         Top             =   1560
         Width           =   2175
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
         Left            =   4440
         TabIndex        =   2
         Top             =   1080
         Width           =   2175
      End
      Begin vkUserContolsXP.vkLabel vkLabel2 
         Height          =   255
         Left            =   3600
         TabIndex        =   25
         Top             =   1200
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   1320
         TabIndex        =   1
         Top             =   600
         Width           =   5295
      End
      Begin vkUserContolsXP.vkLabel vkLabel1 
         Height          =   255
         Left            =   360
         TabIndex        =   0
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
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
   End
End
Attribute VB_Name = "frmEnterMarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iSubjCount As Integer

Private Sub cmbBatch_Click()
    iBatch = cmbBatch.Text
    Call cmbRegNo_Load(cmbRegNo)
    Call cmbRegNo_Click
    Call SubjCode_Load
    iSubjCount = GetSubjCount(iSem, iDept, iBatch) - 1
    If cmbBatch.Text > 2007 Then
        lblExternals.Caption = "Grade"
    Else
        lblExternals.Caption = "Externals"
    End If
End Sub

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub cmdInsert_Click()
    On Error Resume Next
    Dim rs As New ADODB.Recordset
    Dim qr As String
    Dim i As Integer
    Dim j As Integer
    For i = 0 To iSubjCount
        If txtInternals(i).Text = "" Or txtExternals(i).Text = "" Then
             MsgBox "Some Fields are Empty"
             Exit Sub
        End If
    Next
    rs.CursorLocation = adUseClient
    If iBatch > 2007 Then
        For i = 0 To iSubjCount
            qr = "insert into studmarks(REGNO,SEMNO,DEPT,BATCH,SUBJCODE,INTERNALS,GRADE) values('" & cmbRegNo.Text & "','" & cmbSem.Text & "'," & iDept & "," & Mid(cmbRegNo.Text, 4, 2) & ",'" & txtSubjcode(i).Text & "','" & txtInternals(i).Text & "','" & txtExternals(i).Text & "')"
            rs.Open qr, conn, adOpenDynamic, adLockOptimistic, 1
        Next i
    Else
        For i = 0 To iSubjCount
            qr = "insert into studmarks(REGNO,SEMNO,DEPT,BATCH,SUBJCODE,INTERNALS,EXTERNALS) values('" & cmbRegNo.Text & "','" & cmbSem.Text & "'," & iDept & "," & Mid(cmbRegNo.Text, 4, 2) & ",'" & txtSubjcode(i).Text & "','" & txtInternals(i).Text & "','" & txtExternals(i).Text & "')"
            rs.Open qr, conn, adOpenDynamic, adLockOptimistic, 1
        Next i
    End If
    If Err.Number <> 0 Then
        MsgBox Error & vbCrLf & "Error Number: " & Err.Number
    Else
        MsgBox "Inserted"
    End If
    For j = 0 To iSubjCount
        txtInternals(j).Text = ""
        txtExternals(j).Text = ""
    Next
End Sub





Private Sub fEnterMarks_MouseDown(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_Load()
    'Me.Top = mdiMain.Top + 300
    'Me.Left = mdiMain.Left + 1000
    CreateRoundRectFromWindow Me, 7, 7
    Me.Top = (mdiMain.Height - Me.Height) / 5
    Me.Left = (mdiMain.Width - Me.Width) / 2
    Me.BackColor = mdiMain.BackColor
    Call frmColor(frmEnterMarks)
    
    Dim i As Integer
    For i = 0 To 9
        txtSubjcode(i).Visible = False
        txtInternals(i).Visible = False
        txtExternals(i).Visible = False
    Next
    
    Call cmbDept_Load(cmbDept)
    Call cmbSem_Load(cmbSem)
    iSem = cmbSem.Text
    Call cmbBatch_Load(cmbBatch)
    
    iSubjCount = GetSubjCount(iSem, iDept, iBatch) - 1
    For i = 0 To iSubjCount
        txtSubjcode(i).Visible = True
        txtInternals(i).Visible = True
        txtExternals(i).Visible = True
    Next
    
    Call cmbRegNo_Load(cmbRegNo)
    Call cmbRegNo_Click
    Call SubjCode_Load
End Sub
Private Sub cmbDept_Click()
    Dim i As Integer
    For i = 0 To iSubjCount
        txtSubjcode(i).Text = ""
    Next
    iDept = Department(cmbDept)
    Call cmbRegNo_Load(cmbRegNo)
    Call cmbRegNo_Click
    Call SubjCode_Load
    iSubjCount = GetSubjCount(iSem, iDept, iBatch) - 1
End Sub
Private Sub cmbDept_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmbDept_Click
End Sub
Private Sub cmbSem_Click()
    Dim i As Integer
    iSem = cmbSem.Text
    iSubjCount = GetSubjCount(iSem, iDept, iBatch) - 1
    For i = 0 To 9
        txtSubjcode(i).Visible = False
        txtInternals(i).Visible = False
        txtExternals(i).Visible = False
    Next
    For i = 0 To iSubjCount
        txtSubjcode(i).Visible = True
        txtInternals(i).Visible = True
        txtExternals(i).Visible = True
    Next
    Call SubjCode_Load
End Sub
Private Sub cmbSem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmbSem_Click
End Sub
Private Sub cmbRegNo_Click()
    On Error Resume Next
    Dim rs As New ADODB.Recordset
    Dim reg As String
    Dim sql As String
    reg = cmbRegNo.Text
    rs.CursorLocation = adUseClient
    sql = "select studname from studdetails where regno = '" & reg & "'"
    rs.Open sql, conn, adOpenDynamic, adLockOptimistic, -1
    txtName.Text = rs.Fields("studname")
    txtName.Refresh
    Exit Sub
End Sub
Private Sub cmbRegNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmbRegNo_Click
End Sub
Private Sub SubjCode_Load()
    On Error Resume Next
    Dim rs As New ADODB.Recordset
    Dim sql As String
    Dim i As Integer
    rs.CursorLocation = adUseClient
    sql = "select subjcode from subj where dept = '" & iDept & "' and semno = " & iSem & " and batch = " & Mid(iBatch, 3, 2) & " order by subjcode"
    rs.Open sql, conn, adOpenDynamic, adLockOptimistic, -1
    For i = 0 To iSubjCount
        txtSubjcode(i).Text = rs.Fields("subjcode")
        rs.MoveNext
    Next i
End Sub

