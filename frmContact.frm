VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkusercontrolsxp.ocx"
Begin VB.Form frmContact 
   BorderStyle     =   0  'None
   Caption         =   "Contact Information"
   ClientHeight    =   6720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9825
   Icon            =   "frmContact.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   9825
   ShowInTaskbar   =   0   'False
   Begin JURA.StylerButton cmdClose 
      Height          =   255
      Left            =   9240
      TabIndex        =   39
      Top             =   0
      Width           =   375
      _extentx        =   661
      _extenty        =   450
      caption         =   "X"
      captiondisablecolor=   12236471
      captioneffectcolor=   16777215
      focusdottedrect =   0   'False
      font            =   "frmContact.frx":57E2
      roundedvalue    =   1
   End
   Begin JURA.StylerButton cmdMin 
      Height          =   255
      Left            =   9000
      TabIndex        =   38
      Top             =   0
      Width           =   255
      _extentx        =   450
      _extenty        =   450
      caption         =   "-"
      captiondisablecolor=   12236471
      captioneffectcolor=   16777215
      focusdottedrect =   0   'False
      font            =   "frmContact.frx":580E
      roundedvalue    =   1
   End
   Begin vkUserContolsXP.vkFrame fContact 
      Height          =   6705
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   9810
      _ExtentX        =   17304
      _ExtentY        =   11827
      Caption         =   "Contact"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleColor1     =   33023
      TitleColor2     =   8438015
      TitleGradient   =   2
      TitleHeight     =   360
      BorderColor     =   33023
      RoundAngle      =   5
      BorderWidth     =   2
      Begin vkUserContolsXP.vkLabel lblName 
         Height          =   195
         Left            =   240
         TabIndex        =   37
         Top             =   1680
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   344
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "&Name:"
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
      Begin vkUserContolsXP.vkLabel lblDoB 
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   2160
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   344
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "D.O.&B:"
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
      Begin vkUserContolsXP.vkLabel lblFather 
         Height          =   195
         Left            =   240
         TabIndex        =   35
         Top             =   2640
         Width           =   540
         _ExtentX        =   873
         _ExtentY        =   344
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "&Father:"
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
      Begin vkUserContolsXP.vkLabel lblMother 
         Height          =   195
         Left            =   240
         TabIndex        =   34
         Top             =   3120
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   344
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "&Mother:"
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
      Begin vkUserContolsXP.vkLabel lblRegNo 
         Height          =   315
         Left            =   240
         TabIndex        =   33
         Top             =   1200
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   556
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "&Reg No:"
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
      Begin vkUserContolsXP.vkLabel lblOccupation 
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   3600
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   344
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "&Occupation:"
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
      Begin vkUserContolsXP.vkLabel lblAddress 
         Height          =   195
         Left            =   240
         TabIndex        =   31
         Top             =   4080
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   344
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Address:"
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
      Begin vkUserContolsXP.vkLabel lblCity 
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   4680
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   344
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "City:"
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
      Begin vkUserContolsXP.vkLabel lblPincode 
         Height          =   195
         Left            =   3480
         TabIndex        =   29
         Top             =   4680
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   344
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "PinCode:"
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
      Begin vkUserContolsXP.vkLabel lblState 
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   5160
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   344
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "State:"
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
      Begin vkUserContolsXP.vkLabel lblEmail 
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   5640
         Width           =   540
         _ExtentX        =   847
         _ExtentY        =   344
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "E-Mail:"
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
      Begin vkUserContolsXP.vkLabel lblPhone 
         Height          =   195
         Left            =   5640
         TabIndex        =   26
         Top             =   4680
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   344
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Phone:"
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
      Begin vkUserContolsXP.vkLabel lblMobile 
         Height          =   195
         Left            =   5640
         TabIndex        =   25
         Top             =   5160
         Width           =   540
         _ExtentX        =   900
         _ExtentY        =   344
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Mobile:"
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
      Begin vkUserContolsXP.vkLabel lblGender 
         Height          =   195
         Left            =   3840
         TabIndex        =   24
         Top             =   2280
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   344
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Gender:"
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
      Begin VB.PictureBox Picture 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3135
         Left            =   6480
         ScaleHeight     =   3105
         ScaleWidth      =   2985
         TabIndex        =   23
         Top             =   720
         Width           =   3015
         Begin VB.Image imgStud 
            BorderStyle     =   1  'Fixed Single
            Height          =   3135
            Left            =   -240
            Picture         =   "frmContact.frx":583A
            Stretch         =   -1  'True
            Top             =   0
            Width           =   3255
         End
      End
      Begin vkUserContolsXP.vkFrame vkFrame3 
         Height          =   375
         Index           =   0
         Left            =   4560
         TabIndex        =   22
         Top             =   2160
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BackGradient    =   0
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
         BorderColor     =   33023
         BreakCorner     =   0   'False
      End
      Begin vkUserContolsXP.vkLabel lblDept 
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "&Department:"
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
      Begin MSComDlg.CommonDialog Dialog1 
         Left            =   9240
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin vkUserContolsXP.vkCommand cmdAdd 
         Height          =   495
         Left            =   6360
         TabIndex        =   16
         Top             =   3960
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   873
         Caption         =   "ADD"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   33023
         BorderColor     =   33023
         CustomStyle     =   0
      End
      Begin vkUserContolsXP.vkCommand cmdUpdate 
         Height          =   495
         Left            =   8040
         TabIndex        =   18
         Top             =   6000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Caption         =   "Update"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   33023
         BorderColor     =   33023
         CustomStyle     =   0
      End
      Begin vkUserContolsXP.vkCommand cmdEdit 
         Height          =   495
         Left            =   6360
         TabIndex        =   17
         Top             =   6000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Caption         =   "Edit"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   33023
         BorderColor     =   33023
         CustomStyle     =   0
      End
      Begin vkUserContolsXP.vkFrame vkFrame2 
         Height          =   375
         Left            =   1560
         TabIndex        =   21
         Top             =   1200
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   661
         BackGradient    =   0
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
         BorderColor     =   33023
         BreakCorner     =   0   'False
         DisplayPicture  =   0   'False
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
         ItemData        =   "frmContact.frx":105C6
         Left            =   1560
         List            =   "frmContact.frx":105C8
         TabIndex        =   2
         Top             =   1200
         Width           =   4695
      End
      Begin vkUserContolsXP.vkTextBox txtName 
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   1680
         Width           =   4695
         _ExtentX        =   8281
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
      Begin vkUserContolsXP.vkTextBox txtDoB 
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   2160
         Width           =   2055
         _ExtentX        =   3625
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
      Begin vkUserContolsXP.vkTextBox txtFather 
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   2640
         Width           =   4695
         _ExtentX        =   8281
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
      Begin vkUserContolsXP.vkTextBox txtMother 
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   3120
         Width           =   4695
         _ExtentX        =   8281
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
      Begin vkUserContolsXP.vkTextBox txtOccupation 
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Top             =   3600
         Width           =   4695
         _ExtentX        =   8281
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
      Begin vkUserContolsXP.vkTextBox txtAddress 
         Height          =   375
         Left            =   1560
         TabIndex        =   9
         Top             =   4080
         Width           =   4695
         _ExtentX        =   8281
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
      Begin vkUserContolsXP.vkTextBox txtCity 
         Height          =   375
         Left            =   1560
         TabIndex        =   10
         Top             =   4560
         Width           =   1815
         _ExtentX        =   3201
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
      Begin vkUserContolsXP.vkTextBox txtPinCode 
         Height          =   375
         Left            =   4200
         TabIndex        =   11
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
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox txtState 
         Height          =   375
         Left            =   1560
         TabIndex        =   12
         Top             =   5040
         Width           =   3975
         _ExtentX        =   7011
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
         Enabled         =   0   'False
         BorderColor     =   33023
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox txtEmail 
         Height          =   375
         Left            =   1560
         TabIndex        =   13
         Top             =   5520
         Width           =   8055
         _ExtentX        =   14208
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
      Begin vkUserContolsXP.vkTextBox txtPhone 
         Height          =   375
         Left            =   6240
         TabIndex        =   14
         Top             =   4560
         Width           =   3375
         _ExtentX        =   5953
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
      Begin vkUserContolsXP.vkTextBox txtMobile 
         Height          =   375
         Left            =   6240
         TabIndex        =   15
         Top             =   5040
         Width           =   3375
         _ExtentX        =   5953
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
      Begin VB.ComboBox cmbGender 
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
         Left            =   4560
         TabIndex        =   5
         Top             =   2160
         Width           =   1695
      End
      Begin vkUserContolsXP.vkFrame vkFrame3 
         Height          =   375
         Index           =   1
         Left            =   1560
         TabIndex        =   20
         Top             =   720
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   661
         BackGradient    =   0
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
         BorderColor     =   33023
         BreakCorner     =   0   'False
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
         Left            =   1560
         TabIndex        =   1
         Top             =   720
         Width           =   4695
      End
   End
End
Attribute VB_Name = "frmContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmContact
Private sImageName As String
Public sTemp As String
Dim ContactTop As Integer
Dim State As Integer
Private Sub cmbDept_Click()
    iDept = Department(cmbDept)
    Call cmbRegNo_Load(cmbRegNo)
End Sub

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click()
    txtName.Enabled = True
    txtDoB.Enabled = True
    txtFather.Enabled = True
    txtMother.Enabled = True
    txtOccupation.Enabled = True
    txtAddress.Enabled = True
    txtCity.Enabled = True
    txtPinCode.Enabled = True
    txtState.Enabled = True
    txtEmail.Enabled = True
    txtPhone.Enabled = True
    txtMobile.Enabled = True
    cmbGender.Enabled = True
    cmdAdd.Enabled = True
End Sub
Private Sub cmbRegNo_Click()
    txtName.Text = ""
    txtDoB.Text = ""
    txtFather.Text = ""
    txtMother.Text = ""
    txtOccupation.Text = ""
    txtAddress.Text = ""
    txtCity.Text = ""
    txtPinCode.Text = ""
    txtState.Text = ""
    txtEmail.Text = ""
    txtMobile.Text = ""
    txtPhone.Text = ""
    sTemp = cmbRegNo.Text
    Call Loaddata(sTemp)
End Sub

Private Sub cmbRegNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmbRegNo_Click
End Sub

Private Sub cmdInsert_Click()
    On Error Resume Next
    Dim rs As New ADODB.Recordset
    Dim qr As String
    qr = "insert into studdetails values(" & txtRegNo.Text & "," & txtName.Text & "," & txtDoB.Text & "," & txtFather.Text & "," & txtMother.Text & "," & txtOccupation.Text & "," & txtAddress.Text & "," & txtCity.Text & "," & txtPinCode.Text & "," & txtState.Text & "," & txtEmail.Text & "," & txtPhone.Text & "," & txtMobile.Text & "," & cmbGender.Text & "," & sImageName & ")"
    conn.Execute qr
    rs.Open qr, conn, adOpenDynamic, adLockOptimistic, 1
    rs.Close
    Call CntCtrlHide
    If Err.Number <> 0 Then
        MsgBox Error & vbCrLf & "Error Number: " & Err.Number
    Else
        MsgBox "Updated"
    End If
End Sub


Private Sub cmdMin_Click()
    Dim i As Long
    If State = 1 Then
        State = 0
        cmdMin.Caption = "+"
        For i = 6705 To 310 Step -150
            Me.Height = i
            DoEvents
        Next i
        Me.Height = 310
        fContact.Height = 310
        fContact.BorderWidth = 0
        Me.Top = 100
    Else
        State = 1
        cmdMin.Caption = "--"
        For i = 310 To 6705 Step 150
            Me.Height = i
            DoEvents
        Next i
        Me.Height = 6705
        fContact.Height = 6705
        fContact.BorderWidth = 2
        Me.Top = ContactTop
    End If
End Sub

Private Sub fContact_MouseDown(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_Load()
    CreateRoundRectFromWindow Me, 7, 7
    Me.Top = 250
    Me.Left = (mdiMain.Width - Me.Width) / 2
    Me.BackColor = mdiMain.BackColor
    Call frmColor(frmContact)
    Call CntCtrlHide
    Call cmbDept_Load(cmbDept)
    Call cmbRegNo_Load(cmbRegNo)
    cmbGender.AddItem "Male"
    cmbGender.AddItem "Female"
    State = 1
    ContactTop = Me.Top
End Sub
Private Sub cmdAdd_Click()
    With Dialog1
       .InitDir = App.Path
       .Filter = "JPEG image|*.jpg|GIF image|*.gif|BITMAP image|*.bmp|Icon image|*.ico|Cursor image|*.cur|Panerio image|*.pan"
       .ShowOpen
          If .FileName <> "" Then
             sImageName = .FileName
             imgStud.Picture = LoadPicture(sImageName)
          End If
     End With
End Sub

Private Sub cmdUpdate_Click()
    On Error Resume Next
    Dim rsUpdate As New ADODB.Recordset
    Dim qr As String
    qr = "update studdetails set studname = '" & txtName.Text & "',dob= to_date('" & txtDoB.Text & "','DD-MM-YYYY'),father= '" & txtFather.Text & "',mother= '" & txtMother.Text & "',occupation= '" & txtOccupation.Text & "',address = '" & txtAddress.Text & "',city = '" & txtCity.Text & "',pincode = '" & txtPinCode.Text & "',state= '" & txtState.Text & "',email = '" & txtEmail.Text & "',landline = '" & txtPhone.Text & "',mobile = '" & txtMobile.Text & "',gender = '" & cmbGender.Text & "',image ='" & sImageName & "' Where regno = " & cmbRegNo.Text
    rsUpdate.Open qr, conn, adOpenDynamic, adLockOptimistic
    'Call MoveFile(sImageName)
    Call CntCtrlHide
    If Err.Number <> 0 Then
        MsgBox Error & vbCrLf & "Error Number: " & Err.Number
    Else
        JuraMsgBox ("Updated")
    End If
End Sub

Public Sub Loaddata(sTemp As String)
    Dim rs As New ADODB.Recordset
    Dim qr As String
    qr = "select * from studdetails where regno= " & sTemp
    rs.Open qr, conn, adOpenDynamic, adLockOptimistic, adCmdText
    On Error Resume Next
    With frmContact
      If rs.BOF = True Or rs.EOF = True Then
          Exit Sub
      Else
        .txtName.Text = rs!Studname
        .txtDoB.Text = rs!dob
        .txtFather.Text = rs!father
        .txtMother.Text = rs!mother
        .txtOccupation.Text = rs!occupation
        .txtAddress.Text = rs!Address
        .txtCity.Text = rs!city
        .txtPinCode.Text = rs!pincode
        .txtState.Text = rs!State
        .txtEmail.Text = rs!email
        .txtPhone.Text = rs!landline
        .txtMobile.Text = rs!mobile
        .cmbGender.Text = rs!gender
        .imgStud.Picture = LoadPicture(App.Path & "\images\" & "default.jpg")
        .imgStud.Picture = LoadPicture(rs!Image)
      End If
       .Image.Refresh
    End With
End Sub

Public Sub CntCtrlHide()
    txtName.Enabled = False
    txtDoB.Enabled = False
    txtFather.Enabled = False
    txtMother.Enabled = False
    txtOccupation.Enabled = False
    txtAddress.Enabled = False
    txtCity.Enabled = False
    txtPinCode.Enabled = False
    txtState.Enabled = False
    txtEmail.Enabled = False
    txtPhone.Enabled = False
    txtMobile.Enabled = False
    cmbGender.Enabled = False
    cmdAdd.Enabled = False
End Sub
Private Sub MoveFile(sFileName As String)
    FileCopy sFileName, App.Path & "\Images"
End Sub

Private Sub txtAddress_LostFocus()
    txtAddress.Text = JangidFormat(txtAddress.Text)
End Sub

Private Sub txtCity_Change()
    txtCity.Text = JangidFormat(txtCity.Text)
End Sub

Private Sub txtFather_LostFocus()
    txtFather.Text = JangidFormat(txtFather.Text)
End Sub

Private Sub txtMother_LostFocus()
    txtMother.Text = JangidFormat(txtMother.Text)
End Sub

Private Sub txtName_LostFocus()
    txtName.Text = JangidFormat(txtName.Text)
End Sub

Private Sub txtOccupation_LostFocus()
    txtOccupation.Text = JangidFormat(txtOccupation.Text)
End Sub

Private Sub txtState_LostFocus()
    txtState.Text = JangidFormat(txtState.Text)
End Sub
