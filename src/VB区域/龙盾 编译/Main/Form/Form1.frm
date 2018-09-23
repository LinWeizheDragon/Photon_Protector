VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#15.3#0"; "Codejock.Controls.v15.3.1.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "¹â×Ó·ÀÓùÍø"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11910
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   546
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   794
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Timer Timer_Reread 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   480
      Top             =   1320
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1695
      Left            =   8040
      TabIndex        =   16
      Top             =   240
      Visible         =   0   'False
      Width           =   3495
      Begin VB.Image ICON_DANGER 
         Height          =   765
         Left            =   2640
         Picture         =   "Form1.frx":145F5
         Top             =   960
         Width           =   765
      End
      Begin VB.Image ICON_SAFE 
         Height          =   765
         Left            =   120
         Picture         =   "Form1.frx":16DBB
         Top             =   840
         Width           =   765
      End
      Begin VB.Image ListNum 
         Height          =   1020
         Index           =   9
         Left            =   2400
         Picture         =   "Form1.frx":196A1
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   585
      End
      Begin VB.Image ListNum 
         Height          =   1020
         Index           =   8
         Left            =   1800
         Picture         =   "Form1.frx":1DBEA
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   585
      End
      Begin VB.Image ListNum 
         Height          =   1020
         Index           =   7
         Left            =   1200
         Picture         =   "Form1.frx":223D8
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   585
      End
      Begin VB.Image ListNum 
         Height          =   1020
         Index           =   6
         Left            =   600
         Picture         =   "Form1.frx":26400
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   585
      End
      Begin VB.Image ListNum 
         Height          =   1020
         Index           =   5
         Left            =   3000
         Picture         =   "Form1.frx":2AAB3
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
      Begin VB.Image ListNum 
         Height          =   1020
         Index           =   4
         Left            =   2400
         Picture         =   "Form1.frx":2EE94
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
      Begin VB.Image ListNum 
         Height          =   1020
         Index           =   3
         Left            =   1800
         Picture         =   "Form1.frx":33118
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
      Begin VB.Image ListNum 
         Height          =   1020
         Index           =   2
         Left            =   1200
         Picture         =   "Form1.frx":3782C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
      Begin VB.Image ListNum 
         Height          =   1020
         Index           =   1
         Left            =   600
         Picture         =   "Form1.frx":3BE41
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
      Begin VB.Image ListNum 
         Height          =   1020
         Index           =   0
         Left            =   0
         Picture         =   "Form1.frx":3FB11
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin VB.Timer Timer_Out 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   1440
      Top             =   1320
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   4080
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1200
      Top             =   240
   End
   Begin VB.Frame MainFrame 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   $"Form1.frx":44099
      Height          =   6015
      Index           =   3
      Left            =   150
      TabIndex        =   2
      Top             =   2160
      Width           =   11655
      Begin PhotonProtect.jcbutton Tool_KillFile 
         Height          =   1455
         Left            =   600
         TabIndex        =   22
         Top             =   3720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   2566
         ButtonStyle     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ºÚÌå"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "Ç¿ÖÆÉ¾³ýÎÄ¼þ"
         MousePointer    =   99
         MouseIcon       =   "Form1.frx":4414B
         Picture         =   "Form1.frx":442AD
         PictureHover    =   "Form1.frx":46F17
         PictureAlign    =   6
      End
      Begin PhotonProtect.jcbutton Tool_ProMon 
         Height          =   1455
         Left            =   2520
         TabIndex        =   27
         Top             =   3720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   2566
         ButtonStyle     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ºÚÌå"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "Áú¶Ü½ø³Ì¹ÜÀíÆ÷"
         MousePointer    =   99
         MouseIcon       =   "Form1.frx":4A0AC
         Picture         =   "Form1.frx":4A20E
         PictureHover    =   "Form1.frx":4CE78
         PictureAlign    =   6
      End
      Begin PhotonProtect.jcbutton Tool_MoreTool 
         Height          =   1455
         Left            =   6240
         TabIndex        =   28
         Top             =   3720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   2566
         ButtonStyle     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ºÚÌå"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "Ð¡¹¤¾ß"
         MousePointer    =   99
         MouseIcon       =   "Form1.frx":5000D
         Picture         =   "Form1.frx":5016F
         PictureHover    =   "Form1.frx":52DD9
         PictureAlign    =   6
      End
      Begin PhotonProtect.jcbutton Tool_FixFolders 
         Height          =   1455
         Left            =   4440
         TabIndex        =   29
         Top             =   3720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   2566
         ButtonStyle     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ºÚÌå"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "UÅÌÎÄ¼þ¼Ð»Ö¸´"
         MousePointer    =   99
         MouseIcon       =   "Form1.frx":55F6E
         Picture         =   "Form1.frx":560D0
         PictureHover    =   "Form1.frx":58D3A
         PictureAlign    =   6
      End
      Begin PhotonProtect.jcbutton Tool_Repair 
         Height          =   1455
         Left            =   600
         TabIndex        =   31
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   2566
         ButtonStyle     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ºÚÌå"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "Â©¶´ÐÞ¸´"
         MousePointer    =   99
         MouseIcon       =   "Form1.frx":5BECF
         Picture         =   "Form1.frx":5C031
         PictureHover    =   "Form1.frx":5EC9B
         PictureAlign    =   6
      End
      Begin PhotonProtect.jcbutton Tool_Clear 
         Height          =   1455
         Left            =   2520
         TabIndex        =   32
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   2566
         ButtonStyle     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ºÚÌå"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "ÏµÍ³ÇåÀí"
         MousePointer    =   99
         MouseIcon       =   "Form1.frx":61E30
         Picture         =   "Form1.frx":61F92
         PictureHover    =   "Form1.frx":64BFC
         PictureAlign    =   6
      End
      Begin PhotonProtect.jcbutton Tool_Imp 
         Height          =   1455
         Left            =   4440
         TabIndex        =   33
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   2566
         ButtonStyle     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ºÚÌå"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "ÏµÍ³ÓÅ»¯"
         MousePointer    =   99
         MouseIcon       =   "Form1.frx":67D91
         Picture         =   "Form1.frx":67EF3
         PictureHover    =   "Form1.frx":6AB5D
         PictureAlign    =   6
      End
      Begin VB.Image Image5 
         Height          =   450
         Left            =   360
         Picture         =   "Form1.frx":6DCF2
         Top             =   240
         Width           =   4500
      End
      Begin VB.Image Image4 
         Height          =   450
         Left            =   360
         Picture         =   "Form1.frx":71960
         Top             =   3120
         Width           =   4500
      End
   End
   Begin VB.Frame MainFrame 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6015
      Index           =   2
      Left            =   150
      TabIndex        =   1
      Top             =   2160
      Width           =   11655
      Begin PhotonProtect.jcbutton btn_Update 
         Height          =   495
         Left            =   3960
         TabIndex        =   24
         Top             =   3120
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         ButtonStyle     =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ËÎÌå"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16765357
         Caption         =   "¼ì²é¸üÐÂ"
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "²úÆ·Éý¼¶"
         BeginProperty Font 
            Name            =   "ºÚÌå"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1800
         TabIndex        =   25
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label VerINFO 
         BackColor       =   &H00FFFFFF&
         Caption         =   "³ÌÐò°æ±¾£º"
         BeginProperty Font 
            Name            =   "ËÎÌå"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   1455
         Left            =   2040
         TabIndex        =   23
         Top             =   1440
         Width           =   4335
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00C0C0C0&
         Height          =   2895
         Left            =   1680
         Shape           =   4  'Rounded Rectangle
         Top             =   1080
         Width           =   6615
      End
      Begin VB.Image Image2 
         Height          =   1200
         Left            =   6600
         Picture         =   "Form1.frx":74D07
         Top             =   1320
         Width           =   1200
      End
   End
   Begin VB.Frame MainFrame 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6015
      Index           =   1
      Left            =   150
      TabIndex        =   0
      Top             =   2160
      Width           =   11655
      Begin VB.Frame fra_Log 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   5895
         Left            =   11160
         TabIndex        =   8
         Top             =   0
         Width           =   11500
         Begin PhotonProtect.jcbutton jcbutton2 
            Height          =   375
            Left            =   480
            TabIndex        =   11
            Top             =   1440
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            ButtonStyle     =   5
            ShowFocusRect   =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ËÎÌå"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   11169024
            Caption         =   "½ø³Ì·À»¤ÈÕÖ¾"
            ForeColor       =   16777215
         End
         Begin PhotonProtect.jcbutton btn_Out 
            Height          =   1215
            Left            =   0
            TabIndex        =   10
            Top             =   1800
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   2143
            ButtonStyle     =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ËÎÌå"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16765357
            Caption         =   "<<"
         End
         Begin VB.TextBox LogTextShow 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "ºÚÌå"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3615
            Left            =   360
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   9
            Top             =   1920
            Width           =   10935
         End
         Begin PhotonProtect.jcbutton jcbutton3 
            Height          =   375
            Left            =   2520
            TabIndex        =   12
            Top             =   1440
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            ButtonStyle     =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ËÎÌå"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   11169024
            Caption         =   "Çý¶¯·À»¤ÈÕÖ¾"
            ForeColor       =   16777215
         End
         Begin PhotonProtect.jcbutton jcbutton4 
            Height          =   375
            Left            =   4560
            TabIndex        =   13
            Top             =   1440
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            ButtonStyle     =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ËÎÌå"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   11169024
            Caption         =   "×¢²á±í·À»¤ÈÕÖ¾"
            ForeColor       =   16777215
         End
         Begin PhotonProtect.jcbutton jcbutton5 
            Height          =   375
            Left            =   6600
            TabIndex        =   14
            Top             =   1440
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            ButtonStyle     =   5
            ShowFocusRect   =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ËÎÌå"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   11169024
            Caption         =   "UÅÌ·À»¤ÈÕÖ¾"
            ForeColor       =   16777215
         End
         Begin PhotonProtect.jcbutton jcbutton6 
            Height          =   375
            Left            =   8640
            TabIndex        =   15
            Top             =   1440
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            ButtonStyle     =   5
            ShowFocusRect   =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ËÎÌå"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   11169024
            Caption         =   "ÎÄ¼þÏµÍ³·À»¤ÈÕÖ¾"
            ForeColor       =   16777215
         End
         Begin VB.Image Tm4 
            Height          =   1020
            Left            =   9840
            Picture         =   "Form1.frx":78061
            Stretch         =   -1  'True
            Top             =   120
            Width           =   585
         End
         Begin VB.Image Tm3 
            Height          =   1020
            Left            =   9120
            Picture         =   "Form1.frx":7C5E9
            Stretch         =   -1  'True
            Top             =   120
            Width           =   585
         End
         Begin VB.Image Tm2 
            Height          =   1020
            Left            =   8400
            Picture         =   "Form1.frx":80BFE
            Stretch         =   -1  'True
            Top             =   120
            Width           =   585
         End
         Begin VB.Image Tm1 
            Height          =   1020
            Left            =   7680
            Picture         =   "Form1.frx":848CE
            Stretch         =   -1  'True
            Top             =   120
            Width           =   585
         End
         Begin VB.Image Image1 
            Height          =   1320
            Left            =   100
            Picture         =   "Form1.frx":88E56
            Top             =   0
            Width           =   11370
         End
      End
      Begin PhotonProtect.DSButton ds_Pro 
         Height          =   375
         Left            =   5880
         TabIndex        =   7
         Top             =   1680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
      End
      Begin PhotonProtect.DSButton ds_Reg 
         Height          =   375
         Left            =   5640
         TabIndex        =   17
         Top             =   4680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
      End
      Begin PhotonProtect.DSButton ds_USB 
         Height          =   375
         Left            =   7560
         TabIndex        =   18
         Top             =   5040
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
      End
      Begin PhotonProtect.DSButton ds_Driver 
         Height          =   375
         Left            =   5880
         TabIndex        =   30
         Top             =   2040
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
      End
      Begin PhotonProtect.jcbutton jcbutton1 
         Height          =   375
         Left            =   2280
         TabIndex        =   34
         Top             =   4920
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         ButtonStyle     =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ËÎÌå"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777088
         Caption         =   "Ë¢ÐÂ×´Ì¬"
      End
      Begin PhotonProtect.DSButton ds_Protect 
         Height          =   375
         Left            =   2520
         TabIndex        =   36
         Top             =   4360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "×ÔÎÒ±£»¤£º"
         BeginProperty Font 
            Name            =   "ËÎÌå"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1440
         TabIndex        =   35
         Top             =   4440
         Width           =   1335
      End
      Begin VB.Image Image3 
         Height          =   6090
         Left            =   0
         Picture         =   "Form1.frx":908E2
         Top             =   0
         Width           =   11700
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1275
      Left            =   150
      TabIndex        =   3
      Top             =   2160
      Width           =   8400
      Begin XtremeSuiteControls.ProgressBar Progress 
         Height          =   220
         Left            =   1440
         TabIndex        =   26
         Top             =   670
         Width           =   6615
         _Version        =   983043
         _ExtentX        =   11668
         _ExtentY        =   388
         _StockProps     =   93
         Value           =   50
         Scrolling       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
         BarColor        =   16761024
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   80
         Left            =   240
         Top             =   900
         Width           =   855
      End
      Begin VB.Image SafeDanger 
         Height          =   765
         Left            =   240
         Picture         =   "Form1.frx":AB675
         Top             =   120
         Width           =   765
      End
      Begin VB.Label LabelSafe 
         BackStyle       =   0  'Transparent
         Caption         =   "°²È«³Ì¶È£ºÖÐ"
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   3240
         TabIndex        =   5
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label LabelTip 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "»¶Ó­Ê¹ÓÃ¹â×Ó·ÀÓùÍø ÊµÊ±·À»¤Î´¿ªÆô"
         BeginProperty Font 
            Name            =   "ºÚÌå"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   240
         Width           =   6615
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5955
      Left            =   8640
      TabIndex        =   19
      Top             =   2160
      Width           =   3120
      Begin PhotonProtect.jcbutton ProgramUpdate 
         Height          =   1335
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   2355
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ºÚÌå"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "²úÆ·Éý¼¶"
         ForeColor       =   49152
         MousePointer    =   99
         MouseIcon       =   "Form1.frx":ADE3B
         Picture         =   "Form1.frx":ADF9D
         PictureHover    =   "Form1.frx":B1307
      End
      Begin PhotonProtect.jcbutton WebSite 
         Height          =   1335
         Left            =   240
         TabIndex        =   21
         Top             =   1680
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   2355
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ºÚÌå"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "¹Ù·½ÍøÕ¾"
         ForeColor       =   49152
         MousePointer    =   99
         MouseIcon       =   "Form1.frx":B449C
         Picture         =   "Form1.frx":B45FE
         PictureHover    =   "Form1.frx":BA1BE
      End
   End
   Begin VB.Menu mnuTray 
      Caption         =   "Ö÷²Ëµ¥"
      Visible         =   0   'False
      Begin VB.Menu mnuWebSite 
         Caption         =   "¹Ù·½ÍøÕ¾"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "¹ØÓÚ ¹â×Ó·ÀÓùÍø"
      End
      Begin VB.Menu mnuShow 
         Caption         =   "ÏÔÊ¾Ö÷½çÃæ"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "ÍË³ö"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_VCENTER = &H4
Private Const DT_SINGLELINE = &H20
Private Const SW_SHOW = 5
Dim IsOutLog As Boolean 'ÈÕÖ¾´°ÌåÊÇ·ñµ¯³öÁË
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private WithEvents c_DirectUI As GDirectUI.CDirectUI
Attribute c_DirectUI.VB_VarHelpID = -1

'//Æ¤·ôÎÄ¼þÂ·¾¶
Private m_SkinPath  As String
Private m_hBGDC     As Long
Private m_hCtrDC    As Long
Private m_hTabDC    As Long
Private m_hPgDC1    As Long
Private m_hPgDC2    As Long
Private m_hPgDC3    As Long
Private m_hPgDC4    As Long
Private m_hBtnDC    As Long
Private m_hNBtnDC   As Long


Function GetDirectory(Optional Msg) As String
Dim bInfo As BROWSEINFO
Dim Path As String
Dim r As Long, x As Long, Pos As Integer
' Root folder = Desktop
bInfo.pidlRoot = 0&
' Title in the dialog
If IsMissing(Msg) Then
bInfo.lpszTitle = "Ñ¡ÔñÉ¨ÃèÄ¿Â¼"
Else
bInfo.lpszTitle = Msg
End If
' Type of directory to return
bInfo.ulFlags = &H1
' Display the dialog
x = SHBrowseForFolder(bInfo)
' Parse the result
Path = Space$(512)
r = SHGetPathFromIDList(ByVal x, ByVal Path)
If r Then
Pos = InStr(Path, Chr$(0))
GetDirectory = Left(Path, Pos - 1)
Else
GetDirectory = ""
End If
End Function
Private Sub btn_Out_Click()
If IsOutLog Then
IsOutLog = False '·µ»ØÖÐ
btn_Out.Caption = ">>"
Timer_Out.Enabled = True
Else
IsOutLog = True 'µ¯³öÖÐ
btn_Out.Caption = "<<"
Timer_Out.Enabled = True
'Æô¶¯¶¨Ê±Æ÷
End If
End Sub

Private Sub btn_Update_Click()
On Error Resume Next
If Not Dir(App.Path & "\ProgramUpdate.exe") = "" Then
OpenFile App.Path & "\ProgramUpdate.exe", , vbNormalFocus
End If
End Sub

Private Sub btn_Update_MouseEnter()
btn_Update.Caption = "µã»÷¿ªÊ¼¸üÐÂ"

End Sub

Private Sub btn_Update_MouseLeave()
btn_Update.Caption = "¼ì²é¸üÐÂ"
End Sub

Private Sub c_DirectUI_OnControlClick(ByVal Key As String)
On Error Resume Next
    '//¿Ø¼þµ¥»÷ c_DirectUI.Handler ÊÇ´°¿Ú¿ØÖÆº¯ÊýµÄ¼¯ºÏ
    Select Case Key
        Case "btnClose": Me.Hide
        Case "btnMin": c_DirectUI.Handler.HitMinimizeButton
        Case "btnMenu"
             PopupMenu mnuTray, , , , mnuShow
'            c_DirectUI.Handler.MsgBox "¸ÃÊ¾ÀýµÄc32bppDIB»æÍ¼ÓÐÐ©Âý£¬ËùÒÔ²ÉÓÃGDI»æÍ¼£¬Èç¹ûÒªÊ¹ÓÃGDI+ÇëÊ¹ÓÃÆäËû»æÍ¼Ä£¿é¡£", _
'                                      vbInformation Or vbOKOnly, "GDirectUI"
        Case "stLogo"
            OpenUrl "www.dvmsc.com"
        Case "radSD", "radFH", "radSJ", "radGJ"
            If c_DirectUI.ControlCollection(Key).Visible Then Exit Sub
            c_DirectUI.ControlCollection("radSD").Visible = False
            c_DirectUI.ControlCollection("radFH").Visible = False
            c_DirectUI.ControlCollection("radSJ").Visible = False
            c_DirectUI.ControlCollection("radGJ").Visible = False
            Dim i
            For i = 1 To 3
              MainFrame(i).Visible = False
            Next
            Select Case Key
             Case "radSD"
               'MainFrame(0).Visible = True
             Case "radFH"
               MainFrame(1).Visible = True
             Case "radSJ"
               MainFrame(2).Visible = True
             Case "radGJ"
               MainFrame(3).Visible = True
            End Select
            ReRead
            c_DirectUI.ControlCollection(Key).Visible = True
        Case "btnQSSM"
            DoScan True, True, ""
        Case "btnQPSM"
         frmChoose.Show
        Case "btnZDSM"
            Dim Target As String
            Target = GetDirectory
            If Target <> "" Then
            DoScan False, False, Target
            End If
    End Select
    ReRead
End Sub

Private Sub c_DirectUI_OnDrawBackground(ByVal hdc As Long, ByVal cX As Long, ByVal cY As Long)
    '//»æÖÆ´°¿Ú±³¾° c_DirectUI.Painter ÊÇ»æÍ¼º¯ÊýµÄ¼¯ºÏ
    With c_DirectUI
        .Painter.FillColor hdc, 0, 0, cX, cY, &H0
        .Painter.BitBlt hdc, 1, 1, cX, cY, m_hBGDC, 0, 0, vbSrcCopy
        .Handler.SetRoundRectRgn 6, 6
    End With
End Sub

Private Sub c_DirectUI_OnDrawControl(ByVal Key As String, ByVal CtlID As Long, _
                                    ByVal PartID As Long, ByVal State As Long, _
                                    ByVal hdc As Long, ByVal cX As Long, _
                                    ByVal cY As Long, SkinDefault As Boolean)
    '//»æÖÆ¿Ø¼þ£¬Òª½øÐÐ×Ô¶¨Òå»æÖÆ Çë·µ»ØSkinDefaultÎªTrue
    With c_DirectUI.Painter
        Select Case Key
            Case "btnClose"
                .BitBlt hdc, 0, 0, cX, cY, m_hCtrDC, 68, State * cY, vbSrcCopy
                SkinDefault = True
            Case "btnMin"
                .BitBlt hdc, 0, 0, cX, cY, m_hCtrDC, 34, State * cY, vbSrcCopy
                SkinDefault = True
            Case "btnMenu"
                .BitBlt hdc, 0, 0, cX, cY, m_hCtrDC, 0, State * cY, vbSrcCopy
                SkinDefault = True
            
            Case "stPage"
                .GridRect hdc, 0, 0, cX, cY, vbBlack, vbWhite, 1, 1, 1, 1
                
            Case "pg1"
                .BitBlt hdc, 0, 0, cX, cY, m_hPgDC1, 0, 0, vbSrcCopy
            Case "pg2"
                .BitBlt hdc, 0, 0, cX, cY, m_hPgDC2, 0, 0, vbSrcCopy
            Case "pg3"
                .BitBlt hdc, 0, 0, cX, cY, m_hPgDC3, 0, 0, vbSrcCopy
            Case "pg4"
                .BitBlt hdc, 0, 0, cX, cY, m_hPgDC4, 0, 0, vbSrcCopy
            
            Case "btn1", "btn2", "btn3", "btn4", "btn5", "btn6", "btn7"
                .GridBlt hdc, 0, 0, cX, cY, m_hNBtnDC, 0, State * 19, 48, 19, 3, 3, 3, 3, RGB(255, 0, 255)
                .DrawText hdc, c_DirectUI.Control(Key).Text, 3, 3, cX - 6, cY - 6, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
                SkinDefault = True
                
            Case "btnQSSM"
                .BitBlt hdc, 0, 0, cX, cY, m_hBtnDC, 0, State * cY, vbSrcCopy
                SkinDefault = True
            Case "btnQPSM"
                .BitBlt hdc, 0, 0, cX, cY, m_hBtnDC, cX, State * cY, vbSrcCopy
                SkinDefault = True
            Case "btnZDSM"
                .BitBlt hdc, 0, 0, cX, cY, m_hBtnDC, cX * 2, State * cY, vbSrcCopy
                SkinDefault = True
                
            Case "radSD", "radFH", "radSJ", "radGJ"
                Dim SrcX As Long
                If Key = "radSD" Then SrcX = 0
                If Key = "radFH" Then SrcX = cX
                If Key = "radSJ" Then SrcX = cX * 2
                If Key = "radGJ" Then SrcX = cX * 3
                If c_DirectUI.Control(Key).Value Then
                    .BitBlt hdc, 0, 0, cX, cY, m_hTabDC, SrcX, cY * 3, vbSrcCopy
                Else
                    .BitBlt hdc, 0, 0, cX, cY, m_hTabDC, SrcX, cY * State, vbSrcCopy
                End If
                SkinDefault = True
        End Select
    End With
End Sub

Private Sub c_DirectUI_OnMouseDown(ByVal Button As Long, ByVal Shift As Long, ByVal x As Long, ByVal y As Long)
    '//´°¿ÚÊó±ê°´ÏÂ
    If Button = 1 Then c_DirectUI.Handler.HitCaption
End Sub

Public Function ReSetButton()
ds_Pro.Reset
ds_Driver.Reset
ds_USB.Reset
ds_Reg.Reset
ds_Protect.Reset
End Function

Private Sub Cmd_Driver_MouseEnter()
ReRead
End Sub

Private Sub Cmd_Driver_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
ReSetButton
End Sub

Private Sub Cmd_Pro_MouseEnter()
ReRead
End Sub

Private Sub Cmd_Pro_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
ReSetButton
End Sub





Private Sub Cmd_Reg_MouseEnter()
ReRead
End Sub

Private Sub Cmd_Reg_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
ReSetButton
End Sub

Private Sub Cmd_USB_MouseEnter()
ReRead
End Sub

Private Sub Cmd_USB_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
ReSetButton
End Sub

Private Sub DSButton1_GotFocus()
MsgBox ".."

End Sub



Private Sub Command1_Click()
MsgBox GetDirectory
End Sub

Private Sub form_Initialize()
Call CreatTray(Me, "¹â×Ó·ÀÓùÍø - ÔËÐÐÖÐ", "¹â×Ó·ÀÓùÍø", "¹â×Ó·ÀÓùÍøÒÑÆô¶¯", 4)
'Dim Pathinfo As String
'If Len(App.Path) > 20 Then
'Pathinfo = Left(App.Path, 17) & "..."
'Else
'Pathinfo = App.Path
'End If
VerINFO.Caption = Replace("³ÌÐòÃû£º¹â×Ó·ÀÓùÍø||" & _
                  "||³ÌÐò°æ±¾£º" & App.Major & "." & App.Minor & _
                  "||Î¬»¤£ºÁúÒ÷Ìì±ä||" & _
                  "¹Ù·½ÍøÕ¾£ºwww.dvmsc.com", "||", vbCrLf)
                  
    InitCommonControls
    'µ÷Õû´°ÌåÎ»ÖÃ
    fra_Log.Left = fra_Log.Width - btn_Out.Width
    fra_Log.Width = btn_Out.Width
    IsOutLog = False
    'Î´µ¯³ö
    Load frmSkin
If Command = "-start" Then
frmMain.Hide
End If
Load frmInit
Load frmRec

frmRec.Show
ChangeNum
ReRead
End Sub

Private Sub Form_Load()

On Error Resume Next
Progress.Value = 5

IniPath = App.Path & "\Set.ini"

    Dim i   As Long
    Dim x
            For x = 1 To 3
              MainFrame(x).Visible = False
              
              
            Next
            

    
    
    Set c_DirectUI = New CDirectUI
    m_SkinPath = App.Path
    If Right$(m_SkinPath, 1) <> "\" Then m_SkinPath = m_SkinPath & "\"
    m_SkinPath = m_SkinPath & "Skin\"

    Call LoadSkin
    
    With c_DirectUI
        '//¸øÖ¸¶¨´°¿Ú½øÐÐDirectUI
        .Attach Me.hwnd
        
        '//ÔÊÐíµÄ×îÐ¡¿í¶ÈºÍ×îÐ¡¸ß¶È
        .MinWidth = 800
        .MinHeight = 580
        
        '//Ôö¼ÓÒ»¸ö¿Ø¼þ¼¯ºÏ
        .AddCollection New CCollection
        With .ControlCollection(0)
            '//¿Ø¼þ¼¯ºÏµÄKey
            .Key = "NC_CONTROL"
            '//Ìí¼Ó¿Ø¼þ (Ä¿Ç°ÔÝÊ±Ã»ÓÐ½âÊÍxml¹¦ÄÜ)
            .AddControl New CButton, , "btnClose", , c_DirectUI.Width - 57, 0, 45, 18
            .AddControl New CButton, , "btnMin", , c_DirectUI.Width - 91, 0, 34, 18
            .AddControl New CButton, , "btnMenu", , c_DirectUI.Width - 125, 0, 34, 18
            .Control("btnClose").MousePointer = vbGDUIHand
            .Control("btnMin").MousePointer = vbGDUIHand
            .Control("btnMenu").MousePointer = vbGDUIHand
        End With
        
        .AddCollection New CCollection
        With .ControlCollection(1)
            .Key = "UI_ELEMENT"
            .AddControl New CStatic, , "stLogo", , 23, 18, 206, 72
            .Control("stLogo").MousePointer = vbGDUIHand
            
            .AddControl New CStatic, , "stPage", , 9, 141, 782, 409
            .AddControl New CStatic, "³ÌÐò°æ±¾" & App.Major & "." & App.Minor, "stVer", , 11, 560, 128, 24
            .AddControl New CStatic, "", "stTip", , 143, 560, 128, 24
            .AddControl New CStatic, "»¶Ó­Ê¹ÓÃ¹â×Ó·ÀÓùÍø", _
                                "stHlp", , c_DirectUI.Width - 404, 560, 400, 24
            '//¿Ø¼þÎÄ×ÖÑÕÉ«
            .Control("stVer").TextColor = &HFFFFFF
            .Control("stTip").TextColor = &HFFFFFF
            .Control("stHlp").TextColor = &HFFFFFF
            
            .AddControl New CRadioButton, "²¡¶¾²éÉ±", "radSD", , 19, 103, 122, 40
            .Control("radSD").Value = True
            .Control("radSD").MousePointer = vbGDUIHand

            .AddControl New CRadioButton, "ÊµÊ±·À»¤", "radFH", , 143, 103, 122, 40
            .Control("radFH").MousePointer = vbGDUIHand
            
            .AddControl New CRadioButton, "ÔÚÏßÉý¼¶", "radSJ", , 267, 103, 122, 40
            .Control("radSJ").MousePointer = vbGDUIHand
            
            .AddControl New CRadioButton, "¹¤¾ß´óÈ«", "radGJ", , 391, 103, 122, 40
            .Control("radGJ").MousePointer = vbGDUIHand
            
        End With
        
        .AddCollection New CCollection
        With .ControlCollection(2)
            .Key = "radSD"
            .AddControl New CStatic, , "pg1", , 10, 142, 781, 408
            
            .AddControl New CButton, "¿ìËÙÉ¨Ãè", "btnQSSM", , 34, 374, 152, 151
            .AddControl New CButton, "È«ÅÌÉ¨Ãè", "btnQPSM", , 213, 374, 152, 151
            .AddControl New CButton, "Ö¸¶¨Î»ÖÃÉ¨Ãè", "btnZDSM", , 392, 374, 152, 151
            '//¿Ø¼þÊó±êÐÎ×´
            .Control("btnQSSM").MousePointer = vbGDUIHand
            .Control("btnQPSM").MousePointer = vbGDUIHand
            .Control("btnZDSM").MousePointer = vbGDUIHand
        End With
        
        .AddCollection New CCollection
        With .ControlCollection(3)
            .Key = "radFH"
            .Visible = False
            .AddControl New CStatic, , "pg2", , 10, 142, 781, 408
            
'            .AddControl New CButton, "Á¢¼´¿ªÆô", "btn1", , 628, 176, 79, 24
'            .AddControl New CButton, "¹Ø±Õ", "btn2", , 659, 247, 48, 19
'            .AddControl New CButton, "¹Ø±Õ", "btn3", , 659, 279, 48, 19
'            .AddControl New CButton, "¹Ø±Õ", "btn4", , 659, 311, 48, 19
'            .AddControl New CButton, "¹Ø±Õ", "btn5", , 659, 343, 48, 19
'            .AddControl New CButton, "¹Ø±Õ", "btn6", , 659, 375, 48, 19
'
'            .Control("btn1").MousePointer = vbGDUIHand
'            .Control("btn2").MousePointer = vbGDUIHand
'            .Control("btn3").MousePointer = vbGDUIHand
'            .Control("btn4").MousePointer = vbGDUIHand
'            .Control("btn5").MousePointer = vbGDUIHand
'            .Control("btn6").MousePointer = vbGDUIHand
        End With
        
        .AddCollection New CCollection
        With .ControlCollection(4)
            .Key = "radSJ"
            .Visible = False
            .AddControl New CStatic, , "pg3", , 10, 142, 781, 408
            
            .AddControl New CButton, "ÁË½â", "btn7", , 150, 366, 88, 24
            .Control("btn7").MousePointer = vbGDUIHand
        End With
        
        .AddCollection New CCollection
        With .ControlCollection(5)
            .Key = "radGJ"
            .Visible = False
            .AddControl New CStatic, , "pg4", , 10, 142, 781, 408
        End With
   End With
   Timer1.Enabled = True
   ds_Pro.SetClickMode 1
ds_Driver.SetClickMode 2
ds_Reg.SetClickMode 3
ds_USB.SetClickMode 4
ds_Protect.SetClickMode 5
Timer_Reread.Enabled = True
End Sub

Private Sub LoadSkin()
    Dim dib As New c32bppDIB
        
    dib.LoadPicture_File m_SkinPath & "bg_theme-Sky.png"
    m_hBGDC = c_DirectUI.Painter.CreateMemDC(dib.Width, dib.Height)
    dib.Render m_hBGDC
    
    dib.LoadPicture_File m_SkinPath & "Control.png"
    m_hCtrDC = c_DirectUI.Painter.CreateMemDC(dib.Width, dib.Height)
    dib.Render m_hCtrDC
    
    dib.LoadPicture_File m_SkinPath & "TAB.png"
    m_hTabDC = c_DirectUI.Painter.CreateMemDC(dib.Width, dib.Height)
    dib.Render m_hTabDC
    
    dib.LoadPicture_File m_SkinPath & "Page1.png"
    m_hPgDC1 = c_DirectUI.Painter.CreateMemDC(dib.Width, dib.Height)
    dib.Render m_hPgDC1

    dib.LoadPicture_File m_SkinPath & "Page2.png"
    m_hPgDC2 = c_DirectUI.Painter.CreateMemDC(dib.Width, dib.Height)
    dib.Render m_hPgDC2

    dib.LoadPicture_File m_SkinPath & "Page3.png"
    m_hPgDC3 = c_DirectUI.Painter.CreateMemDC(dib.Width, dib.Height)
    dib.Render m_hPgDC3

    dib.LoadPicture_File m_SkinPath & "Page4.png"
    m_hPgDC4 = c_DirectUI.Painter.CreateMemDC(dib.Width, dib.Height)
    dib.Render m_hPgDC4

    dib.LoadPicture_File m_SkinPath & "Button.png"
    m_hBtnDC = c_DirectUI.Painter.CreateMemDC(dib.Width, dib.Height)
    dib.Render m_hBtnDC
    
    dib.LoadPicture_File m_SkinPath & "btn.png"
    m_hNBtnDC = c_DirectUI.Painter.CreateMemDC(dib.Width, dib.Height)
    dib.Render m_hNBtnDC
End Sub

Private Sub DestroySkin()
    c_DirectUI.Painter.DeleteDC m_hBGDC
    c_DirectUI.Painter.DeleteDC m_hCtrDC
    c_DirectUI.Painter.DeleteDC m_hTabDC
    c_DirectUI.Painter.DeleteDC m_hPgDC1
    c_DirectUI.Painter.DeleteDC m_hPgDC2
    c_DirectUI.Painter.DeleteDC m_hPgDC3
    c_DirectUI.Painter.DeleteDC m_hPgDC4
    c_DirectUI.Painter.DeleteDC m_hBtnDC
    c_DirectUI.Painter.DeleteDC m_hNBtnDC
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
If c_DirectUI.Handler.MsgBox("ÄúÈ·¶¨ÒªÍË³ö¹â×Ó·ÀÓùÍøÂð£¿" & vbCrLf & "ÌáÊ¾£ºÏÖÔÚ»¹ÓÐÒ»µãµãÐ¡ÎÊÌâ£¬Îª·ÀÖ¹ÍË³öºóµÄ´íÎóÌáÊ¾£¬ÍË³öºóÐèÒªÊÖ¶¯½áÊøÖ÷½ø³Ì¡£", vbYesNo, "¹â×Ó·ÀÓùÍø") = vbNo Then
Cancel = 1
Exit Sub
End If
Me.Hide

strShare = "Protect"
SuperSleep 1
strShare = "Protect.Unload"
SuperSleep 1
'Call ShowTip("¹â×Ó·ÀÓùÍø", "³ÌÐòÕýÔÚ¹Ø±ÕÏà¹Ø½ø³Ì£¬Íê³Éºó×Ô¶¯ÍË³öÍÐÅÌ......", 4)
If exitproc("USBRTA.exe") = True Then
  strShare = "USBRTA"
  SuperSleep 1
  strShare = "USBRTA.Close"
  SuperSleep 1
End If

If exitproc("ProcessRTA.exe") = True Then
  strShare = "ProcessRTA"
  SuperSleep 1
  strShare = "ProcessRTA.Unload"
  SuperSleep 1
End If
If exitproc("RegRTA.exe") = True Then
  strShare = "RegRTA"
  SuperSleep 1
  strShare = "RegRTA.Unload"
End If

'AddInfo "³ÌÐò¿ªÊ¼ÍË³ö£¬±£´æÈÕÖ¾......"
'Dim MyFSO As New FileSystemObject
'Dim DataNum As Integer
'Dim LogStr As String
'Dim LogNum As Integer
'LogStr = "¹â×Ó·ÀÓùÍøÖ÷³ÌÐòÈÕÖ¾" & vbCrLf & "Éú³ÉÊ±¼ä£º" & Now & vbCrLf & _
'"-------------------------" & vbCrLf
'LogNum = InfoView.ListItems.Count
'Do Until LogNum = 0
'LogStr = LogStr & InfoView.ListItems(LogNum).Text & ":" & InfoView.ListItems(LogNum).SubItems(1) & vbCrLf
'LogNum = LogNum - 1
'Loop
'DataNum = MyFSO.GetFolder(App.Path & "\Data\").Files.Count + 1
'Open App.Path & "\Data\¹â×Ó·ÀÓùÍøÈÕÖ¾-" & DataNum & ".log" For Append As #1
'Print #1, LogStr
'Close #1
'

If UnloadMode = vbAppWindows Then
    Dim i
    For Each i In Forms
    Unload i
    Next
End
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call DestroySkin
    Set c_DirectUI = Nothing
    UnloadTray
    Dim i
    For Each i In Forms
    Unload i
    Next
    Unload frmMain
    
End Sub


Public Function SuperSleep(DealyTime As Single) '´Ë´¦Ô­Îªlong£¬ÐÞ¸ÄÎªsingle¿ÉÑÓÊ±1ms :SK<2<8h
Dim TimerCount As Single
    TimerCount = Timer + DealyTime 'Ôö¼ÓXÃë ZJ9x6|q
    While TimerCount - Timer > 0
        DoEvents
        Sleep 1
    Wend
   
End Function





Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
ReSetButton
End Sub

Private Sub jcbutton1_Click()
ReRead
End Sub

Private Sub jcbutton2_Click()
ReadLogFile 1
End Sub

Private Sub jcbutton3_Click()
ReadLogFile 3

End Sub

Private Sub jcbutton4_Click()
ReadLogFile 4
End Sub

Private Sub jcbutton5_Click()
ReadLogFile 5
End Sub

Private Sub jcbutton6_Click()
ReadLogFile 2
End Sub

Private Sub jcbutton7_Click()

End Sub

Private Sub MainFrame_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
ReSetButton

End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuExit_Click()
Unload frmMain
End Sub

Private Sub mnuShow_Click()
frmMain.Show
End Sub

Private Sub mnuWebSite_Click()
OpenUrl "www.dvmsc.com"
End Sub

Private Sub ProgramUpdate_Click()
On Error Resume Next
If Not Dir(App.Path & "\ProgramUpdate.exe") = "" Then
OpenFile App.Path & "\ProgramUpdate.exe", , vbNormalFocus
End If
End Sub

Private Sub ProgramUpdate_MouseEnter()
ProgramUpdate.ForeColor = &H8080FF
End Sub

Private Sub ProgramUpdate_MouseLeave()
ProgramUpdate.ForeColor = &HC000&
End Sub

Private Sub Timer_Out_Timer()
Debug.Print fra_Log.Left
If IsOutLog Then
'»ØÊÕ²Ù×÷
If Not fra_Log.Width <= btn_Out.Width Then
 fra_Log.Width = fra_Log.Width - 100
 fra_Log.Left = fra_Log.Left + 100
Else 'Íê³ÉÁË
 Timer_Out.Enabled = False
End If
Else '·Å³ö²Ù×÷
If Not fra_Log.Left = 0 Then
 fra_Log.Width = fra_Log.Width + 100
 fra_Log.Left = fra_Log.Left - 100
Else
 Timer_Out.Enabled = False
End If
End If
End Sub

Private Sub Timer1_Timer()
'Dim i
'For i = 20 To 50
'CoolBar1.SetValue i
'SuperSleep 0.1
'Next
Timer1.Enabled = False
End Sub


Public Function SetStatus(ByVal mode)


End Function

Public Function ReadLogFile(ByVal LType As Integer)
On Error GoTo Err:
Dim strText
Dim strPath
Select Case LType

Case 1 '½ø³ÌÈÕÖ¾
strPath = App.Path & "\ProLog.dat"
Case 2
strPath = App.Path & "\FileLog.dat"
Case 3
strPath = App.Path & "\DrvLog.dat"
Case 4
strPath = App.Path & "\RegLog.dat"
Case 5
strPath = App.Path & "\USBLog.dat"
End Select
If Dir(strPath) <> "" Then
Dim Lines As String
Dim NextLine As String
Open strPath For Input As #1
Do While Not EOF(1)
Line Input #1, NextLine
Lines = Lines & NextLine & vbCrLf
Loop
strText = Lines
If strText = "" Then strText = "ÔÝÎÞÈÕÖ¾"
LogTextShow.Text = strText
Close #1
Else
strText = "ÔÝÎÞÈÕÖ¾"
LogTextShow.Text = strText
End If
Exit Function
Err:
strText = "¶ÁÈ¡´íÎó"
LogTextShow.Text = strText
End Function



Private Sub Tool_Clear_Click()
On Error Resume Next
If Not Dir(App.Path & "\PhotonClear.exe", vbHidden Or vbSystem Or vbNormal Or vbReadOnly) = "" Then
OpenFile App.Path & "\PhotonClear.exe"
End If
End Sub

Private Sub Tool_FixFolders_Click()
On Error Resume Next
If Not Dir(App.Path & "\Tools\FixFolders\¹â×Ó·ÀÓùÍø-ÒÆ¶¯ÅÌÒþ²ØÎÄ¼þ¼ÐÐÞ¸´¹¤¾ß.exe", vbHidden Or vbSystem Or vbNormal Or vbReadOnly) = "" Then
OpenFile App.Path & "\Tools\FixFolders\¹â×Ó·ÀÓùÍø-ÒÆ¶¯ÅÌÒþ²ØÎÄ¼þ¼ÐÐÞ¸´¹¤¾ß.exe"
End If
End Sub

Private Sub Tool_Imp_Click()
On Error Resume Next
If Not Dir(App.Path & "\PhotonMajorization.exe", vbHidden Or vbSystem Or vbNormal Or vbReadOnly) = "" Then
OpenFile App.Path & "\PhotonMajorization.exe"
End If
End Sub

Private Sub Tool_KillFile_Click()
On Error Resume Next
If Not Dir(App.Path & "\Tools\KillFIles\KillFile.exe", vbHidden Or vbSystem Or vbNormal Or vbReadOnly) = "" Then
OpenFile App.Path & "\Tools\KillFIles\KillFile.exe"
End If
End Sub

Private Sub Tool_MoreTool_Click()
On Error Resume Next
If Not Dir(App.Path & "\Tools\USBTools\", vbDirectory Or vbHidden Or vbSystem) = "" Then
OpenFile App.Path & "\Tools\USBTools\"
End If
End Sub

Private Sub Tool_ProMon_Click()
On Error Resume Next
If Not Dir(App.Path & "\Tools\ProcessMonitor\ProcessMonitor.exe", vbHidden Or vbSystem Or vbNormal Or vbReadOnly) = "" Then
OpenFile App.Path & "\Tools\ProcessMonitor\ProcessMonitor.exe"
End If
End Sub

Private Sub Tool_Repair_Click()
On Error Resume Next
If Not Dir(App.Path & "\PhotonRepair.exe", vbHidden Or vbSystem Or vbNormal Or vbReadOnly) = "" Then
OpenFile App.Path & "\PhotonRepair.exe"
End If
End Sub

Private Sub WebSite_Click()
OpenUrl "www.dvmsc.com"
End Sub

Private Sub WebSite_MouseEnter()
WebSite.ForeColor = &H8080FF
End Sub

Private Sub WebSite_MouseLeave()
WebSite.ForeColor = &HC000&
End Sub
Private Sub OpenUrl(tUrl As String)
On Error Resume Next
'==º¯Êý£ºÊ¹ÓÃÄ¬ÈÏä¯ÀÀÆ÷´ò¿ªÖ¸¶¨ÍøÒ³==
    ShellExecute Me.hwnd, "Open", tUrl, 0, 0, 0
End Sub
Public Function ReRead()
On Error Resume Next
If exitproc("ProcessRTA.exe") Then
ds_Pro.SetStatus True
ds_Driver.SetStatus True
Else
ds_Pro.SetStatus False
ds_Driver.SetStatus False
End If
If exitproc("RegRTA.exe") Then
ds_Reg.SetStatus True
Else
ds_Reg.SetStatus False
End If
If exitproc("USBRTA.exe") Then
ds_USB.SetStatus True
Else
ds_USB.SetStatus False
End If
If exitproc("ProtectProcess.exe") Then
ds_Protect.SetStatus True
Else
ds_Protect.SetStatus False
End If

Dim SafeNum As Integer
SafeNum = 0
If ds_Pro.GetStatus Then
SafeNum = SafeNum + 1
End If
If ds_Driver.GetStatus Then
SafeNum = SafeNum + 1
End If
If ds_Reg.GetStatus Then
SafeNum = SafeNum + 1
End If
If ds_USB.GetStatus Then
SafeNum = SafeNum + 1
End If
If ds_Protect.GetStatus Then
SafeNum = SafeNum + 1
End If
Select Case SafeNum
Case 0 'ÎÞ·À»¤
Progress.Value = 5
LabelSafe.ForeColor = &HFF&
LabelSafe.Caption = "°²È«³Ì¶È£º²î"
LabelTip.Caption = "»¶Ó­Ê¹ÓÃ¹â×Ó·ÀÓùÍø  ÊµÊ±·À»¤Î´¿ªÆô ½¨ÒéÄúÂíÉÏ¿ªÆô£¡"
SafeDanger.Picture = ICON_DANGER.Picture
Case 1
Progress.Value = 20
LabelSafe.ForeColor = &HFF&
LabelSafe.Caption = "°²È«³Ì¶È£º²î"
LabelTip.Caption = "»¶Ó­Ê¹ÓÃ¹â×Ó·ÀÓùÍø  ÊµÊ±·À»¤Î´È«²¿¿ªÆô ½¨ÒéÄúÂíÉÏ¿ªÆô£¡"
SafeDanger.Picture = ICON_DANGER.Picture
Case 2
Progress.Value = 40
LabelSafe.ForeColor = &H80FF&
LabelSafe.Caption = "°²È«³Ì¶È£ºÖÐ"
LabelTip.Caption = "»¶Ó­Ê¹ÓÃ¹â×Ó·ÀÓùÍø  ÊµÊ±·À»¤Î´È«²¿¿ªÆô ½¨ÒéÄúÂíÉÏ¿ªÆô£¡"
SafeDanger.Picture = ICON_DANGER.Picture
Case 3
Progress.Value = 60
LabelSafe.ForeColor = &H80FF&
LabelSafe.Caption = "°²È«³Ì¶È£ºÖÐ"
LabelTip.Caption = "»¶Ó­Ê¹ÓÃ¹â×Ó·ÀÓùÍø  ÊµÊ±·À»¤Î´È«²¿¿ªÆô ½¨ÒéÄúÂíÉÏ¿ªÆô£¡"
SafeDanger.Picture = ICON_SAFE.Picture
Case 4
Progress.Value = 60
LabelSafe.ForeColor = &H80FF&
LabelSafe.Caption = "°²È«³Ì¶È£ºÖÐ"
LabelTip.Caption = "»¶Ó­Ê¹ÓÃ¹â×Ó·ÀÓùÍø  ÊµÊ±·À»¤Î´È«²¿¿ªÆô ½¨ÒéÄúÂíÉÏ¿ªÆô£¡"
SafeDanger.Picture = ICON_SAFE.Picture
Case 5
Progress.Value = 100
LabelSafe.ForeColor = &H8000&
LabelSafe.Caption = "°²È«³Ì¶È£º¸ß"
LabelTip.Caption = "»¶Ó­Ê¹ÓÃ¹â×Ó·ÀÓùÍø  ÊµÊ±·À»¤È«²¿¿ªÆô"
SafeDanger.Picture = ICON_SAFE.Picture
End Select
End Function

Public Function DoScan(ByVal Op_Sim As Boolean, ByVal Op_AllDisk As Boolean, ByVal ScanTarget As String)
'Op_Sim:ÊÇ·ñ¿ìËÙÉ¨Ãè
'Op_AllDisk:ÊÇ·ñÈ«ÅÌÉ¨Ãè
'ScanTarget:É¨ÃèÂ·¾¶

If exitproc("ScanMod.exe") = False Then
Shell App.Path & "\ScanMod.exe"
Do Until exitproc("ScanMod.exe") = True
SuperSleep 1
Loop
End If

 Dim Way As String
 Dim Path As String
 Dim drivename As Drive
 If Op_Sim = True Then
  '¸ùÄ¿Â¼É¨Ãè
  Way = "Sim"
 Else
  'Éî²ã´ÎÉ¨Ãè
  Way = "Adv"
 End If
 
 If Op_AllDisk = True Then
   strShare = ""
   SuperSleep 1
   strShare = "ScanMod.Scan." & Way & "AllDisk"
 Else
  If Right(ScanTarget, 1) = "\" Then
   Path = Left(ScanTarget, Len(ScanTarget) - 1)
   strShare = ""
   SuperSleep 1
   strShare = "ScanMod.Scan." & Way & Path
   Debug.Print "ScanMod.Scan." & Way & Path
   Else
   Path = ScanTarget
   strShare = ""
   SuperSleep 1
   strShare = "ScanMod.Scan." & Way & Path
   Debug.Print "ScanMod.Scan." & Way & Path
   End If
 End If

End Function
