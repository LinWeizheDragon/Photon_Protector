VERSION 5.00
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.3#0"; "Codejock.SkinFramework.v15.3.1.ocx"
Object = "{2B4B5F62-B44F-4B34-A682-587182855142}#1.0#0"; "SFTabControl.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "龙盾"
   ClientHeight    =   9015
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   12015
   Icon            =   "frmMainForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMainForm.frx":0CCA
   ScaleHeight     =   9015
   ScaleWidth      =   12015
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer ScrollControl 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   6480
      Top             =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   4200
      TabIndex        =   54
      Top             =   360
      Width           =   1215
   End
   Begin VB.Timer Auto 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   480
      Top             =   720
   End
   Begin 龙盾.jcbutton jcbutton1 
      Height          =   1335
      Left            =   10680
      TabIndex        =   4
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2355
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      Caption         =   "退出"
      ForeColor       =   33023
      MousePointer    =   4
      Picture         =   "frmMainForm.frx":16060C
      PictureAlign    =   6
   End
   Begin VB.Timer InitTimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame MainFrame 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6495
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   8880
      Width           =   11775
      Begin 龙盾.jcbutton jcbutton5 
         Height          =   1455
         Left            =   600
         TabIndex        =   26
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2566
         ButtonStyle     =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "龙盾强制删除文件"
         MousePointer    =   99
         MouseIcon       =   "frmMainForm.frx":163866
         Picture         =   "frmMainForm.frx":1639C8
         PictureAlign    =   5
      End
      Begin 龙盾.jcbutton jcbutton6 
         Height          =   1455
         Left            =   2100
         TabIndex        =   27
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2566
         ButtonStyle     =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "龙盾进程管理器"
         MousePointer    =   99
         MouseIcon       =   "frmMainForm.frx":166632
         Picture         =   "frmMainForm.frx":166794
         PictureAlign    =   5
      End
      Begin 龙盾.jcbutton jcbutton8 
         Height          =   1455
         Left            =   3600
         TabIndex        =   38
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2566
         ButtonStyle     =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "U盘百宝箱"
         MousePointer    =   99
         MouseIcon       =   "frmMainForm.frx":1693FE
         Picture         =   "frmMainForm.frx":169560
         PictureAlign    =   5
      End
      Begin 龙盾.jcbutton jcbutton9 
         Height          =   1455
         Left            =   5100
         TabIndex        =   39
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2566
         ButtonStyle     =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "金山毒霸U盘专杀"
         MousePointer    =   99
         MouseIcon       =   "frmMainForm.frx":16C1CA
         Picture         =   "frmMainForm.frx":16C32C
         PictureAlign    =   5
      End
      Begin 龙盾.jcbutton jcbutton10 
         Height          =   1455
         Left            =   6600
         TabIndex        =   40
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2566
         ButtonStyle     =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "360U盘专杀"
         MousePointer    =   99
         MouseIcon       =   "frmMainForm.frx":16EF96
         Picture         =   "frmMainForm.frx":16F0F8
         PictureAlign    =   5
      End
      Begin 龙盾.jcbutton jcbutton11 
         Height          =   1455
         Left            =   8100
         TabIndex        =   41
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2566
         ButtonStyle     =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "更多小工具"
         MousePointer    =   99
         MouseIcon       =   "frmMainForm.frx":171D62
         Picture         =   "frmMainForm.frx":171EC4
         PictureAlign    =   5
      End
   End
   Begin VB.Frame MainFrame 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   8880
      Width           =   11775
      Begin 龙盾.jcbutton jcbutton4 
         Height          =   1575
         Left            =   7440
         TabIndex        =   25
         Top             =   4080
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   2778
         ButtonStyle     =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "扫描！"
         Picture         =   "frmMainForm.frx":174B2E
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "目标"
         Height          =   2055
         Left            =   2160
         TabIndex        =   17
         Top             =   720
         Width           =   8895
         Begin VB.OptionButton Op_Path 
            BackColor       =   &H00FFFFFF&
            Caption         =   "指定路径"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   840
            Width           =   1335
         End
         Begin VB.OptionButton Op_AllDisk 
            BackColor       =   &H00FFFFFF&
            Caption         =   "全盘扫描"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   480
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton Cmd_Explorer 
            Caption         =   "浏览"
            Height          =   300
            Left            =   7320
            TabIndex        =   20
            Top             =   1560
            Width           =   975
         End
         Begin VB.TextBox Text1 
            Height          =   300
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   1200
            Width           =   8055
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "扫描目标："
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "选项"
         Height          =   1095
         Left            =   600
         TabIndex        =   15
         Top             =   2880
         Width           =   10455
         Begin VB.OptionButton Op_Sim 
            BackColor       =   &H00FFFFFF&
            Caption         =   "根目录扫描（速度快，扫描目标下根目录的文件）"
            Height          =   375
            Left            =   600
            TabIndex        =   23
            Top             =   600
            Width           =   5295
         End
         Begin VB.OptionButton Op_Adv 
            BackColor       =   &H00FFFFFF&
            Caption         =   "深层次扫描（速度慢、扫描目录下所有文件）"
            Height          =   375
            Left            =   600
            TabIndex        =   16
            Top             =   240
            Value           =   -1  'True
            Width           =   5295
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1920
         Left            =   120
         Picture         =   "frmMainForm.frx":17AF60
         ScaleHeight     =   1920
         ScaleWidth      =   1920
         TabIndex        =   13
         Top             =   120
         Width           =   1920
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "扫描计算机"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   735
         Left            =   2120
         TabIndex        =   14
         Top             =   0
         Width           =   5415
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "扫描计算机"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   2160
         TabIndex        =   24
         Top             =   0
         Width           =   5415
      End
   End
   Begin VB.Frame MainFrame 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6015
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   11775
      Begin VB.PictureBox Pic_WMI 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4440
         MouseIcon       =   "frmMainForm.frx":1877AA
         MousePointer    =   99  'Custom
         Picture         =   "frmMainForm.frx":1878FC
         ScaleHeight     =   330
         ScaleWidth      =   1155
         TabIndex        =   7
         Top             =   2400
         Width           =   1155
      End
      Begin VB.PictureBox Pic_Ring0 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4440
         MouseIcon       =   "frmMainForm.frx":188264
         MousePointer    =   99  'Custom
         Picture         =   "frmMainForm.frx":1883B6
         ScaleHeight     =   330
         ScaleWidth      =   1155
         TabIndex        =   8
         Top             =   2880
         Width           =   1155
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   7200
         Picture         =   "frmMainForm.frx":188D1E
         ScaleHeight     =   2055
         ScaleWidth      =   3495
         TabIndex        =   28
         Top             =   480
         Width           =   3495
      End
      Begin VB.PictureBox Pic_USB 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4440
         MouseIcon       =   "frmMainForm.frx":1B8B08
         MousePointer    =   99  'Custom
         Picture         =   "frmMainForm.frx":1B8C5A
         ScaleHeight     =   330
         ScaleWidth      =   1155
         TabIndex        =   12
         Top             =   1080
         Width           =   1155
      End
      Begin VB.PictureBox Pic_Reg 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4440
         MouseIcon       =   "frmMainForm.frx":1B95C2
         MousePointer    =   99  'Custom
         Picture         =   "frmMainForm.frx":1B9714
         ScaleHeight     =   330
         ScaleWidth      =   1155
         TabIndex        =   10
         Top             =   3360
         Width           =   1155
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "杜绝危险操作"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   255
         Left            =   1680
         TabIndex        =   36
         Top             =   2040
         Width           =   4335
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "系统防御"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   495
         Left            =   240
         TabIndex        =   35
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "坚守入侵系统的大门"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   255
         Left            =   1800
         TabIndex        =   34
         Top             =   720
         Width           =   4335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "入口防御"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   495
         Left            =   240
         TabIndex        =   33
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "龙盾 实时防护"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   975
         Left            =   1440
         TabIndex        =   31
         Top             =   0
         Width           =   5655
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "龙盾 实时防护"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   975
         Left            =   1500
         TabIndex        =   32
         Top             =   30
         Width           =   5655
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "初级实时防护（WMI）"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         MouseIcon       =   "frmMainForm.frx":1BA07C
         MousePointer    =   99  'Custom
         TabIndex        =   30
         Top             =   2420
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "高级实时防护（Ring0）"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         MouseIcon       =   "frmMainForm.frx":1BA1CE
         MousePointer    =   99  'Custom
         TabIndex        =   29
         Top             =   2900
         Width           =   2775
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "U盘安全防护"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         MouseIcon       =   "frmMainForm.frx":1BA320
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   1095
         Width           =   3015
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "注册表防护"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   9
         Top             =   3380
         Width           =   3015
      End
   End
   Begin 龙盾.jcbutton jcbutton7 
      Height          =   1335
      Left            =   9360
      TabIndex        =   37
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2355
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      Caption         =   "关于"
      ForeColor       =   16711680
      MousePointer    =   4
      Picture         =   "frmMainForm.frx":1BA472
      PictureAlign    =   6
   End
   Begin VB.Frame MainFrame 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "U盘插入防护"
      Height          =   6375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   2040
      Width           =   11775
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   2895
         Left            =   5520
         TabIndex        =   48
         Top             =   0
         Width           =   6135
         Begin VB.PictureBox Picture3 
            BorderStyle     =   0  'None
            Height          =   2895
            Left            =   0
            Picture         =   "frmMainForm.frx":1BD6CC
            ScaleHeight     =   2895
            ScaleWidth      =   6135
            TabIndex        =   51
            Top             =   0
            Width           =   6135
            Begin VB.TextBox LogTextShow 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   2700
               Left            =   360
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   3  'Both
               TabIndex        =   52
               Text            =   "frmMainForm.frx":1F7438
               Top             =   100
               Width           =   5700
            End
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "信息"
         Height          =   2295
         Left            =   4680
         TabIndex        =   45
         Top             =   3240
         Width           =   5895
         Begin MSComctlLib.ListView InfoView 
            Height          =   1815
            Left            =   240
            TabIndex        =   55
            Top             =   240
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   3201
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin 龙盾.jcbutton jcbutton3 
         Height          =   2175
         Left            =   -120
         TabIndex        =   42
         ToolTipText     =   "实时防护"
         Top             =   120
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   3836
         ButtonStyle     =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "实时防护"
         MousePointer    =   99
         MouseIcon       =   "frmMainForm.frx":1F743E
         Picture         =   "frmMainForm.frx":1F75A0
         PictureAlign    =   6
      End
      Begin 龙盾.jcbutton jcbutton2 
         Height          =   2175
         Left            =   0
         TabIndex        =   43
         ToolTipText     =   "扫描计算机"
         Top             =   3600
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   3836
         ButtonStyle     =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "扫描"
         MousePointer    =   99
         MouseIcon       =   "frmMainForm.frx":1FD9D2
         Picture         =   "frmMainForm.frx":1FDB34
         PictureAlign    =   6
      End
      Begin 龙盾.jcbutton jcbutton12 
         Height          =   2175
         Left            =   2520
         TabIndex        =   44
         ToolTipText     =   "实用工具"
         Top             =   3360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   3836
         ButtonStyle     =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "实用工具"
         MousePointer    =   99
         MouseIcon       =   "frmMainForm.frx":203F66
         Picture         =   "frmMainForm.frx":2040C8
         PictureAlign    =   6
      End
      Begin 龙盾.jcbutton Btn_Process1 
         Height          =   495
         Left            =   2160
         TabIndex        =   46
         Top             =   0
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   1085
         ButtonStyle     =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "进程创建防护"
         Picture         =   "frmMainForm.frx":20A4FA
         PictureAlign    =   2
         CaptionAlign    =   0
      End
      Begin 龙盾.jcbutton Btn_Drive1 
         Height          =   495
         Left            =   2160
         TabIndex        =   47
         Top             =   990
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   873
         ButtonStyle     =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "驱动加载防护"
         Picture         =   "frmMainForm.frx":20BA4C
         PictureAlign    =   2
         CaptionAlign    =   0
      End
      Begin 龙盾.jcbutton Btn_Reg1 
         Height          =   495
         Left            =   2160
         TabIndex        =   49
         Top             =   1485
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   873
         ButtonStyle     =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "注册表写入防护"
         Picture         =   "frmMainForm.frx":20CF9E
         PictureAlign    =   2
         CaptionAlign    =   0
      End
      Begin 龙盾.jcbutton Btn_USB1 
         Height          =   495
         Left            =   2160
         TabIndex        =   50
         Top             =   1980
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   873
         ButtonStyle     =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "U盘插入防护"
         Picture         =   "frmMainForm.frx":20E4F0
         PictureAlign    =   2
         CaptionAlign    =   0
      End
      Begin 龙盾.jcbutton Btn_File1 
         Height          =   495
         Left            =   2160
         TabIndex        =   53
         Top             =   495
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   873
         ButtonStyle     =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "文件系统防护"
         Picture         =   "frmMainForm.frx":20FA42
         PictureAlign    =   2
         CaptionAlign    =   0
      End
      Begin 龙盾.jcbutton Btn_Process2 
         Height          =   495
         Left            =   2160
         TabIndex        =   56
         Top             =   0
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   1085
         ButtonStyle     =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "进程创建防护"
         Picture         =   "frmMainForm.frx":210F94
         PictureAlign    =   2
         CaptionAlign    =   0
      End
      Begin 龙盾.jcbutton Btn_File2 
         Height          =   495
         Left            =   2160
         TabIndex        =   57
         Top             =   495
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   873
         ButtonStyle     =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "文件系统防护"
         Picture         =   "frmMainForm.frx":2124E6
         PictureAlign    =   2
         CaptionAlign    =   0
      End
      Begin 龙盾.jcbutton Btn_Drive2 
         Height          =   495
         Left            =   2160
         TabIndex        =   58
         Top             =   990
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   873
         ButtonStyle     =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "驱动加载防护"
         Picture         =   "frmMainForm.frx":213A38
         PictureAlign    =   2
         CaptionAlign    =   0
      End
      Begin 龙盾.jcbutton Btn_Reg2 
         Height          =   495
         Left            =   2160
         TabIndex        =   59
         Top             =   1485
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   873
         ButtonStyle     =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "注册表写入防护"
         Picture         =   "frmMainForm.frx":214F8A
         PictureAlign    =   2
         CaptionAlign    =   0
      End
      Begin 龙盾.jcbutton Btn_USB2 
         Height          =   495
         Left            =   2160
         TabIndex        =   60
         Top             =   1980
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   873
         ButtonStyle     =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "U盘插入防护"
         Picture         =   "frmMainForm.frx":2164DC
         PictureAlign    =   2
         CaptionAlign    =   0
      End
      Begin VB.Line ShowLine 
         BorderColor     =   &H000080FF&
         BorderStyle     =   2  'Dash
         BorderWidth     =   3
         X1              =   2040
         X2              =   2040
         Y1              =   120
         Y2              =   2400
      End
   End
   Begin SFTabControlPro.SFTabControl MainTab 
      Height          =   7095
      Left            =   0
      Top             =   1560
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   12515
   End
   Begin VB.Image Img_ON 
      Height          =   330
      Left            =   5880
      Picture         =   "frmMainForm.frx":217A2E
      Top             =   480
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Image Img_OFF 
      Height          =   330
      Left            =   5880
      Picture         =   "frmMainForm.frx":218354
      Top             =   120
      Visible         =   0   'False
      Width           =   1155
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   0
      Top             =   8520
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label lbl1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Info1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   8700
      Width           =   7215
   End
   Begin VB.Menu mnuTray 
      Caption         =   "trayMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "显示/隐藏主窗口"
      End
      Begin VB.Menu mnuSwp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "退出"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONONFO) As Long

Private Type OSVERSIONONFO
dwOSVersionInfoSize As Long
dwMajorVersion As Long
dwMinorVersion As Long
dwBuildNumber As Long
dwPlatformld As Long
dwCSDVersion As String * 128
End Type



Const TH32CS_SNAPHEAPLIST = &H1
Const TH32CS_SNAPPROCESS = &H2
Const TH32CS_SNAPTHREAD = &H4
Const TH32CS_SNAPMODULE = &H8
Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Const TH32CS_INHERIT = &H80000000
Const MAX_PATH As Integer = 260
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long


Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

  Private Declare Function ReleaseCapture Lib "user32" () As Long
  Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
  Private Const HTCAPTION = 2
  Private Const WM_NCLBUTTONDOWN = &HA1







Dim Check As Boolean
Public strShare As New CSharedString

'=================正式部分=================
Public Function SystemVer() As Variant
Dim Osinfor As OSVERSIONONFO, StrOsName As String
Osinfor.dwOSVersionInfoSize = Len(Osinfor)
GetVersionEx Osinfor
Select Case Osinfor.dwPlatformld
       Case 0
            StrOsName = "Windows 32s"
       Case 1
          Select Case Osinfor.dwMinorVersion
                 Case 0
                      StrOsName = "Windows 95"
                 Case 10
                      StrOsName = "Windows 98"
                 Case 90
                      StrOsName = "Windows Mellinnium"
          End Select
       Case 2
          Select Case Osinfor.dwMajorVersion
                 Case 3
                      StrOsName = "WindowsNT 3.51"
                 Case 4
                      StrOsName = "WindowsNT 4.0"
                 Case 5
                      Select Case Osinfor.dwMinorVersion
                             Case 0
                                  StrOsName = "Windows 2000"
                             Case 1
                                  StrOsName = "Windows XP"
                             Case 2
                                  StrOsName = "Windows 2003"
                      End Select
                 Case 6
                      Select Case Osinfor.dwMinorVersion
                             Case 0
                                  StrOsName = "Windows Vista"
                             Case 1
                                  StrOsName = "Windows 7"
                      End Select
         End Select
       Case Else
            StrOsName = "未知系统版本"
       End Select
       SystemVer = StrOsName
End Function
Public Function exitproc(ByVal exefile As String) As Boolean
Dim r
exitproc = False
    Dim hSnapShot As Long, uProcess As PROCESSENTRY32
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
    uProcess.dwSize = Len(uProcess)
    r = Process32First(hSnapShot, uProcess)
       Do While r
        If Left$(uProcess.szExeFile, IIf(InStr(1, uProcess.szExeFile, Chr$(0)) > 0, InStr(1, uProcess.szExeFile, Chr$(0)) - 1, 0)) = exefile Then
        exitproc = True
        Exit Do
        End If
        'Retrieve information about the next process recorded in our system snapshot
        r = Process32Next(hSnapShot, uProcess)
    Loop
End Function












Private Sub Auto_Timer()
'If exitproc("ProcessRTA.exe") = False Then '如果没开启
'Shell App.Path & "\ProcessRTA.exe"
'End If
End Sub

Private Sub Btn_Drive1_Click()
ReadLogFile 3
MoveText
End Sub

Private Sub Btn_Drive2_Click()
ReadLogFile 3
MoveText
End Sub

Private Sub Btn_File1_Click()
ReadLogFile 2
MoveText
End Sub

Private Sub Btn_File2_Click()
ReadLogFile 2
MoveText
End Sub

Private Sub Btn_Process1_Click()
ReadLogFile 1
MoveText
End Sub

Private Sub Btn_Process2_Click()
ReadLogFile 1
MoveText
End Sub

Private Sub Btn_Reg1_Click()
ReadLogFile 4
MoveText
End Sub

Private Sub Btn_Reg2_Click()
ReadLogFile 4
MoveText
End Sub

Private Sub Btn_USB1_Click()
ReadLogFile 5
MoveText
End Sub

Private Sub Btn_USB2_Click()
ReadLogFile 5
MoveText
End Sub

Private Sub Cmd_Explorer_Click()
Dim FilePath As String, i As Long
    ShowDir Me.hwnd, FilePath
    '加载目录，使用自定义函数ShowDir
    Text1.Text = FilePath
End Sub





Private Sub Command1_Click()
jcbutton18.Picture = LoadPicture(App.Path & "\Res\On1.ico")
End Sub

Private Sub Form_Activate()
Call rgnform(Me, 10, 10) '调用子过程

End Sub

Private Sub Form_Load()
If Command = "-start" Then
Me.Hide
End If
Check = False
strShare.Create "TestShareString" '内存共享通信，类建立
Call CreatTray(Me, "龙盾", "龙盾", "程序正在加载……", 4)
'-------------皮肤控件加载----------------
Dim FileName As String
Dim IniFile As String
FileName = App.Path & "\Skin\Office2007.cjstyles"
IniFile = "NormalBlue.ini"
SkinFramework1.LoadSkin FileName, IniFile
SkinFramework1.ApplyWindow Me.hwnd
SkinFramework1.ApplyOptions = SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics
InitTimer.Enabled = True
Dim i
For i = 0 To MainFrame.Count - 1
MainFrame(i).Visible = False
Next
MainFrame(0).Visible = True
Me.Caption = "龙盾  " & App.Major & "." & App.Minor & "." & App.Revision
Info1.Caption = "主程序版本：" & App.Major & "." & App.Minor & "." & App.Revision _
& "    文件病毒库版本：" & ReadString("Ver", "Ver", App.Path & "\FileData\Version.ini") _
& "    进程病毒库版本：" & ReadString("Ver", "Ver", App.Path & "\ProcessData\Version.ini")
Load_Main = True

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 '拖动窗体
  If Button = 1 Then
  ReleaseCapture
  SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
  End If
End Sub




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
Cancel = 1
Me.Hide
Exit Sub
End If
If MsgBox("您正在退出龙盾主程序，一旦退出将关闭所有防护功能，您确认退出吗？", vbYesNo, "龙盾-退出提示") = vbNo Then
Cancel = 1
Exit Sub
End If
Me.Hide
Call ShowTip("龙盾", "程序正在关闭相关进程，完成后自动退出托盘......", 4)
If exitproc("USBRTA.exe") = True Then
  strShare = "USBRTA"
  SuperSleep 1
  strShare = "USBRTA.Close"
  SuperSleep 1
  AddInfo "U盘实时防护关闭......"
End If
If frmWatch.Working = True Then
Unload frmWatch
End If
If exitproc("ProcessRTA.exe") = True Then
  Auto.Enabled = False
  strShare = "ProcessRTA"
  SuperSleep 1
  strShare = "ProcessRTA.Unload"
  SuperSleep 1
  AddInfo "高级实时防护关闭......"
End If
If exitproc("RegRTA.exe") = True Then
  strShare = "RegRTA"
  SuperSleep 1
  strShare = "RegRTA.Unload"
  AddInfo "注册表实时防护关闭......"
End If


AddInfo "程序开始退出，保存日志......"
Dim MyFSO As New FileSystemObject
Dim DataNum As Integer
Dim LogStr As String
Dim LogNum As Integer
LogStr = "龙盾主程序日志" & vbCrLf & "生成时间：" & Now & vbCrLf & _
"-------------------------" & vbCrLf
LogNum = InfoView.ListItems.Count
Do Until LogNum = 0
LogStr = LogStr & InfoView.ListItems(LogNum).Text & ":" & InfoView.ListItems(LogNum).SubItems(1) & vbCrLf
LogNum = LogNum - 1
Loop
DataNum = MyFSO.GetFolder(App.Path & "\Data\").Files.Count + 1
Open App.Path & "\Data\龙盾日志-" & DataNum & ".log" For Append As #1
Print #1, LogStr
Close #1


If UnloadMode = vbAppWindows Then
End
End If
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then
Me.WindowState = vbNormal
Me.Hide
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'DeleteObject outrgn '将圆角区域使用的所有系统资源释放

Dim fm As Form
For Each fm In Forms
Unload fm
Next
End Sub

Private Sub InitTimer_Timer()

InitTimer.Enabled = False
End Sub

Private Sub jcbutton1_Click()
Unload Me
End Sub

Private Sub jcbutton10_Click()
If Dir(App.Path & "\Tools\USBTools\killer_autorun.exe", vbNormal Or vbHidden Or vbSystem Or vbReadOnly) <> "" Then
Shell App.Path & "\Tools\USBTools\killer_autorun.exe", vbNormalFocus
Else
MsgBox "未找到此小工具，有可能是安装不完整，请重新安装或还原程序。", vbInformation, "未找到小工具"
End If
End Sub

Private Sub jcbutton11_Click()
If Dir(App.Path & "\Tools\USBTools\U盘工具合集", vbDirectory Or vbNormal Or vbHidden Or vbSystem Or vbReadOnly) <> "" Then
Shell "Explorer.exe " & App.Path & "\Tools\USBTools\U盘工具合集", vbNormalFocus
Else
MsgBox "未找到此小工具，有可能是安装不完整，请重新安装或还原程序。", vbInformation, "未找到小工具"
End If
End Sub

Private Sub jcbutton12_Click()
MainTab.DefaultTab = 3
End Sub



Private Sub jcbutton13_Click()
ReadLogFile 1
MoveText
End Sub

Private Sub jcbutton14_Click()
ReadLogFile 2
MoveText
End Sub

Private Sub jcbutton15_Click()
ReadLogFile 3
MoveText
End Sub

Private Sub jcbutton16_Click()
ReadLogFile 4
MoveText
End Sub

Private Sub jcbutton17_Click()
ReadLogFile 5
MoveText
End Sub

Private Sub jcbutton18_Click()
ReRead
End Sub

Private Sub jcbutton2_Click()
MainTab.DefaultTab = 2
End Sub

Private Sub jcbutton3_Click()
MainTab.DefaultTab = 1
End Sub

Private Sub jcbutton4_Click()
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
  '根目录扫描
  Way = "Sim"
 Else
  '深层次扫描
  Way = "Adv"
 End If
 
 If Op_AllDisk.Value = True Then
   strShare = ""
   SuperSleep 1
   strShare = "ScanMod.Scan." & Way & "AllDisk"
   AddInfo "启动全盘扫描......"
 Else
  If Right(Text1.Text, 1) = "\" Then
   Path = Left(Text1.Text, Len(Text1.Text) - 1)
   strShare = ""
   SuperSleep 1
   strShare = "ScanMod.Scan." & Way & Path
   Debug.Print "ScanMod.Scan." & Way & Path
   AddInfo "启动对""" & Path & """的扫描......"
   Else
   Path = Text1.Text
   strShare = ""
   SuperSleep 1
   strShare = "ScanMod.Scan." & Way & Path
   Debug.Print "ScanMod.Scan." & Way & Path
   AddInfo "启动对""" & Path & """的扫描......"
   End If
 End If



End Sub

Private Sub jcbutton5_Click()
If Dir(App.Path & "\Tools\KillFiles\KillFile.exe", vbNormal Or vbHidden Or vbSystem Or vbReadOnly) <> "" Then
Shell App.Path & "\Tools\KillFiles\KillFile.exe", vbNormalFocus
Else
MsgBox "未找到此小工具，有可能是安装不完整，请重新安装或还原程序。", vbInformation, "未找到小工具"
End If
End Sub

Private Sub jcbutton6_Click()
If Dir(App.Path & "\Tools\ProcessMonitor\ProcessMonitor.exe", vbNormal Or vbHidden Or vbSystem Or vbReadOnly) <> "" Then
Shell App.Path & "\Tools\ProcessMonitor\ProcessMonitor.exe", vbNormalFocus
Else
MsgBox "未找到此小工具，有可能是安装不完整，请重新安装或还原程序。", vbInformation, "未找到小工具"
End If
End Sub

Private Sub jcbutton7_Click()
frmAbout.Show
End Sub

Private Sub jcbutton8_Click()
If Dir(App.Path & "\Tools\USBTools\upbbx.exe", vbNormal Or vbHidden Or vbSystem Or vbReadOnly) <> "" Then
Shell App.Path & "\Tools\USBTools\upbbx.exe", vbNormalFocus
Else
MsgBox "未找到此小工具，有可能是安装不完整，请重新安装或还原程序。", vbInformation, "未找到小工具"
End If
End Sub

Private Sub jcbutton9_Click()
If Dir(App.Path & "\Tools\USBTools\kavudisk.exe", vbNormal Or vbHidden Or vbSystem Or vbReadOnly) <> "" Then
Shell App.Path & "\Tools\USBTools\kavudisk.exe", vbNormalFocus
Else
MsgBox "未找到此小工具，有可能是安装不完整，请重新安装或还原程序。", vbInformation, "未找到小工具"
End If
End Sub

Private Sub Label1_Click()
frmWatch.Show
frmWatch.Visible = True
End Sub

Private Sub Label2_Click()
If exitproc("ProcessRTA.exe") = True Then
  strShare = "ProcessRTA"
  SuperSleep 1
 strShare = "ProcessRTA.Show"
End If
End Sub



Private Sub Label4_Click()
'frmUSBOption.Show
End Sub

Private Sub MainFrame_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Check = True Then Exit Sub
If exitproc("ProcessRTA.exe") = True Then
  ChangeIcon True, False, Pic_Ring0
Else
  ChangeIcon False, False, Pic_Ring0
End If
If exitproc("RegRTA.exe") = True Then
  ChangeIcon True, False, Pic_Reg
Else
  ChangeIcon False, False, Pic_Reg
End If
If exitproc("USBRTA.exe") = True Then
  ChangeIcon True, False, Pic_USB
Else
  ChangeIcon False, False, Pic_USB
End If
If frmWatch.Working = True Then
  ChangeIcon True, False, Pic_WMI
Else
  ChangeIcon False, False, Pic_WMI
End If

Check = True
End Sub

Private Sub MainTab_ChangeTab(ByVal dwCurIndex As Long)
On Error Resume Next
Dim i
For i = 0 To MainFrame.Count - 1
MainFrame(i).Visible = False
Next
MainFrame(dwCurIndex).Visible = True
If dwCurIndex = 1 Then
'如果是打开防火墙界面
ReRead
End If
If dwCurIndex = 2 Then
Cmd_Explorer.Enabled = Op_Path.Value
End If
End Sub

Private Function ChangeIcon(ByVal OnOff As Boolean, ByVal Focus As Boolean, ByRef Picturebox As Picturebox)
'-------------改变图标状态----------
If OnOff = True Then
  If Focus = True Then
    Picturebox.Picture = LoadPicture(App.Path & "\Res\On1.ico")
  Else
    Picturebox.Picture = LoadPicture(App.Path & "\Res\On2.ico")
  End If
Else
  If Focus = True Then
    Picturebox.Picture = LoadPicture(App.Path & "\Res\Off1.ico")
  Else
    Picturebox.Picture = LoadPicture(App.Path & "\Res\Off2.ico")
  End If
End If
End Function

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuShow_Click()
frmMain.Show
End Sub

Private Sub Op_AllDisk_Click()
Cmd_Explorer.Enabled = False
End Sub

Private Sub Op_Path_Click()
Cmd_Explorer.Enabled = True
End Sub

Private Sub Pic_Reg_Click()

If exitproc("RegRTA.exe") = True Then
  strShare = "RegRTA"
  SuperSleep 1
  strShare = "RegRTA.Unload"
  Call WriteString("Main", "RegRTA", 0, IniPath)
  AddInfo "注册表实时防护关闭......"
Else
  'ChangeIcon False, True, Pic_Reg
  '开启模块
  If Not Dir(App.Path & "\RegRTA.exe") = "" Then
  Shell App.Path & "\RegRTA.exe"
  End If
  Call WriteString("Main", "RegRTA", 1, IniPath)
  AddInfo "注册表实时防护关闭......"
End If
SuperSleep 0.5
ReRead
End Sub

Private Sub Pic_Reg_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If exitproc("RegRTA.exe") = True Then
  ChangeIcon True, True, Pic_Reg
Else
  ChangeIcon False, True, Pic_Reg
End If
Check = False
End Sub

Private Sub Pic_Ring0_Click()
If SystemVer <> "Windows XP" Then
 If MsgBox("您的操作系统是 " & SystemVer & " ，目前未测试兼容性" & _
 "您一定要继续操作吗？有可能造成蓝屏。", vbOKCancel, "操作系统兼容性未明") = vbCancel Then
 Exit Sub
End If
End If
If exitproc("ProcessRTA.exe") = True Then
  Auto.Enabled = False
  strShare = "ProcessRTA"
  SuperSleep 1
  strShare = "ProcessRTA.Unload"
  Call WriteString("Main", "ProcessRTA", 0, IniPath)
  AddInfo "高级实时防护关闭......"
Else
   Auto.Enabled = True
  ChangeIcon False, True, Pic_Ring0
  '开启模块
  If Not Dir(App.Path & "\ProcessRTA.exe") = "" Then
  Shell App.Path & "\ProcessRTA.exe", vbNormalFocus
  End If
  Call WriteString("Main", "ProcessRTA", 1, IniPath)
  AddInfo "高级实时防护开启......"
End If
ReRead
End Sub

Private Sub Pic_Ring0_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If exitproc("ProcessRTA.exe") = True Then
  ChangeIcon True, True, Pic_Ring0
Else
  ChangeIcon False, True, Pic_Ring0
End If
Check = False
End Sub

Private Sub Pic_USB_Click()

If exitproc("USBRTA.exe") = True Then
  strShare = "USBRTA"
  SuperSleep 1
  strShare = "USBRTA.Close"
  Call WriteString("Main", "USBRTA", 0, IniPath)
  AddInfo "U盘实时防护关闭......"
Else
  'ChangeIcon False, True, Pic_Reg
  '开启模块
  If Not Dir(App.Path & "\USBRTA.exe") = "" Then
  Shell App.Path & "\USBRTA.exe"
  End If
  Call WriteString("Main", "USBRTA", 1, IniPath)
  AddInfo "U盘实时防护开启......"
End If
SuperSleep 0.5
ReRead
End Sub

Private Sub Pic_USB_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If exitproc("USBRTA.exe") = True Then
  ChangeIcon True, True, Pic_USB
Else
  ChangeIcon False, True, Pic_USB
End If
Check = False
End Sub

Private Sub Pic_WMI_Click()
If frmWatch.Working = True Then
Unload frmWatch
Call WriteString("Main", "WMI", 0, IniPath)
Else
  If frmWatch Is Nothing Then
    frmWatch.Show
  Else
    AddInfo "初始化初级实时防护……"
    Unload frmWatch
    frmWatch.Show
  End If
Call WriteString("Main", "WMI", 1, IniPath)
End If
ReRead
End Sub

Private Sub Pic_WMI_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If frmWatch.Working = True Then
 ChangeIcon True, True, Pic_WMI
Else
 ChangeIcon False, True, Pic_WMI
End If
Check = False
End Sub

Public Sub ReRead()
If frmWatch.Working = True Then
  ChangeIcon True, False, Pic_WMI
  Btn_Process2.Visible = True
  Btn_Process1.Visible = False
Else
  ChangeIcon False, False, Pic_WMI
  Btn_Process2.Visible = False
  Btn_Process1.Visible = True
End If
If exitproc("ProcessRTA.exe") = True Then
  ChangeIcon True, False, Pic_Ring0
  Btn_Process2.Visible = True
  Btn_Process1.Visible = False
Else
  ChangeIcon False, False, Pic_Ring0
'  If frmWatch.Working = True Then '如果已经开启过一个
'
'  End If
End If
If exitproc("RegRTA.exe") = True Then
  ChangeIcon True, False, Pic_Reg
  Btn_Reg2.Visible = True
  Btn_Reg1.Visible = False
Else
  ChangeIcon False, False, Pic_Reg
  Btn_Reg2.Visible = False
  Btn_Reg1.Visible = True
End If
If exitproc("USBRTA.exe") = True Then
  ChangeIcon True, False, Pic_USB
  Btn_USB2.Visible = True
  Btn_USB1.Visible = False
Else
  ChangeIcon False, False, Pic_USB
  Btn_USB2.Visible = False
  Btn_USB1.Visible = True
End If
End Sub


'Private Sub ReReadFoces()
'If frmWatch.Working = True Then
'  ChangeIcon True, True, Pic_WMI
'Else
'  ChangeIcon False, True, Pic_WMI
'End If
'If exitproc("ProcessRTA") = True Then
'  ChangeIcon True, True, Pic_Ring0
'Else
'  ChangeIcon False, True, Pic_Ring0
'End If
'End Sub


Public Function ReadLogFile(ByVal LType As Integer)
'On Error GoTo Err:
Dim strText
Dim strPath
Select Case LType

Case 1 '进程日志
strPath = App.Path & "\ProLog.dat"
Case 2
strPath = App.Path & "\FileLog.dat"
Case 3
strPath = App.Path & "\DriveLog.dat"
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
If strText = "" Then strText = "暂无日志"
LogTextShow.Text = strText
Close #1
Else
strText = "暂无日志"
LogTextShow.Text = strText
End If
Exit Function
Err:
strText = "读取错误"
LogTextShow.Text = strText
End Function

Public Function MoveText()

LogTextShow.Top = 0 - LogTextShow.Height
ScrollControl.Enabled = True
End Function

Private Sub Timer1_Timer()

End Sub

Private Sub ScrollControl_Timer()
LogTextShow.Top = LogTextShow.Top + 100

If LogTextShow.Top >= 100 Then
LogTextShow.Top = 100
ScrollControl.Enabled = False
End If
End Sub
