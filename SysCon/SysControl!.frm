VERSION 5.00
Begin VB.Form FrmSysCon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SysControl! 2"
   ClientHeight    =   6525
   ClientLeft      =   -105
   ClientTop       =   1380
   ClientWidth     =   6570
   Icon            =   "SysControl!.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   6570
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnSysInfo 
      Caption         =   "System Info"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      TabIndex        =   47
      Top             =   2520
      Width           =   615
   End
   Begin VB.CheckBox OptOnTop 
      BackColor       =   &H00CCCCCC&
      Caption         =   "Stay on Top"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2520
      Width           =   615
   End
   Begin VB.CheckBox OptRegSS 
      BackColor       =   &H00CCCCCC&
      Caption         =   "Register as screen saver"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1740
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Disable User Control"
      Top             =   2520
      Width           =   855
   End
   Begin VB.CheckBox ChkShowExInfo 
      Caption         =   "Show Extended Program Information"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   39
      Top             =   3210
      Width           =   3075
   End
   Begin VB.Timer TmrMain 
      Interval        =   1
      Left            =   4560
      Top             =   1260
   End
   Begin VB.CommandButton BtnHelp 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   12
      Top             =   2520
      Width           =   435
   End
   Begin VB.CommandButton BtnHighlightHWnd 
      Caption         =   "Highlight HWnds!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   11
      Top             =   2520
      Width           =   735
   End
   Begin VB.Frame FrameSelItem 
      Caption         =   "Selectable Items"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   6440
      Begin VB.CommandButton BtnListParents 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Parents"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1980
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   2100
         Width           =   675
      End
      Begin VB.CommandButton BtnListChildren 
         Caption         =   "Children"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1980
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   1860
         Width           =   675
      End
      Begin VB.CommandButton BtnListAll 
         Caption         =   "HWnd Tree"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Lists all HWnds in a hierarchical list"
         Top             =   1860
         Width           =   555
      End
      Begin VB.CommandButton BtnExtended 
         Caption         =   "&Extended Options"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5580
         TabIndex        =   14
         Top             =   1860
         Width           =   795
      End
      Begin VB.TextBox TxtNewParent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   3180
         MaxLength       =   10
         TabIndex        =   24
         Top             =   1080
         Width           =   795
      End
      Begin VB.CommandButton BtnMe 
         Caption         =   "Me"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4020
         TabIndex        =   23
         ToolTipText     =   "Gets Console's HWnd"
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton BtnLink 
         Caption         =   "Link"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3180
         TabIndex        =   22
         ToolTipText     =   "Move the target to new HWnd"
         Top             =   1380
         Width           =   1215
      End
      Begin VB.CommandButton DragThing 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Drag mouse from here to target!"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   505
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1280
         Width           =   1095
      End
      Begin VB.TextBox TxtTarget 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   279
         Left            =   4430
         MaxLength       =   10
         MultiLine       =   -1  'True
         TabIndex        =   16
         Text            =   "SysControl!.frx":0442
         Top             =   510
         Width           =   855
      End
      Begin VB.CommandButton BtnClose 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   15
         Top             =   2100
         Width           =   555
      End
      Begin VB.ListBox LstParents 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1905
         IntegralHeight  =   0   'False
         ItemData        =   "SysControl!.frx":044D
         Left            =   60
         List            =   "SysControl!.frx":0454
         TabIndex        =   9
         Top             =   420
         Width           =   1335
      End
      Begin VB.CommandButton BtnKill 
         Caption         =   "Kill"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   19
         Top             =   1860
         Width           =   555
      End
      Begin VB.CommandButton BtnCapTarget 
         Caption         =   "Capture Picture"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4380
         TabIndex        =   20
         Top             =   1860
         Width           =   675
      End
      Begin VB.CommandButton BtnHide 
         Caption         =   "&Hide"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   25
         Top             =   2100
         Width           =   555
      End
      Begin VB.CommandButton BtnShow 
         Caption         =   "&Show"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   26
         Top             =   1860
         Width           =   555
      End
      Begin VB.CommandButton BtnInfo 
         Caption         =   "Get/Set Misc."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3180
         TabIndex        =   27
         Top             =   1860
         Width           =   675
      End
      Begin VB.CommandButton BtnSelStart 
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5940
         TabIndex        =   13
         ToolTipText     =   "Shortcut"
         Top             =   660
         Width           =   435
      End
      Begin VB.CommandButton BtnSelTaskBar 
         Caption         =   "Taskbar"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   17
         ToolTipText     =   "Shortcut"
         Top             =   420
         Width           =   615
      End
      Begin VB.CommandButton BtnSelDesktop 
         Caption         =   "Desktop"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5640
         TabIndex        =   18
         ToolTipText     =   "Shortcut"
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton BtnClickError 
         Caption         =   "Lost Capture"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5580
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1860
         Width           =   795
      End
      Begin VB.Label lblDragging 
         BackStyle       =   0  'Transparent
         Caption         =   "Drag mouse to target!"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3180
         TabIndex        =   42
         Top             =   1860
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "or"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   5400
         TabIndex        =   33
         Top             =   870
         Width           =   195
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00808080&
         X1              =   5580
         X2              =   5580
         Y1              =   1140
         Y2              =   1020
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00808080&
         X1              =   5580
         X2              =   5460
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00808080&
         X1              =   5280
         X2              =   5400
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         X1              =   5280
         X2              =   5280
         Y1              =   840
         Y2              =   960
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         X1              =   5640
         X2              =   5640
         Y1              =   780
         Y2              =   900
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         X1              =   5640
         X2              =   5520
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   5460
         X2              =   5640
         Y1              =   1020
         Y2              =   780
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   5310
         X2              =   5575
         Y1              =   870
         Y2              =   1135
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "HWnd Parent:"
         Height          =   255
         Left            =   3120
         TabIndex        =   32
         Top             =   300
         UseMnemonic     =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "HWnd:"
         Height          =   195
         Left            =   4500
         TabIndex        =   31
         Top             =   300
         UseMnemonic     =   0   'False
         Width           =   675
      End
      Begin VB.Label LblHWndP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "{none yet}"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   3300
         TabIndex        =   30
         Top             =   480
         UseMnemonic     =   0   'False
         Width           =   735
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "New Parent:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3180
         TabIndex        =   29
         Top             =   900
         UseMnemonic     =   0   'False
         Width           =   795
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "And press Enter"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   435
         Left            =   4500
         TabIndex        =   28
         Top             =   780
         UseMnemonic     =   0   'False
         Width           =   675
      End
      Begin VB.Shape Light 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   225
         Left            =   5180
         Shape           =   3  'Circle
         Top             =   375
         Width           =   225
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Parents going up:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   180
         UseMnemonic     =   0   'False
         Width           =   1215
      End
      Begin VB.Label LblParentInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "{Choose Parent for Info}"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1395
         Left            =   1440
         TabIndex        =   10
         Top             =   420
         UseMnemonic     =   0   'False
         Width           =   1380
      End
   End
   Begin VB.Frame FrameHotkey 
      Caption         =   "On Ctrl+Alt+A"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   60
      TabIndex        =   3
      Top             =   2460
      Width           =   1575
      Begin VB.OptionButton OptParent 
         Caption         =   "Undo ""New Parent"""
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   6
         Top             =   420
         Width           =   1455
      End
      Begin VB.OptionButton OptCap 
         Caption         =   "Take Screenshot"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   5
         Top             =   180
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.CommandButton BtnCaptureScr 
      Caption         =   "Screen Capture"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4140
      TabIndex        =   2
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton BtnExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   1
      Top             =   2520
      Width           =   465
   End
   Begin VB.Frame FrameProgInfo 
      Caption         =   "Running Programs"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2910
      Left            =   60
      TabIndex        =   34
      Top             =   3540
      Visible         =   0   'False
      Width           =   6435
      Begin VB.TextBox TxtLocation 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   2520
         Width           =   5415
      End
      Begin VB.Frame FramePriority 
         Caption         =   "Priority"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1395
         Left            =   5160
         TabIndex        =   48
         Top             =   240
         Width           =   1095
         Begin VB.OptionButton OptPriority 
            Caption         =   "Idle"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   64
            Left            =   120
            TabIndex        =   52
            Top             =   1090
            Width           =   855
         End
         Begin VB.OptionButton OptPriority 
            Caption         =   "Normal"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   32
            Left            =   120
            TabIndex        =   51
            Top             =   790
            Width           =   915
         End
         Begin VB.OptionButton OptPriority 
            Caption         =   "High"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   128
            Left            =   120
            TabIndex        =   50
            Top             =   490
            Width           =   855
         End
         Begin VB.OptionButton OptPriority 
            Caption         =   "Realtime"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   256
            Left            =   120
            TabIndex        =   49
            Top             =   190
            Width           =   915
         End
      End
      Begin VB.ListBox LstProcess 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1320
         ItemData        =   "SysControl!.frx":0469
         Left            =   120
         List            =   "SysControl!.frx":046B
         TabIndex        =   36
         Top             =   300
         Width           =   3855
      End
      Begin VB.CommandButton BtnActivate 
         Caption         =   "Activate Selected"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4080
         TabIndex        =   41
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton BtnSendText 
         Caption         =   "Send Text to Selected"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4080
         TabIndex        =   40
         Top             =   300
         Width           =   975
      End
      Begin VB.CommandButton BtnProgRefresh 
         Caption         =   "Refresh ListBox"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4080
         TabIndex        =   35
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label LblProcessName 
         BackStyle       =   0  'Transparent
         Caption         =   "Location:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   2550
         UseMnemonic     =   0   'False
         Width           =   795
      End
      Begin VB.Label LblProcess 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "{select a program}"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   120
         TabIndex        =   37
         Top             =   1680
         UseMnemonic     =   0   'False
         Width           =   2175
      End
   End
   Begin VB.Menu MnuMain 
      Caption         =   "Main (only visible in Extended pop-up)"
      Visible         =   0   'False
      Begin VB.Menu MnuGetPos 
         Caption         =   "Get &Position of Target"
      End
      Begin VB.Menu MnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuLock 
         Caption         =   "&Lock Updating"
      End
      Begin VB.Menu MnuUnlock 
         Caption         =   "&Unlock Updating"
      End
      Begin VB.Menu MnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuEnable 
         Caption         =   "&Enable"
      End
      Begin VB.Menu MnuDisable 
         Caption         =   "&Disable"
      End
      Begin VB.Menu MnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuStayOnTopOn 
         Caption         =   "Stay-On-Top On"
      End
      Begin VB.Menu MnuStayOnTopOff 
         Caption         =   "Stay-On-Top Off"
      End
      Begin VB.Menu MnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu MnuTargetMin 
         Caption         =   "&Minimize Target"
      End
      Begin VB.Menu MnuTargetNorm 
         Caption         =   "&Normalize Target"
      End
      Begin VB.Menu MnuTargetMax 
         Caption         =   "Ma&ximize Target"
      End
      Begin VB.Menu MnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFlashWin 
         Caption         =   "&Flash Target's Bar"
      End
      Begin VB.Menu MnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu MnuEditStyleYourSelf 
         Caption         =   "E&dit the Style Yourself!"
      End
      Begin VB.Menu MnuSendMessage 
         Caption         =   "&Send a Message"
      End
   End
End
Attribute VB_Name = "FrmSysCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************
'*  System Control 2.0          *
'*  By Wesley Crossman          *
'*  wesley_crossman@yahoo.com   *
'*  Last Updated: Jan. 2001     *
'*******************************************************
'* If you would like to use this source, feel free to  *
'* do so. All I ask is that you send me a copy of your *
'* program if possible. I would really please me to    *
'* know that my note-in-a-bottle went somewhere! :-)   *
'* Also, if you need help with your project, feel free *
'* to ask me. I have tons of time I don't use and      *
'* would be priviledged to help a fellow programmer.   *
'* Besides, I get bored sometimes!                     *
'*******************************************************
'* I would like to thank the KPD-Team (allapi.net) for *
'* their excellent API-Guide and Dan Appleman for      *
'* writing the "Visual Basic Programmers Guide to the  *
'* Win32 API".                                         *
'*******************************************************

DefLng A-Z

'used for the "Undo New Parent" hotkey
Dim hC, hP
'for enumeration of processes
Dim Process(0 To 300) As PROCESSENTRY32
'Dragging: true if selecting target
Dim Dragging As Byte
'ListType: see FrmSysCon.RefreshList
Dim ListType As Byte
'DragThing
Dim Mouse As POINTAPI

'Set this if you have a large screen, or if you use the
'"Running Programs" tab often.
Const ConstShowProgInfoByDefault = 0
'If you change the forms dimensions, change it here too.
Const ConstMaxHeight = 6900, ConstMinHeight = 3870

Sub SetProcessInfoMode(pMode As Byte)
If pMode Then
 TxtLocation.BackColor = &HFFFFFF
 TxtLocation.Enabled = 1
 FramePriority.Enabled = 1
Else
 LblProcess.Caption = "{select a program}"
 OptPriority(32).Value = 0
 OptPriority(64).Value = 0
 OptPriority(128).Value = 0
 OptPriority(256).Value = 0
 FramePriority.Enabled = 0
 TxtLocation.BackColor = &HE0E0E0
 TxtLocation.Text = ""
 TxtLocation.Enabled = 0
End If
End Sub

'toggles between dragging and normal mode (n=1 on normal)
Sub SetMode(n As Byte)
BtnInfo.Visible = n
BtnShow.Visible = n
BtnHide.Visible = n
BtnCapTarget.Visible = n
BtnKill.Visible = n
BtnClose.Visible = n
BtnExtended.Visible = n
Dragging = n Xor 1
lblDragging.Visible = n Xor 1
End Sub

'This is a recursive sub designed to list all HWnds to LstParents
Sub ListWindows(ByVal Parent&, ByVal Depth%)
On Error GoTo errhandle

Do
 'find next window
 lc = FindWindowEx(Parent, lc, vbNullString, vbNullString)
 'if there isn't another window, it will return 0
 If lc Then
  'add item to LstParents, accounting for the parent/child relationship
  If GetParent(lc) = Parent Then
   'if HWnd is too deeply nested, reduce the indentation for that one
   If Depth > 4 Then overedge = 1 Else overedge = 0
   'to change the level of indentation, alter the "* 3" part
   LstParents.AddItem Space(Depth * 3 - overedge) & lc
   'begin another branch of recursion targeting lc
   ListWindows lc, Depth + 1
  End If
 Else
  'parent has no more children so exit
  Exit Do
 End If
Loop

errhandle:
End Sub

'changes status of prog to show status of target through NewSetting
Sub LightSwitch(NewSetting As Byte)
'NewSetting=1:Green NewSetting=0:Red
Light.FillColor = IIf(NewSetting, &H11FF11, &H5555FF)
TxtNewParent.Enabled = NewSetting
End Sub

'gets list of target's parents and puts them in the listbox
Sub RefreshList()
LstParents.Enabled = 1
LstParents.Clear

Select Case ListType
Case 1
 LblParentInfo = "{Choose Parent for Info}"
 n = HWCustom
 If IsWindow(HWCustom) Then
  For A = 1 To 300
   n = GetParent(n)
   LstParents.AddItem n
   If n = 0 Then Exit Sub
  Next
 Else
  LstParents.AddItem "{Select target}"
  LstParents.Enabled = 0
 End If
Case 2
 LblParentInfo = "{Choose Child for Info}"
 If IsWindow(HWCustom) Then
  ListWindows HWCustom, 0
 Else
  LstParents.AddItem "{Select target}"
  LstParents.Enabled = 0
 End If
Case 3
 LblParentInfo = "{Choose HWnd for Info}"
 ListWindows 0, 0
End Select
End Sub

'capture a picture of screen
Sub CapScr()
FrmViewPic.Cls
'create a device context of the screen
ndc = CreateDC("DISPLAY", vbNullString, vbNullString, ByVal 0&)
'copy BitBlt to capture the screen and copy it to ViewPic
BitBlt FrmViewPic.hdc, 0, 0, ScreenDim.Right, ScreenDim.Bottom, ndc, 0, 0, vbSrcCopy
'delete context to free memory
DeleteDC ndc
FrmViewPic.Show
End Sub

'set the target of the prog to HWNew
Sub SetTarget(ByVal HWNew&)
If IsWindow(HWNew) Then
 HWCustom = HWNew
 TxtTarget = HWNew
 LblHWndP = GetParent(HWNew)
 LightSwitch 1 'on
Else
 HWCustom = 0
 LblHWndP = "{invalid child}"
 LightSwitch 0 'off
End If
End Sub

Private Sub BtnActivate_Click()
On Error GoTo errhandle

If LstProcess.ListIndex = -1 Then Exit Sub
'activate the highlighted program
AppActivate Process(LstProcess.ListIndex).th32ProcessID
Exit Sub

errhandle:
If Err.Number = 5 Then MsgBox "This program cannot be activated." & vbCrLf & "It may not have a window to do so!": Exit Sub
MsgBox Err.Description
End Sub

Private Sub BtnClickError_Click()
SetMode 1 'set prog back to normal
End Sub

Private Sub BtnClose_Click()
If IsWindow(HWCustom) = 0 Then Exit Sub
'check if target is essential, otherwise close
If HWCustom <> hwnd And HWCustom <> BtnExit.hwnd And HWCustom <> DragThing.hwnd And HWCustom <> FrameSelItem.hwnd Then
 'close target
 If MsgBox("Are you sure?", vbYesNo, "Close the Target") = vbYes Then PostMessage HWCustom, &H10, 0, 0
Else
 MsgBox "Sorry, you can't close essential functions."
End If
End Sub

Private Sub BtnCapTarget_Click()
If IsWindow(HWCustom) = 0 Then Exit Sub
FrmViewPic.Cls
'create a device context of the screen
ndc = CreateDC("DISPLAY", vbNullString, vbNullString, ByVal 0&)
'get rectangle of target
GetWindowRect HWCustom, WinDim
With WinDim
 'use BitBlt to capture the target and copy it to ViewPic
 BitBlt FrmViewPic.hdc, 0, 0, .Right - .Left, .Bottom - .Top, ndc, .Left, .Top, vbSrcCopy
End With
'delete context to free memory
DeleteDC ndc
FrmViewPic.Show
End Sub

Private Sub BtnCaptureScr_Click()
CapScr
End Sub

Private Sub BtnListAll_Click()
Label3 = "All HWnds!:"
BtnListChildren.BackColor = &H8000000F
BtnListAll.BackColor = &HC0FFC0 'green
BtnListParents.BackColor = &H8000000F
ListType = 3
RefreshList
End Sub

Private Sub BtnListChildren_Click()
Label3 = "All Children:"
BtnListChildren.BackColor = &HC0FFC0 'green
BtnListAll.BackColor = &H8000000F
BtnListParents.BackColor = &H8000000F
ListType = 2
LblParentInfo = "{Choose Child for Info}"
RefreshList
End Sub

Private Sub BtnListParents_Click()
Label3 = "Parents going up:"
BtnListChildren.BackColor = &H8000000F
BtnListAll.BackColor = &H8000000F
BtnListParents.BackColor = &HC0FFC0 'green
ListType = 1
RefreshList
End Sub

Private Sub BtnProgRefresh_Click()
Dim uProcess As PROCESSENTRY32
'get snapshot of programs
hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
'on error
If hSnapShot = 0 Then Exit Sub
'clean all the previous prog entries out of the listbox
LstProcess.Clear
'clean out the all previous prog records
Erase Process
'prepare for API call
uProcess.dwSize = Len(uProcess)
'get first program
r = Process32First(hSnapShot, uProcess)
'loop while progs are still being found
Do While r
 'copy prog data from temp to main process array
 Process(proc) = uProcess
 Select Case uProcess.pcPriClassBase
 Case 4: s$ = "Idle"
 Case 8: s$ = "Normal"
 Case 13: s$ = "High"
 Case 24: s$ = "Realtime"
 End Select
 'add name to Process listbox
 LstProcess.AddItem GetShortFileTitle(uProcess.szExeFile) & "     (" & s & " Priority)"
 'get next program
 r = Process32Next(hSnapShot, uProcess)
 proc = proc + 1
Loop
'close handle to free resource
CloseHandle hSnapShot
'reset info
SetProcessInfoMode 0
End Sub

Private Sub BtnSysInfo_Click()
Dim MemStat As MEMORYSTATUS, osinfo As OSVERSIONINFO, s$
Dim Serial&, VName$, FSName$, strsave$, tmp$, tcd As String * 30

'get memory statistics
GlobalMemoryStatus MemStat
'set length of type for API
osinfo.dwOSVersionInfoSize = Len(osinfo)
'get Windows version info
GetVersionEx osinfo
'pad strings for API
VName = Space(255)
FSName = Space(255)
'Get the volume information
GetVolumeInformation "C:\", VName, 255, Serial, 0, 0, FSName, 255
'remove nulls from volume info
VName = StripNulls(VName)
FSName = StripNulls(FSName)
'pad strings
tmp = Space(255)
'get all active drives
GetLogicalDriveStrings 255, tmp
'scan for all drives from "a:\" to "z:\"
For A = Asc("a") To Asc("z")
 'if tmp contains drive, add to strsave
 If InStr(tmp, Chr(A) & ":\") Then
  'add a comma if necessary and attach drive letter
  strsave = strsave & IIf(strsave > "", ",", "") & Chr(A)
 End If
Next

'prepare & show messagebox
s = "You have " & MemStat.dwTotalPhys \ 1024 ^ 2 & " Mb of physical memory."
s = s & vbCrLf & "You have " & MemStat.dwAvailPhys \ 1024 ^ 2 & " Mb of physical memory free."
s = s & vbCrLf & "You are using " & (MemStat.dwTotalPhys - MemStat.dwAvailPhys) \ 1024 ^ 2 & " Mb of physical memory."
s = s & vbCrLf & "You are using " & (MemStat.dwTotalVirtual - MemStat.dwAvailVirtual) \ 1024 ^ 2 & " Mb of virtual memory."
s = s & vbCrLf & vbCrLf & "Windows Version: " & osinfo.dwMajorVersion & "." & osinfo.dwMinorVersion
s = s & vbCrLf & "Windows Platform: " & IIf(osinfo.dwPlatformId = 1, "95/98/ME", "NT/2000")
s = s & vbCrLf & vbCrLf & "Drive C's File System: " & FSName
s = s & vbCrLf & "Drive C's Volume Name: " & VName
s = s & vbCrLf & "Drive C's Serial Number: " & Serial
s = s & vbCrLf & "Existing Drives: " & strsave
s = s & vbCrLf & vbCrLf & "Slow Computer: " & CBool(GetSystemMetrics(73))
s = s & vbCrLf & "Number of Mouse Buttons: " & GetSystemMetrics(43)
s = s & vbCrLf & "Computer is on a Network: " & CBool(GetSystemMetrics(63) And 1)
s = s & vbCrLf & "Running in Safe Mode: " & CBool(GetSystemMetrics(67))
MsgBox s
End Sub

Private Sub ChkShowExInfo_Click()
If ChkShowExInfo.Value Then
 FrameProgInfo.Visible = 1
 Height = ConstMaxHeight
 'this will move SysCon up if extended form would go over screen bottom
 If Top + Height > Screen.Height Then
  'center form vertically
  Top = Screen.Height \ 2 - Height \ 2
 End If
 BtnProgRefresh_Click 'list programs
Else
 FrameProgInfo.Visible = 0
 Height = ConstMinHeight
End If
End Sub

Private Sub DragThing_KeyPress(KeyAscii As Integer)
MsgBox "Sorry, you have to drag the button with the mouse."
End Sub

Private Sub DragThing_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'SetCapture hooks the form's mouse events
'into a global system, allowing the user to select
'a target (see Form's mouse functions)
SetCapture hwnd
'go into dragging mode
Dragging = 1
SetMode 0 'dragging
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'DragThing hooks the form's mouse events
'into a global system, allowing the user to select
'a target (see DragThing_MouseDown)

If Dragging Then
 GetCursorPos Mouse
 hw = WindowFromPoint(Mouse.X, Mouse.Y)
 'if it isn't the currently displayed HWnd
 If hw <> HWCustom Then
  'set up user's display
  SetTarget hw
  'pad string
  s$ = Space(256)
  'get the type of target
  r = GetClassName(HWCustom, s, 256)
  lblDragging = "Drag mouse to target!" & vbCrLf & "Class: " & s
  'listtype being 3 means 'List All' is on
  If ListType <> 3 Then RefreshList
 End If
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'return prog to normal operation on mouse release
SetMode 1
End Sub

Private Sub Form_Paint()
'if VB takes its ontopness, turn it back on
If GetStyleBitValue(-20, hwnd, &H8) = 0 And OptOnTop.Value Then
 SetWindowPos hwnd, -1, 0, 0, 0, 0, 3
End If
End Sub

Private Sub LstProcess_Click()
On Error Resume Next
'get the currently selected item
li = LstProcess.ListIndex
'if nothing was selected somehow
If li = -1 Then Exit Sub
'the number of semi-independent "mini-programs" in the main prog
LblProcess = "Number of Threads: " & Process(li).cntThreads & vbCrLf
'get various IDs for user's possible needs
LblProcess = LblProcess & "Process ID: " & Process(li).th32ProcessID & vbCrLf
LblProcess = LblProcess & "Parent ID: " & Process(li).th32ParentProcessID & vbCrLf
'get the name of the program
TxtLocation = Process(li).szExeFile
'enable the option buttons
SetProcessInfoMode 1
'get the process priority for the option buttons
OptPriority(GetPriority(Process(li).th32ProcessID)).Value = 1
End Sub

Private Sub BtnSendText_Click()
On Error GoTo errhandle

'-1 on none selected
If LstProcess.ListIndex = -1 Then Exit Sub
s$ = InputBox("Send Text")
If s = "" Then Exit Sub
'select the program target
AppActivate Process(LstProcess.ListIndex).th32ProcessID
'insert a slight delay for application switching
Sleep 100
'send the text that you entered to the program target
SendKeys s
Exit Sub

errhandle:
If Err.Number = 5 Then
 MsgBox "This program cannot be activated to send text." & vbCrLf & "It may not have a window to do so!"
Else
 MsgBox Err.Description
End If
End Sub

Private Sub MnuEditStyleYourSelf_Click()
'show the Edit Style Dialog modally
If HWCustom Then FrmCustStyleEdit.Show 1, Me
End Sub

Private Sub MnuStayOnTopOff_Click()
If IsWindow(HWCustom) Then
 SetWindowPos HWCustom, -2, 0, 0, 0, 0, 3
End If
End Sub

Private Sub MnuStayOnTopOn_Click()
If IsWindow(HWCustom) Then
 SetWindowPos HWCustom, -1, 0, 0, 0, 0, 3
End If
End Sub

Private Sub MnuTargetMax_Click()
'maximize target
If IsWindow(HWCustom) And HWCustom <> FrameSelItem.hwnd And HWCustom <> FrameHotkey.hwnd And FrameProgInfo.hwnd Then
 ShowWindowAsync HWCustom, 3
ElseIf IsWindow(HWCustom) Then
 MsgBox "Sorry, you can't maximize this frame." & vbCrLf & "This would impede normal functioning."
End If
End Sub

Private Sub MnuSendMessage_Click()
'show the Send Message Dialog modally
If HWCustom Then FrmSendMessage.Show 1, Me
End Sub

Private Sub MnuTargetMin_Click()
'minimize target
If IsWindow(HWCustom) Then ShowWindowAsync HWCustom, 6
End Sub

Private Sub MnuTargetNorm_Click()
'"restore" target
If IsWindow(HWCustom) Then ShowWindowAsync HWCustom, 9
End Sub

Private Sub OptPriority_Click(Index As Integer)
li = LstProcess.ListIndex
'on none selected
If li = -1 Then Exit Sub
'get priority for select listbox target
cp = GetPriority(Process(li).th32ProcessID)
'if it isn't set to that already
If Index <> cp Then
 'If you select realtime
 If Index = 256 Then
  If MsgBox("Are you sure you want set this program to realtime?" & vbCrLf & "If the program is not made for this (most are not), the computer could hang.", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
   'refresh priority info
   LstProcess_Click
   Exit Sub
  End If
 End If
 'set priority for selected item
 SetPriority Process(li).th32ProcessID, Index
End If
End Sub

Private Sub txtTarget_KeyPress(KeyAscii As Integer)
On Error GoTo errhandle

'when user presses enter, select new HWnd
If KeyAscii = 13 Then
 KeyAscii = 0
 SetTarget TxtTarget
End If
Exit Sub

errhandle:
HWCustom = 0
LblHWndP = "{error}"
LightSwitch 0 'off
End Sub

Private Sub BtnExtended_Click()
'popup extended menu
PopupMenu MnuMain
End Sub

Private Sub BtnHelp_Click()
FrmHelp.Show
End Sub

Private Sub BtnHighlightHWnd_Click()
On Error Resume Next 'just in case
'get display HWnd
ndc = CreateDC("DISPLAY", vbNullString, vbNullString, ByVal 0&)
'lock window updating
LockWindowUpdate GetDesktopWindow
For A% = 0 To ScreenDim.Right
 For b% = 0 To ScreenDim.Bottom
  'set color to each HWnd and attempt to make adjacent HWnd numbers more distinct
  SetPixelV ndc, A, b, ((WindowFromPoint(A, b) * 25) Or &HA5ABCB) Mod 2147483647
 Next
Next
'wait four seconds
Sleep 4000
'unlock updating
LockWindowUpdate 0
'delete resource to free memory
DeleteDC ndc
'make sure that the window is active for the 3 WindowStates
'by making it active (using Stay On Top)
SetWindowPos hwnd, -1, 0, 0, 0, 0, 3
'if 'Stay on Top' is not depressed, deactivate previous command
If OptOnTop.Value = 0 Then SetWindowPos hwnd, -2, 0, 0, 0, 0, 3
'refresh screen by maximizing, minimizing, and restoring SysCon
WindowState = 2: WindowState = 1: WindowState = 0
End Sub

Private Sub OptOnTop_Click()
If OptOnTop.Value Then
 '"stay on top" on
 SetWindowPos hwnd, -1, 0, 0, 0, 0, 3
Else
 '"stay on top" off
 SetWindowPos hwnd, -2, 0, 0, 0, 0, 3
End If
End Sub

Private Sub LstParents_Click()
'show extended info on selected in LstParents
s& = LstParents.List(LstParents.ListIndex)
'this signifies the main window
If s = 0 Then LblParentInfo = "Main Screen": Exit Sub
'check if listbox target still exists
If IsWindow(s&) = 0 Then LblParentInfo = "{Listbox Target Destroyed}": Exit Sub

'pad string with spaces for API
t$ = Space(256)
'get classname
GetClassName s, t, 256
t = "Class Name: " & vbCrLf & StripNulls(t) & vbCrLf & vbCrLf

'gets the target's standard style properties
l = GetWindowLong(s, -16)
t = t & "Visible: " & CBool(l And &H10000000) & vbCrLf
'call special functions to acquire the name of owning process
t = t & "Owner: " & GetShortFileTitle(GetProcessOwner(s)) & vbCrLf
'pad string with spaces for API
tmp$ = Space(70)
'get window's text
GetWindowText s, tmp, 70
'if window has text (versus simply being padded with nulls)
If InStr(1, tmp, Chr(0)) > 1 Then t = t & "Text: " & tmp
'move temporary string to label
LblParentInfo = t
End Sub

Private Sub LstParents_DblClick()
'select HWnd for controls (make it program target)
n$ = LstParents.List(LstParents.ListIndex)
'if item is 0, it signifies the screen
If n <> 0 Then
 If IsWindow(n) Then
  SetTarget n 'set target to selected listbox item
  If ListType < 3 Then RefreshList 'if listbox is not set to "HWnd Tree"
 End If
End If
End Sub

'free ANY window lock
Private Sub MnuUnlock_Click()
LockWindowUpdate 0
End Sub

'causes the titlebar to flash uCount (5) times
Private Sub MnuFlashWin_Click()
Dim FlashInfo As FLASHWINFO
If IsWindow(HWCustom) Then
 FlashInfo.dwFlags = 7
 FlashInfo.cbSize = Len(FlashInfo)
 FlashInfo.dwTimeout = 0 'default cursor blink rate
 FlashInfo.hwnd = HWCustom
 'specifies the number of times to flash the window
 FlashInfo.uCount = 5
 FlashWindowEx FlashInfo
End If
End Sub

Private Sub MnuLock_Click()
If IsWindow(HWCustom) = 0 Then Exit Sub
LockWindowUpdate 0 'unlock any locked
LockWindowUpdate HWCustom 'lock updating in target
End Sub

'register as screen saver (thereby disabling Alt+Tab, Ctrl+Alt+Del, etc.)
Private Sub OptRegSS_Click()
If OptRegSS.Value Then
 SystemParametersInfo 97, 1, "1", 0
Else
 SystemParametersInfo 97, 0, "1", 0
End If
End Sub

'shortcut for filling 'New Parent' textbox with SysCon's HWnd
Private Sub BtnMe_Click()
If TxtNewParent.Enabled Then TxtNewParent = hwnd
End Sub

Private Sub BtnLink_Click()
'steal target from another window, etc.
On Error Resume Next
If IsWindow(HWCustom) = 0 Then Exit Sub
SetParentAPI HWCustom, TxtNewParent
hC = TxtTarget
hP = LblHWndP
RefreshList
End Sub

Private Sub MnuDisable_Click()
If IsWindow(HWCustom) = 0 Then Exit Sub
'check if target is essential
If HWCustom <> hwnd And HWCustom <> FrameSelItem.hwnd And HWCustom <> BtnExtended.hwnd Then
 'run the subroutine and check for errors
 If SetStyleBitValue(-16, HWCustom, &H8000000, 1) = 0 Then
  MsgBox "Sorry, this target can't be disabled."
 End If
Else
 MsgBox "Sorry, you can't disable essential functions."
End If
End Sub

Private Sub MnuEnable_Click()
If IsWindow(HWCustom) = 0 Then Exit Sub
'run the subroutine and check for errors
If SetStyleBitValue(-16, HWCustom, &H8000000, 0) = 0 Then
 MsgBox "Sorry, this target can't be enabled."
End If
End Sub

Private Sub BtnExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
'registers Ctrl+Alt+A for hotkey (see Timer1)
ret = RegisterHotKey(hwnd, &HBFFF&, MOD_ALT Or MOD_CONTROL, vbKeyA)
'if ret is 0 hotkey registration malfunctioned
If ret = 0 Then MsgBox "Hotkey did not activate properly. Hotkey feature may not function."
'get screen dimensions for future uses
GetWindowRect GetDesktopWindow, ScreenDim
'change the display depending on the following constant
If ConstShowProgInfoByDefault Then
 Height = Height - 255
 ChkShowExInfo.Visible = 0
 FrameProgInfo.Visible = 1
 FrameProgInfo.Top = FrameProgInfo.Top - 255
 BtnProgRefresh_Click
Else
 Height = ConstMinHeight
End If
'set LstParents to list parents (as opposed to other settings)
ListType = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
'remove hotkey hook
UnregisterHotKey hwnd, &HBFFF&
'deactivate 'Register as SS' setting
SystemParametersInfo 97, 0, "1", 0
'unload all
Unload FrmHelp
Unload FrmViewPic
Unload FrmHWInfo
Unload FrmCustStyleEdit
Unload FrmSendMessage
End Sub

Private Sub MnuGetPos_Click()
Dim n As RECT, s$
If IsWindow(HWCustom) = 0 Then Exit Sub
'get target dimensions
GetWindowRect HWCustom, n
s = s & "X: " & n.Left & vbCrLf & "Y: " & n.Top
s = s & vbCrLf & "Width: " & n.Right - n.Left
s = s & vbCrLf & "Height: " & n.Bottom - n.Top
MsgBox s
End Sub

Private Sub BtnHide_Click()
If IsWindow(HWCustom) = 0 Then Exit Sub
'check if target is essential, otherwise hide
If HWCustom <> hwnd And HWCustom <> FrameSelItem.hwnd Then
 'hide target w/o waiting (prevents possible freezes)
 ShowWindowAsync HWCustom, vbHide
Else
 MsgBox "Sorry, you can't hide essential functions."
End If
End Sub

Private Sub BtnShow_Click()
'show target without focus or state change
If IsWindow(HWCustom) Then ShowWindowAsync HWCustom, 8
End Sub

Private Sub BtnInfo_Click()
'erase previous changes to infobox and shows it again
If IsWindow(HWCustom) Then Unload FrmHWInfo: FrmHWInfo.Show
End Sub

Private Sub BtnKill_Click()
If IsWindow(HWCustom) = 0 Then Exit Sub
'check if target is essential, otherwise kill
If HWCustom <> hwnd And HWCustom <> BtnExit.hwnd And HWCustom <> DragThing.hwnd And HWCustom <> FrameSelItem.hwnd Then
 If MsgBox("Are you sure?", vbYesNo, "Kill the Target") = vbYes Then DestroyWindow HWCustom
Else
 MsgBox "Sorry, you can't kill essential functions."
End If
End Sub

Private Sub BtnSelDesktop_Click()
'select desktop for target
hwc = FindWindow(vbNullString, "Program Manager")
SetTarget hwc
'update the listbox
RefreshList
End Sub

Private Sub BtnSelStart_Click()
'find the bar on which the Start button resides
tWnd = FindWindow("Shell_TrayWnd", vbNullString)
'select the Start button for target
hwc = FindWindowEx(tWnd, ByVal 0&, "BUTTON", vbNullString)
SetTarget hwc
'update the listbox
RefreshList
End Sub

Private Sub BtnSelTaskBar_Click()
'select the bar on which the Start button resides for target
hwc = FindWindow("Shell_TrayWnd", vbNullString)
SetTarget hwc
'update the listbox
RefreshList
End Sub

Private Sub TmrMain_Timer()
Dim Message As Msg
'checks for a hotkey message by using PeekMessage
'(hotkey is not 100% reliable, but it works on the first or second try)
For A = 1 To 10
 'wait for an incoming message
 WaitMessage
 'check for a hotkey message to this window
 If PeekMessage(Message, hwnd, WM_HOTKEY, WM_HOTKEY, PM_REMOVE) Then
  If OptCap.Value Then CapScr Else SetParentAPI hC, hP: RefreshList
 End If
Next

'check validity of current target
If IsWindow(HWCustom) = 0 And HWCustom Then
 HWCustom = 0
 'make light red
 LightSwitch 0 'off
 'update parent hwnd label
 LblHWndP = "was: " & vbCrLf & LblHWndP
End If
End Sub
