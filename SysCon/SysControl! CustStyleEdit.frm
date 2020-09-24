VERSION 5.00
Begin VB.Form FrmCustStyleEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom Style Edit"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4245
   Icon            =   "SysControl! CustStyleEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnInfo 
      Caption         =   "?"
      Height          =   315
      Left            =   1440
      TabIndex        =   12
      Top             =   480
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Style to Change"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   60
      TabIndex        =   9
      Top             =   840
      Width           =   1635
      Begin VB.OptionButton OptExtended 
         Caption         =   "Extended Style"
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
         Left            =   60
         TabIndex        =   11
         Top             =   420
         Width           =   1275
      End
      Begin VB.OptionButton OptStandard 
         Caption         =   "Standard Style"
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
         Left            =   60
         TabIndex        =   10
         Top             =   180
         Value           =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.CommandButton BtnGo 
      Caption         =   "Get Current Value"
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
      Left            =   840
      TabIndex        =   8
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton BtnApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   2280
      Width           =   1335
   End
   Begin VB.OptionButton OptOff 
      Caption         =   "Off"
      Enabled         =   0   'False
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   1920
      Width           =   675
   End
   Begin VB.OptionButton OptOn 
      Caption         =   "On"
      Enabled         =   0   'False
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   1680
      Width           =   675
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   60
      MaxLength       =   13
      TabIndex        =   3
      Text            =   "&h0"
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton BtnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1500
      TabIndex        =   1
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   2280
      Width           =   1395
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   $"SysControl! CustStyleEdit.frx":000C
      Height          =   1995
      Left            =   1860
      TabIndex        =   6
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a Mask"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   1635
   End
End
Attribute VB_Name = "FrmCustStyleEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefLng A-Z
Dim Mode 'for Extended or Standard style mode

Private Sub BtnApply_Click()
On Error GoTo errhandle

If OptOn.Enabled Then
 If SetStyleBitValue(Mode, HWCustom, Text1, OptOn.Value) = 0 Then
  MsgBox "Sorry, that style bit apparently can't be changed."
 End If
End If
Exit Sub

errhandle:
MsgBox Err.Description
End Sub

Private Sub BtnCancel_Click()
Unload Me
End Sub

Private Sub BtnInfo_Click()
Dim t$
t = "Style Editing Info:" & vbCrLf & vbCrLf
t = t & "First, enter a style." & vbCrLf
t = t & "Next, select the style ""library"" to use." & vbCrLf
t = t & "Finally, get the current value & pick the value you want." & vbCrLf & vbCrLf
t = t & "Useful Properties:" & vbCrLf & vbCrLf
t = t & "WS_DISABLED, &h8000000, Standard" & vbCrLf
t = t & "Will disable the target on ""on""." & vbCrLf & vbCrLf
t = t & "WS_EX_TRANSPARENT, &h20, Extended" & vbCrLf
t = t & "Won't make the target transparent but will cause crazy effects." & vbCrLf & vbCrLf
t = t & "..." & vbCrLf
t = t & "many more (however, most are too task specific to be here)"
MsgBox t
End Sub

Private Sub BtnOK_Click()
BtnApply_Click
Unload Me
End Sub

Private Sub BtnGo_Click()
On Error GoTo errhandle
If GetStyleBitValue(Mode, HWCustom, Text1) Then
 OptOn.Value = 1
 OptOff.Value = 0
Else
 OptOn.Value = 0
 OptOff.Value = 1
End If
OptOn.Enabled = 1
OptOff.Enabled = 1
Exit Sub

errhandle:
MsgBox Err.Description
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then BtnOK_Click
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
SetWindowPos hwnd, -1, 0, 0, 0, 0, 3 'stay on top
Mode = -16 'start with standard mode
End Sub

Private Sub OptExtended_Click()
Mode = -20 'extended
End Sub

Private Sub OptStandard_Click()
Mode = -16 'standard
End Sub
