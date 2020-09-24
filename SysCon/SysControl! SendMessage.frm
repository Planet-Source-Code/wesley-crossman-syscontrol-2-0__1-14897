VERSION 5.00
Begin VB.Form FrmSendMessage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Send Message"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   Icon            =   "SysControl! SendMessage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnInfo 
      Caption         =   "?"
      Height          =   315
      Left            =   1320
      TabIndex        =   10
      ToolTipText     =   "Lists Useful Commands"
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton BtnApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   2940
      TabIndex        =   8
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton BtnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1500
      TabIndex        =   7
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   60
      TabIndex        =   6
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox TxtlParam 
      Height          =   285
      Left            =   3000
      MaxLength       =   11
      TabIndex        =   5
      Text            =   "0"
      Top             =   360
      Width           =   1155
   End
   Begin VB.TextBox TxtwParam 
      Height          =   285
      Left            =   1800
      MaxLength       =   11
      TabIndex        =   3
      Text            =   "0"
      Top             =   360
      Width           =   1155
   End
   Begin VB.TextBox TxtMsg 
      Height          =   285
      Left            =   60
      MaxLength       =   13
      TabIndex        =   1
      Text            =   "&H0"
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Note: Be careful! There are certain messages that are not good to send. It would be beneficial to have a book on the subject. "
      Height          =   615
      Left            =   60
      TabIndex        =   9
      Top             =   720
      UseMnemonic     =   0   'False
      Width           =   4155
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   1680
      X2              =   1680
      Y1              =   60
      Y2              =   720
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "lParam:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3000
      TabIndex        =   4
      Top             =   60
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "wParam:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1800
      TabIndex        =   2
      Top             =   60
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Message Number:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1335
   End
End
Attribute VB_Name = "FrmSendMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnApply_Click()
On Error GoTo errhandle
PostMessage HWCustom, TxtMsg, TxtwParam, TxtlParam
Exit Sub

errhandle:
MsgBox Err.Description
End Sub

Private Sub BtnCancel_Click()
Unload Me
End Sub

Private Sub BtnOK_Click()
BtnApply_Click
Unload Me
End Sub

Private Sub BtnInfo_Click()
Dim t$
t = "Useful Messages" & vbCrLf & vbCrLf
t = t & "WM_CLOSE, &h10, Close a target." & vbCrLf
t = t & "WM_COPY, &h301, Copy selected text to clipboard." & vbCrLf
t = t & "WM_CUT, &h300, Cut Selected text to clipboard." & vbCrLf
t = t & "WM_PASTE, &h302, Paste text from clipboard." & vbCrLf
t = t & "WM_UNDO, &h304, Undo the most recent text operation." & vbCrLf
t = t & "WM_LIMITTEXT, &hc5, Limit text to {wParam} letters." & vbCrLf
t = t & "WM_PAINT, &hf, Tell window to redraw." & vbCrLf
MsgBox t
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then BtnOK_Click
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
SetWindowPos HWnd, -1, 0, 0, 0, 0, 3 'stay on top
End Sub
