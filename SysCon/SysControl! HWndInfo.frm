VERSION 5.00
Begin VB.Form FrmHWInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Info"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   Icon            =   "SysControl! HWndInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtP 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   300
      Width           =   2355
   End
   Begin VB.CommandButton BtnRefPic 
      Caption         =   "Refresh All"
      Height          =   315
      Left            =   4200
      TabIndex        =   5
      Top             =   2640
      Width           =   2355
   End
   Begin VB.PictureBox SPic 
      AutoRedraw      =   -1  'True
      Height          =   2295
      Left            =   4200
      ScaleHeight     =   149
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   153
      TabIndex        =   4
      Top             =   300
      Width           =   2355
      Begin VB.Label LblSorry 
         Alignment       =   2  'Center
         Caption         =   "Sorry! the target is hidden!"
         Height          =   255
         Left            =   180
         TabIndex        =   9
         Top             =   1020
         Width           =   1935
      End
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   60
      TabIndex        =   3
      Top             =   3000
      Width           =   6495
   End
   Begin VB.CommandButton BtnSet 
      Caption         =   "Change Text"
      Height          =   315
      Left            =   2460
      TabIndex        =   2
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox TxtContents 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   2460
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   300
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Properties:"
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
      TabIndex        =   8
      Top             =   60
      UseMnemonic     =   0   'False
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Picture of Target (if visible):"
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
      Left            =   4200
      TabIndex        =   6
      Top             =   60
      UseMnemonic     =   0   'False
      Width           =   1995
   End
   Begin VB.Label Label1 
      Caption         =   "Text:"
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
      Left            =   2460
      TabIndex        =   1
      Top             =   60
      UseMnemonic     =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "FrmHWInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnOK_Click()
Unload Me
End Sub

Private Sub BtnSet_Click()
'save text to target
SetWindowText HWCustom, TxtContents
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'if user presses ESC or Enter, exit
If KeyAscii = 27 Or KeyAscii = 13 Then Unload Me
End Sub

Private Sub Form_Load()
'set caption to read HWnd
Caption = "Info (HWnd " & HWCustom & ")"

'stay on top
SetWindowPos hwnd, -1, 0, 0, 0, 0, 3

'**** get properties ****

'check for target validity
If IsWindow(HWCustom) = 0 Then
 BtnSet.Enabled = 0
 TxtContents.Locked = 1
 Caption = "Info (target HWnd no longer valid)"
 Exit Sub
End If

'gets the class name
s$ = Space(256)
GetClassName HWCustom, s, 256
t$ = "Class Name: " & vbCrLf & StripNulls(s) & vbCrLf & vbCrLf

'process name collection
t = t & "Owner Process: " & vbCrLf & GetProcessOwner(HWCustom) & vbCrLf & vbCrLf

'gets the target's standard style properties
wl = GetWindowLong(HWCustom, -16)
'("wl And" code removes value from style and CBool converts it to 1 or 0)
t = t & "Border: " & CBool(wl And &H800000) & vbCrLf
t = t & "Caption: " & CBool(wl And &HC00000) & vbCrLf
t = t & "Child: " & CBool(wl And &H40000000) & vbCrLf
t = t & "Disabled: " & CBool(wl And &H8000000) & vbCrLf
t = t & "Tabstoppable: " & CBool(wl And &H10000) & vbCrLf
t = t & "Visible: " & CBool(wl And &H10000000) & vbCrLf
t = t & "Popup Window: " & CBool(wl And &H80000000) & vbCrLf
t = t & "Icon in Corner: " & CBool(wl And &H80000) & vbCrLf
t = t & "Unicode: " & CBool(IsWindowUnicode(HWCustom)) & vbCrLf

'gets the extended window style
l = GetWindowLong(HWCustom, -20)
t = t & "Accepts Files: " & CBool(l And &H10) & vbCrLf
t = t & "Stays on Top: " & CBool(l And &H8) & vbCrLf

'adds note
t = t & vbCrLf & "Note:" & vbCrLf & "Some of these properties may not apply to some targets. If the property doesn't seem right to you, it probably isn't applicable."

'set the temporary string into the Info box
TxtP = t

'**** get text ****

'pad the string with spaces for API
n$ = Space(2000)
'get text with a max length of 2000
GetWindowText HWCustom, n, 2000
'set textbox with string
TxtContents = n

'**** get picture of target ****

Dim windim2 As RECT
SPic.Cls
'get hdc from the display
ndc = CreateDC("DISPLAY", vbNullString, vbNullString, ByVal 0&)
'get rectangle of picture
GetWindowRect SPic.hwnd, windim2
'get rectangle of target
GetWindowRect HWCustom, WinDim
'get mode of resizing
gs = GetStretchBltMode(hdc)
'set the best quality mode
SetStretchBltMode hdc, 4
With WinDim
 'use BitBlt to capture the target and copy it to ViewPic
 'get width & height of picture
 pw = windim2.Right - windim2.Left
 ph = windim2.Bottom - windim2.Top
 'get width & height of target
 sw = .Right - .Left
 sh = .Bottom - .Top
 'copies the picture from the screen to the picturebox if visible
 StretchBlt SPic.hdc, 0, 0, pw, ph, ndc, .Left, .Top, sw + 1, sh, vbSrcCopy
 'check for visibility
 If (wl And &H10000000) Then
  LblSorry.Visible = 0
 Else
  LblSorry.Visible = 1
 End If
End With
'restore original mode
SetStretchBltMode hdc, gs
'delete context to free memory
DeleteDC ndc
End Sub

Private Sub BtnRefPic_Click()
'rerun starting code
Form_Load
End Sub
