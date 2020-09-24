VERSION 5.00
Begin VB.Form FrmViewPic 
   AutoRedraw      =   -1  'True
   Caption         =   "ViewPic"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   Icon            =   "SysControl! CaptureBox.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu MnuPicture 
      Caption         =   "&Picture"
      Begin VB.Menu MnuSave 
         Caption         =   "&Save Picture"
      End
      Begin VB.Menu MnuClip 
         Caption         =   "C&opy 2 Clipboard"
      End
      Begin VB.Menu MnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu MnuClose 
         Caption         =   "&Close Window"
      End
   End
End
Attribute VB_Name = "FrmViewPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" (pOpenfilename As OpenFileName) As Long

Private Type OpenFileName
 lStructSize As Long
 hwndOwner As Long
 hInstance As Long
 lpstrFilter As String
 lpstrCustomFilter As String
 nMaxCustFilter As Long
 nFilterIndex As Long
 lpstrFile As String
 nMaxFile As Long
 lpstrFileTitle As String
 nMaxFileTitle As Long
 lpstrInitialDir As String
 lpstrTitle As String
 flags As Long
 nFileOffset As Integer
 nFileExtension As Integer
 lpstrDefExt As String
 lCustData As Long
 lpfnHook As Long
 lpTemplateName As String
End Type

Private Sub MnuClip_Click()
'set picture to clipboard
Clipboard.SetData FrmViewPic.Image
End Sub

Private Sub MnuClose_Click()
Unload Me
End Sub

Private Sub MnuSave_Click()
Dim Filename$, SaveFileDialog As OpenFileName
On Error Resume Next

10 With SaveFileDialog
 .lStructSize = Len(SaveFileDialog)
 .hwndOwner = HWnd
 .hInstance = App.hInstance
 .lpstrFilter = "Bitmap" + Chr(0) + "*.bmp"
 .lpstrFile = Space(254)
 .nMaxFile = 255
 .lpstrFileTitle = Space(254)
 .nMaxFileTitle = 255
 .lpstrInitialDir = CurDir
 .lpstrTitle = "Save Screenshot"
 .flags = 0
End With

'get filename
GetSaveFileName SaveFileDialog
'format filename
Filename = Trim(SaveFileDialog.lpstrFile)
Filename = Left(Filename, Len(Filename) - 1)
'if they entered nothing or canceled
If Filename = "" Then Exit Sub
'If there is an extension that does not fit, change it.
If LCase(Right(Filename, 4)) <> ".bmp" Then Filename = Filename & ".bmp"
'check for existing file by that name
If Dir(Filename) > "" Then
 If MsgBox("Are you sure you want to overwrite the existing file by that name?", vbYesNo, Filename) = vbNo Then GoTo 10
End If

'switch to error reporting mode for actual save
On Error GoTo errhandle
SavePicture Image, Filename
Exit Sub

errhandle:
MsgBox Err.Description
End Sub
