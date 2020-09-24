Attribute VB_Name = "ModSysCon"
Public Type RECT
 Left As Long
 Top As Long
 Right As Long
 Bottom As Long
End Type

Public Type POINTAPI
 X As Long
 Y As Long
End Type

Public Type Msg
 hwnd As Long
 Message As Long
 wParam As Long
 lParam As Long
 time As Long
 pt As POINTAPI
End Type

Public Type FLASHWINFO
 cbSize As Long
 hwnd As Long
 dwFlags As Long
 uCount As Long
 dwTimeout As Long
End Type

Public Type PROCESSENTRY32
 dwSize As Long
 cntUsage As Long
 th32ProcessID As Long
 th32DefaultHeapID As Long
 th32ModuleID As Long
 cntThreads As Long
 th32ParentProcessID As Long
 pcPriClassBase As Long
 dwFlags As Long
 szExeFile As String * 260
End Type

Public Type MEMORYSTATUS
 dwLength As Long
 dwMemoryLoad As Long
 dwTotalPhys As Long
 dwAvailPhys As Long
 dwTotalPageFile As Long
 dwAvailPageFile As Long
 dwTotalVirtual As Long
 dwAvailVirtual As Long
End Type

Public Type OSVERSIONINFO
 dwOSVersionInfoSize As Long
 dwMajorVersion As Long
 dwMinorVersion As Long
 dwBuildNumber As Long
 dwPlatformId As Long
 szCSDVersion As String * 128
End Type

Declare Function CreateToolhelp32Snapshot& Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long)
Declare Function Process32First& Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32)
Declare Function Process32Next& Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32)
Declare Function GetWindowThreadProcessId& Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long)
Declare Function GetFileTitle% Lib "comdlg32" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer)
Declare Function GetVersionEx& Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO)
Declare Function GetVolumeInformation& Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long)
Declare Function GetLogicalDriveStrings& Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String)
Declare Function GetSystemMetrics& Lib "user32" (ByVal nIndex As Long)
Declare Function GetPriorityClass& Lib "kernel32" (ByVal hProcess As Long)
Declare Function SetPriorityClass& Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long)
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

Declare Function FindWindow& Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String)
Declare Function ShowWindowAsync& Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long)
Declare Function SystemParametersInfo& Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long)
Declare Function FindWindowEx& Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String)
Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Declare Function SetParentAPI& Lib "user32" Alias "SetParent" (ByVal hWndChild As Long, ByVal hWndNewParent As Long)
Declare Function GetParentAPI& Lib "user32" Alias "GetParent" (ByVal hwnd As Long)
Declare Function GetDesktopWindow& Lib "user32" ()
Declare Function CreateDC& Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByVal lpInitData As Any)
Declare Function DeleteDC& Lib "gdi32" (ByVal hdc As Long)
Declare Function SetPixelV& Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long)
Declare Function DestroyWindow& Lib "user32" (ByVal hwnd As Long)
Declare Function GetWindowRect& Lib "user32" (ByVal hwnd As Long, lpRect As RECT)
Declare Function GetWindowText& Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long)
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long)
Declare Function RegisterHotKey& Lib "user32" (ByVal hwnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long)
Declare Function UnregisterHotKey& Lib "user32" (ByVal hwnd As Long, ByVal id As Long)
Declare Function PeekMessage& Lib "user32" Alias "PeekMessageA" (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long)
Declare Function WaitMessage& Lib "user32" ()
Declare Function SetParent& Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long)
Declare Function GetParent& Lib "user32" (ByVal hwnd As Long)
Declare Function WindowFromPoint& Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long)
Declare Function GetCursorPos& Lib "user32" (lpPoint As POINTAPI)
Declare Function LockWindowUpdate& Lib "user32" (ByVal hwndLock As Long)
Declare Function FlashWindowEx Lib "user32" (pfwi As FLASHWINFO) As Boolean
Declare Function IsWindow& Lib "user32" (ByVal hwnd As Long)
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function BitBlt& Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long)
Declare Function PostMessage& Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function IsWindowUnicode& Lib "user32" (ByVal hwnd As Long)
Declare Function SetWindowText& Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String)
Declare Function StretchBlt& Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long)
Declare Function GetStretchBltMode& Lib "gdi32" (ByVal hdc As Long)
Declare Function SetStretchBltMode& Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long)
Declare Function SetCapture& Lib "user32" (ByVal hwnd As Long)
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'used for storing the HWND for operations
Public HWCustom&
'screen capturing
Public WinDim As RECT, ScreenDim As RECT

'constants for API calls
Public Const RSP_SIMPLE_SERVICE = 1, RSP_UNREGISTER_SERVICE = 0
Public Const MOD_ALT = &H1, MOD_CONTROL = &H2, MOD_SHIFT = &H4
Public Const PM_REMOVE = &H1, WM_HOTKEY = &H312
Public Const TH32CS_SNAPPROCESS = 2

'takes a long path (drive:\folder\fname.ext) and converts it to it's
'short file name (fname.ext)
Function GetShortFileTitle$(ByVal FName$)
Dim buffer As String
'pad buffer with spaces for API
buffer = Space(255)
'use commondialog's API to get short name
n = GetFileTitle(FName, buffer, 255)
'save for return
GetShortFileTitle = StripNulls(buffer)
End Function

'get the name of the owning process
Function GetProcessOwner$(ByVal HWnd2Check&)
Dim TmpProc As PROCESSENTRY32

'get process id
GetWindowThreadProcessId HWnd2Check, processnum
'get snapshot of programs
hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
'on error
If hSnapShot = 0 Then Exit Function
TmpProc.dwSize = Len(TmpProc)
'get first program
r = Process32First(hSnapShot, TmpProc)
Do While r
 'check for the program owner of HWnd2Check
 If processnum = TmpProc.th32ProcessID Then
  GetProcessOwner = UCase(GetShortFileTitle(TmpProc.szExeFile))
 End If
 'get next program
 r = Process32Next(hSnapShot, TmpProc)
Loop
'free resource
CloseHandle hSnapShot
End Function

'Inputs: Mode (-16 for standard properties, -20 for extended)
'Inputs: TargetHWnd (specify HWnd for property change)
'Inputs: Mask (the number mask to apply to the HWnd property)
'Inputs: ValueNum (value to set masked property to)
'Return: 2 on success, 1 on already there, 0 on error
Function SetStyleBitValue(ByVal Mode%, ByVal TargetHWnd&, ByVal Mask&, ByVal ValueNum As Boolean) As Byte
'check for an illegal style
If Mode <> -16 And Mode <> -20 Then Exit Function
'get style in the mode specified
n = GetWindowLong(TargetHWnd, Mode)
'since the value is to be one or zero, check to make sure
'it isn't already (without this check, it will always
'reverse the value of the num!)
If ValueNum Then
 If (n And Mask) = 0 Then n = n Xor Mask Else SetStyleBitValue = 1: Exit Function
Else
 If (n And Mask) Then n = n Xor Mask Else SetStyleBitValue = 1: Exit Function
End If
'save the style with change
SetWindowLong TargetHWnd, Mode, n
'retrieve the style again
n = GetWindowLong(TargetHWnd, Mode)
'make sure the selected property changed
If ValueNum Then
 If (n And Mask) Then SetStyleBitValue = 2
Else
 If (n And Mask) = 0 Then SetStyleBitValue = 2
End If
End Function

'Inputs: Mode (-16 for standard properties, -20 for extended)
'Inputs: TargetHWnd (specify HWnd for property change)
'Inputs: Mask (the number mask to apply to the HWnd property)
'Return: Value of Style Bit
Function GetStyleBitValue(ByVal Mode%, ByVal TargetHWnd&, ByVal Mask&) As Byte
'check for an illegal style
If Mode <> -16 And Mode <> -20 Then Exit Function
'get style in the mode specified
n& = GetWindowLong(TargetHWnd, Mode)
GetStyleBitValue = IIf(n And Mask, 1, 0)
End Function


Function GetPriority%(ByVal ProcId&)
'creates link to process
lprocesshandle = OpenProcess(0, 0, ProcId)
'on error return error code & exit
If lprocesshandle = 0 Then GetPriority = 0: Exit Function
'get the priority
GetPriority = GetPriorityClass(lprocesshandle)
'disconnect link
CloseHandle lprocesshandle
End Function

Sub SetPriority(ByVal ProcId&, ByVal Prio%)
'create link to process
lprocesshandle = OpenProcess(0, 0, ProcId)
'on error exit
If lprocesshandle = 0 Then Exit Sub
'set priority
SetPriorityClass lprocesshandle, Prio
'disconnect link
CloseHandle lprocesshandle
End Sub

'search for first null and return everything before it
Function StripNulls$(ByVal TempStr$)
'get location of null
n% = InStr(TempStr, Chr(0))
'If one exists, remove all after it; otherwise return string unchanged.
If n > 0 Then
 StripNulls = Left(TempStr, InStr(TempStr, Chr(0)) - 1)
Else
 StripNulls = TempStr
End If
End Function
