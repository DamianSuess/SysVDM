Attribute VB_Name = "fuct_wnd05"
Option Explicit
'[ http://www.vbstatic.net/
'[ ---------------------------
'[ fuct22.bas
'[ fuct33.bas
'[ fuctWnd1-20.bas
'[ fuct_wnd04   feb-4th 2004
'[ ---------------------------

Public waitEND As Boolean     '[ used with waitforclass(...)

Public Type POINTAPI
  x As Long
  Y As Long
End Type
Public Type RECT
  left As Long
  top As Long
  right As Long
  bottom As Long
End Type

'[ Mouse Pointer
Public Declare Function CreateCaret Lib "user32" (ByVal hWnd As Long, ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function GetCaretBlinkTime Lib "user32" () As Long
Public Declare Function SetCaretBlinkTime Lib "user32" (ByVal wMSeconds As Long) As Long
Public Declare Function DestroyCaret Lib "user32" () As Long
Public Declare Function HideCaret Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ShowCaret Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetCaretPos Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long
Public Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Declare Function BeginDeferWindowPos Lib "user32" (ByVal nNumWindows As Long) As Long
Public Declare Function DeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long, ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function EndDeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'[ Windows on Screen
Public Declare Function MapWindowPoints Lib "user32" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function ChildWindowFromPoint Lib "user32" (ByVal hWnd As Long, ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

'[ General API
Public Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsZoomed Lib "user32" (ByVal hWnd As Long) As Boolean
Public Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long   '[ used w/  GetFocus(), SetFocus()
Public Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetModuleFileName Lib "Kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hWnd As Long, ByVal wFlag As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetPath Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, lpTypes As Byte, ByVal nSize As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetTopWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long

'Public Declare Function MoveWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetFocusEX Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function PostMessageByString Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

'[ Subclass
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowWord Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Integer

'[ drop shadow
Public Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const CS_DROPSHADOW = &H20000
Public Const GCL_STYLE = (-26)

'[ Text-Font Width
Public Type SIZE
  cx As Long
  cy As Long
End Type
Public Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZE) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

'[ Get System Version
Private Type TOSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type
Public Declare Function GetVersionEx Lib "Kernel32" Alias "GetVersionExA" (lpVersionInformation As TOSVERSIONINFO) As Long
Public Const VER_PLATFORM_WIN32s = 0         'Windows 3.x is running, using Win32s
Public Const VER_PLATFORM_WIN32_WINDOWS = 1  'Windows 95 or 98 is running.
Public Const VER_PLATFORM_WIN32_NT = 2       'Windows NT is running.

'[ wall paper shit
Public Declare Function SysParamNFO_SCREEN Lib "user32.dll" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uiParam As Long, pvParam As Any, ByVal fWinIni As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long

Public Const SPI_SETDESKWALLPAPER = 20
Public Const SPIF_SENDWININICHANGE = &H2
Public Const SPIF_UPDATEINIFILE = &H1

'[ Window pos
Public Const HWND_BOTTOM = 1
Public Const HWND_BROADCAST = &HFFFF&
Public Const HWND_DESKTOP = 0
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1

Public Enum ZPOS_CLASS
  ZPOS_Normal = 0
  ZPOS_TopMost = 1
  ZPOS_Bottom = 2
  ZPOS_Desktop = 3
End Enum


Public Const SWP_NOMOVE = &H2           ' Do not reposition window
Public Const SWP_NOSIZE = &H1           ' Do not re-size window
Public Const SWP_SHOWWINDOW = &H40      ' Make window visible/active
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

' ShowWindow() Commands
Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_NORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10
Public Const SW_MAX = 10

'[ GetWindow() Constants
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_OWNER = 4
Public Const GW_CHILD = 5
Public Const GW_MAX = 5

' Window field offsets for GetWindowLong() and GetWindowWord()
Public Const GWL_WNDPROC = (-4)
Public Const GWL_HINSTANCE = (-6)
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_USERDATA = (-21)
Public Const GWL_ID = (-12)

'[ Window Style [tasks]
Public Const WS_EX_APPWINDOW = &H40000
Public Const WS_EX_TOOLWINDOW = &H80
Public Const WS_EX_TOPMOST = &H8&

'[ window style
Public Const WS_BORDER = &H800000         ' Window has a border
Public Const WS_CAPTION = &HC00000
Public Const WS_CHILD = &H40000000
Public Const WS_CLIPSIBLINGS = &H4000000  ' can clip windows
Public Const WS_DISABLED = &H8000000
Public Const WS_GROUP = &H20000           ' Window is top of group
Public Const WS_MINIMIZE = &H20000000     ' Style bit 'is minimized'
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_THICKFRAME = &H40000      ' Window has thick border
Public Const WS_TABSTOP = &H10000         ' Window has tabstop
Public Const WS_VISIBLE = &H10000000      ' Window is not hidden

' Class field offsets for GetClassLong() and GetClassWord()
Public Const GCL_MENUNAME = (-8)
Public Const GCL_HBRBACKGROUND = (-10)
Public Const GCL_HCURSOR = (-12)
Public Const GCL_HICON = (-14)
Public Const GCL_HMODULE = (-16)
Public Const GCL_CBWNDEXTRA = (-18)
Public Const GCL_CBCLSEXTRA = (-20)
Public Const GCL_WNDPROC = (-24)
'Public Const GCL_STYLE = (-26)
Public Const GCW_ATOM = (-32)

'[ Window Messages
Public Const WM_CLOSE = &H10
Public Const WM_CHAR = &H102
Public Const WM_COMMAND = &H111
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_ENABLE = &HA
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYFIRST = &H100
Public Const WM_KEYLAST = &H108
Public Const WM_KEYUP = &H101
Public Const WM_SETTEXT = &HC
Public Const WM_SETFOCUS = &H7
Public Const WM_SETCURSOR = &H20
Public Const WM_SHOWWINDOW = &H18
Public Const WM_USER = &H400
Public Const WM_WINDOWPOSCHANGED = &H47
Public Const WM_WINDOWPOSCHANGING = &H46
Public Const WM_MOVE = &H3
'----------------------------
Public Const WM_MOUSEFIRST = &H200
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MOUSELAST = &H209

Public Const VK_SPACE = &H20


'[ Listbox messages
Public Const LB_ADDSTRING = &H180
Public Const LB_INSERTSTRING = &H181
Public Const LB_DELETESTRING = &H182
Public Const LB_SELITEMRANGEEX = &H183
Public Const LB_RESETCONTENT = &H184
Public Const LB_SETSEL = &H185
Public Const LB_SETCURSEL = &H186
Public Const LB_GETSEL = &H187
Public Const LB_GETCURSEL = &H188
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_GETCOUNT = &H18B
Public Const LB_SELECTSTRING = &H18C
Public Const LB_DIR = &H18D
Public Const LB_GETTOPINDEX = &H18E
Public Const LB_FINDSTRING = &H18F
Public Const LB_GETSELCOUNT = &H190
Public Const LB_GETSELITEMS = &H191
Public Const LB_SETTABSTOPS = &H192
Public Const LB_GETHORIZONTALEXTENT = &H193
Public Const LB_SETHORIZONTALEXTENT = &H194
Public Const LB_SETCOLUMNWIDTH = &H195
Public Const LB_ADDFILE = &H196
Public Const LB_SETTOPINDEX = &H197
Public Const LB_GETITEMRECT = &H198
Public Const LB_GETITEMDATA = &H199
Public Const LB_SETITEMDATA = &H19A
Public Const LB_SELITEMRANGE = &H19B
Public Const LB_SETANCHORINDEX = &H19C
Public Const LB_GETANCHORINDEX = &H19D
Public Const LB_SETCARETINDEX = &H19E
Public Const LB_GETCARETINDEX = &H19F
Public Const LB_SETITEMHEIGHT = &H1A0
Public Const LB_GETITEMHEIGHT = &H1A1
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_SETLOCALE = &H1A5
Public Const LB_GETLOCALE = &H1A6
Public Const LB_SETCOUNT = &H1A7
Public Const LB_MSGMAX = &H1A8

'[ Button Control Messages
Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1
Public Const BM_GETSTATE = &HF2
Public Const BM_SETSTATE = &HF3
Public Const BM_SETSTYLE = &HF4

'[ Combo Box Notification Codes
Public Const CBN_ERRSPACE = (-1)
Public Const CBN_SELCHANGE = 1
Public Const CBN_DBLCLK = 2
Public Const CBN_SETFOCUS = 3
Public Const CBN_KILLFOCUS = 4
Public Const CBN_EDITCHANGE = 5
Public Const CBN_EDITUPDATE = 6
Public Const CBN_DROPDOWN = 7
Public Const CBN_CLOSEUP = 8
Public Const CBN_SELENDOK = 9
Public Const CBN_SELENDCANCEL = 10

'[ Combo Box messages
Public Const CB_GETEDITSEL = &H140
Public Const CB_LIMITTEXT = &H141
Public Const CB_SETEDITSEL = &H142
Public Const CB_ADDSTRING = &H143
Public Const CB_DELETESTRING = &H144
Public Const CB_DIR = &H145
Public Const CB_GETCOUNT = &H146
Public Const CB_GETCURSEL = &H147
Public Const CB_GETLBTEXT = &H148
Public Const CB_GETLBTEXTLEN = &H149
Public Const CB_INSERTSTRING = &H14A
Public Const CB_RESETCONTENT = &H14B
Public Const CB_FINDSTRING = &H14C
Public Const CB_SELECTSTRING = &H14D
Public Const CB_SETCURSEL = &H14E
Public Const CB_SHOWDROPDOWN = &H14F
Public Const CB_GETITEMDATA = &H150
Public Const CB_SETITEMDATA = &H151
Public Const CB_GETDROPPEDCONTROLRECT = &H152
Public Const CB_SETITEMHEIGHT = &H153
Public Const CB_GETITEMHEIGHT = &H154
Public Const CB_SETEXTENDEDUI = &H155
Public Const CB_GETEXTENDEDUI = &H156
Public Const CB_GETDROPPEDSTATE = &H157
Public Const CB_FINDSTRINGEXACT = &H158
Public Const CB_SETLOCALE = &H159
Public Const CB_GETLOCALE = &H15A
Public Const CB_MSGMAX = &H15B

'[ Edit Control Messages
Public Const EM_GETSEL = &HB0
Public Const EM_SETSEL = &HB1
Public Const EM_GETRECT = &HB2
Public Const EM_SETRECT = &HB3
Public Const EM_SETRECTNP = &HB4
Public Const EM_SCROLL = &HB5
Public Const EM_LINESCROLL = &HB6
Public Const EM_SCROLLCARET = &HB7
Public Const EM_GETMODIFY = &HB8
Public Const EM_SETMODIFY = &HB9
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINEINDEX = &HBB
Public Const EM_SETHANDLE = &HBC
Public Const EM_GETHANDLE = &HBD
Public Const EM_GETTHUMB = &HBE
Public Const EM_LINELENGTH = &HC1
Public Const EM_REPLACESEL = &HC2
Public Const EM_GETLINE = &HC4
Public Const EM_LIMITTEXT = &HC5
Public Const EM_CANUNDO = &HC6
Public Const EM_UNDO = &HC7
Public Const EM_FMTLINES = &HC8
Public Const EM_LINEFROMCHAR = &HC9
Public Const EM_SETTABSTOPS = &HCB
Public Const EM_SETPASSWORDCHAR = &HCC
Public Const EM_EMPTYUNDOBUFFER = &HCD
Public Const EM_GETFIRSTVISIBLELINE = &HCE
Public Const EM_SETREADONLY = &HCF
Public Const EM_SETWORDBREAKPROC = &HD0
Public Const EM_GETWORDBREAKPROC = &HD1
Public Const EM_GETPASSWORDCHAR = &HD2

'[ Edit Control Notification Codes
Public Const EN_SETFOCUS = &H100
Public Const EN_KILLFOCUS = &H200
Public Const EN_CHANGE = &H300
Public Const EN_UPDATE = &H400
Public Const EN_ERRSPACE = &H500
Public Const EN_MAXTEXT = &H501
Public Const EN_HSCROLL = &H601
Public Const EN_VSCROLL = &H602

Public Function SplitFix(ByVal Expression As String, _
  Optional ByVal Delimiter As String = " ", _
  Optional ByVal Limit As Long = -1, _
  Optional Compare As VbCompareMethod = vbBinaryCompare) As Variant
  
  '[ x = SplitFix("Hello my  name is   Fuct") ' ---> 'hello my name is fuct'
  '[ For i = LBound(x) To UBound(x): MsgBox x(i): Next i
  Dim varItems As Variant, i&
  varItems = Split(Expression, Delimiter, Limit, Compare)
  For i = LBound(varItems) To UBound(varItems)
    If Len(varItems(i)) = 0 Then varItems(i) = Delimiter
  Next i
  SplitFix = filter(varItems, Delimiter, False)
End Function
Public Function ArrSize(arr()) As Long '[ takes in Variant[] array
  On Error GoTo err_hand
  
  Dim cnt&
  cnt = UBound(arr())
  ArrSize = cnt
  
Exit Function
err_hand:
  
  ArrSize = -1
  
End Function
Public Sub sizeListbox(lst As ListBox)
 Dim ndx&, iLen%, range&
 
 If lst.ListCount = 0 Then Exit Sub
 
 For ndx = 0 To lst.ListCount - 1
  iLen = GetTextSize(lst.hWnd, lst.List(ndx))
  If iLen > range Then range = iLen
 Next ndx
 
 Call SendMessage(lst.hWnd, &H415, range * 15, 0&)
  ' Call AddScrollList(lst, range)
  '[ Public Sub AddScrollList(lst As ListBox, range&)
  'Debug.Print "lstLen: " & range
End Sub

Public Function GetTextSize(hWnd&, Txt$) As Long
 Dim iDC&, lpSize As SIZE, wnd&, iWdth&
 '[ wnd = box.hWnd
 iDC = GetDC(hWnd)

 Call GetTextExtentPoint32(iDC, Txt, Len(Txt), lpSize)
 iWdth = lpSize.cx

 Call ReleaseDC(hWnd, iDC)

 GetTextSize = iWdth '* 15
End Function
Public Function ArrSize_s(arr() As String) As Long '[ takes in String[] array
  On Error GoTo err_hand
  
  Dim cnt&
  cnt = UBound(arr())
  ArrSize_s = cnt
  
Exit Function
err_hand:
  
  ArrSize_s = -1
  
End Function


Public Function comp(word1$, word2$) As Boolean
  '[ my most favorate function ever written
  '[ with FixString(str, chr32rmv) & FindChildByClass(wnd, class, ndx)
  '[ just behind
  
  If LCase(word1) = LCase(word2) Then
    comp = True
  Else
    comp = False
  End If
End Function

Public Function DirExists(ByVal path$) As Boolean
  Dim ret$  '[ never know when you need this,
            '[ if you find an err0r tell me~!~!    - fuct
  path = Trim(path)
  If Len(path) = 0 Then DirExists = False: Exit Function
  If right(path, 1) <> "\" Then path = path & "\"
  
  ret = Dir(path, vbDirectory)
  If (ret = ".") Then
    DirExists = True
  Else
    DirExists = False
  End If
End Function


Public Sub MoveForm2(ByVal wnd As Long)
  ReleaseCapture
  SendMessage wnd, &HA1, 2, 0&
End Sub

Public Function CurDirFull(Optional full As Boolean = True)
  '[ returns full path (OR) w. "\"
  Dim buff$, ret&, pos&
  
  buff = String(512, 0)
  ret = GetModuleFileName(0, buff, 512)
  buff = fixstring(buff)
  
  If full = True Then
    CurDirFull = buff
  Else
    pos = InStrRev(buff, "\")
    CurDirFull = Mid(buff, 1, pos)
  End If
End Function
Public Function CurrentDir(Optional slash As Boolean = False) As String '[ fuct 1998
  '[ 'returns path w/ or w/o "\"
  Dim cDir$, pth$
  pth = App.path
  cDir = pth
  If Mid(pth, Len(pth), 1) = Chr$(89 + 3) Then
    If slash = False Then _
      cDir = Mid(pth, 1, Len(pth) - 1)
  Else
    If slash = True Then _
      cDir = pth & "\"
  End If
  CurrentDir = cDir
End Function


Public Function IsWin2000Plus() As Boolean
  Dim OSV As TOSVERSIONINFO
  IsWin2000Plus = False
  OSV.dwOSVersionInfoSize = Len(OSV)
  If GetVersionEx(OSV) = 1 Then
     IsWin2000Plus = (OSV.dwPlatformId = VER_PLATFORM_WIN32_NT) And (OSV.dwMajorVersion >= 5)
  End If
  
   'Dim os   As OSVERSIONINFO
   'Dim lRet As Long
   'IsWin2000 = False   ' Check for windows 2000
   'os.dwOSVersionInfoSize = Len(os)  ' set the size of the structure
   'lRet = GetVersionEx(os)           ' read Windows's version information
   'If os.dwPlatformId = VER_PLATFORM_WIN32_NT Then' Check for Win32 & NT/2000
   '   If (os.dwMajorVersion < 5) Then
   '      IsWin2000 = False
   '   Else: IsWin2000 = True
   '   End If
   'End If
End Function
Public Sub do3d(wnd&)
  Dim ret&
  ret = GetClassLong(wnd, GCL_STYLE)
  Call SetClassLong(wnd, GCL_STYLE, (ret Or CS_DROPSHADOW))
  
End Sub

Public Sub Click(wnd&)
  Call SendMessage(wnd, WM_LBUTTONDOWN, 0, 0): DoEvents
  Call SendMessage(wnd, WM_LBUTTONUP, 0, 0): DoEvents
End Sub

Public Sub Click2(mButton As Long)
  Call SendMessage(mButton&, WM_KEYDOWN, VK_SPACE, 0&)
  Call SendMessage(mButton&, WM_KEYUP, VK_SPACE, 0&)
End Sub
Public Sub Click3(wnd&) '[ pure fuct.up shit
  Call SendMessageByLong(GetParent(wnd), WM_COMMAND, GetWindowWord(wnd, -12), ByVal CLng(wnd))
End Sub
Public Sub ClickChk(wnd&)
  '[ good for message boxes
  Const MK_LBUTTON = &H1
  Call SendMessage(wnd, WM_LBUTTONDOWN, MK_LBUTTON, 0): DoEvents
  Call SendMessage(wnd, BM_SETSTATE, 1, 0)
  
  Call SendMessage(wnd, WM_LBUTTONUP, 0, 0): DoEvents
  Call SendMessage(wnd, BM_SETSTATE, 0, 0)
End Sub


Public Function FileExists(pth As String) As Boolean '[ fuct22.bas
  On Error GoTo FileExistError
  Dim attrib&
  attrib = vbNormal Or vbReadOnly Or vbHidden Or vbSystem
  '[ vbNormal = 0
  '[ vbReadOnly = 1
  '[ vbHidden = 2
  '[ vbSystem = 4
  If Len(Trim(pth)) = 0 Then
    FileExists = False
    Exit Function
  ElseIf Len(Dir$(pth, attrib)) Then
    FileExists = True
  Else
    FileExists = False
  End If
Exit Function
FileExistError:
  Select Case Err
   Case 76: FileExists = False   'Path not found ie: CD-ROM disk not present
   Case 75: FileExists = False   'Path/File access error
   Case 68: FileExists = False   'Device unavailable
  End Select
End Function


Public Function FindChild(par&, Class$, Title$, Optional logic As Boolean = True) As Long
  Dim wnd&, cl$, tx$
  If par = 0 Then Exit Function
  wnd = GetWindow(GetWindow(par, GW_CHILD), GW_HWNDFIRST)
  Do: DoEvents
    cl = GetClass(wnd)
    tx = fixstring(GrabWndText(wnd))
    'If (LCase(cl) = LCase(Class)) And (LCase(tx) = LCase(Title)) Then Exit Do
    If (LCase(cl) = LCase(Class)) Then
      If logic = True Then '[ default value
        If (LCase(tx) = LCase(Title)) Then Exit Do
      Else: If InStr(1, LCase(tx), LCase(Title), vbTextCompare) Then Exit Do
      End If
    End If
    wnd = GetWindow(wnd, GW_HWNDNEXT)
  Loop Until wnd = 0
  FindChild = wnd
End Function
Public Function FindChildByTitle(par&, Title$, Optional logic As Boolean = True) As Long
  Dim wnd&, Txt$ '[ set logic=TRUE if you know EXACT title
  If par = 0 Then Exit Function '[ def=TRUE
  wnd = GetWindow(GetWindow(par, GW_CHILD), GW_HWNDFIRST)
  Do: DoEvents
    Txt = fixstring(GrabWndText(wnd))
    If logic = True Then '[ default value
      If LCase(Txt) = LCase(Title) Then Exit Do
    Else
      If InStr(1, LCase(Txt), LCase(Title), vbTextCompare) Then Exit Do
    End If
    wnd = GetWindow(wnd, GW_HWNDNEXT)
  Loop Until wnd = 0
  FindChildByTitle = wnd
End Function
Public Function FindChildByClass(par&, cls_nme$, Optional indx& = 1) As Long
  Dim wnd&, Txt$, ndx&
  ndx = 0
  If par = 0 Then Exit Function
  If indx <= 0 Then indx = 1
  
  wnd = GetWindow(GetWindow(par, GW_CHILD), GW_HWNDFIRST)
  Do: DoEvents
    Txt = GetClass(wnd)
    If LCase(Txt) = LCase(cls_nme) Then
      ndx = ndx + 1
      If ndx >= indx Then Exit Do
    End If
    wnd = GetWindow(wnd, GW_HWNDNEXT)
  Loop Until wnd = 0
  FindChildByClass = wnd
End Function
Public Function GetClass(hWnd As Long) '[ fuct98.bas
  Dim Txt$, x&
  Txt = String(500, 0) '[ dont need a big cls.name
  x = GetClassName(hWnd, Txt, 500)
  GetClass = fixstring(Txt)
End Function
Public Function fixstring(ByVal Txt As String, Optional rmvSpace As Boolean = False) As String
  If InStr(1, Txt, Chr(0), 1) Then
    Txt = Mid(Txt, 1, InStr(1, Txt, Chr(0), 1) - 1)
    If rmvSpace = True Then Txt = Trim(Txt)
  End If
  fixstring = Txt
End Function

Public Sub ontop(wnd As Long, pos As ZPOS_CLASS)
  '[ fuct_wnd05
  Dim zPos&, wndProgMan&, wndPar&
  
  zPos = HWND_NOTOPMOST
  If pos = ZPOS_Normal Then
    zPos = HWND_NOTOPMOST
    Call SetParent(wnd, 0)
    
  ElseIf pos = ZPOS_TopMost Then
    zPos = HWND_TOPMOST
    Call SetParent(wnd, 0)
    
  ElseIf pos = ZPOS_Bottom Then
    zPos = HWND_BOTTOM
    Call SetParent(wnd, 0)
    
  ElseIf pos = ZPOS_Desktop Then
    '[ HWND wndPar = GetAncestor(wnd, GA_PARENT)
    wndProgMan = FindWindow("Progman", "Program Manager")
    '[ If (wndProgMan And (wndPar <> wndProgMan)) Then
    Call SetParent(wnd, wndProgMan)
    '[ End If
  End If
  'if (zPos != ZPOSITION_ONDESKTOP && (wndPar != GetDesktopWindow()))
  '    SetParent(wnd, NULL);

  
  Call SetWindowPos(wnd, zPos, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE)
  
'[ onBottom
'DeskH = GethWndByWinTitle("Program Manager")
'Call SetParent(frm.hWnd, DeskH)

  '[ OLD---[ fuct22.bas ]---------------------
  '[ Public Sub ontop(wnd As Long, pos As Boolean)
  '[ If pos = True Then
  '[   Call SetWindowPos(wnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, Flags)
  '[ Else: Call SetWindowPos(wnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, Flags)
  '[ End If
End Sub


Public Function IntToHex(numbr&, Optional Length% = 4, Optional headr As Boolean = True) As String
  '[ Make number into "0x40D0" format
  
  Dim ret As String, sNum As String, i%
  sNum = Hex(numbr)
  i = Abs(Length - Len(sNum)) '[ ABS(..) for error trapping only
  ret = String(i, "0") & sNum
  If headr = True Then ret = "0x" & ret
  IntToHex = ret
End Function

Public Function Hex2Int(Data As String) As String
  Dim cnt As Double
  For cnt = 1 To Len(Data) Step 2
    Hex2Int = Hex2Int & CInt(Val("&H" & Mid$(Data, cnt, 2)))
  Next cnt
End Function
Public Function RShift(ByRef iValue As Long, ByRef iShift As Long)
  RShift = CLng(iValue \ (2 ^ iShift)) '[   same as >> operator
End Function


Public Function LShift(ByRef iValue As Long, ByRef iShift As Long)
  LShift = iValue * (2 ^ iShift) '[   same as << operator
End Function

Public Sub StrToUni(sTxt$, ByteArray() As Byte)
  '[ fuct0x1
  ByteArray = StrConv(sTxt, vbFromUnicode)
End Sub

Public Function UniToStr(ByteArray() As Byte) As String
  '[ fuct0x1
  UniToStr = StrConv(ByteArray, vbUnicode)
End Function

Public Function GrabWndText(hWnd As Long) As String
  Dim x&, ln&, Txt$
  ln = SendMessage(hWnd, WM_GETTEXTLENGTH, 0, 0)
  Txt = Space(ln)
  x = SendMessageByString(hWnd, WM_GETTEXT, ln + 1, Txt)
  GrabWndText = Txt
End Function
Public Function MsgBoxText(Optional sTitle$ = vbNullString) As String
  Dim i&, ii&, wnd&, t$
  wnd = FindWindow("#32770", sTitle)
  If wnd <> 0 Then
    i = FindChildByClass(wnd, "static", 2)
    If ii = 0 Then
      t = GrabWndText(i)
    Else: t = GrabWndText(ii)
    End If
  Else
    t = "<no msgbox>"
  End If
  MsgBoxText = t
End Function

Public Function MakeLong(LoWord As Integer, HiWord As Integer) As Long
  '[ creates a Long value using Low and High integers
  '[ useful when converting code from C++
  Dim nLoWord As Long
  If LoWord% < 0 Then
    nLoWord& = LoWord% + &H10000
  Else
    nLoWord& = LoWord%
  End If
  MakeLong& = CLng(nLoWord&) Or (HiWord% * &H10000)
End Function

Public Function MakeWord(LoByte As Byte, HiByte As Byte) As Integer  '[ fuct0x1
  '[ creates an integer value using Low and High bytes
  '[ useful when converting code from C++
  Dim nLoByte As Integer
  If LoByte < 0 Then
    nLoByte = LoByte + &H100
  Else
    nLoByte = LoByte
  End If
  MakeWord = CInt(nLoByte) Or (HiByte * &H100)
End Function
Public Function HiByte(WordIn As Integer) As Byte  '[ fuct0x1
  If WordIn% And &H8000 Then
    HiByte = &H80 Or ((WordIn% And &H7FFF) \ &HFF)
  Else
    HiByte = WordIn% \ 256
  End If
End Function

Public Function HiWord(LongIn As Long) As Integer  '[ fuct0x1
  HiWord% = (LongIn& And &HFFFF0000) \ &H10000
End Function
Public Function LoByte(WordIn As Integer) As Byte
  LoByte = WordIn% And &HFF&
End Function

Public Function LoWord(LongIn As Long) As Integer
  Dim L As Long
  L& = LongIn& And &HFFFF&
  If L& > &H7FFF Then
    LoWord% = L& - &H10000
  Else
    LoWord% = L&
  End If
End Function

Public Function InIDE() As Boolean
  On Error GoTo InIDEError
  
  InIDE = False
  Debug.Print 1 / 0
Exit Function
InIDEError:
  InIDE = True
Exit Function
End Function
Public Function Percent(Complete&, Total&, TotalOutput&) As Long
  On Error Resume Next
  Percent = Int(Complete / Total * TotalOutput)
End Function

Public Sub runEX(wnd&, ByVal capp$) '[ fuct22.bas - 1998
  Dim win&, mnu&, mCnt&, i&, i2&, i3&, subMnu&, subCnt&
  Dim sSubMnu&, sSubCnt&, Txt$, x&, id&, clk&
  win = wnd
  If win = 0 Then Exit Sub
  mnu = GetMenu(win)
  mCnt = GetMenuItemCount(mnu)
  For i = 0 To mCnt - 1
    subMnu = GetSubMenu(mnu, i)
    subCnt = GetMenuItemCount(subMnu)
    For i2 = 0 To subCnt - 1
      sSubMnu = GetSubMenu(subMnu, i2)
      If sSubMnu Then
        sSubCnt = GetMenuItemCount(sSubMnu)
        For i3 = 0 To sSubCnt - 1
          Txt = Space(256)
          x = GetMenuString(sSubMnu, i3, Txt, 256, WM_USER)
          Txt = fixstring(Txt)
'          Debug.Print "s3(" & Txt & "),(" & capp & ")"
          If InStr(1, LCase(Txt), LCase(capp), vbTextCompare) Then
            id = GetMenuItemID(sSubMnu, i3)
'            Debug.Print "mnuID: " & id
            clk = PostMessage(win, WM_COMMAND, id, 0)
            DoEvents: Exit Sub
          End If
        Next i3
      End If
      Txt = Space(256)
      x = GetMenuString(subMnu, i2, Txt, 256, WM_USER)
      Txt = fixstring(Txt)
'      Debug.Print "s2(" & Txt & "),(" & capp & ")"
      If InStr(1, LCase(Txt), LCase(capp), vbTextCompare) Then
        id = GetMenuItemID(subMnu, i3)
        clk = PostMessage(win, WM_COMMAND, id, 0)
        DoEvents: Exit Sub
      End If
      DoEvents
    Next i2
    DoEvents
  Next i
End Sub
  Public Sub closewindow(Window As Long)
    Call PostMessage(Window&, WM_CLOSE, 0&, 0&)
End Sub
Sub sendText(wnd&, Txt$)
  Call SendMessageByString(wnd, WM_SETTEXT, 0, Txt)
End Sub
Public Sub trace(Text$, ctl As Control)  '[ fuct99.bas
  ctl.SelStart = Len(ctl.Text)
  If ctl.Text = "" Then
    ctl.SelText = Text
  Else
    ctl.SelText = vbCrLf & Text
  End If
  ctl.SelLength = 0
End Sub
Public Sub Pause(duration#): Dim startTime#  '[ fuct22.bas
  startTime = Timer
  Do While Timer - startTime < duration
    DoEvents
  Loop
End Sub
Public Function RemoveSpace(ByVal ttx As String) As String  '[ fuct22.bas
  Dim pos&, Txt$
  Txt = ttx
  While InStr(Txt, Chr(32)) <> 0: DoEvents
    pos = InStr(Txt, Chr(32))
    Txt = Mid(Txt, 1, pos - 1) & Mid(Txt, pos + 1)
  Wend
  RemoveSpace = Txt
End Function
Public Function waitforclass(par&, Class$, Optional indx& = 1) As Long
  Dim wnd&
  '[ remove 'rem' statements if you need this
  '[ waitEND = False
  Do: Pause 0.1
    wnd = FindChildByClass(par, Class, indx)
  Loop Until (wnd <> 0) '[ Or (waitEND = True)
  waitforclass = wnd
  
End Function
Public Function waitfortitle(par&, Txt$) As Long
  Dim wnd&
  '[ remove 'rem' statements if you need this
  '[ waitEND = False
  Do: Pause 0.1
    wnd = FindChildByTitle(par, Txt, True)
  Loop Until (wnd <> 0) '[  Or (waitEND = True)
  waitfortitle = wnd
  
End Function




