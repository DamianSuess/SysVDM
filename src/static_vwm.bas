Attribute VB_Name = "old_mod"
Option Explicit
Public m_TaskVisible As Boolean   '[ used for EnumWND()
Public m_MaxDesks As Integer      '[ max windows per desktop <- warning limitations

'=====================================
Public Type TaskINFO
  sClass As String
  sTitle As String
  hWnd As Long
  subWnd() As Long
End Type
Public DeskTasks() As TaskINFO
Public TempTasks() As TaskINFO
'-------------------------------------
Public Type SystemINFO
 ' DeskTop As Long         '[ desktop hwnd
  Width As Long           '[ ScreenWidth    1152 (or) 1024
  Height As Long          '[ ScreenHeight   864        768
  CurrentDesk As Integer  '[ Which Desk We're On  [0 to MAX-1]
  LastDesk As Integer     '[ Desk were Leaving    [0 to MAX-1]
  VDM_Count As Integer    '[ [0-3]  (max-1)
End Type
Public SYS As SystemINFO
'++++++++++++++++++++++++++++++++++++
'  Public Type TaskOLD
'    sClass As String
'    sText As String
'    hWnd As Long
'  End Type
'  Public DeskOLD() As TaskOLD
'++++++++++++++++++++++++++++++++++++



'Public Const WM_SHOWWINDOW = &H18


'Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
'Public Const SWP_HIDEWINDOW = &H80
'Public Const SWP_NOMOVE = &H2
'Public Const SWP_NOSIZE = &H1
'Public Const SWP_NOZORDER = &H4
'Public Const SWP_SHOWWINDOW = &H40


Public Declare Function EnumDesktopWindows Lib "user32" (ByVal hDesktop As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
Public Declare Function EnumDesktops Lib "user32" Alias "EnumDesktopsA" (ByVal hwinsta As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
'==========================
'Public Const HWND_TOP = 0 ' Move to top of z-order
'Public Const SWP_NOSIZE = &H1 ' Do not re-size window
'Public Const SWP_NOMOVE = &H2 ' Do not reposition window
'Public Const SWP_SHOWWINDOW = &H40 ' Make window visible/active
'Public Const GW_HWNDFIRST = 0 ' Get first Window handle
'Public Const GW_HWNDNEXT = 2 ' Get next window handle
'Public Const GWL_STYLE = (-16) ' Get Window's style bits
'Public Const SW_RESTORE = 9 ' Restore window
'Public Const WS_MINIMIZE = &H20000000 ' Style bit 'is minimized'
'Public Const WS_VISIBLE = &H10000000 ' Window is not hidden
'Public Const WS_BORDER = &H800000 ' Window has a border
' Other bits that are normally set include:

'[ write .ini
Public Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Sub writeini(xKey As String, xSub As String, ByVal buff$)
  Dim pth$
  pth = CurrentDir(True) & "vdm.ini"
  Call WritePrivateProfileString(xKey, xSub, buff, pth)
End Sub

Public Function readini(xKey As String, xSub As String, Optional def$ = "") As String
  Dim buff$, pth$
  buff = String(65000, Chr(0))
  pth = CurrentDir(True) & "vdm.ini"
  Call GetPrivateProfileString(xKey, ByVal xSub, def, buff, 65000, pth)
  
  readini = fixstring(buff)
End Function


Public Function EnumChildProc(ByVal hWnd As Long, ByVal lParam As Long) As Boolean
  
  Dim sClass$, sText$
  sClass = GetClass(hWnd)
  sText = GrabWndText(hWnd)
  '[ fMain.List(1).AddItem "[" & hwnd & "][" & sClass & "][" & sText & "]"
  
  EnumChildProc = True
End Function


Public Function EnumWndProc(ByVal hWnd As Long, ByVal lParam As Long) As Boolean
  Dim clas$, Txt$
  Dim desk&, parn&
  Dim dbg$
  Dim exSty As Long, bHasNoOwner As Boolean
  
'[   If IsWindowVisible(hWnd) > 0 Then
      
  bHasNoOwner = (GetWindow(hWnd, GW_OWNER) = 0)
  exSty = GetWindowLong(hWnd, GWL_EXSTYLE)
  
  If (((exSty And WS_EX_TOOLWINDOW) = 0) And bHasNoOwner) Or _
     ((exSty And WS_EX_APPWINDOW) And Not bHasNoOwner) Then


'  If (exSty Or GW_OWNER) Or ((exSty And WS_EX_TOOLWINDOW) = 0) Or (exSty And WS_EX_APPWINDOW) Then
  'If (exSty And WS_EX_TOOLWINDOW) Then
     '[ If (exSty And WS_EX_TOPMOST) Then Debug.Print "TOP"
    
    
    clas = GetClass(hWnd)
    Txt = GrabWndText(hWnd)
    
    '[ If m_TaskVisible = True Then
      If (IsWindowVisible(hWnd) > 0) Then
        Dim ndx%
        ndx = UBound(TempTasks())
        ReDim Preserve TempTasks(ndx + 1)
        TempTasks(ndx).sClass = clas
        TempTasks(ndx).sTitle = Txt
        TempTasks(ndx).hWnd = SYS.CurrentDesk
        
        dbg = "wnd(show): 0x"
        dbg = dbg & Hex(hWnd) & " [" & clas & "][" & Txt & "]"
        fMain.List(SYS.CurrentDesk).AddItem dbg
      End If
    '[ End If
  End If
  If hWnd < 1 Then
    Debug.Print "#####################"
  End If
  
  EnumWndProc = True
End Function
Public Function GetTasks(bVisible As Boolean, tDesk() As TaskINFO)
  m_TaskVisible = bVisible
  ReDim TempTasks(0)
  Call EnumWindows(AddressOf EnumWndProc, &H0)
End Function

'Public Function TaskWindow(wnd&) As Boolean
'  Dim style&
'  style = GetWindowLong(wnd, GWL_STYLE)
'  If (style And (WS_VISIBLE Or WS_BORDER)) = (WS_VISIBLE Or WS_BORDER) Then
'    TaskWindow = True
'  Else
'    TaskWindow = True
'  End If
'End Function

Public Sub ShiftWnd(wnd, direct&)
  If wnd = 0 Then Exit Sub
  
  Dim r As RECT, xL&, xT&, xW&, xH&, xR&, xB&
  Call GetWindowRect(wnd, r)
  
  xL = r.left
  xT = r.top
  xR = r.right
  xB = r.bottom
  xW = r.right - r.left
  xH = r.bottom - r.top
  
  Dim x&, Y&, nWidth&, nHeight&, bool As Boolean
  
  x = (xL + direct)
  Y = xT  '[ were not doing height changes yet
  '[
  '[ All methods work just fine.. jut dont work w/ VB-IDE
  '[
  '==[ meth 1 ]=========
  Call MoveWindow(wnd, x, Y, xW, xH, True)
  
  '==[ meth 2 ]=========
  'Dim dwp&
  'dwp = BeginDeferWindowPos(1 + 16)
  'Call DeferWindowPos(dwp, wnd, &H0, x, y, xW, xH, SWP_NOZORDER Or SWP_NOACTIVATE)
  'Call EndDeferWindowPos(dwp)
  
  '==[ meth 3 ]=========
  'Dim flagz&
  'flagz = SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
  'Call SetWindowPos(wnd, 0, x, y, xW, xH, flagz)
  
End Sub

'Public Sub VDM_Init(Optional Refreshing As Boolean = False)
'  If Refreshing = True Then
'    '[ Set AT PROG-START
'    SYS.LastDesk = 0      '[ StartPOS
'    SYS.CurrentDesk = 0   '[ StartPOS
'    SYS.VDM_Count = 3     '[ default = 3
'    m_MaxDesks = 256    '[ default = 256
'  End If
'    '[  this will hold 256 windows per desktop
'  ReDim DeskWnd(SYS.VDM_Count, m_MaxDesks)
'End Sub


