Attribute VB_Name = "cmdial0g_2"
Option Explicit
'[ do this cause fuct told you to

'[ History
'[ -----------------------
'[  01-04-2004   added random updates
'[  01-06-2004   added color support for VDM

Type CHOOSECOLORS
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Public Type SelectedColor
    colr As Long ' OLE_COLOR
    bCanceled As Boolean
End Type
Public Const CC_RGBINIT = &H1
Public Const CC_FULLOPEN = &H2
Public Const CC_PREVENTFULLOPEN = &H4
Public Const CC_SHOWHELP = &H8
Public Const CC_ENABLEHOOK = &H10
Public Const CC_ENABLETEMPLATE = &H20
Public Const CC_ENABLETEMPLATEHANDLE = &H40
Public Const CC_SOLIDCOLOR = &H80
Public Const CC_ANYCOLOR = &H100
Public Const COLOR_FLAGS = CC_FULLOPEN Or CC_ANYCOLOR Or CC_RGBINIT

Private Type OPENFILENAME
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
  pvReserved As Long 'new Win2000 / WinXP members
  dwReserved As Long 'new Win2000 / WinXP members
  FlagsEx    As Long 'new Win2000 / WinXP members
End Type

Public Const OFN_READONLY = &H1
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_SHOWHELP = &H10
'Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOLONGNAMES = &H40000 ' force no long names for 4.x modules
Public Const OFN_EXPLORER = &H80000 ' new look commdlg
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_LONGNAMES = &H200000 ' force long names for 3.x modules
Public Const OFN_ENABLESIZING As Long = &H800000
Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHAREWARN = 0

Public Const OFS_MAXPATHNAME As Long = 260

'Public Function CurrentDir() 'returns path w/o "\"
'  '[ fuct 1998
'  Dim cDir$
'  If Mid(App.Path, Len(App.Path), 1) = Chr$(89 + 3) Then
'    cDir = Mid(App.Path, 1, Len(App.Path) - 1)
'  Else: cDir = App.Path
'  End If
'  CurrentDir = cDir
'End Function

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
'Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONTS) As Long
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLORS) As Long
'Private Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLGS) As Long

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long


Private Type TOSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type
Private Declare Function GetVersionEx Lib "Kernel32" Alias "GetVersionExA" (lpVersionInformation As TOSVERSIONINFO) As Long

Private Const VER_PLATFORM_WIN32s = 0        'Windows 3.x is running, using Win32s
Private Const VER_PLATFORM_WIN32_WINDOWS = 1 'Windows 95 or 98 is running.
Private Const VER_PLATFORM_WIN32_NT = 2      'Windows NT is running.

' =  OFN_EXPLORER Or OFN_LONGNAMES Or _
' OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY
' OFS_FILE_SAVE_FLAGS Or OFN_ENABLEHOOK Or OFN_ENABLESIZING

Public Const OFS_FILE_OPEN_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
Public Const OFS_FILE_SAVE_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY '[ OFS_FILE_SAVE_FLAGS Or OFN_ENABLEHOOK Or OFN_ENABLESIZING
Public Const OFN_FUCT = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_CREATEPROMPT Or OFN_ENABLESIZING Or OFN_EXPLORER Or OFN_LONGNAMES


Public Sub lockWnd(wnd&)
  Call LockWindowUpdate(wnd)
End Sub
Public Function saveCMD(frm As Form, filter$, Title$, initDir$, Optional flags = OFN_FUCT, Optional defFile$ = "") As String
  
  Dim OFN As OPENFILENAME, a&
  OFN.lStructSize = Len(OFN)
  OFN.hwndOwner = frm.hWnd
  OFN.hInstance = App.hInstance
  If right$(filter, 1) <> "|" Then filter = filter + "|"


  For a = 1 To Len(filter)
    If Mid$(filter, a, 1) = "|" Then Mid$(filter, a, 1) = Chr$(0)
  Next
  If defFile = "" Then defFile = Space$(254)
  OFN.lpstrFilter = filter
  OFN.lpstrFile = defFile '['[ .sFile = defFile & Chr(0)
  OFN.nMaxFile = 255
  OFN.lpstrFileTitle = Space$(254)
  OFN.nMaxFileTitle = 255
  OFN.lpstrInitialDir = initDir
  OFN.lpstrTitle = Title
  OFN.flags = flags
  
  a = GetSaveFileName(OFN)
  
  '[ added by fuct 9.3.03 ]'
  Dim tmp$, arr$, ext%, file$
  ext = OFN.nFilterIndex
  file = OFN.lpstrFile
  tmp = OFN.lpstrFile
  If ext > 0 Then
    arr = Split(filter, Chr(0))((ext * 2) - 1)
    If arr <> "" Then
      arr = fixstring(arr, True)
      tmp = fixstring(tmp, True)
      If arr = "*.*" Then
        file = tmp
      Else
        file = Replace(arr, "*", tmp)
      End If
    End If
  End If
  '-[ end fuct ]-'

  If (a) Then
    saveCMD = Trim$(file) 'ofn.lpstrFile)
  Else
    saveCMD = ""
  End If
End Function
Private Function RemoveSpace(ByVal ttx As String) As String
  Dim pos&, Txt$
  Txt = ttx
  While InStr(Txt, Chr(32)) <> 0: DoEvents
    pos = InStr(Txt, Chr(32))
    Txt = Mid(Txt, 1, pos - 1) & Mid(Txt, pos + 1)
  Wend
  RemoveSpace = Txt
End Function
Public Sub dupekill2(lst As Control)
 Dim dTxt$, rTxt$, i&, x&, tCnt&, ccp$, nndx&
 dTxt = "": rTxt = ""
 tCnt = lst.ListCount - 1
 
 'ccp = fMain.Caption
 
 For x = 0 To tCnt
  rTxt = TrimSuper(lst.List(x))
  For i = x To tCnt
   If i = x Then i = i + 1
   dTxt = TrimSuper(lst.List(i))
   
   If rTxt = dTxt Then
    lst.RemoveItem i
    tCnt = lst.ListCount - 1
    i = i - 1
   End If
   
'   nndx = nndx + 1
'   If nndx Mod 512 = 0 Then
'    nndx = 0
'    fMain.Caption = "duplicate: " & Percent(x, tCnt, 100) & "% complete..."
'    DoEvents
'   End If
  Next i
 Next x
 
' fMain.Caption = ccp
End Sub
Private Function TrimSuper(Txt As String) As String
  TrimSuper = LCase(RemoveSpace(Txt))
  DoEvents
End Function
Public Function openCMD(frm As Form, filter$, Title$, initDir$, Optional iFlags As Long = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST) As String
    
  Dim OFN As OPENFILENAME, a&
  OFN.lStructSize = Len(OFN)
  OFN.hwndOwner = frm.hWnd
  OFN.hInstance = App.hInstance
  If right$(filter, 1) <> "|" Then filter = filter + "|"
  
  
  For a = 1 To Len(filter)
    If Mid$(filter, a, 1) = "|" Then Mid$(filter, a, 1) = Chr$(0)
  Next
  OFN.lpstrFilter = filter
  OFN.lpstrFile = Space$(254)
  OFN.nMaxFile = 255
  OFN.lpstrFileTitle = Space$(254)
  OFN.nMaxFileTitle = 255
  OFN.lpstrInitialDir = initDir
  OFN.lpstrTitle = Title
  OFN.flags = iFlags 'OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
  a = GetOpenFileName(OFN)
  
  
  If (a) Then
    openCMD = Trim$(OFN.lpstrFile)
  Else
    openCMD = ""
  End If
End Function
Public Function ShowColor(ByVal hWnd As Long) As SelectedColor
  
  Dim cc As CHOOSECOLORS
  Dim ret&
  
  cc.hwndOwner = hWnd
  cc.lpCustColors = ""
  cc.lStructSize = Len(cc)
  cc.flags = COLOR_FLAGS
  
  ret = ChooseColor(cc)
  If ret Then
    ShowColor.bCanceled = False
    ShowColor.colr = cc.rgbResult  '[ ShowColor.oSelectedColor = cc.rgbResult
  Else
    ShowColor.bCanceled = True
    ShowColor.colr = 0&            '[ ShowColor.oSelectedColor = &H0&
  End If

'  Dim customcolors() As Byte  ' dynamic (resizable) array
'  Dim i As Integer
'  Dim ret As Long
'  Dim hInst As Long
'  Dim Thread As Long
  
'  ParenthWnd = hWnd
'  If ColorDialog.lpCustColors = "" Then
'    ReDim customcolors(0 To 16 * 4 - 1) As Byte  'resize the array
'    For i = LBound(customcolors) To UBound(customcolors)
'      customcolors(i) = 254 ' sets all custom colors to white
'    Next i
'    ColorDialog.lpCustColors = StrConv(customcolors, vbUnicode)  ' convert array
'  End If
'
'  ColorDialog.hwndOwner = hWnd
'  ColorDialog.lStructSize = Len(ColorDialog)
'  ColorDialog.flags = COLOR_FLAGS
'
'  'Set up the CBT hook
'  hInst = GetWindowLong(hWnd, GWL_HINSTANCE)
'  Thread = GetCurrentThreadId()
'  If centerForm = True Then
'    hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
'  Else
'    hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
'  End If
'
'  ret = ChooseColor(ColorDialog)
'  If ret Then
'    ShowColor.bCanceled = False
'    ShowColor.oSelectedColor = ColorDialog.rgbResult
'    Exit Function
'  Else
'    ShowColor.bCanceled = True
'    ShowColor.oSelectedColor = &H0&
'    Exit Function
'  End If
  
End Function
Private Function fixstring(Txt As String, Optional rmvSpace As Boolean = False) As String
  If InStr(1, Txt, Chr(0), 1) Then
    Txt = Mid(Txt, 1, InStr(1, Txt, Chr(0), 1) - 1)
    If rmvSpace = True Then Txt = Trim(Txt)
  End If
  fixstring = Txt
End Function

