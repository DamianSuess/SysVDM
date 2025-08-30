Attribute VB_Name = "x_script"
Option Explicit
Public Script_Buff As String
Public SCRIPT_PATH As String '[ make sure to set this


'Public Type UserVariables '[ used for Chat & some Program Options
'  varName As String
'  varData() As String
'End Type
'Public UserVARS() As UserVariables

Private Const BUFFER_SIZE_256k = 262144   ' 256k buffer (ehhh 256 * 1024)
Private Const BUFFER_SIZE_1Mb = 1048576   ' 1 MegaByte buffer (ehhh 1024 * 1024)
Private Const GENERIC_READ = &H80000000
Private Const FILE_SHARE_READ = &H1
Private Const OPEN_EXISTING = 3
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
'Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Type SECURITY_ATTRIBUTES
   nLength As Long
   lpSecurityDescriptor As Long
   bInheritHandle As Boolean
End Type

'Public Sub GetCustom(ret_arr() As String)
'  Dim aBuff() As String, arr$(), tmp$
'  Dim x_str$, i%, cnt%, pos%, vars%
'
'  aBuff() = Split(Script_Buff, vbCrLf)
'  ReDim arr(0)  '[ set default
'
'  x_str = "start $"
'
'  cnt = 0
'  For i = 0 To UBound(aBuff())
'    tmp = aBuff(i)
'
'    If InStr(LCase(tmp), LCase(x_str)) > 0 Then
'        '[ load strings into
'        '[ UserVARS().varName
'
'      pos = InStr(tmp, Chr(32))
'      If pos > 0 Then
'
'        ReDim Preserve arr(cnt)
'        arr(cnt) = Mid(tmp, pos + 1)
'        cnt = cnt + 1
'
'      End If
'      'bfound = False
'      'For vars = (i + 1) To UBound(aBuff())
'      '  '[ load strings into
'      '  '[ UserVARS().varData()
'      '  redim uservars(
'      '
'      'Next vars
'      'If bool = False Then
'      '  Call MsgBox("Critical Error:  'end' statement not found in scripting file.  " & _
'      '              "Please check your script & reload before an even worse error occurs.", vbCritical, "Script Logic Error!")
'[ Get System Version
Private Type TOSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type
Private Declare Function GetVersionEx Lib "Kernel32" Alias "GetVersionExA" (lpVersionInformation As TOSVERSIONINFO) As Long
Private Const VER_PLATFORM_WIN32s = 0         'Windows 3.x is running, using Win32s
Private Const VER_PLATFORM_WIN32_WINDOWS = 1  'Windows 95 or 98 is running.
Private Const VER_PLATFORM_WIN32_NT = 2       'Windows NT is running.

Public Function AlignData(ArrIN$(), Optional strIndent$ = "") As String
  Dim ret$, ndx&, mx&, tmp$, arr() As String
  
  arr() = ArrIN() '[ remove blank lines
  arr() = SplitFix(Join(arr(), vbCrLf), vbCrLf)
  
  mx = ArrSize_s(arr())
  If mx = -1 Then AlignData = "": Exit Function
  
  ret = ""
  For ndx = 0 To mx
    tmp = Trim(arr(ndx))
    If tmp <> "" Then
      ret = ret & strIndent & tmp & vbCrLf
    End If
  Next ndx
  
  If Len(ret) > 2 Then _
    ret = Mid(ret, 1, Len(ret) - 2)
  
  AlignData = ret
End Function

Public Sub filewrite(path$, buffer$, Optional append As Boolean = False)
  On Error GoTo err_hand:
  
  Dim f%
  f = FreeFile
If append = True Then
  Open path For Append As #f
Else
  Open path For Output As #f
End If
    If Len(buffer) > 2 Then
      If (Right(buffer, 2) = vbCrLf) Then _
        buffer = Mid(buffer, 1, Len(buffer) - 2)
    End If
    
    Print #f, buffer
  Close #f
  
Exit Sub
err_hand:
  
  Call MsgBox("Error Number: " & Err & vbCrLf & _
              "Error Message: " & vbCrLf & Error(Err), _
              vbCritical, "filewrite() error")
End Sub

'      'End If
'    End If
'
'  Next i
'
'  ret_arr = arr()
'
'End Sub
Public Function LoadFunction(ByVal fctn_name As String) As String
      '[ start chat_init
      '[   wadup yall~!
      '[   who all here from cali?
      '[   anyone 21 er oldr, im need someone special to talk to
      '[   anyone here lookin for a lil somthin somthin?
      '[   im board anyone wanna get a lil board wif me..  hehehe
      '[ End

  If Trim(Script_Buff) = "" Then
    LoadFunction = ""
'//    Call dbug_x("Error-LoadFunction: Empty Script File!")
    Exit Function
  End If

  Dim crlf$, x_start$, x_end$, head$, ret_block$
  Dim str_pos As Single, end_pos As Single
  Dim xx_str As Single, xx_len As Single

  crlf = vbCrLf
  ret_block = ""

  x_start = "start "
  x_end = crlf & "end"

  head = x_start & fctn_name & crlf

  str_pos = InStr(LCase(Script_Buff), LCase(head))
  If (str_pos > 0) Then

    end_pos = InStr(str_pos, LCase(Script_Buff), x_end)
    If (end_pos > 0) Then
      '[ LOAD FUNCTION
      'Debug.Print "str-pos: " & str_pos
      'Debug.Print "end-pos: " & end_pos

      If (end_pos = (str_pos + Len(head) - 2)) Then
        '[ there's no data
'//        Call dbug_x("Error-LoadFunction: No Data For Function: '" & fctn_name & "'")
        ret_block = ""
      Else
        xx_str = str_pos + Len(head)
        xx_len = (end_pos - (str_pos + Len(head)))

        ret_block = Mid(Script_Buff, xx_str, xx_len)

      '  Debug.Print "str: " & xx_str
      '  Debug.Print "len: " & xx_len
      '  Debug.Print ret_block
      '  Debug.Print "--------"
      End If
    Else

'//      Call dbug_x("Error-LoadFunction: No '<crlf>end'")
'//      Call dbug_x("  Please Re-Edit your Scripting File.. NOW~!")
      Call MsgBox("Missing 'end' statement, please FIX your script file NOW!", vbCritical)
    End If
  Else
'//    Call dbug_x("Error-LoadFunction: Function not found: '" & fctn_name & "'")
  End If
  LoadFunction = ret_block
  
End Function
Public Function LoadScript(path$) As Boolean

  Dim buff$
  '[ buff = fileload(path, True)
  buff = FileBuffer(path, True)
  
  If buff = "" Then
    Script_Buff = ""
    LoadScript = False
  Else
    Dim arr$(), i As Single, tmp$
    
    Debug.Print "-=-=-=-=-=-="
    arr() = SplitFix(buff, vbCrLf) '[ remove & rejoin w/o empty spaces
    For i = 0 To UBound(arr())
      tmp = arr(i)
      arr(i) = Trim(tmp)
      Debug.Print "'[   " & arr(i)
    Next i
    Debug.Print "-=-=-=-=-=-="
    buff = Join(arr(), vbCrLf)
    
    Script_Buff = buff
    LoadScript = True
  End If
  
End Function
Public Function ScriptErase(VarName$) As Boolean
  Dim pth$, buff$, cnt&
  If Trim(SCRIPT_PATH) = "" Then
    Call MsgBox("File i/o Error:  No script file specified.", vbCritical, "ScriptErase() Error")
    Exit Function
  End If
  
  Dim x_var$, x_end$, x_buff$, x_tmp$
re_start:
  
  x_var = "[" & VarName & "]"
  x_end = vbCrLf & "[end]"
  
  If FileExists(SCRIPT_PATH) = False Then Exit Function
  
  Dim str_pos&, end_pos&, file_endd&, ret%, ln&
  Dim lft$, rgh$, xnd$, nuBuff$
  
  buff = fileload(SCRIPT_PATH, True)
  file_endd = Len(buff)
  str_pos = InStr(1, buff, x_var)
  If (str_pos > 0) Then
    
    end_pos = InStr(str_pos + 1, buff, x_end)
    If (end_pos = 0) Then
      ret = MsgBox("Repair Script NOW!  Missing '[end]' tag." & vbCrLf & vbCrLf & _
                    "Press 'yes' to retry w/ changes made" & vbCrLf & _
                    "Press 'no' to exit erase process.", vbCritical & vbYesNo, "Script Erase Error")
      If ret = vbYes Then GoTo re_start
      ScriptErase = False
      Exit Function
    End If
  Else
    '[ were okay LEAVE~!
    ScriptErase = True
    Exit Function
  End If
  
  Dim LEFT_END&, RIGHT_STR&, x_left$, x_right$
  
  If (str_pos = 1) And (end_pos >= file_endd) Then
                                          Debug.Print "ScriptErase() Only Value Cleared"
    x_buff = vbCrLf
    Call filewrite(SCRIPT_PATH, x_buff)
    ScriptErase = True  '[ ### this isn't called at all ### ]'
   
  ElseIf (str_pos = 1) And (end_pos < file_endd) Then
                                          Debug.Print "ScriptErase() Value at Beginning, Get RIGHT"
    RIGHT_STR = (end_pos + Len(x_end))
    x_right = Mid(buff, RIGHT_STR)
   
    x_buff = x_right
    Call filewrite(SCRIPT_PATH, x_buff)
    ScriptErase = True
   
  ElseIf (str_pos > 1) And (end_pos < file_endd) Then
                                          Debug.Print "ScriptErase() Value in Middle, Get LEFT & RIGHT"
    LEFT_END = str_pos - 1
    RIGHT_STR = (end_pos + Len(x_end))
    
    x_left = Mid(buff, 1, LEFT_END)
    x_right = Mid(buff, RIGHT_STR)
    
    If Right(x_left, 2) = vbCrLf Then
      x_left = Mid(x_left, 1, Len(x_left) - 2)
    End If
    
    x_buff = x_left & vbCrLf
    x_buff = x_buff & x_right
    Call filewrite(SCRIPT_PATH, x_buff)
    ScriptErase = True
   
  ElseIf (str_pos > 1) And (end_pos >= file_endd) Then
                                          Debug.Print "ScriptErase() Value at End, Get LEFT"
    LEFT_END = str_pos - 1
    x_left = Mid(buff, 1, LEFT_END)
   
    x_buff = x_left
    Call filewrite(SCRIPT_PATH, x_buff)
    ScriptErase = True  '[ === this isnt called at all === ]'
   
  Else
    Call MsgBox("ScriptErase()" & vbCrLf & vbCrLf & "Unknown Error occured.", vbCritical)
    ScriptErase = False
  End If
End Function

Private Function SplitFix(ByVal Expression As String, _
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
  SplitFix = Filter(varItems, Delimiter, False)
End Function

Public Function FileBuffer(ptFileName As String, pfSucces As Boolean) As String
  Dim lFile   As Long
  Dim lRet    As Long
  Dim lBytes  As Long
  Dim lSecAtt As SECURITY_ATTRIBUTES
  Dim tBuffer As String
  Dim tDestin As String
  tDestin = ""
  
  If (IsWin2000Plus = True) Then
   lSecAtt.nLength = Len(lSecAtt)   ' size of the structure
   lSecAtt.lpSecurityDescriptor = 0 ' default (normal) level of security
   lSecAtt.bInheritHandle = 1       ' this is the default setting
   
   lFile = CreateFile(ptFileName, GENERIC_READ, FILE_SHARE_READ, ByVal lSecAtt, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
  Else
   lFile = CreateFile(ptFileName, GENERIC_READ, FILE_SHARE_READ, ByVal CLng(0), OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
  End If
  
  If lFile = -1 Then '[ the file could not be opened
    FileBuffer = ""
    pfSucces = False
    Exit Function
  End If
  
  tBuffer = String(BUFFER_SIZE_1Mb, Chr(0))
  lRet = ReadFile(lFile, ByVal tBuffer, BUFFER_SIZE_1Mb, lBytes, ByVal CLng(0))
  
  If lBytes = 0 Then
    '[ Check for EOF
    FileBuffer = ""
    pfSucces = False
  Else
    FileBuffer = ""
    tDestin = Left(tBuffer, lBytes)
    Do While lBytes > 0
      lRet = ReadFile(lFile, ByVal tBuffer, BUFFER_SIZE_1Mb, lBytes, ByVal CLng(0))
      tDestin = tDestin & Left(tBuffer, lBytes)
    Loop
  End If
  lRet = CloseHandle(lFile)
  
  '[ Return contents
  FileBuffer = tDestin
  pfSucces = True
  tDestin = ""  '[ Clear memory
End Function
Private Function IsWin2000Plus() As Boolean
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


Public Function fileload(ByVal path$, Optional by_line As Boolean = False) As String
  
  If FileExists(path) = False Then
    fileload = ""
    Exit Function
  End If
  
  Dim buff$, f%, tmp$
  f = FreeFile
  buff = ""
  
  Open path For Input As #f
    Do Until EOF(f)
      If (by_line = True) Then
        Line Input #f, tmp
        buff = buff & tmp & vbCrLf
      Else
        Input #f, tmp
        buff = buff & tmp
      End If
    Loop
  Close #f
  
  If (by_line = True) Then
    If (Right(buff, 2) = vbCrLf) And Len(buff) > 2 Then _
      buff = Mid(buff, 1, Len(buff) - 2)
  End If
  fileload = buff
End Function


Private Function FileExists(pth As String) As Boolean '[ fuct22.bas
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

Public Function ReadVar(buff$, VarName$) As String
  Dim x_var$, x_end
  Dim pos1&, pos2&, ln&, datlen&, ret$
  
  x_var = "[" & VarName & "]"
  x_end = "[/" & VarName & "]"
  
  pos1 = InStr(1, buff, x_var, 1)
  If (pos1 = 0) Then
    ReadVar = vbNullString
    Exit Function
  End If
  
  pos2 = InStr(pos1 + 1, buff, x_end, 1)
  If (pos2 = 0) Then
    MsgBox "Error: No end tag for variable '" & x_var & "'" & vbCrLf & vbCrLf & _
           "Please repair your script before a major error occurs.", vbCritical, "Script Error"
    ReadVar = vbNullString
    Exit Function
  End If
  
  ln = Len(x_var)
  If (pos1 + ln = pos2) Then
    ReadVar = vbNullString
    Exit Function
  End If
  
  datlen = pos2 - (pos1 + ln)
  ret = Mid(buff, (pos1 + ln), datlen)
  
  ReadVar = ret
End Function

Public Function ScriptRead(ByVal VarName$, RetArr() As String, Optional trim_9_32 As Boolean = False) As Long
  Dim pth$, buff$, cnt&, str_pos&
  
  pth = SCRIPT_PATH
  buff = fileload(pth, True)
  
  cnt = 0
  str_pos = InStr(LCase(buff), "[" & VarName & "]")
  If (str_pos = 0) Then
    ReDim RetArr(0)
    ScriptRead = 0
  Else
    Dim end_pos&
    end_pos = InStr(str_pos + 2, LCase(buff), vbCrLf & "[end]")
    If (end_pos = 0) Then
      ReDim RetArr(0)
      ScriptRead = 0
      Call MsgBox("Script Error:  Variable w/o '[end]' tag." & vbCrLf & vbCrLf & _
                   "Variable: '" & VarName & "'", vbCritical, "Fix Your Script")
      Exit Function
    End If
    Dim i_left&, i_len&, scr_buff$, arr$()
    
    i_left = (str_pos + 2) + Len(VarName)
    i_len = (end_pos) - i_left  '(end_pos - 1) - i_left
    
    scr_buff = Mid(buff, i_left, i_len)
    arr() = SplitFix(scr_buff, vbCrLf)
    cnt = ArrSize_s(arr())
    
    '[ remove chr(9) and Trim() Spaces ============
    If (cnt > -1) And (trim_9_32 = True) Then
      Dim i&, tmp$, ta$()
      For i = LBound(arr()) To UBound(arr())
        tmp = arr(i)
        tmp = Replace(tmp, Chr(9), "")
        tmp = Trim(tmp)
        arr(i) = tmp
      Next i
      tmp = Join(arr(), vbCrLf)       '[ fix blank spaces
      arr() = SplitFix(tmp, vbCrLf)
      cnt = ArrSize_s(arr())
    End If
    '[ ============================================
    RetArr = arr()
    ScriptRead = (cnt + 1)
  End If
  
End Function
Private Function ArrSize_s(arr() As String) As Long '[ takes in String[] array
  On Error GoTo err_hand
  
  Dim cnt&
  cnt = UBound(arr())
  ArrSize_s = cnt
  
Exit Function
err_hand:
  
  ArrSize_s = -1
  
End Function
Public Function ScriptWrite(VarName$, VarData$(), Optional Auto_Indent As String = "  ") As Boolean
  ' Takes In:
  '  str FileLoad(pth$)             DONE
  '  void AppendFile(path$, buff$)  DONE
  '  void FileWrite(pth, buff)      DONE
  '------------------
  '  Need to Finish this Function
  ' and do THROUGH testing
  '
  '  file_endd = LOF(#f)
  ' 1]  (str_pos=1) && (end_pos >= file_endd)
  ' 2]  (str_pos=1) && (end_pos <  file_endd)
  ' 3]  (str_pos>1) && (end_pos <  file_endd)
  ' 4]  (str_pos>1) && (end_pos >= file_endd)
  '__________________________________________
  Dim pth$, buff$, cnt&
  If Trim(SCRIPT_PATH) = "" Then
    Call MsgBox("File i/o Error:  No script file specified.", vbCritical, "ScriptWrite() Error")
    Exit Function
  End If
  
  Dim x_var$, x_end$, x_buff$, x_tmp$
re_start:
  
  x_var = "[" & VarName & "]"
  x_end = vbCrLf & "[end]"
  
  If FileExists(SCRIPT_PATH) = False Then
  
    x_buff = x_var & vbCrLf & AlignData(VarData(), Auto_Indent) & x_end
    Call filewrite(SCRIPT_PATH, x_buff)
    ScriptWrite = False
    Exit Function
  End If
  
  Dim str_pos&, end_pos&, file_endd&, ret%, ln&
  Dim lft$, rgh$, xnd$, nuBuff$
  
  buff = fileload(SCRIPT_PATH, True)
  file_endd = Len(buff)
  str_pos = InStr(1, buff, x_var)
  If (str_pos > 0) Then
    
    end_pos = InStr(str_pos + 1, buff, x_end)
    If (end_pos = 0) Then
      ret = MsgBox("Repair Script NOW!  Missing '[end]' tag." & vbCrLf & vbCrLf & _
                    "Press 'yes' to retry w/ changes made" & vbCrLf & _
                    "Press 'no' to exit saving process.", vbCritical & vbYesNo, "Script Write Error")
      If ret = vbYes Then GoTo re_start
      ScriptWrite = False
      Exit Function
    End If
  Else
    '[ APPEND, IT's a NEW Value
    Debug.Print "ScriptWrite: New Value Saved"
    x_buff = x_var & vbCrLf
    x_buff = x_buff & AlignData(VarData(), Auto_Indent)
    x_buff = x_buff & x_end
    Call filewrite(SCRIPT_PATH, x_buff, True)
    ScriptWrite = True
    Exit Function
  End If
  
  Dim LEFT_END&, RIGHT_STR&, x_left$, x_right$
  
  If (str_pos = 1) And (end_pos >= file_endd) Then
                                          Debug.Print "ScriptWrite() Value Rewritten"
    x_buff = x_var & vbCrLf
    x_buff = x_buff & AlignData(VarData(), Auto_Indent)
    x_buff = x_buff & x_end
    Call filewrite(SCRIPT_PATH, x_buff)
    ScriptWrite = True  '[ ### this isn't called at all ### ]'
    
  ElseIf (str_pos = 1) And (end_pos < file_endd) Then
                                          Debug.Print "ScriptWrite() Value at Beginning, Get RIGHT"
    RIGHT_STR = (end_pos + Len(x_end))
    x_right = Mid(buff, RIGHT_STR)
    
    x_buff = x_var & vbCrLf
    x_buff = x_buff & AlignData(VarData(), Auto_Indent)
    '[ x_buff = x_buff & x_end & vbcrlf & x_right         // will add unwanted spaces
    x_buff = x_buff & x_end & x_right
    Call filewrite(SCRIPT_PATH, x_buff)
    ScriptWrite = True
    
  ElseIf (str_pos > 1) And (end_pos < file_endd) Then
                                          Debug.Print "ScriptWrite() Value in Middle, Get LEFT & RIGHT"
    LEFT_END = str_pos - 1
    RIGHT_STR = (end_pos + Len(x_end))
    x_left = Mid(buff, 1, LEFT_END)
    x_right = Mid(buff, RIGHT_STR)
    
    x_buff = x_left & x_var & vbCrLf
    x_buff = x_buff & AlignData(VarData(), Auto_Indent)
    x_buff = x_buff & x_end & x_right
    Call filewrite(SCRIPT_PATH, x_buff)
    ScriptWrite = True
    
  ElseIf (str_pos > 1) And (end_pos >= file_endd) Then
                                          Debug.Print "ScriptWrite() Value at End, Get LEFT"
    LEFT_END = str_pos - 1
    x_left = Mid(buff, 1, LEFT_END)
    
    x_buff = (x_left & x_var & vbCrLf)
    x_buff = x_buff & AlignData(VarData(), Auto_Indent)
    x_buff = x_buff & x_end
    Call filewrite(SCRIPT_PATH, x_buff)
    ScriptWrite = True  '[ === this isnt called at all === ]'
    
  Else
    Call MsgBox("ScriptWrite()" & vbCrLf & vbCrLf & "Unknown Error occured.", vbCritical)
    ScriptWrite = False
  End If
  
End Function

