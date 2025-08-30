Attribute VB_Name = "vwm_tasks"
Option Explicit
'Public Sub SetBGColor(NEW_BGCOLOR)
'    SetSysColors 1, COLOR_BACKGROUND, NEW_BGCOLOR
'End Sub
'Public Sub GetBGColor()
'    OLD_BGCOLOR = GetSysColor(COLOR_BACKGROUND)
'End Sub


' Public Task Item Structure
Public Type TASK_STRUCT
  TaskClass As String
  TaskTitle As String
  TaskID As Long
  TaskRECT As RECT
  'nDesk(1 To 3) As Integer
End Type

'[ ===( DESKTOPS )=== ]'
Public X_Desk1() As TASK_STRUCT 'Public X_Desk1_COUNT As Long
Public X_Desk2() As TASK_STRUCT 'Public X_Desk2_COUNT As Long
Public X_Desk3() As TASK_STRUCT 'Public X_Desk3_COUNT As Long
Public X_Desk4() As TASK_STRUCT 'Public X_Desk4_COUNT As Long
Public X_Desk5() As TASK_STRUCT
Public X_Desk6() As TASK_STRUCT
Public X_Desk7() As TASK_STRUCT
Public X_Desk8() As TASK_STRUCT
Public X_Desk9() As TASK_STRUCT
Public X_Desk10() As TASK_STRUCT
'[ Public X_LastDesk As Integer
  'Structure filled by FillTaskList Sub call
'[ Public TaskList(1000) As TASK_STRUCT
'[ Public X_TaskCOUNT As Long
Public X_TaskCOUNT(0 To 3) As Long
'// Public X_TaskCOUNT(0 To 9) As Long


'[ ===( STICKEY WINDOWS )=== ]'
Public Const CLASS_START = "[cls]"
Public Const CLASS_END = "[/cls]"
Public Const TITLE_START = "[txt]"
Public Const TITLE_END = "[/txt]"
Public Const SHELL_Wnd = "Shell_TrayWnd"


Public Type WND_STICKY
  sClass As String
  sTitle As String
End Type
Public A_Sticky() As WND_STICKY
Public I_Sticky As Integer      '[ (1 -> MAX)  (0=none)

'[ ===( WALLPAPER)=== ]'
Public Enum wall_style
  CenterWall = 0
  TileWall = 1
  StretchWall = 2
End Enum

Public Const COLOR_BACKGROUND = 1

Public Declare Function SetSysColors Lib "user32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Public Sub ChangeWallpaper(pic_path$, styl As wall_style, Optional bgcolr& = "-1")
  
  Dim cReg As New clsRegistry, ke$, nme_tile$, nme_style$, val_tile$, val_style$
  
  Select Case styl
    Case CenterWall:    val_tile = "0": val_style = "0"
    Case TileWall:      val_tile = "1": val_style = "0"
    Case StretchWall:   val_tile = "0": val_style = "2"
  End Select
    ke = "Control Panel\Desktop"
    nme_tile = "TileWallpaper"
    nme_style = "WallpaperStyle"
      
    If bgcolr > -1 Then
      SetSysColors 1, COLOR_BACKGROUND, bgcolr
    End If
    
    Call cReg.SetValue(eHKEY_CURRENT_USER, ke, nme_tile, val_tile)
    Call cReg.SetValue(eHKEY_CURRENT_USER, ke, nme_style, val_style)
    
  '[ plain background
  If Trim(pic_path) = "" Then pic_path = "(None)"

  Call SystemParametersInfo(SPI_SETDESKWALLPAPER, 2&, ByVal pic_path, (SPIF_SENDWININICHANGE Or SPIF_UPDATEINIFILE))

End Sub
'[
'[ Return TRUE if hWnd is not suppost to be hidden/moved
'[
Public Function DoStickyTest(hWnd&) As Boolean
  Dim ret As Boolean, s_class$, s_title$
  ret = False
  
  s_class = GetClass(hWnd)
  s_title = GrabWndText(hWnd)

  If I_Sticky > 0 Then
    Dim ndx%, tmp_c$, tmp_t$
    Dim bool_c As Boolean, bool_t As Boolean
    
    
    For ndx = 0 To I_Sticky - 1
      tmp_c = A_Sticky(ndx).sClass:   tmp_t = A_Sticky(ndx).sTitle
        bool_c = False:               bool_t = False
        
      If Len(tmp_c) > 0 Then _
        If (comp(s_class, tmp_c) = True) Then bool_c = True
      
      If Len(tmp_t) > 0 Then _
        If (comp(s_title, tmp_t) = True) Then bool_t = True
      
      If Len(tmp_t) > 0 And Len(tmp_c) > 0 Then
        If (bool_t = True) And (bool_c = True) Then _
          ret = True: GoTo END_TEST
      Else
        If bool_t = True Then ret = True: GoTo END_TEST
        If bool_c = True Then ret = True: GoTo END_TEST
      End If
    Next ndx
  Else
    If comp(s_class, SHELL_Wnd) = True Then
      ret = True
    End If
  End If
  
  If (InIDE() = True) Then
    If (comp(GetClass(hWnd), "wndclass_desked_gsk") = True Or _
        comp(GetClass(hWnd), "VbaWindow") = True Or _
        comp(GetClass(hWnd), "VBFloatingPalette") = True Or _
        comp(GetClass(hWnd), "DockingView") = True) Then
      ret = True
    End If
  End If
  
  '==============
  ' old version [pre 0.9]
  'If comp(GetClass(hWnd), "Shell_TrayWnd") Then
  '  ret = True
  'ElseIf comp(GetClass(hWnd), "Winamp v1.x") Then
  '  ret = True
  'ElseIf comp(GrabWndText(hWnd), "Windows Task Manager") Then
  '  ret = True
  'Else
  '  If (InIDE() = True) And (comp(GetClass(hWnd), "wndclass_desked_gsk") = True Or _
  '                           comp(GetClass(hWnd), "VbaWindow") = True Or _
  '                           comp(GetClass(hWnd), "VBFloatingPalette") = True Or _
  '                           comp(GetClass(hWnd), "DockingView") = True) Then
  '    ret = True
  '  Else
  '    ret = False
  '  End If
  'End If

END_TEST:

  DoStickyTest = ret
End Function

Public Sub DrawRects(pbx As PictureBox, deskNdx%)
  
  Dim iDesk%, hWnd&, cnt%, ndx%
  Dim tmpTask() As TASK_STRUCT
  

  Call DeskPULL(deskNdx, tmpTask())
  cnt = X_TaskCOUNT(deskNdx)

  pbx.Cls

  If cnt > 0 Then
    'Debug.Print "Draw Rects-----"
    
'//     fMain.lstApp.Clear
    For ndx = (cnt - 1) To 0 Step -1
      hWnd = tmpTask(ndx).TaskID

      If ndx = 0 Then '[ its last hWnd, make fore window dark
        Call DrawRectsIMG(pbx, deskNdx, hWnd, True)
      Else
        Call DrawRectsIMG(pbx, deskNdx, hWnd)
      End If
    Next ndx
  End If
  
End Sub
Public Sub DrawRectsIMG(pbx As PictureBox, deskNdx%, hWnd&, Optional last_wnd As Boolean = False)
  On Error Resume Next
  
  Dim wnd&, r As RECT, sClass$, i&
  
  Call GetWindowRect(hWnd, r)
  
  Dim RatioX&, RatioY&, ix%, iy%, iw%, ih%
  RatioX = SYS.Width / pbx.Width  '[ pic(deskNdx).Width
  RatioY = SYS.Height / pbx.Height  '[ pic(deskNdx).Height
  
  Dim x1%, x2%, y1%, y2%, colr&
  
  x1 = Int(r.left / RatioX)
  x2 = Int(r.right / RatioX)
  y1 = Int(r.top / RatioY)
  y2 = Int(r.bottom / RatioY)
  
  pbx.Refresh '[ pbx(deskNdx).Refresh
  pbx.AutoRedraw = True '[ pbx(deskNdx).AutoRedraw = True
  pbx.ForeColor = &H0 '[ pbx(deskNdx).ForeColor = &H0
  
  If last_wnd = True Then
    colr = &H808080
  Else: colr = &HC0C0C0
  End If
  'Debug.Print "x1,y1: " & x1 & ", " & y1
  'If x1 = 16 And y1 = 16 Then
  '[
  '[ "IDEOwner"
  '[
  '  MsgBox GetClass(hWnd) & vbCrLf & GrabWndText(hWnd)
  'End If
  pbx.Line (x1, y1)-(x2, y2), , B '[ pbx(deskNdx).Line (x1, y1)-(x2, y2), , B
  pbx.ForeColor = colr    '[ pbx(deskNdx).ForeColor = &HE0E0E0
  pbx.Line (x1 + 1, y1 + 1)-(x2 - 1, y2 - 1), , BF '[ pbx(deskNdx).Line (x1 + 1, y1 + 1)-(x2 - 1, y2 - 1), , BF
  
End Sub


Public Sub DeskPULL(ndx%, newTasks() As TASK_STRUCT)
  
  Select Case ndx
    Case 0: newTasks() = X_Desk1()
    Case 1: newTasks() = X_Desk2()
    Case 2: newTasks() = X_Desk3()
    Case 3: newTasks() = X_Desk4()
    
    Case 4: newTasks() = X_Desk5()
    Case 5: newTasks() = X_Desk6()
    Case 6: newTasks() = X_Desk7()
    Case 7: newTasks() = X_Desk8()
    Case 8: newTasks() = X_Desk9()
    Case 9: newTasks() = X_Desk10()
  End Select
  
End Sub


Public Sub DeskSET(ndx%, dsk() As TASK_STRUCT, cnt As Long)
  
  '[ Setup to hold our global variables
  
  X_TaskCOUNT(ndx) = cnt
  Select Case ndx
    Case 0:      X_Desk1() = dsk()
    Case 1:      X_Desk2() = dsk()
    Case 2:      X_Desk3() = dsk()
    Case 3:      X_Desk4() = dsk()
  
    Case 4:      X_Desk5() = dsk()
    Case 5:      X_Desk6() = dsk()
    Case 6:      X_Desk7() = dsk()
    Case 7:      X_Desk8() = dsk()
    Case 8:      X_Desk9() = dsk()
    Case 9:      X_Desk10() = dsk()
  End Select
End Sub


Public Sub load_sticky()
  '[ this will reload the Type Array of stored Sticky variables
  'Public Type WND_STICKY
  '  sClass As String
  '  sTitle As String
  'End Type
  'Public a_sticky() As WND_STICKY
  'Public i_sticky As Integer      '[ (1 -> MAX)  (0=none)

  '      [cls]Winamp v1.x[/cls]
  '      [cls][/cls][txt]Windows Task Manager[/txt]

  Dim cnt&, arr$()
  
  cnt = ScriptRead("sticky", arr(), True)

  If cnt = 0 Then
    ReDim A_Sticky(0)
    I_Sticky = 1
    A_Sticky(0).sClass = "Shell_TrayWnd"
    
    Call save_sticky
  Else
    Dim ndx&, buff$, ret$
    ReDim A_Sticky(cnt - 1)
    I_Sticky = cnt
    
    For ndx = 0 To cnt - 1
      buff = arr(ndx)
      '
      '  ADD AN ERROR TEST TO CHECK FOR both 'cls' && 'txt' are actual variables
      '
      '  PROBLEMS MAY OCCUR IN FUTURE, Keep in mind 'STUPID PEOPLE'
      '
      ret = ReadVar(buff, "cls")
      If Len(ret) > 0 Then
            A_Sticky(ndx).sClass = ret
      Else: A_Sticky(ndx).sClass = vbNullString
      End If
      
      ret = ReadVar(buff, "txt")
      If Len(ret) > 0 Then
            A_Sticky(ndx).sTitle = ret
      Else: A_Sticky(ndx).sTitle = vbNullString
      End If
      
    Next ndx
    
    '[ ########################################################
    '' check for shell window & add if nessicary
    '[ "Shell_TrayWnd"
    Dim bool As Boolean:    bool = False
    
    For ndx = 0 To cnt - 1
      ret = A_Sticky(ndx).sClass
      If comp(ret, "Shell_TrayWnd") = True Then bool = True
    Next ndx
    If bool = False Then
      ReDim Preserve A_Sticky(cnt)
      I_Sticky = cnt + 1
      A_Sticky(cnt).sClass = "Shell_TrayWnd"
            
      Call save_sticky
    End If
    '[ #########################################################
  End If
End Sub

Public Sub save_sticky()
  '[ this will reload the Type Array of stored Sticky variables
  'Public Type WND_STICKY
  '  sClass As String
  '  sTitle As String
  'End Type
  'Public a_sticky() As WND_STICKY
  'Public i_sticky As Integer      '[ (1 -> MAX)  (0=none)

  '      [cls]Winamp v1.x[/cls]
  '      [cls][/cls][txt]Windows Task Manager[/txt]
  '
  '
  Dim cnt%, sticky_buff$, shel$, sticky_Arr() As String
  Dim ndx%, tmp$, s_cls$, s_txt$, bool As Boolean
  
  
  shel = CLASS_START & "Shell_TrayWnd" & CLASS_END
  cnt = I_Sticky
  sticky_buff = ""

  If cnt > 0 Then
      
    bool = False
    
    For ndx = 0 To cnt - 1
      s_cls = "": s_txt = ""
      
      tmp = A_Sticky(ndx).sClass
      If Len(tmp) > 0 Then _
        s_cls = CLASS_START & tmp & CLASS_END
      
      tmp = A_Sticky(ndx).sTitle
      If Len(tmp) > 0 Then _
        s_txt = TITLE_START & tmp & TITLE_END
        
      sticky_buff = sticky_buff & s_cls & s_txt & vbCrLf
    Next ndx
  End If
  
  If InStr(LCase(sticky_buff), LCase(shel)) = 0 Then
    sticky_buff = sticky_buff & shel
  End If
  
  sticky_Arr() = Split(sticky_buff, vbCrLf)
  bool = ScriptWrite("sticky", sticky_Arr(), "  ")
  
  If bool = False Then
    Debug.Print "Save_Sticky() - Fail"
    MsgBox "Error Saving Sticky Settings", vbCritical, "Save_Sticky()"
  Else
    Debug.Print "Save_Sticky() - Success"
  End If
End Sub
Public Sub WndHide(hWnd&)
  Dim flagz&
  flagz = SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_HIDEWINDOW
  Call SetWindowPos(hWnd, HWND_TOP, 0&, 0&, 0&, 0&, flagz)
End Sub

' Returns if a Process is a Visible Window
Public Function IsTask(hWndTask As Long) As Boolean
    Dim WndStyle As Long
    Const IsTaskStyle = WS_VISIBLE Or WS_BORDER

    WndStyle = GetWindowLong(hWndTask, GWL_STYLE)
    If (WndStyle And IsTaskStyle) = IsTaskStyle Then IsTask = True
End Function

Public Sub FillTaskList(hWnd As Long, tmpTask() As TASK_STRUCT, tmpCount As Long)
'[ Fills the Task structure with captions and hWnd of all active programs
  Dim hWndTask&, intLen&, strTitle$, strClass$, cnt%
  Dim rrect As RECT
  Dim ttmpTask(1000) As TASK_STRUCT
    
  cnt = 0
  '[  process all top-level windows in master window list
  hWnd = FindWindow("Shell_TrayWnd", vbNullString)
  
  hWndTask = GetWindow(hWnd, GW_HWNDFIRST) '[ get first window
  Do While hWndTask '[ repeat for all windows
    If hWndTask <> hWnd And IsTask(hWndTask) Then
    
      '[ ### added 0.9+, 2-12-05
      If DoStickyTest(hWndTask) = False Then
      '[ ==================================
      
        strTitle = GrabWndText(hWndTask)
        strClass = GetClass(hWndTask)
        
        If strTitle = "" Then strTitle = vbNullString
        
        If (Len(strTitle) > 0) Or (Len(strClass) > 0) Then
          If IsWindowVisible(hWnd) > 0 Then
          
            Call GetWindowRect(hWndTask, rrect)
            ttmpTask(cnt).TaskClass = strClass '[ TaskList(cnt).TaskClass = strClass
            ttmpTask(cnt).TaskTitle = strTitle '[ TaskList(cnt).TaskTitle = strTitle
            ttmpTask(cnt).TaskID = hWndTask    '[ TaskList(cnt).TaskID = hwndTask
            ttmpTask(cnt).TaskRECT = rrect
            cnt = cnt + 1
          End If
        End If
      End If
    End If
    hWndTask = GetWindow(hWndTask, GW_HWNDNEXT)
  Loop
  
  tmpTask() = ttmpTask()
  If (cnt > 0) Then
    ReDim Preserve tmpTask(cnt - 1)
    tmpCount = cnt
  Else
    ReDim tmpTask(0)
    tmpCount = 0
  End If
  
End Sub
 Public Sub SwitchTo(hWnd As Long)
  Dim ret As Long
  Dim WStyle As Long ' Window Style bits
  '[ Get style bits for window
  WStyle = GetWindowLong(hWnd, GWL_STYLE)
  '[ If minimized do a restore
  If WStyle And WS_MINIMIZE Then
    ret = ShowWindow(hWnd, SW_RESTORE)
  End If
  
  '[ Move window to top of z-order/activate; no move/resize
  ret = SetWindowPos(hWnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW)
  
End Sub
Public Sub WndShow(hWnd&)
  Dim flagz&
  flagz = SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_SHOWWINDOW
  Call SetWindowPos(hWnd, HWND_TOP, 0&, 0&, 0&, 0&, flagz)
End Sub


