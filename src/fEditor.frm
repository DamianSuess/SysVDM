VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fEditor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Window Editor"
   ClientHeight    =   3975
   ClientLeft      =   4620
   ClientTop       =   3900
   ClientWidth     =   6315
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Height          =   1455
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   16
      Text            =   "fEditor.frx":0000
      Top             =   1200
      Visible         =   0   'False
      Width           =   4995
   End
   Begin MSComctlLib.ImageList imgLst 
      Left            =   0
      Top             =   2460
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fEditor.frx":0082
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fEditor.frx":0156
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   4290
      Top             =   0
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Windowed Processes"
      Height          =   3825
      Left            =   90
      TabIndex        =   4
      Top             =   90
      Width           =   6165
      Begin VB.CommandButton cmdView 
         Caption         =   "make sticky"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   3750
         TabIndex        =   6
         Top             =   3450
         Width           =   1005
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "hide"
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   7
         Top             =   3450
         Width           =   645
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "show"
         Height          =   255
         Index           =   0
         Left            =   2490
         TabIndex        =   8
         Top             =   3450
         Width           =   645
      End
      Begin VB.CommandButton cmdLVRefresh 
         Caption         =   "refresh"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1590
         TabIndex        =   5
         Top             =   3450
         Width           =   735
      End
      Begin VB.CommandButton cmd 
         Caption         =   "wm_close"
         Height          =   255
         Index           =   1
         Left            =   660
         TabIndex        =   10
         Top             =   3450
         Width           =   795
      End
      Begin VB.CommandButton cmd 
         Caption         =   "open"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   11
         Top             =   3450
         Width           =   585
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "close"
         Height          =   255
         Left            =   5340
         TabIndex        =   9
         Top             =   3450
         Width           =   705
      End
      Begin MSComctlLib.ListView lvProc 
         Height          =   3135
         Left            =   90
         TabIndex        =   12
         Top             =   210
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   5530
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Line Line1 
         X1              =   1140
         X2              =   2700
         Y1              =   3570
         Y2              =   3570
      End
   End
   Begin VB.ListBox lstProc 
      Appearance      =   0  'Flat
      Height          =   2505
      Index           =   2
      Left            =   3060
      TabIndex        =   3
      Top             =   4950
      Visible         =   0   'False
      Width           =   2985
   End
   Begin VB.ListBox lstProc 
      Appearance      =   0  'Flat
      Height          =   2505
      Index           =   1
      Left            =   1140
      TabIndex        =   2
      Top             =   4950
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.ListBox lstProc 
      Appearance      =   0  'Flat
      Height          =   2505
      Index           =   0
      Left            =   330
      TabIndex        =   1
      Top             =   4950
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "refresh"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   0
      Top             =   4710
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID:"
      Height          =   165
      Left            =   330
      TabIndex        =   15
      Top             =   4800
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title:"
      Height          =   165
      Left            =   3060
      TabIndex        =   14
      Top             =   4800
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Class Name:"
      Height          =   165
      Left            =   1140
      TabIndex        =   13
      Top             =   4800
      Visible         =   0   'False
      Width           =   795
   End
End
Attribute VB_Name = "fEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Const LVM_FIRST As Long = &H1000
'Private Const LVM_SETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 54)
'Private Const LVS_EX_FLATSB As Long = &H100
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
' SendMessage ListView1.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FLATSB, True
Option Explicit

Private Type x_process
  hWnd As Long
  sClass As String
  sText As String
End Type
Private arrProc() As x_process
Private proc_cnt As Long

Private oLVH As clsLVManip


Private Sub GetTasks()
  
  Dim tsk() As Long '[ store hwnd numbers to windows
  '
  ' // to reduce flicker \\
  ' collect hwnd data and compare the list size to "Private arrTasks()"
  ' if different then
  '   edit the listboxes to accomidate to the size
  ' else
  '   just edit the caption of the list items if needed
  
  
End Sub

Private Sub cmd_Click(Index As Integer)
  Dim sWnd$, wnd&, ndx&
  If lvProc.ListItems.Count = 0 Then Exit Sub
  For ndx = 1 To lvProc.ListItems.Count
    If lvProc.ListItems.Item(ndx).Selected = True Then
      
      'wnd = CLng(lvProc.ListItems.Item(ndx).SubItems(0))
      wnd = CLng(lvProc.ListItems.Item(ndx).Text)
      
      If Index = 0 Then
        'Call ShowWindow(wnd, SW_SHOW)
      Else
        Call SendMessage(wnd, WM_CLOSE, 0, 0)
      End If
    End If
  Next ndx
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdLVRefresh_Click()
  
  proc_cnt = -1
  
  
  Dim hWnd&
  Dim hWndTask&, strTitle$, strClass$
  Dim WndStyle As Long
  
  hWnd = Me.hWnd
   

  '[  process all top-level windows in master window list
  hWnd = FindWindow("Shell_TrayWnd", vbNullString)
  
  hWndTask = GetWindow(hWnd, GW_HWNDFIRST) '[ get first window
  Do While hWndTask '[ repeat for all windows
    'If hWndTask <> hwnd And IsTask(hWndTask) Then
    Const IsTaskStyle = WS_BORDER 'or WS_VISIBLE
    WndStyle = GetWindowLong(hWndTask, GWL_STYLE)
    If (WndStyle And IsTaskStyle) = IsTaskStyle Then
      
      
      strTitle = GrabWndText(hWndTask)
      strClass = GetClass(hWndTask)
      
      If strTitle = "" Then strTitle = vbNullString
      If (Len(strTitle) > 0) Or (Len(strClass) > 0) Then
        proc_cnt = proc_cnt + 1
        ReDim Preserve arrProc(proc_cnt)
'//        If IsWindowVisible(hwnd) > 0 Then    // may use later
        If (hWndTask = Me.hWnd) Or (hWndTask = fMain.hWnd) Then
          '// dont add
        Else
          arrProc(proc_cnt).hWnd = hWndTask
          arrProc(proc_cnt).sClass = strClass
          arrProc(proc_cnt).sText = strTitle
          Debug.Print "wnd: " & hWndTask & "  Class: " & strClass & "    Text: " & strTitle
        End If
        
'//         End If
      End If
    End If
    hWndTask = GetWindow(hWndTask, GW_HWNDNEXT)
  Loop
  
  Debug.Print "########################################################"
  Dim lv_cnt&, lv_ndx&, l_wnd&
  Dim arr_ndx&, found As Boolean
  

  '
  If proc_cnt > -1 Then
    
    If lvProc.ListItems.Count > 0 Then
      For lv_ndx = 1 To lvProc.ListItems.Count
        
        found = False
        If lv_ndx > lvProc.ListItems.Count Then Exit For
        For arr_ndx = 0 To proc_cnt
          If CLng(lvProc.ListItems.Item(lv_ndx).Text) = arrProc(arr_ndx).hWnd Then
            found = True: Exit For
          End If
        Next arr_ndx
        
        If found = False Then
          '// REMOVE ITEM AND RESET COUNT
          Call lvProc.ListItems.Remove(lv_ndx)
          lv_ndx = lv_ndx - 1
        End If
      Next lv_ndx
    End If
    ' --------------------------------------
    Dim lv_item As ListItem, old_wnd&, new_wnd&
    '// Add & Edit all NEW Items
    For arr_ndx = 0 To proc_cnt
      
      new_wnd = arrProc(arr_ndx).hWnd
      
      found = False
      If lvProc.ListItems.Count > 0 Then
        For lv_ndx = 1 To lvProc.ListItems.Count
          
          old_wnd = CLng(lvProc.ListItems.Item(lv_ndx).Text)
          
          If old_wnd = new_wnd Then
            found = True
            Debug.Print "found[0x" & Hex(new_wnd) & "]: update info"
            '// UPDATE Item INFO if needed
            Exit For
          End If
          
        Next lv_ndx
      End If
      
      If found = False Then
        '// Add new item
        Set lv_item = lvProc.ListItems.Add(, , new_wnd)
        lv_item.SubItems(1) = arrProc(arr_ndx).sClass
        lv_item.SubItems(2) = arrProc(arr_ndx).sText
        
        Debug.Print "added new: " & new_wnd
      End If
      
    Next arr_ndx
    
  End If
  
  Const LVM_FIRST = &H1000
  Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
  Const LVSCW_AUTOSIZE_USEHEADER As Long = -2
  
  SendMessage lvProc.hWnd, LVM_SETCOLUMNWIDTH, 2, LVSCW_AUTOSIZE_USEHEADER

  '  // ListView-cleanup
  '  for lv_ndx = 0 to lvProc.MAX
  '    for arr_ndx = 0 to arrProc.MAX
  '      if lvProc(lv_ndx) = arrProc(xxx).hWnd then
  '        found = true
  '    next arr_ndx
  '    if found = false then
  '      //remove item from lvProc.LIST
  '  next lv_ndx
  '
  '  // Add & Edit all New Items
  '  for arr_ndx = 0 to arrProc.MAX
  '    for lv_ndx = 0 to lvProc.MAX
  '      if lvProc(lv_ndx) = arrProc(xxx).hWnd then
  '      {
  '        found = true
  '        update item if needed
  '      }
  '    next lv_ndx
  '    if found = false then
  '      // add new item
  '  next arr_ndx
End Sub
Private Sub cmdRefresh_Click()
  Dim hWnd&
  hWnd = Me.hWnd
  
  Call lstProc(0).Clear
  Call lstProc(1).Clear
  Call lstProc(2).Clear
  
  
  Dim hWndTask&, intLen&, strTitle$, strClass$, cnt%
  Dim rrect As RECT
  'Dim ttmpTask(1000) As TASK_STRUCT
  Dim WndStyle As Long
     
   
  cnt = 0
  '[  process all top-level windows in master window list
  hWnd = FindWindow("Shell_TrayWnd", vbNullString)
  
  hWndTask = GetWindow(hWnd, GW_HWNDFIRST) '[ get first window
  Do While hWndTask '[ repeat for all windows
    'If hWndTask <> hwnd And IsTask(hWndTask) Then
    'If IsTask(hWndTask) Then
    
    Const IsTaskStyle = WS_BORDER 'or WS_VISIBLE
    WndStyle = GetWindowLong(hWndTask, GWL_STYLE)
    If (WndStyle And IsTaskStyle) = IsTaskStyle Then 'IsTask = True
    
      '[ ### added 0.9+, 2-12-05
'      If DoStickyTest(hWndTask) = False Then
      '[ ==================================
        
        strTitle = GrabWndText(hWndTask)
        strClass = GetClass(hWndTask)
        
        If strTitle = "" Then strTitle = vbNullString
        
        If (Len(strTitle) > 0) Or (Len(strClass) > 0) Then
'          If IsWindowVisible(hwnd) > 0 Then
          
            Call GetWindowRect(hWndTask, rrect)
            Call lstProc(0).AddItem("0x" & Hex(hWndTask)) '[ ttmpTask(cnt).TaskID = hWndTask    '// TaskList(cnt).TaskID = hwndTask
            Call lstProc(1).AddItem(strClass) '[ ttmpTask(cnt).TaskClass = strClass '// TaskList(cnt).TaskClass = strClass
            Call lstProc(2).AddItem(strTitle) '[ ttmpTask(cnt).TaskTitle = strTitle '// TaskList(cnt).TaskTitle = strTitle
            '[ ttmpTask(cnt).TaskRECT = rrect
            cnt = cnt + 1
'          End If
        End If
'      End If
    End If
    hWndTask = GetWindow(hWndTask, GW_HWNDNEXT)
  Loop
  
'  tmpTask() = ttmpTask()
'  If (cnt > 0) Then
'    ReDim Preserve tmpTask(cnt - 1)
'    tmpCount = cnt
'  Else
'    ReDim tmpTask(0)
'    tmpCount = 0
'  End If

End Sub


Private Sub cmdView_Click(Index As Integer)
  
  Dim sWnd$, wnd&, ndx&
  If lvProc.ListItems.Count = 0 Then Exit Sub
  For ndx = 1 To lvProc.ListItems.Count
    If lvProc.ListItems.Item(ndx).Selected = True Then
    
      'wnd = CLng(lvProc.ListItems.Item(ndx).SubItems(0))
      wnd = CLng(lvProc.ListItems.Item(ndx).Text)
      If Index = 0 Then
        Call ShowWindow(wnd, SW_SHOW)
      Else
        Call ShowWindow(wnd, SW_HIDE)
      End If
    End If
  Next ndx
  
End Sub


Private Sub Form_Load()
  
  lvProc.ColumnHeaders.Add , , "hWnd:", 794.83          '[ 795
  lvProc.ColumnHeaders.Add , , "Class Name:", 1904.88   '[ 1905
  lvProc.ColumnHeaders.Add , , "Window Text:", 2984.88  '[ 2985
  
  lvProc.ColumnHeaders(1).Tag = "NUMERIC"
  
  lvProc.FullRowSelect = True
  lvProc.MultiSelect = False
  lvProc.LabelEdit = lvwManual
  lvProc.LabelWrap = False
  lvProc.View = lvwReport
  
  
  Call cmdLVRefresh_Click 'cmdRefresh_Click
  
  
  
  '[ list view sorter
  Set oLVH = New clsLVManip
  lvProc.ColumnHeaderIcons = imgLst
  With oLVH
    .Initialise lvProc
    .SortColumnContent 1
  End With
End Sub


Private Sub Form_Unload(Cancel As Integer)
Set oLVH = Nothing
End Sub


Private Sub lstProc_Click(ndx As Integer)
  Dim i%
  i = lstProc(ndx).ListIndex
  lstProc(0).ListIndex = i
  lstProc(1).ListIndex = i
  lstProc(2).ListIndex = i
  
End Sub


Private Sub lvProc_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  'Sort Any of the Column Headers Clicked
  oLVH.SortColumnContent ColumnHeader.Index
End Sub


Private Sub tmr_Timer()
  Dim i%
  i = lstProc(0).ListIndex
  
  Call cmdRefresh_Click
  
  lstProc(0).ListIndex = i
  lstProc(1).ListIndex = i
  lstProc(2).ListIndex = i
  
End Sub


