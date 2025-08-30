VERSION 5.00
Begin VB.Form fMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1425
   ClientLeft      =   3810
   ClientTop       =   2175
   ClientWidth     =   2175
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   95
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   145
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSetup 
      Caption         =   "Setup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdSwitch 
      Caption         =   "switch"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2370
      TabIndex        =   4
      Top             =   1230
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2370
      TabIndex        =   3
      Top             =   990
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox defTab 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   990
      Picture         =   "fMain.frx":08CA
      ScaleHeight     =   525
      ScaleWidth      =   195
      TabIndex        =   28
      Top             =   690
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox defButton 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   600
      Picture         =   "fMain.frx":0E84
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   26
      Top             =   690
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Frame Frame1 
      Caption         =   "old meth  v0.1.x"
      Height          =   2355
      Left            =   2340
      TabIndex        =   14
      Top             =   1560
      Visible         =   0   'False
      Width           =   3975
      Begin VB.CommandButton Command3 
         Caption         =   "shift-right"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1170
         TabIndex        =   19
         Top             =   510
         Width           =   885
      End
      Begin VB.CommandButton Command5 
         Caption         =   "hide-swp"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2220
         TabIndex        =   17
         Top             =   510
         Width           =   885
      End
      Begin VB.CommandButton desk 
         Caption         =   "desk1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   810
         Width           =   765
      End
      Begin VB.CommandButton desk 
         Caption         =   "desk2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   1950
         TabIndex        =   23
         Top             =   840
         Width           =   675
      End
      Begin VB.ListBox List 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Index           =   1
         Left            =   1950
         TabIndex        =   22
         Top             =   1080
         Width           =   1875
      End
      Begin VB.CommandButton Command1 
         Caption         =   "rect"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   510
         Width           =   915
      End
      Begin VB.CommandButton Command2 
         Caption         =   "shift-left"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1170
         TabIndex        =   20
         Top             =   270
         Width           =   885
      End
      Begin VB.CommandButton Command4 
         Caption         =   "show-swp"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2220
         TabIndex        =   18
         Top             =   270
         Width           =   885
      End
      Begin VB.CommandButton Command7 
         Caption         =   "GetChilds"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   270
         Width           =   915
      End
      Begin VB.ListBox List 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   1785
      End
   End
   Begin VB.PictureBox picNdx 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   60
      Index           =   3
      Left            =   1380
      ScaleHeight     =   2
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   13
      Top             =   480
      Width           =   480
   End
   Begin VB.PictureBox picNdx 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   60
      Index           =   2
      Left            =   915
      ScaleHeight     =   2
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   12
      Top             =   480
      Width           =   480
   End
   Begin VB.PictureBox picNdx 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   60
      Index           =   1
      Left            =   450
      ScaleHeight     =   2
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   11
      Top             =   480
      Width           =   480
   End
   Begin VB.PictureBox picNdx 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   60
      Index           =   0
      Left            =   -15
      ScaleHeight     =   2
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   10
      Top             =   480
      Width           =   480
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   450
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   2
      Top             =   -15
      Width           =   480
   End
   Begin VB.CommandButton cmdDrawRect 
      Caption         =   "rect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2370
      TabIndex        =   9
      Top             =   780
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdGather 
      Caption         =   "Gather"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   3
      Left            =   1380
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   7
      Top             =   -15
      Width           =   480
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   2
      Left            =   915
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   6
      Top             =   -15
      Width           =   480
   End
   Begin VB.ListBox lstApp 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   3270
      TabIndex        =   5
      Top             =   690
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.Timer tmrVDM 
      Interval        =   750
      Left            =   1260
      Top             =   690
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   -15
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   0
      Top             =   -15
      Width           =   480
      Begin VB.Label xxWND 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   75
         Index           =   0
         Left            =   330
         TabIndex        =   29
         Top             =   360
         Visible         =   0   'False
         Width           =   105
      End
   End
   Begin VB.PictureBox cmdExit 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   135
      Left            =   1890
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   25
      Top             =   360
      Width           =   135
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FF80FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   1860
      ScaleHeight     =   525
      ScaleWidth      =   195
      TabIndex        =   27
      Top             =   0
      Width           =   195
   End
   Begin VB.Menu mnu 
      Caption         =   "mnu"
      Begin VB.Menu mnu_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_about 
         Caption         =   "sys-vdm 0.0.0"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_opt 
         Caption         =   "Setup Options"
      End
      Begin VB.Menu mnu_edi 
         Caption         =   "Process Editor"
      End
      Begin VB.Menu mnu_sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_lock 
         Caption         =   "Lock VDM"
      End
      Begin VB.Menu mnu_gather 
         Caption         =   "Gather Desktops"
      End
      Begin VB.Menu mnu_tasks 
         Caption         =   "Tasks"
         Begin VB.Menu mnu_tasks_ndx 
            Caption         =   "[ ... task list ... ]"
            Index           =   0
         End
      End
      Begin VB.Menu mnu_sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_desk 
         Caption         =   "Desk 1"
         Index           =   0
      End
      Begin VB.Menu mnu_desk 
         Caption         =   "Desk 2"
         Index           =   1
      End
      Begin VB.Menu mnu_desk 
         Caption         =   "Desk 3"
         Index           =   2
      End
      Begin VB.Menu mnu_desk 
         Caption         =   "Desk 4"
         Index           =   3
      End
      Begin VB.Menu mnu_sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_exit 
         Caption         =   "Close VDM"
      End
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XS, YS As Integer


Private Sub BtnPress(iState%) '[ (picNdx%, iState%)
  Dim pos%
  'If picNdx = 0 Then
    Const xWidth = 9
    Const xHeight = 9
    cmdExit.Cls
    cmdExit.BorderStyle = 0
    cmdExit.AutoRedraw = True
    If iState = 0 Then
      pos = 0           '[ button up
    Else: pos = 10      '[ button down
    End If
    cmdExit.PaintPicture defButton, 0, 0, xWidth, xHeight, pos, 0, xWidth, xHeight
    
  'ElseIf picNdx = 1 Then
  '  '[ do another button... prog later
  '
  'End If
End Sub

Private Sub DeskBGrnd(ndx%)
  
End Sub

Private Sub DeskSwitch(ndx%)
    '[ 1) get current tasks
    '[ 2) store them into the proper desktop
    '[ 3) Hide CURRENT LIST
    '[ 4) get list of new desk
    '[ 5) show NewDesktop, If Any
    '[ 6) Set .LastDesk = CurrentDesk
  
  
  If ndx > SYS.VDM_Count Then
    MsgBox "VDM_OOB:  DeskSwitch(" & ndx & ")", vbCritical, "Desktop Switch Out Of Bounds"
    Exit Sub
  End If
  
' not used picNdx(idx).BackColor = &HE0E0E0
' current  picNdx(s_ndx).BackColor = &HFF8080

  SYS.CurrentDesk = ndx
  If SYS.LastDesk = SYS.CurrentDesk Then Exit Sub

  picNdx(SYS.LastDesk).BackColor = &HE0E0E0
  picNdx(ndx).BackColor = &HFFC0C0:         DoEvents
  
  Me.Enabled = False
  Debug.Print "##################"
  Debug.Print "====[ DESK " & ndx & " ]===="
  Dim cnt%, i%, hWnd&
  Dim tmpTask() As TASK_STRUCT, tmpCnt&
  
  Call FillTaskList(Me.hWnd, tmpTask(), tmpCnt)   '[ -=-=-=-=-=-=-=- #-1 (GET CURRENT TASKS)
    
  cnt = tmpCnt - 1       '[ -=-=-=-=-=-=-=- #-2 (STORE CURRENT)
  If cnt > -1 Then
    Dim sClass$, sTitle$
    
    Call DeskSET(SYS.LastDesk, tmpTask(), tmpCnt)
    
    For i = 0 To cnt          '[ -=-=-=-=-=-=-=- #-3  (HIDE CURRNT DESK)
      hWnd = tmpTask(i).TaskID
      
      If (hWnd <> Me.hWnd) Then '[ And (DoStickyTest(hWnd) = False) Then
        Call WndHide(hWnd)
        DoEvents
      End If
    Next i
  End If
  
  
  picNdx(ndx).BackColor = &HFF8080: DoEvents
  
  
  
  Dim tmpTaskNEW() As TASK_STRUCT ', tmpCnt As Long
  
  Call DeskPULL(SYS.CurrentDesk, tmpTaskNEW())
  tmpCnt = X_TaskCOUNT(SYS.CurrentDesk) - 1
  
  If (tmpCnt > -1) Then       '[ -=-=-=-=-=-=-=- #-5 (SHOW NEW DESK, If Any)
'//    lstApp.Clear
    
    '[ Change Background
    Call DeskBGrnd(ndx)
    
    '[ Show new windows
    For i = tmpCnt To 0 Step (-1)
      hWnd = tmpTaskNEW(i).TaskID
      Call WndShow(hWnd)  '[ Event FORCE TO SHOW Our Window
        
        '[ Debug our tasks
        sTitle = tmpTaskNEW(i).TaskTitle
        sClass = tmpTaskNEW(i).TaskClass
'//        lstApp.AddItem "[0x" & Hex(hWnd) & "][" & sClass & "][" & sTitle & "]"
'//        lstApp.ItemData(lstApp.NewIndex) = hWnd
        DoEvents
    Next i
  End If
    
  '[ -=-=-=-=-=-=-=- #-6 (SET .Last == .Current)
  picNdx(ndx).BackColor = &HFF0000
  DoEvents
  SYS.LastDesk = SYS.CurrentDesk
  Me.Enabled = True
End Sub

Private Sub Load_ME()
  Dim intg As Integer, byt As Byte, ret As String
  
  '[ Load Sticky Windows into Memory ================
  Call load_sticky
  
  '[ Menu Setup ====================
  mnu.Enabled = False
  mnu.Visible = False
  mnu_about.Caption = "sys-vdm v" & App.Major & "." & App.Minor & "." & App.Revision
    intg = CInt(readini("opt", "lockform", "0"))
  If intg = 1 Then mnu_lock.Checked = True Else mnu_lock.Checked = False
  
  '[ Form Setup ====================
  Dim RC As RECT
  Const SPI_GETWORKAREA = 48
  Call SysParamNFO_SCREEN(SPI_GETWORKAREA, 0, RC, 0)
    RC.top = RC.top * Screen.TwipsPerPixelY
    RC.bottom = RC.bottom * Screen.TwipsPerPixelY
    RC.left = RC.left * Screen.TwipsPerPixelX
    RC.right = RC.right * Screen.TwipsPerPixelX
  Me.Height = (picTab.Height * 15) + (15 * 2)
  Me.Width = (cmdExit.left * 15 + cmdExit.Width * 15) + (15 * 4)
  
  Dim xx&, yy&
  ret = readini("opt", "xy-pos", "-1;-1")
    xx = Split(ret, ";")(0)
    yy = Split(ret, ";")(1)
    
    'Debug.Print "sys.width  " & RC.right
    'Debug.Print "sys.height " & RC.bottom
    'Debug.Print "ini.top    " & yy
    'Debug.Print "ini.lft    " & xx
  If (ret = "") Or (InStr(ret, ";") = 0) Or _
     ((yy > RC.bottom) Or (xx > RC.right)) Or (yy = -1 Or xx = -1) Then
    Me.top = (RC.top + RC.bottom) - Me.Height - (15 * 3)
    Me.left = (RC.left + RC.right) - Me.Width - (15 * 1)
  Else
    If (yy + Me.Height) > RC.bottom Then
          Me.top = (RC.top + RC.bottom) - Me.Height - (15 * 3)
    Else: Me.top = yy
    End If
    
    If ((xx + Me.Width) > RC.right) Then
          Me.left = (RC.left + RC.right) - Me.Width - (15 * 1)
    Else: Me.left = xx
    End If
  End If
    'Debug.Print ".left      " & Me.left
    'Debug.Print ".top       " & Me.top
    'Debug.Print ".width     " & Me.Width
    'Debug.Print ".height    " & Me.Height
  
  '[ GRAPHICS =======================
  'PROGRAM THIS FOR SKINS ################################
  '
  picTab.Picture = defTab.Picture
  Call BtnPress(0)
  '
  '
  ' ######################################################
  '
  
  Me.Show
  
  '[ Transparancy ===================
  byt = CByte(readini("opt", "trans", "180"))
  Call SetLayered(Me.hWnd, True, byt)
  
  '[ position =======================
  intg = CInt(readini("opt", "zorder", "0"))
  Select Case intg
    Case 0 ': Debug.Print "Form ON-TOP" '[ ontop
      Call ontop(fMain.hWnd, ZPOS_TopMost)
    Case 1 ': Debug.Print "Form ON-BOTTOM" '[ on desktop
      Call ontop(fMain.hWnd, ZPOS_Desktop)
    Case 2 ': Debug.Print "Form none" '[ no zorder
      Call ontop(fMain.hWnd, ZPOS_Normal)
  End Select
  
  '[ VDM Refresh ====================
  intg = CInt(readini("opt", "refresh", "750"))
  If (intg > &H7FFF) Or (intg < 99) Then intg = 750
  tmrVDM.Interval = intg
                    
End Sub

Private Sub cmdExit_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  'cmdExit.BackColor = &H808080
  Call BtnPress(1)
End Sub


Private Sub cmdExit_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  'cmdExit.BackColor = &HE0E0E0
  Call BtnPress(0)
  
    Debug.Print "x" & x
    Debug.Print "y" & Y
  If (x < 0) Or (x > cmdExit.Width) Then Exit Sub
  If (Y < 0) Or (Y > cmdExit.Height) Then Exit Sub
  
  
  If Button = 1 Then
    
    Dim xx%, yy%
    'vbPopupMenuLeftAlign
    'vbPopupMenuCenterAlign
    'vbPopupMenuRightAlign
    xx = x - (cmdExit.Width / 15)
    yy = Y - ((cmdExit.Height * 15) + (cmdExit.top)) - (fMain.Height / 15)
    'Debug.Print "xx" & xx
    'Debug.Print "yy" & yy
    Call Me.PopupMenu(mnu, vbPopupMenuLeftAlign, xx, yy, mnu_about)
  End If
End Sub

Private Sub cmdGather_Click()
'[ 1) Show All Tasks
'[ 2) Clear ALL BUFFERS()
'[ 3) Save Current List to .CurrentDesk
    
  
  Dim cnt&, hWnd&, flagz&
  Dim tmpDesk() As TASK_STRUCT
  
  Dim ndx%, idx&
  For ndx = 0 To SYS.VDM_Count
    
    pic(ndx).Cls
    
    Call DeskPULL(ndx, tmpDesk())
    cnt = X_TaskCOUNT(ndx) - 1
    If (cnt > -1) Then
    
      For idx = 0 To cnt
        hWnd = tmpDesk(idx).TaskID
        Call WndShow(hWnd): DoEvents
      Next idx
      
    End If
  Next ndx
  
  '[ =================================== ]'
  '[ =====( 2) clear out buff )========= ]'
  '[ =================================== ]'
  
  ReDim X_Desk1(0): X_TaskCOUNT(0) = 0
  ReDim X_Desk2(0): X_TaskCOUNT(1) = 0
  ReDim X_Desk3(0): X_TaskCOUNT(2) = 0
  ReDim X_Desk4(0): X_TaskCOUNT(3) = 0
  
'  ReDim X_Desk5(0): X_TaskCOUNT(5) = 0
'  ReDim X_Desk6(0): X_TaskCOUNT(6) = 0
'  ReDim X_Desk7(0): X_TaskCOUNT(7) = 0
'  ReDim X_Desk8(0): X_TaskCOUNT(8) = 0
'  ReDim X_Desk9(0): X_TaskCOUNT(9) = 0
'  ReDim X_Desk10(0): X_TaskCOUNT(10) = 0
  '[ =================================== ]'
  '[ =====( 3) put all in one )========= ]'
  '[ =================================== ]'
  ReDim tmpDesk(0)
  cnt = 0
  Call FillTaskList(Me.hWnd, tmpDesk(), cnt)  '[ -=-=-=-=-=-=-=- #-1 (GET CURRENT TASKS)
  
  If cnt > -1 Then '[ cnt = X_TaskCOUNT - 1       '[ -=-=-=-=-=-=-=- #-2 (STORE CURRENT)
  '[
  '[  cnt > 0 ????
  '[
    Call DeskSET(0, tmpDesk(), cnt)
    Call DrawRects(pic(SYS.CurrentDesk), SYS.CurrentDesk)
    
  End If
End Sub

Private Sub cmdRefresh_Click()
'[  Dim tmpTask() As TASK_STRUCT, tmpCnt&
'[  Call FillTaskList(Me.hWnd, tmpTask(), tmpCnt)
'[  Call DeskSET(SYS.CurrentDesk, tmpTask(), tmpCnt)
End Sub
Private Sub cmdSetup_Click()  '[ 32*15=480
  
  SYS.VDM_Count = (4 - 1)
  SYS.CurrentDesk = (0)
  SYS.LastDesk = (0)
  '[ SYS.DeskTop = GetDesktopWindow()

  Dim r As RECT
  Call GetWindowRect(GetDesktopWindow(), r)
  SYS.Height = r.bottom ':  Debug.Print "SYS.Width:  " & SYS.Width
  SYS.Width = r.right ':    Debug.Print "SYS.Height: " & SYS.Height
  Debug.Print "-----------"
  
  Dim tmpTask() As TASK_STRUCT, tmpCnt&
  Call FillTaskList(Me.hWnd, tmpTask(), tmpCnt) '[ 2) Get our settings
  Call DeskSET(0, tmpTask(), tmpCnt)            '[ 3) save our settings
                                                '[ 4) Draw it all
  Call DrawRects(pic(SYS.CurrentDesk), SYS.CurrentDesk)
  
'[ use this for Debug Options

'  Dim i&, wnd, sTitle$, sClass$
'  For i = 0 To tmpCnt - 1
'    wnd = tmpTask(i).TaskID
'    sTitle = tmpTask(i).TaskTitle
'    sClass = tmpTask(i).TaskClass
''//    lstApp.AddItem "[0x" & Hex(hWnd) & "][" & sClass & "][" & sTitle & "]"
''//    lstApp.ItemData(lstApp.NewIndex) = wnd
'  Next i
      

End Sub

Private Sub cmdSwitch_Click()
'[  Dim hWnd As Long
'[  If lstApp.ListIndex < 0 Then Beep: Exit Sub
'[
'[  hWnd = lstApp.ItemData(lstApp.ListIndex)
'[  Call SwitchTo(hWnd)
End Sub

Private Sub Command1_Click()
'  Dim r As RECT
'  Dim ndx%, sClass$, sText$, wnd&
'  ndx = List(0).ListIndex
'  If ndx = -1 Then MsgBox "need item": Exit Sub
'
'  sClass = TempTasks(ndx).sClass
'  sText = TempTasks(ndx).sTitle
'
'  wnd = FindWindow(sClass, sText):      Debug.Assert wnd
'  Call GetWindowRect(wnd, r)
'
'  Debug.Print "-------------------"
'  Debug.Print "class: " & sClass
'  Debug.Print "left: " & r.left
'  Debug.Print "top: " & r.top
'  Debug.Print "width: " & r.right - r.left
'  Debug.Print "height: " & r.bottom - r.top
End Sub

Private Sub Command2_Click()
'  Dim r As RECT
'  Dim ndx%, sClass$, sText$, wnd&
'  ndx = List(0).ListIndex
'  If ndx = -1 Then MsgBox "need item": Exit Sub
'
'  sClass = TempTasks(ndx).sClass
'  sText = TempTasks(ndx).sTitle
'  wnd = FindWindow(sClass, sText)
'
'  Debug.Assert wnd
'
'  Call ShiftWnd(wnd, -SYS.Width)
'  'Call GetWindowRect(wnd, r)
'  'Debug.Print "left: " & r.left
'  'Debug.Print "top: " & r.top
'  'Debug.Print "right: " & r.right
'  'Debug.Print "bottom: " & r.bottom
End Sub


Private Sub Command3_Click()
'  Dim r As RECT
'  Dim ndx%, sClass$, sText$, wnd&
'  ndx = List(0).ListIndex
'  If ndx = -1 Then MsgBox "need item": Exit Sub
'
'  sClass = TempTasks(ndx).sClass
'  sText = TempTasks(ndx).sTitle
'  wnd = FindWindow(sClass, sText)
'
'  Debug.Assert wnd
'
'  Call ShiftWnd(wnd, SYS.Width)
'  Debug.Print "Right-Shift"
End Sub


Private Sub Command4_Click()
'  Dim flagz&
'  Dim ndx%, sClass$, sText$, wnd&
'  ndx = List(0).ListIndex
'  If ndx = -1 Then MsgBox "need item": Exit Sub
'
'  sClass = TempTasks(ndx).sClass
'  sText = TempTasks(ndx).sTitle
'  wnd = FindWindow(sClass, sText)
'
'  flagz = SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
'  Call SetWindowPos(wnd, HWND_TOP, 0&, 0&, 0&, 0&, flagz)
'  '[ SendMessage wnd, WM_SHOWWINDOW, 1, 3
End Sub

Private Sub Command5_Click()
'  Dim flagz&
'  Dim ndx%, sClass$, sText$, wnd&
'  ndx = List(0).ListIndex
'  If ndx = -1 Then MsgBox "need item": Exit Sub
'
'  sClass = TempTasks(ndx).sClass
'  sText = TempTasks(ndx).sTitle
'  wnd = FindWindow(sClass, sText)
'  Debug.Print "0x" & Hex(wnd)
'  flagz = SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_HIDEWINDOW
'  Call SetWindowPos(wnd, HWND_TOP, 0&, 0&, 0&, 0&, flagz)
'
'  '[ SendMessage wnd, WM_SHOWWINDOW, 0, 1
End Sub


Private Sub cmdDrawRect_Click()
'[  Call DrawRects(pic(SYS.CurrentDesk), SYS.CurrentDesk)
End Sub


Private Sub Command7_Click()
'[  Dim ndx%, sClass$, sText$, wnd&
'[  ndx = List(0).ListIndex
'[  If ndx = -1 Then MsgBox "need item": Exit Sub
'[
'[  List(1).Clear
'[
'[  sClass = TempTasks(ndx).sClass
'[  sText = TempTasks(ndx).sTitle
'[  wnd = FindWindow(sClass, sText):      Debug.Assert wnd
'[
'[  Call EnumChildWindows(wnd, AddressOf EnumChildProc, &H0)
End Sub

Private Sub desk_Click(Index As Integer)
'  '[ Dim ndx%
'  '[ ndx = Index + 1
'  '[ Call SwitchDesk(ndx)
'
'  SYS.CurrentDesk = (Index)
'  Dim tINF() As TaskINFO
'  List(Index).Clear
'  Call GetTasks(True, tINF())
End Sub


Private Sub Form_Load()

  If App.PrevInstance = True Then
    Call MsgBox("Only one instance of Sys-VDM allowed at a time." & vbCrLf & vbCrLf & _
                "Note:" & vbCrLf & "If you cannot see Sys-VDM, check .ini file for Transparancy level [default=220]", vbInformation)
    End
    Exit Sub
  End If
  
  SCRIPT_PATH = CurrentDir(True) & "sys-vdm.dat"
  
  Call Load_ME
  
  Call cmdSetup_Click
  
  
    '[ Self Update ]'
  If readini("opt", "update", "1") = "1" Then
    Dim bool As Boolean
    bool = fUpdate.CheckUpdate
    If bool = True Then
      fUpdate.Visible = True
      Call ontop(fUpdate.hWnd, ZPOS_TopMost)
      Pause 0.1: Beep
      Pause 0.1: Beep
      Pause 0.1: Beep
      Pause 0.1: Beep
      
      Call fUpdate.ZOrder(0)
    Else
      Unload fUpdate
    End If
  End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = 1 And mnu_lock.Checked = False Then
    Call MoveForm2(Me.hWnd)
  End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
  Call cmdGather_Click
  
  DoEvents
  
  Unload Me
  
  End
End Sub

Private Sub List_DblClick(ndx%)
'  Dim li%, buff$
'  li = List(ndx).ListIndex
'  If li = -1 Then Exit Sub
'
'  buff = List(ndx).List(li)
'  Debug.Print "++++++++++++++++++"
'  Debug.Print buff
'
'  '===========================
'  Dim indx%, sClass$, sText$, wnd&
'  indx = List(0).ListIndex
'  If indx = -1 Then MsgBox "need item": Exit Sub
'
'  sClass = TempTasks(indx).sClass
'  sText = TempTasks(indx).sTitle
'  wnd = FindWindow(sClass, sText)
'  Debug.Print "wnd: 0x" & wnd
End Sub
Private Sub mnu_desk_Click(Index As Integer)
  Call pic_Click(Index)
End Sub

Private Sub mnu_edi_Click()
  fEditor.Show
End Sub

Private Sub mnu_exit_Click()
  Dim ret%
  If readini("opt", "confirm", "1") = "1" Then
    ret = MsgBox("Are you sure you want to close Sys-VDM?", _
            vbQuestion + vbYesNo, "Closing so soon?")
  Else
    ret = vbYes
  End If
  
  If ret = vbYes Then
  
    Call cmdGather_Click
    DoEvents
    Dim xy$
    xy = (Me.left & ";" & Me.top)
    Call writeini("opt", "xy-pos", xy)
    
    End
    
    Exit Sub
  End If
  
End Sub

Private Sub mnu_gather_Click()
  Call cmdGather_Click
End Sub

Private Sub mnu_lock_Click()
  If mnu_lock.Checked = True Then
    mnu_lock.Checked = False
    Call writeini("opt", "lockform", "0")
  Else
    mnu_lock.Checked = True
    Call writeini("opt", "lockform", "1")
  End If
End Sub

Private Sub mnu_opt_Click()
  fSetup.Show
End Sub

Private Sub pic_Click(ndx As Integer)

  '[ private
  Call DeskSwitch(ndx)
  
  On Error GoTo err_hand
  Dim buff$, path$, ret$, bg_color&
  Dim x_ndx%, x_view As wall_style
  
  If readini("desks", "custom", "0") = "1" Then
  
    buff = readini("desks", "desk" & (ndx + 1))
    If InStr(buff, ";") <> 0 Then
      
      path = Trim(Split(buff, ";")(2))        '[ extract path
      ret = Trim(Split(buff, ";")(1))         '[ extract bgColor
      If (IsNumeric(ret) = True) Then
        If (CLng(ret) > -1) Then
          bg_color = CLng(ret)
        End If
      End If
      
      x_ndx = CInt(Split(buff, ";")(0)) - 1   '[ extract view type
      If x_ndx = 0 Then
        x_view = StretchWall
      ElseIf x_ndx = 1 Then
        x_view = CenterWall
      ElseIf x_ndx = 2 Then
        x_view = TileWall
      Else
        Call MsgBox("Error selecting Wallpaper style. Invalid refrence number on desktop " & _
                    ndx + 1 & vbCrLf & vbCrLf & " 0=Stretch 1=Center 2=Tiled", vbCritical, "DeskSwitch()")
      End If
  
      If FileExists(path) = False Then
        Call MsgBox("Invalid Image Path on Desktop " & ndx + 1, vbCritical)
      Else
        Call ChangeWallpaper(path, x_view, bg_color)
      End If
    End If
  End If

Exit Sub
err_hand:
  
  
  Call MsgBox("Error: " & Err & vbCrLf & Error(Err), vbCritical, "DeskSwitch()")


End Sub


Private Sub picNdx_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = 1 And mnu_lock.Checked = False Then
    Call MoveForm2(Me.hWnd)
  End If
End Sub


Private Sub picTab_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = 1 And mnu_lock.Checked = False Then
    Call MoveForm2(Me.hWnd)
  End If
End Sub


Private Sub tmrVDM_Timer()
  Dim tmpTask() As TASK_STRUCT, tmpCnt As Long
'//  Debug.Print "==[ tmrVDM_Timer() ]====="
               
  Call FillTaskList(Me.hWnd, tmpTask(), tmpCnt)
  
  
  '[ 1) Get new task list
  '[  refresh VDM If()
  '[  = Foreground window is different
  '[  = Rect pos changed on any hWnd
  '[  = (count hWnds) != (Last count hWnds)
  '[
  '[ 2) If any contitions met, Refresh RECTS
  '[
  Dim ndx%, nCnt%, ddsk() As TASK_STRUCT
  Dim bFail As Boolean: bFail = False
  
  ndx = SYS.CurrentDesk
  Call DeskPULL(ndx, ddsk())
  nCnt = (X_TaskCOUNT(ndx))
  
  '[[ Debug.Print "nCnt: " & nCnt & "   ;;; tmpCnt(): " & tmpCnt
  
  '[Debug.Print "xTaskCOUNT(" & ndx & "): " & nCnt
  '[Debug.Print "UBound(tmpTask()): " & UBound(tmpTask())
  
  '=========================

  If nCnt <> tmpCnt Then
    bFail = True
    Call DeskSET(ndx, tmpTask(), tmpCnt)
'//    Debug.Print "tmrVDM: != wnd[].count"
  Else
    
    If (nCnt) > 0 Then
    
      Dim ii%, rect1 As RECT, rect2 As RECT
      For ii = 0 To (nCnt - 1)
        rect1 = tmpTask(ii).TaskRECT
        rect2 = ddsk(ii).TaskRECT
        '[ A window have moved, refresh images
        If (rect1.right <> rect2.right) Or (rect1.bottom <> rect2.bottom) Then
          bFail = True
          Call DeskSET(SYS.CurrentDesk, tmpTask, tmpCnt)
'//          Debug.Print "tmrVDM: wnd[].rect <>"
          Exit For
        End If
      Next ii
    End If
  End If
  
  
  If bFail = True Then
'//    Debug.Print "REDRAW"
    Call DrawRects(pic(SYS.CurrentDesk), SYS.CurrentDesk)
  End If
  
  '========================================
'  Static s_ndx As Integer
'  Dim tmpNdx%
'  tmpNdx = SYS.CurrentDesk
'  If s_ndx <> tmpNdx Then
'    s_ndx = tmpNdx
'    Dim idx%
'    For idx = 0 To SYS.VDM_Count
'      If idx <> s_ndx Then _
'        picNdx(idx).BackColor = &HE0E0E0
'    Next idx
'    picNdx(s_ndx).BackColor = &HFF8080
'  End If
  
'//  Debug.Print "==[ end ]================"
  
End Sub

