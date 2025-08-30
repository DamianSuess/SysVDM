VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form fUpdate 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3060
   ClientLeft      =   7125
   ClientTop       =   2385
   ClientWidth     =   3150
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
   Icon            =   "fUpdate.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   3150
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton fCommand1 
      Caption         =   "Update Executable"
      Height          =   765
      Left            =   1860
      TabIndex        =   7
      Top             =   1920
      Width           =   1125
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Download Update"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1260
      Width           =   2685
   End
   Begin VB.CommandButton Command2 
      Caption         =   "save link to..."
      Height          =   315
      Left            =   30
      TabIndex        =   2
      Top             =   2190
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdCheckUpdate 
      Caption         =   "check update"
      Height          =   285
      Left            =   30
      TabIndex        =   3
      Top             =   1830
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSWinsockLib.Winsock wskU 
      Left            =   0
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Would you like to update from:  Current Version: v0.6.66     Online Version: v0.7.77"
      Height          =   615
      Left            =   390
      TabIndex        =   5
      Top             =   510
      Width           =   2355
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "An update has been found:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label lbClose 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " X "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2910
      TabIndex        =   0
      ToolTipText     =   " -cancel changes- "
      Top             =   0
      Width           =   225
   End
   Begin VB.Label lbMove 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "    -self updating system-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "fUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private old_dir$
Private nwY2&, nwX2&
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Update_ONLINE As Boolean
Private Sub dl_link(ByVal path$)
  Const SW_NORMAL = 1
  '[ path = "http://www.vbstatic.net/update/FILE.exe"
  
  Call ShellExecute(Me.hWnd, "open", path$, vbNullString, vbNullString, SW_NORMAL)

End Sub

Private Function DownloadFile(URL As String, LocalFilename As String) As Boolean
    Dim lngRetVal As Long
    lngRetVal = URLDownloadToFile(0, URL, LocalFilename, 0, 0)
    If lngRetVal = 0 Then DownloadFile = True

'[ ret = DownloadFile("http://www.URL Of The File", _
'[ "c:\Which Dir To Save The File To.jpg, exe etc.")

End Function




Private Sub MoveForm(f As Form, ByVal btn, ByVal x, ByVal Y): On Error Resume Next
  '[ fuct98.bas
  If btn = 1 Then
    f.Move (f.left + (x - nwX2) / 6), (f.top + (Y - nwY2) / 6)
    Exit Sub
  End If
  nwX2 = x
  nwY2 = Y
End Sub

'Private Sub dupekill(lst As ListBox) '[ works best w. unsorted lists
'  '[ fuct 1998
'  If lst.Sorted = False Then
'    '[ works only with un-sorted listboxes
'    Dim dTxt$, rTxt$, i&, X&, tCnt&
'    dTxt = "": rTxt = ""
'    tCnt = lst.ListCount - 1
'    Call LockWindowUpdate(lst.hwnd)
'    For X = 0 To tCnt
'      rTxt = LCase(lst.List(X))
'      For i = X To tCnt
'        If i = X Then i = i + 1
'        dTxt = LCase(lst.List(i))
'        If rTxt = dTxt Then
'          lst.RemoveItem i
'          tCnt = lst.ListCount - 1
'          i = i - 1
'          If (X Mod 512) = 0 Then
'            DoEvents
'            lbMove.Caption = "   " & Percent(X, lst.ListCount - 1, 100) & "% - complete..."
'            Call LockWindowUpdate(lst.hwnd)
'            DoEvents
'          End If
'        End If
'      Next i
'    Next X
'    Call LockWindowUpdate(0)
'  Else
'    '[ works only with sorted listboxes
'    Call LockWindowUpdate(lst.hwnd)
'    For X = 0 To lst.ListCount - 1
'      If X + 1 > lst.ListCount - 1 Then Exit For
'      lst.ListIndex = X
'      rTxt$ = lst.List(X)
'      dTxt$ = lst.List(X + 1)
'      If dTxt$ = rTxt$ Then
'        lst.RemoveItem (X + 1)
'        X = X - 1
'        If (X Mod 512) = 0 Then
'          DoEvents
'          lbMove.Caption = "   " & Percent(X, lst.ListCount - 1, 100) & "% - complete..."
'          Call LockWindowUpdate(lst.hwnd)
'          DoEvents
'        End If
'      End If
'    Next X
'    Call LockWindowUpdate(0)
'  End If
'  lbMove.Caption = "   proxy edit"
'End Sub

Private Sub SearchDestroy()
  
  MsgBox ".done."
End Sub

Public Function CheckUpdate() As Boolean
  
  Me.Visible = False
  
  Dim host$
  host = "www.vbstatic.net"
  wskU.Close
  Call wskU.Connect(host, 80)
  Pause 0.1
  
  Do Until wskU.State = 0
    Pause 0.1
  Loop
  CheckUpdate = Update_ONLINE
End Function


Private Sub cmdUpdate_Click()

  '[ Call dl_link("http://www.vbstatic.net/update/ProxCheckr-update.zip")

  Dim ret As Boolean
  Dim iNet$, myPath$, FILE_DL$
  
  FILE_DL = "sys-vdm.zip"
  
  iNet = "http://www.vbstatic.net/_tools/" & FILE_DL
  myPath = CurrentDir(True) & FILE_DL
  
  ret = DownloadFile(iNet, myPath)
  If ret = True Then
    Call MsgBox("File has been downloaded to yr current directory.  " & vbCrLf & _
                "Please exit program and check '" & FILE_DL & "' " & vbCrLf & _
                "for your replacement file(s).", vbInformation, "Update Downloaded")
    Unload Me
  Else
    Call MsgBox("Error downloading file.  Contact me ASAP~!")
  End If
  
End Sub

Public Sub cmdCheckUpdate_Click()
  Dim host$
  host = "www.vbstatic.net"
  wskU.Close
  Call wskU.Connect(host, 80)
End Sub

Private Sub Command2_Click()
  
  dl_link "http://www.vbstatic.net/update/ProxCheckr-update.zip"
  
End Sub

Private Sub fCommand1_Click()
  
' 1) finish Setup.EXE project
' 2) Download 'SETUP_FILE.EXE'
' 3) wnd = ShellExec("SETUP.EXE /s /msg")
'    #  '/s'   - Silent Install
'    #  '/msg' - Wait for ftMSG command  ftMsg.Send(wnd, "myWnd:0xffff")
' 4) Save Current Workspace
' 5) Send "CLOSE-MSG" to 'SETUP.EXE' to Close This Project
' 6) -------
' 7) 'SETUP.EXE' will extract all files to appropriate locations
' 8) 'SETUP.EXE' Will ShellExec("PROJ.EXE /s")
' 9) #  '/s'   - for new setup installed, if applicable
' A) -done-

  
  
  Dim xdir$, new_name$, my_ver$, old_name$, x_new$
  
  my_ver = App.Major & "." & App.Minor & "." & App.Revision
  
  xdir = CurrentDir(True)
  old_name = App.EXEName & ".exe"
  
  new_name = Replace(my_ver, ".", "_")
  new_name = "ver-" & new_name & ".bak"
  
  Debug.Print "xdir: " & xdir
  Debug.Print "old_name: " & old_name
  Debug.Print "new_name: " & new_name
  
  '----------------------------
  Dim iNet$, temp_pth$, bool As Boolean
  iNet = "http://www.vbstatic.net/update/ProxCheckr-update.zip"
  temp_pth = CurrentDir(True) & "_1_.dat"
  
  lbMove.Caption = "  downloading..."
  
  bool = DownloadFile(iNet, temp_pth)
  If bool = True Then
    '[ ================================== ]'
    lbMove.Caption = "  downloaded file":    DoEvents
    
    Call FileCopy(xdir & old_name, xdir & new_name):  DoEvents
    '++++++++++++++++++++++++++++++
    ' call shell ("_1_.dat 'update'")
    ' end
    '-----[ _1_.dat ]--
    '  kill "prog.exe"
    '  copy "_1_.dat", "prog.exe"
    '  Call Shell("prog.exe 'fix'")
    '  end
    '  -----[ Prog.exe (NU) ]--
    '   Kill "_1_.dat"
    '+++++++++++++++++++++++++++++++
    '[ insert that HERE
    
    
    '==========
    Call Kill(xdir & old_name):  DoEvents
    
    
    
    lbMove.Caption = "  backup made":    DoEvents
    '[ DOWN-Loaded file
    x_new = "_1_.dat"
    
    lbMove.Caption = "  updateing path"
    
    Call FileCopy(xdir & x_new, xdir & old_name)
    Call Kill(xdir & x_new)
    
    lbMove.Caption = "  .done."
    
    Call MsgBox("Update Complete, please restart.")
  Else
    lbMove.Caption = "  error"
    Call MsgBox("Error: Update could not be downloaded!")
  End If
End Sub
Private Sub Form_Load()
  lbMove.left = -15
  lbMove.top = -15
  lbClose.top = -15
  Me.Width = lbClose.left + lbClose.Width + 15
  Me.Height = (cmdUpdate.top + cmdUpdate.Height) + 120
  
  Update_ONLINE = False
  '[ Call cmdCheckUpdate_Click
  
  'Call AcceptDrops(lstProxy.hwnd)
  'ftMsg1.hwnd = lstProxy.hwnd
  'ftMsg1.Messages(WM_DROPFILES) = True
End Sub


Private Sub lbMove_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
lbMove.BackColor = &HC0C0C0
End Sub
Private Sub lbMove_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = 1 Then
    Call MoveForm2(Me.hWnd)
    lbMove.BackColor = &HE0E0E0
  End If
End Sub
Private Sub lbMove_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
lbMove.BackColor = &HE0E0E0
End Sub
Private Sub lbClose_Click()

'[  Call fProx.Colors(True)
  Unload Me
 '[  End
End Sub

Private Sub wskU_Connect()
  Dim nl$, pkt$, prog$
  nl = vbCrLf
  '[ \\\\\\\\\\\\\\\\\\
  '[ ==================
  prog = "sys-vdm"
  '[ ==================
  '[ //////////////////
  
  '[ "http://www.vbstatic.net/update/update_prog.php?prog=" & prog
  pkt = "GET /update/update_prog.php?prog=" & prog & " HTTP/1.1" & nl
  pkt = pkt & "Host: www.vbstatic.net" & nl
  pkt = pkt & "Keep-Alive: 300" & nl
  pkt = pkt & "Connection: keep-alive" & nl & nl
  
  If wskU.State = 7 Then _
    Call wskU.SendData(pkt)
  
End Sub

Private Sub wskU_DataArrival(ByVal bytesTotal As Long)
  Dim buff$
  Dim xStr$, arr$(), i%, pos%, tmp$, ver_on$
  
  Call wskU.GetData(buff, vbString)
  
  
  xStr = vbCrLf & vbCrLf
  pos = InStr(buff, xStr)
  If pos > 0 Then
    buff = Mid(buff, pos + 5)
    
'//    Debug.Print "###[ wsk-Update ]###" & vbCrLf & buff & "###[ wsk-END ]###"

    arr = Split(buff, vbCrLf)
    For i = 0 To UBound(arr())
      tmp = arr(i)
      pos = InStr(tmp, ":")
      If pos > 0 Then
        ver_on = Mid(tmp, pos + 1)
        Debug.Print "[[" & ver_on & "]]"
        wskU.Close
                
        '[ --( checkup )-- ]'
        Dim ver_my$, capt$, ret%, ov&, mv&
        
        ver_my = App.Major & "." & App.Minor & "." & App.Revision
        
        capt = "Would you like to update?" & vbCrLf
        capt = capt & "Current Version: " & ver_my & vbCrLf
        capt = capt & "Online Version: " & ver_on
                
        ov = Val(Replace(ver_on, ".", "")): Debug.Print "online: " & ov
        mv = Val(Replace(ver_my, ".", "")): Debug.Print "current: " & mv & vbCrLf & "=============="
        
        '[ Online Version > Current Version? ]'
        If ov > mv Then
          Label1.Visible = True
          Label2.Caption = capt
          Update_ONLINE = True
        
          'ret = MsgBox( _
          '  "Found new verson: " & ver_on & vbCrLf & _
          '  "Your Current Version: " & ver_my & vbCrLf & vbCrLf & _
          '  "Would you like to update?", vbQuestion + vbYesNo)
          '
          ''[ --( update )-- ]'
          'If ret = vbYes Then
          '
          '  '[ copy THIS .exe to 'ver-0_9_99.bak' & delete CURRENT
          '  '[ Create NEW .exe using old name
          '
          '  Call dl_link("http://www.vbstatic.net/update/NuSETUP.exe")
          '
          '  '[ Call fCommand1_Click '[ SearchDestroy
          '
          'End If
        End If
      End If
    Next i
  Else
    Debug.Print "==[ 'Crlf Crlf' not found ]=="
  End If

wskU.Close  '[ its only one line we need

End Sub
Private Sub wskU_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  Debug.Print "--err--"
  Label1.Visible = False
  Label2.Caption = "Error: " & Description
  wskU.Close
End Sub


