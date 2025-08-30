VERSION 5.00
Begin VB.Form fSetup 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Virtual Desktop Manager - [VDM] Setup"
   ClientHeight    =   5970
   ClientLeft      =   6690
   ClientTop       =   2655
   ClientWidth     =   7815
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
   Icon            =   "fSetup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   Begin VB.Frame windo 
      Caption         =   "Sticky Windows  [dbl-click to edit]"
      Height          =   2895
      Index           =   2
      Left            =   3990
      TabIndex        =   12
      Top             =   2940
      Width           =   3795
      Begin VB.CommandButton cmdSaveSticky 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2970
         TabIndex        =   36
         Top             =   510
         Width           =   735
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Reload"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2970
         TabIndex        =   35
         ToolTipText     =   "Refresh Sticky-Window List"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdEditSticky 
         Caption         =   "Edit Item"
         Height          =   255
         Left            =   2850
         TabIndex        =   34
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton cmdRmvSticky 
         Caption         =   "Remove"
         Height          =   255
         Left            =   2100
         TabIndex        =   15
         Top             =   840
         Width           =   735
      End
      Begin VB.PictureBox picWnd 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   120
         MouseIcon       =   "fSetup.frx":000C
         Picture         =   "fSetup.frx":08D6
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   18
         Top             =   390
         Width           =   240
      End
      Begin VB.ListBox lstSticky 
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
         Height          =   1665
         IntegralHeight  =   0   'False
         ItemData        =   "fSetup.frx":0E60
         Left            =   90
         List            =   "fSetup.frx":0E62
         TabIndex        =   17
         Top             =   1140
         Width           =   3615
      End
      Begin VB.CommandButton cmdAddSticky 
         Caption         =   "Add Window"
         Height          =   255
         Left            =   990
         TabIndex        =   16
         Top             =   840
         Width           =   1125
      End
      Begin VB.TextBox txtClass 
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
         Height          =   255
         Left            =   810
         TabIndex        =   14
         Text            =   "[none]"
         Top             =   240
         Width           =   2115
      End
      Begin VB.TextBox txtTitle 
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
         Height          =   255
         Left            =   810
         TabIndex        =   13
         Text            =   "[none]"
         Top             =   510
         Width           =   2115
      End
      Begin VB.CommandButton cmdHelpSticky 
         Caption         =   "¿ Help ?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   90
         TabIndex        =   30
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "class:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   1
         Left            =   420
         TabIndex        =   20
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "title:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   0
         Left            =   450
         TabIndex        =   19
         Top             =   570
         Width           =   255
      End
   End
   Begin VB.Frame windo 
      Caption         =   "Backgrounds"
      Height          =   2895
      Index           =   1
      Left            =   60
      TabIndex        =   7
      Top             =   2940
      Width           =   3795
      Begin VB.CheckBox chkBGColr 
         Height          =   225
         Left            =   1860
         TabIndex        =   51
         ToolTipText     =   "Check if you want to modify background color."
         Top             =   2070
         Width           =   195
      End
      Begin VB.PictureBox picBGColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1680
         ScaleHeight     =   285
         ScaleWidth      =   525
         TabIndex        =   49
         Top             =   1740
         Width           =   555
      End
      Begin VB.PictureBox picEdit 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3450
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   48
         Top             =   1770
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.HScrollBar hsDeskEdit 
         Height          =   225
         Left            =   2130
         Max             =   10
         Min             =   1
         TabIndex        =   45
         Top             =   990
         Value           =   1
         Width           =   1575
      End
      Begin VB.PictureBox picPreview 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   945
         Left            =   2310
         ScaleHeight     =   915
         ScaleWidth      =   1215
         TabIndex        =   44
         Top             =   1410
         Width           =   1245
      End
      Begin VB.Frame FrameWall 
         Appearance      =   0  'Flat
         Caption         =   "Image Style"
         ForeColor       =   &H80000008&
         Height          =   1035
         Left            =   240
         TabIndex        =   39
         Top             =   1320
         Width           =   1365
         Begin VB.OptionButton optStyle 
            Caption         =   "Tiled"
            Height          =   195
            Index           =   2
            Left            =   90
            TabIndex        =   42
            ToolTipText     =   "Tiled"
            Top             =   780
            Width           =   795
         End
         Begin VB.OptionButton optStyle 
            Caption         =   "Center"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   41
            ToolTipText     =   "Center"
            Top             =   510
            Value           =   -1  'True
            Width           =   945
         End
         Begin VB.OptionButton optStyle 
            Caption         =   "Stretch"
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   40
            ToolTipText     =   "Stretch"
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdFindWall 
         Caption         =   "..."
         Height          =   285
         Left            =   3390
         TabIndex        =   38
         Top             =   2460
         Width           =   345
      End
      Begin VB.TextBox txtWallPath 
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
         Height          =   285
         Left            =   150
         TabIndex        =   37
         Text            =   "(path to wallpaper)"
         Top             =   2460
         Width           =   3225
      End
      Begin VB.HScrollBar hsDesk 
         Height          =   225
         Left            =   2490
         Max             =   10
         Min             =   2
         TabIndex        =   10
         Top             =   570
         Value           =   4
         Width           =   1215
      End
      Begin VB.CheckBox chkBk 
         Appearance      =   0  'Flat
         Caption         =   "Custom Backgrounds"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   150
         TabIndex        =   9
         Top             =   270
         Width           =   1905
      End
      Begin VB.CommandButton cmdWallApply 
         Caption         =   "Apply Changes To:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   8
         Top             =   960
         Width           =   1665
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "[2-10]"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   165
         Index           =   8
         Left            =   1710
         TabIndex        =   52
         Top             =   570
         Width           =   330
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   150
         X2              =   3660
         Y1              =   855
         Y2              =   855
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "BG Color:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   7
         Left            =   1650
         TabIndex        =   50
         Top             =   1530
         Width           =   615
      End
      Begin VB.Label lbl 
         Caption         =   "**Use '(none)' for no background image"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   6
         Left            =   2280
         TabIndex        =   47
         Top             =   150
         Width           =   1305
      End
      Begin VB.Label lblWallEdit 
         Alignment       =   2  'Center
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1860
         TabIndex        =   46
         Top             =   990
         Width           =   225
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000015&
         Index           =   0
         X1              =   120
         X2              =   3690
         Y1              =   870
         Y2              =   870
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Number of Desktops:"
         Height          =   195
         Index           =   5
         Left            =   150
         TabIndex        =   43
         Top             =   570
         Width           =   1515
      End
      Begin VB.Label lblDesks 
         Alignment       =   2  'Center
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2220
         TabIndex        =   11
         Top             =   570
         Width           =   225
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   180
      TabIndex        =   2
      Top             =   2520
      Width           =   1425
   End
   Begin VB.Frame windo 
      Caption         =   "Settings"
      Height          =   2895
      Index           =   0
      Left            =   1740
      TabIndex        =   1
      Top             =   0
      Width           =   3795
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "check"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2250
         TabIndex        =   31
         Top             =   450
         Width           =   705
      End
      Begin VB.HScrollBar hsTrans 
         Height          =   195
         LargeChange     =   5
         Left            =   180
         Max             =   255
         TabIndex        =   28
         Top             =   1320
         Value           =   255
         Width           =   2025
      End
      Begin VB.CheckBox chkExit 
         Appearance      =   0  'Flat
         Caption         =   "Confirmation on exit"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   180
         TabIndex        =   27
         Top             =   990
         Value           =   1  'Checked
         Width           =   1995
      End
      Begin VB.CheckBox chkStartup 
         Appearance      =   0  'Flat
         Caption         =   "automatic startup with windows"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   180
         TabIndex        =   26
         Top             =   720
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin VB.CheckBox chkUpdate 
         Appearance      =   0  'Flat
         Caption         =   "automatic self-update"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   180
         TabIndex        =   23
         Top             =   450
         Value           =   1  'Checked
         Width           =   2085
      End
      Begin VB.PictureBox picIco 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   3210
         Picture         =   "fSetup.frx":0E64
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   22
         Top             =   120
         Width           =   480
      End
      Begin VB.TextBox txtRefresh 
         Alignment       =   2  'Center
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
         Height          =   255
         Left            =   180
         TabIndex        =   21
         Text            =   "750"
         ToolTipText     =   "VWM Timer refresh rate [ def=750 "
         Top             =   1620
         Width           =   495
      End
      Begin VB.Frame frame 
         Appearance      =   0  'Flat
         Caption         =   "z-order setting"
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   0
         Left            =   90
         TabIndex        =   3
         Top             =   1950
         Width           =   3615
         Begin VB.OptionButton optZ 
            Appearance      =   0  'Flat
            Caption         =   "force on top"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   6
            Top             =   270
            Value           =   -1  'True
            Width           =   1245
         End
         Begin VB.OptionButton optZ 
            Appearance      =   0  'Flat
            Caption         =   "force on desktop"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   1080
            TabIndex        =   5
            Top             =   540
            Width           =   1635
         End
         Begin VB.OptionButton optZ 
            Appearance      =   0  'Flat
            Caption         =   "no z-order"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   2310
            TabIndex        =   4
            Top             =   270
            Width           =   1125
         End
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "99 < ## <  32767"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   4
         Left            =   2610
         TabIndex        =   33
         Top             =   1740
         Width           =   1005
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "note:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   3
         Left            =   2610
         TabIndex        =   32
         Top             =   1560
         Width           =   315
      End
      Begin VB.Label lblTranz 
         AutoSize        =   -1  'True
         Caption         =   "Alpha Level: 255"
         Height          =   195
         Left            =   2280
         TabIndex        =   29
         Top             =   1320
         Width           =   1200
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "VWM Refresh rate (ms)"
         Height          =   195
         Index           =   2
         Left            =   780
         TabIndex        =   25
         ToolTipText     =   "VWM Timer refresh rate [ def=750 "
         Top             =   1650
         Width           =   1680
      End
      Begin VB.Label lblCapt 
         AutoSize        =   -1  'True
         Caption         =   "sys-vwm v0.0.0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   780
         TabIndex        =   24
         Top             =   210
         Width           =   1530
      End
   End
   Begin VB.ListBox lst 
      Height          =   2400
      ItemData        =   "fSetup.frx":172E
      Left            =   60
      List            =   "fSetup.frx":173B
      TabIndex        =   0
      Top             =   60
      Width           =   1635
   End
   Begin VB.Menu mnuSticky 
      Caption         =   "sticky"
      Begin VB.Menu cmdSticky_add 
         Caption         =   "Add New Window"
      End
      Begin VB.Menu cmdSticky_remove 
         Caption         =   "Remove Current Window"
      End
      Begin VB.Menu cmdSticky_edit 
         Caption         =   "Edit Item"
      End
      Begin VB.Menu cmdSticky_split 
         Caption         =   "-"
      End
      Begin VB.Menu cmdSticky_pre 
         Caption         =   "Add Preset"
         Begin VB.Menu cmdSticky_pre_winampMain 
            Caption         =   "Winamp (Main Window)"
         End
         Begin VB.Menu cmdSticky_pre_winampAll 
            Caption         =   "Winamp (All Windows)"
         End
         Begin VB.Menu cmdSticky_pre_taskmngr 
            Caption         =   "Task Manager"
         End
         Begin VB.Menu cmdSticky_pre_reg 
            Caption         =   "Regedit"
         End
         Begin VB.Menu cmdSticky_pre_spypp 
            Caption         =   "Microsoft Spy++**"
         End
         Begin VB.Menu cmdSticky_pre_cpp 
            Caption         =   "**Microsoft Visual C++**"
         End
         Begin VB.Menu cmdSticky_pre_smartftp 
            Caption         =   "SmartFTP v**"
         End
      End
   End
End
Attribute VB_Name = "fSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PictureThis As StdPicture
'Private imgEdit As clsImage

Private Type POINTAPI
  x As Long
  Y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long


Private Sub LoadImage(path$, SavePreview As Boolean, Optional SaveTo As String = "")
  
  Dim imgEdit As New clsImage
  Call imgEdit.ReadImageInfo(path)
  Call picEdit.PaintPicture(LoadPicture(path), 0, 0, imgEdit.Width, imgEdit.Height)
  
  If SavePreview = True Then
    Set picEdit.Picture = LoadPicture(path) '[ one more time.. could be removed
    Call SavePicture(picEdit.Picture, SaveTo)
  End If
  
  Call picPreview.PaintPicture(LoadPicture(path), _
      0, 0, picPreview.ScaleWidth, picPreview.ScaleHeight, 0, 0)

End Sub

Private Sub LoadOptions()
  Dim intg As Integer, byt As Byte
  
  '[ ===( MAIN )=== ]'
  byt = CByte(readini("opt", "trans", "180"))
  '[ lblTranz.Caption = Percent((255 - CLng(byt)), 255, 100) & "% Transparency"
  lblTranz.Caption = "Alpha Level: " & CLng(byt)
  hsTrans.Value = byt
  intg = CInt(readini("opt", "zorder", "0"))
    optZ(intg).Value = True
  intg = CInt(readini("opt", "refresh", "750"))
    txtRefresh.Text = intg
    txtRefresh.ForeColor = &H80000008
  intg = CInt(readini("opt", "update", "1"))
    chkUpdate.Value = intg
  intg = CInt(readini("opt", "startup", "0"))
    chkStartup.Value = intg
  intg = CInt(readini("desks", "custom", "0"))
    chkBk.Value = intg
  
  Call RefreshSticky
  Call LoadWallpaper(1)
  
End Sub
Private Sub LoadWallpaper(ndx%)

On Error GoTo err_hand

  Dim buff$, path$, x_ndx%, ret$
  
  buff = readini("desks", "desk" & ndx)
  If InStr(buff, ";") <> 0 Then
    
    path = Trim(Split(buff, ";")(2))
    x_ndx = CInt(Split(buff, ";")(0)) - 1
    ret = Trim(Split(buff, ";")(1))
    If (IsNumeric(ret) = True) Then
      If (CLng(ret) > -1) Then
        picBGColor.BackColor = CLng(ret)
        chkBGColr.Value = 1
      Else
        picBGColor.BackColor = GetSysColor(COLOR_BACKGROUND)
        chkBGColr.Value = 0
      End If
    End If
    

    If FileExists(path) = False Then
      Call MsgBox("Invalid Image Path on Desktop " & ndx, vbCritical)
    Else
      optStyle(x_ndx).Value = True
      
      txtWallPath.Text = path
      Call LoadImage(path, False)
      
    End If
    
  End If
Exit Sub
err_hand:
  Call MsgBox("Error: " & Err & vbCrLf & Error(Err), vbCritical, "LoadWallpaper")
End Sub


Private Sub RefreshSticky()

  lstSticky.Clear
  Call load_sticky
  
  '[ ===( STICKY )=== ]'
  If I_Sticky >= 1 Then
    Dim x_cls$, x_txt$, ndx&, tmp_c$, tmp_t$
    
    For ndx = 0 To I_Sticky - 1
      tmp_c = "": tmp_t = ""
      
      x_cls = A_Sticky(ndx).sClass
      If Len(x_cls) > 0 Then
        tmp_c = "[cls]" & x_cls & "[/cls]"
      End If
      
      x_txt = A_Sticky(ndx).sTitle
      If Len(x_txt) > 0 Then
        tmp_t = "[txt]" & x_txt & "[/txt]"
      End If
      
      Call lstSticky.AddItem(tmp_c & tmp_t)
      
    Next ndx
  End If

End Sub


Private Sub chkBk_Click()
  Dim cust%
  If chkBk.Value = 0 Then cust = 0 Else cust = 1
  Call writeini("desks", "custom", cust)
End Sub

Private Sub chkExit_Click()
  Dim ini$
  If chkExit.Value = 1 Then
        ini = "1"
  Else: ini = "0"
  End If
  
  Call writeini("opt", "confirm", ini)
End Sub

Private Sub chkStartup_Click()
  Dim ini$
   Dim cReg As New clsRegistry, ke$, pth$, var$
  
    ke = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
    var = "sys-vdm"
  If chkStartup.Value = 1 Then
    ini = "1"
    Call writeini("opt", "startup", ini)
    
    If InIDE() = False Then
      pth = CurDirFull()
      Call cReg.SetValue(eHKEY_LOCAL_MACHINE, ke, var, pth)
    End If
    
  Else
    ini = "0"
    Call writeini("opt", "startup", ini)
    
    If cReg.KeyExists(eHKEY_LOCAL_MACHINE, ke) = True Then
      Call cReg.DeleteValue(eHKEY_LOCAL_MACHINE, ke, var)
    End If
    
  End If
  
  
End Sub

Private Sub chkUpdate_Click()
  Dim ini$
  If chkUpdate.Value = 1 Then
        ini = "1"
  Else: ini = "0"
  End If
  
  Call writeini("opt", "update", ini)
End Sub


Private Sub cmdAddSticky_Click()
  Dim s_cls$, s_txt$
  
  If comp(txtClass, "[none]") Then txtClass = ""
  If comp(txtTitle, "[none]") Then txtTitle = ""
  If txtClass = "" And txtTitle = "" Then Exit Sub
  
  cmdSaveSticky.Enabled = True  '[ changes may of been made~!
  
  s_cls = CLASS_START & txtClass & CLASS_END
  s_txt = TITLE_START & txtTitle & TITLE_END
  
  Call lstSticky.AddItem(s_cls & s_txt)
    
End Sub

Private Sub cmdWallApply_Click()
  Dim xdesk$, xval$
  Dim img_path$, img_pathOG$, img_pos As wall_style
  Dim cur_desk%, pth_save$
  Dim bg_color&
  
  cur_desk = hsDeskEdit.Value
  img_path = txtWallPath.Text
  img_pathOG = txtWallPath.Text
  
  If comp(img_path, "(none)") Then
'//    bg_color = picBGColor.BackColor
    xval = "2"      '[ just for the hell of it
    GoTo skip_all
  ElseIf FileExists(img_path) = False Then
    '[ insert handler for "(none)"
    Call MsgBox("Invalid Wallpaper Path!", vbCritical)
    Exit Sub
  End If

'[ ###[ Save as BitMap ]###################
  pth_save = CurrentDir(True) & "_skin\"
  If DirExists(pth_save) = False Then Call MkDir(pth_save)
  img_path = pth_save & "desk-" & cur_desk & ".bmp"
  
  Call LoadImage(img_pathOG, True, img_path)
  
  'Call picEdit.PaintPicture(LoadPicture(img_path), 0, 0, imgEdit.Width, imgEdit.Height, 0, 0)
  'Call SavePicture(picEdit.Picture, img_path)
  ' Set picEdit.Picture = LoadPicture(img_path) '[ one more time.. could be removed
  
'[ ############################

  xdesk = "desk" & cur_desk
  
  If optStyle(0).Value = True Then      '[ stretched
    
    img_pos = StretchWall
    '[ bg_color = "-1" '[ set to -1 cause there is no bg-color
    xval = "1"
    
  ElseIf optStyle(1).Value = True Then  '[ center
    
    img_pos = CenterWall
    '[ bg_color = picBGColor.BackColor
    xval = "2"
    
  ElseIf optStyle(2).Value = True Then  '[ tiled
    
    img_pos = TileWall
    '[ bg_color = "-1"   '[ set to -1 cause there is no bg-color
    bg_color = picBGColor.BackColor
    xval = "3"
    
  Else
    Call MsgBox("Please select an Image Style!", vbCritical)
    Exit Sub
  End If
  
skip_all:
  
  '[ modify the backgound color.. set here for non-2k+ (nt)users
  If chkBGColr.Value = 1 Then bg_color = picBGColor.BackColor Else bg_color = -1
  
  xval = xval & ";" & bg_color & ";" & img_path
  
  Call writeini("desks", xdesk, xval)
  
  If chkBk.Value = 1 Then
      
    Debug.Print "Changed Wallpaper[" & SYS.CurrentDesk & "]"
    Debug.Print "sys: " & SYS.CurrentDesk
    Debug.Print "hs: " & hsDeskEdit.Value
    
    If (hsDeskEdit.Value) = (SYS.CurrentDesk + 1) Then
      Debug.Print "Changed Wallpaper[" & SYS.CurrentDesk & "]"
      
      Call ChangeWallpaper(img_path, img_pos, bg_color)
    End If
  End If
End Sub

Public Function ActiveDesktop() As Boolean

    Dim tmpLong&
    tmpLong = FindWindow("Progman", vbNullString)
    tmpLong = FindWindowEx(tmpLong, 0&, "SHELLDLL_DefView", vbNullString)
    tmpLong = FindWindowEx(tmpLong, 0&, "Internet Explorer_Server", _
    vbNullString)


    If tmpLong > 0 Then
        ActiveDesktop = True
    Else
        ActiveDesktop = False
    End If

End Function

Private Sub cmdClose_Click()
  If cmdSaveSticky.Enabled = True Then
    Dim ret%
    ret = MsgBox("Changes were made to Sticky Window list, do you wish to continue " & _
                 "with out saving?", vbYesNo + vbQuestion, "Changes were made")
    If ret = vbNo Then Exit Sub
  End If
  
  Unload fSetup
End Sub

Private Sub cmdEditSticky_Click()
  Dim ndx%, sMsg$, def$, ret$
  
  ndx = lstSticky.ListIndex
  If (ndx > -1) Then
    cmdSaveSticky.Enabled = True  '[ changes may of been made~!
    
    '
    '3,taskmgr.exe                      // EXE NAME
    '9,MusicMatch Jukebox
    '9,*- Receiving mail
    '9,*- Delivering mail
    '9,Microsoft Spy++*                 // USE OF WILD CARDS
    '
    
    def = lstSticky.List(ndx)
    sMsg = sMsg & "To edit hWnd to be Sticky Remember:" & vbCrLf
    sMsg = sMsg & "" & vbCrLf
    sMsg = sMsg & "    [cls] CLASS_NAME [/cls]" & vbCrLf
    sMsg = sMsg & "    [txt] CLASS_TITLE [/txt]" & vbCrLf
    sMsg = sMsg & "" & vbCrLf
    sMsg = sMsg & "If you dont have a class or caption leave it "
    sMsg = sMsg & "blank w/o spaces in the middle." & vbCrLf
    
    ret = Trim(InputBox(sMsg, "Edit Sticky Window", def))
    If ret <> "" Then
      lstSticky.List(ndx) = ret
    End If
  End If
End Sub

Private Sub cmdFindWall_Click()
  Dim filt$, xflg&, titl$, path$
  
  titl = "Select Wallpaper for Desktop " & hsDeskEdit.Value
  filt = "Common Imgaes|*.jpg;*.gif;*.png;*.bmp|All Files|*.*|"
  xflg = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST Or OFN_OVERWRITEPROMPT Or OFN_ENABLESIZING Or OFN_EXPLORER Or OFN_LONGNAMES
  
  path = openCMD(Me, filt, titl, "_skin", xflg)
  
  If path = "" Then Exit Sub
  txtWallPath = path
  
  Call LoadImage(path, False)
  
End Sub

Private Sub cmdHelpSticky_Click()
  Dim mx$
  Call ontop(Me.hWnd, ZPOS_Normal)
  mx = mx & "Sticky Window Help:" & vbCrLf & vbCrLf

  mx = mx & "Click & drag the cross-hair to select your window that you wish to" & vbCrLf
  mx = mx & "remain on all Desktops ('sticky').  To make that Window sticky just" & vbCrLf
  mx = mx & "click 'Add Window' and a simple script will be added." & vbCrLf
  mx = mx & "" & vbCrLf
  mx = mx & "If during switch of desktops there are a few Toolbars\Windows left" & vbCrLf
  mx = mx & "behind you will need to add those items by hand (for now)." & vbCrLf
  mx = mx & "" & vbCrLf
  mx = mx & "To 'Edit' script be very careful to STICK TO THE RULES, you can" & vbCrLf
  mx = mx & "cause a program to malfunction if you edit this improperly" & vbCrLf
  mx = mx & "" & vbCrLf
  mx = mx & "Current Editable Tags:" & vbCrLf
  mx = mx & "    [cls] CLASS_NAME [/cls]" & vbCrLf
  mx = mx & "    [txt] CLASS_TITLE [/txt]" & vbCrLf
  'mx = mx & "    [exe] EXE_NAME [/exe]" & vbCrLf
  mx = mx & "" & vbCrLf
  mx = mx & "For any extra documentation goto:   http://www.vbStatic.net/"

  
  Call MsgBox(mx, vbInformation, "Sticky Window Help")
  Call ontop(Me.hWnd, ZPOS_TopMost)
End Sub


Private Sub cmdRefresh_Click()
  Call RefreshSticky
  cmdSaveSticky.Enabled = False
End Sub

Private Sub cmdRmvSticky_Click()
  Dim ndx%
  ndx = lstSticky.ListIndex
  If ndx = -1 Then Exit Sub
  
  Dim tmp$, ret%
  tmp = lstSticky.List(ndx)
  
  ret = MsgBox("Are you sure you want to remove Window with values:" & vbCrLf & _
                tmp & vbCrLf & "From your Sticky Windows list?", vbYesNo + vbQuestion, _
                "Removing Sticky Window?")
                
  If ret = vbYes Then
    cmdSaveSticky.Enabled = True  '[ changes may of been made~!
    Call lstSticky.RemoveItem(ndx)
  End If
  
End Sub

Private Sub cmdSaveSticky_Click()
  '[ Add Listbox items to a_sticky() & i_sticky& GLOBALS
  Dim cnt%
  cnt = lstSticky.ListCount
  I_Sticky = cnt
  
  If cnt > 0 Then
    ReDim A_Sticky(I_Sticky - 1)
    Dim ndx%, tmp$, s_cls$, s_txt$
    
    For ndx = 0 To I_Sticky - 1
      tmp = lstSticky.List(ndx)
      
      s_cls = "": s_txt = ""
      
      s_cls = ReadVar(tmp, "cls")
      If Len(s_cls) > 0 Then
            A_Sticky(ndx).sClass = s_cls
      Else: A_Sticky(ndx).sClass = vbNullString
      End If
      
      s_txt = ReadVar(tmp, "txt")
      If Len(s_txt) > 0 Then
            A_Sticky(ndx).sTitle = s_txt
      Else: A_Sticky(ndx).sTitle = vbNullString
      End If
      
    Next ndx
    
    Call save_sticky
    cmdSaveSticky.Enabled = False
    
  Else
    
    Call ScriptErase("sticky")
    cmdSaveSticky.Enabled = False
  End If
  
End Sub
Private Sub cmdUpdate_Click()
  Dim bool As Boolean
  bool = fUpdate.CheckUpdate
  If bool = True Then
    fUpdate.Visible = True
    Call ontop(fUpdate.hWnd, ZPOS_TopMost)
    Call fUpdate.ZOrder(0)
  Else
    Call MsgBox("You have the most updated version according to 'vbStatic.net'.", vbInformation)
    Unload fUpdate
  End If
End Sub

Private Sub DeskToolbar1_OpenAnimateStarts()

End Sub

Private Sub Form_Load()
  mnuSticky.Visible = False
  mnuSticky.Enabled = False
  
  Dim i%
  For i = 0 To 2
    windo(i).top = 0
    windo(i).left = 1740
  Next i
  
  Call windo(0).ZOrder(0)
  
  Call LoadOptions
  
  

  
  Me.Height = windo(0).Height + (15 * 25)
  Me.Width = windo(0).left + windo(0).Width + (15 * 10)
  lblCapt.Caption = "sys-vdm v" & _
                    App.Major & "." & App.Minor & "." & App.Revision
  lblDesks.Caption = SYS.VDM_Count
  hsDesk.Value = SYS.VDM_Count
  '[ Call ontop(fSetup.hWnd, ZPOS_TopMost)
  
  cmdSaveSticky.Enabled = False
  Me.Show
  'Call sizeListbox(lstBK)
  
End Sub

Private Sub hsDesk_Change()
lblDesks = hsDesk.Value
End Sub

Private Sub hsDesk_Scroll()
lblDesks = hsDesk.Value
End Sub


Private Sub hsDeskEdit_Change()
  lblWallEdit = hsDeskEdit.Value
  Call LoadWallpaper(hsDeskEdit.Value)
  
End Sub

Private Sub hsDeskEdit_Scroll()
lblWallEdit = hsDeskEdit.Value
End Sub


Private Sub hsTrans_Change()
  Dim ini$, var As Byte
  
  var = hsTrans.Value
  ini = Trim(CStr(var))
  
  '[ lblTranz.Caption = Percent((255 - CLng(var)), 255, 100) & "% Transparency"
  lblTranz.Caption = "Alpha Level: " & CLng(var)
  '
  '[ Error, Will stay as not layered when switched off.
  '[
  '[ If var = 255 Then
  '[       Call SetLayered(fMain.hWnd, False)
  '[ Else: Call SetLayered(fMain.hWnd, True, var)
  '[ End If
  Call SetLayered(fMain.hWnd, True, var)
  
  Call writeini("opt", "trans", ini)
End Sub

Private Sub hsTrans_Scroll()
  Call hsTrans_Change
End Sub


Private Sub lst_Click()
  Dim i%
  i = lst.ListIndex
  If i > -1 Then
    Call windo(i).ZOrder(0)
  End If
End Sub


'Private Sub lstBK_Click()
'Call sizeListbox(lstBK)
'End Sub

Private Sub lstSticky_DblClick()
  Call cmdEditSticky_Click
End Sub


Private Sub optZ_Click(Index As Integer)
  Dim ini$
  Select Case Index
    Case 0: Debug.Print "Form ON-TOP"  '[ ontop
      Call ontop(fMain.hWnd, ZPOS_TopMost)
    Case 1: Debug.Print "Form ON-BOTTOM"  '[ on desktop
      Call ontop(fMain.hWnd, ZPOS_Desktop)
    Case 2: Debug.Print "Form none" '[ no zorder
      Call ontop(fMain.hWnd, ZPOS_Normal)
  End Select
  
  ini = Trim(CStr(Index))
  Call writeini("opt", "zorder", ini)
  
End Sub

Private Sub picBGColor_Click()
  Dim sc As SelectedColor
  
  sc = ShowColor(Me.hWnd)
  If sc.bCanceled = False Then
    picBGColor.BackColor = sc.colr
  End If
  
End Sub

Private Sub picWnd_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Set PictureThis = picWnd.Picture '[ save the picture of the ring in the PictureThis variable
  picWnd.MousePointer = 99 '[ set the cursor equal to the ring picture
  picWnd.Picture = Me.Picture '[ set picture1.picture equal to nothing
End Sub


Private Sub picWnd_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Dim hWnd&, ttx$, CPos As POINTAPI
  Static OldWindow&
  
  If Button = 1 Then
  
    Call GetCursorPos(CPos)
    hWnd = WindowFromPoint(CPos.x, CPos.Y)
    '[ hWnd = GetParent(hWnd)
    If hWnd <> OldWindow Then
      txtClass.Text = GetClass(hWnd)
      ttx = GrabWndText(hWnd)
      If ttx <> "" Then
        txtTitle.Text = ttx
      Else
        txtTitle.Text = "" '// "[none]"
      End If

    End If
  End If
End Sub


Private Sub picWnd_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  picWnd.MousePointer = 0           '[ set the cursor equal to its default
  Set picWnd.Picture = PictureThis  '[ set picture1.picture to the picture of the ring
  'Call Command1_Click                 '[ Click the Command1 button
End Sub


Private Sub txtRefresh_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 8, 13
      If KeyAscii = 13 Then
        Dim ms%, ini$
        
        KeyAscii = 0
        'txtRefresh.ForeColor = &H0
        ms = CInt(Trim(txtRefresh.Text))
        If (ms > &H7FFF) Or (ms < 99) Then  '[ 32767 Then
          txtRefresh.ForeColor = vbRed
        Else
          
          
          fMain.tmrVDM.Interval = ms
          ini = Trim(CStr(ms))
          Call writeini("opt", "refresh", ini)
          txtRefresh.ForeColor = &H80000008
        End If
      Else
        txtRefresh.ForeColor = vbRed
      End If
      
    Case Else
      KeyAscii = 0
  End Select
End Sub


