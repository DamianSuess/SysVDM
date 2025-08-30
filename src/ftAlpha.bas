Attribute VB_Name = "ftAlpha"
Option Explicit
'[
'[ i am special i'm usin fuct's shit... ooohh!
'[ ...
'[    (and now a special msg from fuct himself)
'[
'[ i am jacks internal rage
'[ i am jacks total disgust
'[ i am jacks chaped ass..
'[
'[  ~ fuct

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hWnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long
Private Type POINTAPI
  x As Long
  Y As Long
End Type
Private Type SIZE
  cx As Long
  cy As Long
End Type
Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const AC_SRC_OVER = &H0
Private Const AC_SRC_ALPHA = &H1
Private Const AC_SRC_NO_PREMULT_ALPHA = &H1
Private Const AC_SRC_NO_ALPHA = &H2
Private Const AC_DST_NO_PREMULT_ALPHA = &H10
Private Const AC_DST_NO_ALPHA = &H20
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private lret As Long
Public Sub fadeout(frm As Form)
    '[ --- fade-out --- ]'
  If IsWin2000Plus() = True Then
    Dim cnt&
    
    If readini("opt", "shade") = "1" Then cnt = 220 Else cnt = 255
    
    Call SetLayered(frm.hWnd, True, cnt)
    frm.Visible = True
     
    Do: Pause 0.01
      Call SetLayered(frm.hWnd, True, cnt)
     cnt = cnt - 5.5
    Loop Until cnt <= 0
    Call SetLayered(frm.hWnd, True, 0)
  Else
    frm.Visible = False
  End If
End Sub
Public Sub fadein(frm As Form)
    '[ --- fade-in --- ]'
  If IsWin2000Plus() = True Then
    Dim i&, Max&, t&
    i = 0
    
    If readini("opt", "shade") = "1" Then Max = 220 Else Max = 255
    
    Call SetLayered(frm.hWnd, True, 0)
    frm.Visible = True
    
    Do: Pause 0.01
      Call SetLayered(frm.hWnd, True, i)
      i = i + 5.5
    Loop Until i >= Max
    Call SetLayered(frm.hWnd, True, Max)
  Else
    frm.Visible = True
  End If
End Sub
Public Function CheckLayered(ByVal hWnd As Long) As Boolean

  If IsWin2000Plus() = True Then
  
    lret = GetWindowLong(hWnd, GWL_EXSTYLE)
    If (lret And WS_EX_LAYERED) = WS_EX_LAYERED Then
      CheckLayered = True
    Else
      CheckLayered = False
    End If
  Else
    CheckLayered = False
  End If
End Function
Public Sub SetLayered(ByVal hWnd As Long, SetAs As Boolean, Optional ByVal bAlpha As Byte = 220)

  If IsWin2000Plus() = True Then
  
    lret = GetWindowLong(hWnd, GWL_EXSTYLE)
    If SetAs = True Then
      lret = lret Or WS_EX_LAYERED
    Else: lret = lret And Not WS_EX_LAYERED
    End If
    Call SetWindowLong(hWnd, GWL_EXSTYLE, lret)
    Call SetLayeredWindowAttributes(hWnd, 0, bAlpha, LWA_ALPHA)
    
  End If
End Sub
