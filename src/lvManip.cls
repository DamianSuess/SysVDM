VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLVManip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : CListViewManipulator
' DateTime  : 01 January 2003 12:00
' Author    : Nicholas Adie
' Company   : Brokenmould Limited
' Web       : www.nickadie.com & www.brokenmould.com
' Purpose   : ListView Manipulator
'---------------------------------------------------------------------------------------
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const LVM_FIRST = &H1000
Private Const LVM_GETTOPINDEX As Long = (LVM_FIRST + 39)
Private Const LVM_GETCOUNTPERPAGE = (LVM_FIRST + 40)
Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Private Const LVSCW_AUTOSIZE As Long = -1
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

' Object Definition
Private lvwWork                         As MSComctlLib.ListView

Public Sub Initialise(ByRef lvwSource As MSComctlLib.ListView)

  '[ Initialise Internal ListView
  Set lvwWork = lvwSource
  
  '[ Reset Sorted Key
  lvwWork.Sorted = False

End Sub
Public Property Get VisibleItemCount() As Long

  '[ Return Error if not Initialised
  If lvwWork Is Nothing Then Err.Raise 445
  
  '[ Return the Visible Item Count
  VisibleItemCount = SendMessage(lvwWork.Hwnd, LVM_GETCOUNTPERPAGE, 0&, ByVal 0&)
       
End Property
Public Property Get VisibleItemTopIndex() As Long

  '[ Return Error if not Initialised
  If lvwWork Is Nothing Then Err.Raise 445
  
  '[ Return the Visible Item Count
  VisibleItemTopIndex = SendMessage(lvwWork.Hwnd, LVM_GETTOPINDEX, 0&, ByVal 0&) + 1
       
End Property
Public Sub SortColumnContent(ByVal ColumnIndex As Long)
  On Error GoTo CATCH
  
  Static sOrder     As Boolean
  Dim lngCount      As Long
  Dim varContent    As Variant
  Dim lngSortColumn As Long
 
  '[ Return Error if Not Initialised
  If lvwWork Is Nothing Then Err.Raise 445
  
  '[ Sort Column
  With lvwWork
  
    '[ Add Sort/Temp Column
    If Not .ColumnHeaders.Item(.ColumnHeaders.Count).Text = "~TEMP~" Then
      .ColumnHeaders.Add , , "~TEMP~", 0
    End If
    
    '[ Do Column Checks
    If .ColumnHeaders(ColumnIndex).Tag = "DATE" Then
      
      '[ - Set Sort Column
      lngSortColumn = .ColumnHeaders.Count
      For lngCount = 1 To .ListItems.Count
        If (ColumnIndex - 1) = 0 Then
              .ListItems(lngCount).SubItems(lngSortColumn - 1) = Format$(.ListItems(lngCount).Text, "yyyymmdd")
        Else: .ListItems(lngCount).SubItems(lngSortColumn - 1) = Format$(.ListItems(lngCount).SubItems(ColumnIndex - 1), "yyyymmdd")
        End If
      Next lngCount
      
    ElseIf .ColumnHeaders(ColumnIndex).Tag = "NUMERIC" Then
      
      ' - Set Sort Column
      lngSortColumn = .ColumnHeaders.Count
      For lngCount = 1 To .ListItems.Count
        If (ColumnIndex - 1) = 0 Then
              .ListItems(lngCount).SubItems(lngSortColumn - 1) = Format$(.ListItems(lngCount).Text, "000000000000.00")
        Else: .ListItems(lngCount).SubItems(lngSortColumn - 1) = Format$(.ListItems(lngCount).SubItems(ColumnIndex - 1), "000000000000.00")
        End If
      Next
    Else
      
      ' - Set Sort Column
      lngSortColumn = ColumnIndex
      
    End If
    
    '[ Check if Change in Column Thus Sequence
    If Not .Sorted Then
      sOrder = False
    ElseIf Not (lngSortColumn - 1) = .SortKey Then
      sOrder = False
    Else
      sOrder = Not sOrder
    End If
     
    'Do Sort
    .SortKey = (lngSortColumn - 1)
    .SortOrder = Abs(sOrder)
    .Sorted = True
     
    'Set the Sort Icon
    '-----------------
    For lngCount = 1 To .ColumnHeaders.Count
      If lngCount = ColumnIndex Then
            .ColumnHeaders(lngCount).Icon = .SortOrder + 1
      Else: .ColumnHeaders(lngCount).Icon = 0
      End If
    Next lngCount
    
    '[ Ensure Select Column Visible
    .SelectedItem.EnsureVisible
    
  End With
 
CATCH:

End Sub
