VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum eImageType
    itUNKNOWN = 0
    itGIF = 1
    itJPEG = 2
    itPNG = 3
    itBMP = 4
    itPCX = 5
    itTIFF = 6
End Enum

Private Type WordBytes
    byte1 As Byte
    byte2 As Byte
End Type

Private Type DWordBytes
    byte1 As Byte
    byte2 As Byte
    byte3 As Byte
    byte4 As Byte
End Type

Private Type WordWrapper
    Value As Integer
End Type

Private Type DWordWrapper
    Value As Long
End Type

Private Const TIFF_WIDTH As Integer = 256
Private Const TIFF_HEIGHT As Integer = 257
Private Const TIFF_BITSPERSAMPLE As Integer = 258
Private Const TIFF_BYTE As Integer = 1
Private Const TIFF_WORD As Integer = 3
Private Const TIFF_DWORD As Integer = 4

Private bBuf() As Byte
Private m_Width As Long
Private m_Height As Long
Private m_Depth As Byte
Private m_ImageType As eImageType
Private m_FileSize As Long

Public Property Get Width() As Long
    Width = m_Width
End Property

Public Property Get Height() As Long
    Height = m_Height
End Property

Public Property Get Depth() As Byte
    Depth = m_Depth
End Property

Public Property Get ImageType() As eImageType
    ImageType = m_ImageType
End Property

Public Property Get FileSize() As Long
    FileSize = m_FileSize
End Property

Public Sub ReadImageInfo(sFileName As String)
    
    Dim pIFD As Long
    Dim NumEntries As Long
    Dim pEntry As Long
    Dim i As Long
    Dim iFN As Integer
    
    m_Width = 0
    m_Height = 0
    m_Depth = 0
    m_FileSize = 0
    m_ImageType = itUNKNOWN
    iFN = FreeFile
    Open sFileName For Binary As iFN
    m_FileSize = LOF(iFN)
    ReDim bBuf(m_FileSize)
    Get #iFN, 1, bBuf()
    Close iFN
'PNG File
    If bBuf(0) = 137 And bBuf(1) = 80 And bBuf(2) = 78 Then
        m_ImageType = itPNG
        
        Select Case bBuf(25)
            Case 0: m_Depth = bBuf(24)
            Case 2: m_Depth = bBuf(24) * 3
            Case 3: m_Depth = 8
            Case 4: m_Depth = bBuf(24) * 2
            Case 6: m_Depth = bBuf(24) * 4
            Case Else
                m_ImageType = itUNKNOWN
        End Select
        
        If m_ImageType Then
            m_Width = BEWord(18)
            m_Height = BEWord(22)
        End If
    End If
' GIF
    If bBuf(0) = 71 And bBuf(1) = 73 And bBuf(2) = 70 Then
        m_ImageType = itGIF
        m_Width = LEWord(6)
        m_Height = LEWord(8)
        m_Depth = (bBuf(10) And 7) + 1
    End If
' BMP
    If bBuf(0) = 66 And bBuf(1) = 77 Then
        m_ImageType = itBMP
        m_Width = LEWord(18)
        m_Height = LEWord(22)
        m_Depth = bBuf(28)
    End If
' PCX
    If bBuf(0) = 10 Then
        Dim Xmin As Long
        Dim Xmax As Long
        Dim Ymin As Long
        Dim Ymax As Long
        
        m_ImageType = itPCX
        Xmin = LEWord(4)
        Ymin = LEWord(6)
        Xmax = LEWord(8)
        Ymax = LEWord(10)
        
        m_Width = (Xmax - Xmin) + 1
        m_Height = (Ymax - Ymin) + 1
        m_Depth = bBuf(65) * bBuf(3)
    End If
' TIFF
    If bBuf(0) = 73 And bBuf(1) = 73 And bBuf(2) = 42 Then
        pIFD = LEDWord(4)
        NumEntries = LEWord(pIFD)
        For i = 0 To NumEntries - 1
            pEntry = pIFD + 2 + (12 * i)
            If LEWord(pEntry) = TIFF_WIDTH Then
                Select Case LEWord(pEntry + 2)
                    Case TIFF_BYTE
                        m_Width = bBuf(pEntry + 8)
                    Case TIFF_WORD
                        m_Width = LEWord(pEntry + 8)
                    Case TIFF_DWORD
                        m_Width = LEDWord(pEntry + 8)
                    Case Else
                        m_Width = 0
                End Select
            End If
            ' find height
            If LEWord(pEntry) = TIFF_HEIGHT Then
                Select Case LEWord(pEntry + 2)
                    Case TIFF_BYTE
                        m_Height = bBuf(pEntry + 8)
                    Case TIFF_WORD
                        m_Height = LEWord(pEntry + 8)
                    Case TIFF_DWORD
                        m_Height = LEDWord(pEntry + 8)
                    Case Else
                        m_Height = 0
                End Select
            End If
            If LEWord(pEntry) = TIFF_BITSPERSAMPLE Then
                Select Case LEWord(pEntry + 2)
                    Case TIFF_BYTE
                        m_Depth = bBuf(pEntry + 8)
                    Case TIFF_WORD
                        m_Depth = LEWord(pEntry + 8)
                    Case TIFF_DWORD
                        m_Depth = LEDWord(pEntry + 8)
                    Case Else
                        m_Depth = 0
                End Select
            End If
            
        Next
        If m_Width > 0 And m_Height > 0 And m_Depth > 0 Then
            m_ImageType = itTIFF
        Else
            m_ImageType = itUNKNOWN
            m_Width = 0
            m_Height = 0
            m_Depth = 0
        End If
        
    End If
' TIFF - big endian
    If bBuf(0) = 77 And bBuf(1) = 77 And bBuf(3) = 42 Then
        pIFD = BEDWord(4)
        NumEntries = BEWord(pIFD)
        For i = 0 To NumEntries - 1
            pEntry = pIFD + 2 + (12 * i)
            If BEWord(pEntry) = TIFF_WIDTH Then
                Select Case BEWord(pEntry + 2)
                    Case TIFF_BYTE
                        m_Width = bBuf(pEntry + 8)
                    Case TIFF_WORD
                        m_Width = BEWord(pEntry + 8)
                    Case TIFF_DWORD
                        m_Width = BEDWord(pEntry + 8)
                    Case Else
                        m_Width = 0
                End Select
            End If
            ' find height
            If BEWord(pEntry) = TIFF_HEIGHT Then
                Select Case BEWord(pEntry + 2)
                    Case TIFF_BYTE
                        m_Height = bBuf(pEntry + 8)
                    Case TIFF_WORD
                        m_Height = BEWord(pEntry + 8)
                    Case TIFF_DWORD
                        m_Height = BEDWord(pEntry + 8)
                    Case Else
                        m_Height = 0
                End Select
            End If
            ' find depth
            If BEWord(pEntry) = TIFF_BITSPERSAMPLE Then
                Select Case BEWord(pEntry + 2)
                    Case TIFF_BYTE
                        m_Depth = bBuf(pEntry + 8)
                    Case TIFF_WORD
                        m_Depth = BEWord(pEntry + 8)
                    Case TIFF_DWORD
                        m_Depth = BEDWord(pEntry + 8)
                    Case Else
                        m_Depth = 0
                End Select
            End If
        Next
        If m_Width > 0 And m_Height > 0 And m_Depth > 0 Then
            m_ImageType = itTIFF
        Else
            m_ImageType = itUNKNOWN
            m_Width = 0
            m_Height = 0
            m_Depth = 0
        End If
        
    End If
' Check for JPEG
    If m_ImageType = itUNKNOWN Then
        Dim lPos As Long
        Do
            If (bBuf(lPos) = &HFF And bBuf(lPos + 1) = &HD8 And bBuf(lPos + 2) = &HFF) _
            Or (lPos >= m_FileSize - 10) Then Exit Do
            lPos = lPos + 1
        Loop
        lPos = lPos + 2
        If lPos >= m_FileSize - 10 Then Exit Sub
        Do
            Do
                If bBuf(lPos) = &HFF And bBuf(lPos + 1) <> &HFF Then Exit Do
                lPos = lPos + 1
                If lPos >= m_FileSize - 10 Then Exit Sub
            Loop
            lPos = lPos + 1
            If (bBuf(lPos) >= &HC0) And (bBuf(lPos) <= &HC3) Then Exit Do
            lPos = lPos + BEWord(lPos + 1)
            If lPos >= m_FileSize - 10 Then Exit Sub
        Loop
        m_ImageType = itJPEG
        m_Height = BEWord(lPos + 4)
        m_Width = BEWord(lPos + 6)
        m_Depth = bBuf(lPos + 8) * 8
    End If
    ReDim bBuf(0)
End Sub

Private Function LEWord(position As Long) As Long
    Dim x1 As WordBytes
    Dim x2 As WordWrapper
    x1.byte1 = bBuf(position)
    x1.byte2 = bBuf(position + 1)
    LSet x2 = x1
    LEWord = x2.Value
End Function

Private Function BEWord(position As Long) As Long
    Dim x1 As WordBytes
    Dim x2 As WordWrapper
    x1.byte1 = bBuf(position + 1)
    x1.byte2 = bBuf(position)
    LSet x2 = x1
    BEWord = x2.Value
End Function

Private Function LEDWord(position As Long) As Long
    Dim x1 As DWordBytes
    Dim x2 As DWordWrapper
    x1.byte1 = bBuf(position)
    x1.byte2 = bBuf(position + 1)
    x1.byte3 = bBuf(position + 2)
    x1.byte4 = bBuf(position + 3)
    LSet x2 = x1
    LEDWord = x2.Value
End Function

Private Function BEDWord(position As Long) As Long
    Dim x1 As DWordBytes
    Dim x2 As DWordWrapper
    x1.byte1 = bBuf(position + 3)
    x1.byte2 = bBuf(position + 2)
    x1.byte3 = bBuf(position + 1)
    x1.byte4 = bBuf(position)
    LSet x2 = x1
    BEDWord = x2.Value
End Function
