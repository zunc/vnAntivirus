Attribute VB_Name = "mGradient"
'================================================
' Module:        mGradient.bas
' Author:        Carles P.V. - 2005
' Dependencies:  None
' Last revision: 2005.05.13
'================================================

Option Explicit

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Const DIB_RGB_COLORS As Long = 0
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long

'//

Public Enum GradientDirectionCts
    [gdHorizontal] = 0
    [gdVertical] = 1
    [gdDownwardDiagonal] = 2
    [gdUpwardDiagonal] = 3
End Enum

Public Sub PaintGradient(ByVal hDC As Long, _
                         ByVal x As Long, _
                         ByVal y As Long, _
                         ByVal Width As Long, _
                         ByVal Height As Long, _
                         ByVal Color1 As Long, _
                         ByVal Color2 As Long, _
                         ByVal GradientDirection As GradientDirectionCts _
                         )

  Dim uBIH    As BITMAPINFOHEADER
  Dim lBits() As Long
  Dim lGrad() As Long
  
  Dim R1      As Long
  Dim G1      As Long
  Dim B1      As Long
  Dim R2      As Long
  Dim G2      As Long
  Dim B2      As Long
  Dim dR      As Long
  Dim dG      As Long
  Dim dB      As Long
  
  Dim Scan    As Long
  Dim i       As Long
  Dim iEnd    As Long
  Dim iOffset As Long
  Dim j       As Long
  Dim jEnd    As Long
  Dim iGrad   As Long
  
    '-- A minor check
    If (Width < 1 Or Height < 1) Then Exit Sub
    
    '-- Decompose colors
    Color1 = Color1 And &HFFFFFF
    R1 = Color1 Mod &H100&
    Color1 = Color1 \ &H100&
    G1 = Color1 Mod &H100&
    Color1 = Color1 \ &H100&
    B1 = Color1 Mod &H100&
    Color2 = Color2 And &HFFFFFF
    R2 = Color2 Mod &H100&
    Color2 = Color2 \ &H100&
    G2 = Color2 Mod &H100&
    Color2 = Color2 \ &H100&
    B2 = Color2 Mod &H100&
    
    '-- Get color distances
    dR = R2 - R1
    dG = G2 - G1
    dB = B2 - B1
    
    '-- Size gradient-colors array
    Select Case GradientDirection
        Case [gdHorizontal]
            ReDim lGrad(0 To Width - 1)
        Case [gdVertical]
            ReDim lGrad(0 To Height - 1)
        Case Else
            ReDim lGrad(0 To Width + Height - 2)
    End Select
    
    '-- Calculate gradient-colors
    iEnd = UBound(lGrad())
    If (iEnd = 0) Then
        '-- Special case (1-pixel wide gradient)
        lGrad(0) = (B1 \ 2 + B2 \ 2) + 256 * (G1 \ 2 + G2 \ 2) + 65536 * (R1 \ 2 + R2 \ 2)
      Else
        For i = 0 To iEnd
            lGrad(i) = B1 + (dB * i) \ iEnd + 256 * (G1 + (dG * i) \ iEnd) + 65536 * (R1 + (dR * i) \ iEnd)
        Next i
    End If
    
    '-- Size DIB array
    ReDim lBits(Width * Height - 1) As Long
    iEnd = Width - 1
    jEnd = Height - 1
    Scan = Width
    
    '-- Render gradient DIB
    Select Case GradientDirection
        
        Case [gdHorizontal]
        
            For j = 0 To jEnd
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(i - iOffset)
                Next i
                iOffset = iOffset + Scan
            Next j
        
        Case [gdVertical]
        
            For j = jEnd To 0 Step -1
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(j)
                Next i
                iOffset = iOffset + Scan
            Next j
            
        Case [gdDownwardDiagonal]
            
            iOffset = jEnd * Scan
            For j = 1 To jEnd + 1
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(iGrad)
                    iGrad = iGrad + 1
                Next i
                iOffset = iOffset - Scan
                iGrad = j
            Next j
            
        Case [gdUpwardDiagonal]
            
            iOffset = 0
            For j = 1 To jEnd + 1
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(iGrad)
                    iGrad = iGrad + 1
                Next i
                iOffset = iOffset + Scan
                iGrad = j
            Next j
    End Select
    
    '-- Define DIB header
    With uBIH
        .biSize = 40
        .biPlanes = 1
        .biBitCount = 32
        .biWidth = Width
        .biHeight = Height
    End With
    
    '-- Paint it!
    Call StretchDIBits(hDC, x, y, Width, Height, 0, 0, Width, Height, lBits(0), uBIH, DIB_RGB_COLORS, vbSrcCopy)
End Sub


