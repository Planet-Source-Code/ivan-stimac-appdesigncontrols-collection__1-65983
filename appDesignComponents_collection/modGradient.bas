Attribute VB_Name = "modGradient"
' Modude Gradient
' by UserXP
' Do it what ever you want
' =========================
Option Explicit

'Declare Function GetTickCount Lib "kernel32" () As Long
Type TRIVERTEX
    X As Long
    Y As Long
    Red As Integer 'Ushort value
    Green As Integer 'Ushort value
    Blue As Integer 'ushort value
    Alpha As Integer 'ushort
End Type
Type GRADIENT_RECT
    UpperLeft As Long  'In reality this is a UNSIGNED Long
    LowerRight As Long 'In reality this is a UNSIGNED Long
End Type
Private Type GRADIENT_TRIANGLE
    Vertex1 As Long
    Vertex2 As Long
    Vertex3 As Long
End Type

Const GRADIENT_FILL_RECT_H As Long = &H0 'In this mode, two endpoints describe a
'rectangle. The rectangle is
'defined to have a constant color (specified by the TRIVERTEX structure) for the
'left and right edges. GDI interpolates
'the color from the top to bottom edge and fills the interior.
Const GRADIENT_FILL_RECT_V  As Long = &H1 'In this mode, two endpoints describe
'a rectangle. The rectangle
' is defined to have a constant color (specified by the TRIVERTEX structure)
'for the top and bottom edges. GDI interpolates
' the color from the top to bottom edge and fills the interior.
Const GRADIENT_FILL_TRIANGLE As Long = &H2 'In this mode, an array of TRIVERTEX
'structures is passed to GDI
'along with a list of array indexes that describe separate triangles. GDI performs
'linear interpolation between triangle vertices
'and fills the interior. Drawing is done directly in 24- and 32-bpp modes. Dithering is performed in 16-, 8.4-, and 1-bpp mode.
Const GRADIENT_FILL_OP_FLAG As Long = &HFF

Enum GRADIENT_FILL_RECT
    FillHor = GRADIENT_FILL_RECT_H
    FillVer = GRADIENT_FILL_RECT_V
End Enum
Enum GRADIENT_TO_CORNER
    All
    TopLeft
    TopRight
    BottomLeft
    BottomRight
End Enum
Enum CRADIENT_DIRECTION
    DirectionSlash
    DirectionBackSlash
End Enum

Declare Function GradientFillTriangle Lib "msimg32" Alias "GradientFill" _
    (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, _
    pMesh As GRADIENT_TRIANGLE, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" _
    (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, _
    pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long

Declare Function GetPixel Lib "gdi32" _
    (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Function DoGradient(ToPic_Pixels As PictureBox, FromColor As Long, ToColor As Long, Optional DrawHorVer As GRADIENT_FILL_RECT = FillHor, Optional Left As Long = 0, Optional Top As Long = 0, Optional Width As Long = -1, Optional Height As Long = -1) As Boolean
    Dim vert(1) As TRIVERTEX
    Dim gRect As GRADIENT_RECT
    Dim R As Byte, G As Byte, B As Byte
    
    If Width < 0 Then Width = ToPic_Pixels.Width
    If Height < 0 Then Height = ToPic_Pixels.Height
    
    Long2RGB FromColor, R, G, B
    With vert(0)
        .X = Left
        .Y = Top
        .Red = Val("&h" & Hex(R) & "00")
        .Green = Val("&h" & Hex(G) & "00")
        .Blue = Val("&h" & Hex(B) & "00")
        .Alpha = 0&
    End With
    
    Long2RGB ToColor, R, G, B
    With vert(1)
        .X = Left + Width
        .Y = Top + Height
        .Red = Val("&h" & Hex(R) & "00")
        .Green = Val("&h" & Hex(G) & "00")
        .Blue = Val("&h" & Hex(B) & "00")
        .Alpha = 0&
    End With

    gRect.UpperLeft = 0
    gRect.LowerRight = 1

    DoGradient = GradientFillRect(ToPic_Pixels.hdc, vert(0), 2, gRect, 1, DrawHorVer)
    
End Function

'With a little modification can be for 4, 5, 6, ... colors
Function DoGradient3Colors(ToPic_Pixels As PictureBox, FromColor As Long, MiddleColor As Long, ToColor As Long, Optional DrawHorVer As GRADIENT_FILL_RECT = FillHor, Optional Left As Long = 0, Optional Top As Long = 0, Optional Width As Long = -1, Optional Height As Long = -1) As Boolean
    Dim vert(3) As TRIVERTEX
    Dim gRect(1) As GRADIENT_RECT
    Dim R As Byte, G As Byte, B As Byte
    Dim RR As Long, GG As Long, BB As Long
    
    If Width < 0 Then Width = ToPic_Pixels.Width
    If Height < 0 Then Height = ToPic_Pixels.Height
    
    Long2RGB FromColor, R, G, B
    With vert(0)
        .X = Left
        .Y = Top
        .Red = Val("&h" & Hex(R) & "00")
        .Green = Val("&h" & Hex(G) & "00")
        .Blue = Val("&h" & Hex(B) & "00")
        .Alpha = 0&
    End With
    
    Long2RGB MiddleColor, R, G, B
    With vert(1)
        .X = Left + (Width / IIf(DrawHorVer = FillHor, 2, 1))
        .Y = Top + (Height / IIf(DrawHorVer = FillHor, 1, 2))
        .Red = Val("&h" & Hex(R) & "00")
        .Green = Val("&h" & Hex(G) & "00")
        .Blue = Val("&h" & Hex(B) & "00")
        .Alpha = 0&
    End With
        
    With vert(2)
        .X = Left + IIf(DrawHorVer = FillHor, Width / 2, 0)
        .Y = Top + IIf(DrawHorVer = FillHor, 0, Height / 2)
        .Red = Val("&h" & Hex(R) & "00")
        .Green = Val("&h" & Hex(G) & "00")
        .Blue = Val("&h" & Hex(B) & "00")
        .Alpha = 0&
    End With
    
    Long2RGB ToColor, R, G, B
    With vert(3)
        .X = Left + Width
        .Y = Top + Height
        .Red = Val("&h" & Hex(R) & "00")
        .Green = Val("&h" & Hex(G) & "00")
        .Blue = Val("&h" & Hex(B) & "00")
        .Alpha = 0&
    End With
    
    gRect(0).UpperLeft = 0
    gRect(0).LowerRight = 1
    gRect(1).UpperLeft = 2
    gRect(1).LowerRight = 3
    DoGradient3Colors = GradientFillRect(ToPic_Pixels.hdc, vert(0), 4, gRect(0), 2, DrawHorVer)
    
End Function

Sub DoGradient45Colors(ToPic As PictureBox, colorTopLeft As Long, colorTopRight As Long, colorBottomLeft As Long, colorBottomRight As Long, Optional colorCenter, Optional Left As Long = 0, Optional Top As Long = 0, Optional Width As Long = -1, Optional Height As Long = -1, Optional ByVal ApplyGradientToCorner As GRADIENT_TO_CORNER = All, Optional DirectionForAll As CRADIENT_DIRECTION = DirectionSlash)
    Dim vert(4) As TRIVERTEX
    Dim gTRi(1) As GRADIENT_TRIANGLE
    Dim R As Byte, G As Byte, B As Byte
    Dim TohDC As Long
    
    TohDC = ToPic.hdc
    
    If Not IsMissing(colorCenter) Then ApplyGradientToCorner = All
    If Width < 0 Then Width = ToPic.Width
    If Height < 0 Then Height = ToPic.Height
    
    Long2RGB colorTopLeft, R, G, B
    vert(0).X = Left
    vert(0).Y = Top
    vert(0).Red = Val("&h" & Hex(R) & "00")
    vert(0).Green = Val("&h" & Hex(G) & "00")
    vert(0).Blue = Val("&h" & Hex(B) & "00")
    vert(0).Alpha = 0&
    
    Long2RGB colorTopRight, R, G, B
    vert(1).X = Left + Width
    vert(1).Y = Top
    vert(1).Red = Val("&h" & Hex(R) & "00")
    vert(1).Green = Val("&h" & Hex(G) & "00")
    vert(1).Blue = Val("&h" & Hex(B) & "00")
    vert(1).Alpha = 0&
    
    Long2RGB colorBottomRight, R, G, B
    vert(2).X = Left + Width
    vert(2).Y = Top + Height
    vert(2).Red = Val("&h" & Hex(R) & "00")
    vert(2).Green = Val("&h" & Hex(G) & "00")
    vert(2).Blue = Val("&h" & Hex(B) & "00")
    vert(2).Alpha = 0&
    
    Long2RGB colorBottomLeft, R, G, B
    vert(3).X = Left
    vert(3).Y = Top + Height
    vert(3).Red = Val("&h" & Hex(R) & "00")
    vert(3).Green = Val("&h" & Hex(G) & "00")
    vert(3).Blue = Val("&h" & Hex(B) & "00")
    vert(3).Alpha = 0&
    
    Dim n1 As Long, n2 As Long, n3 As Long
    If ApplyGradientToCorner = TopLeft Then
        n1 = 1: n2 = 3: n3 = 0
    ElseIf ApplyGradientToCorner = TopRight Then
        n1 = 0: n2 = 2: n3 = 1
    ElseIf ApplyGradientToCorner = BottomLeft Then
        n1 = 0: n2 = 2: n3 = 3
    ElseIf ApplyGradientToCorner = BottomRight Then
        n1 = 1: n2 = 3: n3 = 2
    Else
        n1 = 0: n2 = 1: n3 = 3
    End If
    gTRi(0).Vertex1 = n1
    gTRi(0).Vertex2 = n2
    gTRi(0).Vertex3 = n3
        
    gTRi(1).Vertex1 = 1
    gTRi(1).Vertex2 = 2
    gTRi(1).Vertex3 = 3
    
    If ApplyGradientToCorner = All Then
        If DirectionForAll = DirectionSlash Then
            gTRi(0).Vertex1 = 0: gTRi(0).Vertex2 = 1: gTRi(0).Vertex3 = 2
            gTRi(1).Vertex1 = 0: gTRi(1).Vertex2 = 2: gTRi(1).Vertex3 = 3
        End If
        If IsMissing(colorCenter) Then
            GradientFillTriangle TohDC, vert(0), 4, gTRi(0), 2, GRADIENT_FILL_TRIANGLE
        Else
            Long2RGB CLng(colorCenter), R, G, B
            vert(4).X = Left + Width / 2
            vert(4).Y = Top + Height / 2
            vert(4).Red = Val("&h" & Hex(R) & "00")
            vert(4).Green = Val("&h" & Hex(G) & "00")
            vert(4).Blue = Val("&h" & Hex(B) & "00")
            vert(4).Alpha = 0&
        
            gTRi(0).Vertex1 = 0: gTRi(0).Vertex2 = 1: gTRi(0).Vertex3 = 4
            GradientFillTriangle TohDC, vert(0), 5, gTRi(0), 1, GRADIENT_FILL_TRIANGLE
            gTRi(0).Vertex1 = 1: gTRi(0).Vertex2 = 2: gTRi(0).Vertex3 = 4
            GradientFillTriangle TohDC, vert(0), 5, gTRi(0), 1, GRADIENT_FILL_TRIANGLE
            gTRi(0).Vertex1 = 2: gTRi(0).Vertex2 = 3: gTRi(0).Vertex3 = 4
            GradientFillTriangle TohDC, vert(0), 5, gTRi(0), 1, GRADIENT_FILL_TRIANGLE
            gTRi(0).Vertex1 = 3: gTRi(0).Vertex2 = 0: gTRi(0).Vertex3 = 4
            GradientFillTriangle TohDC, vert(0), 5, gTRi(0), 1, GRADIENT_FILL_TRIANGLE
        End If
    Else
        GradientFillTriangle TohDC, vert(0), 4, gTRi(0), 1, GRADIENT_FILL_TRIANGLE
    End If
        
End Sub

Function Long2RGB(nColor As Long, Red As Byte, Green As Byte, Blue As Byte)
    Red = (nColor And &HFF&)
    Green = (nColor And &HFF00&) / &H100
    Blue = (nColor And &HFF0000) / &H10000
End Function




