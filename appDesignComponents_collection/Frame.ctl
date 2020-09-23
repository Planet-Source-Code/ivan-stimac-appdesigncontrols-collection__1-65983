VERSION 5.00
Begin VB.UserControl Frame 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   FillStyle       =   0  'Solid
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "Frame.ctx":0000
End
Attribute VB_Name = "Frame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long

Private BC As OLE_COLOR, FillC As OLE_COLOR, FillC2 As OLE_COLOR
Private FC  As OLE_COLOR, FCDisabled As OLE_COLOR
Private BDRC  As OLE_COLOR, bdrC2 As OLE_COLOR
Private enbl As Boolean

Private strCap As String

Public Enum eFill
    All
    OnlyInside
End Enum
Private mFill As eFill

Public Enum fr_eStyle
    Standard
    CaptionInside
End Enum
Private mfStyle As fr_eStyle

'-----------------------------------------------------------------------------------

'read properties
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    Set Font = PropBag.ReadProperty("Font", Ambient.Font) 'UserControl.Font
    BC = PropBag.ReadProperty("BackColor", vbButtonFace)
    FillC = PropBag.ReadProperty("FillColor", vbHighlight)
    FillC2 = PropBag.ReadProperty("FillColorCaption", vbHighlight)
    FC = PropBag.ReadProperty("ForeColor", vbBlack)
    FCDisabled = PropBag.ReadProperty("ForeColorDisabled", vbBlack)
    
    BDRC = PropBag.ReadProperty("BorderColor", vbBlack)
    bdrC2 = PropBag.ReadProperty("BorderColorCaption", vbBlack)
    
    enbl = PropBag.ReadProperty("Enabled", True)
    mFill = PropBag.ReadProperty("Fill", All)
    strCap = PropBag.ReadProperty("Caption", "Frame")
    
    mfStyle = PropBag.ReadProperty("Style", Standard)
    
    UserControl.BackColor = BC
    reDraw
End Sub

Private Sub UserControl_Resize()
    reDraw
End Sub

'write properties
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next
    PropBag.WriteProperty "Font", UserControl.Font, Ambient.Font
    PropBag.WriteProperty "BackColor", BC, vbHighlight
    PropBag.WriteProperty "FillColor", FillC, vbHighlight
    PropBag.WriteProperty "FillColorCaption", FillC2, vbHighlight
    PropBag.WriteProperty "ForeColor", FC, vbBlack
    PropBag.WriteProperty "ForeColorDisabled", FCDisabled, vbBlack
    
    PropBag.WriteProperty "BorderColor", BDRC, vbBlack
    PropBag.WriteProperty "BorderColorCaption", bdrC2, vbBlack
    
    PropBag.WriteProperty "Enabled", enbl, True
    PropBag.WriteProperty "Fill", mFill, All
    PropBag.WriteProperty "Caption", strCap, "Frame"
    
    PropBag.WriteProperty "Style", mfStyle, Standard
    
End Sub

Private Sub UserControl_Initialize()
    BC = vbButtonFace
    FillC = vbHighlight
    FillC2 = vbHighlight
    
    FC = vbBlack
    FCDisabled = vbBlack
    
    BDRC = vbActiveBorder
    bdrC2 = vbActiveBorder
    
    enbl = True
    strCap = "Frame"
    mFill = All
    mfStyle = Standard
End Sub


Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property
Public Property Set Font(ByVal nFON As Font)
    Set UserControl.Font = nFON
    reDraw
    PropertyChanged "Font"
End Property
'back color
Public Property Get BackColor() As OLE_COLOR
    BackColor = BC
End Property
Public Property Let BackColor(ByVal nC As OLE_COLOR)
    BC = nC
    UserControl.BackColor = BC
    reDraw
    PropertyChanged "BackColor"
End Property
'fill color
Public Property Get FillColor() As OLE_COLOR
    FillColor = FillC
End Property
Public Property Let FillColor(ByVal nC As OLE_COLOR)
    FillC = nC
    reDraw
    PropertyChanged "FillColor"
End Property
'fill color caption
Public Property Get FillColorCaption() As OLE_COLOR
    FillColorCaption = FillC2
End Property
Public Property Let FillColorCaption(ByVal nC As OLE_COLOR)
    FillC2 = nC
    reDraw
    PropertyChanged "FillColorCaption"
End Property
'fore color
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = FC
End Property
Public Property Let ForeColor(ByVal nC As OLE_COLOR)
    FC = nC
    reDraw
    PropertyChanged "ForeColor"
End Property
'fore color disabled
Public Property Get ForeColorDisabled() As OLE_COLOR
    ForeColorDisabled = FCDisabled
End Property
Public Property Let ForeColorDisabled(ByVal nC As OLE_COLOR)
    FCDisabled = nC
    reDraw
    PropertyChanged "ForeColorDisabled"
End Property

'border color
Public Property Get BorderColor() As OLE_COLOR
    BorderColor = BDRC
End Property
Public Property Let BorderColor(ByVal nC As OLE_COLOR)
    BDRC = nC
    reDraw
    PropertyChanged "BorderColor"
End Property
'border color capt
Public Property Get BorderColorCaption() As OLE_COLOR
    BorderColorCaption = bdrC2
End Property
Public Property Let BorderColorCaption(ByVal nC As OLE_COLOR)
    bdrC2 = nC
    reDraw
    PropertyChanged "BorderColorCaption"
End Property

Public Property Get Enabled() As Boolean
    Enabled = enbl
End Property
Public Property Let Enabled(ByVal nV As Boolean)
    enbl = nV
    reDraw
    PropertyChanged "Enabled"
End Property

Public Property Get Fill() As eFill
    Fill = mFill
End Property
Public Property Let Fill(ByVal nV As eFill)
    mFill = nV
    reDraw
    PropertyChanged "Fill"
End Property

Public Property Get Style() As fr_eStyle
    Style = mfStyle
End Property
Public Property Let Style(ByVal nV As fr_eStyle)
    mfStyle = nV
    reDraw
    PropertyChanged "Style"
End Property

Public Property Get Caption() As String
    Caption = strCap
End Property
Public Property Let Caption(ByVal nV As String)
    strCap = nV
    reDraw
    PropertyChanged "Caption"
End Property

Private Function reDraw()
    Dim i As Integer, ucSH As Integer, ucSW As Integer, cRadius As Integer
    UserControl.Cls
    UserControl.BackColor = BC
    UserControl.FillColor = FillC
    ucSH = UserControl.ScaleHeight
    ucSW = UserControl.ScaleWidth
    
    cRadius = ((ucSH + ucSW) / 2) / 40
    
    UserControl.Enabled = True
    If enbl = True Then
        UserControl.ForeColor = FC
    Else
        UserControl.ForeColor = FCDisabled
    End If
    
    'draw rounded rectangle
    Circle (cRadius, UserControl.TextHeight("H") + cRadius), cRadius, BDRC, 3.14 / 2, 3.14
    Line (cRadius, UserControl.TextHeight("H"))-(ucSW - cRadius, UserControl.TextHeight("H")), BDRC
    Circle (ucSW - cRadius - 1, UserControl.TextHeight("H") + cRadius), cRadius, BDRC, 0, 3.14 / 2
    Line (ucSW - 1, UserControl.TextHeight("H") + cRadius)-(ucSW - 1, ucSH - cRadius), BDRC
    Circle (ucSW - cRadius - 1, ucSH - cRadius - 1), cRadius, BDRC, (1.5) * (3.14), 0
    Line (cRadius, ucSH - 1)-(ucSW - 1 - cRadius, ucSH - 1), BDRC
    Circle (cRadius, ucSH - cRadius - 1), cRadius, BDRC, 3.14, 1.5 * 3.14
    Line (0, UserControl.TextHeight("H") + cRadius - 1)-(0, ucSH - cRadius + 1), BDRC
    'fill inside box
    If mFill = OnlyInside Then
        ExtFloodFill UserControl.hdc, ucSW / 2, ucSH / 2, UserControl.Point(ucSW / 2, ucSH / 2), 1
    End If
    
    If strCap <> "" Then
        If mfStyle = Standard Then
            'hide line where coming caption
            Line (cRadius + 30 + UserControl.TextWidth(strCap), UserControl.TextHeight("H"))-(cRadius + 20, UserControl.TextHeight("H")), BC
            '
            UserControl.CurrentX = cRadius + 25
            UserControl.CurrentY = UserControl.TextHeight("H") - UserControl.TextHeight("H") / 2
        ElseIf mfStyle = CaptionInside Then
            cRadius = UserControl.TextHeight(strCap) / 2
            UserControl.FillColor = FillC2
            'drawing box for caption
            Line (cRadius + 30, UserControl.TextHeight("H"))-(cRadius + 30, UserControl.TextHeight("H") * 2), bdrC2
            Circle (cRadius * 2 + 30, UserControl.TextHeight("H") * 2), cRadius, bdrC2, 3.14, 1.5 * 3.14
            Line (cRadius * 2 + 30 - 1, UserControl.TextHeight("H") * 2 + cRadius)-(cRadius * 2 + 30 + 10 + UserControl.TextWidth(strCap), UserControl.TextHeight("H") * 2 + cRadius), bdrC2
            Circle (cRadius * 2 + 30 + 10 + UserControl.TextWidth(strCap), UserControl.TextHeight("H") * 2), cRadius, bdrC2, 1.5 * 3.14, 0
            Line (cRadius * 3 + 30 + 10 + UserControl.TextWidth(strCap), UserControl.TextHeight("H"))-(cRadius * 3 + 30 + 10 + UserControl.TextWidth(strCap), UserControl.TextHeight("H") * 2), bdrC2
            ExtFloodFill UserControl.hdc, cRadius + 32, UserControl.TextHeight("H") + 2, UserControl.Point(cRadius + 32, UserControl.TextHeight("H") + 2), 1
            'hide line at top of this box
            Line (cRadius + 30, UserControl.TextHeight("H"))-(cRadius * 3 + 30 + 10 + UserControl.TextWidth(strCap), UserControl.TextHeight("H")), FillC2
            UserControl.CurrentX = cRadius * 2 + 30 + UserControl.TextHeight("H") / 2
            UserControl.CurrentY = UserControl.TextHeight("H") + 3 '+ UserControl.TextHeight("H") / 2
            
        End If
        UserControl.Print strCap
    End If
    UserControl.Enabled = enbl
End Function
