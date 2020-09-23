VERSION 5.00
Begin VB.UserControl Counter 
   AutoRedraw      =   -1  'True
   ClientHeight    =   2475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5895
   FillStyle       =   0  'Solid
   ScaleHeight     =   165
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   393
   ToolboxBitmap   =   "Counter.ctx":0000
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2940
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1860
      Top             =   660
   End
End
Attribute VB_Name = "Counter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
'api
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long


'
Private BC As OLE_COLOR, FillC2 As OLE_COLOR, buttBC As OLE_COLOR, buttHoverBC As OLE_COLOR
Private FC  As OLE_COLOR, FCDisabled  As OLE_COLOR, FCButt As OLE_COLOR, FCHover As OLE_COLOR
Private BDRC  As OLE_COLOR, BDRCDisabled  As OLE_COLOR

Private enbl As Boolean
Private mVal As Integer, minV As Integer, maxV As Integer

Private hoverButt As Byte

Private num1 As Integer
Public Event Click()


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
'fill color caption
Public Property Get FillColorValue() As OLE_COLOR
    FillColorValue = FillC2
End Property
Public Property Let FillColorValue(ByVal nC As OLE_COLOR)
    FillC2 = nC
    reDraw
    PropertyChanged "FillColorValue"
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
'fore color button
Public Property Get ForeColorButton() As OLE_COLOR
    ForeColorButton = FCButt
End Property
Public Property Let ForeColorButton(ByVal nC As OLE_COLOR)
    FCButt = nC
    reDraw
    PropertyChanged "ForeColorButton"
End Property
'fore color button hover
Public Property Get ForeColorButtonHover() As OLE_COLOR
    ForeColorButtonHover = FCHover
End Property
Public Property Let ForeColorButtonHover(ByVal nC As OLE_COLOR)
    FCHover = nC
    reDraw
    PropertyChanged "ForeColorButtonHover"
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
Public Property Get BorderColorDisabled() As OLE_COLOR
    BorderColorDisabled = BDRCDisabled
End Property
Public Property Let BorderColorDisabled(ByVal nC As OLE_COLOR)
    BDRCDisabled = nC
    reDraw
    PropertyChanged "BorderColorDisabled"
End Property

Public Property Get Enabled() As Boolean
    Enabled = enbl
End Property
Public Property Let Enabled(ByVal nV As Boolean)
    enbl = nV
    reDraw
    PropertyChanged "Enabled"
End Property


Public Property Get Value() As Integer
    Value = mVal
End Property
Public Property Let Value(ByVal nV As Integer)
    mVal = nV
    reDraw
    PropertyChanged "Value"
End Property

Public Property Get MinValue() As Integer
    MinValue = minV
End Property
Public Property Let MinValue(ByVal nV As Integer)
    minV = nV
    reDraw
    PropertyChanged "MinValue"
End Property

Public Property Get MaxValue() As Integer
    MaxValue = maxV
End Property
Public Property Let MaxValue(ByVal nV As Integer)
    maxV = nV
    reDraw
    PropertyChanged "MaxValue"
End Property

'border color butt
Public Property Get BackColorButton() As OLE_COLOR
    BackColorButton = buttBC
End Property
Public Property Let BackColorButton(ByVal nC As OLE_COLOR)
    buttBC = nC
    reDraw
    PropertyChanged "BackColorButton"
End Property
'border color hover
Public Property Get BackColorButtonHover() As OLE_COLOR
    BackColorButtonHover = buttHoverBC
End Property
Public Property Let BackColorButtonHover(ByVal nC As OLE_COLOR)
    buttHoverBC = nC
    reDraw
    PropertyChanged "BackColorButtonHover"
End Property

Private Sub Timer2_Timer()
    If Timer2.Interval = 1 Then Timer2.Interval = 500
    If mVal + num1 >= minV And mVal + num1 <= maxV Then mVal = mVal + num1
    If num1 = -1 Then
        reDraw 1
    ElseIf num1 = 1 Then
        reDraw 0, 1
    End If
    
    'decreasing interval to increase change speed
    If Timer2.Interval > 50 Then Timer2.Interval = Timer2.Interval - 50
    RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
    enbl = True
    buttBC = vbButtonFace
    buttHoverBC = vbHighlight
    BC = vbButtonFace
    
    FCHover = vbBlack
    FCButt = vbBlack
    FCDisabled = vbBlack
    FC = vbWhite
    
    FillC2 = vbWhite
    
    minV = 0
    maxV = 10
End Sub

'timer: checking is mouse over control
Private Sub Timer1_Timer()
    Dim lpPos As POINTAPI
    Dim lhWnd As Long
    GetCursorPos lpPos
    lhWnd = WindowFromPoint(lpPos.X, lpPos.Y)
    If lhWnd <> UserControl.hWnd And hoverButt <> 0 Then
        'if NOT redraw control
        reDraw 0, 0
        hoverButt = 0
        Timer1.Enabled = False
        Timer2.Enabled = False
        num1 = 0
    End If
    
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim btsW As Integer, ucSW As Integer
    ucSW = UserControl.ScaleWidth
    btsW = (UserControl.TextWidth(">") + 20) * 2
    
    Timer2.Interval = 1
    If X <= ucSW - (btsW) / 2 + 3 And X >= ucSW - (btsW) + 10 Then
        'If mVal > minV Then mVal = mVal - 1
        num1 = -1
        Timer2.Enabled = True
        reDraw 1
    ElseIf X >= ucSW - (btsW) / 2 + 6 And X <= ucSW - 1 Then
        'If mVal < maxV Then mVal = mVal + 1
        num1 = 1
        Timer2.Enabled = True
        reDraw 0, 1
    End If
    hoverButt = 0
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim mHbutt As Byte, btsW As Integer, ucSW As Integer
    Timer1.Enabled = True
    ucSW = UserControl.ScaleWidth
    btsW = (UserControl.TextWidth(">") + 20) * 2
    
    If X <= ucSW - (btsW) / 2 + 3 And X >= ucSW - (btsW) + 10 Then
        If hoverButt <> 1 Then reDraw 1, 0
        hoverButt = 1
    ElseIf X >= ucSW - (btsW) / 2 + 6 And X <= ucSW - 1 Then
        If hoverButt <> 2 Then reDraw 0, 1
        hoverButt = 2
    Else
        If hoverButt <> 0 Then reDraw 0, 0
        hoverButt = 0
    End If

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer2.Enabled = False
    num1 = 0
End Sub

'-----------------------------------------------------------------------------------

'read properties
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   ' On Error Resume Next
    Set Font = PropBag.ReadProperty("Font", Ambient.Font) 'UserControl.Font
    BC = PropBag.ReadProperty("BackColor", vbButtonFace)
    buttBC = PropBag.ReadProperty("BackColorButton", vbButtonFace)
    buttHoverBC = PropBag.ReadProperty("BackColorButtonHover", vbHighlight)
    
    FillC2 = PropBag.ReadProperty("FillColorValue", vbHighlight)
    FC = PropBag.ReadProperty("ForeColor", vbBlack)
    FCDisabled = PropBag.ReadProperty("ForeColorDisabled", vbBlack)
    
    FCButt = PropBag.ReadProperty("ForeColorButton", vbBlack)
    FCHover = PropBag.ReadProperty("ForeColorButtonHover", vbBlack)
    
    BDRC = PropBag.ReadProperty("BorderColor", vbBlack)
    BDRCDisabled = PropBag.ReadProperty("BorderColorCaption", vbBlack)
    
    enbl = PropBag.ReadProperty("Enabled", True)
    mVal = PropBag.ReadProperty("Value", 0)
    minV = PropBag.ReadProperty("MinValue", 0)
    maxV = PropBag.ReadProperty("MaxValue", 10)
    
    UserControl.BackColor = BC
    reDraw
End Sub

Private Sub UserControl_Resize()
    reDraw
End Sub

'write properties
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    'On Error Resume Next
    PropBag.WriteProperty "Font", UserControl.Font, Ambient.Font
    PropBag.WriteProperty "BackColor", BC, vbHighlight
    PropBag.WriteProperty "BackColorButton", buttBC, vbHighlight
    PropBag.WriteProperty "BackColorButtonHover", buttHoverBC, vbHighlight
    
    PropBag.WriteProperty "FillColorValue", FillC2, vbHighlight
    PropBag.WriteProperty "ForeColor", FC, vbBlack
    PropBag.WriteProperty "ForeColorDisabled", FCDisabled, vbBlack
    
    PropBag.WriteProperty "BorderColor", BDRC, vbBlack
    PropBag.WriteProperty "BorderColorCaption", BDRCDisabled, vbBlack
    
    PropBag.WriteProperty "ForeColorButton", FCButt, vbBlack
    PropBag.WriteProperty "ForeColorButtonHover", FCHover, vbBlack
    
    PropBag.WriteProperty "Enabled", enbl, True
    PropBag.WriteProperty "Value", mVal, 0
    PropBag.WriteProperty "MinValue", minV, 0
    PropBag.WriteProperty "MaxValue", maxV, 10
End Sub

Private Function reDraw(Optional B1 As Byte = 0, Optional B2 As Byte = 0)
    Dim mSpace As Integer, ucSH As Integer, ucSW As Integer, cRadius As Integer
    Dim btsW As Integer
    
    Dim bordC As OLE_COLOR, buttForeC As OLE_COLOR, foreC As OLE_COLOR
    
    UserControl.Cls
    UserControl.BackColor = BC
    mSpace = 7
    
    '
    UserControl.Enabled = enbl
    'set control height if font is changed
    UserControl.Height = (UserControl.TextHeight("H") + mSpace) * Screen.TwipsPerPixelY
    
    ucSH = UserControl.ScaleHeight
    ucSW = UserControl.ScaleWidth
    
    'buttons width
    btsW = (UserControl.TextWidth(">") + 20) * 2
    'set colors
    If enbl = True Then
        bordC = BDRC
        buttForeC = FCButt
        foreC = FC
    Else
        bordC = BDRCDisabled
        buttForeC = FCDisabled
        foreC = FCDisabled
    End If
    '
    UserControl.FillColor = FillC2
    'draw value field
    UserControl.Line (0, 0)-(ucSW - btsW, ucSH - 1), bordC, B
    
    UserControl.ForeColor = foreC
    UserControl.CurrentX = 3
    UserControl.CurrentY = ucSH / 2 - UserControl.TextHeight("H") / 2
    UserControl.Print Value
    '
    'reset colors
    UserControl.FillColor = buttBC
    UserControl.ForeColor = buttForeC
    '1=hover
    If B1 = 1 Then
        UserControl.FillColor = buttHoverBC
        UserControl.ForeColor = FCHover
    End If
    'draw button '<'
    UserControl.Line (ucSW - (btsW) / 2 + 3, 0)-(ucSW - (btsW) + 10, ucSH - 1), bordC, B
    UserControl.CurrentX = (ucSW - (btsW) / 2 + 3 + ucSW - (btsW) + 10) / 2 - UserControl.TextWidth(">") / 2
    UserControl.CurrentY = ucSH / 2 - UserControl.TextHeight("H") / 2
    UserControl.Print "<"
    '
    'reset colors
    UserControl.FillColor = buttBC
    UserControl.ForeColor = buttForeC
    'if butt2 hover
    If B2 = 1 Then
        UserControl.FillColor = buttHoverBC
        UserControl.ForeColor = FCHover
    End If
    'draw button '>'
    UserControl.Line (ucSW - (btsW) / 2 + 6, 0)-(ucSW - 1, ucSH - 1), bordC, B
    UserControl.CurrentX = (ucSW - (btsW) / 2 + 6 + ucSW - 1 / 2) / 2 - UserControl.TextWidth(">") / 2
    UserControl.CurrentY = ucSH / 2 - UserControl.TextHeight("H") / 2
    UserControl.Print ">"
 
End Function
