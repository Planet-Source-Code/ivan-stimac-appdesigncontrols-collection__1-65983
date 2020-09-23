VERSION 5.00
Begin VB.UserControl ItemList 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   238
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "ItemList.ctx":0000
   Begin VB.PictureBox pcTMP2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1320
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2220
      Top             =   1320
   End
   Begin VB.Image imgIc 
      Appearance      =   0  'Flat
      Height          =   240
      Index           =   0
      Left            =   4320
      Picture         =   "ItemList.ctx":0312
      Stretch         =   -1  'True
      Top             =   3060
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "ItemList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'api
'Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
'Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
Private Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Private Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )


'properties
Private icWid As Integer, icHeig As Integer
'captions
Private lstCaptions As New Collection, isAded As Boolean
'disabled items
Private lstDisabled As New Collection
'colors
Private BC As OLE_COLOR, BCHover As OLE_COLOR, BCSelected As OLE_COLOR, BCDisabled As OLE_COLOR, BCSep As OLE_COLOR
Private FC As OLE_COLOR, FCHover As OLE_COLOR, FCSelected As OLE_COLOR, FCDisabled As OLE_COLOR
Private mskC As OLE_COLOR
'show line option, line style and line color
Private showLn As Boolean, lnStyle As Byte, lnC As OLE_COLOR

'vAlign
Public Enum eVertAlign
    Top
    Center
    Down
End Enum
Private mVAlign As eVertAlign

'hAlign
Public Enum eHorAlign
    eLeft
    eRight
End Enum
Private mHAlign As eHorAlign

'items count
Private itmCount As Integer
'selected item and hover item
Private selectedIND As Integer, hoverIND As Integer

'upSpace=SpacingTop, mSpace=spacing top and down
Private mSpace As Integer, upSpace As Integer

Private mEODM As eOLEDM

'events
Public Event Click()
Public Event OLECompleteDrag1(Effect As Long, lstIndex As Integer)
Public Event OLEDragDrop1(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, lstIndex As Integer)
Public Event OLEDragOver1(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer, lstIndex As Integer)
Public Event OLESetData1(Data As DataObject, DataFormat As Integer, lstIndex As Integer)
Public Event OLEStartDrag1(Data As DataObject, AllowedEffects As Long, lstIndex As Integer)

Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)



'-----------------------------properties-----------------------------------------
'--------------------------------------------------------------------------------

'OLEDropMode
Public Property Get OLEDropMode() As eOLEDM
    OLEDropMode = mEODM
End Property
Public Property Let OLEDropMode(ByVal nC As eOLEDM)
    mEODM = nC
    UserControl.OLEDropMode = mEODM
    PropertyChanged "OLEDropMode"
End Property

Public Property Get SpacingTop() As Integer
    SpacingTop = upSpace
End Property
Public Property Let SpacingTop(ByVal nV As Integer)
    upSpace = nV
    reDraw selectedIND
End Property


'--------------------------------------------------------------------------------
Public Property Get List(ByVal Index As Integer) As String
    List = lstCaptions.Item(Index + 1)
End Property
Public Property Let List(ByVal Index As Integer, strLst As String)
    replaceData lstCaptions, Index + 1, strLst
    reDraw selectedIND
End Property

Public Property Get ListIndex() As Integer
    ListIndex = selectedIND
End Property
Public Property Let ListIndex(ByVal Index As Integer)
    reDraw Index
    RaiseEvent Click
End Property

Public Property Get ListCount() As Integer
    ListCount = imgIc.Count
End Property

'----------------------------------------------------------------------------------
'
Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property
Public Property Set Font(ByVal nFON As Font)
    Set UserControl.Font = nFON
    reDraw selectedIND
    PropertyChanged "Font"
End Property
'back color
Public Property Get BackColor() As OLE_COLOR
    BackColor = BC
End Property
Public Property Let BackColor(ByVal nC As OLE_COLOR)
    BC = nC
    UserControl.BackColor = BC
    reDraw selectedIND
    PropertyChanged "BackColor"
End Property
'back color hover
Public Property Get BackColorHover() As OLE_COLOR
    BackColorHover = BCHover
End Property
Public Property Let BackColorHover(ByVal nC As OLE_COLOR)
    BCHover = nC
    reDraw selectedIND
    PropertyChanged "BackColorHover"
End Property
'back color selected
Public Property Get BackColorSelected() As OLE_COLOR
    BackColorSelected = BCSelected
End Property
Public Property Let BackColorSelected(ByVal nC As OLE_COLOR)
    BCSelected = nC
    reDraw selectedIND
    PropertyChanged "BackColorSelected"
End Property
'back color disabled
Public Property Get BackColorDisabled() As OLE_COLOR
    BackColorDisabled = BCDisabled
End Property
Public Property Let BackColorDisabled(ByVal nC As OLE_COLOR)
    BCDisabled = nC
    reDraw selectedIND
    PropertyChanged "BackColorDisabled"
End Property
'back color hover
Public Property Get BackColorSeparator() As OLE_COLOR
    BackColorSeparator = BCSep
End Property
Public Property Let BackColorSeparator(ByVal nC As OLE_COLOR)
    BCSep = nC
    reDraw selectedIND
    PropertyChanged "BackColorSeparator"
End Property
'fore color
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = FC
End Property
Public Property Let ForeColor(ByVal nC As OLE_COLOR)
    FC = nC
    reDraw selectedIND
    PropertyChanged "ForeColor"
End Property
'fore color
Public Property Get ForeColorHover() As OLE_COLOR
    ForeColorHover = FCHover
End Property
Public Property Let ForeColorHover(ByVal nC As OLE_COLOR)
    FCHover = nC
    reDraw selectedIND
    PropertyChanged "ForeColorHover"
End Property
'fore color
Public Property Get ForeColorSelected() As OLE_COLOR
    ForeColorSelected = FCSelected
End Property
Public Property Let ForeColorSelected(ByVal nC As OLE_COLOR)
    FCSelected = nC
    reDraw selectedIND
    PropertyChanged "ForeColorSelected"
End Property
'
Public Property Get ForeColorDisabled() As OLE_COLOR
    ForeColorDisabled = FCDisabled
End Property
Public Property Let ForeColorDisabled(ByVal nC As OLE_COLOR)
    FCDisabled = nC
    reDraw selectedIND
    PropertyChanged "ForeColorDisabled"
End Property
'line color
Public Property Get LineColor() As OLE_COLOR
    LineColor = lnC
End Property
Public Property Let LineColor(ByVal nC As OLE_COLOR)
    lnC = nC
    reDraw selectedIND
    PropertyChanged "LineColor"
End Property
'mask color
Public Property Get MaskColor() As OLE_COLOR
    MaskColor = mskC
End Property
Public Property Let MaskColor(ByVal nC As OLE_COLOR)
    mskC = nC
    reDraw selectedIND
    PropertyChanged "MaskColor"
End Property
'show line
Public Property Get ShowLines() As Boolean
    ShowLines = showLn
End Property
Public Property Let ShowLines(ByVal nV As Boolean)
    showLn = nV
    reDraw selectedIND
    PropertyChanged "ShowLines"
End Property
'line style
Public Property Get LineStyle() As Integer
    LineStyle = lnStyle
End Property
Public Property Let LineStyle(ByVal nV As Integer)
    lnStyle = nV
    reDraw selectedIND
    PropertyChanged "LineStyle"
End Property
'image
Public Property Get IconWidth() As Integer
    IconWidth = icWid
End Property
Public Property Let IconWidth(ByVal nV As Integer)
    icWid = nV
    reDraw selectedIND
    PropertyChanged "IconWidth"
End Property
'
Public Property Get IconHeight() As Integer
    IconHeight = icHeig
End Property
Public Property Let IconHeight(ByVal nV As Integer)
    icHeig = nV
    reDraw selectedIND
    PropertyChanged "IconHeight"
End Property

'
Public Property Get VAlign() As eVertAlign
    VAlign = mVAlign
End Property
Public Property Let VAlign(ByVal nV As eVertAlign)
    mVAlign = nV
    reDraw selectedIND
    PropertyChanged "VAlign"
End Property

'
Public Property Get HAlign() As eHorAlign
    HAlign = mHAlign
End Property
Public Property Let HAlign(ByVal nV As eHorAlign)
    mHAlign = nV
    reDraw selectedIND
    PropertyChanged "HAlign"
End Property

'------------------------functions and subs-------------------------------
'-------------------------------------------------------------------------

'timer: checking is mouse over control
Private Sub Timer1_Timer()
    Dim lpPos As POINTAPI
    Dim lhWnd As Long
    GetCursorPos lpPos
    lhWnd = WindowFromPoint(lpPos.X, lpPos.Y)
    If lhWnd <> UserControl.hWnd And hoverIND >= 0 Then
        'if NOT redraw control
        reDraw selectedIND
    End If
    
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag1(Effect, hoverIND)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop1(Data, Effect, Button, Shift, X, Y, hoverIND)
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver1(Data, Effect, Button, Shift, X, Y, State, hoverIND)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData1(Data, DataFormat, hoverinf)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag1(Data, AllowedEffects, hoverIND)
End Sub

'read properties
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    Set Font = PropBag.ReadProperty("Font", Ambient.Font) 'UserControl.Font
    BC = PropBag.ReadProperty("BackColor", vbButtonFace)
    BCHover = PropBag.ReadProperty("BackColorHover", vbButtonFace)
    BCSelected = PropBag.ReadProperty("BackColorSelected", vbButtonFace)
    BCDisabled = PropBag.ReadProperty("BackColorDisabled", vbButtonFace)
    BCSep = PropBag.ReadProperty("BackColorSeparator", vbButtonFace)
    
    FC = PropBag.ReadProperty("ForeColor", vbBlack)
    FCHover = PropBag.ReadProperty("ForeColorHover", vbBlack)
    FCSelected = PropBag.ReadProperty("ForeColorSelected", vbBlack)
    FCDisabled = PropBag.ReadProperty("ForeColorDisabled", vbBlack)
    
    lnC = PropBag.ReadProperty("LineColor", vbBlack)
    lnStyle = PropBag.ReadProperty("LineStyle", 0)
    showLn = PropBag.ReadProperty("ShowLines", True)
    
    icWid = PropBag.ReadProperty("IconWidth", 16)
    icHeig = PropBag.ReadProperty("IconHeight", 16)
    
    mVAlign = PropBag.ReadProperty("VAlign", Top)
    mHAlign = PropBag.ReadProperty("HAlign", eLeft)
    
    upSpace = PropBag.ReadProperty("SpacingTop", 10)
    
    mskC = PropBag.ReadProperty("MaskColor", vbWhite)
    
    mEODM = PropBag.ReadProperty("OLEDropMode", None)
    UserControl.OLEDropMode = mEODM
    
    UserControl.BackColor = BC
    reDraw selectedIND
End Sub

Private Sub UserControl_Resize()
    reDraw selectedIND
End Sub

'write properties
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next
    PropBag.WriteProperty "Font", UserControl.Font, Ambient.Font
    
    PropBag.WriteProperty "BackColor", BC, vbButtonFace
    PropBag.WriteProperty "BackColorHover", BCHover, vbButtonFace
    PropBag.WriteProperty "BackColorSelected", BCSelected, vbButtonFace
    PropBag.WriteProperty "BackColorDisabled", BCDisabled, vbButtonFace
    PropBag.WriteProperty "BackColorSeparator", BCSep, vbButtonFace
    
    PropBag.WriteProperty "ForeColor", FC, vbBlack
    PropBag.WriteProperty "ForeColorHover", FCHover, vbBlack
    PropBag.WriteProperty "ForeColorSelected", FCSelected, vbBlack
    PropBag.WriteProperty "ForeColorDisabled", FCDisabled, vbBlack
    
    PropBag.WriteProperty "LineColor", lnC, vbBlack
    PropBag.WriteProperty "LineStyle", lnStyle, 0
    PropBag.WriteProperty "ShowLines", showLn, True
    
    
    PropBag.WriteProperty "IconWidth", icWid, 16
    PropBag.WriteProperty "IconHeight", icHeig, 16
    
    PropBag.WriteProperty "VAlign", mVAlign, Top
    PropBag.WriteProperty "HAlign", mHAlign, eLeft
    
    PropBag.WriteProperty "SpacingTop", upSpace, 10
    
    PropBag.WriteProperty "MaskColor", mskC, vbWhite
    PropBag.WriteProperty "OLEDropMode", mEODM, None
End Sub

'click on image
Private Sub imgIc_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index <> selectedIND And isInList(lstDisabled, Index) <> True Then reDraw Index
    RaiseEvent Click
End Sub
'image mouse move
Private Sub imgIc_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index <> hoverIND And isInList(lstDisabled, Index) <> True Then reDraw selectedIND, Index
End Sub

'defaults values
Private Sub UserControl_Initialize()
    icWid = 16
    icHeig = 16
    itmCount = 0
    
    BC = vbButtonFace
    BCHover = vbButtonFace
    BCSelected = vbButtonFace
    BCDisabled = vbButtonFace
    BCSep = vbButtonFace
    
    'lnC = vbButtonShadow
    lnC = vbBlack
    showLn = True
    lnStyle = 0

    FC = vbBlack
    FCHover = vbBlack
    FCSelected = vbBlack
    FCDisabled = &H80000011
    
    mHAlign = eLeft
    
    lstCaptions.Add "ItemList"
    itmCount = 1
    
    mSpace = 5: upSpace = 10
    
    mVAlign = Top
    
    mskC = vbWhite
End Sub

'add item
Public Function addItem(ByVal strCaption As String, ByVal imgIcon As StdPicture)
    'if is not first item, loadin new image
    If isAded = False Then itmCount = 0
    If itmCount > 0 Then
        Load imgIc(imgIc.Count)
    End If
    '
    imgIc(itmCount).Visible = True
    imgIc(itmCount).ZOrder 0
    'image size
    imgIc(itmCount).Stretch = False
    imgIc(itmCount).Width = icWid
    imgIc(itmCount).Height = icHeig

    
    'set image
    Set imgIc(itmCount).Picture = imgIcon
    'imgDisabled(itmCount).Picture = pcTMP.Image
    If isAded = False Then
        replaceData lstCaptions, 1, strCaption
        isAded = True

    Else
        lstCaptions.Add strCaption
    End If
    
    itmCount = itmCount + 1
    reDraw selectedIND
End Function

'draw control
Private Function reDraw(Optional ByVal sIndex As Integer = -1, Optional hIndex As Integer = -1)
    If itmCount <= 0 Then Exit Function
    '
    Dim i As Integer, ucSH As Integer, ucSW As Integer
    Dim mY As Integer, startX As Integer
    Dim PE As New ascPaintEffects, mPIC1 As New StdPicture
    '
    If hIndex <> hoverIND And hIndex >= 0 Then Timer1.Enabled = True
    pcTMP2.BackColor = mskC
    '
    UserControl.Cls
    UserControl.BackColor = BC
    ucSH = UserControl.ScaleHeight
    ucSW = UserControl.ScaleWidth
    mY = mSpace + upSpace: startX = 10
    '
    If mVAlign = Center Then
       ' mY = (ucSH + upSpace * 2) / 2 - (mSpace * 2 + imgIc.Count * icHeig + upSpace) / 2 '- (icHeig + mSpace * 2 + upSpace) / 2
        mY = (ucSH + upSpace * 2) / 2 - (mSpace * 2 + icHeig + 1 + upSpace) * imgIc.Count / 2
    ElseIf mVAlign = Down Then
        mY = ucSH - (mSpace * 2 + icHeig + 1 + upSpace) * imgIc.Count
    End If
    mY = mY + 1
    
    'draw first line
    If showLn Then
        UserControl.DrawStyle = lnStyle
        UserControl.Line (0, mY - upSpace - mSpace - 1)-(ucSW, mY - upSpace - mSpace - 1), lnC
    End If
    
    For i = 0 To imgIc.Count - 1
        imgIc(i).Visible = False
            
        If mHAlign = eLeft Then
            imgIc(i).Left = startX
'            imgDisabled(i).Left = startX
'            imgSelected(i).Left = startX
        Else
            imgIc(i).Left = ucSW - startX - icWid
'            imgDisabled(i).Left = ucSW - startX - icWid
'            imgSelected(i).Left = ucSW - startX - icWid
        End If
        imgIc(i).Top = mY
'        imgDisabled(i).Top = mY
'        imgSelected(i).Top = mY
        UserControl.DrawStyle = 0

        '
        If lstCaptions.Item(i + 1) <> "-" Then
            'disabled
            If isInList(lstDisabled, i) = True Then
                UserControl.ForeColor = FCDisabled
                UserControl.Line (0, mY - mSpace - upSpace)-(ucSW, mY + icHeig + mSpace), BCDisabled, BF
                'imgDisabled(i).Visible = True
                'imgSelected(i).Visible = False
                'imgIc(i).Visible = False
            'selected button
            ElseIf sIndex = i Then
                UserControl.ForeColor = FCSelected
                UserControl.Line (0, mY - mSpace - upSpace)-(ucSW, mY + icHeig + mSpace), BCSelected, BF
'                imgDisabled(i).Visible = False
'                imgSelected(i).Visible = True
'                imgIc(i).Visible = False
            'hover
            ElseIf hIndex = i Then
                UserControl.ForeColor = FCHover
                UserControl.Line (0, mY - mSpace - upSpace)-(ucSW, mY + icHeig + mSpace), BCHover, BF
'                imgDisabled(i).Visible = False
'                imgSelected(i).Visible = False
'                imgIc(i).Visible = True
            'normal button
            Else
                UserControl.ForeColor = FC
                UserControl.Line (0, mY - mSpace - upSpace)-(ucSW, mY + icHeig + mSpace), BC, BF
'                imgDisabled(i).Visible = False
'                imgSelected(i).Visible = False
'                imgIc(i).Visible = True
            End If
            'draw line
            If showLn Then
                UserControl.DrawStyle = lnStyle
                UserControl.Line (0, mY + icHeig + mSpace)-(ucSW, mY + icHeig + mSpace), lnC
            End If
            'text position
            If mHAlign = eLeft Then
                UserControl.CurrentX = startX + icWid + mSpace
            Else
                UserControl.CurrentX = ucSW - startX - icWid - mSpace - UserControl.TextWidth(lstCaptions.Item(i + 1))
            End If
            UserControl.CurrentY = mY + icHeig / 2 - UserControl.TextHeight("H") / 2
            'print text
            UserControl.Print lstCaptions.Item(i + 1)
            
            '
            pcTMP2.Cls
            pcTMP2.Width = imgIc(i).Width
            pcTMP2.Height = imgIc(i).Height
            pcTMP2.Picture = imgIc(i).Picture
            '
            Set mPIC1 = PowerResize(pcTMP2.Image, icWid, icHeig)
            pcTMP2.Cls
            pcTMP2.Width = icWid
            pcTMP2.Height = icHeig
            pcTMP2.Picture = mPIC1
            If isInList(lstDisabled, i) = True Then
                PE.PaintDisabledPictureEx UserControl.hdc, imgIc(i).Left, imgIc(i).Top, icWid, icHeig, pcTMP2.Picture, 0, 0, mskC
                'PE.PaintDisabledPicture UserControl.hdc, pcTMP2.Picture, imgIc(i).Left, imgIc(i).Top, imgIc(i).Width, imgIc(i).Height, 0, 0, mskC
            Else
                PE.PaintTransparentPicture UserControl.hdc, pcTMP2.Picture, imgIc(i).Left, imgIc(i).Top, icWid, icHeig, 0, 0, mskC
            End If
            mY = mY + mSpace * 2 + icHeig + 1 + upSpace
            
        'draw separator
        Else
            'draw separator
            UserControl.Line (0, mY - mSpace - upSpace)-(ucSW, mY + icHeig + mSpace), BCSep, BF
            imgIc(i).Visible = False
            mY = mY + mSpace * 2 + icHeig + 1 + upSpace
            'draw line
            If showLn Then
                UserControl.DrawStyle = lnStyle
                UserControl.Line (0, mY - mSpace - upSpace - 1)-(ucSW, mY - mSpace - upSpace - 1), lnC
            End If
        End If
    Next i
    
    'set new values
    selectedIND = sIndex
    hoverIND = hIndex
End Function
'
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If itmCount <= 0 Then Exit Sub
    Dim i As Integer
    For i = 0 To imgIc.Count - 1
        'if y of mouse is in button range
        If Y >= imgIc(i).Top - mSpace - upSpace And Y <= imgIc(i).Top + mSpace + icHeig Then
            'if is not separator
            If lstCaptions.Item(i + 1) <> "-" And isInList(lstDisabled, i) <> True Then reDraw i
        End If
    Next i
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub
'
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If itmCount <= 0 Then Exit Sub
    Dim i As Integer, hoverButton As Integer
    hoverButton = -1
    For i = 0 To imgIc.Count - 1
        If Y >= imgIc(i).Top - mSpace - upSpace And Y <= imgIc(i).Top + mSpace + icHeig Then
            If lstCaptions.Item(i + 1) <> "-" And isInList(lstDisabled, i) <> True Then
                hoverButton = i
                Exit For
            ElseIf isInList(lstDisabled, i) = True Then
                hoverButton = i
            End If
        End If
    Next i
    
    If hoverButton <> hoverIND Then reDraw selectedIND, hoverButton
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub



Public Function EnableItem(ByVal Index As Integer)
    On Error Resume Next
    lstDisabled.Remove getDataIndex(lstDisabled, Index)
    reDraw selectedIND
End Function

Public Function DisableItem(ByVal Index As Integer)
    On Error Resume Next
    If isInList(lstDisabled, Index) <> True Then lstDisabled.Add Index
    reDraw selectedIND
End Function


