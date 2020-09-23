VERSION 5.00
Begin VB.UserControl ImageMenu 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   FillStyle       =   0  'Solid
   MaskColor       =   &H00FFFFFF&
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "ImageMenu.ctx":0000
   Begin VB.PictureBox pcTMP 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3360
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   2
      Top             =   2340
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.PictureBox pcTMP2 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2700
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   1
      Top             =   2940
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2220
      Top             =   1200
   End
   Begin VB.PictureBox pcItm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000011&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   315
      Index           =   0
      Left            =   3540
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   0
      Top             =   1860
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "ImageMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type

'Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
'Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private BC As OLE_COLOR, BDRC As OLE_COLOR
Private itmBC As OLE_COLOR, itmBCHover As OLE_COLOR, itmBCSelected As OLE_COLOR, itmBCDisabled As OLE_COLOR
Private itmFC As OLE_COLOR, itmFCHover As OLE_COLOR, itmFCSelected As OLE_COLOR, itmFCDisabled As OLE_COLOR
Private itmBDRC As OLE_COLOR, itmBDRCSelected As OLE_COLOR, itmBDRCHover As OLE_COLOR, itmBDRCDisabled As OLE_COLOR
Private mskC As OLE_COLOR

Private lstCaptions As New Collection, isAded As Boolean
Private disabledItems As New Collection

Private icWid As Integer, icHeig As Integer
Public Enum eRecType
    eRectangle
    eRoundRect
End Enum
Private mRecType As eRecType

Private mEODM As eOLEDM

Private selectedIND As Integer, hoverIND As Integer

Private mSpacing As Integer

Private PE As New ascPaintEffects
'-------events------------------------
Public Event Click()
Public Event OLECompleteDrag1(Effect As Long, lstIndex As Integer)
Public Event OLEDragDrop1(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, lstIndex As Integer)
Public Event OLEDragOver1(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer, lstIndex As Integer)
Public Event OLESetData1(Data As DataObject, DataFormat As Integer, lstIndex As Integer)
Public Event OLEStartDrag1(Data As DataObject, AllowedEffects As Long, lstIndex As Integer)

Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'-------------------------------------------------------------
Private lastX As Integer, lastY As Integer

'---------------------------------------------------------------
'OLEDropMode
Public Property Get OLEDropMode() As eOLEDM
    OLEDropMode = mEODM
End Property
Public Property Let OLEDropMode(ByVal nC As eOLEDM)
    mEODM = nC
    UserControl.OLEDropMode = mEODM
    PropertyChanged "OLEDropMode"
End Property

Public Property Get List(ByVal Index As Integer) As String
    List = lstCaptions.Item(Index + 1)
End Property
Public Property Let List(ByVal Index As Integer, strLst As String)
    replaceData lstCaptions, Index + 1, strLst
    reDraw
End Property

Public Property Get ListIndex() As Integer
    ListIndex = selectedIND
End Property
Public Property Let ListIndex(ByVal Index As Integer)
    selectedIND = Index
    reDraw
    RaiseEvent Click
End Property

Public Property Get ListCount() As Integer
    ListCount = lstCaptions.Count
End Property

'----------------------------------------------------------------------------------
'
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
'back color butt
Public Property Get ButtonBackColor() As OLE_COLOR
    ButtonBackColor = itmBC
End Property
Public Property Let ButtonBackColor(ByVal nC As OLE_COLOR)
    itmBC = nC
    reDraw
    PropertyChanged "ButtonBackColor"
End Property
'back color hover
Public Property Get ButtonBackColorHover() As OLE_COLOR
    ButtonBackColorHover = itmBCHover
End Property
Public Property Let ButtonBackColorHover(ByVal nC As OLE_COLOR)
    itmBCHover = nC
    reDraw
    PropertyChanged "ButtonBackColorHover"
End Property
'back color selected
Public Property Get ButtonBackColorSelected() As OLE_COLOR
    ButtonBackColorSelected = itmBCSelected
End Property
Public Property Let ButtonBackColorSelected(ByVal nC As OLE_COLOR)
    itmBCSelected = nC
    reDraw
    PropertyChanged "ButtonBackColorSelected"
End Property
'back color disabled
Public Property Get ButtonBackColorDisabled() As OLE_COLOR
    ButtonBackColorDisabled = itmBCDisabled
End Property
Public Property Let ButtonBackColorDisabled(ByVal nC As OLE_COLOR)
    itmBCDisabled = nC
    reDraw
    PropertyChanged "ButtonBackColorDisabled"
End Property

'fore color
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = itmFC
End Property
Public Property Let ForeColor(ByVal nC As OLE_COLOR)
    itmFC = nC
    reDraw
    PropertyChanged "ForeColor"
End Property
'fore color
Public Property Get ForeColorHover() As OLE_COLOR
    ForeColorHover = itmFCHover
End Property
Public Property Let ForeColorHover(ByVal nC As OLE_COLOR)
    itmFCHover = nC
    reDraw
    PropertyChanged "ForeColorHover"
End Property
'fore color
Public Property Get ForeColorSelected() As OLE_COLOR
    ForeColorSelected = itmFCSelected
End Property
Public Property Let ForeColorSelected(ByVal nC As OLE_COLOR)
    itmFCSelected = nC
    reDraw
    PropertyChanged "ForeColorSelected"
End Property
'
Public Property Get ForeColorDisabled() As OLE_COLOR
    ForeColorDisabled = itmFCDisabled
End Property
Public Property Let ForeColorDisabled(ByVal nC As OLE_COLOR)
    itmFCDisabled = nC
    reDraw
    PropertyChanged "ForeColorDisabled"
End Property
'line color
Public Property Get ButtonBorderColor() As OLE_COLOR
    ButtonBorderColor = itmBDRC
End Property
Public Property Let ButtonBorderColor(ByVal nC As OLE_COLOR)
    BDRC = nC
    reDraw
    PropertyChanged "ButtonBorderColor"
End Property
'line color
Public Property Get ButtonBorderColorHover() As OLE_COLOR
    ButtonBorderColorHover = itmBDRCHover
End Property
Public Property Let ButtonBorderColorHover(ByVal nC As OLE_COLOR)
    itmBDRCHover = nC
    reDraw
    PropertyChanged "ButtonBorderColorHover"
End Property
'line color
Public Property Get ButtonBorderColorSelected() As OLE_COLOR
    ButtonBorderColorSelected = itmBDRCSelected
End Property
Public Property Let ButtonBorderColorSelected(ByVal nC As OLE_COLOR)
    itmBDRCSelected = nC
    reDraw
    PropertyChanged "ButtonBorderColorSelected"
End Property
'line color
Public Property Get ButtonBorderColorDisabled() As OLE_COLOR
    ButtonBorderColorDisabled = itmBDRCDisabled
End Property
Public Property Let ButtonBorderColorDisabled(ByVal nC As OLE_COLOR)
    itmBDRCDisabled = nC
    reDraw
    PropertyChanged "ButtonBorderColorDisabled"
End Property

'mask color
Public Property Get MaskColor() As OLE_COLOR
    MaskColor = mskC
End Property
Public Property Let MaskColor(ByVal nC As OLE_COLOR)
    mskC = nC
    reDraw
    PropertyChanged "MaskColor"
End Property


'image
Public Property Get IconWidth() As Integer
    IconWidth = icWid
End Property
Public Property Let IconWidth(ByVal nV As Integer)
    icWid = nV
    reDraw
    PropertyChanged "IconWidth"
End Property
'
Public Property Get IconHeight() As Integer
    IconHeight = icHeig
End Property
Public Property Let IconHeight(ByVal nV As Integer)
    icHeig = nV
    reDraw
    PropertyChanged "IconHeight"
End Property

Public Property Get Shape() As eRecType
    Shape = mRecType
End Property
Public Property Let Shape(ByVal nV As eRecType)
     mRecType = nV
     reDraw
     PropertyChanged "Shape"
End Property


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
    RaiseEvent OLESetData1(Data, DataFormat, hoverIND)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag1(Data, AllowedEffects, hoverIND)
End Sub

'-------------------------------------------------------------

'read properties
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    Set Font = PropBag.ReadProperty("Font", Ambient.Font) 'UserControl.Font
    BC = PropBag.ReadProperty("BackColor", vbButtonFace)
    itmBC = PropBag.ReadProperty("ButtonBackColor")
    itmBCHover = PropBag.ReadProperty("ButtonBackColorHover")
    itmBCSelected = PropBag.ReadProperty("ButtonBackColorSelected")
    itmBCDisabled = PropBag.ReadProperty("ButtonBackColordisabled")
    
    itmFC = PropBag.ReadProperty("ForeColor")
    itmFCHover = PropBag.ReadProperty("ForeColorHover")
    itmFCSelected = PropBag.ReadProperty("ForeColorSelected")
    itmFCDisabled = PropBag.ReadProperty("ForeColorDisabled")
    
    itmBDRC = PropBag.ReadProperty("ButtonBorderColor")
    itmBDRCHover = PropBag.ReadProperty("ButtonBorderColorHover")
    itmBDRCSelected = PropBag.ReadProperty("ButtonBorderColorSelected")
    itmBDRCDisabled = PropBag.ReadProperty("ButtonBorderColorDisabled")
    
    icWid = PropBag.ReadProperty("IconWidth", 32)
    icHeig = PropBag.ReadProperty("IconHeight", 32)
    
    mskC = PropBag.ReadProperty("MaskColor", vbWhite)
    
    mRecType = PropBag.ReadProperty("Shape", 0)
    
    mEODM = PropBag.ReadProperty("OLEDropMode", None)
    UserControl.OLEDropMode = mEODM

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
    
    PropBag.WriteProperty "BackColor", BC, vbButtonFace
    
    PropBag.WriteProperty "ButtonBackColor", itmBC, vbButtonFace
    PropBag.WriteProperty "ButtonBackColorSelected", itmBCSelected, vbButtonFace
    PropBag.WriteProperty "ButtonBackColorDisabled", itmBCDisabled, vbButtonFace
    PropBag.WriteProperty "ButtonBackColorHover", itmBCHover, vbButtonFace
    
    PropBag.WriteProperty "ForeColor", itmFC, vbBlack
    PropBag.WriteProperty "ForeColorHover", itmFCHover, vbBlack
    PropBag.WriteProperty "ForeColorSelected", itmFCSelected, vbBlack
    PropBag.WriteProperty "ForeColorDisabled", itmFCDisabled, vbBlack
    
    PropBag.WriteProperty "ButtonBorderColor", itmBDRC
    PropBag.WriteProperty "ButtonBorderColorSelected", itmBDRCSelected
    PropBag.WriteProperty "ButtonBorderColorHover", itmBDRCHover
    PropBag.WriteProperty "ButtonBorderColorDisabled", itmBDRCDisabled
    
    PropBag.WriteProperty "IconWidth", icWid, 32
    PropBag.WriteProperty "IconHeight", icHeig, 32
    
    PropBag.WriteProperty "MaskColor", mskC, vbWhite
    PropBag.WriteProperty "Shape", mRecType, 0
    
    PropBag.WriteProperty "OLEDropMode", mEODM, None
    
End Sub
Private Sub Timer1_Timer()
    Dim lpPos As POINTAPI
    Dim lhWnd As Long
    GetCursorPos lpPos
    lhWnd = WindowFromPoint(lpPos.X, lpPos.Y)
    If lhWnd <> UserControl.hWnd And hoverIND >= 0 Then
        'if NOT redraw control
        hoverIND = -1
        reDraw
    End If
End Sub

Private Sub UserControl_Initialize()
    BC = vbWhite
    BDRC = vbHighlight
    
    itmBC = vbWhite
    itmBCHover = vbActiveBorder
    itmBCSelected = vbHighlight
    itmBCDisabled = vbWhite
    
    itmFC = vbBlack
    itmFCHover = vbBlack
    itmFCSelected = vbWhite
    itmFCDisabled = &H80000011
    
    itmBDRC = vbBlack
    itmBDRCSelected = vbActiveBorder
    itmBDRCHover = vbHighlight
    itmBDRCDisabled = &H80000011
    
    mskC = vbWhite
    
    lstCaptions.Add "ImageMenu"
    
    
    mSpacing = 5
    
    icWid = 32
    icHeig = 32
    
    mRecType = eRoundRect
    
    selectedIND = 0
    hoverIND = -1
End Sub

'-
Private Function reDraw()
    Dim i As Integer, ucSH As Integer, ucSW As Integer, cRadius As Integer
    Dim mY As Integer, bSize As Integer
    Dim mPIC1 As New StdPicture
    UserControl.Cls
    UserControl.BackColor = BC
    ucSH = UserControl.ScaleHeight
    ucSW = UserControl.ScaleWidth
    mY = mSpacing
    
    bSize = icHeig + UserControl.TextHeight("H") + mSpacing * 3
    
    
    For i = 1 To lstCaptions.Count
        If isInList(disabledItems, i - 1) Then
            UserControl.FillColor = itmBCDisabled
            If mRecType = eRectangle Then
                UserControl.Line (mSpacing, mY)-(ucSW - mSpacing, mY + bSize), itmBDRCDisabled, B
            ElseIf mRecType = eRoundRect Then
                UserControl.FillColor = itmBCDisabled
                UserControl.ForeColor = itmBDRCDisabled
                RoundRect UserControl.hdc, mSpacing, mY, ucSW - mSpacing, mY + bSize, 10, 10
            End If
            UserControl.ForeColor = itmFCDisabled
            
        ElseIf selectedIND = i - 1 Then
            UserControl.FillColor = itmBCSelected
            If mRecType = eRectangle Then
                UserControl.Line (mSpacing, mY)-(ucSW - mSpacing, mY + bSize), itmBDRCSelected, B
            ElseIf mRecType = eRoundRect Then
                UserControl.FillColor = itmBCSelected
                UserControl.ForeColor = itmBDRCSelected
                RoundRect UserControl.hdc, mSpacing, mY, ucSW - mSpacing, mY + bSize, 10, 10
            End If
            UserControl.ForeColor = itmFCSelected
        ElseIf hoverIND = i - 1 Then
            UserControl.FillColor = itmBCHover
            If mRecType = eRectangle Then
                UserControl.Line (mSpacing, mY)-(ucSW - mSpacing, mY + bSize), itmBDRCHover, B
            ElseIf mRecType = eRoundRect Then
                UserControl.FillColor = itmBCHover
                UserControl.ForeColor = itmBDRCHover
                RoundRect UserControl.hdc, mSpacing, mY, ucSW - mSpacing, mY + bSize, 10, 10
            End If
            UserControl.ForeColor = itmFCHover
        Else
            UserControl.FillColor = itmBC
            If mRecType = eRectangle Then
                UserControl.Line (mSpacing, mY)-(ucSW - mSpacing, mY + bSize), itmBDRC, B
            ElseIf mRecType = eRoundRect Then
                UserControl.FillColor = itmBC
                UserControl.ForeColor = itmBDRC
                RoundRect UserControl.hdc, mSpacing, mY, ucSW - mSpacing, mY + bSize, 10, 10
            End If
            UserControl.ForeColor = itmFC
        End If
        
'
'        pcTMP2.Width = icWid
'        pcTMP2.Height = icHeig
'        Set mPIC1 = PowerResize(pcItm(i - 1).Picture, icWid, icHeig)
'        pcTMP2.Picture = mPIC1
        
        'add image to picture box and then use picture box image
        '   because PowerResize can't work with icons
        pcTMP.BackColor = mskC
        pcTMP.Cls
        pcTMP.Width = pcItm(i - 1).Width
        pcTMP.Height = pcItm(i - 1).Height
        pcTMP.Picture = pcItm(i - 1).Picture
            '
        Set mPIC1 = PowerResize(pcTMP.Image, icWid, icHeig)
        pcTMP2.Cls
        pcTMP2.Width = icWid
        pcTMP2.Height = icHeig
        pcTMP2.Picture = mPIC1
        
        If isInList(disabledItems, i - 1) <> True Then
            If lstCaptions.Item(i) <> "" Then
                PE.PaintTransparentPicture UserControl.hdc, pcTMP2.Picture, ucSW / 2 - icWid / 2, mY + mSpacing, icWid, icHeig, , , mskC
            Else
                PE.PaintTransparentPicture UserControl.hdc, pcTMP2.Picture, ucSW / 2 - icWid / 2, mY + bSize / 2 - icHeig / 2, icWid, icHeig, , , mskC
            End If
        Else
            If lstCaptions.Item(i) <> "" Then
                PE.PaintDisabledPicture UserControl.hdc, pcTMP2.Picture, ucSW / 2 - icWid / 2, mY + mSpacing, icWid, icHeig, , , mskC
            Else
                PE.PaintDisabledPicture UserControl.hdc, pcTMP2.Picture, ucSW / 2 - icWid / 2, mY + bSize / 2 - icHeig / 2, icWid, icHeig, , , mskC
            End If
        End If
        UserControl.CurrentY = mY + icWid + mSpacing * 2
        UserControl.CurrentX = ucSW / 2 - UserControl.TextWidth(lstCaptions.Item(i)) / 2
        UserControl.Print lstCaptions.Item(i)
        mY = mY + bSize + mSpacing
    Next i
End Function

'-
Public Function addItem(ByVal strCaption As String, ByVal imgPic As StdPicture)
    If isAded = False Then
        replaceData lstCaptions, 1, strCaption
        isAded = True
    Else
        lstCaptions.Add strCaption
    End If
    If lstCaptions.Count > 1 Then Load pcItm(pcItm.Count)
    
    pcItm(lstCaptions.Count - 1).Visible = False
    pcItm(lstCaptions.Count - 1).Width = icWid
    pcItm(lstCaptions.Count - 1).Height = icHeig
    Set pcItm(lstCaptions.Count - 1).Picture = imgPic
    reDraw
End Function

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer, ucSH As Integer, ucSW As Integer, cRadius As Integer
    Dim mY As Integer, bSize As Integer, mInd As Integer
    ucSH = UserControl.ScaleHeight
    ucSW = UserControl.ScaleWidth
    mY = mSpacing
    
    'button height
    bSize = icHeig + UserControl.TextHeight("H") + mSpacing * 3
    'finding what button is selected
    For i = 1 To lstCaptions.Count
        'if y is in range of i-1 button then redraw
        If Y >= mY And Y <= mY + bSize Then
            mInd = i - 1
            'but first check if button is selected or disabled
            If mInd <> selectedIND And isInList(disabledItems, mInd) = False Then
                selectedIND = mInd
                hoverIND = -1
                reDraw
                Exit For
            End If
        End If
        mY = mY + bSize + mSpacing
    Next i
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer, ucSH As Integer, ucSW As Integer, cRadius As Integer
    Dim mY As Integer, bSize As Integer, mInd As Integer
    'UserControl.Cls
    'UserControl.BackColor = BC
    ucSH = UserControl.ScaleHeight
    ucSW = UserControl.ScaleWidth
    mY = mSpacing
    
    
    
    If X = lastX And Y = lastY Then Exit Sub
    lastX = X
    lastY = Y
    bSize = icHeig + UserControl.TextHeight("H") + mSpacing * 3
    For i = 1 To lstCaptions.Count
        If Y >= mY And Y <= mY + bSize Then
            mInd = i - 1
            If mInd <> selectedIND And mInd <> hoverIND And isInList(disabledItems, mInd) = False Then
                hoverIND = mInd
                reDraw
                Exit For
            Else
                hoverIND = mInd
                Exit For
            End If
        'if hovering over empty space on control
        Else
            If i = lstCaptions.Count Then
                If hoverIND <> -1 Then
                    hoverIND = -1
                    reDraw
                End If
            End If
        End If
        mY = mY + bSize + mSpacing
    Next i
    Timer1.Enabled = True
    
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Public Function EnableItem(ByVal Index As Integer)
    On Error Resume Next
    disabledItems.Remove getDataIndex(disabledItems, Index)
    reDraw
End Function

Public Function DisableItem(ByVal Index As Integer)
    On Error Resume Next
    If isInList(disabledItems, Index) <> True Then disabledItems.Add Index
    reDraw
End Function

