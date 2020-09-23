VERSION 5.00
Begin VB.UserControl sTab 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8400
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   238
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   PropertyPages   =   "Tab.ctx":0000
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   560
   ToolboxBitmap   =   "Tab.ctx":0031
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2880
      Top             =   1200
   End
End
Attribute VB_Name = "sTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long

'colors
Private BCControl As OLE_COLOR
Private BCSelected As OLE_COLOR
Private FC  As OLE_COLOR, FCSelected As OLE_COLOR, FCHover As OLE_COLOR
'
Private BC As OLE_COLOR, BDRC As OLE_COLOR, BDRCSelected As OLE_COLOR
Private BDRShadow As OLE_COLOR, BDRShadSelected As OLE_COLOR

'style
Public Enum eStyle
    XP
    NET
    Professional
    NewEdition
    RoundTabs
End Enum
Private mStyle As eStyle

'orjentation
Public Enum eTabOrj
    sTabOrjentationTop
    sTabOrjentationBottom
End Enum
Private meOrij As eTabOrj

'font
Private fntHover As Font, FNT As Font, fntSEL As Font


Private mEODM As eOLEDM

'
'captions
Private lstCaptions As New Collection, tbCnt As Integer
'start x and button width
Private lstX As New Collection, lstButtWid As New Collection

'spacing
Private sideSpacing As Integer, upSpacing As Integer, downSpacing As Integer
Private upSpacingSel As Integer


'selected tab
Private selectedIND As Integer, hoverIND As Integer

'
Private mX As Integer

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


'OLEDropMode
Public Property Get OLEDropMode() As eOLEDM
    OLEDropMode = mEODM
End Property
Public Property Let OLEDropMode(ByVal nC As eOLEDM)
    mEODM = nC
    UserControl.OLEDropMode = mEODM
    PropertyChanged "OLEDropMode"
End Property

'
Public Property Get SpacingTop() As Integer
    SpacingTop = upSpacing
End Property
Public Property Let SpacingTop(ByVal nV As Integer)
    upSpacing = nV
    reDraw selectedIND
End Property
'
Public Property Get SpacingTopSelected() As Integer
    SpacingTopSelected = upSpacingSel
End Property
Public Property Let SpacingTopSelected(ByVal nV As Integer)
    upSpacingSel = nV
    reDraw selectedIND
End Property
Public Property Get SpacingDown() As Integer
    SpacingDown = downSpacing
End Property
Public Property Let SpacingDown(ByVal nV As Integer)
    downSpacing = nV
    reDraw selectedIND
End Property

Public Property Get SpacingSides() As Integer
    SpacingSides = sideSpacing
End Property
Public Property Let SpacingSides(ByVal nV As Integer)
    sideSpacing = nV
    reDraw selectedIND
End Property

'--------------------------------------------------------------------------------
Public Property Get List(ByVal Index As Integer) As String
Attribute List.VB_ProcData.VB_Invoke_Property = "Tabs"
    List = lstCaptions.Item(Index + 1)
End Property
Public Property Let List(ByVal Index As Integer, strLst As String)
    replaceData lstCaptions, Index + 1, strLst
    reDraw selectedIND
End Property

Public Property Get ListIndex() As Integer
Attribute ListIndex.VB_ProcData.VB_Invoke_Property = "Tabs"
    ListIndex = selectedIND
End Property
Public Property Let ListIndex(ByVal Index As Integer)
    If Index >= 0 And Index < tbCnt Then
        reDraw Index
        RaiseEvent Click
    Else
        MsgBox "Invalid property value!", vbCritical
    End If
End Property

Public Property Get ListCount() As Integer
    ListCount = tbCnt
End Property
Public Property Let ListCount(ByVal nV As Integer)
    Dim i As Integer, tmpCNT As Integer
    tbCnt = nV
    tmpCNT = lstCaptions.Count
    If tbCnt > lstCaptions.Count Then
        For i = 1 To tbCnt - lstCaptions.Count
            lstCaptions.Add "Tab " & tmpCNT + i
        Next i
    ElseIf tbCnt < lstCaptions.Count Then
        For i = lstCaptions.Count To tbCnt + 1 Step -1
            lstCaptions.Remove i
        Next i
    End If
    reDraw selectedIND
    PropertyChanged "ListCount"
End Property

'----------------------------------------------------------------------------------
'
Public Property Get Font() As Font
    Set Font = FNT
End Property
Public Property Set Font(ByVal nFON As Font)
    Set FNT = nFON
    Set UserControl.Font = FNT
    reDraw selectedIND
    PropertyChanged "Font"
End Property
'font hover
Public Property Get FontHover() As Font
    Set FontHover = fntHover
End Property
Public Property Set FontHover(ByVal nFON As Font)
    Set fntHover = nFON
    reDraw selectedIND
    PropertyChanged "FontHover"
End Property
'font selected
Public Property Get FontSelected() As Font
    Set FontSelected = fntSEL
End Property
Public Property Set FontSelected(ByVal nFON As Font)
    Set fntSEL = nFON
    reDraw selectedIND
    PropertyChanged "FontSelected"
End Property
'back color
Public Property Get BackColor() As OLE_COLOR
    BackColor = BCControl
End Property
Public Property Let BackColor(ByVal nC As OLE_COLOR)
    BCControl = nC
    UserControl.BackColor = BCControl
    reDraw selectedIND
    PropertyChanged "BackColor"
End Property
'back color
Public Property Get BackColorNormal() As OLE_COLOR
    BackColorNormal = BC
End Property
Public Property Let BackColorNormal(ByVal nC As OLE_COLOR)
    BC = nC
    UserControl.BackColor = BCControl
    reDraw selectedIND
    PropertyChanged "BackColorNormal"
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

'fore color
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = FC
End Property
Public Property Let ForeColor(ByVal nC As OLE_COLOR)
    FC = nC
    reDraw selectedIND
    PropertyChanged "ForeColor"
End Property
'fore color selected
Public Property Get ForeColorSelected() As OLE_COLOR
    ForeColorSelected = FCSelected
End Property
Public Property Let ForeColorSelected(ByVal nC As OLE_COLOR)
    FCSelected = nC
    reDraw selectedIND
    PropertyChanged "ForeColorSelected"
End Property
'fore color hover
Public Property Get ForeColorHover() As OLE_COLOR
    ForeColorHover = FCHover
End Property
Public Property Let ForeColorHover(ByVal nC As OLE_COLOR)
    FCHover = nC
    reDraw selectedIND
    PropertyChanged "ForeColorHover"
End Property
'
'
'Public Property Get ForeColorDisabled() As OLE_COLOR
'    ForeColorDisabled = FCDisabled
'End Property
'Public Property Let ForeColorDisabled(ByVal nC As OLE_COLOR)
'    FCDisabled = nC
'    reDraw selectedIND
'    PropertyChanged "ForeColorDisabled"
'End Property
'
'border color
Public Property Get BorderColor() As OLE_COLOR
    BorderColor = BDRC
End Property
Public Property Let BorderColor(ByVal nC As OLE_COLOR)
    BDRC = nC
    reDraw selectedIND
    PropertyChanged "BorderColor"
End Property
'border shadow color
Public Property Get BorderShadowColor() As OLE_COLOR
    BorderShadowColor = BDRShadow
End Property
Public Property Let BorderShadowColor(ByVal nC As OLE_COLOR)
    BDRShadow = nC
    reDraw selectedIND
    PropertyChanged "BorderShadowColor"
End Property
'border shadow color
Public Property Get BorderShadowColorSelected() As OLE_COLOR
    BorderShadowColorSelected = BDRShadSelected
End Property
Public Property Let BorderShadowColorSelected(ByVal nC As OLE_COLOR)
    BDRShadSelected = nC
    reDraw selectedIND
    PropertyChanged "BorderShadowColorSelected"
End Property
'border color selected
Public Property Get BorderColorSelected() As OLE_COLOR
    BorderColorSelected = BDRCSelected
End Property
Public Property Let BorderColorSelected(ByVal nC As OLE_COLOR)
    BDRCSelected = nC
    reDraw selectedIND
    PropertyChanged "BorderColorSelected"
End Property

Public Property Get Style() As eStyle
    Style = mStyle
End Property
Public Property Let Style(ByVal nS As eStyle)
    mStyle = nS
    reDraw selectedIND
    PropertyChanged "Style"
End Property

'orientation
Public Property Get TabOrientation() As eTabOrj
    TabOrientation = meOrij
End Property
Public Property Let TabOrientation(ByVal nS As eTabOrj)
    meOrij = nS
    reDraw selectedIND
    PropertyChanged "TabOrientation"
End Property


Private Sub Timer1_Timer()
    Dim lpPos As POINTAPI
    Dim lhWnd As Long
    GetCursorPos lpPos
    lhWnd = WindowFromPoint(lpPos.X, lpPos.Y)
    If lhWnd <> UserControl.hWnd And hoverIND >= 0 Then
        'if NOT redraw control
        hoverIND = -1
        reDraw selectedIND
        Timer1.Enabled = False
    End If
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer, mInd As Integer
    
    For i = 1 To lstX.Count - 1
        If X > lstX.Item(i) And X < lstX.Item(i + 1) Then
            'if x is in range of i-1 button then redraw
            If hoverIND <> i - 1 Then
                hoverIND = i - 1
                reDraw selectedIND
                Exit For
            End If
        ElseIf i = lstX.Count - 1 Then
            'if x is in range of last button then redraw
            If X > lstX.Item(i + 1) And X < lstX.Item(i + 1) + lstButtWid.Item(i + 1) Then
                If hoverIND <> i Then
                    hoverIND = i
                    reDraw selectedIND
                End If
            End If
        End If
    Next i
    Timer1.Enabled = True
    RaiseEvent MouseMove(Button, Shift, X, Y)
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

'-----------------------------------------------------------------------------------

'read properties
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim i As Integer
    On Error Resume Next
    Set FNT = PropBag.ReadProperty("Font", Ambient.Font) 'UserControl.Font
    Set fntHover = PropBag.ReadProperty("FontHover", Ambient.Font)
    Set fntSEL = PropBag.ReadProperty("FontSelected", Ambient.Font)
    
    BCControl = PropBag.ReadProperty("BackColor", vbHighlight)
    BC = PropBag.ReadProperty("BackColorNormal", vbHighlight)
    BCSelected = PropBag.ReadProperty("BackColorSelected", vbButtonFace)
    
    FC = PropBag.ReadProperty("ForeColor", vbWhite)
    FCSelected = PropBag.ReadProperty("ForeColorSelected", vbBlack)
    FCHover = PropBag.ReadProperty("ForeColorHover", vbWhite)
    
    BDRC = PropBag.ReadProperty("BorderColor", vbBlack)
    BDRShadow = PropBag.ReadProperty("BorderShadowColor", vbBlack)
    
    BDRCSelected = PropBag.ReadProperty("BorderColorSelected", vbBlack)
    BDRShadSelected = PropBag.ReadProperty("BorderShadowColorSelected", vbBlack)
    
    selectedIND = PropBag.ReadProperty("ListIndex", 0)
    
    tbCnt = PropBag.ReadProperty("ListCount", 1)
    
    For i = 1 To tbCnt
        If lstCaptions.Count < i Then
            lstCaptions.Add PropBag.ReadProperty("List" & i, "")
        Else
            replaceData lstCaptions, i, PropBag.ReadProperty("List" & i, "")
        End If
    Next i
    
    upSpacing = PropBag.ReadProperty("SpacingTop", 0)
    upSpacingSel = PropBag.ReadProperty("SpacingTopSelected", 5)
    downSpacing = PropBag.ReadProperty("SpacingDown", 5)
    sideSpacing = PropBag.ReadProperty("SpacingSides", 5)
    
    mStyle = PropBag.ReadProperty("Style", XP)
    
    mEODM = PropBag.ReadProperty("OLEDropMode", None)
    
    meOrij = PropBag.ReadProperty("TabOrientation", sTabOrjentationTop)

    
    Set UserControl.Font = FNT
    UserControl.BackColor = BCControl
    UserControl.OLEDropMode = mEODM
    reDraw selectedIND
End Sub

Private Sub UserControl_Resize()
    reDraw selectedIND
End Sub

'write properties
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Dim i As Integer
    On Error Resume Next
    PropBag.WriteProperty "Font", FNT, Ambient.Font
    PropBag.WriteProperty "FontHover", fntHover, Ambient.Font
    PropBag.WriteProperty "FontSelected", fntSEL, Ambient.Font
    
    PropBag.WriteProperty "BackColor", BCControl, vbHighlight
    PropBag.WriteProperty "BackColorNormal", BC, vbHighlight
    PropBag.WriteProperty "BackColorSelected", BCSelected, vbButtonFace
    
    PropBag.WriteProperty "ForeColor", FC, vbWhite
    PropBag.WriteProperty "ForeColorSelected", FCSelected, vbBlack
    PropBag.WriteProperty "ForeColorHover", FCHover, vbWhite
    
    PropBag.WriteProperty "BorderColor", BDRC, vbBlack
    PropBag.WriteProperty "BorderShadowColor", BDRShadow, vbBlack
    
    PropBag.WriteProperty "BorderColorSelected", BDRCSelected, vbBlack
    PropBag.WriteProperty "BorderShadowColorSelected", BDRShadSelected, vbBlack
    
    PropBag.WriteProperty "SpacingTop", upSpacing, 0
    PropBag.WriteProperty "SpacingTopSelected", upSpacingSel, 5
    PropBag.WriteProperty "SpacingDown", downSpacing, 5
    PropBag.WriteProperty "SpacingSides", sideSpacing, 5
    
    PropBag.WriteProperty "ListIndex", selectedIND, 0
    
    PropBag.WriteProperty "Style", mStyle, XP
    PropBag.WriteProperty "OLEDropMode", mEODM, None
    PropBag.WriteProperty "TabOrientation", meOrij, sTabOrjentationTop
    
    PropBag.WriteProperty "ListCount", tbCnt, 1
    
    For i = 1 To lstCaptions.Count
        PropBag.WriteProperty "List" & i, lstCaptions.Item(i), ""
    Next i
    
End Sub

Private Sub UserControl_Initialize()
    selectedIND = 0
    hoverIND = -1
    sideSpacing = 10
    upSpacing = 3
    upSpacingSel = 0
    downSpacing = 3
    
    BCControl = vbHighlight
    'BC = BCControl
    BC = vbHighlight
    BCSelected = vbButtonFace
    FC = vbWhite
    FCHover = vbWhite
    FCSelected = vbBlack
    
    BDRC = &H404040
    BDRShadow = vbButtonFace
    
    BDRCSelected = vbBlack
    BDRShadSelected = vbButtonShadow
    
    tbCnt = 1
    
    lstCaptions.Add "Tab 1"
'    lstCaptions.Add "Tab 2"
'    lstCaptions.Add "Tab 3"


    
    Set FNT = UserControl.Font
    Set fntSEL = UserControl.Font
    Set fntHover = UserControl.Font
    
    mStyle = XP
    mEODM = None
    meOrij = sTabOrjentationTop
End Sub


'add item
'Public Function addItem(ByVal strCaption As String)
'    lstCaptions.Add strCaption
'    reDraw selectedIND
'End Function

Public Function reDraw(ByVal sIndex As Integer)
    If lstCaptions.Count = 0 Then Exit Function
    Dim i As Integer, ucSH As Integer, ucSW As Integer
    Dim buttWid As Integer
    
    Dim selX As Integer, selWid As Integer, selIND As Integer
    selectedIND = sIndex
    
    UserControl.Cls
    UserControl.BackColor = BCControl
    ucSH = UserControl.ScaleHeight
    ucSW = UserControl.ScaleWidth
    mX = sideSpacing
    
    'down border
    If meOrij = sTabOrjentationTop Then
        Line (0, ucSH - 1)-(ucSW, ucSH - 1), BDRCSelected
    Else
        Line (0, 0)-(ucSW, 0), BCSelected
        Line (0, 1)-(ucSW, 1), BDRCSelected
    End If
    
    For i = 0 To tbCnt - 1
        'find button (item) width
        If i = sIndex Then
            Set UserControl.Font = fntSEL
        ElseIf i = hoverIND Then
            Set UserControl.Font = fntHover
        Else
            Set UserControl.Font = FNT
        End If
        buttWid = UserControl.TextWidth(lstCaptions.Item(i + 1)) + sideSpacing * 2
        'if button x is not in list then add it
        If i + 1 > lstX.Count Then
            lstX.Add mX
            lstButtWid.Add buttWid + 2 + 10 + sideSpacing * 2 + 6
        'else replace it with new
        Else
            replaceData lstX, i + 1, mX
            replaceData lstButtWid, i + 1, buttWid + 2 + 10 + sideSpacing * 2 + 6
        End If
        
        'if i is selected then rember it
        If i = sIndex Then
            selX = mX
            selWid = buttWid
            selIND = i
        'draw hover button
        ElseIf i = hoverIND Then
            drawButton False, mX, buttWid, i, True
        'else draw normal button
        Else
            drawButton False, mX, buttWid, i
        End If
        'change x for next button
        If mStyle = XP Then
            mX = mX + buttWid + 2 + 10 + sideSpacing * 2 + 6
        ElseIf mStyle = NET Then
            mX = mX + buttWid + sideSpacing * 2
        ElseIf mStyle = Professional Then
            mX = mX + buttWid + sideSpacing * 2
            If i = sIndex Then mX = mX + (mX + sideSpacing * 2 + buttWid - (mX + sideSpacing + buttWid)) * (ucSH - ucSH / 1.5) / (ucSH / 1.5)
        ElseIf mStyle = NewEdition Then
            mX = mX + buttWid + sideSpacing * 2
        ElseIf mStyle = RoundTabs Then
            mX = mX + buttWid + sideSpacing * 2 + 7
        End If
        
    Next i
    'draw selected button
    Set UserControl.Font = fntSEL
    drawButton True, selX, selWid, selIND
End Function
'draw button
Private Function drawButton(ByVal isSelected As Boolean, ByVal startX As Integer, bWid As Integer, ByVal Index As Integer, Optional isHover As Boolean = False)
    Dim ucSH As Integer, ucSW As Integer
    Dim strCap As String
    Dim stepXSel As Single
    
    strCap = lstCaptions.Item(Index + 1)
    
    ucSH = UserControl.ScaleHeight
    ucSW = UserControl.ScaleWidth
    '-----------------------------------------------------------------------------------------
    '-----------------------------------------------------------------------------------------
    '-------------- orientation = TOP --------------------------------------------------------
    '-----------------------------------------------------------------------------------------
    '-----------------------------------------------------------------------------------------
    If meOrij = sTabOrjentationTop Then
        'if drawing selected/active button
        If isSelected Then
            UserControl.ForeColor = FCSelected
            If mStyle = XP Then
                Line (startX, ucSH - 1)-(startX, upSpacingSel + 5), BDRCSelected
                Circle (startX + 4, upSpacingSel + 8), 4, BDRCSelected, 3.14 / 2, 3.14
                Line (startX + 4, upSpacingSel + 4)-(startX + 4 + bWid + sideSpacing * 2, upSpacingSel + 4), BDRCSelected
                Circle (startX + 2 + bWid + sideSpacing * 2, upSpacingSel + 8), 4, BDRCSelected, 0, 3.14 / 2
                Line (startX + 6 + bWid + sideSpacing * 2, upSpacingSel + 8)-(startX + 6 + bWid + sideSpacing * 2 + 10, ucSH - 1), BDRCSelected
                UserControl.FillColor = BCSelected
                ExtFloodFill UserControl.hdc, startX + 2, ucSH - 2, UserControl.Point(startX + 2, ucSH - 2), 1
                'hide line at botton of button
                Line (startX, ucSH - 1)-(startX + 6 + bWid + sideSpacing * 2 + 10, ucSH - 1), BCSelected
                'draw shadow
                Line (startX + 4, upSpacingSel + 4 + 1)-(startX + 4 + bWid + sideSpacing * 2, upSpacingSel + 4 + 1), BDRShadSelected
                'Circle (startX + 2 + bWid + sideSpacing * 2 - 1, upspacingsel + 8 + 1), 4, BDRShadSelected, 0, 3.14 / 2
                Line (startX + 6 + bWid + sideSpacing * 2 - 1, upSpacingSel + 8)-(startX + 6 + bWid + sideSpacing * 2 + 10 - 1, ucSH - 1), BDRShadSelected
            ElseIf mStyle = NET Then
                'MsgBox index & vbCrLf & selectedIND
                Line (startX, ucSH - 1)-(startX + sideSpacing, upSpacingSel + 5), BDRCSelected
                Line (startX + sideSpacing, upSpacingSel + 5)-(startX + sideSpacing + bWid, upSpacingSel + 5), BDRCSelected
                Line (startX + sideSpacing + bWid, upSpacingSel + 5)-(startX + sideSpacing * 2 + bWid, ucSH - 1), BDRCSelected
                'hide line at botton of selected button
                Line (startX, ucSH - 1)-(startX + sideSpacing * 2 + bWid, ucSH - 1), BCSelected
                'draw shadow
                Line (startX + sideSpacing + 1, upSpacingSel + 5 + 1)-(startX + sideSpacing + bWid - 1, upSpacingSel + 5 + 1), BDRShadSelected
                Line (startX + sideSpacing + bWid - 1, upSpacingSel + 5)-(startX + sideSpacing * 2 + bWid - 1, ucSH - 1), BDRShadSelected
                'redraw line
                Line (startX + sideSpacing, upSpacingSel + 5)-(startX + sideSpacing + bWid, upSpacingSel + 5), BDRCSelected
                UserControl.FillColor = BCSelected
                ExtFloodFill UserControl.hdc, startX + bWid - 2, ucSH - 2, UserControl.Point(startX + bWid - 2, ucSH - 2), 1
            ElseIf mStyle = Professional Then
                stepXSel = (startX + sideSpacing * 2 + bWid - (startX + sideSpacing + bWid)) * (ucSH - ucSH / 1.5) / (ucSH / 1.5)
                
                Line (startX, ucSH - 1)-(startX, upSpacingSel + 5), BDRCSelected
                Line (startX, upSpacingSel + 5)-(startX + sideSpacing + bWid, upSpacingSel + 5), BDRCSelected
                Line (startX + sideSpacing + bWid, upSpacingSel + 5)-(startX + sideSpacing * 2 + bWid + stepXSel, ucSH - 1), BDRCSelected
                'Line (startX + sideSpacing * 2 + bWid, ucSH / 1.5)-(startX + sideSpacing * 2 + bWid, ucSH - 1), BDRCSelected
                'draw shadow
                Line (startX + 1, upSpacingSel + 5 + 1)-(startX + sideSpacing + bWid, upSpacingSel + 5 + 1), BDRShadSelected
                Line (startX + sideSpacing + bWid - 1, upSpacingSel + 5)-(startX + sideSpacing * 2 + bWid + stepXSel - 1, ucSH - 1), BDRShadSelected
                'Line (startX + sideSpacing * 2 + bWid - 1, ucSH / 1.5)-(startX + sideSpacing * 2 + bWid - 1, ucSH - 1), BDRShadSelected
                'redraw line
                Line (startX, upSpacingSel + 5)-(startX + sideSpacing + bWid, upSpacingSel + 5), BDRCSelected
                'hide line at botton of selected button
                Line (startX, ucSH - 1)-(startX + sideSpacing * 2 + bWid + stepXSel, ucSH - 1), BCSelected
                UserControl.FillColor = BCSelected
                ExtFloodFill UserControl.hdc, startX + bWid - 2, ucSH - 2, UserControl.Point(startX + bWid - 2, ucSH - 2), 1
            ElseIf mStyle = NewEdition Then
                Line (startX, ucSH - 1)-(startX, upSpacingSel + 5), BDRCSelected
                Line (startX, upSpacingSel + 5)-(startX + sideSpacing + bWid, upSpacingSel + 5), BDRCSelected
                Line (startX + sideSpacing + bWid, upSpacingSel + 5)-(startX + sideSpacing * 2 + bWid, ucSH / 1.5), BDRCSelected
                Line (startX + sideSpacing * 2 + bWid, ucSH / 1.5)-(startX + sideSpacing * 2 + bWid, ucSH - 1), BDRCSelected
                'draw shadow
                Line (startX + 1, upSpacingSel + 5 + 1)-(startX + sideSpacing + bWid, upSpacingSel + 5 + 1), BDRShadSelected
                Line (startX + sideSpacing + bWid - 1, upSpacingSel + 5)-(startX + sideSpacing * 2 + bWid - 1, ucSH / 1.5), BDRShadSelected
                Line (startX + sideSpacing * 2 + bWid - 1, ucSH / 1.5)-(startX + sideSpacing * 2 + bWid - 1, ucSH - 1), BDRShadSelected
                'redraw line
                Line (startX, upSpacingSel + 5)-(startX + sideSpacing + bWid, upSpacingSel + 5), BDRCSelected
                 'hide line at botton of selected button
                Line (startX, ucSH - 1)-(startX + sideSpacing * 2 + bWid, ucSH - 1), BCSelected
                UserControl.FillColor = BCSelected
                ExtFloodFill UserControl.hdc, startX + bWid - 2, ucSH - 2, UserControl.Point(startX + bWid - 2, ucSH - 2), 1
            ElseIf mStyle = RoundTabs Then
                Line (startX, ucSH - 1)-(startX, upSpacingSel + 5), BDRCSelected
                Circle (startX + 4, upSpacingSel + 8), 4, BDRCSelected, 3.14 / 2, 3.14
                Line (startX + 4, upSpacingSel + 4)-(startX + 4 + bWid + sideSpacing * 2, upSpacingSel + 4), BDRCSelected
                Circle (startX + 2 + bWid + sideSpacing * 2, upSpacingSel + 8), 4, BDRCSelected, 0, 3.14 / 2
                Line (startX + 2 + bWid + sideSpacing * 2 + 4, ucSH - 1)-(startX + 2 + bWid + sideSpacing * 2 + 4, upSpacingSel + 5), BDRCSelected
                UserControl.FillColor = BCSelected
                ExtFloodFill UserControl.hdc, startX + 2, ucSH - 2, UserControl.Point(startX + 2, ucSH - 2), 1
                'hide line at botton of button
                Line (startX, ucSH - 1)-(startX + 2 + bWid + sideSpacing * 2 + 4, ucSH - 1), BCSelected
                'draw shadow
                Line (startX + 4, upSpacingSel + 4 + 1)-(startX + 4 + bWid + sideSpacing * 2, upSpacingSel + 4 + 1), BDRShadSelected
                'Line (startX + 2 + bWid + sideSpacing * 2 + 4 - 1, ucSH - 1)-(startX + 2 + bWid + sideSpacing * 2 + 4 - 1, upspacingsel + 5), BDRShadSelected
            End If
        'if drawing normal/inactive or hover button
        Else
            If isHover = True Then
                UserControl.ForeColor = FCHover
            Else
                UserControl.ForeColor = FC
            End If
            
            If mStyle = XP Then
                If Index <> 0 And Index <> selectedIND + 1 Then
                    Line (startX, ucSH - 5)-(startX, upSpacing + 5), BDRC
                    Line (startX + 1, ucSH - 5)-(startX + 1, upSpacing + 5), BDRShadow
                End If
            ElseIf mStyle = NET Then
                Line (startX, ucSH - 2)-(startX + sideSpacing, upSpacing + 5), BDRC
                Line (startX + sideSpacing, upSpacing + 5)-(startX + sideSpacing + bWid, upSpacing + 5), BDRC
                Line (startX + sideSpacing + bWid, upSpacing + 5)-(startX + sideSpacing * 2 + bWid, ucSH - 1), BDRC
                'shadow
                Line (startX + sideSpacing + 1, upSpacing + 5 + 1)-(startX + sideSpacing + bWid - 1, upSpacing + 5 + 1), BDRShadow
                Line (startX + sideSpacing + bWid - 1, upSpacing + 5)-(startX + sideSpacing * 2 + bWid - 1, ucSH - 1), BDRShadow
                'redraw line at top of button becuse there is one pixel of shadow line
                Line (startX + sideSpacing, upSpacing + 5)-(startX + sideSpacing + bWid, upSpacing + 5), BDRC
                UserControl.FillColor = BC
                ExtFloodFill UserControl.hdc, startX + 2, ucSH - 2, UserControl.Point(startX + 2, ucSH - 2), 1
            ElseIf mStyle = Professional Then
                Line (startX, ucSH - 1)-(startX, upSpacing + 5), BDRC
                Line (startX, upSpacing + 5)-(startX + sideSpacing + bWid, upSpacing + 5), BDRC
                If Index <> tbCnt - 1 Then
                    Line (startX + sideSpacing + bWid, upSpacing + 5)-(startX + sideSpacing * 2 + bWid, ucSH / 1.5), BDRC
                    Line (startX + sideSpacing * 2 + bWid, ucSH / 1.5)-(startX + sideSpacing * 2 + bWid, ucSH - 1), BDRC
                    'draw shadow
                    Line (startX + 1, upSpacing + 5 + 1)-(startX + sideSpacing + bWid, upSpacing + 5 + 1), BDRShadow
                    Line (startX + sideSpacing + bWid - 1, upSpacing + 5)-(startX + sideSpacing * 2 + bWid - 1, ucSH / 1.5), BDRShadow
                    Line (startX + sideSpacing * 2 + bWid - 1, ucSH / 1.5)-(startX + sideSpacing * 2 + bWid - 1, ucSH - 1), BDRShadow
                    'redraw line
                    Line (startX, upSpacing + 5)-(startX + sideSpacing + bWid, upSpacing + 5), BDRC
                Else
                    stepXSel = (startX + sideSpacing * 2 + bWid - (startX + sideSpacing + bWid)) * (ucSH - ucSH / 1.5) / (ucSH / 1.5)
                    Line (startX + sideSpacing + bWid, upSpacing + 5)-(startX + sideSpacing * 2 + bWid + stepXSel, ucSH - 1), BDRC
                    'draw shadow
                    Line (startX + 1, upSpacing + 5 + 1)-(startX + sideSpacing + bWid, upSpacing + 5 + 1), BDRShadow
                    Line (startX + sideSpacing + bWid - 1, upSpacing + 5)-(startX + sideSpacing * 2 + bWid + stepXSel - 1, ucSH - 1), BDRShadow
                    'redraw line
                    Line (startX, upSpacing + 5)-(startX + sideSpacing + bWid, upSpacing + 5), BDRC
                    'hide line at botton of selected button
                    Line (startX, ucSH - 1)-(startX + sideSpacing * 2 + bWid + stepXSel, ucSH - 1), BDRC
                End If
                UserControl.FillColor = BC
                ExtFloodFill UserControl.hdc, startX + bWid - 2, ucSH - 2, UserControl.Point(startX + bWid - 2, ucSH - 2), 1
            ElseIf mStyle = NewEdition Then
                Line (startX, ucSH - 1)-(startX, upSpacing + 5), BDRC
                Line (startX, upSpacing + 5)-(startX + sideSpacing + bWid, upSpacing + 5), BDRC
                Line (startX + sideSpacing + bWid, upSpacing + 5)-(startX + sideSpacing * 2 + bWid, ucSH / 1.5), BDRC
                Line (startX + sideSpacing * 2 + bWid, ucSH / 1.5)-(startX + sideSpacing * 2 + bWid, ucSH - 1), BDRC
                'draw shadow
                Line (startX + 1, upSpacing + 5 + 1)-(startX + sideSpacing + bWid, upSpacing + 5 + 1), BDRShadow
                Line (startX + sideSpacing + bWid - 1, upSpacing + 5)-(startX + sideSpacing * 2 + bWid - 1, ucSH / 1.5), BDRShadow
                Line (startX + sideSpacing * 2 + bWid - 1, ucSH / 1.5)-(startX + sideSpacing * 2 + bWid - 1, ucSH - 1), BDRShadow
                'redraw line
                Line (startX, upSpacing + 5)-(startX + sideSpacing + bWid, upSpacing + 5), BDRC
                UserControl.FillColor = BC
                ExtFloodFill UserControl.hdc, startX + bWid - 2, ucSH - 2, UserControl.Point(startX + bWid - 2, ucSH - 2), 1
            ElseIf mStyle = RoundTabs Then
                Line (startX, ucSH - 1)-(startX, upSpacing + 5), BDRC
                Circle (startX + 4, upSpacing + 8), 4, BDRC, 3.14 / 2, 3.14
                Line (startX + 4, upSpacing + 4)-(startX + 4 + bWid + sideSpacing * 2, upSpacing + 4), BDRC
                Circle (startX + 2 + bWid + sideSpacing * 2, upSpacing + 8), 4, BDRC, 0, 3.14 / 2
                Line (startX + 2 + bWid + sideSpacing * 2 + 4, ucSH - 1)-(startX + 2 + bWid + sideSpacing * 2 + 4, upSpacing + 5), BDRC
                UserControl.FillColor = BC
                ExtFloodFill UserControl.hdc, startX + 2, ucSH - 2, UserControl.Point(startX + 2, ucSH - 2), 1
                'hide line at botton of button
                'Line (startX, ucSH - 1)-(startX + 2 + bWid + sideSpacing * 2 + 4, ucSH - 1), BC
                'draw shadow
                Line (startX + 4, upSpacing + 4 + 1)-(startX + 4 + bWid + sideSpacing * 2, upSpacing + 4 + 1), BDRShadow
                'Line (startX + 2 + bWid + sideSpacing * 2 + 4 - 1, ucSH - 1)-(startX + 2 + bWid + sideSpacing * 2 + 4 - 1, upSpacing + 5), BDRShadSelected
            End If
        End If
     '-----------------------------------------------------------------------------------------
    '-----------------------------------------------------------------------------------------
    '-------------- orientation = BOTTOM --------------------------------------------------------
    '-----------------------------------------------------------------------------------------
    '-----------------------------------------------------------------------------------------
    ElseIf meOrij = sTabOrjentationBottom Then
        'if drawing selected/active button
        If isSelected Then
            UserControl.ForeColor = FCSelected
            If mStyle = XP Then
                Line (startX, 1)-(startX, ucSH - upSpacingSel - 5), BDRCSelected
                Circle (startX + 4, ucSH - upSpacingSel - 8), 4, BDRCSelected, 3.14, 3.14 * 1.5
                Line (startX + 4, ucSH - upSpacingSel - 4)-(startX + 4 + bWid + sideSpacing * 2, ucSH - upSpacingSel - 4), BDRCSelected
                Circle (startX + 2 + bWid + sideSpacing * 2, ucSH - upSpacingSel - 8), 4, BDRCSelected, 3.14 * 1.5, 0
                Line (startX + 6 + bWid + sideSpacing * 2, ucSH - upSpacingSel - 8)-(startX + 6 + bWid + sideSpacing * 2 + 10, 1), BDRCSelected
                UserControl.FillColor = BCSelected
                ExtFloodFill UserControl.hdc, startX + 2, 2, UserControl.Point(startX + 2, 2), 1
                'hide line at TOP of button
                Line (startX, 1)-(startX + 6 + bWid + sideSpacing * 2 + 10, 1), BCSelected
                'draw shadow
                Line (startX + 4, ucSH - upSpacingSel - 4 - 1)-(startX + 4 + bWid + sideSpacing * 2, ucSH - upSpacingSel - 4 - 1), BDRShadSelected
                'Circle (startX + 2 + bWid + sideSpacing * 2 - 1, upspacingsel + 8 + 1), 4, BDRShadSelected, 0, 3.14 / 2
                Line (startX + 6 + bWid + sideSpacing * 2 - 1, ucSH - upSpacingSel - 8)-(startX + 6 + bWid + sideSpacing * 2 + 10 - 1, 1), BDRShadSelected
            ElseIf mStyle = NET Then
                'MsgBox index & vbCrLf & selectedIND
                Line (startX, 1)-(startX + sideSpacing, ucSH - upSpacingSel - 5), BDRCSelected
                Line (startX + sideSpacing, ucSH - upSpacingSel - 5)-(startX + sideSpacing + bWid, ucSH - upSpacingSel - 5), BDRCSelected
                Line (startX + sideSpacing + bWid, ucSH - upSpacingSel - 5)-(startX + sideSpacing * 2 + bWid, 1), BDRCSelected
                'hide line at botton of selected button
                Line (startX, 1)-(startX + sideSpacing * 2 + bWid, 1), BCSelected
                'draw shadow
                Line (startX + sideSpacing + 1, ucSH - upSpacingSel - 5 - 1)-(startX + sideSpacing + bWid - 1, ucSH - upSpacingSel - 5 - 1), BDRShadSelected
                Line (startX + sideSpacing + bWid - 1, ucSH - upSpacingSel - 5)-(startX + sideSpacing * 2 + bWid - 1, 1), BDRShadSelected
                'redraw line
                Line (startX + sideSpacing, ucSH - upSpacingSel - 5)-(startX + sideSpacing + bWid, ucSH - upSpacingSel - 5), BDRCSelected
                UserControl.FillColor = BCSelected
                ExtFloodFill UserControl.hdc, startX + bWid - 2, 2, UserControl.Point(startX + bWid - 2, 2), 1
            ElseIf mStyle = Professional Then
                stepXSel = (startX + sideSpacing * 2 + bWid - (startX + sideSpacing + bWid)) * (ucSH - ucSH / 1.5) / (ucSH / 1.5)
                Line (startX, 1)-(startX, ucSH - upSpacingSel - 5), BDRCSelected
                Line (startX, ucSH - upSpacingSel - 5)-(startX + sideSpacing + bWid, ucSH - upSpacingSel - 5), BDRCSelected
                Line (startX + sideSpacing + bWid, ucSH - upSpacingSel - 5)-(startX + sideSpacing * 2 + bWid + stepXSel, 1), BDRCSelected
                'draw shadow
                Line (startX + 1, ucSH - upSpacingSel - 5 - 1)-(startX + sideSpacing + bWid, ucSH - upSpacingSel - 5 - 1), BDRShadSelected
                Line (startX + sideSpacing + bWid - 1, ucSH - upSpacingSel - 5)-(startX + sideSpacing * 2 + bWid + stepXSel - 1, 1), BDRShadSelected
                'redraw line
                Line (startX, ucSH - upSpacingSel - 5)-(startX + sideSpacing + bWid, ucSH - upSpacingSel - 5), BDRCSelected
                'hide line at botton of selected button
                Line (startX, 1)-(startX + sideSpacing * 2 + bWid + stepXSel, 1), BCSelected
                UserControl.FillColor = BCSelected
                ExtFloodFill UserControl.hdc, startX + bWid - 2, 2, UserControl.Point(startX + bWid - 2, 2), 1
            ElseIf mStyle = NewEdition Then
                Line (startX, 1)-(startX, ucSH - upSpacingSel - 5), BDRCSelected
                Line (startX, ucSH - upSpacingSel - 5)-(startX + sideSpacing + bWid, ucSH - upSpacingSel - 5), BDRCSelected
                Line (startX + sideSpacing + bWid, ucSH - upSpacingSel - 5)-(startX + sideSpacing * 2 + bWid, ucSH - ucSH / 1.5), BDRCSelected
                Line (startX + sideSpacing * 2 + bWid, ucSH - ucSH / 1.5)-(startX + sideSpacing * 2 + bWid, 1), BDRCSelected
                'draw shadow
                Line (startX + 1, ucSH - upSpacingSel - 5 - 1)-(startX + sideSpacing + bWid, ucSH - upSpacingSel - 5 - 1), BDRShadSelected
                Line (startX + sideSpacing + bWid - 1, ucSH - upSpacingSel - 5)-(startX + sideSpacing * 2 + bWid - 1, ucSH - ucSH / 1.5), BDRShadSelected
                Line (startX + sideSpacing * 2 + bWid - 1, ucSH - ucSH / 1.5)-(startX + sideSpacing * 2 + bWid - 1, 1), BDRShadSelected
                'redraw line
                Line (startX, ucSH - upSpacingSel - 5)-(startX + sideSpacing + bWid, ucSH - upSpacingSel - 5), BDRCSelected
                 'hide line at botton of selected button
                Line (startX, 1)-(startX + sideSpacing * 2 + bWid, 1), BCSelected
                UserControl.FillColor = BCSelected
                ExtFloodFill UserControl.hdc, startX + bWid - 2, 2, UserControl.Point(startX + bWid - 2, 2), 1
            ElseIf mStyle = RoundTabs Then
                Line (startX, 1)-(startX, ucSH - upSpacingSel - 5), BDRCSelected
                Circle (startX + 4, ucSH - upSpacingSel - 8), 4, BDRCSelected, 3.14, 3.14 * 1.5
                Line (startX + 4, ucSH - upSpacingSel - 4)-(startX + 4 + bWid + sideSpacing * 2, ucSH - upSpacingSel - 4), BDRCSelected
                Circle (startX + 2 + bWid + sideSpacing * 2, ucSH - upSpacingSel - 8), 4, BDRCSelected, 3.14 * 1.5, 0
                Line (startX + 2 + bWid + sideSpacing * 2 + 4, 1)-(startX + 2 + bWid + sideSpacing * 2 + 4, ucSH - upSpacingSel - 5), BDRCSelected
                UserControl.FillColor = BCSelected
                ExtFloodFill UserControl.hdc, startX + bWid - 2, 2, UserControl.Point(startX + bWid - 2, 2), 1
                'hide line at botton of button
                Line (startX, 1)-(startX + 2 + bWid + sideSpacing * 2 + 4, 1), BCSelected
                'draw shadow
                Line (startX + 4, ucSH - upSpacingSel - 4 - 1)-(startX + 4 + bWid + sideSpacing * 2, ucSH - upSpacingSel - 4 - 1), BDRShadSelected
            End If
        'if drawing normal/inactive or hover button
        Else
            If isHover = True Then
                UserControl.ForeColor = FCHover
            Else
                UserControl.ForeColor = FC
            End If
            
            If mStyle = XP Then
                If Index <> 0 And Index <> selectedIND + 1 Then
                    Line (startX, 5)-(startX, ucSH - upSpacing - 5), BDRC
                    Line (startX + 1, 5)-(startX + 1, ucSH - upSpacing - 5), BDRShadow
                End If
            ElseIf mStyle = NET Then
                 'MsgBox index & vbCrLf & selectedIND
                Line (startX, 1)-(startX + sideSpacing, ucSH - upSpacing - 5), BDRC
                Line (startX + sideSpacing, ucSH - upSpacing - 5)-(startX + sideSpacing + bWid, ucSH - upSpacing - 5), BDRC
                Line (startX + sideSpacing + bWid, ucSH - upSpacing - 5)-(startX + sideSpacing * 2 + bWid, 1), BDRC
                'draw shadow
                Line (startX + sideSpacing + 1, ucSH - upSpacing - 5 - 1)-(startX + sideSpacing + bWid - 1, ucSH - upSpacing - 5 - 1), BDRShadow
                Line (startX + sideSpacing + bWid - 1, ucSH - upSpacing - 5)-(startX + sideSpacing * 2 + bWid - 1, 1), BDRShadow
                'redraw line
                Line (startX + sideSpacing, ucSH - upSpacing - 5)-(startX + sideSpacing + bWid, ucSH - upSpacing - 5), BDRC
                UserControl.FillColor = BC
                ExtFloodFill UserControl.hdc, startX + bWid - 2, 2, UserControl.Point(startX + bWid - 2, 2), 1
            ElseIf mStyle = Professional Then
                Line (startX, 1)-(startX, ucSH - upSpacing - 5), BDRC
                Line (startX, ucSH - upSpacing - 5)-(startX + sideSpacing + bWid, ucSH - upSpacing - 5), BDRC
                If Index <> tbCnt - 1 Then
                    Line (startX + sideSpacing + bWid, ucSH - upSpacing - 5)-(startX + sideSpacing * 2 + bWid, ucSH - ucSH / 1.5), BDRC
                    Line (startX + sideSpacing * 2 + bWid, ucSH - ucSH / 1.5)-(startX + sideSpacing * 2 + bWid, 1), BDRC
                    'draw shadow
                    Line (startX + 1, ucSH - upSpacing - 5 - 1)-(startX + sideSpacing + bWid, ucSH - upSpacing - 5 - 1), BDRShadow
                    Line (startX + sideSpacing + bWid - 1, ucSH - upSpacing - 5)-(startX + sideSpacing * 2 + bWid - 1, ucSH - ucSH / 1.5), BDRShadow
                    Line (startX + sideSpacing * 2 + bWid - 1, ucSH - ucSH / 1.5)-(startX + sideSpacing * 2 + bWid - 1, 1), BDRShadow
                    'redraw line
                    Line (startX, ucSH - upSpacing - 5)-(startX + sideSpacing + bWid, ucSH - upSpacing - 5), BDRC
                Else
                    stepXSel = (startX + sideSpacing * 2 + bWid - (startX + sideSpacing + bWid)) * (ucSH - ucSH / 1.5) / (ucSH / 1.5)
                    Line (startX + sideSpacing + bWid, ucSH - upSpacing - 5)-(startX + sideSpacing * 2 + bWid + stepXSel, 1), BDRC
                    'draw shadow
                    Line (startX + 1, ucSH - upSpacing - 5 - 1)-(startX + sideSpacing + bWid, ucSH - upSpacing - 5 - 1), BDRShadow
                    Line (startX + sideSpacing + bWid - 1, ucSH - upSpacing - 5)-(startX + sideSpacing * 2 + bWid + stepXSel - 1, 1), BDRShadow
                    'redraw line
                    Line (startX, ucSH - upSpacing - 5)-(startX + sideSpacing + bWid, ucSH - upSpacing - 5), BDRC
                    'hide line at botton of selected button
                    Line (startX, 1)-(startX + sideSpacing * 2 + bWid + stepXSel, 1), BDRC
                End If
                UserControl.FillColor = BC
                ExtFloodFill UserControl.hdc, startX + bWid - 2, 2, UserControl.Point(startX + bWid - 2, 2), 1
            ElseIf mStyle = NewEdition Then
                Line (startX, 1)-(startX, ucSH - upSpacing - 5), BDRC
                Line (startX, ucSH - upSpacing - 5)-(startX + sideSpacing + bWid, ucSH - upSpacing - 5), BDRC
                Line (startX + sideSpacing + bWid, ucSH - upSpacing - 5)-(startX + sideSpacing * 2 + bWid, ucSH - ucSH / 1.5), BDRC
                Line (startX + sideSpacing * 2 + bWid, ucSH - ucSH / 1.5)-(startX + sideSpacing * 2 + bWid, 1), BDRC
                'draw shadow
'                Line (startX + 1, upSpacing - 5 - 1)-(startX + sideSpacing + bWid, upSpacing - 5 - 1), BDRShadow
'                Line (startX + sideSpacing + bWid - 1, upSpacing - 5)-(startX + sideSpacing * 2 + bWid - 1, ucSH - ucSH / 1.5), BDRShadow
'                Line (startX + sideSpacing * 2 + bWid - 1, ucSH - ucSH / 1.5)-(startX + sideSpacing * 2 + bWid - 1, 1), BDRShadow
                Line (startX, upSpacing - 5)-(startX + sideSpacing + bWid, upSpacing - 5), BDRC

            ElseIf mStyle = RoundTabs Then
                Line (startX, 1)-(startX, ucSH - upSpacing - 5), BDRC
                Circle (startX + 4, ucSH - upSpacing - 8), 4, BDRC, 3.14, 3.14 * 1.5
                Line (startX + 4, ucSH - upSpacing - 4)-(startX + 4 + bWid + sideSpacing * 2, ucSH - upSpacing - 4), BDRC
                Circle (startX + 2 + bWid + sideSpacing * 2, ucSH - upSpacing - 8), 4, BDRC, 3.14 * 1.5, 0
                Line (startX + 2 + bWid + sideSpacing * 2 + 4, 1)-(startX + 2 + bWid + sideSpacing * 2 + 4, ucSH - upSpacing - 5), BDRC
                UserControl.FillColor = BC
                ExtFloodFill UserControl.hdc, startX + bWid - 2, 2, UserControl.Point(startX + bWid - 2, 2), 1
                'hide line at botton of button
                'draw shadow
                Line (startX + 4, ucSH - upSpacing - 4 - 1)-(startX + 4 + bWid + sideSpacing * 2, ucSH - upSpacing - 4 - 1), BDRShadow
            End If
        End If
    End If
    
    'print caption
    If mStyle = XP Then
        UserControl.CurrentX = (startX + startX + 6 + bWid + sideSpacing * 2) / 2 - UserControl.TextWidth(strCap) / 2
    ElseIf mStyle = NET Then
        UserControl.CurrentX = (startX + startX + bWid + sideSpacing * 2) / 2 - UserControl.TextWidth(strCap) / 2
    ElseIf mStyle = Professional Or mStyle = NewEdition Then
        UserControl.CurrentX = (startX + startX + bWid + sideSpacing * 2) / 2 - UserControl.TextWidth(strCap) / 2
    ElseIf mStyle = RoundTabs Then
        UserControl.CurrentX = (startX + startX + bWid + sideSpacing * 2 + 7) / 2 - UserControl.TextWidth(strCap) / 2
    End If
    
    If meOrij = sTabOrjentationTop Then
        UserControl.CurrentY = ucSH - downSpacing - UserControl.TextHeight("H")
    ElseIf meOrij = sTabOrjentationBottom Then
        UserControl.CurrentY = downSpacing
    End If
    UserControl.Print strCap
End Function

'Public Function clear()
'    clearColl lstCaptions
'    lstCaptions.Add "Tab 1"
'End Function

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    
    For i = 1 To lstX.Count - 1
        If X > lstX.Item(i) And X < lstX.Item(i + 1) Then
            'if x is in range of i-1 button then redraw
            reDraw i - 1
        ElseIf i = lstX.Count - 1 Then
            'if x is in range of last button then redraw
            If X > lstX.Item(i + 1) And X < lstX.Item(i + 1) + lstButtWid.Item(i + 1) Then
                reDraw i
            End If
        End If
    Next i
    RaiseEvent MouseDown(Button, Shift, X, Y)
    
End Sub
