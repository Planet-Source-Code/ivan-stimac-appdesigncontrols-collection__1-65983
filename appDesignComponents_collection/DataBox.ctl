VERSION 5.00
Begin VB.UserControl DataBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8235
   FillColor       =   &H000000FF&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Terminal"
      Size            =   9
      Charset         =   255
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   43
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   549
End
Attribute VB_Name = "DataBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'events
Public Event Change()
Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Event GotFocus1()
Public Event LostFocus1()
Public Event EnterFocus1()
Public Event ExitFocus1()
'-------------------------------------------------------------------------------------
'properties
Private BC As OLE_COLOR, BCDisabled As OLE_COLOR
Private FC As OLE_COLOR, FCDisabled As OLE_COLOR
Private FCFormat As OLE_COLOR, FCFormatDisabled As OLE_COLOR

Public Enum eAppearance
    dBoxFlat
    dBox3D
End Enum
Private meAppearance As eAppearance

Public Enum eBdrStyle
    dBoxNone
    dBoxFixedSingle
End Enum
Private meBdrStyle As eBdrStyle

Public Enum eDataType
    dBoxAll
    dBoxOnlyNumbers
    dBoxOnlyCharacters
End Enum
Private meData As eDataType

Public Enum eStartFrom
    dBoxLeft
    dBoxRight
End Enum
Private meStartFrom As eStartFrom

Private charSpacing As Integer

Private ctlLocked As Boolean


'data format ( ##.##.#### or ##/##/#### or #####.## )   '
Private strDataFormat As String
Private strFormatedData As String
'
Private enbl As Boolean

'----------------------------------------------------------------------
'private vars
'like:07062006                  current char ( selected char )
Private strInputData As String, charIndex As Integer
'
Private mSelLen As Integer
'
Private ucHaveFocus As Boolean




'----------------------------------------------------------------
'-------------- properties
'----enums
Public Property Get DataType() As eDataType
    DataType = meData
End Property
Public Property Let DataType(ByVal nV As eDataType)
    meData = nV
    reDraw
    PropertyChanged "DataType"
End Property
'
Public Property Get BorderStyle() As eBdrStyle
    BorderStyle = meBdrStyle
End Property
Public Property Let BorderStyle(ByVal nV As eBdrStyle)
    meBdrStyle = nV
    reDraw
    PropertyChanged "BorderStyle"
End Property
'
Public Property Get Appearance() As eAppearance
    Appearance = meAppearance
End Property
Public Property Let Appearance(ByVal nV As eAppearance)
    meAppearance = nV
    reDraw
    PropertyChanged "Appearance"
End Property
'
Public Property Get TextAlign() As eStartFrom
    TextAlign = meStartFrom
End Property
Public Property Let TextAlign(ByVal nV As eStartFrom)
    meStartFrom = nV
    reDraw
    PropertyChanged "TextAlign"
End Property
'----colors
'
Public Property Get BackColor() As OLE_COLOR
    BackColor = BC
End Property
Public Property Let BackColor(ByVal nV As OLE_COLOR)
    BC = nV
    reDraw
    PropertyChanged "BackColor"
End Property
'
Public Property Get BackColorDisabled() As OLE_COLOR
    BackColorDisabled = BCDisabled
End Property
Public Property Let BackColorDisabled(ByVal nV As OLE_COLOR)
    BCDisabled = nV
    reDraw
    PropertyChanged "BackColorDisabled"
End Property
'
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = FC
End Property
Public Property Let ForeColor(ByVal nV As OLE_COLOR)
    FC = nV
    reDraw
    PropertyChanged "ForeColor"
End Property
'
Public Property Get ForeColorDisabled() As OLE_COLOR
    ForeColorDisabled = FCDisabled
End Property
Public Property Let ForeColorDisabled(ByVal nV As OLE_COLOR)
    FCDisabled = nV
    reDraw
    PropertyChanged "ForeColorDisabled"
End Property
'
Public Property Get ForeColorFormat() As OLE_COLOR
    ForeColorFormat = FCFormat
End Property
Public Property Let ForeColorFormat(ByVal nV As OLE_COLOR)
    FCFormat = nV
    reDraw
    PropertyChanged "ForeColorFormat"
End Property
'
Public Property Get ForeColorFormatDisabled() As OLE_COLOR
    ForeColorFormatDisabled = FCFormatDisabled
End Property
Public Property Let ForeColorFormatDisabled(ByVal nV As OLE_COLOR)
    FCFormatDisabled = nV
    reDraw
    PropertyChanged "ForeColorFormatDisabled"
End Property
'----strings
'
Public Property Get InputFormat() As String
    InputFormat = strDataFormat
End Property
Public Property Let InputFormat(ByVal nV As String)
    strDataFormat = nV
    reDraw
    PropertyChanged "InputFormat"
End Property
'
Public Property Get FormatedData() As String
    FormatedData = strFormatedData
End Property
'----numbers
Public Property Get LetterSpacing() As Integer
    LetterSpacing = charSpacing
End Property
Public Property Let LetterSpacing(ByVal nV As Integer)
    charSpacing = nV
    reDraw
    PropertyChanged "LetterSpacing"
End Property
'
Public Property Get Data(ByVal dataIndex As Integer) As String
    If dataIndex > getDataCount Then
        MsgBox "There is no " & dataIndex & " data count!", vbCritical
    Else
        Data = getDataString(dataIndex)
        
    End If
End Property
Public Property Let Data(ByVal dataIndex As Integer, ByVal nV As String)
    If dataIndex > getDataCount Then
        MsgBox "There is no " & dataIndex & " data count!", vbCritical
    Else
        setDataString dataIndex, nV
    End If
End Property
'----bool
Public Property Get Enabled() As Boolean
    Enabled = enbl
End Property
Public Property Let Enabled(ByVal nV As Boolean)
    enbl = nV
    reDraw
    PropertyChanged "Enabled"
End Property
'
Public Property Get Locked() As Boolean
    Locked = ctlLocked
End Property
Public Property Let Locked(ByVal nV As Boolean)
    ctlLocked = nV
    PropertyChanged "Locked"
End Property
'----font
Public Property Get Font() As StdFont
    Set Font = UserControl.Font
End Property
Public Property Set Font(ByVal nV As StdFont)
    Set UserControl.Font = nV
    reDraw
    PropertyChanged "Font"
End Property
'----------------------------------------------------------------
'-------------- functions
Private Function reDraw()
    Dim mX As Integer, mY As Integer, ucSH As Integer, ucSW As Integer
    Dim mX1 As Integer
    '                 tmpUIC is selected char from strInputData
    Dim i As Integer, z As Integer, tmpUIC As String
    Dim tempString As String, tmpINPData As String, tmpCHIndex As Integer
    
    UserControl.Cls
    strFormatedData = ""
    
    UserControl.Enabled = enbl
    
    UserControl.Appearance = meAppearance
    UserControl.BorderStyle = meBdrStyle
    
    If meStartFrom <> dBoxLeft Then
        If strInputData = "" Then strInputData = String(getCharCount, "-")
    End If

    
    If UserControl.Enabled = True Then
        UserControl.BackColor = BC
    Else
        UserControl.BackColor = BCDisabled
    End If
    
    ucSH = UserControl.ScaleHeight
    ucSW = UserControl.ScaleWidth
    '
    
    
    mY = 5
    If meStartFrom = dBoxLeft Then
        mX = 5
    Else
        mX = ucSW - ((charSpacing + UserControl.TextWidth("W")) * Len(strDataFormat)) - 5
    End If
    z = 0
    For i = 1 To Len(strDataFormat)
        tmpUIC = ""
        If Mid(strDataFormat, i, 1) = "#" Then
            z = z + 1
            If Len(strInputData) >= z Then
                tmpUIC = Mid(strInputData, z, 1)
            End If
            
'            If meStartFrom = dBoxLeft Then
                tmpCHIndex = charIndex
'            Else
'                tmpCHIndex = Len(tmpINPData) - charIndex + 1
'            End If
            
            
            'draw highlight text
            If z = tmpCHIndex And ucHaveFocus = True Then
                UserControl.ForeColor = vbHighlightText
                UserControl.Line (mX - 1, mY)-(mX + UserControl.TextWidth(tmpUIC), mY + UserControl.TextHeight("H")), vbHighlight, BF
            Else
                If UserControl.Enabled = True Then
                    UserControl.ForeColor = FC
                Else
                    UserControl.ForeColor = FCDisabled
                End If
            End If
                
            UserControl.CurrentY = mY
            UserControl.CurrentX = mX
            mX1 = mX

            
            If tmpUIC <> "" And tmpUIC <> "-" Then
                UserControl.Print tmpUIC
                strFormatedData = strFormatedData & tmpUIC
                mX = mX + UserControl.TextWidth(tmpUIC) + charSpacing
            Else
                mX = mX + UserControl.TextWidth("W") + charSpacing
                strFormatedData = strFormatedData & "_"
            End If
            
            
            
            UserControl.Line (mX1, mY + UserControl.TextHeight("H") + charSpacing)-(mX - charSpacing, mY + UserControl.TextHeight("H") + charSpacing), FC
            
        Else
            If UserControl.Enabled = True Then
                UserControl.ForeColor = FCFormat
            Else
                UserControl.ForeColor = FCFormatDisabled
            End If
            UserControl.CurrentY = mY
            UserControl.CurrentX = mX
            UserControl.Print Mid(strDataFormat, i, 1)
            strFormatedData = strFormatedData & Mid(strDataFormat, i, 1)
            mX = mX + UserControl.TextWidth(Mid(strDataFormat, i, 1)) + charSpacing
        End If
    Next i
    
End Function


Private Function getCharCount() As Integer
    Dim i As Integer, z As Integer
    z = 0
    If strDataFormat = "" Then
        getCharCount = -1
    Else
        For i = 1 To Len(strDataFormat)
            If Mid(strDataFormat, i, 1) = "#" Then z = z + 1
        Next i
    End If
    If z = 0 Then getCharCount = -1 Else getCharCount = z
End Function

Private Function getDataCount() As Integer
    Dim i As Integer, z As Integer
    Dim isNewData As Boolean
    z = 0: isNewData = False
    If strDataFormat = "" Then
        getDataCount = -1
    Else
        For i = 1 To Len(strDataFormat)
            If Mid(strDataFormat, i, 1) = "#" Then
                If isNewData = False Then
                    isNewData = True
                    z = z + 1
                End If
            ElseIf isNewData = True Then
                isNewData = False
            End If
        Next i
    End If
    getDataCount = z
End Function

Private Function getDataString(ByVal dataIndex As Integer) As String
    Dim i As Integer, z As Integer, k As Integer, strTMP As String
    Dim isNewData As Boolean
    z = 0: k = 0: isNewData = False: strTMP = ""
    
    If strDataFormat = "" Then
        getDataString = ""
    Else
        For i = 1 To Len(strDataFormat)
            If Mid(strDataFormat, i, 1) = "#" Then
                If isNewData = False Then
                    isNewData = True
                    z = z + 1
                End If
                k = k + 1
                If z = dataIndex Then strTMP = strTMP & Mid(strInputData, k, 1)

            ElseIf isNewData = True Then
                isNewData = False
            End If
        Next i
    End If
'    If meStartFrom = dBoxLeft Then
        getDataString = strTMP
'    Else
'        getDataString = rewString(strTMP)
'    End If
End Function

Private Function setDataString(ByVal dataIndex As Integer, ByVal nString As String)
    Dim i As Integer, z As Integer, k As Integer, strTMP As String
    Dim isNewData As Boolean
    z = 0: k = 0: isNewData = False: strTMP = ""
    Dim dataStart As Integer, dataLen As Integer
    
    dataStart = 0: dataLen = 0
    If strDataFormat = "" Then
        Exit Function
    Else
        For i = 1 To Len(strDataFormat)
            If Mid(strDataFormat, i, 1) = "#" Then
                If isNewData = False Then
                    isNewData = True
                    z = z + 1
                End If
                k = k + 1
                If z = dataIndex Then
                    If dataStart = 0 Then dataStart = k
                    dataLen = dataLen + 1
                End If

            ElseIf isNewData = True Then
                isNewData = False
                'Exit For
            End If
        Next i
    End If
    If dataLen > nString Then
        nString = nString & String(dataLen - Len(nString), " ")
    ElseIf dataLen < nString Then
        nString = Mid(nString, 1, dataLen)
    End If
    
    If dataLen <> 0 Then
        If Len(strInputData) < dataStart Then
            If Len(strInputData) < dataStart - 1 Then
                strInputData = strInputData & String(dataStart - 1 - Len(strInputData), " ")
            End If
        End If
        If dataStart = 1 Then
            strInputData = nString & Mid(strInputData, dataLen + 1)
        Else
            strInputData = Mid(strInputData, 1, dataStart - 1) & nString & Mid(strInputData, dataStart + dataLen)
        End If

        reDraw
    End If

End Function

Private Function mDataLen(ByVal dataIndex As Integer) As Integer
    Dim i As Integer, z As Integer, tmpLen As Integer
    Dim isNewData As Boolean
    isNewData = False
    z = 0: tmpLen = 0
    If strDataFormat = "" Then
        Exit Function
    Else
        For i = 1 To Len(strDataFormat)
            If Mid(strDataFormat, i, 1) = "#" Then
                If isNewData = False Then
                    isNewData = True
                    z = z + 1
                End If
                If z = dataIndex Then
                    tmpLen = tmpLen + 1
                End If
            ElseIf isNewData = True Then
                isNewData = False
            End If
        Next i
    End If
    mDataLen = tmpLen
End Function


Private Function rewString(ByVal strString As String) As String
    Dim i As Integer, strTMP As String
    strTMP = ""
    For i = Len(strString) To 1 Step -1
        strTMP = strTMP & Mid(strString, i, 1)
    Next i
    rewString = strTMP
End Function

'----------------------------------------------------------------
'-------------- public functions
Public Function SelStart(ByVal sStart As Integer)
    charIndex = SelStart
    reDraw
End Function

Public Function SetFocus()
    UserControl.SetFocus
End Function

Public Function Clear()
    strInputData = ""
    reDraw
End Function

'----------------------------------------------------------------
'-------------- usercontrol events
Private Sub UserControl_Initialize()
    'if "" then it's control like text box
    strDataFormat = ""
    strInputData = ""
    
    meData = dBoxOnlyNumbers
    meAppearance = dBox3D
    meBdrStyle = dBoxFixedSingle
    meStartFrom = dBoxLeft
    
    BC = vbWhite
    BCDisabled = vbActiveBorder
    FC = vbBlack
    FCDisabled = &H80000011
    
    FCFormat = vbRed
    FCFormatDisabled = &H808080
    ucHaveFocus = False
    charIndex = 1
    mSelLen = 0
    charSpacing = 3
    
    ctlLocked = False
    
    enbl = True
    reDraw
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If ctlLocked = True Then Exit Sub
    Dim isOk As Boolean
    Dim tmpStr As String, i As Integer
    isOk = False
    Select Case KeyCode
        'left
        Case 37
            If charIndex > 1 Then charIndex = charIndex - 1
        'right
        Case 39
            If charIndex < Len(strInputData) Then charIndex = charIndex + 1
        'delete, backspace
        Case 46, 8
            strInputData = Mid(strInputData, 1, charIndex - 1) & Mid(strInputData, charIndex + 1)
            If meStartFrom = dBoxLeft Then
                If charIndex > Len(strInputData) Then charIndex = charIndex - 1
                If charIndex < 1 Then charIndex = 1
            End If

        Case Else
            If meData = dBoxAll Then
                If KeyCode >= 32 Then isOk = True
            ElseIf meData = dBoxOnlyNumbers Then
                If KeyCode >= 48 And KeyCode <= 57 Then isOk = True
            ElseIf meData = dBoxOnlyCharacters Then
                If KeyCode >= 65 And KeyCode <= 90 Or (KeyCode >= 97 And KeyCode <= 122) Then
                    isOk = True
                End If
            End If
    End Select
    
    If isOk = True Then
        If meStartFrom = dBoxRight Then
            tmpStr = ""
            For i = 1 To Len(strInputData)
                If Mid(strInputData, i, 1) <> "-" Then
                    tmpStr = tmpStr & Mid(strInputData, i, 1)
                End If
            Next i
            strInputData = tmpStr
        End If
        'MsgBox strInputData
        
        If charIndex = Len(strInputData) + 1 Then
            strInputData = strInputData & Chr(KeyCode)
        Else
            If charIndex <= 1 Then
                strInputData = Chr(KeyCode) & Mid(strInputData, 2)
            Else
                strInputData = Mid(strInputData, 1, charIndex - 1) & Chr(KeyCode) & Mid(strInputData, charIndex + 1)
            End If
        End If
        '
        If charIndex < getCharCount Then charIndex = charIndex + 1
        If meStartFrom = dBoxRight Then
            If getCharCount > Len(strInputData) Then
                strInputData = String(getCharCount - Len(strInputData), "-") & strInputData
            End If
        End If
        RaiseEvent Change
    End If
    '
    If meStartFrom = dBoxRight Then
        If getCharCount > Len(strInputData) Then
            strInputData = String(getCharCount - Len(strInputData), "-") & strInputData
        End If
    End If
    reDraw
End Sub


Private Sub UserControl_LostFocus()
    ucHaveFocus = False
    reDraw
    RaiseEvent LostFocus1
End Sub

Private Sub UserControl_GotFocus()
    ucHaveFocus = True
    reDraw
    RaiseEvent GotFocus1
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_EnterFocus()
    RaiseEvent EnterFocus1
End Sub

Private Sub UserControl_ExitFocus()
    RaiseEvent ExitFocus1
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer, z As Integer, mX As Integer, mX1 As Integer, ucSW As Integer
    
    ucSW = UserControl.ScaleWidth
    
    If meStartFrom = dBoxLeft Then
        mX = 5
    Else
        mX = ucSW - ((charSpacing + UserControl.TextWidth("W")) * Len(strDataFormat)) - 5
    End If
    For i = 1 To Len(strDataFormat)
        If Mid(strDataFormat, i, 1) = "#" Then
            z = z + 1
            mX1 = mX
            If Mid(strInputData, z, 1) <> "" And Mid(strInputData, z, 1) <> "-" Then
                mX = mX + UserControl.TextWidth(Mid(strInputData, z, 1)) + charSpacing
            Else
                If z - Len(strInputData) = 1 Then
                    mX = mX + UserControl.TextWidth("W") + charSpacing
                End If
            End If
        Else
            If meStartFrom = dBoxLeft Then
                mX = mX + UserControl.TextWidth(Mid(strDataFormat, i, 1)) + charSpacing
            Else
                mX = mX + UserControl.TextWidth(Mid(strDataFormat, i, 1)) + charSpacing
                charIndex = Len(strInputData)
                'z = z - 1
                'Exit For
            End If
        End If
        'if click on char between mx1 and mx then select it
        If X >= mX1 And X <= mX Then
            If z > Len(strInputData) Then
                z = Len(strInputData) + 1
            End If
            charIndex = z
            reDraw
            Exit For
        End If
    Next i
    
    
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    
    BC = PropBag.ReadProperty("BackColor", vbWhite)
    BCDisabled = PropBag.ReadProperty("BackColorDisabled", vbActiveBorder)
    FC = PropBag.ReadProperty("ForeColor", vbBlack)
    FCDisabled = PropBag.ReadProperty("ForeColorDisabled", &H80000011)
    
    FCFormat = PropBag.ReadProperty("ForeColorFormat", vbRed)
    FCFormatDisabled = PropBag.ReadProperty("ForeColorFormatDisabled", &H80000011)
    
    strDataFormat = PropBag.ReadProperty("InputFormat", "")
    meData = PropBag.ReadProperty("DataType", 0)
    meBdrStyle = PropBag.ReadProperty("BorderStyle", 1)
    meAppearance = PropBag.ReadProperty("Appearance", 1)
    
    charSpacing = PropBag.ReadProperty("LetterSpacing", 3)
    
    
    enbl = PropBag.ReadProperty("Enabled", True)
    ctlLocked = PropBag.ReadProperty("Locked", False)
    
    meStartFrom = PropBag.ReadProperty("TextAlign", 0)
    reDraw
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
     PropBag.WriteProperty "Font", UserControl.Font, Ambient.Font
    
    PropBag.WriteProperty "BackColor", BC, vbWhite
    PropBag.WriteProperty "BackColorDisabled", BCDisabled, vbActiveBorder
    PropBag.WriteProperty "ForeColor", FC, vbBlack
    PropBag.WriteProperty "ForeColorDisabled", FCDisabled, &H80000011
    
    PropBag.WriteProperty "ForeColorFormat", FCFormat, vbRed
    PropBag.WriteProperty "ForeColorFormatDisabled", FCFormatDisabled, &H80000011
    
    
    PropBag.WriteProperty "InputFormat", strDataFormat, ""
    PropBag.WriteProperty "DataType", meData, 0
    PropBag.WriteProperty "BorderStyle", meBdrStyle, 1
    PropBag.WriteProperty "Appearance", meAppearance, 1
    
    PropBag.WriteProperty "Enabled", enbl, True
    PropBag.WriteProperty "Locked", ctlLocked, False
    
    PropBag.WriteProperty "TextAlign", meStartFrom, 0
    
    PropBag.WriteProperty "LetterSpacing", charSpacing, 0
End Sub
