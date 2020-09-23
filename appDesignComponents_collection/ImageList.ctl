VERSION 5.00
Begin VB.UserControl ImageList 
   ClientHeight    =   5310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6990
   InvisibleAtRuntime=   -1  'True
   PropertyPages   =   "ImageList.ctx":0000
   ScaleHeight     =   5310
   ScaleWidth      =   6990
   ToolboxBitmap   =   "ImageList.ctx":0028
   Begin VB.PictureBox mPC 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   3300
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   1
      Top             =   1080
      Width           =   675
   End
   Begin VB.PictureBox pcTMP 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   1560
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   0
      Top             =   180
      Width           =   675
   End
   Begin VB.Image itmIMG 
      Height          =   495
      Index           =   0
      Left            =   1080
      Top             =   1440
      Width           =   615
   End
   Begin VB.Image imgPic 
      Height          =   540
      Left            =   0
      Picture         =   "ImageList.ctx":033A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   540
   End
End
Attribute VB_Name = "ImageList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Private imgCount As Integer
Private mskC As OLE_COLOR

Public Function AddImage(ByVal imgImage As StdPicture)
    If imgCount > 0 Then Load itmIMG(imgCount)
    itmIMG(imgCount).Visible = False
    Set itmIMG(imgCount).Picture = imgImage
   '' Set mImages(imgCount) = imgImage
    Set mImage(imgCount) = itmIMG(imgCount).Picture
    'lstImages.Add imgImage
    imgCount = imgCount + 1

End Function

Private Sub pcTMP_Paint()
    pcTMP.Picture = pcTMP.Image
End Sub

Private Sub UserControl_Initialize()
    imgCount = 0
    mskC = vbWhite
End Sub



Private Sub UserControl_Resize()
    UserControl.Width = imgPic.Width
    UserControl.Height = imgPic.Height
End Sub


Public Function RemoveItem(ByVal itmIndex As Integer)
    Dim i As Integer, mP As Integer
    Dim haveItem As Boolean
    haveItem = False
    'lstimages.Remove itmIndex + 1
    For i = 0 To imgCount - 1
        'sort images
        If i = itmIndex Then
            mP = i
            i = i + 1
            haveItem = True
        End If
        If haveItem Then
            itmIMG(mP).Picture = itmIMG(i).Picture
            'mImages(mP) = mImages(i)
            mP = mP + 1
        End If
    Next i
    'unload last item
    If haveItem Then
        Unload itmIMG(itmIMG.Count - 1)
        imgCount = imgCount - 1
    End If
End Function

Public Property Get mImage(ByVal Index As Integer) As StdPicture
    Dim PE As New ascPaintEffects
    Dim mPic As New StdPicture
    Set mImage = itmIMG(Index).Picture
End Property
Public Property Set mImage(ByVal Index As Integer, Image As StdPicture)
    'Set mImages(Index) = Image
    Set itmIMG(Index).Picture = Image
    PropertyChanged "mImage"
End Property

Public Property Get ImageCount() As Integer
    ImageCount = imgCount
End Property
Public Property Let ImageCount(ByVal nV As Integer)
    imgCount = nV
    PropertyChanged "ImageCount"
End Property

Public Property Get MaskColor() As OLE_COLOR
    MaskColor = mskC
End Property
Public Property Let MaskColor(ByVal nC As OLE_COLOR)
    mskC = nC
    PropertyChanged "MaskColor"
End Property

Public Function getOriginalImage(ByVal Index As Integer) As StdPicture
    Set getOriginalImage = itmIMG(Index).Picture
End Function

Public Function Clear()
    Dim i As Integer
    For i = 1 To itmIMG.Count - 1
        Unload itmIMG(i)
    Next i
    imgCount = 0
End Function

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Dim i As Integer
    PropBag.WriteProperty "ImageCount", imgCount, 0
    PropBag.WriteProperty "MaskColor", mskC, vbWhite
    For i = 0 To imgCount - 1
        'PropBag.WriteProperty "mImage" & i, mImages(i), Nothing
        PropBag.WriteProperty "mImage" & i, itmIMG(i).Picture, Nothing
    Next i
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim i As Integer
    imgCount = PropBag.ReadProperty("ImageCount", 0)
    mskC = PropBag.ReadProperty("MaskColor", vbWhite)
    'ReDim mImages(imgCount) As New StdPicture
    For i = 0 To imgCount - 1
        'Set mImages(i) = PropBag.ReadProperty("mImage" & i, Nothing)
        If itmIMG.Count - 1 < i Then Load itmIMG(i)
        Set itmIMG(i).Picture = PropBag.ReadProperty("mImage" & i, Nothing)
    Next i
End Sub
