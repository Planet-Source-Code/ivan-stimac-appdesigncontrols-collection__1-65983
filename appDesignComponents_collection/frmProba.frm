VERSION 5.00
Begin VB.Form frmProba 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   594
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   635
   StartUpPosition =   3  'Windows Default
   Begin appDesignComponents.Frame Frame3 
      Height          =   5355
      Left            =   2460
      TabIndex        =   12
      Top             =   3000
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   9446
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      BorderColor     =   -2147483630
      BorderColorCaption=   -2147483638
      Begin appDesignComponents.Frame Frame7 
         Height          =   1935
         Left            =   360
         TabIndex        =   13
         Top             =   780
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   3413
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         FillColorCaption=   -2147483633
         BorderColor     =   -2147483635
         BorderColorCaption=   -2147483632
         Caption         =   "Counters"
         Style           =   1
         Begin appDesignComponents.Counter Counter1 
            Height          =   300
            Left            =   600
            TabIndex        =   14
            Top             =   720
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483633
            BackColorButton =   12640511
            BackColorButtonHover=   8438015
            FillColorValue  =   12632319
            ForeColor       =   12582912
            BorderColor     =   -2147483636
         End
         Begin appDesignComponents.Counter Counter2 
            Height          =   300
            Left            =   600
            TabIndex        =   15
            Top             =   1140
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483633
            BackColorButton =   -2147483633
            FillColorValue  =   16777215
            MaxValue        =   100
         End
      End
   End
   Begin appDesignComponents.Frame Frame1 
      Height          =   1035
      Left            =   120
      TabIndex        =   1
      Top             =   660
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1826
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      BorderColor     =   -2147483647
      BorderColorCaption=   -2147483638
      Caption         =   "NewEdition style tab"
      Begin appDesignComponents.sTab sTab1 
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   420
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontSelected {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   4210752
         BorderShadowColor=   -2147483633
         BorderShadowColorSelected=   -2147483632
         SpacingTop      =   2
         SpacingTopSelected=   2
         SpacingDown     =   3
         SpacingSides    =   10
         ListIndex       =   1
         Style           =   3
         ListCount       =   2
         List1           =   "Untitled 1"
         List2           =   "Untitled 2"
      End
   End
   Begin appDesignComponents.ImageList imgLST 
      Left            =   900
      Top             =   7620
      _ExtentX        =   953
      _ExtentY        =   953
      ImageCount      =   4
      mImage0         =   "frmProba.frx":0000
      mImage1         =   "frmProba.frx":0552
      mImage2         =   "frmProba.frx":0AA4
      mImage3         =   "frmProba.frx":0FF6
   End
   Begin appDesignComponents.sTab Tab1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontSelected {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   4210752
      BorderShadowColor=   -2147483633
      BorderShadowColorSelected=   -2147483632
      SpacingTop      =   2
      SpacingTopSelected=   0
      SpacingDown     =   3
      SpacingSides    =   10
      OLEDropMode     =   1
      ListCount       =   4
      List1           =   "XP Style Tab"
      List2           =   "Untitled 2"
      List3           =   "Untitled 3"
      List4           =   "appDesignControls Collection"
   End
   Begin appDesignComponents.Frame Frame2 
      Height          =   1035
      Left            =   4800
      TabIndex        =   3
      Top             =   660
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1826
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      BorderColor     =   -2147483647
      BorderColorCaption=   -2147483638
      Caption         =   "NET Style tab"
      Begin appDesignComponents.sTab sTab2 
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   420
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontSelected {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   4210752
         BorderShadowColor=   -2147483633
         BorderShadowColorSelected=   -2147483632
         SpacingTop      =   2
         SpacingTopSelected=   2
         SpacingDown     =   3
         SpacingSides    =   10
         ListIndex       =   1
         Style           =   1
         ListCount       =   2
         List1           =   "Untitled 1"
         List2           =   "Untitled 2"
      End
   End
   Begin appDesignComponents.Frame Frame4 
      Height          =   1035
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1826
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      BorderColor     =   -2147483647
      BorderColorCaption=   -2147483638
      Caption         =   "Professional style tab"
      Begin appDesignComponents.sTab sTab3 
         Height          =   375
         Left            =   0
         TabIndex        =   6
         Top             =   360
         Width           =   4600
         _ExtentX        =   8123
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontSelected {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   4210752
         BorderShadowColor=   -2147483633
         BorderShadowColorSelected=   -2147483632
         SpacingTop      =   2
         SpacingTopSelected=   2
         SpacingDown     =   3
         SpacingSides    =   10
         ListIndex       =   1
         Style           =   2
         ListCount       =   2
         List1           =   "Untitled 1"
         List2           =   "Untitled 2"
      End
   End
   Begin appDesignComponents.Frame Frame5 
      Height          =   1035
      Left            =   4800
      TabIndex        =   7
      Top             =   1800
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1826
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      BorderColor     =   -2147483647
      BorderColorCaption=   -2147483638
      Caption         =   "Rounde style tab"
      Begin appDesignComponents.sTab sTab4 
         Height          =   375
         Left            =   0
         TabIndex        =   8
         Top             =   360
         Width           =   4600
         _ExtentX        =   8123
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontSelected {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorNormal =   -2147483632
         BorderColor     =   4210752
         BorderShadowColor=   -2147483633
         BorderShadowColorSelected=   -2147483632
         SpacingTop      =   2
         SpacingTopSelected=   0
         SpacingDown     =   3
         SpacingSides    =   10
         ListIndex       =   1
         Style           =   4
         ListCount       =   2
         List1           =   "Untitled 1"
         List2           =   "Untitled 2"
      End
   End
   Begin appDesignComponents.Frame Frame6 
      Height          =   5355
      Left            =   7680
      TabIndex        =   9
      Top             =   3000
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   9446
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      BorderColor     =   -2147483636
      BorderColorCaption=   -2147483638
      Caption         =   ""
      Begin appDesignComponents.ImageMenu ImageMenu1 
         Height          =   4635
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   8176
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonBackColor =   12632319
         ButtonBackColorSelected=   8438015
         ButtonBackColorDisabled=   -2147483637
         ButtonBackColorHover=   12640511
         ForeColorSelected=   16777215
         ForeColorDisabled=   -2147483631
         ButtonBorderColor=   0
         ButtonBorderColorSelected=   0
         ButtonBorderColorHover=   -2147483635
         ButtonBorderColorDisabled=   -2147483631
         Shape           =   1
         OLEDropMode     =   1
      End
   End
   Begin appDesignComponents.ItemList ItemList1 
      Height          =   4935
      Left            =   0
      TabIndex        =   11
      Top             =   3300
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   8705
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorHover  =   -2147483638
      BackColorSelected=   -2147483635
      BackColorSeparator=   -2147483628
      ForeColorSelected=   -2147483634
      ForeColorDisabled=   -2147483631
      LineColor       =   -2147483636
      IconWidth       =   32
      IconHeight      =   32
      VAlign          =   1
      HAlign          =   1
      OLEDropMode     =   1
   End
   Begin appDesignComponents.sTab sTab5 
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   8520
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontSelected {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColorHover  =   -2147483634
      BorderColor     =   4210752
      BorderShadowColor=   -2147483633
      BorderShadowColorSelected=   -2147483632
      SpacingTop      =   2
      SpacingTopSelected=   2
      SpacingDown     =   3
      SpacingSides    =   10
      OLEDropMode     =   1
      TabOrientation  =   1
      ListCount       =   4
      List1           =   "XP Style Tab"
      List2           =   "Untitled 2"
      List3           =   "Untitled 3"
      List4           =   "appDesignControls Collection"
   End
End
Attribute VB_Name = "frmProba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit










Private Sub Form_Load()
    'populate item list
    Me.ItemList1.addItem "Item list", Me.imgLST.mImage(0)
    Me.ItemList1.addItem "Documents", Me.imgLST.mImage(1)
    Me.ItemList1.addItem "Refresh", Me.imgLST.mImage(2)
    Me.ItemList1.addItem "Flash", Me.imgLST.mImage(3)
    'disable one item
    Me.ItemList1.DisableItem 2
    
    'populate image menu
    Me.ImageMenu1.addItem "Item list", Me.imgLST.mImage(0)
    Me.ImageMenu1.addItem "Documents", Me.imgLST.mImage(1)
    Me.ImageMenu1.addItem "Refresh", Me.imgLST.mImage(2)
    Me.ImageMenu1.addItem "Flash", Me.imgLST.mImage(3)
    'disable one item
    Me.ImageMenu1.DisableItem 2
    
End Sub

Private Sub ImageMenu1_OLEDragDrop1(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, lstIndex As Integer)
    MsgBox "DragDrop file: " & Data.Files.Item(1) & vbCrLf & "on button index: " & lstIndex
End Sub

Private Sub ItemList1_OLEDragDrop1(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, lstIndex As Integer)
    MsgBox "DragDrop file: " & Data.Files.Item(1) & vbCrLf & "on button index: " & lstIndex
End Sub

Private Sub Tab1_OLEDragDrop1(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, lstIndex As Integer)
    MsgBox "DragDrop file: " & Data.Files.Item(1) & vbCrLf & "on button index: " & lstIndex
End Sub
