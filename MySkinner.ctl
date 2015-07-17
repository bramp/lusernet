VERSION 5.00
Begin VB.UserControl MySkinner 
   ClientHeight    =   3930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5220
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3930
   ScaleWidth      =   5220
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Skinner"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   960
   End
End
Attribute VB_Name = "MySkinner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'This app skins a form

'local variable(s) to hold property value(s)
Private mvarimgNE As StdPicture 'local copy
Private mvarimgN As StdPicture 'local copy
Private mvarimgNW As StdPicture 'local copy
Private mvarimgE As StdPicture 'local copy
Private mvarimgW As StdPicture 'local copy
Private mvarimgSE As StdPicture 'local copy
Private mvarimgS As StdPicture 'local copy
Private mvarimgSW As StdPicture 'local copy
Private mvarBackColor As OLE_COLOR 'local copy
Private mvarForeColor1 As OLE_COLOR 'local copy
Private mvarForeColor2 As OLE_COLOR 'local copy

Private Declare Function BitBlt Lib "gdi32" (ByVal DestDC As Long, ByVal DestX As Long, ByVal DestY As Long, ByVal Width As Long, ByVal Height As Long, ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Rop As RasterOpConstants) As Long

'Most optimal way to skin a form:
'Log(W*P*Log(2))/Log(2) + 1
'Where W is width of form, p is width of pixel,
'and the solution is how many copys of p should be made before tiling
'Allways round it down

'Public Sub OldRepaint(ByRef target As Form)
'    Dim i As Long
'
'    target.AutoRedraw = True
'    target.BorderStyle = 0
'    target.ScaleMode = vbPixels
'
'    Dim hiMeticPerPixelX As Long
'    Dim hiMeticPerPixelY As Long
'
'    hiMeticPerPixelX = target.ScaleX(1, vbPixels, vbHimetric)
'    hiMeticPerPixelY = target.ScaleY(1, vbPixels, vbHimetric)
'
'    target.Cls
'
'    'Tile Across the top
'    For i = 0 To target.ScaleWidth Step (mvarimgN.Width / hiMeticPerPixelX)
'        target.PaintPicture mvarimgN, i, 0
'    Next i
'
'    'Tile Across the right
'    For i = 0 To target.ScaleHeight Step (mvarimgE.Height / hiMeticPerPixelY)
'        target.PaintPicture mvarimgE, target.ScaleWidth - (mvarimgE.Width / hiMeticPerPixelY), i
'    Next i
'
'    'Tile Across the bottom
'    For i = 0 To target.ScaleWidth Step (mvarimgS.Width / hiMeticPerPixelX)
'        target.PaintPicture mvarimgS, i, target.ScaleHeight - (mvarimgS.Height / hiMeticPerPixelX)
'    Next i
'
'    'Tile Across the left
'    For i = 0 To target.ScaleHeight Step (mvarimgW.Height / hiMeticPerPixelY)
'        target.PaintPicture mvarimgW, 0, i
'    Next i
'
'    'Stick Corners in
'    target.PaintPicture mvarimgNW, 0, 0
'    target.PaintPicture mvarimgNE, target.ScaleWidth - (mvarimgNE.Width / hiMeticPerPixelY), 0
'    target.PaintPicture mvarimgSE, target.ScaleWidth - (mvarimgSE.Width / hiMeticPerPixelY), target.ScaleHeight - (mvarimgSE.Height / hiMeticPerPixelX)
'    target.PaintPicture mvarimgSW, 0, target.ScaleHeight - (mvarimgSW.Height / hiMeticPerPixelX)
'
'    'Make sure everything catchs up
'    'DoEvents
'End Sub

Public Sub Repaint(ByRef target As Form)
    
    target.AutoRedraw = True
    target.BorderStyle = 0
    target.ScaleMode = vbPixels
    
    target.BackColor = mvarBackColor
    target.Cls
            
    'Tile Across the top
    TileX target, mvarimgN, 0
    
    'Tile Across the bottom
    TileX target, mvarimgS, target.ScaleHeight - target.ScaleY(mvarimgS.Height, vbHimetric, vbPixels)
    
    'Tile Across the left
    TileY target, mvarimgW, 0
    
    'Tile Across the right
    TileY target, mvarimgE, target.ScaleWidth - target.ScaleX(mvarimgE.Width, vbHimetric, vbPixels)
    
    'Stick Corners in
    target.PaintPicture mvarimgNW, 0, 0
    target.PaintPicture mvarimgNE, target.ScaleWidth - target.ScaleX(mvarimgNE.Width, vbHimetric, vbPixels), 0
    target.PaintPicture mvarimgSE, target.ScaleWidth - target.ScaleX(mvarimgSE.Width, vbHimetric, vbPixels), target.ScaleHeight - target.ScaleY(mvarimgSE.Height, vbHimetric, vbPixels)
    target.PaintPicture mvarimgSW, 0, target.ScaleHeight - target.ScaleY(mvarimgSW.Height, vbHimetric, vbPixels)
    
End Sub

Public Property Let BackColor(ByVal vData As OLE_COLOR)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.imgSW = Form1
    mvarBackColor = vData
End Property


Public Property Get BackColor() As OLE_COLOR
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.imgSW
    BackColor = mvarBackColor
End Property

Public Property Let ForeColor1(ByVal vData As OLE_COLOR)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.imgSW = Form1
    mvarForeColor1 = vData
End Property


Public Property Get ForeColor1() As OLE_COLOR
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.imgSW
    ForeColor1 = mvarForeColor1
End Property

Public Property Let ForeColor2(ByVal vData As OLE_COLOR)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.imgSW = Form1
    mvarForeColor2 = vData
End Property


Public Property Get ForeColor2() As OLE_COLOR
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.imgSW
    ForeColor2 = mvarForeColor2
End Property

Public Property Set imgSW(ByVal vData As StdPicture)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.imgSW = Form1
    Set mvarimgSW = vData
End Property


Public Property Get imgSW() As StdPicture
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.imgSW
    Set imgSW = mvarimgSW
End Property



Public Property Set imgS(ByVal vData As StdPicture)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.imgS = Form1
    Set mvarimgS = vData
End Property


Public Property Get imgS() As StdPicture
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.imgS
    Set imgS = mvarimgS
End Property



Public Property Set imgSE(ByVal vData As StdPicture)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.imgSE = Form1
    Set mvarimgSE = vData
End Property


Public Property Get imgSE() As StdPicture
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.imgSE
    Set imgSE = mvarimgSE
End Property



Public Property Set imgW(ByVal vData As StdPicture)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.imgW = Form1
    Set mvarimgW = vData
End Property


Public Property Get imgW() As StdPicture
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.imgW
    Set imgW = mvarimgW
End Property



Public Property Set imgE(ByVal vData As StdPicture)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.imgE = Form1
    Set mvarimgE = vData
End Property


Public Property Get imgE() As StdPicture
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.imgE
    Set imgE = mvarimgE
End Property



Public Property Set imgNW(ByVal vData As StdPicture)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.imgNW = Form1
    Set mvarimgNW = vData
End Property


Public Property Get imgNW() As StdPicture
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.imgNW
    Set imgNW = mvarimgNW
End Property



Public Property Set imgN(ByVal vData As StdPicture)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.imgN = Form1
    Set mvarimgN = vData
End Property


Public Property Get imgN() As StdPicture
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.imgN
    Set imgN = mvarimgN
End Property

Public Property Set imgNE(ByVal vData As StdPicture)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.imgNE = Form1
    Set mvarimgNE = vData
End Property


Public Property Get imgNE() As StdPicture
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.imgNE
    Set imgNE = mvarimgNE
End Property

Private Sub UserControl_Resize()
    UserControl.Width = lblLabel.Width
    UserControl.Height = lblLabel.Height
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    Set mvarimgNE = PropBag.ReadProperty("imgNE")
    Set mvarimgN = PropBag.ReadProperty("imgN")
    Set mvarimgNW = PropBag.ReadProperty("imgNW")
    Set mvarimgE = PropBag.ReadProperty("imgE")
    Set mvarimgW = PropBag.ReadProperty("imgW")
    Set mvarimgSE = PropBag.ReadProperty("imgSE")
    Set mvarimgS = PropBag.ReadProperty("imgS")
    Set mvarimgSW = PropBag.ReadProperty("imgSW")
    
    mvarBackColor = PropBag.ReadProperty("BackColor")
    mvarForeColor1 = PropBag.ReadProperty("ForeColor1")
    mvarForeColor2 = PropBag.ReadProperty("ForeColor2")
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "imgNE", mvarimgNE
    PropBag.WriteProperty "imgN", mvarimgN
    PropBag.WriteProperty "imgNW", mvarimgNW
    PropBag.WriteProperty "imgE", mvarimgE
    PropBag.WriteProperty "imgW", mvarimgW
    PropBag.WriteProperty "imgSE", mvarimgSE
    PropBag.WriteProperty "imgS", mvarimgS
    PropBag.WriteProperty "imgSW", mvarimgSW
    
    PropBag.WriteProperty "BackColor", mvarBackColor
    PropBag.WriteProperty "ForeColor1", mvarForeColor1
    PropBag.WriteProperty "ForeColor2", mvarForeColor2
End Sub


Public Sub TileX(Dest As Form, Pict As StdPicture, offset As Long)
Dim lCurW As Long
Dim lCurH As Long
Dim lMaxW As Long
Dim lMaxH As Long

  With Dest

    lMaxW = .ScaleWidth
    lMaxH = .ScaleHeight

    lCurW = .ScaleX(Pict.Width, vbHimetric, vbPixels)
    lCurH = .ScaleY(Pict.Height, vbHimetric, vbPixels)

    'Paints the top-left image
    Call .PaintPicture(Pict, 0, offset)

    'Tiles the first row
    Do While lCurW < lMaxW
      Call BitBlt(.hDC, lCurW, offset, lCurW, lCurH, .hDC, 0, offset, vbSrcCopy)
      lCurW = lCurW + lCurW
    Loop

    'Tiles vertically
'    Do While lCurH < lMaxH
'      Call BitBlt(.hDC, 0, lCurH, lCurW, lCurH, .hDC, 0, 0, vbSrcCopy)
'      lCurH = lCurH + lCurH
'    Loop

  End With

End Sub

Public Sub TileY(Dest As Form, Pict As StdPicture, offset As Long)
Dim lCurW As Long
Dim lCurH As Long
Dim lMaxW As Long
Dim lMaxH As Long

  With Dest

    lMaxW = .ScaleWidth
    lMaxH = .ScaleHeight

    lCurW = .ScaleX(Pict.Width, vbHimetric, vbPixels)
    lCurH = .ScaleY(Pict.Height, vbHimetric, vbPixels)

    'Paints the top-left image
    Call .PaintPicture(Pict, offset, 0)

    'Tiles the first row
'    Do While lCurW < lMaxW
'      Call BitBlt(.hDC, lCurW, 0, lCurW, lCurH, .hDC, 0, 0, vbSrcCopy)
'      lCurW = lCurW + lCurW
'    Loop

    'Tiles vertically
    Do While lCurH < lMaxH
      Call BitBlt(.hDC, offset, lCurH, lCurW, lCurH, .hDC, offset, 0, vbSrcCopy)
      lCurH = lCurH + lCurH
    Loop

  End With

End Sub

