Attribute VB_Name = "Module2"
Public Declare Function GdiTransparentBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Public Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean

Public Type POINTS2D
    x As Long
    y As Long
End Type

Public Declare Function PlgBlt Lib "gdi32" (ByVal hdcDest As Long, _
                        lpPoint As POINTS2D, _
                        ByVal hdcSrc As Long, _
                        ByVal nXSrc As Long, _
                        ByVal nYSrc As Long, _
                        ByVal nWidth As Long, _
                        ByVal nHeight As Long, _
                        ByVal hbmMask As Long, _
                        ByVal xMask As Long, _
                        ByVal yMask As Long) As Long

Public Const NotPI = 3.14159265238 / 180

Public Sub Rotate(ByRef picDestHdc As Long, xPos As Long, yPos As Long, _
                  ByVal Angle As Long, _
                  ByRef picSrcHdc As Long, srcXoffset As Long, srcYoffset As Long, _
                  ByVal srcWidth As Long, ByVal srcHeight As Long)

  '## Rotate - Rotates an image.
  '##
  '## PicDestHdc      = the hDc of the target picturebox (ie. Picture2.hdc )
  '## xPos            = the target coordinates (note that the image will be centered around these
  '## yPos              coordinates).
  '## Angle           = Rotate Angle (0-360)
  '## PicSrcHdc       = The source image to rotate (ie. Picture1.hdc )
  '## srcXoffset      = The offset coordinates within the Source Image to grab.
  '## srcYoffset
  '## srcWidth        = The width/height of the source image to grab.
  '## srcHeight
  '##
  '## Returns: Nothing.

  Dim Points(3) As POINTS2D
  Dim DefPoints(3) As POINTS2D
  Dim ThetS As Single, ThetC As Single
  Dim Ret As Long
    
    'SET LOCAL AXIS / ALIGNMENT
    Points(0).x = -srcWidth * 0.5
    Points(0).y = -srcHeight * 0.5
    
    Points(1).x = Points(0).x + srcWidth
    Points(1).y = Points(0).y
    
    Points(2).x = Points(0).x
    Points(2).y = Points(0).y + srcHeight
    
    'ROTATE AROUND Z-AXIS
    ThetS = Sin(Angle * NotPI)
    ThetC = Cos(Angle * NotPI)
    
    DefPoints(0).x = (Points(0).x * ThetC - Points(0).y * ThetS) + xPos
    DefPoints(0).y = (Points(0).x * ThetS + Points(0).y * ThetC) + yPos
    
    DefPoints(1).x = (Points(1).x * ThetC - Points(1).y * ThetS) + xPos
    DefPoints(1).y = (Points(1).x * ThetS + Points(1).y * ThetC) + yPos
    
    DefPoints(2).x = (Points(2).x * ThetC - Points(2).y * ThetS) + xPos
    DefPoints(2).y = (Points(2).x * ThetS + Points(2).y * ThetC) + yPos
    
    PlgBlt picDestHdc, DefPoints(0), picSrcHdc, srcXoffset, srcYoffset, srcWidth, srcHeight, 0, 0, 0
    
End Sub



