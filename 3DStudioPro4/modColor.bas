Attribute VB_Name = "modColor"
Option Explicit

Public Const sng1Div3     As Single = 0.3333333
Public Const sng1Div255   As Single = 0.0039215

Public Function ColorSet(Red As Integer, Green As Integer, Blue As Integer) As COLORRGB

    ColorSet.R = Red
    ColorSet.G = Green
    ColorSet.B = Blue

End Function

Public Function ColorScale(C1 As COLORRGB, Scalar As Single) As COLORRGB

    ColorScale.R = C1.R * Scalar
    ColorScale.G = C1.G * Scalar
    ColorScale.B = C1.B * Scalar

End Function

Public Function ColorAdd(C1 As COLORRGB, C2 As COLORRGB) As COLORRGB

    ColorAdd.R = C1.R + C2.R
    ColorAdd.G = C1.G + C2.G
    ColorAdd.B = C1.B + C2.B

End Function

Public Function ColorSub(C1 As COLORRGB, C2 As COLORRGB) As COLORRGB

    ColorSub.R = C1.R - C2.R
    ColorSub.G = C1.G - C2.G
    ColorSub.B = C1.B - C2.B

End Function

Public Function ColorPlus(C1 As COLORRGB, Val As Integer) As COLORRGB

    ColorPlus.R = C1.R + Val
    ColorPlus.G = C1.G + Val
    ColorPlus.B = C1.B + Val
    Call ColorLimit(ColorPlus)

End Function

Function ColorInterpolate(C1 As COLORRGB, C2 As COLORRGB, Alpha As Single) As COLORRGB

    ColorInterpolate.R = ((C2.R - C1.R) * Alpha) + C1.R
    ColorInterpolate.G = ((C2.G - C1.G) * Alpha) + C1.G
    ColorInterpolate.B = ((C2.B - C1.B) * Alpha) + C1.B

End Function

Public Sub ColorLimit(C1 As COLORRGB)

    If (C1.R > 255) Then C1.R = 255 Else If (C1.R < 0) Then C1.R = 0
    If (C1.G > 255) Then C1.G = 255 Else If (C1.G < 0) Then C1.G = 0
    If (C1.B > 255) Then C1.B = 255 Else If (C1.B < 0) Then C1.B = 0

End Sub

Public Function ColorAverage(C1 As COLORRGB, C2 As COLORRGB, C3 As COLORRGB) As COLORRGB

    ColorAverage.R = CInt((C1.R + C2.R + C3.R) * sng1Div3)
    ColorAverage.G = CInt((C1.G + C2.G + C3.G) * sng1Div3)
    ColorAverage.B = CInt((C1.B + C2.B + C3.B) * sng1Div3)
    Call ColorLimit(ColorAverage)

End Function

Public Function ColorGray(C1 As COLORRGB) As Integer

    ColorGray = CInt((C1.R + C1.G + C1.B) * sng1Div3)
    If (ColorGray > 255) Then ColorGray = 255
    
End Function

Public Function ColorDiffuse(C1 As COLORRGB, C2 As COLORRGB) As COLORRGB

    ColorDiffuse.R = C1.R * C2.R * sng1Div255
    ColorDiffuse.G = C1.G * C2.G * sng1Div255
    ColorDiffuse.B = C1.B * C2.B * sng1Div255

End Function

Public Function ColorLongToRGB(lColor As Long) As COLORRGB

    ColorLongToRGB.R = (lColor And &HFF&)
    ColorLongToRGB.G = (lColor And &HFF00&) / &H100&
    ColorLongToRGB.B = (lColor And &HFF0000) / &H10000

End Function

Public Function ColorRGBToLong(C1 As COLORRGB) As Long

    ColorRGBToLong = RGB(C1.R, C1.G, C1.B)

End Function

Public Function ColorBGRToLong(C1 As COLORRGB) As Long

    ColorBGRToLong = RGB(C1.B, C1.G, C1.R)

End Function


