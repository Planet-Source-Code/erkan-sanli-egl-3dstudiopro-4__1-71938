Attribute VB_Name = "modHSL"
Option Explicit

Private Const dbl1Div240 As Double = 0.004166666667
Private Const sng1Div40 As Single = 0.025

Public Type COLORHSL
    Hue     As Integer
    Sat     As Integer
    Lum     As Integer
End Type

Public Enum ColorType
    tRGB
    tHSL
End Enum

Public Function RGBtoHSL(RGBCol As COLORRGB) As COLORHSL
    
    Dim R       As Integer
    Dim G       As Integer
    Dim B       As Integer
    Dim H       As Double
    Dim S       As Double
    Dim L       As Double
    Dim cMax    As Integer
    Dim cMin    As Integer
    Dim RDelta  As Double
    Dim GDelta  As Double
    Dim BDelta  As Double
    Dim cMinus  As Long
    Dim cPlus   As Long

    R = RGBCol.R
    G = RGBCol.G
    B = RGBCol.B
    cMax = iMax(iMax(R, G), B)                          'Highest and lowest
    cMin = iMin(iMin(R, G), B)                          'color values
    cMinus = cMax - cMin                                'Used to simplify the
    cPlus = cMax + cMin                                 'calculations somewhat.
    L = ((cPlus * 240) + 255) / 510                     'Luminance
    If cMax = cMin Then                                 'Greyscale
        S = 0
        H = 160
    Else
        If L <= 120 Then                                'Saturation
            S = ((cMinus * 240) + 0.5) / cPlus
        Else
            S = ((cMinus * 240) + 0.5) / (510 - cPlus)
        End If
        RDelta = (((cMax - R) * 40) + 0.5) / cMinus 'Hue
        GDelta = (((cMax - G) * 40) + 0.5) / cMinus
        BDelta = (((cMax - B) * 40) + 0.5) / cMinus
        Select Case cMax
            Case CLng(R)
                H = BDelta - GDelta
            Case CLng(G)
                H = 80 + RDelta - BDelta
            Case CLng(B)
                H = 160 + GDelta - RDelta
        End Select
        If H < 0 Then H = H + 240
    End If
    RGBtoHSL.Hue = CInt(H)
    RGBtoHSL.Lum = CInt(L)
    RGBtoHSL.Sat = CInt(S)

End Function

Public Function HSLtoRGB(HueLumSat As COLORHSL) As COLORRGB

    Dim R       As Double
    Dim G       As Double
    Dim B       As Double
    Dim H       As Double
    Dim L       As Double
    Dim S       As Double
    Dim Magic1  As Double
    Dim Magic2  As Double

    H = HueLumSat.Hue
    L = HueLumSat.Lum
    S = HueLumSat.Sat
'Greyscale
    If HueLumSat.Sat = 0 Then
        R = L * 1.0625
        G = R
        B = R
    Else
        If L <= 120 Then
            Magic2 = (L * (240 + S) + 0.5) * dbl1Div240
        Else
            Magic2 = L + S - ((L * S) + 0.5) * dbl1Div240
        End If
        Magic1 = 2 * L - Magic2
        R = (HUEtoRGB(Magic1, Magic2, H + 80) * 255 + 0.5) * dbl1Div240
        G = (HUEtoRGB(Magic1, Magic2, H) * 255 + 0.5) * dbl1Div240
        B = (HUEtoRGB(Magic1, Magic2, H - 80) * 255 + 0.5) * dbl1Div240
    End If
    
    HSLtoRGB = ColorSet(CInt(R), CInt(G), CInt(B))
    Call ColorLimit(HSLtoRGB)
End Function

Private Function HUEtoRGB(mag1 As Double, mag2 As Double, ByVal Hue As Double) As Double

    If Hue < 0 Then Hue = Hue + 240 Else If Hue > 240 Then Hue = Hue - 240
    Select Case Hue
        Case Is < 40
            HUEtoRGB = mag1 + 20 + sng1Div40 * (mag2 - mag1) * Hue
        Case Is < 120
            HUEtoRGB = mag2
        Case Is < 160
            HUEtoRGB = mag1 + 20 + sng1Div40 * (mag2 - mag1) * (160 - Hue)
        Case Else
            HUEtoRGB = mag1
    End Select

End Function

Private Function iMax(A As Integer, B As Integer) As Integer

    iMax = IIf(A > B, A, B)

End Function

Private Function iMin(A As Integer, B As Integer) As Integer

    iMin = IIf(A < B, A, B)

End Function

Public Function ColorLimits(ByVal iColor As Integer) As Integer

    If iColor < 0 Then
        ColorLimits = 0
    ElseIf iColor > 255 Then
        ColorLimits = 255
    Else
        ColorLimits = iColor
    End If
    
End Function

