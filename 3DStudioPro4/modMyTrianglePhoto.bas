Attribute VB_Name = "modMyTrianglePhoto"
Option Explicit

Private Type COLORRGBS
    R     As Single
    G     As Single
    B     As Single
End Type

Private Type PHOTEL 'PHOTo realistic face pixELs
    Y1      As Single
    U1      As Single
    V1      As Single
    Y2      As Single
    U2      As Single
    V2      As Single
    C1      As COLORRGBS
    C2      As COLORRGBS
    Used    As Boolean
End Type

Private Photels()  As PHOTEL

Public Sub DrawTrianglePhoto(idxMesh As Integer, idxFace As Long, Alpha As Single, Optional Col As Integer = 128)
    
    Dim minX As Long, maxX As Long
    Dim X1 As Long, Y1 As Long
    Dim X2 As Long, Y2 As Long
    Dim X3 As Long, Y3 As Long
    Dim U1 As Single, V1 As Single
    Dim U2 As Single, V2 As Single
    Dim U3 As Single, V3 As Single
    Dim C1 As COLORRGBS
    Dim C2 As COLORRGBS
    Dim C3 As COLORRGBS
    Dim rgbColor As COLORRGB

'Get points values
    With Meshs(idxMesh)
        X1 = .Screen(.Faces(idxFace).A).X
        Y1 = .Screen(.Faces(idxFace).A).Y
        X2 = .Screen(.Faces(idxFace).B).X
        Y2 = .Screen(.Faces(idxFace).B).Y
        X3 = .Screen(.Faces(idxFace).C).X
        Y3 = .Screen(.Faces(idxFace).C).Y
        Call CalculateUV(idxMesh, idxFace)
        U1 = .TScreen(.Faces(idxFace).A).U
        V1 = .TScreen(.Faces(idxFace).A).V
        U2 = .TScreen(.Faces(idxFace).B).U
        V2 = .TScreen(.Faces(idxFace).B).V
        U3 = .TScreen(.Faces(idxFace).C).U
        V3 = .TScreen(.Faces(idxFace).C).V
        C1 = ConvertSngColor(.Vertices(.Faces(idxFace).A).ColorS)
        C2 = ConvertSngColor(.Vertices(.Faces(idxFace).B).ColorS)
        C3 = ConvertSngColor(.Vertices(.Faces(idxFace).C).ColorS)
    End With
'Redim Photels
    minX = IIf(X1 < X2, X1, X2)
    minX = IIf(minX < X3, minX, X3)
    maxX = IIf(X1 > X2, X1, X2)
    maxX = IIf(maxX > X3, maxX, X3)
    ReDim Photels(minX To maxX)
'Line Interpolation
    Call LineInterpolatePhoto(X1, Y1, X2, Y2, U1, V1, U2, V2, C1, C2)
    Call LineInterpolatePhoto(X2, Y2, X3, Y3, U2, V2, U3, V3, C2, C3)
    Call LineInterpolatePhoto(X3, Y3, X1, Y1, U3, V3, U1, V1, C3, C1)
'Limits
    If minX < CanRect.L Then minX = CanRect.L
    If maxX > CanRect.R Then maxX = CanRect.R
'Fill
    For minX = minX To maxX
        FillPhoto minX, Meshs(idxMesh).Faces(idxFace).idxMat, Alpha, Col
    Next

End Sub

Private Sub LineInterpolatePhoto(ByVal X1 As Long, ByVal Y1 As Single, _
                                 ByVal X2 As Long, ByVal Y2 As Single, _
                                 ByVal U1 As Single, ByVal V1 As Single, _
                                 ByVal U2 As Single, ByVal V2 As Single, _
                                 C1 As COLORRGBS, C2 As COLORRGBS)

    Dim DeltaX  As Double, DeltaY   As Double
    Dim StartX  As Double, StartY   As Double
    Dim EndX    As Double, EndY     As Double
    Dim StepX   As Double, StepY    As Double
    Dim absDeltaY As Double
    Dim idx     As Long
    
    Dim StepC   As COLORRGBS
    Dim DeltaC  As COLORRGBS
    Dim StartC  As COLORRGBS
    Dim EndC    As COLORRGBS
    
    Dim StartU  As Single, StartV   As Single
    Dim EndU    As Single, EndV     As Single
    Dim StepU   As Single, StepV    As Single
    Dim DeltaU  As Single, DeltaV   As Single

    If X1 < X2 Then
        StartX = X1:    EndX = X2
        StartY = Y1:    EndY = Y2
        StartC = C1:    EndC = C2
        StartU = U1:    EndU = U2
        StartV = V1:    EndV = V2
    Else
        StartX = X2:    EndX = X1
        StartY = Y2:    EndY = Y1
        StartC = C2:    EndC = C1
        StartU = U2:    EndU = U1
        StartV = V2:    EndV = V1
    End If
    DeltaX = EndX - StartX
    DeltaY = EndY - StartY
    DeltaC.R = EndC.R - StartC.R
    DeltaC.G = EndC.G - StartC.G
    DeltaC.B = EndC.B - StartC.B
    DeltaU = EndU - StartU
    DeltaV = EndV - StartV
    
    If DeltaX > Abs(DeltaY) Then
        StepY = Div(DeltaY, DeltaX)
        StepC.R = Div(DeltaC.R, DeltaX)
        StepC.G = Div(DeltaC.G, DeltaX)
        StepC.B = Div(DeltaC.B, DeltaX)
        StepU = Div(DeltaU, DeltaX)
        StepV = Div(DeltaV, DeltaX)
        For StartX = StartX To EndX
            With Photels(StartX)
                If .Used Then
                    If .Y1 < Fix(StartY) Then .Y1 = Fix(StartY): .C1 = StartC: .U1 = StartU:  .V1 = StartV
                    If .Y2 > Fix(StartY) Then .Y2 = Fix(StartY): .C2 = StartC: .U2 = StartU:  .V2 = StartV
                Else
                    .Y1 = Fix(StartY): .C1 = StartC: .U1 = StartU:  .V1 = StartV
                    .Y2 = Fix(StartY): .C2 = StartC: .U2 = StartU:  .V2 = StartV
                    .Used = True
                End If
            End With
            StartY = StartY + StepY
            StartC.R = StartC.R + StepC.R
            StartC.G = StartC.G + StepC.G
            StartC.B = StartC.B + StepC.B
            StartU = StartU + StepU
            StartV = StartV + StepV
        Next
    Else
        absDeltaY = Abs(DeltaY)
        StepX = Div(DeltaX, absDeltaY)
        StepY = Div(DeltaY, absDeltaY)
        StepC.R = Div(DeltaC.R, absDeltaY)
        StepC.G = Div(DeltaC.G, absDeltaY)
        StepC.B = Div(DeltaC.B, absDeltaY)
        StepU = Div(DeltaU, absDeltaY)
        StepV = Div(DeltaV, absDeltaY)
        For idx = 0 To absDeltaY
            With Photels(StartX)
                If .Used Then
                    If .Y1 < Fix(StartY) Then .Y1 = Fix(StartY): .C1 = StartC: .U1 = StartU:  .V1 = StartV
                    If .Y2 > Fix(StartY) Then .Y2 = Fix(StartY): .C2 = StartC: .U2 = StartU:  .V2 = StartV
                Else
                    .Y1 = Fix(StartY): .C1 = StartC: .U1 = StartU:  .V1 = StartV
                    .Y2 = Fix(StartY): .C2 = StartC: .U2 = StartU:  .V2 = StartV
                    .Used = True
                End If
            End With
            StartX = StartX + StepX
            StartY = StartY + StepY
            StartC.R = StartC.R + StepC.R
            StartC.G = StartC.G + StepC.G
            StartC.B = StartC.B + StepC.B
            StartU = StartU + StepU
            StartV = StartV + StepV
        Next
    End If

End Sub

Private Sub FillPhoto(X As Long, idxMat As Integer, Alpha As Single, Col As Integer)

    Dim DeltaY  As Single
    Dim StepU   As Single
    Dim StepV   As Single
    Dim StepC   As COLORRGBS
    Dim minY    As Single
    Dim maxY    As Single
    Dim rgbTex  As COLORRGB
    Dim rgbCan  As COLORRGB
    Dim rgbGrad As COLORRGB
    Dim rgbColorS As COLORRGB
    Dim Gray As Integer
    Dim GrayS As Integer
    
    
    On Error Resume Next
    
    With Photels(X)
        DeltaY = .Y1 - .Y2
        StepU = Div(.U1 - .U2, DeltaY)
        StepV = Div(.V1 - .V2, DeltaY)
        StepC.R = Div(.C1.R - .C2.R, DeltaY)
        StepC.G = Div(.C1.G - .C2.G, DeltaY)
        StepC.B = Div(.C1.B - .C2.B, DeltaY)

        If .Y2 < CanRect.T Then
            minY = CanRect.T
            .U2 = .U2 + (StepU * Abs(CanRect.T - .Y2))
            .V2 = .V2 + (StepV * Abs(CanRect.T - .Y2))
        Else
            minY = .Y2
        End If
        maxY = IIf(.Y1 > CanRect.D, CanRect.D, .Y1)
        

        If (Alpha = 0) Then
            For minY = minY To maxY
                Call FindTexel(idxMat, .U2, .V2)
                rgbTex = ColorLongToRGB(Materials(idxMat).dibTexT.mapArray(Fix(.U2), Fix(.V2)))
                Gray = (ColorGray(SngColorToRGB(.C2)) - Col) * 0.5
                rgbCan = ColorPlus(rgbTex, Gray)
                dibCanvas.mapArray(X, minY) = ColorRGBToLong(rgbCan)
                .U2 = .U2 + StepU
                .V2 = .V2 + StepV
                .C2.R = .C2.R + StepC.R
                .C2.G = .C2.G + StepC.G
                .C2.B = .C2.B + StepC.B
            Next
        Else
            For minY = minY To maxY
                Call FindTexel(idxMat, .U2, .V2)
                rgbTex = ColorLongToRGB(Materials(idxMat).dibTexT.mapArray(Fix(.U2), Fix(.V2)))
                Gray = (ColorGray(SngColorToRGB(.C2)) - Col) * 0.5
                rgbTex = ColorPlus(rgbTex, Gray)
                rgbCan = ColorLongToRGB(dibCanvas.mapArray(X, minY))
                rgbCan = ColorInterpolate(rgbTex, rgbCan, Alpha)
                dibCanvas.mapArray(X, minY) = ColorRGBToLong(rgbCan)
                .U2 = .U2 + StepU
                .V2 = .V2 + StepV
            Next
        End If
    End With

End Sub

Private Function ConvertSngColor(C1 As COLORRGB) As COLORRGBS
    
    ConvertSngColor.R = CSng(C1.R)
    ConvertSngColor.G = CSng(C1.G)
    ConvertSngColor.B = CSng(C1.B)

End Function

Private Function SngColorToRGB(C1 As COLORRGBS) As COLORRGB

    If (C1.R > 255) Then C1.R = 255 Else If (C1.R < 0) Then C1.R = 0
    If (C1.G > 255) Then C1.G = 255 Else If (C1.G < 0) Then C1.G = 0
    If (C1.B > 255) Then C1.B = 255 Else If (C1.B < 0) Then C1.B = 0
    SngColorToRGB.R = CInt(C1.R)
    SngColorToRGB.G = CInt(C1.G)
    SngColorToRGB.B = CInt(C1.B)

End Function

