Attribute VB_Name = "modMyTriangleGradient"
Option Explicit

'Draw gradient triangle with transparent

Private Type COLORRGBS
    R     As Single
    G     As Single
    B     As Single
End Type

Private Type GRATEL 'GRADient face pixELs for transparent
    Y1      As Single
    C1      As COLORRGBS
    Y2      As Single
    C2      As COLORRGBS
    Used    As Boolean
End Type

Private Gratels()  As GRATEL

Public Sub DrawTriangleGradientA(idxMesh As Integer, idxFace As Long, Alpha As Single)
    
    Dim minX As Double, maxX As Double
    Dim X1 As Long, Y1 As Long
    Dim X2 As Long, Y2 As Long
    Dim X3 As Long, Y3 As Long
    Dim C1 As COLORRGBS
    Dim C2 As COLORRGBS
    Dim C3 As COLORRGBS
    
'Get points values
    With Meshs(idxMesh)
        X1 = .Screen(.Faces(idxFace).A).X
        Y1 = .Screen(.Faces(idxFace).A).Y
        C1 = ConvertSngColor(.Vertices(.Faces(idxFace).A).ColorS)
        
        X2 = .Screen(.Faces(idxFace).B).X
        Y2 = .Screen(.Faces(idxFace).B).Y
        C2 = ConvertSngColor(.Vertices(.Faces(idxFace).B).ColorS)
        
        X3 = .Screen(.Faces(idxFace).C).X
        Y3 = .Screen(.Faces(idxFace).C).Y
        C3 = ConvertSngColor(.Vertices(.Faces(idxFace).C).ColorS)
    End With
'Redim Gratels
    minX = IIf(X1 < X2, X1, X2)
    minX = IIf(minX < X3, minX, X3)
    maxX = IIf(X1 > X2, X1, X2)
    maxX = IIf(maxX > X3, maxX, X3)
    ReDim Gratels(minX To maxX)

'Line Interpolation
    Call LineInterpolateGrad(X1, Y1, X2, Y2, C1, C2)
    Call LineInterpolateGrad(X2, Y2, X3, Y3, C2, C3)
    Call LineInterpolateGrad(X3, Y3, X1, Y1, C3, C1)
'Fill
    For minX = minX To maxX
        FillGradientA minX, Alpha
    Next
    
End Sub

Private Sub LineInterpolateGrad(ByVal X1 As Long, ByVal Y1 As Single, ByVal X2 As Long, ByVal Y2 As Single, C1 As COLORRGBS, C2 As COLORRGBS)

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

    If X1 < X2 Then
        StartX = X1:    EndX = X2
        StartY = Y1:    EndY = Y2
        StartC = C1:    EndC = C2
    Else
        StartX = X2:    EndX = X1
        StartY = Y2:    EndY = Y1
        StartC = C2:    EndC = C1
    End If
    DeltaX = EndX - StartX
    DeltaY = EndY - StartY
    DeltaC.R = EndC.R - StartC.R
    DeltaC.G = EndC.G - StartC.G
    DeltaC.B = EndC.B - StartC.B

    If DeltaX > Abs(DeltaY) Then
        StepY = Div(DeltaY, DeltaX)
        StepC.R = Div(DeltaC.R, DeltaX)
        StepC.G = Div(DeltaC.G, DeltaX)
        StepC.B = Div(DeltaC.B, DeltaX)
        For StartX = StartX To EndX
            With Gratels(StartX)
                If .Used Then
                    If .Y1 < Fix(StartY) Then .Y1 = Fix(StartY): .C1 = StartC
                    If .Y2 > Fix(StartY) Then .Y2 = Fix(StartY): .C2 = StartC
                Else
                    .Y1 = Fix(StartY): .C1 = StartC
                    .Y2 = Fix(StartY): .C2 = StartC
                    .Used = True
                End If
            End With
            StartY = StartY + StepY
            StartC.R = StartC.R + StepC.R
            StartC.G = StartC.G + StepC.G
            StartC.B = StartC.B + StepC.B
        Next
    Else
        absDeltaY = Abs(DeltaY)
        StepX = Div(DeltaX, absDeltaY)
        StepY = Div(DeltaY, absDeltaY)
        StepC.R = Div(DeltaC.R, absDeltaY)
        StepC.G = Div(DeltaC.G, absDeltaY)
        StepC.B = Div(DeltaC.B, absDeltaY)
        For idx = 0 To absDeltaY
            With Gratels(StartX)
                If .Used Then
                    If .Y1 < Fix(StartY) Then .Y1 = Fix(StartY): .C1 = StartC
                    If .Y2 > Fix(StartY) Then .Y2 = Fix(StartY): .C2 = StartC
                Else
                    .Y1 = Fix(StartY): .C1 = StartC
                    .Y2 = Fix(StartY): .C2 = StartC
                    .Used = True
                End If
            End With
            StartX = StartX + StepX
            StartY = StartY + StepY
            StartC.R = StartC.R + StepC.R
            StartC.G = StartC.G + StepC.G
            StartC.B = StartC.B + StepC.B
        Next
    End If

End Sub

Private Sub FillGradientA(X As Double, Alpha As Single)

    Dim DeltaY  As Single
    Dim StepC   As COLORRGBS
    Dim minY    As Single
    Dim maxY    As Single
    Dim rgbCan  As COLORRGB
    Dim rgbGrad As COLORRGB

    On Error Resume Next
    
    With Gratels(X)
        DeltaY = .Y1 - .Y2
        StepC.R = Div(.C1.R - .C2.R, DeltaY)
        StepC.G = Div(.C1.G - .C2.G, DeltaY)
        StepC.B = Div(.C1.B - .C2.B, DeltaY)
        minY = .Y2 + 1
        maxY = .Y1
        For minY = minY To maxY
            rgbCan = ColorLongToRGB(dibCanvas.mapArray(X, minY))
            rgbGrad = SngColorToRGB(.C2)
            rgbCan = ColorInterpolate(rgbGrad, rgbCan, Alpha)
            dibCanvas.mapArray(X, minY) = ColorBGRToLong(rgbCan)
            .C2.R = .C2.R + StepC.R
            .C2.G = .C2.G + StepC.G
            .C2.B = .C2.B + StepC.B
        Next
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
