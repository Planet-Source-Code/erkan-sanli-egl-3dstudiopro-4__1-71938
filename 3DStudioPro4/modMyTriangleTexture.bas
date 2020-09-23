Attribute VB_Name = "modMyTriangleTexture"
Option Explicit

Private Type TEXEL 'TEXturized face pixELs
    Y1      As Single
    U1      As Single
    V1      As Single
    Y2      As Single
    U2      As Single
    V2      As Single
    Used    As Boolean
End Type

Private Texels()  As TEXEL

Public Sub DrawTriangleTex(idxMesh As Integer, idxFace As Long, Alpha As Single, Lit As Boolean, Optional Col As Integer = 128)
    
    Dim minX As Long, maxX As Long
    Dim X1 As Long, Y1 As Long
    Dim X2 As Long, Y2 As Long
    Dim X3 As Long, Y3 As Long
    Dim U1 As Single, V1 As Single
    Dim U2 As Single, V2 As Single
    Dim U3 As Single, V3 As Single
    
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
    End With
'Redim Texels
    minX = IIf(X1 < X2, X1, X2)
    minX = IIf(minX < X3, minX, X3)
    maxX = IIf(X1 > X2, X1, X2)
    maxX = IIf(maxX > X3, maxX, X3)
    ReDim Texels(minX To maxX)
'Line Interpolation
    Call LineInterpolateTex(X3, Y3, X2, Y2, U3, V3, U2, V2)
    Call LineInterpolateTex(X2, Y2, X1, Y1, U2, V2, U1, V1)
    Call LineInterpolateTex(X1, Y1, X3, Y3, U1, V1, U3, V3)
'Limits
    If minX < CanRect.L Then minX = CanRect.L
    If maxX > CanRect.R Then maxX = CanRect.R
'Fill
    For minX = minX To maxX
        FillTex minX, Meshs(idxMesh).Faces(idxFace).idxMat, Alpha, Lit, Col
    Next

End Sub

Private Sub LineInterpolateTex(ByVal X1 As Long, ByVal Y1 As Single, _
                               ByVal X2 As Long, ByVal Y2 As Single, _
                               ByVal U1 As Single, ByVal V1 As Single, _
                               ByVal U2 As Single, ByVal V2 As Single)
    Dim DeltaX  As Long
    Dim StepY   As Single
    Dim StepU   As Single
    Dim StepV   As Single

    If X1 < X2 Then
        DeltaX = X2 - X1
        StepY = Div(Y2 - Y1, DeltaX)
        StepU = Div(U2 - U1, DeltaX)
        StepV = Div(V2 - V1, DeltaX)
        For X1 = X1 To X2
            With Texels(X1)
                If .Used Then
                    If .Y1 < Fix(Y1) Then .Y1 = Fix(Y1): .U1 = U1:  .V1 = V1
                    If .Y2 > Fix(Y1) Then .Y2 = Fix(Y1): .U2 = U1:  .V2 = V1
                Else
                    .Y1 = Fix(Y1): .U1 = U1:  .V1 = V1
                    .Y2 = Fix(Y1): .U2 = U1:  .V2 = V1
                    .Used = True
                End If
            End With
            Y1 = Y1 + StepY
            U1 = U1 + StepU
            V1 = V1 + StepV
        Next
    Else
        DeltaX = X1 - X2
        StepY = Div(Y1 - Y2, DeltaX)
        StepU = Div(U1 - U2, DeltaX)
        StepV = Div(V1 - V2, DeltaX)
        For X2 = X2 To X1
            With Texels(X2)
                If .Used Then
                    If .Y1 < Fix(Y2) Then .Y1 = Fix(Y2): .U1 = U2:  .V1 = V2
                    If .Y2 > Fix(Y2) Then .Y2 = Fix(Y2): .U2 = U2:  .V2 = V2
                Else
                    .Y1 = Fix(Y2): .U1 = U2:  .V1 = V2
                    .Y2 = Fix(Y2): .U2 = U2:  .V2 = V2
                    .Used = True
                End If
            End With
            Y2 = Y2 + StepY
            U2 = U2 + StepU
            V2 = V2 + StepV
        Next
   End If

End Sub

Private Sub FillTex(X As Long, idxMat As Integer, Alpha As Single, Lit As Boolean, Optional Col As Integer = 128)

    Dim DeltaY  As Single
    Dim StepU   As Single
    Dim StepV   As Single
    Dim minY    As Single
    Dim maxY    As Single
    Dim rgbTex  As COLORRGB
    Dim rgbCan  As COLORRGB
    
    On Error Resume Next
    
    With Texels(X)
        DeltaY = .Y1 - .Y2
        StepU = Div(.U1 - .U2, DeltaY)
        StepV = Div(.V1 - .V2, DeltaY)
        If .Y2 < CanRect.T Then
            minY = CanRect.T
            .U2 = .U2 + (StepU * Abs(CanRect.T - .Y2))
            .V2 = .V2 + (StepV * Abs(CanRect.T - .Y2))
        Else
            minY = .Y2
        End If
        maxY = IIf(.Y1 > CanRect.D, CanRect.D, .Y1)
        
        If (Alpha = 0) Then
            If Lit Then
                For minY = minY To maxY
                    Call FindTexel(idxMat, .U2, .V2)
                    rgbTex = ColorLongToRGB(Materials(idxMat).dibTexT.mapArray(Fix(.U2), Fix(.V2)))
                    rgbTex = ColorPlus(rgbTex, Col)
                    dibCanvas.mapArray(X, minY) = ColorRGBToLong(rgbTex)
                    .U2 = .U2 + StepU
                    .V2 = .V2 + StepV
                Next
            Else
                For minY = minY To maxY
                    Call FindTexel(idxMat, .U2, .V2)
                    dibCanvas.mapArray(X, minY) = Materials(idxMat).dibTexT.mapArray(Fix(.U2), Fix(.V2))
                    .U2 = .U2 + StepU
                    .V2 = .V2 + StepV
                Next
            End If
        Else
            If Lit Then
                For minY = minY To maxY
                    Call FindTexel(idxMat, .U2, .V2)
                    rgbTex = ColorLongToRGB(Materials(idxMat).dibTexT.mapArray(Fix(.U2), Fix(.V2)))
                    rgbTex = ColorPlus(rgbTex, Col)
                    rgbCan = ColorLongToRGB(dibCanvas.mapArray(X, minY))
                    rgbCan = ColorInterpolate(rgbTex, rgbCan, Alpha)
                    dibCanvas.mapArray(X, minY) = ColorRGBToLong(rgbCan)
                    .U2 = .U2 + StepU
                    .V2 = .V2 + StepV
                Next
            Else
                For minY = minY To maxY
                    Call FindTexel(idxMat, .U2, .V2)
                    rgbTex = ColorLongToRGB(Materials(idxMat).dibTexT.mapArray(Fix(.U2), Fix(.V2)))
                    rgbCan = ColorLongToRGB(dibCanvas.mapArray(X, minY))
                    rgbCan = ColorInterpolate(rgbTex, rgbCan, Alpha)
                    dibCanvas.mapArray(X, minY) = ColorRGBToLong(rgbCan)
                    .U2 = .U2 + StepU
                    .V2 = .V2 + StepV
                Next
            End If
        End If
    End With

End Sub

Public Function Div(ByVal R1 As Single, ByVal R2 As Single) As Single
    
    If R2 = 0 Then R2 = ApproachVal
    Div = CSng(R1 / R2)

End Function

Public Sub FindTexel(idx As Integer, U As Single, V As Single)
    
    Dim texWidth As Long
    Dim texHeight As Long
    
    texWidth = Materials(idx).dibTexT.mapDIB.Width - 1
    texHeight = Materials(idx).dibTexT.mapDIB.Height - 1
    
    If U > texWidth Then
        Do
            U = U - texWidth
        Loop Until U < texWidth
    End If
    
    If U < 0 Then
        Do
            U = U + texWidth
        Loop Until U > 0
    End If
    
    If V > texHeight Then
        Do
            V = V - (texHeight)
        Loop Until V < texHeight
    End If
    
    If V < 0 Then
        Do
            V = V + (texHeight)
        Loop Until V > 0
    End If
    
End Sub

