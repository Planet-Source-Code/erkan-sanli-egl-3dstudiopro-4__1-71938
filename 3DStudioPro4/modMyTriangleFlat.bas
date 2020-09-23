Attribute VB_Name = "modMyTriangleFlat"
Option Explicit

Private Type FLATEL 'FLAT face pixELs for transparency
    Y1      As Single
    Y2      As Single
    Used    As Boolean
End Type

Private Flatels()  As FLATEL

Public Sub DrawTriangleFlatA(idxMesh As Integer, idxFace As Long, Alpha As Single, Lit As Boolean)
    
    Dim minX As Long, maxX As Long
    Dim X1 As Long, Y1 As Long
    Dim X2 As Long, Y2 As Long
    Dim X3 As Long, Y3 As Long
    Dim rgbColor    As COLORRGB
    
'Get points values and shaded color
    With Meshs(idxMesh)
        X1 = .Screen(.Faces(idxFace).A).X
        Y1 = .Screen(.Faces(idxFace).A).Y
        X2 = .Screen(.Faces(idxFace).B).X
        Y2 = .Screen(.Faces(idxFace).B).Y
        X3 = .Screen(.Faces(idxFace).C).X
        Y3 = .Screen(.Faces(idxFace).C).Y
        If Lit Then
            rgbColor = ColorAverage(.Vertices(.Faces(idxFace).A).ColorS, _
                                    .Vertices(.Faces(idxFace).B).ColorS, _
                                    .Vertices(.Faces(idxFace).C).ColorS)
        Else
            rgbColor = Materials(.Faces(idxFace).idxMat).DiffuseColor
        End If
    End With
        
'Redim Flatels
    minX = IIf(X1 < X2, X1, X2)
    minX = IIf(minX < X3, minX, X3)
    maxX = IIf(X1 > X2, X1, X2)
    maxX = IIf(maxX > X3, maxX, X3)
    ReDim Flatels(minX To maxX)
'Line Interpolation
    Call LineInterpolateFlat(X3, Y3, X2, Y2)
    Call LineInterpolateFlat(X2, Y2, X1, Y1)
    Call LineInterpolateFlat(X1, Y1, X3, Y3)
'Limits
    If minX < CanRect.L Then minX = CanRect.L
    If maxX > CanRect.R Then maxX = CanRect.R
'Fill
    For minX = minX To maxX
        FillFlatA minX, rgbColor, Alpha
    Next
    
End Sub
'ok
Private Sub LineInterpolateFlat(ByVal X1 As Long, ByVal Y1 As Single, ByVal X2 As Long, ByVal Y2 As Single)

    Dim DeltaX  As Long
    Dim StepY   As Single
    
    If X1 < X2 Then
        DeltaX = X2 - X1
        StepY = Div(Y2 - Y1, DeltaX)
        For X1 = X1 To X2
            With Flatels(X1)
                If .Used Then
                    If .Y1 < Fix(Y1) Then .Y1 = Fix(Y1)
                    If .Y2 > Fix(Y1) Then .Y2 = Fix(Y1)
                Else
                    .Y1 = Fix(Y1)
                    .Y2 = Fix(Y1)
                    .Used = True
                End If
            End With
            Y1 = Y1 + StepY
        Next
    Else
        DeltaX = X1 - X2
        StepY = Div(Y1 - Y2, DeltaX)
        For X2 = X2 To X1
            With Flatels(X2)
                If .Used Then
                    If .Y1 < Fix(Y2) Then .Y1 = Fix(Y2)
                    If .Y2 > Fix(Y2) Then .Y2 = Fix(Y2)
                Else
                    .Y1 = Fix(Y2)
                    .Y2 = Fix(Y2)
                    .Used = True
                End If
            End With
            Y2 = Y2 + StepY
        Next
   End If

End Sub
'ok
Private Sub FillFlatA(X As Long, rgbColor As COLORRGB, Alpha As Single)

    Dim DeltaY  As Single
    Dim StepU   As Single
    Dim StepV   As Single
    Dim minY    As Single
    Dim maxY    As Single
    Dim rgbCan  As COLORRGB
    Dim lColor  As Long
    On Error Resume Next
    
    With Flatels(X)
        DeltaY = .Y1 - .Y2
        minY = IIf(.Y2 < CanRect.T, CanRect.T, .Y2)
        maxY = IIf(.Y1 > CanRect.D, CanRect.D, .Y1)
        
        Select Case Alpha '0=opaque 1=transparent
'            Case Is = 0
            Case Is = 1
                For minY = minY To maxY
                    dibCanvas.mapArray(X, minY) = dibBack.mapArray(X, minY)
                Next
            Case Else
                For minY = minY To maxY
                    rgbCan = ColorLongToRGB(dibCanvas.mapArray(X, minY))
                    rgbCan = ColorInterpolate(rgbColor, rgbCan, Alpha)
                    dibCanvas.mapArray(X, minY) = ColorBGRToLong(rgbCan)
                Next
        End Select
    End With

End Sub

