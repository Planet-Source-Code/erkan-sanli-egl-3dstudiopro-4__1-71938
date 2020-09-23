Attribute VB_Name = "modMyLineTexture"
Option Explicit

Public Sub DrawLineTex(ByVal X1 As Long, ByVal Y1 As Single, _
                        ByVal X2 As Long, ByVal Y2 As Single, _
                        ByVal U1 As Single, ByVal V1 As Single, _
                        ByVal U2 As Single, ByVal V2 As Single, idxMat As Integer)

    Dim DeltaX  As Single, DeltaY   As Single
    Dim StartX  As Single, StartY   As Single
    Dim EndX    As Single, EndY     As Single
    Dim StepX   As Single, StepY    As Single
    Dim absDeltaY As Single
    Dim idx     As Long
    Dim StepU   As Single
    Dim StepV   As Single
    Dim DeltaU  As Single, DeltaV   As Single
    Dim StartU  As Single, StartV   As Single
    Dim EndU    As Single, EndV     As Single

    If X1 < X2 Then
        StartX = X1:    EndX = X2
        StartY = Y1:    EndY = Y2
        StartU = U1:    EndU = U2
        StartV = V1:    EndV = V2
    Else
        StartX = X2:    EndX = X1
        StartY = Y2:    EndY = Y1
        StartU = U2:    EndU = U1
        StartV = V2:    EndV = V1
    End If
    
    DeltaX = EndX - StartX
    DeltaY = EndY - StartY
    DeltaU = EndU - StartU
    DeltaV = EndV - StartV
    If DeltaX > Abs(DeltaY) Then
        StepY = Div(DeltaY, DeltaX)
        StepU = Div(DeltaU, DeltaX)
        StepV = Div(DeltaV, DeltaX)
        For StartX = StartX To EndX
            If StartX > CanRect.L And StartX < CanRect.R And StartY > CanRect.T And StartY < CanRect.D Then
                Call FindTexel(idxMat, StartU, StartV)
                dibCanvas.mapArray(StartX, StartY) = Materials(idxMat).dibTexT.mapArray(Fix(StartU), Fix(StartV))
            End If
            StartY = StartY + StepY
            StartU = StartU + StepU
            StartV = StartV + StepV
        Next
    Else
        absDeltaY = Abs(DeltaY)
        StepX = Div(DeltaX, absDeltaY)
        StepY = Div(DeltaY, absDeltaY)
        StepU = Div(DeltaU, absDeltaY)
        StepV = Div(DeltaV, absDeltaY)
        For idx = 0 To absDeltaY
            If StartX > CanRect.L And StartX < CanRect.R And StartY > CanRect.T And StartY < CanRect.D Then
                Call FindTexel(idxMat, StartU, StartV)
                dibCanvas.mapArray(StartX, StartY) = Materials(idxMat).dibTexT.mapArray(Fix(StartU), Fix(StartV))
            End If
            StartX = StartX + StepX
            StartY = StartY + StepY
            StartU = StartU + StepU
            StartV = StartV + StepV
        Next
    End If

End Sub


