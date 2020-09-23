Attribute VB_Name = "modClipper"
Option Explicit

Public Poly()    As POINTAPI

Public Function FaceStat(idxMesh As Integer, idxFace As Long, Optional Box As Boolean = False, Optional Wire As Boolean = False) As CanStat
    
    Dim P1 As POINTAPI
    Dim P2 As POINTAPI
    Dim P3 As POINTAPI
    
    Dim AB As Boolean
    Dim BC As Boolean
    Dim CA As Boolean
 
    With Meshs(idxMesh)
        If Box Then
            P1 = .BoxScreen(.BoxFace(idxFace).A)
            P2 = .BoxScreen(.BoxFace(idxFace).B)
            P3 = .BoxScreen(.BoxFace(idxFace).C)
            AB = .BoxFace(idxFace).AB
            BC = .BoxFace(idxFace).BC
            CA = .BoxFace(idxFace).CA
        Else
            P1 = .Screen(.Faces(idxFace).A)
            P2 = .Screen(.Faces(idxFace).B)
            P3 = .Screen(.Faces(idxFace).C)
            If .FaceInfo Then
                AB = .Faces(idxFace).AB
                BC = .Faces(idxFace).BC
                CA = .Faces(idxFace).CA
            Else
                AB = True
                BC = True
                CA = True
            End If
        End If
    End With
    
    FaceStat = ClipTriangle(P1, P2, P3, AB, BC, CA, Wire)

End Function

Private Function ClipTriangle(P1 As POINTAPI, P2 As POINTAPI, P3 As POINTAPI, AB As Boolean, BC As Boolean, CA As Boolean, Optional Wire As Boolean = False) As CanStat
 
    Dim OA1 As POINTAPI, OA2 As POINTAPI
    Dim OB1 As POINTAPI, OB2 As POINTAPI
    Dim OC1 As POINTAPI, OC2 As POINTAPI
    Dim A1 As Byte, A2 As Byte, A3 As Byte
    Dim R1 As Byte, R2 As Byte, R3 As Byte

    ReDim Poly(0)

'CanvasCorner========================================================
'Big face from canvas

    If AddCorner(P1, P2, P3, CanRect.L, CanRect.T) And _
       AddCorner(P1, P2, P3, CanRect.L, CanRect.D) And _
       AddCorner(P1, P2, P3, CanRect.R, CanRect.T) And _
       AddCorner(P1, P2, P3, CanRect.R, CanRect.D) Then
       ClipTriangle = CanClip
       Exit Function
    End If
    
'Find intersections
    Call ClipLine(P1, P2, OA1, OA2)
    Call ClipLine(P2, P3, OB1, OB2)
    Call ClipLine(P3, P1, OC1, OC2)
    
    A1 = IIf(Accept2D(P1.X, P1.Y, P2.X, P2.Y), 1, 0)
    A2 = IIf(Accept2D(P2.X, P2.Y, P3.X, P3.Y), 1, 0)
    A3 = IIf(Accept2D(P3.X, P3.Y, P1.X, P1.Y), 1, 0)
    R1 = IIf(OA1.X = 0 And OA1.Y = 0 And OA2.X = 0 And OA2.Y = 0, 1, 0)
    R2 = IIf(OB1.X = 0 And OB1.Y = 0 And OB2.X = 0 And OB2.Y = 0, 1, 0)
    R3 = IIf(OC1.X = 0 And OC1.Y = 0 And OC2.X = 0 And OC2.Y = 0, 1, 0)
    
'Completly outside=======================================================
    
    If A1 = 1 And A2 = 1 And A3 = 1 And R1 = 1 And R2 = 1 And R3 = 1 Then
        ClipTriangle = CanOut
        Exit Function
    End If

'Completely inside=======================================================
    
    If A1 = 0 And A2 = 0 And A3 = 0 And R1 = 0 And R2 = 0 And R3 = 0 Then
        ClipTriangle = CanIn
        ReDim Poly(3)
        Poly(1) = P1
        Poly(2) = P2
        Poly(3) = P3
        Exit Function
    End If

'Region is intersect triangle==========================================
    ClipTriangle = CanClip
    If Wire Then
        'AB
        If AB Then
            If A1 = 1 Then
                If R1 = 0 Then Call Add2Point(OA1, OA2)
            Else
                Call Add2Point(P1, P2)
            End If
        End If
        'BC
        If BC Then
            If A2 = 1 Then
                If R2 = 0 Then Call Add2Point(OB1, OB2)
            Else
                Call Add2Point(P2, P3)
            End If
        End If
        'CA
        If CA Then
            If A3 = 1 Then
                If R3 = 0 Then Call Add2Point(OC1, OC2)
            Else
                Call Add2Point(P3, P1)
            End If
        End If
    Else
        'AB
            If A1 = 1 Then
                If R1 = 0 Then Call AddPoint(OA1): Call AddPoint(OA2)
            Else
                Call AddPoint(P1): Call AddPoint(P2)
            End If
        'BC
            If A2 = 1 Then
                If R2 = 0 Then Call AddPoint(OB1): Call AddPoint(OB2)
            Else
                Call AddPoint(P2): Call AddPoint(P3)
            End If
        'CA
            If A3 = 1 Then
                If R3 = 0 Then Call AddPoint(OC1): Call AddPoint(OC2)
            Else
                Call AddPoint(P3): Call AddPoint(P1)
            End If
    End If

End Function

Private Sub ClipLine(P1 As POINTAPI, P2 As POINTAPI, OP1 As POINTAPI, OP2 As POINTAPI)

    Dim PX1!, PY1!, PX2!, PY2!, U1!, U2!, Dx!, Dy!, P!, Q!, R!, CT As Byte

    U1 = 0
    U2 = 1
    PX1 = P1.X
    PY1 = P1.Y
    PX2 = P2.X
    PY2 = P2.Y
    Dx = (PX2 - PX1)
    Dy = (PY2 - PY1)
    P = -Dx
    Q = (PX1 - CanRect.L)
 
    If (P < 0) Then
        R = (Q / P)
        If (R > U2) Then CT = 1 Else If (R > U1) Then U1 = R
    ElseIf (P > 0) Then
        R = (Q / P)
        If (R < U1) Then CT = 1 Else If (R < U2) Then U2 = R
    ElseIf (Q < 0) Then
        CT = 1
    End If

    If CT = 0 Then
        P = Dx
        Q = (CanRect.R - PX1)
        
        If (P < 0) Then
            R = (Q / P)
            If (R > U2) Then CT = 1 Else If (R > U1) Then U1 = R
        ElseIf (P > 0) Then
            R = (Q / P)
            If (R < U1) Then CT = 1 Else If (R < U2) Then U2 = R
        ElseIf (Q < 0) Then
            CT = 1
        End If
    
        If CT = 0 Then
            P = -Dy
            Q = (PY1 - CanRect.T)
            If (P < 0) Then
                R = (Q / P)
                If (R > U2) Then CT = 1 Else If (R > U1) Then U1 = R
            ElseIf (P > 0) Then
                R = (Q / P)
                If (R < U1) Then CT = 1 Else If (R < U2) Then U2 = R
            ElseIf (Q < 0) Then
                CT = 1
            End If
            
            If CT = 0 Then
                P = Dy
                Q = (CanRect.D - PY1)
                If (P < 0) Then
                    R = (Q / P)
                    If (R > U2) Then CT = 1 Else If (R > U1) Then U1 = R
                ElseIf (P > 0) Then
                    R = (Q / P)
                    If (R < U1) Then CT = 1 Else If (R < U2) Then U2 = R
                ElseIf (Q < 0) Then
                    CT = 1
                End If
                
                If CT = 0 Then
                    If (U2 < 1) Then PX2 = (PX1 + (U2 * Dx)): PY2 = (PY1 + (U2 * Dy))
                    If (U1 > 0) Then PX1 = (PX1 + (U1 * Dx)): PY1 = (PY1 + (U1 * Dy))
                    OP1.X = PX1
                    OP1.Y = PY1
                    OP2.X = PX2
                    OP2.Y = PY2
                End If
            End If
        End If
    End If

End Sub

Function Accept2D(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long) As Boolean

    If CanRect.L > X1 Or X1 > CanRect.R Or _
       CanRect.T > Y1 Or Y1 > CanRect.D Or _
       CanRect.L > X2 Or X2 > CanRect.R Or _
       CanRect.T > Y2 Or Y2 > CanRect.D Then Accept2D = True

End Function

Function IsInTriangle(P1 As POINTAPI, P2 As POINTAPI, P3 As POINTAPI, PX As Long, PY As Long) As Boolean

    Dim Val1 As Double, Val2 As Double, Val3 As Double
    Dim P1x As Double, P1y As Double
    Dim P2x As Double, P2y As Double
    Dim P3x As Double, P3y As Double
    
    P1x = P1.X: P1y = P1.Y
    P2x = P2.X: P2y = P2.Y
    P3x = P3.X: P3y = P3.Y
    
    Val1 = (P1x - PX) * (P2y - PY) - (P2x - PX) * (P1y - PY)
    Val2 = (P2x - PX) * (P3y - PY) - (P3x - PX) * (P2y - PY)
    Val3 = (P3x - PX) * (P1y - PY) - (P1x - PX) * (P3y - PY)
    
    If (Val1 > 0 And Val2 > 0 And Val3 > 0) Or _
       (Val1 < 0 And Val2 < 0 And Val3 < 0) Then IsInTriangle = True

End Function

Private Function AddCorner(P1 As POINTAPI, P2 As POINTAPI, P3 As POINTAPI, CX As Long, CY As Long) As Boolean
    
    If IsInTriangle(P1, P2, P3, CX, CY) = True Then
        ReDim Preserve Poly(UBound(Poly) + 1)
        Poly(UBound(Poly)).X = CX
        Poly(UBound(Poly)).Y = CY
        AddCorner = True
    End If

End Function

Private Sub AddPoint(P1 As POINTAPI)

    Dim idx As Long
    Dim twin As Boolean

    For idx = 0 To UBound(Poly)
        If Poly(idx).X = P1.X And Poly(idx).Y = P1.Y Then twin = True
    Next
    If twin = False Then
        ReDim Preserve Poly(UBound(Poly) + 1)
        Poly(UBound(Poly)) = P1
    End If

End Sub

Private Sub Add2Point(P1 As POINTAPI, P2 As POINTAPI)

    ReDim Preserve Poly(UBound(Poly) + 2)
    Poly(UBound(Poly) - 1) = P1
    Poly(UBound(Poly)) = P2

End Sub

