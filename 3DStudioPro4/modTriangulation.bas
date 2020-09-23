Attribute VB_Name = "modTriangulation"
Option Explicit

Public Function Triangulate(Faces() As FACE) As Long
    
    Dim Verts() As POINTAPI
    Dim NVert%
    Dim Edges() As Long
    Dim Complete() As Boolean
    Dim XMin&, XMax&, YMin&, YMax&, XMid&, YMid&
    Dim I%, J%, K%, NTri%, NEdge&, Inc As Boolean
    Dim Dx!, Dy!, DMax!, XC!, YC!, R!
    Dim DDMax!
    
    ReDim Complete(UBound(Faces))
    ReDim Edges(2, (UBound(Faces) * 3))
    
    Verts = Poly
    NVert = UBound(Verts)
    XMin = Verts(1).X
    YMin = Verts(1).Y
    XMax = XMin
    YMax = YMin
    For I = 2 To NVert
       If Verts(I).X < XMin Then XMin = Verts(I).X
       If Verts(I).X > XMax Then XMax = Verts(I).X
       If Verts(I).Y < YMin Then YMin = Verts(I).Y
       If Verts(I).Y > YMax Then YMax = Verts(I).Y
    Next I
    Dx = XMax - XMin
    Dy = YMax - YMin
    DMax = IIf(Dx > Dy, Dx, Dy)
    XMid = (XMax + XMin) * 0.5
    YMid = (YMax + YMin) * 0.5
    
    ReDim Preserve Verts(NVert + 3)
    DDMax = DMax * 2
    Verts(NVert + 1).X = XMid - DDMax
    Verts(NVert + 1).Y = YMid - DMax
    Verts(NVert + 2).X = XMid
    Verts(NVert + 2).Y = YMid + DDMax
    Verts(NVert + 3).X = XMid + DDMax
    Verts(NVert + 3).Y = YMid - DMax
    Faces(1).A = NVert + 1
    Faces(1).B = NVert + 2
    Faces(1).C = NVert + 3
    Complete(1) = False
    NTri = 1

    For I = 1 To NVert
        NEdge = 0
        J = 0
        Do
            J = (J + 1)
            If (Complete(J) = False) Then
                If InCircumCircle(Verts(I).X, Verts(I).Y, _
                                  Verts(Faces(J).A).X, Verts(Faces(J).A).Y, _
                                  Verts(Faces(J).B).X, Verts(Faces(J).B).Y, _
                                  Verts(Faces(J).C).X, Verts(Faces(J).C).Y, _
                                  XC, YC) Then
                    Edges(1, NEdge + 1) = Faces(J).A
                    Edges(2, NEdge + 1) = Faces(J).B
                    Edges(1, NEdge + 2) = Faces(J).B
                    Edges(2, NEdge + 2) = Faces(J).C
                    Edges(1, NEdge + 3) = Faces(J).C
                    Edges(2, NEdge + 3) = Faces(J).A
                    Faces(J).A = Faces(NTri).A
                    Faces(J).B = Faces(NTri).B
                    Faces(J).C = Faces(NTri).C
                    Complete(J) = Complete(NTri)
                    NEdge = (NEdge + 3)
                    J = (J - 1)
                    NTri = (NTri - 1)
                End If
            End If
        Loop While (J < NTri)
        
        For J = 1 To (NEdge - 1)
          If (Edges(1, J) <> 0) And (Edges(2, J) <> 0) Then
              For K = (J + 1) To NEdge
                  If (Edges(1, K) <> 0) And (Edges(2, K) <> 0) Then
                      If (Edges(1, J) = Edges(2, K)) Then
                          If (Edges(2, J) = Edges(1, K)) Then
                              Edges(1, J) = 0
                              Edges(2, J) = 0
                              Edges(1, K) = 0
                              Edges(2, K) = 0
                          End If
                      End If
                  End If
              Next K
          End If
        Next J
    
        For J = 1 To NEdge
            If (Edges(1, J) <> 0) And (Edges(2, J) <> 0) Then
                NTri = (NTri + 1)
                Faces(NTri).A = Edges(1, J)
                Faces(NTri).B = Edges(2, J)
                Faces(NTri).C = I
                Complete(NTri) = False
            End If
        Next J
    Next I

    I = 0
    Do
        I = (I + 1)
        If (Faces(I).A > NVert) Or (Faces(I).B > NVert) Or (Faces(I).C > NVert) Then
            Faces(I).A = Faces(NTri).A
            Faces(I).B = Faces(NTri).B
            Faces(I).C = Faces(NTri).C
            I = (I - 1)
            NTri = (NTri - 1)
        End If
    Loop While (I < NTri)
    Triangulate = NTri

End Function

Function InCircumCircle(XX&, YY&, X1&, Y1&, X2&, Y2&, X3&, Y3&, CX!, CY!) As Boolean

 ' FUNCTION : InCircumCircle
 ' =========================
 '
 ' RETURNED VALUES:
 '
 ' - Function   : InCircumCircle : Boolean
 ' - Parameters : CX, CY, R      : Single
 '
 ' Return true if the point (XX, YY) lies inside the CircumCircle
 '  made up by points (X1, Y1) (X2, Y2) (X3, Y3).
 ' The CircumCircle centre is returned in (CX, CY) and the radius as R.

 Dim RSqr!, DRSqr!, Dx!, Dy!, M1!, M2!, MX1!, MX2!, MY1!, MY2!

    If Y1 - Y2 < ApproachVal And _
       Y2 - Y1 < ApproachVal And _
       Y2 - Y3 < ApproachVal And _
       Y3 - Y2 < ApproachVal Then Exit Function
    
    If Y2 - Y1 < ApproachVal And Y1 - Y2 < ApproachVal Then
        M2 = -(X3 - X2) / (Y3 - Y2)
        MX2 = (X2 + X3) * 0.5
        MY2 = (Y2 + Y3) * 0.5
        CX = (X2 + X1) * 0.5
        CY = (M2 * (CX - MX2)) + MY2
    ElseIf Y3 - Y2 < ApproachVal And Y2 - Y3 < ApproachVal Then
        M1 = -(X2 - X1) / (Y2 - Y1)
        MX1 = (X1 + X2) * 0.5
        MY1 = (Y1 + Y2) * 0.5
        CX = (X3 + X2) * 0.5
        CY = (M1 * (CX - MX1)) + MY1
    Else
        M1 = -(X2 - X1) / (Y2 - Y1)
        M2 = -(X3 - X2) / (Y3 - Y2)
        MX1 = (X1 + X2) * 0.5
        MX2 = (X2 + X3) * 0.5
        MY1 = (Y1 + Y2) * 0.5
        MY2 = (Y2 + Y3) * 0.5
        CX = ((M1 * MX1 - M2 * MX2) + (MY2 - MY1)) / (M1 - M2)
        CY = (M1 * (CX - MX1)) + MY1
    End If

    Dx = X2 - CX
    Dy = Y2 - CY
    RSqr = Dx * Dx + Dy * Dy
    Dx = XX - CX
    Dy = YY - CY
    DRSqr = Dx * Dx + Dy * Dy
    If (DRSqr < RSqr) Then InCircumCircle = True

End Function


