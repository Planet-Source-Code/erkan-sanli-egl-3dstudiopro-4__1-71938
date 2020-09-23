Attribute VB_Name = "modGradientAPI"
Option Explicit

Private Const GRADIENT_FILL_TRIANGLE    As Long = &H2

Private Type TRIVERTEX
    X           As Long
    Y           As Long
    Red         As Integer
    Green       As Integer
    Blue        As Integer
    Alpha       As Integer
End Type

Private Type GRADIENT_TRIANGLE
    Vertex1     As Long
    Vertex2     As Long
    Vertex3     As Long
End Type

Private Declare Function GradientFillTriangle Lib "msimg32" Alias "GradientFill" ( _
                                        ByVal hDC As Long, pVertex As TRIVERTEX, _
                                        ByVal dwNumVertex As Long, _
                                        pMesh As GRADIENT_TRIANGLE, _
                                        ByVal dwNumMesh As Long, _
                                        ByVal dwMode As Long) As Long
Private triVert(2) As TRIVERTEX

Public Sub DrawTriangleGradient(idxMesh As Integer, idxFace As Long)

    Dim Vert(2) As TRIVERTEX
    Dim gTri    As GRADIENT_TRIANGLE
    
    With Meshs(idxMesh)
    
        Vert(0).X = .Screen(.Faces(idxFace).A).X
        Vert(0).Y = .Screen(.Faces(idxFace).A).Y
        Vert(0).Red = ConvertUShort(.Vertices(.Faces(idxFace).A).ColorS.R)
        Vert(0).Green = ConvertUShort(.Vertices(.Faces(idxFace).A).ColorS.G)
        Vert(0).Blue = ConvertUShort(.Vertices(.Faces(idxFace).A).ColorS.B)
        
        Vert(1).X = .Screen(.Faces(idxFace).B).X
        Vert(1).Y = .Screen(.Faces(idxFace).B).Y
        Vert(1).Red = ConvertUShort(.Vertices(.Faces(idxFace).B).ColorS.R)
        Vert(1).Green = ConvertUShort(.Vertices(.Faces(idxFace).B).ColorS.G)
        Vert(1).Blue = ConvertUShort(.Vertices(.Faces(idxFace).B).ColorS.B)
    
        Vert(2).X = .Screen(.Faces(idxFace).C).X
        Vert(2).Y = .Screen(.Faces(idxFace).C).Y
        Vert(2).Red = ConvertUShort(.Vertices(.Faces(idxFace).C).ColorS.R)
        Vert(2).Green = ConvertUShort(.Vertices(.Faces(idxFace).C).ColorS.G)
        Vert(2).Blue = ConvertUShort(.Vertices(.Faces(idxFace).C).ColorS.B)
    
    End With
    gTri.Vertex1 = 0
    gTri.Vertex2 = 1
    gTri.Vertex3 = 2
    Call GradientFillTriangle(dibCanvas.mapDIB.hDC, Vert(0), 3, gTri, 1, GRADIENT_FILL_TRIANGLE)

End Sub

Public Sub DrawPolygonGradient(idxMesh As Integer, idxFace As Long)

    Dim Tri(2) As TRIVERTEX
    Dim Vert(2) As TRIVERTEX
    Dim ClipTri() As TRIVERTEX
    
    Dim gTri    As GRADIENT_TRIANGLE
    
    Dim NumFaces    As Long
    Dim Faces(15)   As FACE
    Dim idx         As Long
    
    On Error GoTo Escape
    
    NumFaces = Triangulate(Faces)
    If NumFaces > 0 Then
        With Meshs(idxMesh)
        
            Tri(0).X = .Screen(.Faces(idxFace).A).X
            Tri(0).Y = .Screen(.Faces(idxFace).A).Y
            Tri(0).Red = .Vertices(.Faces(idxFace).A).ColorS.R
            Tri(0).Green = .Vertices(.Faces(idxFace).A).ColorS.G
            Tri(0).Blue = .Vertices(.Faces(idxFace).A).ColorS.B
            
            Tri(1).X = .Screen(.Faces(idxFace).B).X
            Tri(1).Y = .Screen(.Faces(idxFace).B).Y
            Tri(1).Red = .Vertices(.Faces(idxFace).B).ColorS.R
            Tri(1).Green = .Vertices(.Faces(idxFace).B).ColorS.G
            Tri(1).Blue = .Vertices(.Faces(idxFace).B).ColorS.B
        
            Tri(2).X = .Screen(.Faces(idxFace).C).X
            Tri(2).Y = .Screen(.Faces(idxFace).C).Y
            Tri(2).Red = .Vertices(.Faces(idxFace).C).ColorS.R
            Tri(2).Green = .Vertices(.Faces(idxFace).C).ColorS.G
            Tri(2).Blue = .Vertices(.Faces(idxFace).C).ColorS.B
        
        End With
        
'Bary Centric Coords

        Dim D As Single, U As Single, V As Single, W As Single
        Dim X1 As Single, Y1 As Single
        Dim X2 As Single, Y2 As Single
        Dim X3 As Single, Y3 As Single
        Dim PX As Single, PY As Single
        
        X1 = CSng(Tri(0).X)
        X2 = CSng(Tri(1).X)
        X3 = CSng(Tri(2).X)
        Y1 = CSng(Tri(0).Y)
        Y2 = CSng(Tri(1).Y)
        Y3 = CSng(Tri(2).Y)
        D = Div(1, (((X2 - X1) * (Y3 - Y1)) - ((Y2 - Y1) * (X3 - X1))))
        
        ReDim ClipTri(UBound(Poly))
        For idx = 1 To UBound(Poly)
            ClipTri(idx).X = Poly(idx).X
            ClipTri(idx).Y = Poly(idx).Y
            PX = CSng(ClipTri(idx).X)
            PY = CSng(ClipTri(idx).Y)
            U = ((X2 - PX) * (Y3 - PY) - (Y2 - PY) * (X3 - PX)) * D
            V = ((X3 - PX) * (Y1 - PY) - (Y3 - PY) * (X1 - PX)) * D
            W = 1 - (U + V)
            ClipTri(idx).Red = (U * Tri(0).Red) + (V * Tri(1).Red) + (W * Tri(2).Red)
            ClipTri(idx).Green = (U * Tri(0).Green) + (V * Tri(1).Green) + (W * Tri(2).Green)
            ClipTri(idx).Blue = (U * Tri(0).Blue) + (V * Tri(1).Blue) + (W * Tri(2).Blue)
            If (ClipTri(idx).Red > 255) Then ClipTri(idx).Red = 255 Else If (ClipTri(idx).Red < 0) Then ClipTri(idx).Red = 0
            If (ClipTri(idx).Green > 255) Then ClipTri(idx).Green = 255 Else If (ClipTri(idx).Green < 0) Then ClipTri(idx).Green = 0
            If (ClipTri(idx).Blue > 255) Then ClipTri(idx).Blue = 255 Else If (ClipTri(idx).Blue < 0) Then ClipTri(idx).Blue = 0
            ClipTri(idx).Red = ConvertUShort(ClipTri(idx).Red)
            ClipTri(idx).Green = ConvertUShort(ClipTri(idx).Green)
            ClipTri(idx).Blue = ConvertUShort(ClipTri(idx).Blue)
        Next idx
        For idx = 0 To NumFaces
            Vert(0) = ClipTri(Faces(idx).A)
            Vert(1) = ClipTri(Faces(idx).B)
            Vert(2) = ClipTri(Faces(idx).C)
            gTri.Vertex1 = 0
            gTri.Vertex2 = 1
            gTri.Vertex3 = 2
            Call GradientFillTriangle(dibCanvas.mapDIB.hDC, Vert(0), 3, gTri, 1, GRADIENT_FILL_TRIANGLE)
        Next
    End If

Escape:

End Sub

Private Function ConvertUShort(Color As Integer) As Integer
    
    Dim Unsigned As Long
    
    Unsigned = Color * 256&
    If Unsigned < &H8000& Then
        ConvertUShort = CInt(Unsigned)
    Else
        ConvertUShort = CInt(Unsigned - &H10000)
    End If
        
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Background~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub GradientBack()

    Dim gTri As GRADIENT_TRIANGLE
    gTri.Vertex1 = 0
    gTri.Vertex2 = 1
    gTri.Vertex3 = 2
    Call GradientFillTriangle(dibBack.mapDIB.hDC, triVert(0), 3, gTri, 1, GRADIENT_FILL_TRIANGLE)

End Sub

Public Sub Gradient0()
    
    triVert(0).X = 0
    triVert(0).Y = 0
    triVert(0).Red = ConvertUShort(BColor1.R)
    triVert(0).Green = ConvertUShort(BColor1.G)
    triVert(0).Blue = ConvertUShort(BColor1.B)
    
    triVert(1).X = CanRect.R
    triVert(1).Y = 0
    triVert(1).Red = ConvertUShort(BColor1.R)
    triVert(1).Green = ConvertUShort(BColor1.G)
    triVert(1).Blue = ConvertUShort(BColor1.B)

    triVert(2).X = 0
    triVert(2).Y = CanRect.D
    triVert(2).Red = ConvertUShort(BColor1.R)
    triVert(2).Green = ConvertUShort(BColor1.G)
    triVert(2).Blue = ConvertUShort(BColor1.B)
    
    Call GradientBack
    
    triVert(0).X = CanRect.R
    triVert(0).Y = CanRect.D
    
    Call GradientBack

End Sub

Public Sub Gradient1()
    
    triVert(0).X = 0
    triVert(0).Y = 0
    triVert(0).Red = ConvertUShort(BColor1.R)
    triVert(0).Green = ConvertUShort(BColor1.G)
    triVert(0).Blue = ConvertUShort(BColor1.B)
    
    triVert(1).X = CanRect.R
    triVert(1).Y = 0
    triVert(1).Red = ConvertUShort(BColor1.R)
    triVert(1).Green = ConvertUShort(BColor1.G)
    triVert(1).Blue = ConvertUShort(BColor1.B)

    triVert(2).X = 0
    triVert(2).Y = CanRect.D
    triVert(2).Red = ConvertUShort(BColor2.R)
    triVert(2).Green = ConvertUShort(BColor2.G)
    triVert(2).Blue = ConvertUShort(BColor2.B)
    
    Call GradientBack
    
    triVert(0).X = CanRect.R
    triVert(0).Y = CanRect.D
    triVert(0).Red = ConvertUShort(BColor2.R)
    triVert(0).Green = ConvertUShort(BColor2.G)
    triVert(0).Blue = ConvertUShort(BColor2.B)
    
    Call GradientBack

End Sub

Public Sub Gradient2()
    
    triVert(0).X = 0
    triVert(0).Y = 0
    triVert(0).Red = ConvertUShort(BColor1.R)
    triVert(0).Green = ConvertUShort(BColor1.G)
    triVert(0).Blue = ConvertUShort(BColor1.B)
    
    triVert(1).X = CanRect.R
    triVert(1).Y = 0
    triVert(1).Red = ConvertUShort(BColor2.R)
    triVert(1).Green = ConvertUShort(BColor2.G)
    triVert(1).Blue = ConvertUShort(BColor2.B)

    triVert(2).X = 0
    triVert(2).Y = CanRect.D
    triVert(2).Red = ConvertUShort(BColor2.R)
    triVert(2).Green = ConvertUShort(BColor2.G)
    triVert(2).Blue = ConvertUShort(BColor2.B)
    
    Call GradientBack
    
    triVert(0).X = CanRect.R
    triVert(0).Y = CanRect.D
    triVert(0).Red = ConvertUShort(BColor1.R)
    triVert(0).Green = ConvertUShort(BColor1.G)
    triVert(0).Blue = ConvertUShort(BColor1.B)
    
    Call GradientBack

End Sub

Public Sub Gradient3()
    
    triVert(0).X = 0
    triVert(0).Y = 0
    triVert(0).Red = ConvertUShort(BColor1.R)
    triVert(0).Green = ConvertUShort(BColor1.G)
    triVert(0).Blue = ConvertUShort(BColor1.B)
    
    triVert(1).X = CanRect.R
    triVert(1).Y = 0
    triVert(1).Red = ConvertUShort(BColor1.R)
    triVert(1).Green = ConvertUShort(BColor1.G)
    triVert(1).Blue = ConvertUShort(BColor1.B)

    triVert(2).X = OriginX
    triVert(2).Y = OriginY
    triVert(2).Red = ConvertUShort(BColor2.R)
    triVert(2).Green = ConvertUShort(BColor2.G)
    triVert(2).Blue = ConvertUShort(BColor2.B)
    
    Call GradientBack
    
    triVert(0).X = CanRect.R
    triVert(0).Y = CanRect.D
    triVert(0).Red = ConvertUShort(BColor1.R)
    triVert(0).Green = ConvertUShort(BColor1.G)
    triVert(0).Blue = ConvertUShort(BColor1.B)
    
    Call GradientBack

    triVert(1).X = 0
    triVert(1).Y = CanRect.D
    triVert(1).Red = ConvertUShort(BColor1.R)
    triVert(1).Green = ConvertUShort(BColor1.G)
    triVert(1).Blue = ConvertUShort(BColor1.B)
    
    Call GradientBack
    
    triVert(0).X = 0
    triVert(0).Y = 0
    triVert(0).Red = ConvertUShort(BColor1.R)
    triVert(0).Green = ConvertUShort(BColor1.G)
    triVert(0).Blue = ConvertUShort(BColor1.B)
    
    Call GradientBack

End Sub

