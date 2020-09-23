Attribute VB_Name = "modSort"
Option Explicit

Public Function SortVisibleFaces() As Long

    Dim idxMesh     As Integer
    Dim idxFace     As Long
    Dim idxOrder    As Long

    idxOrder = -1
    Erase MeshsOrder
    For idxMesh = 0 To g_idxMesh
        With Meshs(idxMesh)
            For idxFace = 0 To .idxFace
                .Faces(idxFace).Normal = FaceNormal(.Vertices(.Faces(idxFace).A).VectorsT, _
                                                    .Vertices(.Faces(idxFace).B).VectorsT, _
                                                    .Vertices(.Faces(idxFace).C).VectorsT)
                If .Faces(idxFace).Normal.Z > 0 Then
                    idxOrder = idxOrder + 1
                    ReDim Preserve MeshsOrder(idxOrder)
                    MeshsOrder(idxOrder).ZValue = (.Vertices(.Faces(idxFace).A).VectorsT.Z + _
                                                   .Vertices(.Faces(idxFace).B).VectorsT.Z + _
                                                   .Vertices(.Faces(idxFace).C).VectorsT.Z)
                    MeshsOrder(idxOrder).idxFaceO = idxFace
                    MeshsOrder(idxOrder).idxMeshO = idxMesh
                End If
            Next
        End With
    Next
    If idxOrder > -1 Then SortFaces 0, idxOrder
    SortVisibleFaces = idxOrder

End Function

Public Function SortUnVisibleFaces() As Long

    Dim idxMesh     As Integer
    Dim idxFace     As Long
    Dim idxOrder    As Long

    idxOrder = -1
    Erase MeshsOrder
    For idxMesh = 0 To g_idxMesh
        With Meshs(idxMesh)
            For idxFace = 0 To .idxFace
                .Faces(idxFace).Normal = FaceNormal(.Vertices(.Faces(idxFace).A).VectorsT, _
                                                    .Vertices(.Faces(idxFace).B).VectorsT, _
                                                    .Vertices(.Faces(idxFace).C).VectorsT)
                If .Faces(idxFace).Normal.Z < 0 Then
                    idxOrder = idxOrder + 1
                    ReDim Preserve MeshsOrder(idxOrder)
                    MeshsOrder(idxOrder).ZValue = (.Vertices(.Faces(idxFace).A).VectorsT.Z + _
                                                   .Vertices(.Faces(idxFace).B).VectorsT.Z + _
                                                   .Vertices(.Faces(idxFace).C).VectorsT.Z)
                    MeshsOrder(idxOrder).idxFaceO = idxFace
                    MeshsOrder(idxOrder).idxMeshO = idxMesh
                End If
            Next
        End With
    Next
    If idxOrder > -1 Then SortFaces 0, idxOrder
    SortUnVisibleFaces = idxOrder

End Function

Private Sub SortFaces(ByVal First As Long, ByVal Last As Long)

    Dim FirstIdx    As Long
    Dim MidIdx      As Long
    Dim LastIdx     As Long
    Dim MidVal      As Single
    Dim TempOrder   As ORDER
    
    If (First < Last) Then
            MidIdx = (First + Last) * 0.5
            MidVal = MeshsOrder(MidIdx).ZValue
            FirstIdx = First
            LastIdx = Last
            Do
                Do While MeshsOrder(FirstIdx).ZValue < MidVal
                    FirstIdx = FirstIdx + 1
                Loop
                Do While MeshsOrder(LastIdx).ZValue > MidVal
                    LastIdx = LastIdx - 1
                Loop
                If (FirstIdx <= LastIdx) Then
                    TempOrder = MeshsOrder(LastIdx)
                    MeshsOrder(LastIdx) = MeshsOrder(FirstIdx)
                    MeshsOrder(FirstIdx) = TempOrder
                    FirstIdx = FirstIdx + 1
                    LastIdx = LastIdx - 1
                End If
            Loop Until FirstIdx > LastIdx

            If (LastIdx <= MidIdx) Then
                SortFaces First, LastIdx
                SortFaces FirstIdx, Last
            Else
                SortFaces FirstIdx, Last
                SortFaces First, LastIdx
            End If
    End If

End Sub
