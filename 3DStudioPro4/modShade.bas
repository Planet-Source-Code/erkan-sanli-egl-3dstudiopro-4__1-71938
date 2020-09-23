Attribute VB_Name = "modShade"
Option Explicit

Public Function Shade(idxMesh As Integer, idxFace As Long, Normal As VECTOR, idxVert As Long) As COLORRGB
    
    Dim Alpha           As Single 'Not transparency value , this is angle
    Dim Beta            As Single
    Dim Epsilon         As Single
    Dim Gamma           As Single
    Dim VectorT         As VECTOR
    Dim Color           As COLORRGB
    
    VectorT = Meshs(idxMesh).Vertices(idxVert).VectorsT
    Color = Materials(Meshs(idxMesh).Faces(idxFace).idxMat).DiffuseColor
    With Lights(0)
        If .Enabled = True Then
            .OriginT = MatrixMultiplyVector(Cameras(0).ViewMatrix, .Origin)
            .DirectionT = MatrixMultiplyVector(Cameras(0).ViewMatrix, .Direction)

            If .Falloff > 0 Then
                Epsilon = VectorAngle(.DirectionT, VectorSubtract(VectorT, .OriginT))
                If .Falloff <> .Hotspot Then
                    Epsilon = (.Falloff - Epsilon) / (.Falloff - .Hotspot)
                    If Epsilon < 0 Then Epsilon = 0
                    If Epsilon > 1 Then Epsilon = 1
                Else
                    Exit Function
                End If
            Else
                Epsilon = 1
            End If
            
            Alpha = VectorAngle(VectorSubtract(.OriginT, VectorT), Normal) * .Diffusion
            If Alpha < 0 Then Alpha = 0
            Beta = VectorAngle(VectorReflect(VectorSubtract(VectorT, .OriginT), Normal), VectorSubtract(Cameras(0).WorldPosition, VectorT)) * .Specular
            If Beta < 0 Then Beta = 0
            
            If .AttenEnable Then
                If .DarkRange <> .BrightRange Then
                    Gamma = (.DarkRange - VectorDistance(VectorT, .OriginT)) / (.DarkRange - .BrightRange)
                    If Gamma < 0 Then Gamma = 0
                    If Gamma > 1 Then Gamma = 1
                End If
            Else
                Gamma = 1
            End If
            Shade = ColorScale(.Color, (Alpha + Beta) * Gamma * Epsilon)
            Shade = ColorInterpolate(Shade, Color, 0.5)
            Shade = ColorAdd(Shade, .Ambiance)
        End If
    End With
    Call ColorLimit(Shade)

End Function
