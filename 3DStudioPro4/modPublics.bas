Attribute VB_Name = "modPublics"
Option Explicit

Global Const ApproachVal As Single = 0.000001

Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal hStretchMode As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Enum VisualStyle
    Dot
    Box
    Wireframe
    WireframeMap
    WireframeBFC
    Flat
    Gouraud
    Mapped
    Photo
End Enum

Public Enum CanStat
    CanOut       'Face's all points are out canvas area
    CanClip      'Face's any edge intersection canvas edge
    CanIn        'Face's all points are in canvas area
End Enum

Public Enum BackType
    Blank
    SingleColor
    GradType1
    GradType2
    GradType3
    BackPicture
End Enum

Public Type CANVASRECT
    L               As Long     'L eft
    R               As Long     'R ight
    T               As Long     'T op
    D               As Long     'D own
End Type

Public Type VECTOR
    X               As Single
    Y               As Single
    Z               As Single
    W               As Single
End Type

Public Type COLORRGB
    R               As Integer
    G               As Integer
    B               As Integer
End Type

Public Type VERTEX
    Vectors         As VECTOR
    VectorsT        As VECTOR
    VectorsS        As VECTOR
    ColorS          As COLORRGB
    Used            As Boolean
End Type

Public Type FACE
    A               As Long
    B               As Long
    C               As Long
    AB              As Boolean
    BC              As Boolean
    CA              As Boolean
    Mapping         As Boolean
    Normal          As VECTOR
    idxMat          As Integer
End Type

Public Type POINTAPI
    X               As Long
    Y               As Long
End Type

Public Type ORDER
    idxMeshO        As Integer
    idxFaceO        As Long
    ZValue          As Single
End Type

Public Type MATRIX
    rc11 As Single: rc12 As Single: rc13 As Single: rc14 As Single
    rc21 As Single: rc22 As Single: rc23 As Single: rc24 As Single
    rc31 As Single: rc32 As Single: rc33 As Single: rc34 As Single
    rc41 As Single: rc42 As Single: rc43 As Single: rc44 As Single
End Type

Public Type MAPCOORS
    U               As Single
    V               As Single
End Type

Public Type MESH
    Name            As String
    
    idxVert         As Long
    idxFace         As Integer
    idxTVert        As Long
    
    
    Vertices()      As VERTEX
    Faces()         As FACE
    Screen()        As POINTAPI
    FaceV()         As ORDER
    FaceInfo        As Boolean
    MapCoorsOK      As Boolean
    
    TVerts()        As MAPCOORS
    TScreen()       As MAPCOORS

    Box(7)          As VERTEX
    BoxScreen(7)    As POINTAPI
    BoxFace(11)     As FACE
    
'    Rotation        As VECTOR
'    Translation     As VECTOR
'    Scales          As VECTOR
'    WorldMatrix     As MATRIX
End Type

Public Type CAMERA
    WorldPosition   As VECTOR
    LookAtPoint     As VECTOR
    VUP             As VECTOR
    PRP             As VECTOR
    Zoom            As Single
    FOV             As Single
    Yaw             As Single
    ClipFar         As Single
    ClipNear        As Single
    ViewMatrix      As MATRIX
End Type

Public Type LIGHT
    Origin      As VECTOR
    OriginT     As VECTOR
    Direction   As VECTOR
    DirectionT  As VECTOR
    
    Color       As COLORRGB
    Falloff     As Single
    Hotspot     As Single
    BrightRange As Single
    DarkRange   As Single

    Ambiance    As COLORRGB
    Diffusion   As Single
    Specular    As Single
    AttenEnable As Boolean
    Enabled     As Boolean
End Type

Public Type DIB
    mapDIB          As clsDIB
    mapArray()      As Long
End Type

Public Type MATERIAL
    MatName         As String
    
    Ambiance        As COLORRGB
    Diffusion       As Single
    Specular        As Single
    DiffuseColor    As COLORRGB
    
    MapUse          As Boolean
    MapName         As String
    MapFileName     As String
    UOffset         As Single
    VOffset         As Single
    UTiling         As Single
    VTiling         As Single
    Angle           As Single
    Transparency    As Single
    Mirroring       As Byte
    Flipping        As Byte
    
    dibTex          As DIB
    dibTexT         As DIB
    
    TwoSided        As Boolean
End Type

Public Meshs()          As MESH
Public MeshsOrder()     As ORDER
Public MeshsRotation    As VECTOR
Public MeshsTranslation As VECTOR
Public MeshsScales      As VECTOR
Public MeshsWorldMatrix As MATRIX
Public FileVer          As Long
Public MeshVer          As Long

Public Cameras(0)   As CAMERA
Public Lights(0)    As LIGHT
Public Materials()  As MATERIAL

Public CanRect      As CANVASRECT

Public cdiLoad      As clsCommonDialog
Public cf3DS        As clsFile3DS
'Public cfOBJ        As clsFileOBJ

Public dibCanvas    As DIB
Public dibBack      As DIB

Public CanvasWidth  As Long
Public CanvasHeight As Long

Public LoadComplete As Boolean
Public VStyle       As VisualStyle
Public Path3DS      As String
Public g_blnDoubleSide   As Boolean
Public g_blnLight   As Boolean

Public g_idxMesh    As Integer
Public g_idxMat     As Integer

Public OriginX     As Long
Public OriginY     As Long

Public BColor1   As COLORRGB
Public BColor2   As COLORRGB
Public BType        As BackType
Public BFileName    As String

Private ReLocalTra As VECTOR
Private ReLocalSca As VECTOR



Public Sub ResetMeshParameters()

    MeshsRotation = VectorSet(0, 0, 0)
    MeshsTranslation = ReLocalTra
    MeshsScales = ReLocalSca

End Sub

Public Sub ResetCameraParameters()

    With Cameras(0)
        .WorldPosition = VectorSet(0, 0, 120)
        .LookAtPoint = VectorSet(0, 0, 0)
        .VUP = VectorSet(0, 1, 0)
        .PRP = VectorSet(0, 0, 1)
        .Zoom = 1
        .FOV = ConvertZoomtoFOV(.Zoom)
        .Yaw = 0
        .ClipFar = 0
        .ClipNear = -1000
    End With
    
End Sub

Public Sub ResetLightParameters()

    With Lights(0)
        .Origin = VectorSet(0, 0, 120)
        .Direction = VectorSet(0, 0, 0)
        .Color = ColorSet(130, 130, 130)
        .Ambiance = ColorSet(0, 0, 0)
        .Diffusion = 2
        .Specular = 2
        .Falloff = 5
        .Hotspot = 0.5
        .DarkRange = 100
        .BrightRange = 0
        .AttenEnable = False 'True
        .Enabled = True
    End With

End Sub

Public Sub ResetMaterialParameters(idxMat As Integer)

    With Materials(idxMat)
        .MatName = "Default"
        
        '.Ambiance
        '.Diffusion
        '.Specular
        
        .MapName = "Default"
        .MapFileName = ""
        .UOffset = 0
        .VOffset = 0
        .UTiling = 1
        .VTiling = 1
        .Angle = 0
        
        .Mirroring = 0
        .Flipping = 0
        
        Set .dibTex.mapDIB = Nothing
        Set .dibTexT.mapDIB = Nothing
        Erase .dibTex.mapArray
        Erase .dibTexT.mapArray
        
        .TwoSided = False
    End With
        
End Sub

Public Sub CreateBox(idxMesh As Integer)

    Dim idx         As Long
    Dim MinVector   As VECTOR
    Dim MaxVector   As VECTOR

    With Meshs(idxMesh)
        MaxVector = .Vertices(0).Vectors
        MinVector = .Vertices(0).Vectors
        For idx = 1 To .idxVert
            If MaxVector.X < .Vertices(idx).Vectors.X Then MaxVector.X = .Vertices(idx).Vectors.X
            If MaxVector.Y < .Vertices(idx).Vectors.Y Then MaxVector.Y = .Vertices(idx).Vectors.Y
            If MaxVector.Z < .Vertices(idx).Vectors.Z Then MaxVector.Z = .Vertices(idx).Vectors.Z
            If MinVector.X > .Vertices(idx).Vectors.X Then MinVector.X = .Vertices(idx).Vectors.X
            If MinVector.Y > .Vertices(idx).Vectors.Y Then MinVector.Y = .Vertices(idx).Vectors.Y
            If MinVector.Z > .Vertices(idx).Vectors.Z Then MinVector.Z = .Vertices(idx).Vectors.Z
        Next
        
'           +Y
'           ^
'           3--------2
'          /|       /|
'         / |      / |
'        7--+-----6  |
'        |  |     |  |
'        |  0-----+--1 >+X
'        | /      | /
'        |/       |/
'        4--------5
'       /
'      +Z
        .Box(0).Vectors = MinVector
        .Box(1).Vectors = MinVector:  .Box(1).Vectors.X = MaxVector.X
        .Box(2).Vectors = MaxVector:  .Box(2).Vectors.Z = MinVector.Z
        .Box(3).Vectors = MinVector:  .Box(3).Vectors.Y = MaxVector.Y
        .Box(4).Vectors = MinVector:  .Box(4).Vectors.Z = MaxVector.Z
        .Box(5).Vectors = MaxVector:  .Box(5).Vectors.Y = MinVector.Y
        .Box(6).Vectors = MaxVector
        .Box(7).Vectors = MaxVector:  .Box(7).Vectors.X = MinVector.X
        .BoxFace(0).A = 0: .BoxFace(0).B = 1: .BoxFace(0).C = 2: .BoxFace(0).AB = True: .BoxFace(0).BC = True: .BoxFace(0).CA = False
        .BoxFace(1).A = 2: .BoxFace(1).B = 3: .BoxFace(1).C = 0: .BoxFace(1).AB = True: .BoxFace(1).BC = True: .BoxFace(1).CA = False
        .BoxFace(2).A = 4: .BoxFace(2).B = 5: .BoxFace(2).C = 6: .BoxFace(2).AB = True: .BoxFace(2).BC = True: .BoxFace(2).CA = False
        .BoxFace(3).A = 6: .BoxFace(3).B = 7: .BoxFace(3).C = 4: .BoxFace(3).AB = True: .BoxFace(3).BC = True: .BoxFace(3).CA = False
        
        .BoxFace(4).A = 1: .BoxFace(4).B = 2: .BoxFace(4).C = 6: .BoxFace(4).AB = True: .BoxFace(4).BC = True: .BoxFace(4).CA = False
        .BoxFace(5).A = 6: .BoxFace(5).B = 5: .BoxFace(5).C = 1: .BoxFace(5).AB = True: .BoxFace(5).BC = True: .BoxFace(5).CA = False
        
        .BoxFace(6).A = 0: .BoxFace(6).B = 3: .BoxFace(6).C = 7: .BoxFace(6).AB = True: .BoxFace(6).BC = True: .BoxFace(6).CA = False
        .BoxFace(7).A = 7: .BoxFace(7).B = 4: .BoxFace(7).C = 0: .BoxFace(7).AB = True: .BoxFace(7).BC = True: .BoxFace(7).CA = False
        
        .BoxFace(8).A = 2: .BoxFace(8).B = 3: .BoxFace(8).C = 7: .BoxFace(8).AB = True: .BoxFace(8).BC = True: .BoxFace(8).CA = False
        .BoxFace(9).A = 7: .BoxFace(9).B = 6: .BoxFace(9).C = 2: .BoxFace(9).AB = True: .BoxFace(9).BC = True: .BoxFace(9).CA = False
        
        .BoxFace(10).A = 4: .BoxFace(10).B = 5: .BoxFace(10).C = 1: .BoxFace(10).AB = True: .BoxFace(10).BC = True: .BoxFace(10).CA = False
        .BoxFace(11).A = 1: .BoxFace(11).B = 0: .BoxFace(11).C = 4: .BoxFace(11).AB = True: .BoxFace(11).BC = True: .BoxFace(11).CA = False
    End With

End Sub

Public Function VerifyText(txt As TextBox) As Single

    VerifyText = IIf(IsNumeric(txt.Text), CSng(txt.Text), 0)

End Function

Public Sub ReLocaleMeshs()

    Dim idxMesh         As Integer
    Dim idxVect         As Long
    Dim idxBox          As Integer
    Dim MinVector       As VECTOR
    Dim MaxVector       As VECTOR
    Dim MinVectorAll    As VECTOR
    Dim MaxVectorAll    As VECTOR
    Dim CenterVectorAll As VECTOR
    Dim Dx As Single
    Dim Dy As Single
    Dim MaxH As Single
    
    MaxVectorAll = Meshs(0).Box(0).Vectors
    MinVectorAll = Meshs(0).Box(0).Vectors
    For idxMesh = 0 To g_idxMesh
        With Meshs(idxMesh)
            MaxVector = .Box(0).Vectors
            MinVector = .Box(0).Vectors
            For idxBox = 1 To UBound(.Box)
                If MaxVector.X < .Box(idxBox).Vectors.X Then MaxVector.X = .Box(idxBox).Vectors.X
                If MaxVector.Y < .Box(idxBox).Vectors.Y Then MaxVector.Y = .Box(idxBox).Vectors.Y
                If MaxVector.Z < .Box(idxBox).Vectors.Z Then MaxVector.Z = .Box(idxBox).Vectors.Z
                If MinVector.X > .Box(idxBox).Vectors.X Then MinVector.X = .Box(idxBox).Vectors.X
                If MinVector.Y > .Box(idxBox).Vectors.Y Then MinVector.Y = .Box(idxBox).Vectors.Y
                If MinVector.Z > .Box(idxBox).Vectors.Z Then MinVector.Z = .Box(idxBox).Vectors.Z
            Next
            If MaxVectorAll.X < MaxVector.X Then MaxVectorAll.X = MaxVector.X
            If MaxVectorAll.Y < MaxVector.Y Then MaxVectorAll.Y = MaxVector.Y
            If MaxVectorAll.Z < MaxVector.Z Then MaxVectorAll.Z = MaxVector.Z
            If MinVectorAll.X > MinVector.X Then MinVectorAll.X = MinVector.X
            If MinVectorAll.Y > MinVector.Y Then MinVectorAll.Y = MinVector.Y
            If MinVectorAll.Z > MinVector.Z Then MinVectorAll.Z = MinVector.Z
        End With
    Next
    CenterVectorAll.X = (MaxVectorAll.X + MinVectorAll.X) * 0.5
    CenterVectorAll.Y = (MaxVectorAll.Y + MinVectorAll.Y) * 0.5
    CenterVectorAll.Z = (MaxVectorAll.Z + MinVectorAll.Z) * 0.5
    CenterVectorAll.W = 1
    
    Dx = MaxVectorAll.X - MinVectorAll.X
    Dy = MaxVectorAll.Y - MinVectorAll.Y
    MaxH = IIf(Dx > Dy, Dx, Dy)
    MaxH = Div(CanRect.D, MaxH) * 0.05
    ReLocalTra = VectorSet(-CenterVectorAll.X, -CenterVectorAll.Y, -CenterVectorAll.Z)
    ReLocalSca = VectorSet(MaxH, MaxH, MaxH)
    Call ResetMeshParameters
    
End Sub

Public Sub CalculateUV(idxMesh As Integer, idxFace As Long)

    Dim idxMat      As Integer
    Dim texWidth    As Long
    Dim texHeight   As Long
    Dim sUTiling    As Single
    Dim sVTiling    As Single
    Dim sUOffset    As Single
    Dim sVOffset    As Single
    Dim Angle       As Single
    Dim sngCosine   As Single
    Dim sngSine     As Single

        With Meshs(idxMesh)
                idxMat = .Faces(idxFace).idxMat
                If Meshs(idxMesh).MapCoorsOK And Materials(idxMat).MapUse Then
                    texWidth = Materials(idxMat).dibTexT.mapDIB.Width - 1
                    texHeight = Materials(idxMat).dibTexT.mapDIB.Height - 1
                    sUTiling = Materials(idxMat).UTiling * texWidth
                    sVTiling = Materials(idxMat).VTiling * texHeight
                    sUOffset = -Materials(idxMat).UOffset * texWidth
                    sVOffset = -Materials(idxMat).VOffset * texHeight
                    'A
                    .TScreen(.Faces(idxFace).A).U = sUOffset + (.TVerts(.Faces(idxFace).A).U * sUTiling)
                    .TScreen(.Faces(idxFace).A).V = sVOffset + (.TVerts(.Faces(idxFace).A).V * sVTiling)
                    If Materials(idxMat).Angle <> 0 Then
                        Angle = ConvertDeg2Rad(Materials(idxMat).Angle)
                        sngCosine = Round(Cos(Angle), 6)
                        sngSine = Round(Sin(Angle), 6)
                        .TScreen(.Faces(idxFace).A).U = (sngCosine * .TScreen(.Faces(idxFace).A).U) - _
                                                        (sngSine * .TScreen(.Faces(idxFace).A).V)
                        .TScreen(.Faces(idxFace).A).V = (sngSine * .TScreen(.Faces(idxFace).A).U) + _
                                                        (sngCosine * .TScreen(.Faces(idxFace).A).V)
                    End If
                    'B
                    .TScreen(.Faces(idxFace).B).U = sUOffset + (.TVerts(.Faces(idxFace).B).U * sUTiling)
                    .TScreen(.Faces(idxFace).B).V = sVOffset + (.TVerts(.Faces(idxFace).B).V * sVTiling)
                    If Materials(idxMat).Angle <> 0 Then
                        Angle = ConvertDeg2Rad(Materials(idxMat).Angle)
                        sngCosine = Round(Cos(Angle), 6)
                        sngSine = Round(Sin(Angle), 6)
                        .TScreen(.Faces(idxFace).B).U = (sngCosine * .TScreen(.Faces(idxFace).B).U) - _
                                                        (sngSine * .TScreen(.Faces(idxFace).B).V)
                        .TScreen(.Faces(idxFace).B).V = (sngSine * .TScreen(.Faces(idxFace).B).U) + _
                                                        (sngCosine * .TScreen(.Faces(idxFace).B).V)
                    End If
                    'C
                    .TScreen(.Faces(idxFace).C).U = sUOffset + (.TVerts(.Faces(idxFace).C).U * sUTiling)
                    .TScreen(.Faces(idxFace).C).V = sVOffset + (.TVerts(.Faces(idxFace).C).V * sVTiling)
                    If Materials(idxMat).Angle <> 0 Then
                        Angle = ConvertDeg2Rad(Materials(idxMat).Angle)
                        sngCosine = Round(Cos(Angle), 6)
                        sngSine = Round(Sin(Angle), 6)
                        .TScreen(.Faces(idxFace).C).U = (sngCosine * .TScreen(.Faces(idxFace).C).U) - _
                                                        (sngSine * .TScreen(.Faces(idxFace).C).V)
                        .TScreen(.Faces(idxFace).C).V = (sngSine * .TScreen(.Faces(idxFace).C).U) + _
                                                        (sngCosine * .TScreen(.Faces(idxFace).C).V)
                    End If
                End If
        End With

End Sub


