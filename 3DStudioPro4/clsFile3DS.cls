VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFile3DS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MAIN3DS                 As Long = &H4D4D
Private Const EDIT3DS                 As Long = &H3D3D
Private Const EDIT_OBJECT             As Long = &H4000
Private Const EDIT_MATERIAL           As Long = &HAFFF

Private Const OBJ_TRIMESH             As Long = &H4100
Private Const FILE_VERSION            As Long = &H2
Private Const MESH_VERSION            As Long = &H3D3E

Private Const MAT_DIFFUSE             As Long = &HA020
Private Const MAT_NAME                As Long = &HA000
Private Const MAT_TEXMAP              As Long = &HA200
Private Const MAT_MAPNAME             As Long = &HA300
Private Const MAT_MAP_TILING          As Long = &HA351
Private Const MAT_MAP_USCALE          As Long = &HA354
Private Const MAT_MAP_VSCALE          As Long = &HA356
Private Const MAT_MAP_UOFFSET         As Long = &HA358
Private Const MAT_MAP_VOFFSET         As Long = &HA35A
Private Const MAT_MAP_ANG             As Long = &HA35C
Private Const MAT_TRANSPARENCY        As Long = &HA050

Private Const TRI_VERTEXL             As Long = &H4110
Private Const TRI_FACEL1              As Long = &H4120
Private Const TRI_MATERIAL            As Long = &H4130
Private Const TRI_MAPPINGCOORS        As Long = &H4140


Private Const NULL_CHUNK              As Long = &H0

Private Type CHUNK
    Header      As Integer '2
    Length      As Long    '4
End Type

Private hFile   As Long
Private Pointer As Long
Private ObjName As String
Private Chnk    As CHUNK
Private ErrorLoad As Boolean

Public Function Read3DS(FileName As String) As Boolean
        
    On Error Resume Next
    
    g_idxMat = -1
    g_idxMesh = -1
        
    Call ResetCameraParameters
    Call ResetLightParameters
    
    Pointer = 1
    hFile = FreeFile
    Open FileName For Binary As #hFile
        While Not EOF(hFile)
            If ErrorLoad = False Then
                Call ReadChunk
            Else
                GoTo Jump
            End If
        Wend
    Close #hFile
    Read3DS = True
    Exit Function
    
Jump:
    Close #hFile
    Read3DS = False

End Function

Private Sub ReadChunk()

    On Error Resume Next
    
    Get #hFile, Pointer, Chnk
    Select Case Chnk.Header

'GO TO SUBCHUNK (skip)
        Case MAIN3DS, EDIT3DS, EDIT_MATERIAL, OBJ_TRIMESH, MAT_TEXMAP, NULL_CHUNK:
                                Chnk.Length = 6 '6=len(chnk)
        Case FILE_VERSION:      Call ReadFileVersion
        Case MESH_VERSION:      Call ReadMeshVersion

'MATERIALS
        Case MAT_NAME:          Call ReadMaterialName
        Case MAT_DIFFUSE:       Call ReadDiffuseColor
        Case MAT_MAPNAME:       Call ReadMapName
        Case MAT_MAP_USCALE:    Call ReadUScale
        Case MAT_MAP_VSCALE:    Call ReadVScale
        Case MAT_MAP_UOFFSET:   Call ReadUOffset
        Case MAT_MAP_VOFFSET:   Call ReadVOffset
        Case MAT_MAP_ANG:       Call ReadMapAng
        Case MAT_TRANSPARENCY:  Call ReadTransparency
        
'OBJECTS
        Case EDIT_OBJECT:       Call ReadObjectName
            
'READ MESH STUFF
        Case TRI_VERTEXL:       Call ReadVertex
        Case TRI_FACEL1:        Call ReadFace
        Case TRI_MATERIAL:      Call ReadFaceMaterial
        Case TRI_MAPPINGCOORS:  Call ReadMappingCoors
    End Select
    Pointer = Pointer + Chnk.Length
    
End Sub
        
Private Sub ReadFileVersion()

    Get #hFile, , FileVer
    
End Sub

Private Sub ReadMeshVersion()

    Get #hFile, , MeshVer

End Sub

Private Sub ReadMaterialName()
    
    Dim strRead As String
    
    strRead = ReadString
    g_idxMat = g_idxMat + 1
    ReDim Preserve Materials(g_idxMat)
    Call ResetMaterialParameters(g_idxMat)
    Materials(g_idxMat).MatName = strRead
    Chnk.Length = Len(strRead) + 7
    
End Sub

Private Sub ReadDiffuseColor()
    
    Dim intRead As Integer
    Dim lngRead As Long
    Dim bytRead As Byte

    Get #hFile, , intRead  'Data type (skip)
    Get #hFile, , lngRead  'Color Length (skip)
    Get #hFile, , bytRead: Materials(g_idxMat).DiffuseColor.R = CInt(bytRead)
    Get #hFile, , bytRead: Materials(g_idxMat).DiffuseColor.G = CInt(bytRead)
    Get #hFile, , bytRead: Materials(g_idxMat).DiffuseColor.B = CInt(bytRead)
    'Get #hFile, , intRead 'SubChunk ID (skip)
    
End Sub

Private Sub ReadMapName()
    
    Materials(g_idxMat).MapFileName = ReadString
    
End Sub

Private Sub ReadUScale()
    
    Get #hFile, , Materials(g_idxMat).UTiling

End Sub

Private Sub ReadVScale()
    
    Get #hFile, , Materials(g_idxMat).VTiling

End Sub

Private Sub ReadUOffset()
    
    Get #hFile, , Materials(g_idxMat).UOffset

End Sub

Private Sub ReadVOffset()
    
    Dim sngRead As Single
    
    Get #hFile, , sngRead
    Materials(g_idxMat).VOffset = -sngRead

End Sub

Private Sub ReadMapAng()
    
    Dim sngRead As Single
    
    Get #hFile, , sngRead
    Materials(g_idxMat).Angle = -sngRead

End Sub

Private Sub ReadTransparency()
    
    Dim intRead As Integer
    Dim lngRead As Long
    
    Get #hFile, , intRead 'skip
    Get #hFile, , lngRead 'skip
    Get #hFile, , intRead
    Materials(g_idxMat).Transparency = intRead * 0.01

End Sub

Private Sub ReadObjectName()

    ObjName = ReadString
    Chnk.Length = Len(ObjName) + 7

End Sub

Private Function ReadString() As String

    Dim strChar   As String
    Dim strAll    As String
    
    strChar = StrConv(InputB(1, #hFile), vbUnicode)
    While strChar <> Chr(0)
        strAll = strAll & strChar
        strChar = StrConv(InputB(1, #hFile), vbUnicode)
    Wend
    ReadString = strAll
  
End Function

Private Sub ReadVertex()
    
    Dim idx         As Long
    Dim intRead     As Integer
    Dim sngRead     As Single
    
    g_idxMesh = g_idxMesh + 1
    ReDim Preserve Meshs(g_idxMesh)

    With Meshs(g_idxMesh)
        .Name = ObjName
        Get #hFile, , intRead
        If intRead > 0 Then
            .idxVert = intRead - 1
            ReDim .Vertices(.idxVert)
            ReDim .Screen(.idxVert)
            For idx = 0 To .idxVert
                Get #hFile, , .Vertices(idx).Vectors.X
                Get #hFile, , .Vertices(idx).Vectors.Y
                Get #hFile, , .Vertices(idx).Vectors.Z
                              .Vertices(idx).Vectors.W = 1
            Next
            Call CreateBox(g_idxMesh)
        Else
            .idxVert = -1 'error
        End If
    End With
        
End Sub

Private Sub ReadFace()
    
    Dim idx         As Long
    Dim intRead     As Integer
    Dim intTemp     As Integer
    Dim intABC()    As Integer

    With Meshs(g_idxMesh)
        Get #hFile, , intRead
        If intRead > 0 Then
            .idxFace = intRead - 1
            ReDim .Faces(.idxFace)
            ReDim intABC(3, .idxFace) '0=A, 1=B, 2=C, 3=AB BC CA
            Get #hFile, , intABC
            For idx = 0 To .idxFace
                .Faces(idx).A = CLng(intABC(0, idx))
                .Faces(idx).B = CLng(intABC(1, idx))
                .Faces(idx).C = CLng(intABC(2, idx))
                
                intTemp = intABC(3, idx) And 1 '0000000X - CA
                .Faces(idx).CA = CBool(intTemp)
                
                intTemp = intABC(3, idx) And 2 '000000X0 - BC
                intTemp = intTemp / 2
                .Faces(idx).BC = CBool(intTemp)
                
                intTemp = intABC(3, idx) And 4 '00000X00 - AB
                intTemp = intTemp / 4
                .Faces(idx).AB = CBool(intTemp)
                
                intTemp = intABC(3, idx) And 8 '0000X000 - Mapping
                intTemp = intTemp / 8
                .Faces(idx).Mapping = CBool(intTemp)
            Next
            .FaceInfo = True
        Else
            .idxFace = -1 'error
        End If
        Chnk.Length = CLng(intRead) * 8 + 8
    End With
    
End Sub

Private Sub ReadFaceMaterial()
    
    Dim idxMat      As Integer
    Dim idxFace     As Integer
    Dim intRead     As Integer
    Dim intReadA()  As Integer
    Dim strRead     As String
    
    strRead = ReadString
    Get #hFile, , intRead
    If intRead > 0 Then
        ReDim intReadA(intRead - 1)
        Get #hFile, , intReadA
        For idxMat = 0 To g_idxMat
            If Materials(idxMat).MatName = strRead Then
                For idxFace = 0 To UBound(intReadA)
                    Meshs(g_idxMesh).Faces(intReadA(idxFace)).idxMat = idxMat
                Next
                Exit For
            End If
        Next
    End If
    
End Sub

Private Sub ReadMappingCoors()

    Dim idx         As Long
    Dim intRead     As Integer

    On Error Resume Next

    With Meshs(g_idxMesh)
        Get #hFile, , intRead
        .idxTVert = intRead - 1
        ReDim .TVerts(.idxTVert)
        ReDim .TScreen(.idxTVert)
        For idx = 0 To .idxTVert
            Get #hFile, , .TVerts(idx).U
            Get #hFile, , .TVerts(idx).V
        Next
        .MapCoorsOK = True
    End With

End Sub

