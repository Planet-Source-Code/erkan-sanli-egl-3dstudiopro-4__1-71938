Attribute VB_Name = "modVisualisation"
Option Explicit

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function Polygon Lib "gdi32.dll" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'Alpha= 0:opaque ,0.01~1:transparent

Public Sub Render(hDC As Long)
        
    DoEvents
    If BType = Blank Then
        
        dibCanvas.mapDIB.Clear dibCanvas.mapArray
    Else
        BitBlt dibCanvas.mapDIB.hDC, 0, 0, CanvasWidth, CanvasHeight, dibBack.mapDIB.hDC, 0, 0, vbSrcCopy
    End If
    Select Case VStyle
        
        Case Dot:           Call RenderDot(g_blnLight)
        
        Case Box:           Call RenderBox
        
        Case Wireframe:     Call RenderWireframe(g_blnLight)
        Case WireframeMap:  Call RenderWireframeTex
        Case WireframeBFC:  Call RenderWireframeBFC(g_blnLight)
        
        Case Flat:          If g_blnDoubleSide Then If SortUnVisibleFaces >= 0 Then Call RenderFlat(g_blnLight)
                            If SortVisibleFaces >= 0 Then Call RenderFlat(g_blnLight)
                
        Case Gouraud:       If g_blnDoubleSide Then If SortUnVisibleFaces >= 0 Then Call RenderGouraud
                            If SortVisibleFaces >= 0 Then Call RenderGouraud
        
        Case Mapped:        If g_blnDoubleSide Then If SortUnVisibleFaces >= 0 Then Call RenderTexture(g_blnLight)
                            If SortVisibleFaces >= 0 Then Call RenderTexture(g_blnLight)
                
        Case Photo:         If g_blnDoubleSide Then If SortUnVisibleFaces >= 0 Then Call RenderPhoto
                            If SortVisibleFaces >= 0 Then Call RenderPhoto
    
    End Select
    'Call DrawFrame
    BitBlt hDC, 0, 0, CanvasWidth, CanvasHeight, dibCanvas.mapDIB.hDC, 0, 0, vbSrcCopy

End Sub
'ok
Private Sub RenderDot(Lit As Boolean)
    
    Dim idxMesh     As Integer
    Dim idxPoint    As Long
    Dim lColor      As Long
    
    Call GetFaceColor(Lit)
    For idxMesh = 0 To g_idxMesh
        With Meshs(idxMesh)
            For idxPoint = 0 To .idxVert
                If PointInCanvas(.Screen(idxPoint)) And ZValInCamera(.Vertices(idxPoint).VectorsT.Z) Then
                    lColor = ColorBGRToLong(.Vertices(idxPoint).ColorS)
                    Call DrawDot(.Screen(idxPoint), lColor)
                End If
            Next idxPoint
        End With
    Next idxMesh
    
End Sub
'ok
Private Sub RenderBox()
    
    Dim idxMesh     As Integer
    Dim idxFace     As Long
    Dim lColor      As Long
    Dim FStat       As CanStat
    
    For idxMesh = 0 To g_idxMesh
        With Meshs(idxMesh)
            If g_idxMat = -1 Then
                lColor = RGB(128, 128, 128)
            Else
                lColor = ColorRGBToLong(Materials(Meshs(idxMesh).Faces(0).idxMat).DiffuseColor)
            End If
            For idxFace = 0 To UBound(.BoxFace)
                FStat = FaceStat(idxMesh, idxFace, True, True)
                Select Case FStat
                    Case CanIn
                        Call DrawLine(.BoxScreen(.BoxFace(idxFace).A), .BoxScreen(.BoxFace(idxFace).B), lColor)
                        Call DrawLine(.BoxScreen(.BoxFace(idxFace).B), .BoxScreen(.BoxFace(idxFace).C), lColor)
                    Case CanClip
                        Call DrawClip(lColor)
                End Select
            Next idxFace
        End With
    Next idxMesh

End Sub

Private Sub RenderWireframe(Lit As Boolean)

    Dim idxMesh     As Integer
    Dim idxFace     As Long
    Dim lColor      As Long
    Dim FStat       As CanStat

    Call GetFaceColor(Lit)
    For idxMesh = 0 To g_idxMesh
        With Meshs(idxMesh)
            If .FaceInfo Then
                For idxFace = 0 To .idxFace
                    If IsInCamera(idxMesh, idxFace) Then
                        FStat = FaceStat(idxMesh, idxFace, , True)
                        Select Case FStat
                            Case CanIn
                                If .Faces(idxFace).AB Then
                                    lColor = ColorRGBToLong(.Vertices(.Faces(idxFace).A).ColorS)
                                    Call DrawLine(.Screen(.Faces(idxFace).A), .Screen(.Faces(idxFace).B), lColor)
                                End If
                                If .Faces(idxFace).BC Then
                                    lColor = ColorRGBToLong(.Vertices(.Faces(idxFace).B).ColorS)
                                    Call DrawLine(.Screen(.Faces(idxFace).B), .Screen(.Faces(idxFace).C), lColor)
                                End If
                                If .Faces(idxFace).CA Then
                                    lColor = ColorRGBToLong(.Vertices(.Faces(idxFace).C).ColorS)
                                    Call DrawLine(.Screen(.Faces(idxFace).C), .Screen(.Faces(idxFace).A), lColor)
                                End If
                            Case CanClip
                                lColor = ColorRGBToLong(Materials(.Faces(idxFace).idxMat).DiffuseColor)
                                Call DrawClip(lColor)
                        End Select
                    End If
                Next idxFace
            Else
                For idxFace = 0 To .idxFace
                    If IsInCamera(idxMesh, idxFace) Then
                        FStat = FaceStat(idxMesh, idxFace, , True)
                        Select Case FStat
                            Case CanIn
                                lColor = ColorRGBToLong(.Vertices(.Faces(idxFace).A).ColorS)
                                Call DrawLine(.Screen(.Faces(idxFace).A), .Screen(.Faces(idxFace).B), lColor)
                                lColor = ColorRGBToLong(.Vertices(.Faces(idxFace).B).ColorS)
                                Call DrawLine(.Screen(.Faces(idxFace).B), .Screen(.Faces(idxFace).C), lColor)
                                lColor = ColorRGBToLong(.Vertices(.Faces(idxFace).C).ColorS)
                                Call DrawLine(.Screen(.Faces(idxFace).C), .Screen(.Faces(idxFace).A), lColor)
                            Case CanClip
                                lColor = ColorRGBToLong(.Vertices(.Faces(idxFace).C).ColorS)
                                Call DrawClip(lColor)
                        End Select
                    End If
                Next idxFace
            End If
        End With
    Next idxMesh
    
End Sub
'ok
Private Sub RenderWireframeBFC(Lit As Boolean)

    Dim idxOrder    As Long
    Dim idxMeshO    As Integer
    Dim idxFaceO    As Long
    Dim lColor      As Long
    Dim FStat       As CanStat

    If SortVisibleFaces < 0 Then Exit Sub
    Call GetFaceColorO(Lit)
    For idxOrder = 0 To UBound(MeshsOrder)
        idxMeshO = MeshsOrder(idxOrder).idxMeshO
        idxFaceO = MeshsOrder(idxOrder).idxFaceO
        If IsInCamera(idxMeshO, idxFaceO) Then
            With Meshs(idxMeshO)
                If .FaceInfo Then
                    FStat = FaceStat(idxMeshO, idxFaceO, , True)
                    Select Case FStat
                        Case CanIn
                            Call DrawTriangleFlatA(idxMeshO, idxFaceO, 1, False)
                            If .Faces(idxFaceO).AB Then
                                lColor = ColorRGBToLong(.Vertices(.Faces(idxFaceO).A).ColorS)
                                Call DrawLine(.Screen(.Faces(idxFaceO).A), .Screen(.Faces(idxFaceO).B), lColor)
                            End If
                            If .Faces(idxFaceO).BC Then
                                lColor = ColorRGBToLong(.Vertices(.Faces(idxFaceO).B).ColorS)
                                Call DrawLine(.Screen(.Faces(idxFaceO).B), .Screen(.Faces(idxFaceO).C), lColor)
                            End If
                            If .Faces(idxFaceO).CA Then
                                lColor = ColorRGBToLong(.Vertices(.Faces(idxFaceO).C).ColorS)
                                Call DrawLine(.Screen(.Faces(idxFaceO).C), .Screen(.Faces(idxFaceO).A), lColor)
                            End If
                        Case CanClip
                            Call DrawTriangleFlatA(idxMeshO, idxFaceO, 1, False)
                            lColor = ColorRGBToLong(.Vertices(.Faces(idxFaceO).A).ColorS)
                            Call DrawClip(lColor)
                    End Select
                Else
                    FStat = FaceStat(idxMeshO, idxFaceO, , True)
                    Select Case FStat
                        Case CanIn
                            Call DrawTriangleFlatA(idxMeshO, idxFaceO, 1, False)
                            lColor = ColorRGBToLong(.Vertices(.Faces(idxFaceO).A).ColorS)
                            Call DrawLine(.Screen(.Faces(idxFaceO).A), .Screen(.Faces(idxFaceO).B), lColor)
                            lColor = ColorRGBToLong(.Vertices(.Faces(idxFaceO).B).ColorS)
                            Call DrawLine(.Screen(.Faces(idxFaceO).B), .Screen(.Faces(idxFaceO).C), lColor)
                            lColor = ColorRGBToLong(.Vertices(.Faces(idxFaceO).C).ColorS)
                            Call DrawLine(.Screen(.Faces(idxFaceO).C), .Screen(.Faces(idxFaceO).A), lColor)
                        Case CanClip
                            Call DrawTriangleFlatA(idxMeshO, idxFaceO, 1, False)
                            lColor = ColorRGBToLong(.Vertices(.Faces(idxFaceO).A).ColorS)
                            Call DrawClip(lColor)
                    End Select
                End If
            End With
        End If
    Next idxOrder

End Sub
'
Private Sub RenderWireframeTex()

    Dim idxMesh     As Integer
    Dim idxFace     As Long
    Dim lColor      As Long
    Dim FStat       As CanStat
    
    Call GetFaceColor(False)
    For idxMesh = 0 To g_idxMesh
        With Meshs(idxMesh)
            If .MapCoorsOK Then
                If .FaceInfo Then
                    For idxFace = 0 To .idxFace
                        If IsInCamera(idxMesh, idxFace) Then
                            FStat = FaceStat(idxMesh, idxFace, , True)
                            If FStat <> CanOut Then
                                If Materials(.Faces(idxFace).idxMat).MapUse Then
                                    Call CalculateUV(idxMesh, idxFace)
                                    If .Faces(idxFace).AB Then
                                        Call DrawLineTex(.Screen(.Faces(idxFace).A).X, .Screen(.Faces(idxFace).A).Y, _
                                                         .Screen(.Faces(idxFace).B).X, .Screen(.Faces(idxFace).B).Y, _
                                                         .TScreen(.Faces(idxFace).A).U, .TScreen(.Faces(idxFace).A).V, _
                                                         .TScreen(.Faces(idxFace).B).U, .TScreen(.Faces(idxFace).B).V, _
                                                         .Faces(idxFace).idxMat)
                                    End If
                                    If .Faces(idxFace).BC Then
                                        Call DrawLineTex(.Screen(.Faces(idxFace).B).X, .Screen(.Faces(idxFace).B).Y, _
                                                         .Screen(.Faces(idxFace).C).X, .Screen(.Faces(idxFace).C).Y, _
                                                         .TScreen(.Faces(idxFace).B).U, .TScreen(.Faces(idxFace).B).V, _
                                                         .TScreen(.Faces(idxFace).C).U, .TScreen(.Faces(idxFace).C).V, _
                                                         .Faces(idxFace).idxMat)
                                    End If
                                    If .Faces(idxFace).CA Then
                                        Call DrawLineTex(.Screen(.Faces(idxFace).C).X, .Screen(.Faces(idxFace).C).Y, _
                                                         .Screen(.Faces(idxFace).A).X, .Screen(.Faces(idxFace).A).Y, _
                                                         .TScreen(.Faces(idxFace).C).U, .TScreen(.Faces(idxFace).C).V, _
                                                         .TScreen(.Faces(idxFace).A).U, .TScreen(.Faces(idxFace).A).V, _
                                                         .Faces(idxFace).idxMat)
                                    End If
                                Else
                                    Select Case FStat
                                        Case CanIn
                                            If .Faces(idxFace).AB Then
                                                lColor = ColorRGBToLong(.Vertices(.Faces(idxFace).A).ColorS)
                                                Call DrawLine(.Screen(.Faces(idxFace).A), .Screen(.Faces(idxFace).B), lColor)
                                            End If
                                            If .Faces(idxFace).BC Then
                                                lColor = ColorRGBToLong(.Vertices(.Faces(idxFace).B).ColorS)
                                                Call DrawLine(.Screen(.Faces(idxFace).B), .Screen(.Faces(idxFace).C), lColor)
                                            End If
                                            If .Faces(idxFace).CA Then
                                                lColor = ColorRGBToLong(.Vertices(.Faces(idxFace).C).ColorS)
                                                Call DrawLine(.Screen(.Faces(idxFace).C), .Screen(.Faces(idxFace).A), lColor)
                                            End If
                                        Case CanClip
                                            lColor = ColorRGBToLong(.Vertices(.Faces(idxFace).A).ColorS)
                                            DrawClip (lColor)
                                    End Select
                                End If
                            End If
                        End If
                    Next idxFace
                Else
                    For idxFace = 0 To .idxFace
                        If IsInCamera(idxMesh, idxFace) Then
                            FStat = FaceStat(idxMesh, idxFace, , True)
                            If FStat <> CanOut Then
                                If Materials(.Faces(idxFace).idxMat).MapUse Then
                                    Call CalculateUV(idxMesh, idxFace)
                                    Call DrawLineTex(.Screen(.Faces(idxFace).A).X, .Screen(.Faces(idxFace).A).Y, _
                                                     .Screen(.Faces(idxFace).B).X, .Screen(.Faces(idxFace).B).Y, _
                                                     .TScreen(.Faces(idxFace).A).U, .TScreen(.Faces(idxFace).A).V, _
                                                     .TScreen(.Faces(idxFace).B).U, .TScreen(.Faces(idxFace).B).V, _
                                                     .Faces(idxFace).idxMat)
                                    Call DrawLineTex(.Screen(.Faces(idxFace).B).X, .Screen(.Faces(idxFace).B).Y, _
                                                     .Screen(.Faces(idxFace).C).X, .Screen(.Faces(idxFace).C).Y, _
                                                     .TScreen(.Faces(idxFace).B).U, .TScreen(.Faces(idxFace).B).V, _
                                                     .TScreen(.Faces(idxFace).C).U, .TScreen(.Faces(idxFace).C).V, _
                                                     .Faces(idxFace).idxMat)
                                    Call DrawLineTex(.Screen(.Faces(idxFace).C).X, .Screen(.Faces(idxFace).C).Y, _
                                                     .Screen(.Faces(idxFace).A).X, .Screen(.Faces(idxFace).A).Y, _
                                                     .TScreen(.Faces(idxFace).C).U, .TScreen(.Faces(idxFace).C).V, _
                                                     .TScreen(.Faces(idxFace).A).U, .TScreen(.Faces(idxFace).A).V, _
                                                     .Faces(idxFace).idxMat)
                                Else
                                    Select Case FStat
                                        Case CanIn
                                            lColor = ColorRGBToLong(.Vertices(.Faces(idxFace).A).ColorS)
                                            Call DrawLine(.Screen(.Faces(idxFace).A), .Screen(.Faces(idxFace).B), lColor)
                                            lColor = ColorRGBToLong(.Vertices(.Faces(idxFace).B).ColorS)
                                            Call DrawLine(.Screen(.Faces(idxFace).B), .Screen(.Faces(idxFace).C), lColor)
                                            lColor = ColorRGBToLong(.Vertices(.Faces(idxFace).C).ColorS)
                                            Call DrawLine(.Screen(.Faces(idxFace).C), .Screen(.Faces(idxFace).A), lColor)
                                        Case CanClip
                                            lColor = ColorRGBToLong(.Vertices(.Faces(idxFace).A).ColorS)
                                            DrawClip (lColor)
                                    End Select
                                End If
                            End If
                        End If
                    Next idxFace
                End If
            Else
                If .FaceInfo Then
                    For idxFace = 0 To .idxFace
                        If IsInCamera(idxMesh, idxFace) Then
                            FStat = FaceStat(idxMesh, idxFace, , True)
                            Select Case FStat
                                Case CanIn
                                    If .Faces(idxFace).AB Then
                                        lColor = ColorRGBToLong(.Vertices(.Faces(idxFace).A).ColorS)
                                        Call DrawLine(.Screen(.Faces(idxFace).A), .Screen(.Faces(idxFace).B), lColor)
                                    End If
                                    If .Faces(idxFace).BC Then
                                        lColor = ColorRGBToLong(.Vertices(.Faces(idxFace).B).ColorS)
                                        Call DrawLine(.Screen(.Faces(idxFace).B), .Screen(.Faces(idxFace).C), lColor)
                                    End If
                                    If .Faces(idxFace).CA Then
                                        lColor = ColorRGBToLong(.Vertices(.Faces(idxFace).C).ColorS)
                                        Call DrawLine(.Screen(.Faces(idxFace).C), .Screen(.Faces(idxFace).A), lColor)
                                    End If
                                Case CanClip
                                    lColor = ColorRGBToLong(.Vertices(.Faces(idxFace).A).ColorS)
                                    DrawClip (lColor)
                            End Select
                        End If
                    Next idxFace
                Else
                    For idxFace = 0 To .idxFace
                        If IsInCamera(idxMesh, idxFace) Then
                            FStat = FaceStat(idxMesh, idxFace, , True)
                            Select Case FStat
                                Case CanIn
                                    lColor = ColorRGBToLong(.Vertices(.Faces(idxFace).A).ColorS)
                                    Call DrawLine(.Screen(.Faces(idxFace).A), .Screen(.Faces(idxFace).B), lColor)
                                    lColor = ColorRGBToLong(.Vertices(.Faces(idxFace).B).ColorS)
                                    Call DrawLine(.Screen(.Faces(idxFace).B), .Screen(.Faces(idxFace).C), lColor)
                                    lColor = ColorRGBToLong(.Vertices(.Faces(idxFace).C).ColorS)
                                    Call DrawLine(.Screen(.Faces(idxFace).C), .Screen(.Faces(idxFace).A), lColor)
                                 Case CanClip
                                    lColor = ColorRGBToLong(.Vertices(.Faces(idxFace).A).ColorS)
                                    DrawClip (lColor)
                            End Select
                        End If
                    Next idxFace
                End If
            End If
        End With
    Next idxMesh
    
End Sub
'ok
Private Sub RenderFlat(Lit As Boolean)
    
    Dim idxOrder    As Long
    Dim idxMeshO    As Integer
    Dim idxFaceO    As Long
    Dim FStat       As CanStat
    Dim Alpha       As Single
        
    Call GetFaceColorO(Lit)
    For idxOrder = 0 To UBound(MeshsOrder)
        idxMeshO = MeshsOrder(idxOrder).idxMeshO
        idxFaceO = MeshsOrder(idxOrder).idxFaceO
        If IsInCamera(idxMeshO, idxFaceO) Then
            FStat = FaceStat(idxMeshO, idxFaceO)
            Alpha = Materials(Meshs(idxMeshO).Faces(idxFaceO).idxMat).Transparency
            If Alpha = 0 Then '0=opaque  0.01~1 = transparent
                Select Case FStat
                    Case CanIn:     Call DrawTriangleFlat(idxMeshO, idxFaceO, Lit)
                    Case CanClip:   Call DrawPolygonFlat(idxMeshO, idxFaceO, Lit)
                End Select
            Else
                If FStat <> CanOut Then Call DrawTriangleFlatA(idxMeshO, idxFaceO, Alpha, Lit)
            End If
        End If
    Next idxOrder
    
End Sub
'ok
Private Sub RenderGouraud()

    Dim idxOrder    As Long
    Dim idxMeshO    As Integer
    Dim idxFaceO    As Long
    Dim FStat       As CanStat
    Dim Alpha       As Single
    
    Call GetFaceColorO
    For idxOrder = 0 To UBound(MeshsOrder)
        idxMeshO = MeshsOrder(idxOrder).idxMeshO
        idxFaceO = MeshsOrder(idxOrder).idxFaceO
        If IsInCamera(idxMeshO, idxFaceO) Then
            FStat = FaceStat(idxMeshO, idxFaceO)
            Alpha = Materials(Meshs(idxMeshO).Faces(idxFaceO).idxMat).Transparency
            If Alpha = 0 Then '0=opaque  0.01~1 = transparent
                Select Case FStat
                    Case CanIn: Call DrawTriangleGradient(idxMeshO, idxFaceO)
                    Case CanClip: Call DrawPolygonGradient(idxMeshO, idxFaceO)
                End Select
            Else
                If FStat <> CanOut Then Call DrawTriangleGradientA(idxMeshO, idxFaceO, Alpha)
            End If
        End If
    Next idxOrder

End Sub

'ok
Private Sub RenderTexture(Lit As Boolean)

    Dim idxOrder    As Long
    Dim idxMeshO    As Integer
    Dim idxFaceO    As Long
    Dim rgbColor    As COLORRGB
    Dim rgbColorS   As COLORRGB
    Dim Gray        As Integer
    Dim FStat       As CanStat
    Dim Alpha       As Single

    Call GetFaceColorO(Lit)
    For idxOrder = 0 To UBound(MeshsOrder)
        idxMeshO = MeshsOrder(idxOrder).idxMeshO
        idxFaceO = MeshsOrder(idxOrder).idxFaceO
        If IsInCamera(idxMeshO, idxFaceO) Then
            FStat = FaceStat(idxMeshO, idxFaceO)
            Alpha = Materials(Meshs(idxMeshO).Faces(idxFaceO).idxMat).Transparency
            If Meshs(idxMeshO).MapCoorsOK And Materials(Meshs(idxMeshO).Faces(idxFaceO).idxMat).MapUse Then
                If FStat <> CanOut Then
                    With Meshs(idxMeshO)
                        rgbColor = Materials(Meshs(idxMeshO).Faces(idxFaceO).idxMat).DiffuseColor
                        rgbColorS = ColorAverage(.Vertices(.Faces(idxFaceO).A).ColorS, _
                                                 .Vertices(.Faces(idxFaceO).B).ColorS, _
                                                 .Vertices(.Faces(idxFaceO).C).ColorS)
                    End With
                    Gray = CSng(rgbColorS.R - rgbColor.R)
                    Call DrawTriangleTex(idxMeshO, idxFaceO, Alpha, Lit, Gray)
                End If
            Else
                If (Alpha = 0) Then '0=opaque  0.01~1 = transparent
                    Select Case FStat
                        Case CanIn:     Call DrawTriangleFlat(idxMeshO, idxFaceO, Lit)
                        Case CanClip:   Call DrawPolygonFlat(idxMeshO, idxFaceO, Lit)
                    End Select
                Else
                    If FStat <> CanOut Then Call DrawTriangleFlatA(idxMeshO, idxFaceO, Alpha, Lit)
                End If
            End If
        End If
    Next
    
End Sub
'ok
Private Sub RenderPhoto()

    Dim idxOrder    As Long
    Dim idxMeshO    As Integer
    Dim idxFaceO    As Long
    Dim rgbColor    As COLORRGB
    Dim rgbColorS   As COLORRGB
    Dim Gray        As Integer
    Dim FStat       As CanStat
    Dim Alpha       As Single

    Call GetFaceColorO(True)
    For idxOrder = 0 To UBound(MeshsOrder)
        idxMeshO = MeshsOrder(idxOrder).idxMeshO
        idxFaceO = MeshsOrder(idxOrder).idxFaceO
        If IsInCamera(idxMeshO, idxFaceO) Then
            FStat = FaceStat(idxMeshO, idxFaceO)
            Alpha = Materials(Meshs(idxMeshO).Faces(idxFaceO).idxMat).Transparency
            If Meshs(idxMeshO).MapCoorsOK Then
                If FStat <> CanOut Then
                    With Meshs(idxMeshO)
                        rgbColor = Materials(Meshs(idxMeshO).Faces(idxFaceO).idxMat).DiffuseColor
                        rgbColorS = ColorAverage(.Vertices(.Faces(idxFaceO).A).ColorS, _
                                                 .Vertices(.Faces(idxFaceO).B).ColorS, _
                                                 .Vertices(.Faces(idxFaceO).C).ColorS)
                    End With
                    Gray = (rgbColor.R + rgbColor.G + rgbColor.B) * sng1Div3 'CSng(rgbColorS.R - rgbColor.R) '* 0.5
                    Call DrawTrianglePhoto(idxMeshO, idxFaceO, Alpha, Gray)
                End If
            
            Else
                If (Alpha = 0) Then '0=opaque  0.01~1 = transparent
                    Select Case FStat
                        Case CanIn:     Call DrawTriangleGradient(idxMeshO, idxFaceO)
                        Case CanClip:   Call DrawPolygonGradient(idxMeshO, idxFaceO)
                    End Select
                Else
                    If FStat <> CanOut Then Call DrawTriangleGradientA(idxMeshO, idxFaceO, Alpha)
                End If
            End If
        End If
    Next
    
End Sub


'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Draws~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'ok
Private Sub DrawDot(P1 As POINTAPI, lColor As Long)

    dibCanvas.mapArray(P1.X, P1.Y) = lColor
    dibCanvas.mapArray(P1.X + 1, P1.Y) = lColor
    dibCanvas.mapArray(P1.X + 1, P1.Y + 1) = lColor
    dibCanvas.mapArray(P1.X, P1.Y + 1) = lColor

End Sub
'ok
Private Sub DrawLine(P1 As POINTAPI, P2 As POINTAPI, lColor As Long)
    
    Dim Pen     As Long
    Dim pAPI    As POINTAPI

    Pen = SelectObject(dibCanvas.mapDIB.hDC, CreatePen(0, 1, lColor))
    MoveToEx dibCanvas.mapDIB.hDC, P1.X, P1.Y, pAPI
    LineTo dibCanvas.mapDIB.hDC, P2.X, P2.Y
    Call DeleteObject(Pen)

End Sub
'ok
Private Sub DrawFrame()

    Dim P1  As POINTAPI
    Dim P2  As POINTAPI
    Dim Col As Long
    
    Col = RGB(200, 200, 250)
    
    P1.X = CanRect.L:   P1.Y = CanRect.T
    P2.X = CanRect.R:   P2.Y = CanRect.T
    Call DrawLine(P1, P2, Col)
    
    P1.X = CanRect.R:   P1.Y = CanRect.T
    P2.X = CanRect.R:   P2.Y = CanRect.D
    Call DrawLine(P1, P2, Col)
    
    P1.X = CanRect.R:   P1.Y = CanRect.D
    P2.X = CanRect.L:   P2.Y = CanRect.D
    Call DrawLine(P1, P2, Col)
    
    P1.X = CanRect.L:   P1.Y = CanRect.D
    P2.X = CanRect.L:   P2.Y = CanRect.T
    Call DrawLine(P1, P2, Col)

End Sub
'ok
Private Sub DrawClip(lColor As Long)
    
    Dim idx As Long

    For idx = 1 To (UBound(Poly) - 1) Step 2
        Call DrawLine(Poly(idx), Poly(idx + 1), lColor)
    Next
    
End Sub
'ok
Private Sub DrawTriangleFlat(idxMesh As Integer, idxFace As Long, Lit As Boolean)

    Dim pAPI(2)     As POINTAPI
    Dim rgbColor    As COLORRGB
    Dim lColor      As Long
    Dim Pen         As Long
    Dim Brush       As Long
    
    With Meshs(idxMesh)
        rgbColor = ColorAverage(.Vertices(.Faces(idxFace).A).ColorS, _
                                .Vertices(.Faces(idxFace).B).ColorS, _
                                .Vertices(.Faces(idxFace).C).ColorS)
        lColor = ColorRGBToLong(rgbColor)
        Pen = SelectObject(dibCanvas.mapDIB.hDC, CreatePen(0, 1, lColor))
        Brush = SelectObject(dibCanvas.mapDIB.hDC, CreateSolidBrush(lColor))
        pAPI(0) = .Screen(.Faces(idxFace).A)
        pAPI(1) = .Screen(.Faces(idxFace).B)
        pAPI(2) = .Screen(.Faces(idxFace).C)
    End With
    Call Polygon(dibCanvas.mapDIB.hDC, pAPI(0), 3)
    Call DeleteObject(Pen)
    Call DeleteObject(Brush)

End Sub
'ok
Private Sub DrawPolygonFlat(idxMesh As Integer, idxFace As Long, Lit As Boolean)

    Dim pAPI(2)     As POINTAPI
    Dim rgbColor    As COLORRGB
    Dim lColor      As Long
    Dim Pen         As Long
    Dim Brush       As Long
    Dim NumFaces    As Long
    Dim Faces(15)   As FACE
    Dim idx         As Long
    
    NumFaces = Triangulate(Faces)
    If NumFaces > 0 Then
        With Meshs(idxMesh)
            rgbColor = ColorAverage(.Vertices(.Faces(idxFace).A).ColorS, _
                                    .Vertices(.Faces(idxFace).B).ColorS, _
                                    .Vertices(.Faces(idxFace).C).ColorS)
            lColor = ColorRGBToLong(rgbColor)
            Pen = SelectObject(dibCanvas.mapDIB.hDC, CreatePen(0, 1, lColor))
            Brush = SelectObject(dibCanvas.mapDIB.hDC, CreateSolidBrush(lColor))
        End With
        For idx = 0 To NumFaces
            pAPI(0) = Poly(Faces(idx).A)
            pAPI(1) = Poly(Faces(idx).B)
            pAPI(2) = Poly(Faces(idx).C)
            Call Polygon(dibCanvas.mapDIB.hDC, pAPI(0), 3)
        Next
        Call DeleteObject(Pen)
        Call DeleteObject(Brush)
    End If

End Sub

Private Sub GetFaceColorO(Optional Lit As Boolean = True)
    
    Dim idxOrder    As Long
    Dim idxMeshO    As Integer
    Dim idxFaceO    As Long
    
'Reset Shaded Colors
    For idxOrder = 0 To UBound(MeshsOrder)
        idxMeshO = MeshsOrder(idxOrder).idxMeshO
        idxFaceO = MeshsOrder(idxOrder).idxFaceO
        With Meshs(idxMeshO)
            .Vertices(.Faces(idxFaceO).A).Used = False
            .Vertices(.Faces(idxFaceO).B).Used = False
            .Vertices(.Faces(idxFaceO).C).Used = False
        End With
    Next idxOrder
        
'Set Shaded Colors
    If Lit Then
        For idxOrder = 0 To UBound(MeshsOrder)
            idxMeshO = MeshsOrder(idxOrder).idxMeshO
            idxFaceO = MeshsOrder(idxOrder).idxFaceO
            Call GetVertexColorS(idxMeshO, idxFaceO)
        Next idxOrder
    Else
        For idxOrder = 0 To UBound(MeshsOrder)
            idxMeshO = MeshsOrder(idxOrder).idxMeshO
            idxFaceO = MeshsOrder(idxOrder).idxFaceO
            Call GetVertexColor(idxMeshO, idxFaceO)
        Next idxOrder
    End If
        
End Sub

Private Sub GetFaceColor(Optional Lit As Boolean = True)
    
    Dim idxMesh     As Integer
    Dim idxFace     As Long
    
    For idxMesh = 0 To g_idxMesh
        With Meshs(idxMesh)
'Reset Shaded Colors
            For idxFace = 0 To .idxFace
                .Vertices(.Faces(idxFace).A).Used = False
                .Vertices(.Faces(idxFace).B).Used = False
                .Vertices(.Faces(idxFace).C).Used = False
            Next idxFace
'Set Shaded Colors
            If Lit Then
                For idxFace = 0 To .idxFace
                    Call GetVertexColorS(idxMesh, idxFace)
                Next idxFace
            Else
                For idxFace = 0 To .idxFace
                    Call GetVertexColor(idxMesh, idxFace)
                Next idxFace
            End If
        End With
    Next idxMesh
    
End Sub

Private Sub GetVertexColorS(idxMesh As Integer, idxFace As Long)
    
    Dim vNormal As VECTOR
    
    With Meshs(idxMesh)
        vNormal = FaceNormal(.Vertices(.Faces(idxFace).A).VectorsS, _
                             .Vertices(.Faces(idxFace).B).VectorsS, _
                             .Vertices(.Faces(idxFace).C).VectorsS)
        If .Vertices(.Faces(idxFace).A).Used = False Then
            .Vertices(.Faces(idxFace).A).ColorS = Shade(idxMesh, idxFace, vNormal, .Faces(idxFace).A)
            .Vertices(.Faces(idxFace).A).Used = True
        End If
        If .Vertices(.Faces(idxFace).B).Used = False Then
            .Vertices(.Faces(idxFace).B).ColorS = Shade(idxMesh, idxFace, vNormal, .Faces(idxFace).B)
            .Vertices(.Faces(idxFace).B).Used = True
        End If
        If .Vertices(.Faces(idxFace).C).Used = False Then
            .Vertices(.Faces(idxFace).C).ColorS = Shade(idxMesh, idxFace, vNormal, .Faces(idxFace).C)
            .Vertices(.Faces(idxFace).C).Used = True
        End If
    End With

End Sub

Private Sub GetVertexColor(idxMesh As Integer, idxFace As Long)
        
    With Meshs(idxMesh)
        If .Vertices(.Faces(idxFace).A).Used = False Then
            .Vertices(.Faces(idxFace).A).ColorS = Materials(.Faces(idxFace).idxMat).DiffuseColor
            .Vertices(.Faces(idxFace).A).Used = True
        End If
        If .Vertices(.Faces(idxFace).B).Used = False Then
            .Vertices(.Faces(idxFace).B).ColorS = Materials(.Faces(idxFace).idxMat).DiffuseColor
            .Vertices(.Faces(idxFace).B).Used = True
        End If
        If .Vertices(.Faces(idxFace).C).Used = False Then
            .Vertices(.Faces(idxFace).C).ColorS = Materials(.Faces(idxFace).idxMat).DiffuseColor
            .Vertices(.Faces(idxFace).C).Used = True
        End If
    End With

End Sub

Private Function PointInCanvas(Point As POINTAPI) As Boolean
    
    If Point.X > CanRect.L And Point.X < CanRect.R And _
       Point.Y > CanRect.T And Point.Y < CanRect.D Then PointInCanvas = True

End Function

Private Function ZValInCamera(ZVal As Single) As Boolean
    
    If (ZVal > Cameras(0).ClipNear) And _
       (ZVal < Cameras(0).ClipFar) Then ZValInCamera = True
                   
End Function

Private Function IsInCamera(idxMesh As Integer, idxFace As Long) As Boolean
    
    With Meshs(idxMesh)
        If (.Vertices(.Faces(idxFace).A).VectorsT.Z > Cameras(0).ClipNear) And _
           (.Vertices(.Faces(idxFace).B).VectorsT.Z > Cameras(0).ClipNear) And _
           (.Vertices(.Faces(idxFace).C).VectorsT.Z > Cameras(0).ClipNear) And _
           (.Vertices(.Faces(idxFace).A).VectorsT.Z < Cameras(0).ClipFar) And _
           (.Vertices(.Faces(idxFace).B).VectorsT.Z < Cameras(0).ClipFar) And _
           (.Vertices(.Faces(idxFace).C).VectorsT.Z < Cameras(0).ClipFar) Then IsInCamera = True
    End With

End Function
