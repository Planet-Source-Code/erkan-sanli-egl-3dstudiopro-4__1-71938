VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileOBJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' Wavefront OBJ File parser
'
Dim hFile       As Long
Private Lines() As String

Public Sub ReadOBJ(FileName As String)
    
    Dim strData     As String
    Dim char        As String
    Dim idx         As Long
    Dim idxVert     As Long
    Dim idxFace     As Long
    Dim Segments()  As String
    
    On Error Resume Next
    
'Reset
    ReDim Meshs(0)
    Call ResetMeshParameters
    
    ReDim Cameras(0)
    Call ResetCameraParameters
    
    ReDim Lights(0)
    Call ResetLightParameters

    hFile = FreeFile
    Open FileName For Input As #hFile
        strData = Input(LOF(1) - 1, #hFile)
    Close #hFile
    
    idxVert = 0
    idxFace = 0
    Erase Lines
    Lines = Split(strData, vbLf)
    With Meshs(0)
        For idx = 0 To UBound(Lines)
            If (Len(Lines(idx))) Then
                Segments = Split(Lines(idx), " ")
                char = Trim(Segments(0))
                Select Case char
                    Case "v"
                        ReDim Preserve .Vertices(idxVert)
                        .Vertices(idxVert).Vectors = GetVectorValue(Lines(idx))
                        idxVert = idxVert + 1
                    Case "f":
                        ReDim Preserve .Faces(idxFace)
                        .Faces(idxFace) = GetFaceValue(Lines(idx))
                        idxFace = idxFace + 1
                End Select
            End If
        Next
        ReDim .Screen(.idxVert)
    End With
    Call CreateBox(0)
    
End Sub

Private Function GetVectorValue(Line As String) As VECTOR

    Dim Value As String
    Dim Segments() As String
    
    Segments = Split(Line, " ")
    
'X Value
    Value = Segments(UBound(Segments) - 2)
    GetVectorValue.X = CSng(Replace(Value, ".", ","))
'Y Value
    Value = Segments(UBound(Segments) - 1)
    GetVectorValue.Y = CSng(Replace(Value, ".", ","))
'Z Value
    Value = Segments(UBound(Segments))
    GetVectorValue.Z = CSng(Replace(Value, ".", ","))
'W Value
    GetVectorValue.W = 1

End Function

Private Function GetFaceValue(Line As String) As FACE

    Dim Value As String
    Dim Values() As String
    Dim Segments() As String
    Dim Segments2() As String
    
    
    Segments = Split(Line, " ")
    
'A Value
    Value = Segments(UBound(Segments) - 2)
    Values = Split(Value, "/")
    GetFaceValue.A = CLng(Values(0)) - 1
'B Value
    Value = Segments(UBound(Segments) - 1)
    Values = Split(Value, "/")
    GetFaceValue.B = CLng(Values(0)) - 1
'C Value
    Value = Segments(UBound(Segments))
    Values = Split(Value, "/")
    GetFaceValue.C = CLng(Values(0)) - 1

End Function

'Public Sub WriteOBJ(FileName As String)
'
'    Dim strTemp     As String
'    Dim idx         As Long
'    Dim idxMesh     As Long
'    Dim idxFace     As Long
'    Dim TotalNumFace As Long
'    Dim strX        As String
'    Dim strY        As String
'    Dim strZ        As String
'
'    On Error GoTo Jump
'
'    If FileExist(FileName) Then Kill FileName
'
'    strTemp = "#" & vbCrLf & "# Create : EGL Ege Graphic Library V1.0" & vbCrLf & "#" & vbCrLf
'
'    hFile = FreeFile
'    Open FileName For Binary As hFile
'
''Description
'        Put #hFile, , strTemp
'
''Vertices
'        For idx = 0 To UBound(Meshs(0).Vertices)
'            strX = GetVal(Meshs(0).Vertices(idx).Vectors.X)
'            strY = GetVal(Meshs(0).Vertices(idx).Vectors.Y)
'            strZ = GetVal(Meshs(0).Vertices(idx).Vectors.Z)
'
'            strTemp = "v  " & strX & " " & strY & " " & strZ & vbCrLf
'            Put #hFile, , strTemp
'        Next
'        strTemp = "# " & CStr(UBound(Meshs(0).Vertices)) & " vertices" & vbCrLf & vbCrLf & "g Object" & vbCrLf
'        Put #hFile, , strTemp
'
''Faces
'        'For idxMesh = 1 To UBound(Meshs(0).Faces)
'            For idxFace = 0 To UBound(Meshs(0).Faces) 'Meshs(0).Meshs(idxMesh).NumFaces
'                With Meshs(0).Faces(idxFace)
'                    strTemp = "f " & CStr(.A) & " " & CStr(.B) & " " & CStr(.C) & vbCrLf
'                    Put #hFile, , strTemp
'                    TotalNumFace = TotalNumFace + 1
'                End With
'            Next
'        'Next
'        strTemp = "# " & CStr(TotalNumFace) & " faces" & vbCrLf & vbCrLf & "g " & vbCrLf
'        Put #hFile, , strTemp
'
'    Close #hFile
'    Exit Sub
'Jump:
'    Close #hFile
'End Sub

'Private Function GetVal(Val As Single) As String
'
'    Val = Round(Val, 6)
'    GetVal = CStr(Val)
'    GetVal = Replace(GetVal, ",", ".")
'
'End Function

