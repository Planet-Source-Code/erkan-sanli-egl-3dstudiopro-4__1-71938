Attribute VB_Name = "modFile"
Option Explicit

Private Type ITEMID
    cb      As Long
    abID    As Integer
End Type

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMID) As Long

Public Function GetMyPicturesFolder(hWnd As Long) As String
    
    Dim RetVal As Long
    Dim tiID As ITEMID
    Dim Folder As String
    
    Folder = Space$(260)
    RetVal = SHGetSpecialFolderLocation(hWnd, &H27, tiID) 'MyPictures= &H27
    RetVal = SHGetPathFromIDList(ByVal tiID.cb, ByVal Folder)
    If RetVal Then GetMyPicturesFolder = Left$(Folder, InStr(1, Folder, Chr$(0)) - 1) & "\"

End Function

Public Function FileExist(FileName As String) As Boolean
    
    On Error Resume Next
    FileExist = CBool(FileLen(FileName))
    
End Function

Public Sub LoadTexture(idx As Integer, FileName As String)
   
    With frmMain
        On Error GoTo Jump
        If FileExist(FileName) Then
            Materials(idx).MapUse = True
            .picLoad.Picture = LoadPicture(FileName)
            Set Materials(idx).dibTex.mapDIB = New clsDIB
            Set Materials(idx).dibTexT.mapDIB = New clsDIB
            Call Materials(idx).dibTex.mapDIB.CreateFromPictureBox(.picLoad, .picLoad.ScaleWidth, .picLoad.ScaleHeight, Materials(idx).dibTex.mapArray, True)
            Call Materials(idx).dibTexT.mapDIB.CreateFromPictureBox(.picLoad, .picLoad.ScaleWidth, .picLoad.ScaleHeight, Materials(idx).dibTexT.mapArray, True)
            Exit Sub
        End If
    End With
    
Jump:
     Materials(idx).MapUse = False

End Sub

Public Sub LoadBackPicture(FileName As String)
    
    With frmMain
        On Error GoTo Jump
        If FileExist(FileName) Then
            .picLoad.Picture = LoadPicture(FileName)
            Call dibBack.mapDIB.Delete(dibBack.mapArray)
            Call dibBack.mapDIB.CreateFromPictureBox(.picLoad, CanRect.R, CanRect.D, dibBack.mapArray)
            BType = BackPicture
            Exit Sub
        End If
    End With

Jump:
    BType = Blank
End Sub

Public Function PicturePath(InitDir As String, Title As String) As String
    
    Set cdiLoad = New clsCommonDialog
    With cdiLoad
        .Filter = "All Picture Files|*.bmp;*.dib;*.gif;*.jpg;*.wmf;*.emf;*.ico;*.cur |" & _
                  "Bitmaps (*.bmp;*.dib)|*.bmp;*.dib|" & _
                  "GIF Images (*.gif)|*.gif|" & _
                  "JPEG Images (*.jpg)|*.jpg,*.jpeg|" & _
                  "All Files (*.*)|*.*"
        .DialogTitle = Title
        .InitDir = InitDir
        .FileName = ""
        .ShowOpen
        PicturePath = .FileName
    End With

End Function

Public Function GetFilePath(strFilePath As String) As String

    Dim FilenameEx  As String
    Dim Length      As Long
    
    FilenameEx = GetFileNameEx(strFilePath)
    Length = Len(strFilePath) - Len(FilenameEx)
    GetFilePath = Left(strFilePath, Length)
    
End Function

Public Function GetFileNameEx(strFilePath As String) As String

    Dim Segments() As String
    
    If Len(strFilePath) <> 0 Then
        Segments = Split(strFilePath, "\")
        GetFileNameEx = Segments(UBound(Segments))
    Else
        GetFileNameEx = "-"
    End If
    
End Function

