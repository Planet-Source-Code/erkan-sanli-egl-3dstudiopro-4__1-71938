VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "3D Studio Pro 4"
   ClientHeight    =   7245
   ClientLeft      =   150
   ClientTop       =   615
   ClientWidth     =   9645
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   483
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   643
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   9840
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   2
      Top             =   2400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer tmrProcess 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9720
      Top             =   1680
   End
   Begin VB.PictureBox picLoad 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   9720
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox picCanvas 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   7200
      Left            =   0
      ScaleHeight     =   478
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   0
      Width           =   9630
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuImport 
         Caption         =   "Import"
         Begin VB.Menu mnuImport3DS 
            Caption         =   "3DS"
         End
         Begin VB.Menu mnuImportOBJ 
            Caption         =   "OBJ"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuExport 
         Caption         =   "Export"
         Begin VB.Menu mnuExportBMP 
            Caption         =   "BMP"
         End
      End
      Begin VB.Menu tire 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuVisualStyleC 
      Caption         =   "&Visual Style"
      Begin VB.Menu mnuVisualStyle 
         Caption         =   "Dot"
         Index           =   0
      End
      Begin VB.Menu mnuVisualStyle 
         Caption         =   "Box"
         Index           =   1
      End
      Begin VB.Menu mnuVisualStyle 
         Caption         =   "Wireframe"
         Index           =   2
      End
      Begin VB.Menu mnuVisualStyle 
         Caption         =   "Wireframe Map"
         Index           =   3
      End
      Begin VB.Menu mnuVisualStyle 
         Caption         =   "Wireframe BFC"
         Index           =   4
      End
      Begin VB.Menu mnuVisualStyle 
         Caption         =   "Flat"
         Index           =   5
      End
      Begin VB.Menu mnuVisualStyle 
         Caption         =   "Gouraud"
         Index           =   6
      End
      Begin VB.Menu mnuVisualStyle 
         Caption         =   "Mapped"
         Index           =   7
      End
      Begin VB.Menu mnuVisualStyle 
         Caption         =   "Photo Realistic"
         Index           =   8
      End
      Begin VB.Menu tire2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLight 
         Caption         =   "Light"
      End
      Begin VB.Menu mnuDoubleSide 
         Caption         =   "Show Backface"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Enabled         =   0   'False
      Begin VB.Menu mnuMesh 
         Caption         =   "Mesh"
      End
      Begin VB.Menu mnuCamera 
         Caption         =   "Camera"
      End
      Begin VB.Menu mnuMaterialEditor 
         Caption         =   "Material"
      End
      Begin VB.Menu mnuBackground 
         Caption         =   "Background"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuControl 
         Caption         =   "Control"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'_____________________________________________________________________________________________________
'Thanks to
'Kaci Lounes : http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=65535&lngWId=1
'shadowmoy   : http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=59365&lngWId=1
'-----------------------------------------------------------------------------------------------------
'
'08.11.2010 Erkan Þanlý (procam)
'_____________________________________________________________________________________________________

'Dim cTimer      As New clsTiming

Private Sub Form_Load()
    
    Set dibCanvas.mapDIB = New clsDIB
    Set dibBack.mapDIB = New clsDIB
    Call MouseInit
    mnuEdit.Enabled = False
    BFileName = App.Path & "\Background\Default.jpg"
    picBack.Picture = LoadPicture(BFileName)
    BColor1 = ColorSet(200, 200, 250)
    BColor1 = ColorSet(200, 250, 250)
    BType = BackPicture
    
End Sub

Private Sub Form_Resize()
    
    DoEvents
    ResizeCanvas

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    tmrProcess.Enabled = False
    dibCanvas.mapDIB.Delete dibCanvas.mapArray
    dibBack.mapDIB.Delete dibBack.mapArray
    Set dibCanvas.mapDIB = Nothing
    Set dibBack.mapDIB = Nothing
    Set cdiLoad = Nothing
    'Set cTimer = Nothing
    Erase Meshs
    Erase Cameras
    End

End Sub

Private Sub mnuBackground_Click()
    
    frmBackground.Show vbModal
    
End Sub

Private Sub mnuCamera_Click()

    frmCamera.Show vbModal, frmMain
    
End Sub

Private Sub mnuControl_Click()
    
    frmControl.Show

End Sub

Private Sub mnuDoubleSide_Click()

    mnuDoubleSide.Checked = Not mnuDoubleSide.Checked
    g_blnDoubleSide = mnuDoubleSide.Checked

End Sub

Private Sub mnuExit_Click()
    
    Unload Me
    
End Sub

Private Sub mnuImport3DS_Click()
    
    Dim idx As Integer
    Dim FileName As String
    
    Set cdiLoad = New clsCommonDialog
    With cdiLoad
        .Filter = "3D Studio File|*.3ds"
        .InitDir = App.Path & "\Sample"
        .FileName = ""
        .ShowOpen
        If Len(.FileName) = 0 Then Exit Sub
        If Not FileExist(.FileName) Then Exit Sub
        Path3DS = GetFilePath(.FileName)
        Set cf3DS = New clsFile3DS
        LoadComplete = cf3DS.Read3DS(.FileName)
        Set cf3DS = Nothing
    End With
    Set cdiLoad = Nothing
    If LoadComplete = True Then
        Call ReLocaleMeshs
        For idx = 0 To g_idxMat
            FileName = Path3DS & Materials(idx).MapFileName
            Call LoadTexture(idx, FileName)
        Next
        For idx = 0 To g_idxMesh
            If Meshs(idx).MapCoorsOK Then
                Call mnuVisualStyle_Click(Mapped)
            Else
                Call mnuVisualStyle_Click(Gouraud)
            End If
        Next
        mnuEdit.Enabled = True
        tmrProcess.Enabled = True
        frmMain.picCanvas.SetFocus
    Else
        mnuEdit.Enabled = False
        MsgBox "Loading Error"
    End If
    
End Sub

'Private Sub mnuImportOBJ_Click()
'
'    Set cdiLoad = New clsCommonDialog
'    With cdiLoad
'        .Filter = "Wavefront Object File |*.obj"
'        .InitDir = App.Path & "\Sample"
'        .DefaultExt = "*.obj"
'        .FileName = ""
'        .ShowOpen
'        If Len(.FileName) = 0 Then Exit Sub
'        If Not FileExist(.FileName) Then Exit Sub
'        Set cfOBJ = New clsFileOBJ
'        Call cfOBJ.ReadOBJ(.FileName)
'        Set cfOBJ = Nothing
'    End With
'    Set cdiLoad = Nothing
'    LoadComplete = True
'    Call mnuVisualStyle_Click(Gouraud)
'    tmrProcess.Enabled = True
'    frmMain.picCanvas.SetFocus
'
'End Sub

Private Sub mnuExportBMP_Click()
    
    If LoadComplete = False Then Exit Sub
    
    Dim cfBMP   As clsFileBMP
    Dim Result  As Long
    Dim strMsg  As String
    Dim FileName As String
    
    tmrProcess.Enabled = False
    Set cdiLoad = New clsCommonDialog
ReOpen:
    With cdiLoad
        .Filter = "24-bit Bitmap |*.bmp"
        .DefaultExt = ".bmp"
        .DialogTitle = "Save Bitmap"
        .InitDir = GetMyPicturesFolder(Me.hWnd)
        .ShowSave
        FileName = Left$(.FileName, InStr(1, .FileName, Chr$(0)) - 1)
        If Len(FileName) <> 0 Then
            If FileExist(FileName) Then
                strMsg = FileName & " already exists." & vbCrLf & "Do you want to replace it."
                Result = MsgBox(strMsg, vbYesNo)
                If Result = vbYes Then
                    Kill FileName
                Else
                    GoTo ReOpen
                End If
            End If
            Set cfBMP = New clsFileBMP
            Call cfBMP.WriteBMP24(FileName)
            Set cfBMP = Nothing
            strMsg = FileName & vbCrLf & " saved."
            MsgBox strMsg
        End If
    End With
    tmrProcess.Enabled = True

End Sub

Private Sub mnuLight_Click()
    
    mnuLight.Checked = Not mnuLight.Checked
    g_blnLight = mnuLight.Checked
    
End Sub

Private Sub mnuMaterialEditor_Click()
    
    frmMaterial.Show vbModal, frmMain

End Sub

Private Sub mnuMesh_Click()
    
    frmMesh.Show

End Sub

Private Sub mnuVisualStyle_Click(Index As Integer)
    
    Dim idx As Integer
    
    For idx = 0 To mnuVisualStyle.UBound
        mnuVisualStyle(idx).Checked = False
    Next
    mnuVisualStyle(Index).Checked = True
    VStyle = Index

End Sub

Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbRightButton Then PopupMenu mnuVisualStyleC

End Sub

Private Sub tmrProcess_Timer()
    
    Dim idx As Integer
    
    If LoadComplete Then
        DoEvents
        'Call cTimer.Reset
        Call RefreshEvents(picCanvas.hWnd)
        For idx = 0 To g_idxMesh
            If Meshs(idx).idxVert > 0 Then
                If VStyle = Box Then
                    Call Screening(Meshs(idx).Box, Meshs(idx).BoxScreen)
                Else
                    Call Screening(Meshs(idx).Vertices, Meshs(idx).Screen)
                End If
            End If
        Next
        Call Render(picCanvas.hDC)
        'Me.Caption = cTimer.Elapsed
    End If
 
End Sub

Private Sub Screening(Vertices() As VERTEX, Screen() As POINTAPI)
    
    Dim idx         As Long
    Dim matOutput   As MATRIX
    Dim ValZoom     As Single
        
    Cameras(0).ViewMatrix = ViewMatrix
    MeshsWorldMatrix = WorldMatrix
    matOutput = MatrixMultiply(MeshsWorldMatrix, Cameras(0).ViewMatrix)
    ValZoom = Cameras(0).Zoom * 1000
    For idx = 0 To UBound(Vertices)
        Vertices(idx).VectorsT = MatrixMultiplyVector(matOutput, Vertices(idx).Vectors)
        Vertices(idx).VectorsS = Vertices(idx).VectorsT
        If Vertices(idx).VectorsT.Z <> 0 Then
            Vertices(idx).VectorsT.X = (Vertices(idx).VectorsT.X / Vertices(idx).VectorsT.Z)
            Vertices(idx).VectorsT.Y = (Vertices(idx).VectorsT.Y / Vertices(idx).VectorsT.Z)
        End If
        Screen(idx).X = Vertices(idx).VectorsT.X * ValZoom + OriginX
        Screen(idx).Y = Vertices(idx).VectorsT.Y * ValZoom + OriginY
    Next idx
    
End Sub

Private Sub ResizeCanvas()
    
    DoEvents
    CanvasWidth = Me.ScaleWidth
    CanvasHeight = Me.ScaleHeight
    OriginX = CLng(CanvasWidth * 0.5)
    OriginY = CLng(CanvasHeight * 0.5)
    picCanvas.Width = CanvasWidth
    picCanvas.Height = CanvasHeight
    Call dibCanvas.mapDIB.Create(CanvasWidth, CanvasHeight, dibCanvas.mapArray)
    Call dibBack.mapDIB.CreateFromPictureBox(picBack, CanvasWidth, CanvasHeight, dibBack.mapArray)
    CanRect.L = picCanvas.ScaleLeft
    CanRect.T = picCanvas.ScaleTop
    CanRect.R = picCanvas.ScaleWidth - 1
    CanRect.D = picCanvas.ScaleHeight - 1
    
End Sub
