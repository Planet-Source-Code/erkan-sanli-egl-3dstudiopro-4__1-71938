VERSION 5.00
Begin VB.Form frmMaterial 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Material Manager"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   544
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   514
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   6480
      TabIndex        =   36
      Top             =   7560
      Width           =   975
   End
   Begin VB.Frame Frame5 
      Caption         =   "Color"
      Height          =   975
      Left            =   3360
      TabIndex        =   28
      Top             =   5280
      Width           =   4215
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   360
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   48
         TabIndex        =   30
         Top             =   360
         Width           =   720
      End
      Begin VB.CommandButton cmdColorMesh 
         Caption         =   "Color"
         Height          =   420
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Transparent"
      Height          =   975
      Left            =   3360
      TabIndex        =   15
      Top             =   6360
      Width           =   4215
      Begin VB.CommandButton cmdTransApply 
         Caption         =   "Apply"
         Height          =   375
         Left            =   3240
         TabIndex        =   34
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtTransparency 
         Height          =   300
         Left            =   360
         TabIndex        =   33
         Text            =   "Transparency"
         Top             =   360
         Width           =   600
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Mirror - Flip"
      Height          =   3375
      Left            =   3360
      TabIndex        =   14
      Top             =   120
      Width           =   4215
      Begin VB.CommandButton cmdMeshTex 
         Caption         =   "Browse"
         Height          =   375
         Left            =   3240
         TabIndex        =   32
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton optFlip 
         Caption         =   "Vertical"
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   27
         Top             =   2040
         Width           =   1095
      End
      Begin VB.OptionButton optFlip 
         Caption         =   "Horizontal"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   26
         Top             =   1800
         Width           =   1095
      End
      Begin VB.OptionButton optFlip 
         Caption         =   "None"
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   25
         Top             =   1560
         Width           =   1095
      End
      Begin VB.PictureBox picPreview 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1800
         Left            =   240
         ScaleHeight     =   120
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   1800
      End
      Begin VB.CommandButton cmdMirrorApply 
         Caption         =   "Apply"
         Height          =   375
         Left            =   3360
         TabIndex        =   21
         Top             =   2760
         Width           =   615
      End
      Begin VB.CheckBox chkMirror 
         Height          =   200
         Index           =   0
         Left            =   2520
         TabIndex        =   17
         Top             =   1560
         Width           =   200
      End
      Begin VB.CheckBox chkMirror 
         Height          =   200
         Index           =   1
         Left            =   2520
         TabIndex        =   16
         Top             =   1800
         Width           =   200
      End
      Begin VB.Label lblLoadTexture 
         Caption         =   "-"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label lblName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "-"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "Flip"
         Height          =   180
         Left            =   3120
         TabIndex        =   24
         Top             =   1320
         Width           =   600
      End
      Begin VB.Label Label9 
         Caption         =   "U"
         Height          =   255
         Left            =   2280
         TabIndex        =   20
         Top             =   1560
         Width           =   135
      End
      Begin VB.Label Label7 
         Caption         =   "V"
         Height          =   255
         Left            =   2280
         TabIndex        =   19
         Top             =   1800
         Width           =   135
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Mirror"
         Height          =   300
         Left            =   2280
         TabIndex        =   18
         Top             =   1320
         Width           =   600
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Coordinates"
      Height          =   1575
      Left            =   3360
      TabIndex        =   2
      Top             =   3600
      Width           =   4215
      Begin VB.CommandButton cmdCoordApply 
         Caption         =   "Apply"
         Height          =   375
         Left            =   3240
         TabIndex        =   22
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtRotate 
         Height          =   300
         Left            =   2760
         TabIndex        =   11
         Text            =   "WRotate"
         Top             =   600
         Width           =   600
      End
      Begin VB.TextBox txtVTiling 
         Height          =   300
         Left            =   1440
         TabIndex        =   6
         Text            =   "VTiling"
         Top             =   960
         Width           =   600
      End
      Begin VB.TextBox txtUTiling 
         Height          =   300
         Left            =   1440
         TabIndex        =   5
         Text            =   "UTiling"
         Top             =   600
         Width           =   600
      End
      Begin VB.TextBox txtVOffset 
         Height          =   300
         Left            =   480
         TabIndex        =   4
         Text            =   "VOffset"
         Top             =   960
         Width           =   600
      End
      Begin VB.TextBox txtUOffset 
         Height          =   300
         Left            =   480
         TabIndex        =   3
         Text            =   "UOffset"
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Label8 
         Caption         =   "W"
         Height          =   255
         Left            =   2520
         TabIndex        =   13
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Rotate"
         Height          =   300
         Left            =   2760
         TabIndex        =   12
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label4 
         Caption         =   "V"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "U"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   135
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Tiling"
         Height          =   300
         Left            =   1440
         TabIndex        =   8
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Offset"
         Height          =   300
         Left            =   480
         TabIndex        =   7
         Top             =   360
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Material List"
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstMaterials 
         Height          =   6690
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmColorMesh     As frmColorDialog
Attribute frmColorMesh.VB_VarHelpID = -1

Dim dibBuffer   As DIB
Dim idxSelMat   As Integer

Private Sub cmdClose_Click()
    
    Unload Me

End Sub

Private Sub Form_Load()

    Dim idx As Integer
        
    For idx = 0 To g_idxMat
        lstMaterials.AddItem CStr(idx) & ": " & Materials(idx).MatName
    Next
    Set dibBuffer.mapDIB = New clsDIB
    
End Sub

Private Sub lstMaterials_Click()
    
    idxSelMat = lstMaterials.ListIndex
    
    If Materials(idxSelMat).MapUse Then
        Call dibBuffer.mapDIB.Delete(dibBuffer.mapArray)
        Call dibBuffer.mapDIB.Create(Materials(idxSelMat).dibTex.mapDIB.Width * 2, Materials(idxSelMat).dibTex.mapDIB.Height * 2, dibBuffer.mapArray)
        lblLoadTexture.Caption = "Loaded"
    Else
        lblLoadTexture.Caption = "File Not Found"
        picPreview.Cls
    End If
    
    txtUOffset.Text = Materials(idxSelMat).UOffset
    txtVOffset.Text = Materials(idxSelMat).VOffset
    txtUTiling.Text = Materials(idxSelMat).UTiling
    txtVTiling.Text = Materials(idxSelMat).VTiling
    txtRotate.Text = Materials(idxSelMat).Angle
    lblName.Caption = Materials(idxSelMat).MapFileName
    picColor.BackColor = ColorRGBToLong(Materials(idxSelMat).DiffuseColor)
    txtTransparency = Materials(idxSelMat).Transparency
    Call RefreshPreview

End Sub

Private Sub chkMirror_Click(Index As Integer)
    
    Dim Ustat As Byte
    Dim Vstat As Byte
    
    If Materials(idxSelMat).MapUse = False Then Exit Sub

    Ustat = IIf(chkMirror(0).Value, 1, 0)
    Vstat = IIf(chkMirror(1).Value, 10, 0)
    Materials(idxSelMat).Mirroring = Ustat + Vstat
    Call RefreshPreview

End Sub

Private Sub optFlip_Click(Index As Integer)
    
    If Materials(idxSelMat).MapUse = False Then Exit Sub

    Materials(idxSelMat).Flipping = CByte(Index)
    Call RefreshPreview

End Sub

Private Sub RefreshPreview()
    
    If Materials(idxSelMat).MapUse = False Then Exit Sub
    
    Call SetStretchBltMode(picPreview.hDC, vbPaletteModeNone)
    
    Select Case Materials(idxSelMat).Mirroring
        
        Case 0
            Call StretchBlt(dibBuffer.mapDIB.hDC, 0, 0, dibBuffer.mapDIB.Width, dibBuffer.mapDIB.Height, _
                            Materials(idxSelMat).dibTex.mapDIB.hDC, 0, 0, Materials(idxSelMat).dibTex.mapDIB.Width, Materials(idxSelMat).dibTex.mapDIB.Height, vbSrcCopy)
        
        Case 1
            Call StretchBlt(dibBuffer.mapDIB.hDC, 0, 0, Materials(idxSelMat).dibTex.mapDIB.Width, dibBuffer.mapDIB.Height, _
                            Materials(idxSelMat).dibTex.mapDIB.hDC, 0, 0, Materials(idxSelMat).dibTex.mapDIB.Width, Materials(idxSelMat).dibTex.mapDIB.Height, vbSrcCopy)
            Call StretchBlt(dibBuffer.mapDIB.hDC, Materials(idxSelMat).dibTex.mapDIB.Width, 0, Materials(idxSelMat).dibTex.mapDIB.Width, dibBuffer.mapDIB.Height, _
                            Materials(idxSelMat).dibTex.mapDIB.hDC, Materials(idxSelMat).dibTex.mapDIB.Width, 0, -Materials(idxSelMat).dibTex.mapDIB.Width, Materials(idxSelMat).dibTex.mapDIB.Height, vbSrcCopy)
        
        Case 10
            Call StretchBlt(dibBuffer.mapDIB.hDC, 0, 0, dibBuffer.mapDIB.Width, Materials(idxSelMat).dibTex.mapDIB.Height, _
                            Materials(idxSelMat).dibTex.mapDIB.hDC, 0, 0, Materials(idxSelMat).dibTex.mapDIB.Width, Materials(idxSelMat).dibTex.mapDIB.Height, vbSrcCopy)
            Call StretchBlt(dibBuffer.mapDIB.hDC, 0, Materials(idxSelMat).dibTex.mapDIB.Height, dibBuffer.mapDIB.Width, Materials(idxSelMat).dibTex.mapDIB.Height, _
                            Materials(idxSelMat).dibTex.mapDIB.hDC, 0, Materials(idxSelMat).dibTex.mapDIB.Height, Materials(idxSelMat).dibTex.mapDIB.Width, -Materials(idxSelMat).dibTex.mapDIB.Height, vbSrcCopy)
        Case 11
            Call StretchBlt(dibBuffer.mapDIB.hDC, 0, 0, Materials(idxSelMat).dibTex.mapDIB.Width, Materials(idxSelMat).dibTex.mapDIB.Height, _
                            Materials(idxSelMat).dibTex.mapDIB.hDC, 0, 0, Materials(idxSelMat).dibTex.mapDIB.Width, Materials(idxSelMat).dibTex.mapDIB.Height, vbSrcCopy)
            Call StretchBlt(dibBuffer.mapDIB.hDC, Materials(idxSelMat).dibTex.mapDIB.Width, 0, Materials(idxSelMat).dibTex.mapDIB.Width, Materials(idxSelMat).dibTex.mapDIB.Height, _
                            Materials(idxSelMat).dibTex.mapDIB.hDC, Materials(idxSelMat).dibTex.mapDIB.Width, 0, -Materials(idxSelMat).dibTex.mapDIB.Width, Materials(idxSelMat).dibTex.mapDIB.Height, vbSrcCopy)
            Call StretchBlt(dibBuffer.mapDIB.hDC, 0, Materials(idxSelMat).dibTex.mapDIB.Height, Materials(idxSelMat).dibTex.mapDIB.Width, Materials(idxSelMat).dibTex.mapDIB.Height, _
                            Materials(idxSelMat).dibTex.mapDIB.hDC, 0, Materials(idxSelMat).dibTex.mapDIB.Height, Materials(idxSelMat).dibTex.mapDIB.Width, -Materials(idxSelMat).dibTex.mapDIB.Height, vbSrcCopy)
            Call StretchBlt(dibBuffer.mapDIB.hDC, Materials(idxSelMat).dibTex.mapDIB.Width, Materials(idxSelMat).dibTex.mapDIB.Height, Materials(idxSelMat).dibTex.mapDIB.Width, Materials(idxSelMat).dibTex.mapDIB.Height, _
                            Materials(idxSelMat).dibTex.mapDIB.hDC, Materials(idxSelMat).dibTex.mapDIB.Width, Materials(idxSelMat).dibTex.mapDIB.Height, -Materials(idxSelMat).dibTex.mapDIB.Width, -Materials(idxSelMat).dibTex.mapDIB.Height, vbSrcCopy)
    End Select
    
    Select Case Materials(idxSelMat).Flipping
        Case 0
            Call StretchBlt(picPreview.hDC, 0, 0, picPreview.ScaleWidth, picPreview.ScaleHeight, _
                            dibBuffer.mapDIB.hDC, 0, 0, dibBuffer.mapDIB.Width, dibBuffer.mapDIB.Height, vbSrcCopy)
        Case 1
            Call StretchBlt(picPreview.hDC, 0, 0, picPreview.ScaleWidth, picPreview.ScaleHeight, _
                            dibBuffer.mapDIB.hDC, 0, dibBuffer.mapDIB.Height, dibBuffer.mapDIB.Width, -dibBuffer.mapDIB.Height, vbSrcCopy)
        Case 2
            Call StretchBlt(picPreview.hDC, 0, 0, picPreview.ScaleWidth, picPreview.ScaleHeight, _
                            dibBuffer.mapDIB.hDC, dibBuffer.mapDIB.Width, 0, -dibBuffer.mapDIB.Width, dibBuffer.mapDIB.Height, vbSrcCopy)
    End Select
    
End Sub

Private Sub cmdCoordApply_Click()
    
    If Materials(idxSelMat).MapUse = False Then Exit Sub

    Materials(idxSelMat).UTiling = VerifyText(txtUTiling)
    Materials(idxSelMat).VTiling = VerifyText(txtVTiling)
    Materials(idxSelMat).UOffset = VerifyText(txtUOffset)
    Materials(idxSelMat).VOffset = VerifyText(txtVOffset)
    Materials(idxSelMat).Angle = VerifyText(txtRotate)

End Sub

Private Sub cmdMirrorApply_Click()

    If Materials(idxSelMat).MapUse = False Then Exit Sub
    
    Select Case Materials(idxSelMat).Flipping
        Case 0
            Call StretchBlt(Materials(idxSelMat).dibTexT.mapDIB.hDC, 0, 0, Materials(idxSelMat).dibTexT.mapDIB.Width, Materials(idxSelMat).dibTexT.mapDIB.Height, _
                            dibBuffer.mapDIB.hDC, 0, 0, dibBuffer.mapDIB.Width, dibBuffer.mapDIB.Height, vbSrcCopy)
        Case 1
            Call StretchBlt(Materials(idxSelMat).dibTexT.mapDIB.hDC, 0, 0, Materials(idxSelMat).dibTexT.mapDIB.Width, Materials(idxSelMat).dibTexT.mapDIB.Height, _
                            dibBuffer.mapDIB.hDC, 0, dibBuffer.mapDIB.Height, dibBuffer.mapDIB.Width, -dibBuffer.mapDIB.Height, vbSrcCopy)
        Case 2
            Call StretchBlt(Materials(idxSelMat).dibTexT.mapDIB.hDC, 0, 0, Materials(idxSelMat).dibTexT.mapDIB.Width, Materials(idxSelMat).dibTexT.mapDIB.Height, _
                            dibBuffer.mapDIB.hDC, dibBuffer.mapDIB.Width, 0, -dibBuffer.mapDIB.Width, dibBuffer.mapDIB.Height, vbSrcCopy)
    End Select

End Sub

Private Sub cmdTransApply_Click()
    
    Materials(idxSelMat).Transparency = VerifyText(txtTransparency)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set dibBuffer.mapDIB = Nothing

End Sub

Private Sub txtRotate_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then cmdCoordApply_Click

End Sub

Private Sub txtTransparency_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then cmdTransApply_Click

End Sub

Private Sub txtUOffset_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then cmdCoordApply_Click

End Sub

Private Sub txtUTiling_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then cmdCoordApply_Click

End Sub

Private Sub txtVOffset_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then cmdCoordApply_Click

End Sub

Private Sub txtVTiling_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then cmdCoordApply_Click

End Sub

Private Sub cmdColorMesh_Click()
    
    If LoadComplete = False Then Exit Sub
    
    If frmColorMesh Is Nothing Then
        Set frmColorMesh = New frmColorDialog
        frmColorMesh.Caption = "Select Mesh Color"
    End If
    frmColorMesh.Show vbModal, frmMaterial
   
End Sub

Private Sub frmColorMesh_Active(R As Integer, G As Integer, B As Integer)

    R = Materials(idxSelMat).DiffuseColor.R
    G = Materials(idxSelMat).DiffuseColor.G
    B = Materials(idxSelMat).DiffuseColor.B

End Sub

Private Sub frmColorMesh_Change(R As Integer, G As Integer, B As Integer)

    Materials(idxSelMat).DiffuseColor = ColorSet(R, G, B)
    picColor.BackColor = ColorRGBToLong(Materials(idxSelMat).DiffuseColor)

End Sub

Private Sub cmdMeshTex_Click()

    Dim FileName As String

    FileName = GetMyPicturesFolder(Me.hWnd)
    FileName = PicturePath(FileName, "Texture File")
    Call LoadTexture(idxSelMat, FileName)

End Sub

