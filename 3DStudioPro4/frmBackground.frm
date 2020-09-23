VERSION 5.00
Begin VB.Form frmBackground 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Background"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   3585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fra 
      Caption         =   "Background"
      Height          =   1695
      Index           =   4
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton cmdColorBack 
         Caption         =   "Color 1"
         Height          =   300
         Index           =   1
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton cmdColorBack 
         Caption         =   "Color 2"
         Height          =   300
         Index           =   2
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton cmdBackType 
         Height          =   375
         Index           =   0
         Left            =   240
         Picture         =   "frmBackground.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   720
         Width           =   500
      End
      Begin VB.CommandButton cmdBackType 
         Height          =   375
         Index           =   1
         Left            =   720
         Picture         =   "frmBackground.frx":02D1
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   720
         Width           =   500
      End
      Begin VB.CommandButton cmdBackType 
         Height          =   375
         Index           =   2
         Left            =   1200
         Picture         =   "frmBackground.frx":05A4
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   720
         Width           =   500
      End
      Begin VB.CommandButton cmdBackType 
         Height          =   375
         Index           =   3
         Left            =   1680
         Picture         =   "frmBackground.frx":0D14
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   720
         Width           =   500
      End
      Begin VB.CommandButton cmdBackType 
         Height          =   375
         Index           =   4
         Left            =   2160
         Picture         =   "frmBackground.frx":1762
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   720
         Width           =   500
      End
      Begin VB.CommandButton cmdBackType 
         Height          =   375
         Index           =   5
         Left            =   2640
         Picture         =   "frmBackground.frx":2074
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   720
         Width           =   500
      End
      Begin VB.CommandButton cmdBackPic 
         Caption         =   "Picture"
         Height          =   300
         Left            =   2280
         TabIndex        =   1
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblBackType 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmBackground"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmBColor1    As frmColorDialog
Attribute frmBColor1.VB_VarHelpID = -1
Private WithEvents frmBColor2    As frmColorDialog
Attribute frmBColor2.VB_VarHelpID = -1

Private Sub cmdColorBack_Click(Index As Integer)
    
    If LoadComplete = False Then Exit Sub
    Select Case Index
        Case 1
            If frmBColor1 Is Nothing Then
                Set frmBColor1 = New frmColorDialog
                frmBColor1.Caption = "Select Back Color 1"
            End If
            frmBColor1.Show vbModal
        Case 2
            If frmBColor2 Is Nothing Then
                Set frmBColor2 = New frmColorDialog
                frmBColor2.Caption = "Select Back Color 2"
            End If
            frmBColor2.Show vbModal
    End Select

End Sub

Private Sub Form_Load()

    cmdColorBack(1).BackColor = RGB(BColor1.R, BColor1.G, BColor1.B)
    cmdColorBack(2).BackColor = RGB(BColor2.R, BColor2.G, BColor2.B)

End Sub

Private Sub frmBColor1_Active(R As Integer, G As Integer, B As Integer)
    
    R = BColor1.R
    G = BColor1.G
    B = BColor1.B
    
End Sub

Private Sub frmBColor2_Active(R As Integer, G As Integer, B As Integer)
    
    R = BColor2.R
    G = BColor2.G
    B = BColor2.B
    
End Sub

Private Sub frmBColor1_Change(R As Integer, G As Integer, B As Integer)
    
    BColor1 = ColorSet(R, G, B)
    cmdColorBack(1).BackColor = RGB(R, G, B)
    cmdBackType_Click (BType)
    
End Sub

Private Sub frmBColor2_Change(R As Integer, G As Integer, B As Integer)
    
    BColor2 = ColorSet(R, G, B)
    cmdColorBack(2).BackColor = RGB(R, G, B)
    cmdBackType_Click (BType)

End Sub

Private Sub cmdBackPic_Click()

    Dim FileName As String
    FileName = PicturePath(App.Path & "\Background\", "Background Picture")
    If FileExist(FileName) Then
        BFileName = FileName
    Else
        BFileName = App.Path & "\Background\Default.jpg"
    End If
    Call cmdBackType_Click(BackPicture)

End Sub

Private Sub cmdBackType_Click(Index As Integer)

    Dim Step As Byte
    Dim FileName As String

    BType = Index
    Select Case BType
        'Case Blank
        Case SingleColor:   Call Gradient0
        Case GradType1:     Call Gradient1
        Case GradType2:     Call Gradient2
        Case GradType3:     Call Gradient3
        Case BackPicture:
            If Not FileExist(BFileName) Then
                Call cmdBackPic_Click
            Else
                Call LoadBackPicture(BFileName)
            End If
    End Select

End Sub


