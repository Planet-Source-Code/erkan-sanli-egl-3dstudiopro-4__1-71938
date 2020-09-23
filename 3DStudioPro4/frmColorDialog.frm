VERSION 5.00
Begin VB.Form frmColorDialog 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Color"
   ClientHeight    =   2280
   ClientLeft      =   765
   ClientTop       =   945
   ClientWidth     =   3555
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   152
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   237
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   14
      Top             =   1680
      Width           =   975
   End
   Begin VB.PictureBox picBase 
      BorderStyle     =   0  'None
      Height          =   1575
      Index           =   0
      Left            =   0
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   225
      TabIndex        =   1
      Top             =   0
      Width           =   3375
      Begin VB.PictureBox picLum 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   600
         ScaleHeight     =   12
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   181
         TabIndex        =   13
         Top             =   1320
         Width           =   2715
         Begin VB.Shape shpLum 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            DrawMode        =   6  'Mask Pen Not
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   225
            Left            =   0
            Top             =   0
            Width           =   30
         End
      End
      Begin VB.PictureBox picSat 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   600
         ScaleHeight     =   12
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   181
         TabIndex        =   12
         Top             =   1080
         Width           =   2715
         Begin VB.Shape shpSat 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            DrawMode        =   6  'Mask Pen Not
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   225
            Left            =   0
            Top             =   0
            Width           =   30
         End
      End
      Begin VB.PictureBox picHue 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   600
         ScaleHeight     =   12
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   181
         TabIndex        =   11
         Top             =   840
         Width           =   2715
         Begin VB.Shape shpHue 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            DrawMode        =   6  'Mask Pen Not
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   225
            Left            =   0
            Top             =   0
            Width           =   30
         End
      End
      Begin VB.PictureBox picBlue 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   600
         ScaleHeight     =   12
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   181
         TabIndex        =   10
         Top             =   600
         Width           =   2715
         Begin VB.Shape shpBlue 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            DrawMode        =   6  'Mask Pen Not
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   225
            Left            =   0
            Top             =   0
            Width           =   30
         End
      End
      Begin VB.PictureBox picGreen 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   600
         ScaleHeight     =   12
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   181
         TabIndex        =   9
         Top             =   360
         Width           =   2715
         Begin VB.Shape shpGreen 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            DrawMode        =   6  'Mask Pen Not
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   225
            Left            =   0
            Top             =   0
            Width           =   30
         End
      End
      Begin VB.PictureBox picRed 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   600
         ScaleHeight     =   12
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   181
         TabIndex        =   2
         Top             =   120
         Width           =   2715
         Begin VB.Shape shpRed 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            DrawMode        =   6  'Mask Pen Not
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   225
            Left            =   0
            Top             =   0
            Width           =   30
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Sat."
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   450
      End
      Begin VB.Label Label1 
         Caption         =   "Lum."
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   450
      End
      Begin VB.Label Label1 
         Caption         =   "Hue"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   450
      End
      Begin VB.Label Label1 
         Caption         =   "Green"
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   450
      End
      Begin VB.Label Label1 
         Caption         =   "Blue"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   450
      End
      Begin VB.Label Label1 
         Caption         =   "Red"
         Height          =   180
         Index           =   5
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   450
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   1680
      Width           =   975
   End
End
Attribute VB_Name = "frmColorDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private NewRGB  As COLORRGB
Private NewHSL  As COLORHSL
Private OldRGB  As COLORRGB

Event Change(R As Integer, G As Integer, B As Integer)
Event Active(R As Integer, G As Integer, B As Integer)

Private Sub cmdCancel_Click()
  
    RaiseEvent Change(OldRGB.R, OldRGB.G, OldRGB.B)
    Me.Hide
    
End Sub

Private Sub cmdOK_Click()
    
    Me.Hide
    
End Sub

Private Sub Form_Activate()
    
    RaiseEvent Active(OldRGB.R, OldRGB.G, OldRGB.B)
    NewRGB = OldRGB
    Call RefreshDisplay(tRGB)
    Call RefreshDisplay(tHSL)

End Sub

Private Sub Form_Load()

    Call FillHue
    
End Sub

Private Sub picRed_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 And X > -1 And X < 180 Then
        shpRed.Move X
        NewRGB.R = CInt(1.424581 * X)
        Call RefreshDisplay(tRGB)
    End If

End Sub

Private Sub picGreen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 And X > -1 And X < 180 Then
        shpGreen.Move X
        NewRGB.G = CInt(1.424581 * X)
        Call RefreshDisplay(tRGB)
    End If
    
End Sub

Private Sub picBlue_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 And X > -1 And X < 180 Then
        shpBlue.Move X
        NewRGB.B = CInt(1.424581 * X)
        Call RefreshDisplay(tRGB)
    End If

End Sub

Private Sub picHue_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 And X > -1 And X < 180 Then
        shpHue.Move X
        NewHSL.Hue = CInt(1.34 * X)
        Call RefreshDisplay(tHSL)
    End If
    
End Sub

Private Sub picSat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 And X > -1 And X < 180 Then
        shpSat.Move X
        shpSat.DrawMode = IIf(X < 60, vbBlackness, vbInvert)
        NewHSL.Sat = CInt(1.34 * X)
        Call RefreshDisplay(tHSL)
    End If
    
End Sub

Private Sub picLum_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 And X > -1 And X < 180 Then
        shpLum.Move X
        NewHSL.Lum = CInt(1.34 * X)
        Call RefreshDisplay(tHSL)
    End If
        
End Sub

Private Sub FillHue()

    Dim Red         As Single
    Dim Green       As Single
    Dim Blue        As Single
    
    Dim stepRed     As Single
    Dim stepGreen   As Single
    Dim stepBlue    As Single
    Dim Step        As Single
    Dim idx         As Integer
    Dim X           As Single
    Dim StepX       As Single

    Step = 6.375
    StepX = 0.754
    For idx = 0 To 240
        If idx < 120 Then
            Red = Step * (80 - idx)
        Else
            Red = Step * (idx - 160)
        End If
        
        If idx < 80 Then
            Green = Step * idx
        Else
            Green = Step * (160 - idx)
        End If
        
        If idx < 160 Then
            Blue = Step * (idx - 80)
        Else
            Blue = Step * (240 - idx)
        End If
        
        stepRed = ColorLimits(Red)
        stepGreen = ColorLimits(Green)
        stepBlue = ColorLimits(Blue)
        picHue.Line (X, 0)-(X + StepX, picHue.ScaleHeight), RGB(stepRed, stepGreen, stepBlue), BF
        X = X + StepX
    Next
    
End Sub

Private Sub FillGradient(picBox As PictureBox, StartColor As COLORRGB, EndColor As COLORRGB)
    
    Dim Red         As Single
    Dim Green       As Single
    Dim Blue        As Single
    Dim stepRed     As Single
    Dim stepGreen   As Single
    Dim stepBlue    As Single
    Dim idx         As Integer
    Dim X           As Single
    Dim StepX       As Single
    Dim StepColor   As COLORRGB

    StepX = 0.71 '(picBox.Width) * sng1Div255
    With StartColor
        If .R = EndColor.R Then
            Red = 0
        ElseIf .R < EndColor.R Then
            Red = (EndColor.R - .R) * sng1Div255
        Else
            Red = -(.R - EndColor.R) * sng1Div255
        End If
        If .G = EndColor.G Then
            Green = 0
        ElseIf .G < EndColor.G Then
            Green = (EndColor.G - .G) * sng1Div255
        Else
            Green = -(.G - EndColor.G) * sng1Div255
        End If
        If .B = EndColor.B Then
            Blue = 0
        ElseIf .B < EndColor.B Then
            Blue = (EndColor.B - .B) * sng1Div255
        Else
            Blue = -(.B - EndColor.B) * sng1Div255
        End If
        stepRed = .R
        stepGreen = .G
        stepBlue = .B
     End With
        For idx = 0 To 255
            picBox.Line (X, 0)-(X + StepX, picBox.ScaleHeight), RGB(stepRed, stepGreen, stepBlue), BF
            X = X + StepX
            stepRed = stepRed + Red
            stepGreen = stepGreen + Green
            stepBlue = stepBlue + Blue
        Next

End Sub

Private Sub FillLum(picBox As PictureBox, Color As COLORRGB)

    Dim Red         As Single
    Dim Green       As Single
    Dim Blue        As Single
    Dim stepRed     As Single
    Dim stepGreen   As Single
    Dim stepBlue    As Single
    Dim idx         As Integer
    Dim X           As Single
    Dim StepX       As Single
    Dim StepColor   As COLORRGB
    Dim StartColor  As COLORRGB
    Dim EndColor    As COLORRGB
    
    StepX = 0.3549 'picBox.Width / (2 * 255)
    StartColor = ColorSet(0, 0, 0)
    EndColor = Color
    With StartColor
        If .R = EndColor.R Then
            Red = 0
        ElseIf .R < EndColor.R Then
            Red = (EndColor.R - .R) * sng1Div255
        Else
            Red = -(.R - EndColor.R) * sng1Div255
        End If
        If .G = EndColor.G Then
            Green = 0
        ElseIf .G < EndColor.G Then
            Green = (EndColor.G - .G) * sng1Div255
        Else
            Green = -(.G - EndColor.G) * sng1Div255
        End If
        If .B = EndColor.B Then
            Blue = 0
        ElseIf .B < EndColor.B Then
            Blue = (EndColor.B - .B) * sng1Div255
        Else
            Blue = -(.B - EndColor.B) * sng1Div255
        End If
        stepRed = .R
        stepGreen = .G
        stepBlue = .B
        For idx = 0 To 255
            picBox.Line (X, 0)-(X + StepX, picBox.ScaleHeight), RGB(stepRed, stepGreen, stepBlue), BF
            X = X + StepX
            stepRed = stepRed + Red
            stepGreen = stepGreen + Green
            stepBlue = stepBlue + Blue
        Next
        StartColor = Color
        EndColor = ColorSet(255, 255, 255)
        If .R = EndColor.R Then
            Red = 0
        ElseIf .R < EndColor.R Then
            Red = (EndColor.R - .R) * sng1Div255
        Else
            Red = -(.R - EndColor.R) * sng1Div255
        End If
        If .G = EndColor.G Then
            Green = 0
        ElseIf .G < EndColor.G Then
            Green = (EndColor.G - .G) * sng1Div255
        Else
            Green = -(.G - EndColor.G) * sng1Div255
        End If
        If .B = EndColor.B Then
            Blue = 0
        ElseIf .B < EndColor.B Then
            Blue = (EndColor.B - .B) * sng1Div255
        Else
            Blue = -(.B - EndColor.B) * sng1Div255
        End If
        stepRed = .R
        stepGreen = .G
        stepBlue = .B
        For idx = 0 To 255
            picBox.Line (X, 0)-(X + StepX, picBox.ScaleHeight), RGB(stepRed, stepGreen, stepBlue), BF
            X = X + StepX
            stepRed = stepRed + Red
            stepGreen = stepGreen + Green
            stepBlue = stepBlue + Blue
        Next
    End With
    
End Sub

Public Sub RefreshDisplay(cType As ColorType)
    
    Select Case cType
        Case tRGB: NewHSL = RGBtoHSL(NewRGB)
        Case tHSL: NewRGB = HSLtoRGB(NewHSL)
    End Select
    With NewRGB
        Call FillGradient(picRed, ColorSet(0, .G, .B), ColorSet(255, .G, .B))
        Call FillGradient(picGreen, ColorSet(.R, 0, .B), ColorSet(.R, 255, .B))
        Call FillGradient(picBlue, ColorSet(.R, .G, 0), ColorSet(.R, .G, 255))
        shpRed.Left = CInt(0.702 * .R)
        shpGreen.Left = CInt(0.702 * .G)
        shpBlue.Left = CInt(0.702 * .B)
    End With
    With NewHSL
        Call FillGradient(picSat, ColorSet(128, 128, 128), NewRGB)
        Call FillLum(picLum, NewRGB)
        shpHue.Left = CInt(0.7458 * .Hue)
        shpSat.Left = CInt(0.7458 * .Sat)
        shpLum.Left = CInt(0.7458 * .Lum)
    End With
    RaiseEvent Change(NewRGB.R, NewRGB.G, NewRGB.B)
    
End Sub


