VERSION 5.00
Begin VB.Form FrmRender 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "3d Render"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9840
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   463
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   656
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicRock 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   10080
      Picture         =   "FrmRender.frx":0000
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   11
      Top             =   3720
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox PicGrass 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   10080
      Picture         =   "FrmRender.frx":0496
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   10
      Top             =   3120
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.CommandButton CmdRotL 
      Caption         =   "< Rotate"
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton CmdDef 
      Caption         =   "Default"
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton CmdRotR 
      Caption         =   "Rotate >"
      Height          =   375
      Left            =   6960
      TabIndex        =   7
      Top             =   6480
      Width           =   1335
   End
   Begin VB.PictureBox PicRaise 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   9960
      ScaleHeight     =   495
      ScaleWidth      =   1095
      TabIndex        =   6
      Top             =   2520
      Width           =   1095
   End
   Begin VB.PictureBox PicSide1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   9960
      ScaleHeight     =   495
      ScaleWidth      =   1095
      TabIndex        =   5
      Top             =   1920
      Width           =   1095
   End
   Begin VB.PictureBox PicSide 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   9960
      ScaleHeight     =   495
      ScaleWidth      =   1095
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.PictureBox PicTile 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   9960
      ScaleHeight     =   495
      ScaleWidth      =   1095
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.PictureBox PicIso 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   9960
      ScaleHeight     =   495
      ScaleWidth      =   1095
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   8400
      TabIndex        =   1
      Top             =   6480
      Width           =   1335
   End
   Begin VB.PictureBox PicGrid 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6300
      Left            =   120
      ScaleHeight     =   420
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   120
      Width           =   9600
   End
End
Attribute VB_Name = "FrmRender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdDef_Click()
    Call Height2Rot
    Call RenderMap
End Sub

Private Sub CmdOk_Click()
    Unload Me
End Sub

Sub GenTile()
    Dim X As Integer
    Dim Y As Integer
    'Creates Tiles
    PicIso.Cls
    PicRaise.Cls
    PicTile.Cls
    PicSide.Cls
    PicSide1.Cls
    
    TileW = 640 / HWidth
    TileH = 320 / HHeight
    TW = Int(TileW / 2)
    TH = Int(TileH / 2)
    
    PicIso.Width = TileW
    PicIso.Height = TileH
    PicTile.Width = TileW
    PicTile.Height = TileH
    PicRaise.Width = TileW
    PicRaise.Height = TileH
    PicSide.Width = TileW
    PicSide.Height = TileH
    PicSide1.Width = TileW + 1
    PicSide1.Height = TileH

    TileW = (640 / HWidth) * 15
    TileH = (320 / HHeight) * 15
    TW = Int(TileW / 2)
    TH = Int(TileH / 2)
    For X = 0 To TileW / 2
        PicIso.Line (X, TH - TH / TW * X)-(X, TH - TH / TW * X + (TH / TW * X) * 2), vbWhite
        PicIso.Line (TileW - X, TH - TH / TW * X)-(TileW - X, TH - TH / TW * X + (TH / TW * X) * 2), vbWhite
    Next
    
    PicTile.Line (TW, 0)-(TileW, TileH / 2), RGB(200, 200, 200)
    PicTile.Line (TW, 0)-(0, TileH / 2), RGB(200, 200, 200)
    PicTile.Line (TW, TileH)-(TileW, TileH / 2), RGB(200, 200, 200)
    PicTile.Line (TW, TileH)-(0, TileH / 2), RGB(200, 200, 200)
    
    For X = 0 To TileW / 2
        PicRaise.Line (X, TH - TH / TW * X)-(X, TH - TH / TW * X + (TH / TW * X) * 2), RGB(175, 175, 175)
        PicRaise.Line (TileW - X, TH - TH / TW * X)-(TileW - X, TH - TH / TW * X + (TH / TW * X) * 2), RGB(175, 175, 175)
    Next
    PicRaise.Line (TW, 0)-(TileW, TileH / 2), vbBlack
    PicRaise.Line (TW, 0)-(0, TileH / 2), vbBlack
    PicRaise.Line (TW, TileH)-(TileW, TileH / 2), vbBlack
    PicRaise.Line (TW, TileH)-(0, TileH / 2), vbBlack
    PicIso.Refresh
    For X = 0 To TileW / 2
        PicSide.Line (X, TH / TW * X)-(X, TH / TW * X + TH), vbWhite
        PicSide.Line (TileW - X, TH / TW * X)-(TileW - X, TH / TW * X + (TileH / 2)), vbWhite
    Next
    For X = 0 To TileW / 2
        PicSide1.Line (X, TH / TW * X)-(X, TH / TW * X + TH), RGB(100, 100, 100)
        PicSide1.Line (TileW - X, TH / TW * X)-(TileW - X, TH / TW * X + TH), RGB(100, 100, 100)
    Next
    PicSide1.Line (0, 0)-(TW, TH), RGB(150, 150, 150)
    PicSide1.Line (TW, TH)-(TileW, 0), RGB(150, 150, 150)
    PicSide1.Line (0, 0)-(0, TH), RGB(150, 150, 150)
    PicSide1.Line (TileW, 0)-(TileW, TileH / 2), RGB(150, 150, 150)
    PicSide1.Line (0, TileH / 2)-(TW, TileH), RGB(150, 150, 150)
    PicSide1.Line (TileW, TH)-(TW, TileH), RGB(150, 150, 150)
    PicSide1.Line (TW, TH)-(TW, TileH), RGB(150, 150, 150)
End Sub

Sub RenderMap()
    Dim X As Integer
    Dim Y As Integer
    Dim TileW As Double
    Dim TileH As Double
    Dim TW As Double
    Dim TH As Double
    
    Dim Before As Long
    Dim Surfaces As Long
    
    'Modified 3d IsoMap rendering Sub
    TileW = 640 / HWidth
    TileH = 320 / HHeight
    TW = TileW / 2
    TH = TileH / 2
    PicGrid.Cls
    Before = Timer
    For X = 1 To HWidth
        For Y = 1 To HHeight
            If RotMap(X, Y) = "X" Then
                Call BitBlt(PicGrid.hDC, 320 + X * TW - Y * TW - TW, Y * TH + X * TH + 64, TileW, TileH, PicIso.hDC, 0, 0, vbSrcPaint)
                Call BitBlt(PicGrid.hDC, 320 + X * TW - Y * TW - TW, Y * TH + X * TH + 64, TileW, TileH, PicTile.hDC, 0, 0, vbSrcAnd)
                Surfaces = Surfaces + 2
            Else
                Call BitBlt(PicGrid.hDC, 320 + X * TW - Y * TW - TW, Y * TH + X * TH + 64 - TH * (RotMap(X, Y) + 1), TileW, TileH, PicIso.hDC, 0, 0, vbSrcPaint)
                Call BitBlt(PicGrid.hDC, 320 + X * TW - Y * TW - TW, Y * TH + X * TH + 64 - TH * (RotMap(X, Y) + 1), TileW, TileH, PicRaise.hDC, 0, 0, vbSrcAnd)
                Surfaces = Surfaces + 2
                For Y1 = 1 To (RotMap(X, Y) + 1)
                    Call BitBlt(PicGrid.hDC, 320 + X * TW - Y * TW - TW, Y * TH + X * TH + 64 - TH * Y1 + TH, TileW + 1, TileH, PicSide.hDC, 0, 0, vbSrcPaint)
                    Call BitBlt(PicGrid.hDC, 320 + X * TW - Y * TW - TW, Y * TH + X * TH + 64 - TH * Y1 + TH, TileW + 1, TileH, PicSide1.hDC, 0, 0, vbSrcAnd)
                    Surfaces = Surfaces + 2
                Next
            End If
        Next
    Next
    PicGrid.Refresh
    FrmRender.Caption = "Rendered " & Surfaces & " Surfaces in " & Timer - Before & " seconds"
End Sub

Sub Height2Rot()
    Dim X As Integer
    Dim Y As Integer
    ReDim RotMap(HWidth, HHeight) As String
    ReDim RotMap1(HWidth, HHeight) As String
    For X = 1 To HWidth
        For Y = 1 To HHeight
            RotMap(X, Y) = HeightMap(X, Y)
        Next
    Next
End Sub

Private Sub CmdRotL_Click()
    Dim X As Integer
    Dim Y As Integer
    For X = 1 To HWidth
        For Y = 1 To HHeight
            RotMap1(X, Y) = RotMap(HHeight + 1 - Y, X)
        Next
    Next
    For X = 1 To HWidth
        For Y = 1 To HHeight
            RotMap(X, Y) = RotMap1(X, Y)
        Next
    Next
    Call RenderMap
End Sub

Private Sub CmdRotR_Click()
    Dim X As Integer
    Dim Y As Integer
    For X = 1 To HWidth
        For Y = 1 To HHeight
            RotMap1(X, Y) = RotMap(Y, HWidth + 1 - X)
        Next
    Next
    For X = 1 To HWidth
        For Y = 1 To HHeight
            RotMap(X, Y) = RotMap1(X, Y)
        Next
    Next
    Call RenderMap
End Sub
