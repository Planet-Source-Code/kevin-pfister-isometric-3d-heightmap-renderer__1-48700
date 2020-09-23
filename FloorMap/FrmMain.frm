VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FloorMap Viewer"
   ClientHeight    =   6645
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   8070
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtHeightMap 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   7815
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "www.MatrixScreensavers.zion.me.uk"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   6360
      Width           =   2610
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Created by Kevin Pfister"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuRandom 
         Caption         =   "Random"
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuRender 
      Caption         =   "Render"
      Begin VB.Menu mnuMake 
         Caption         =   "Make3d"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Randomize Timer
    Call ResetHM
End Sub

Private Sub mnuAbout_Click()
    MsgBox ("FloorMap - Created by Kevin Pfister")
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuMake_Click()
    TxtHeightMap.Text = UCase(TxtHeightMap.Text)
    
    FrmMain.Caption = "FloorMap Viewer - Generating 3d Map"
    Call Render(TxtHeightMap.Text, False)
    FrmMain.Caption = "FloorMap Viewer"
End Sub

Private Sub mnuNew_Click()
    Call ResetHM
End Sub

Sub ResetHM()
    Dim GridS As Integer
    Dim Y As Integer
    GridS = InputBox("GridSize", , 10)
    TxtHeightMap.Text = ""
    For Y = 1 To GridS - 1
        TxtHeightMap.Text = TxtHeightMap.Text & String(GridS, "X") & vbNewLine
    Next
    TxtHeightMap.Text = TxtHeightMap.Text & String(GridS, "X")
End Sub

Sub ProcessLine(Text, LineNo)
    Dim X As Integer
    If LineNo = 1 Then
        ReDim HeightMap(HWidth, LineNo) As String
    Else
        ReDim Preserve HeightMap(HWidth, LineNo) As String
    End If
    For X = 1 To Len(Text)
        If Val(Mid(Text, X, 1)) = 0 And Mid(Text, X, 1) <> "0" Then
            HeightMap(X, LineNo) = "X"
        Else
            HeightMap(X, LineNo) = Mid(Text, X, 1)
        End If
    Next
End Sub

Private Sub mnuRandom_Click()
    Dim X As Integer
    Dim Y As Integer
    GridS = InputBox("GridSize", , 10)
    MaxH = InputBox("Max Height", , 2)
    FrmMain.Caption = "FloorMap Viewer - Generating Random Map"
    TxtHeightMap.Text = ""
    Dim TempRow As String
    For Y = 1 To GridS - 1
        TempRow = ""
        For X = 1 To GridS
            HeightM = Int(Rnd * MaxH)
            If HeightM = 0 Then
                TempRow = TempRow & "X"
            Else
                TempRow = TempRow & HeightM - 1
            End If
        Next
        TxtHeightMap.Text = TxtHeightMap.Text & TempRow & vbNewLine
        DoEvents
    Next
    TempRow = ""
    For X = 1 To GridS
        HeightM = Int(Rnd * MaxH)
        If HeightM = 0 Then
            TempRow = TempRow & "X"
        Else
            TempRow = TempRow & HeightM - 1
        End If
    Next
    TxtHeightMap.Text = TxtHeightMap.Text & TempRow
    FrmMain.Caption = "FloorMap Viewer"
End Sub

Sub Render(Map, Chr13 As Boolean)
    Dim X As Integer
    Dim Y As Integer
    Dim Lines() As String
    If Map = "" Then
        Call MsgBox("There is no Map Data", vbCritical)
        Exit Sub
    End If
    If InStr(1, Map, vbNewLine) = 0 Then
        HWidth = Len(Map)
        HHeight = 1
        Call ProcessLine(Map, 1)
    Else
        If Chr13 = True Then
            Lines() = Split(Map, Chr(13))
        Else
            Lines() = Split(Map, vbNewLine)
        End If
        If UBound(Lines()) = 0 Then
            'error
        Else
            HWidth = Len(Lines(0))
            HHeight = UBound(Lines()) + 1
            For Y = 0 To UBound(Lines())
                Call ProcessLine(Lines(Y), Y + 1)
            Next
        End If
    End If
    FrmRender.Show
    If HWidth > HHeight Then
        For X = HHeight To HWidth
            Call ProcessLine(String(HWidth, "X"), X)
        Next
        HHeight = HWidth
    ElseIf HWidth < HHeight Then
        ReDim TempArray(HWidth, HHeight) As String
        For Y = 1 To HHeight
            For X = 1 To HWidth
                TempArray(X, Y) = HeightMap(X, Y)
            Next
        Next
        ReDim HeightMap(HHeight, HHeight) As String
        For Y = 1 To HHeight
            For X = 1 To HWidth
                HeightMap(X, Y) = TempArray(X, Y)
            Next
            For X = HWidth + 1 To HHeight
                HeightMap(X, Y) = "X"
            Next
        Next
        HWidth = HHeight
    End If
    Call FrmRender.GenTile
    Call FrmRender.Height2Rot
    Call FrmRender.RenderMap
End Sub
