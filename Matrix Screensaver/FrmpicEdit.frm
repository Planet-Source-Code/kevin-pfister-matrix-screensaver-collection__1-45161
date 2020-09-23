VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmpicEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Matrix Picture Editor"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   537
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   712
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Slider SldDif 
      Height          =   375
      Left            =   8880
      TabIndex        =   15
      Top             =   4800
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Min             =   1
      Max             =   100
      SelStart        =   10
      Value           =   10
   End
   Begin VB.CommandButton CmdResize 
      Caption         =   "Resize"
      Height          =   375
      Left            =   8880
      TabIndex        =   14
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton CmdBlackOut 
      Caption         =   "Black Out non-white"
      Height          =   375
      Left            =   8880
      TabIndex        =   13
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton CmdCreateGreen 
      Caption         =   "Create Green Mask"
      Height          =   375
      Left            =   8880
      TabIndex        =   12
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton CmdCreateWhite 
      Caption         =   "Create White Mask"
      Height          =   375
      Left            =   8880
      TabIndex        =   11
      Top             =   3840
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   8880
      Top             =   7440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdLarger 
      Caption         =   "Larger"
      Height          =   375
      Left            =   9720
      TabIndex        =   10
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton CmdSmaller 
      Caption         =   "Smaller"
      Height          =   375
      Left            =   8880
      TabIndex        =   9
      Top             =   3360
      Width           =   855
   End
   Begin VB.PictureBox PicCursor 
      AutoRedraw      =   -1  'True
      Height          =   1695
      Left            =   8880
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   109
      TabIndex        =   8
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton CmdColour 
      Caption         =   "Colour"
      Height          =   375
      Left            =   9360
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.PictureBox PicColour 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8880
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   6
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save Picture"
      Height          =   375
      Left            =   8880
      TabIndex        =   5
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton CmdOpen 
      Caption         =   "Load Picture"
      Height          =   375
      Left            =   8880
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.VScrollBar VScl 
      Enabled         =   0   'False
      Height          =   7575
      Left            =   8520
      SmallChange     =   5
      TabIndex        =   2
      Top             =   120
      Width           =   255
   End
   Begin VB.HScrollBar HScl 
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      SmallChange     =   5
      TabIndex        =   1
      Top             =   7680
      Width           =   8415
   End
   Begin VB.PictureBox PicFrame 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   7575
      Left            =   120
      ScaleHeight     =   505
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   561
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      Begin VB.PictureBox PicMatrix 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5655
         Left            =   0
         ScaleHeight     =   377
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   489
         TabIndex        =   3
         Top             =   0
         Width           =   7335
      End
   End
End
Attribute VB_Name = "FrmpicEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CursorSize As Integer
Dim CurColour As Long
Dim DrawE As Boolean
Private Sub CmdBlackOut_Click()
    Dim Red As Integer
    Dim Green As Integer
    Dim Blue As Integer
    Dim Red1 As Integer
    Dim Green1 As Integer
    Dim Blue1 As Integer
    Call DisableButtons
    For X = 1 To PicMatrix.Width
        For Y = 1 To PicMatrix.Height
            Colour = GetPixel(PicMatrix.hDC, X, Y)
            If Colour <> vbWhite Then
                'easy its the same colour ->change to white
                Call SetPixelV(PicMatrix.hDC, X, Y, vbBlack)
            End If
        Next
        DoEvents
    Next
    PicMatrix.Refresh
    Call EnableButtons
End Sub

Sub DisableButtons()
    CmdCreateWhite.Enabled = False
    CmdCreateGreen.Enabled = False
    CmdBlackOut.Enabled = False
End Sub

Sub EnableButtons()
    CmdCreateWhite.Enabled = True
    CmdCreateGreen.Enabled = True
    CmdBlackOut.Enabled = True
End Sub

Private Sub CmdColour_Click()
    DrawE = False
    CD1.ShowColor
    PicColour.BackColor = CD1.Color
    CurColour = CD1.Color
    Call UpdateCursor
    DrawE = True
End Sub

Private Sub CmdCreateGreen_Click()
    Dim Red As Integer
    Dim Green As Integer
    Dim Blue As Integer
    Dim Red1 As Integer
    Dim Green1 As Integer
    Dim Blue1 As Integer
    Call DisableButtons
    Call GetRgb(CurColour, Red, Green, Blue)
    For X = 1 To PicMatrix.Width
        For Y = 1 To PicMatrix.Height
            Colour = GetPixel(PicMatrix.hDC, X, Y)
            If Colour = CurColour Then
                'easy its the same colour ->change to green
                Call SetPixelV(PicMatrix.hDC, X, Y, RGB(0, 255, 0))
            Else
                'The colour may be totally different or very close
                'change colours close also to white(helps with a gradient
                Call GetRgb(Colour, Red1, Green1, Blue1)
                Diff = Abs(Red - Red1) + Abs(Green - Green1) + Abs(Blue - Blue1)
                If Diff < SldDif Then
                    Call SetPixelV(PicMatrix.hDC, X, Y, RGB(0, 255, 0))
                End If
            End If
        Next
        DoEvents
    Next
    PicMatrix.Refresh
    Call EnableButtons
End Sub

Private Sub CmdCreateWhite_Click()
    Dim Red As Integer
    Dim Green As Integer
    Dim Blue As Integer
    Dim Red1 As Integer
    Dim Green1 As Integer
    Dim Blue1 As Integer
    Call DisableButtons
    Call GetRgb(CurColour, Red, Green, Blue)
    For X = 1 To PicMatrix.Width
        For Y = 1 To PicMatrix.Height
            Colour = GetPixel(PicMatrix.hDC, X, Y)
            If Colour = CurColour Then
                'easy its the same colour ->change to white
                Call SetPixelV(PicMatrix.hDC, X, Y, vbWhite)
            Else
                'The colour may be totally different or very close
                'change colours close also to white(helps with a gradient
                Call GetRgb(Colour, Red1, Green1, Blue1)
                Diff = Abs(Red - Red1) + Abs(Green - Green1) + Abs(Blue - Blue1)
                If Diff < SldDif Then
                    Call SetPixelV(PicMatrix.hDC, X, Y, vbWhite)
                End If
            End If
        Next
        DoEvents
    Next
    PicMatrix.Refresh
    Call EnableButtons
End Sub

Private Sub CmdLarger_Click()
    If CursorSize = 50 Then
        MsgBox ("Cursor Size at maximum")
    Else
        CursorSize = CursorSize + 2
    End If
    Call UpdateCursor
End Sub

Private Sub CmdOpen_Click()
    DrawE = False
    CD1.ShowOpen
    FileName = CD1.FileName
    If FileName = "" Then Exit Sub
    PicMatrix.Picture = LoadPicture(FileName)
    HScl.Enabled = False
    PicMatrix.Left = 0
    PicMatrix.Top = 0
    HScl.Value = 0
    VScl.Value = 0
    If PicMatrix.Width > PicFrame.Width Then
        HScl.Max = PicMatrix.Width - PicFrame.Width
        HScl.Enabled = True
    End If
    VScl.Enabled = False
    If PicMatrix.Height > PicFrame.Height Then
        VScl.Max = PicMatrix.Height - PicFrame.Height
        VScl.Enabled = True
    End If
    DrawE = True
End Sub

Private Sub CmdResize_Click()
    XSize = InputBox("Width", "Width", "1024")
    YSize = InputBox("Height", "Height", "768")
    If XSize = "" Or YSize = "" Then Exit Sub
    If Val(XSize) > PicMatrix.Width Then Exit Sub
    If Val(YSize) > PicMatrix.Height Then Exit Sub
    Call DisableButtons
    For X = 1 To Val(XSize)
        For Y = 1 To Val(YSize)
            Call SetPixelV(PicMatrix.hDC, X, Y, GetPixel(PicMatrix.hDC, PicMatrix.Width / Val(XSize) * X, PicMatrix.Height / Val(YSize) * Y))
        Next
        DoEvents
    Next
    PicMatrix.Width = Val(XSize)
    PicMatrix.Height = Val(YSize)
    PicMatrix.Refresh
    Call EnableButtons
    If PicMatrix.Width > PicFrame.Width Then
        HScl.Max = PicMatrix.Width - PicFrame.Width
        HScl.Enabled = True
    End If
    VScl.Enabled = False
    If PicMatrix.Height > PicFrame.Height Then
        VScl.Max = PicMatrix.Height - PicFrame.Height
        VScl.Enabled = True
    End If
End Sub

Private Sub CmdSave_Click()
    DrawE = False
    CD1.ShowSave
    FileName = CD1.FileName
    If FileName = "" Then Exit Sub
    Call SavePicture(PicMatrix.Image, FileName)
    DrawE = True
End Sub

Private Sub CmdSmaller_Click()
    If CursorSize = 2 Then
        MsgBox ("Cursor Size at minimum")
    Else
        CursorSize = CursorSize - 2
    End If
    Call UpdateCursor
End Sub

Sub UpdateCursor()
    PicCursor.Cls
    For X = 1 To CursorSize
        For Y = 1 To CursorSize
            Call SetPixelV(PicCursor.hDC, Int(PicCursor.Width / 2) - CursorSize / 2 + X, Int(PicCursor.Height / 2) - CursorSize / 2 + Y, CurColour)
        Next
    Next
    PicCursor.Refresh
End Sub

Private Sub Form_Load()
    CursorSize = 2
    Call UpdateCursor
    CurColour = PicColour.BackColor
    DrawE = True
End Sub

Private Sub HScl_Change()
    PicMatrix.Left = 0 - HScl.Value
End Sub

Private Sub PicMatrix_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If DrawE = True Then
            For X1 = 1 To CursorSize
                For Y1 = 1 To CursorSize
                    Call SetPixelV(PicMatrix.hDC, X - CursorSize / 2 + X1, Y - CursorSize / 2 + Y1, CurColour)
                Next
            Next
            PicMatrix.Refresh
        End If
    ElseIf Button = 2 Then
        CurColour = GetPixel(PicMatrix.hDC, X, Y)
        PicColour.BackColor = CurColour
        Call UpdateCursor
    End If
End Sub

Private Sub VScl_Change()
    PicMatrix.Top = 0 - VScl.Value
End Sub
