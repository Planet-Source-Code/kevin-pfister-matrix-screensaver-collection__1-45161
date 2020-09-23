VERSION 5.00
Begin VB.Form FrmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "The Matrix By Kevin Pfister"
   ClientHeight    =   9660
   ClientLeft      =   120
   ClientTop       =   690
   ClientWidth     =   10110
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00008000&
   Icon            =   "matrix.frx":0000
   ScaleHeight     =   40.25
   ScaleMode       =   4  'Character
   ScaleWidth      =   84.25
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer TmrLoad 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   600
      Top             =   1080
   End
   Begin VB.PictureBox PicMovieBuf 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   2640
      ScaleHeight     =   233
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   345
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.PictureBox PicMovie 
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   2640
      ScaleHeight     =   233
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   345
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Timer TmrHallway 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Tag             =   "Falling Code"
      Top             =   1080
   End
   Begin VB.Timer TmrFrameRate 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1080
      Tag             =   "Falling Code"
      Top             =   600
   End
   Begin VB.PictureBox PicMatrix 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   1560
      Picture         =   "matrix.frx":0442
      ScaleHeight     =   79
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   68
      TabIndex        =   0
      Tag             =   "Falling Code"
      Top             =   120
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Timer TmrApply 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   600
      Tag             =   "Startup"
      Top             =   600
   End
   Begin VB.Timer TmrMain3 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   120
      Tag             =   "Knock"
      Top             =   600
   End
   Begin VB.Timer TmrMain2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1080
      Tag             =   "Tracing"
      Top             =   120
   End
   Begin VB.Timer TmrMain1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   600
      Tag             =   "Tracing"
      Top             =   120
   End
   Begin VB.Timer TmrMain 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Tag             =   "Falling Code"
      Top             =   120
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'#######################################################################
'Matrix Screensavers made by Kevin Pfister
'#######################################################################

'READ FIRST

'The program is preset to be a screensaver, this means that it will not run in VB
'To make it work, in properties change the startup object form Sub Main to frmMain.
'If you stop the program from running in Vb ie. Ctrl+Break the cursor will be
'invisible

'General Variables
Dim IntLastXPos As Integer    'For use in Checking the mouse movements
Dim IntLastYPos As Integer 'For use in Checking the mouse movements
Dim IntFrames As Integer
Dim IntFrameRate As Integer
Dim IntActiveScreensaver As Integer
Dim IntX As Integer
Dim IntY As Integer

'Falling Code Variables
Dim IntBackGroundPic() As Integer
Dim IntLengthOfDrop() As Integer   'Length of Dropping column
Dim IntLeading() As Integer   'To hold the IntLeading letters
Dim IntLetter() As Integer   'The symbol
Dim IntColour() As Integer    'The IntColour of the symbol
Dim IntFSize() As Integer    'The IntColour of the symbol
Dim IntFntSize() As Integer
Dim IntLngWaitLngBeforeClear() As Integer        'To hold the length of time LngBefore the symbol fades
Dim IntMaxLength As Integer   'The maximum length of the column
Dim IntMaxLngWait As Integer   'The maximum Waiting time Before clearing
Dim IntDropCols As Integer   'The StrNumber of dropping coloumns
Dim IntFadeSpeed As Integer   'The fading speed of the symbols
Dim IntFromTop As Integer   'If the column starts falling from the top or from a random position
Dim IntWillFade As Integer   'Will the letter fade or not
Dim IntMultipleColours As Integer   'Is it single or multiple Colours
Dim IntFontSize As Integer
Dim LngOneCol As Long
Dim BlnUseBackGround As Boolean
Dim BlnMask As Boolean
Dim BlnHybrid As Boolean
Dim StrImageFile As String
Dim IntCodeColour As Integer
Dim IntReloadStyle As Integer
Dim IntDifferentFontSizes As Integer
Dim IntSuperSpeed As Integer
Dim IntSmallMode As Integer
Dim IntMovieFrames As Integer
Dim StrMoviePath As String
Dim BlnMovie As Boolean
Dim IntFrameNo As Integer
Dim IntMovieOffset As Integer
Dim Frames() As Integer
Dim i

'Tracing Variables
Dim IntYNums(1 To 30) As Integer
Dim IntXNums(1 To 60) As Integer
Dim IntTextDone As Integer   'How much has been drawn to the screen already
Dim IntSTextF As Integer
Dim StrPhoneNo(1 To 11) As String   'The seperate parts of the phone StrNumber
Dim IntAnim As Integer    'Change draw IntColour (1 -> 0 -> 1 -> 0...)
Dim LngXSpace As Long  'Where the StrNumbers are to be drawn
Dim LngYSpace  As Long 'Where the StrNumbers are to be drawn
Dim LngRanNum As Long  'If random StrNumber was choosen
Dim LngTraceCol As Long
Dim LngYCoord As Long
Dim LngXCoord(1 To 11) As Long
Dim LngWait As Long
Dim BlnCols(60) As Boolean    'The different columns, when clearing
Dim BlnPhoneOn(1 To 11) As Boolean 'To IntCheck if the phone StrNumber is to be shown
Dim StrNumber As String    'The phone StrNumber to be traced
Dim StrNumbers(60, 30) As Integer 'all the StrNumbers
Dim StrStartText As String   'Text to be drawn to the screen

'Knock, Knock Variables
Dim IntTxtSpeed(4) As Integer
Dim IntMatrixDone As Integer
Dim IntCurrentTxt As Integer
Dim StrTxtMatrix(4) As String

'Hallway Variables
Dim MatrixPeople(70, 80) As Long
Dim TempX As Integer
Dim TempY As Integer

'#######################################################################
'General Section
'#######################################################################

Private Sub Form_Load()
    Dim IntCurrent As Integer
    Dim IntDoFill As Integer
    Dim IntPNo As Integer
    ShowCursor (0)  'Make the cursor invisible
    FrmMain.WindowState = 2 'make the screensaver maximised
    ForeColor = RGB(0, 220, 0)  'Change the forecolor to the default shade of green
    
    '#######################################################################
    'General Settings
    '#######################################################################
    
    IntActiveScreensaver = Val(GetSetting("Kevin Pfister's Matrix", "Options", "Which", 0))
    
    '#######################################################################
    'Falling Code Settings
    '#######################################################################
    'This section gets the settings from the registry and stores them in the variables
    IntReloadStyle = GetSetting("Kevin Pfister's Matrix", "Options", "Reloaded", 0)
    IntMaxLength = GetSetting("Kevin Pfister's Matrix", "Drops", "MaxDrop", 100)   'Retieve the Maximum length of the columns
    IntMaxLngWait = GetSetting("Kevin Pfister's Matrix", "Drops", "BeforeClean", 200)   'Retieve the maximum LngWaiting time LngBefore clearing the symbol
    IntDropCols = GetSetting("Kevin Pfister's Matrix", "Drops", "DropsRunning", 25)       'Retieve the StrNumber of dropping columns
    IntFadeSpeed = GetSetting("Kevin Pfister's Matrix", "Drops", "FadeSpeed", 2)          'Retieve the fading speed of the columns
    IntFromTop = GetSetting("Kevin Pfister's Matrix", "Options", "FromTop", 1)   'Retieve if the columns start from the top or from a random position
    IntFontSize = Val(GetSetting("Kevin Pfister's Matrix", "Options", "Size", "8"))   'Retieve font size
    IntWillFade = GetSetting("Kevin Pfister's Matrix", "Colour", "Fade", 0)       'Retieve if the symbols fade or not
    IntMultipleColours = GetSetting("Kevin Pfister's Matrix", "Colour", "MColours", 1)   'Retieve if it are different shades of green
    TmrMain.Interval = 1000 / GetSetting("Kevin Pfister's Matrix", "Options", "Frame Rate", 100)
    IntCodeColour = GetSetting("Kevin Pfister's Matrix", "Options", "Colour", 1)
    IntDifferentFontSizes = GetSetting("Kevin Pfister's Matrix", "Options", "Dif Size", 0)
    IntSuperSpeed = GetSetting("Kevin Pfister's Matrix", "Options", "Super", 0)
    
    IntMovieFrames = GetSetting("Kevin Pfister's Matrix", "Options", "MovieFrames", 10)
    StrMoviePath = GetSetting("Kevin Pfister's Matrix", "Options", "MoviePath", 0)
    IntMovieOffset = GetSetting("Kevin Pfister's Matrix", "Options", "MovieOffset", 2)
    FrmMain.ScaleWidth = FrmMain.ScaleWidth * 10 / IntFontSize
    FrmMain.ScaleHeight = FrmMain.ScaleHeight * 10 / IntFontSize
    FrmMain.WindowState = 2
    If Val(GetSetting("Kevin Pfister's Matrix", "Options", "Style", 1)) = 1 Then
        BlnUseBackGround = True
        StrImageFile = GetSetting("Kevin Pfister's Matrix", "Options", "BckImage", "C:\Agent.jpg")
    ElseIf Val(GetSetting("Kevin Pfister's Matrix", "Options", "Style", 1)) = 2 Then
        BlnMask = True
        StrImageFile = GetSetting("Kevin Pfister's Matrix", "Options", "MaskImage", "C:\Agent.jpg")
    ElseIf Val(GetSetting("Kevin Pfister's Matrix", "Options", "Style", 1)) = 3 Then
        BlnHybrid = True
        StrImageFile = GetSetting("Kevin Pfister's Matrix", "Options", "MaskImage", "C:\Agent.jpg")
    ElseIf Val(GetSetting("Kevin Pfister's Matrix", "Options", "Style", 1)) = 4 Then
        BlnMovie = True
    End If
    
    '#######################################################################
    'Tracing Settings
    '#######################################################################
    
    StrNumber = GetSetting("Kevin Pfister's Matrix", "Options", "StrNumber", "0000000000")
    LngRanNum = GetSetting("Kevin Pfister's Matrix", "Options", "Random", 1)
    LngTraceCol = RGB(0, 220, 0)
    LngXSpace = Width / 45
    LngYSpace = Height / 35
    LngYCoord = LngYSpace * 3
    For IntX = 1 To 11
        LngXCoord(IntX) = LngXSpace * (2 + IntX)
    Next
    For IntX = 1 To 60
        IntXNums(IntX) = LngXSpace * (2 + IntX)
    Next
    For IntY = 1 To 30
        IntYNums(IntY) = LngYSpace * (4 + IntY)
    Next
    '#######################################################################
    'Knock Knock Neo Settings
    '#######################################################################
    
    StrTxtMatrix(1) = "Wake up,  Neo. . ."
    IntTxtSpeed(1) = 150
    StrTxtMatrix(2) = "The Matrix has you. . ."
    IntTxtSpeed(2) = 150
    StrTxtMatrix(3) = "Follow the white rabbit."
    IntTxtSpeed(3) = 150
    StrTxtMatrix(4) = "Knock,  Knock,  Neo.."
    IntTxtSpeed(4) = 1
    IntCurrent = 1
    
    Randomize Timer 'randomize the screensaver
    
    If IntActiveScreensaver = 0 Then  'Falling Code
        TmrApply.Enabled = True
    ElseIf IntActiveScreensaver = 1 Then 'Tracing
        For IntDoFill = 1 To 60
            BlnCols(IntDoFill) = 1
        Next
        For IntPNo = 1 To 11
            If LngRanNum = 1 Then
                StrPhoneNo(IntPNo) = Int(Rnd * 9)
            Else
                StrPhoneNo(IntPNo) = Mid(StrNumber, IntPNo, 1)
            End If
        Next
        StrStartText = "Call Trans opt: Rec " + Str$(Date) + " " + Str$(Time) + " Rec:Log> "
        ForeColor = RGB(0, 220, 0)
        ScaleMode = 1
        
        Font = "MS Serif"
        TmrMain1.Enabled = True
    ElseIf IntActiveScreensaver = 2 Then 'Knock,Knock
        Font = "Arial"
        ForeColor = &H9BAC9B
        TmrMain3.Enabled = True
        IntCurrentTxt = 1
    ElseIf IntActiveScreensaver = 3 Then
        TmrHallway.Enabled = True
    End If
End Sub

'#######################################################################
'Falling Code Section
'#######################################################################

Private Sub TmrLoad_Timer()
    'This loads each of the frames of the movie and then copies it to another picturebox where the pixel
    'values can be extracted
    If TmrLoad.Interval = 10000 Then    '10 Seconds at the beginning to allow video to load
        TmrLoad.Interval = 100
    End If
    IntFrameNo = IntFrameNo + 1
    FrmMain.Cls
    FrmMain.Print "Saving Frame:" & Str(IntFrameNo) & " /" & Str(IntMovieFrames)
    'Below sends the command to play the selected frame
    i = mciSendString("play video1 from" & Str((IntFrameNo * IntMovieOffset) - 1) & " to" & Str((IntFrameNo * IntMovieOffset)), 0&, 0, 0)
    'Copys it to the other picture box, I found it won't work if you just use
    'one picturebox
    BitBlt PicMovieBuf.hDC, 0, 0, PicMovieBuf.ScaleWidth, PicMovieBuf.ScaleHeight, PicMovie.hDC, 0, 0, vbSrcCopy
    PicMovieBuf.Refresh
    Dim R1 As Integer
    Dim G1 As Integer
    Dim B1 As Integer
    Dim Temp As Long
    'Save the image to the frame array so it can be shown
    For IntX = 1 To FrmMain.ScaleWidth
        For IntY = 1 To FrmMain.ScaleHeight
            Temp = GetPixel(PicMovieBuf.hDC, Int(PicMovieBuf.ScaleWidth / FrmMain.ScaleWidth * IntX), Int(PicMovieBuf.ScaleHeight / FrmMain.ScaleHeight * IntY))
            GetRgb Temp, R1, G1, B1
            Temp = Int((R1 + G1 + B1) / 3)
            Frames(IntFrameNo, IntX, IntY + 4) = Temp
            Frames(IntFrameNo, IntX, IntY + 4) = (Frames(IntFrameNo, IntX, IntY + 4) + 1) / 100 * 80 + 20
        Next
    Next
    If IntFrameNo = IntMovieFrames Then 'Has it reached the number of selected frames
        IntFrameNo = 0
        TmrLoad.Enabled = False
        PicMovie.Visible = False
        i = mciSendString("close video1", 0&, 0, 0) 'Close the video Important!!
        Font = "Matrix"   'Use the Matrix Font
        Font.Size = IntFontSize
        FrmMain.Cls
        Call MovieStartup
    End If
End Sub

Sub MovieStartup()
    'This starts up the scrolling effect of the code
    Dim DoRand As Integer
    Dim XR As Integer
    Dim YR As Integer
    For DoRand = 1 To IntDropCols 'Create the starting IntDrops
        XR = Int(Rnd * FrmMain.ScaleWidth) + 1  'The IntX position
        YR = Int(Rnd * (FrmMain.ScaleHeight + 5)) + 1   'The IntY position
        IntLengthOfDrop(XR, YR) = Int(Rnd * IntMaxLength)      'The Length of the drop
        IntLeading(XR, YR) = 1 'Make it a IntLeading symbol
        IntLetter(XR, YR) = Int(Rnd * 43) + 65         'Set the letter/symbol
        IntColour(1, XR, YR) = IntBackGroundPic(XR, YR) + Rnd * 20 'Set the IntColour
    Next
    'This just allows the code to fall without being displayed to give a better effect
    For XR = 1 To 50
        Call MovieCol
    Next
    'Puts the sequence to where it started
    IntFrameNo = IntFrameNo - 50
    
    'Start to record the framerate
    TmrFrameRate.Enabled = True
    If IntSuperSpeed = 0 Then
        TmrMain.Enabled = True
        Exit Sub
    End If
    Do
        DoEvents
        FrmMain.WindowState = 2
        FrmMain.Cls
        Call MovieCol
        For IntX = 1 To FrmMain.ScaleWidth
            For IntY = 1 To FrmMain.ScaleHeight + 5
                If IntLetter(IntX, IntY) <> 0 Then
                    Call ShowColor(IntX, IntY)
                End If
            Next
        Next
        IntFrames = IntFrames + 1
    Loop
End Sub

Private Sub TmrMain_Timer()
    'Main loop timer if Superspeed is disabled
    FrmMain.WindowState = 2
    If BlnMovie = True Then
        FrmMain.Cls
        Call MovieCol
        For IntX = 1 To FrmMain.ScaleWidth
            For IntY = 1 To FrmMain.ScaleHeight + 5
                If IntLetter(IntX, IntY) <> 0 Then
                    If IntLeading(IntX, IntY) <> 1 Then
                        Call ShowColor(IntX, IntY)
                    End If
                End If
            Next
        Next
    Else
        If IntMultipleColours = 0 Then
            Call OneIntColour
        Else
            Call MoreThanOneColour
        End If
    End If
    IntFrames = IntFrames + 1
End Sub

Sub MovieCol()
    'Because the Movie Feature uses its own layout, it needs its own Subroutine
    IntFrameNo = IntFrameNo + 1
    If IntFrameNo > IntMovieFrames Then
        IntFrameNo = 1
    End If
    Dim IntDrops As Integer
    Dim IntMakeNew As Integer
    For IntX = 1 To FrmMain.ScaleWidth
        For IntY = 1 To FrmMain.ScaleHeight + 5
            If IntLetter(IntX, IntY) <> 0 Then
                IntColour(1, IntX, IntY) = Frames(IntFrameNo, IntX, IntY) + Rnd * 40
                If IntLeading(IntX, IntY) = 1 Then 'Is it IntLeading
                    If IntY <= FrmMain.ScaleHeight + 4 Then 'Is it smaller than the screen height
                        If IntLengthOfDrop(IntX, IntY) > 0 Then 'Is there still IntDrops in this column
                            If IntLetter(IntX, IntY + 1) <> 0 Then
                                IntLeading(IntX, IntY + 1) = 0
                                IntLngWaitLngBeforeClear(IntX, IntY + 1) = 0
                                IntLetter(IntX, IntY + 1) = 0
                                IntColour(1, IntX, IntY + 1) = 0
                            End If
                            IntLengthOfDrop(IntX, IntY + 1) = IntLengthOfDrop(IntX, IntY) - 1
                            IntLeading(IntX, IntY) = 0
                            IntLeading(IntX, IntY + 1) = 2
                            IntLetter(IntX, IntY + 1) = Int(Rnd * 43) + 65
                            IntColour(1, IntX, IntY + 1) = Frames(IntFrameNo, IntX, IntY + 1) + Rnd * 20 'Set the IntColour
                            IntLngWaitLngBeforeClear(IntX, IntY) = IntMaxLngWait / 2 + Rnd(IntMaxLngWait / 2)
                        Else    'End of Drop(Kill Letter/Symbol)
                            If IntLeading(IntX, IntY) = 1 Then
                                IntLeading(IntX, IntY) = 0
                                IntLngWaitLngBeforeClear(IntX, IntY) = 0
                                IntLetter(IntX, IntY) = 0
                                IntColour(1, IntX, IntY) = 0
                            End If
                        End If
                    Else
                        IntLeading(IntX, IntY) = 0
                        IntLngWaitLngBeforeClear(IntX, IntY) = 0
                        IntLetter(IntX, IntY) = 0
                        IntColour(1, IntX, IntY) = 0
                    End If
                ElseIf IntLeading(IntX, IntY) = 2 Then
                    IntLeading(IntX, IntY) = 1
                    IntDrops = IntDrops + 1
                End If
                If IntLngWaitLngBeforeClear(IntX, IntY) > 0 Then 'Is the Letter/Symbol dieing?
                    IntLngWaitLngBeforeClear(IntX, IntY) = IntLngWaitLngBeforeClear(IntX, IntY) - 1
                    If IntLngWaitLngBeforeClear(IntX, IntY) = 0 Or IntColour(1, IntX, IntY) = 0 Then
                        IntLeading(IntX, IntY) = 0
                        IntLngWaitLngBeforeClear(IntX, IntY) = 0
                        IntLetter(IntX, IntY) = 0
                        IntColour(1, IntX, IntY) = 0
                    End If
                End If
            End If
        Next
    Next
    If IntDrops < IntDropCols Then
        For IntMakeNew = IntDrops To IntDropCols
            IntX = Int(Rnd * FrmMain.ScaleWidth) + 1
            If IntFromTop = 1 Then
                IntY = Int(Rnd * 5) + 1
            Else
                IntY = Int(Rnd * FrmMain.ScaleHeight) + 1
            End If
            If IntLetter(IntX, IntY) > 0 Then
                IntLeading(IntX, IntY) = 0
                IntLngWaitLngBeforeClear(IntX, IntY) = 0
                IntLetter(IntX, IntY) = 0
                IntColour(1, IntX, IntY) = 0
            End If
            IntLengthOfDrop(IntX, IntY) = Int(Rnd * IntMaxLength)
            IntLeading(IntX, IntY) = 1
            IntLetter(IntX, IntY) = 64 + Int(Rnd * 26)
            IntColour(1, IntX, IntY) = Frames(IntFrameNo, IntX, IntY + 1) + Rnd * 20
        Next
    End If
End Sub

Private Sub TmrApply_Timer()
    Dim DoRand As Integer
    Dim XR As Integer
    Dim YR As Integer
    Dim Temp As Long
    Dim Loading As Integer
    Dim AddNum As Integer
    
    TmrApply.Enabled = False
    
    If IntCodeColour = 0 Then
        LngOneCol = RGB(150, 0, 0)
    ElseIf IntCodeColour = 1 Then
        LngOneCol = RGB(0, 150, 0)
    ElseIf IntCodeColour = 2 Then
        LngOneCol = RGB(0, 0, 150)
    End If
    'Change the variable sizes to fit the screen size
    'The extra 1 width is to stop errors when randomly choosing locations and the extra 5 height is for overhang
    ReDim IntLengthOfDrop(FrmMain.ScaleWidth + 1, FrmMain.ScaleHeight + 5) As Integer
    ReDim IntLeading(FrmMain.ScaleWidth + 1, FrmMain.ScaleHeight + 5) As Integer
    ReDim IntLetter(FrmMain.ScaleWidth + 1, FrmMain.ScaleHeight + 5) As Integer
    ReDim IntColour(1 To 2, FrmMain.ScaleWidth + 1, FrmMain.ScaleHeight + 5) As Integer
    ReDim IntLngWaitLngBeforeClear(FrmMain.ScaleWidth + 1, FrmMain.ScaleHeight + 5) As Integer
    ReDim IntBackGroundPic(FrmMain.ScaleWidth + 1, FrmMain.ScaleHeight + 5) As Integer
    ReDim IntFntSize(FrmMain.ScaleWidth + 1, FrmMain.ScaleHeight + 5) As Integer
    If Dir(StrImageFile) = "" Then 'Check to see if file exists
        'If it doesn't use the normal Mode
        BlnUseBackGround = False
        BlnMask = False
        BlnHybrid = False
    End If
    If Dir(StrMoviePath) = "" Then  'Check if the movie exists
        'If not, don't use the movie
        BlnMovie = False
    End If
    If BlnMovie = True Then
        FrmMain.ForeColor = vbGreen
        Font = "Arial"
        Font.Size = 12
        Dim Holder As String
        Dim Todo As String
        'Place the Movie in the center so the User can see it while the program loads
        PicMovie.Left = FrmMain.ScaleWidth / 2 - PicMovie.Width / 2
        PicMovie.Top = FrmMain.ScaleHeight / 2 - PicMovie.Height / 2
        PicMovie.Visible = True
        FrmMain.Cls
        FrmMain.Print "Loading Movie"
        DoEvents
        'Close any movie file that is already open
        i = mciSendString("close all", 0&, 0, 0)
        Holder = PicMovie.hWnd & " Style " & &H40000000
        'Below is very important, all MCI related filenames must be Shortnames or they won't work
        StrMoviePath = GetShortName(StrMoviePath)
        'This opens the Movie and places/Resizes it into the Picturebox where the frames are extracted
        Todo = "open " & StrMoviePath & " Type avivideo Alias video1 parent " & Holder
        i = mciSendString(Todo, 0&, 0, 0)
        'Place the video in the picturebox
        i = mciSendString("put video1 window at 0 0 " & PicMovie.ScaleWidth & " " & PicMovie.ScaleHeight, 0&, 0, 0)
        
        'This checks to see if there is enough frames to fill
        'If the user has selected more frames than there is, it will be adjusted
        Dim mssg As String * 255
        i = mciSendString("set video1 time format frames", 0&, 0, 0)
        i = mciSendString("status video1 length", mssg, 255, 0)
        If IntMovieFrames * IntMovieOffset > Val(mssg) Then
            IntMovieFrames = Val(mssg) / IntMovieOffset
        End If
        'Show the first frame
        i = mciSendString("play video1 from" & Str(0) & " to" & Str(1), 0&, 0, 0)
        'Turn off the videos audio
        i = mciSendString("set video1 audio all off", 0&, 0, 0)
        'Redefine the array to store the information from the Frames
        ReDim Frames(IntMovieFrames, FrmMain.ScaleWidth + 1, FrmMain.ScaleHeight + 5) As Integer
        TmrLoad.Enabled = True
        Exit Sub
    End If
    If BlnUseBackGround = True And IntMultipleColours = 1 Then
        PicMatrix.Picture = LoadPicture(StrImageFile)
        Font = "Arial"
        Font.Size = 12
        AddNum = 1
        Dim R1 As Integer
        Dim G1 As Integer
        Dim B1 As Integer
        For IntX = 1 To FrmMain.ScaleWidth
            Loading = Loading + AddNum
            If Loading = 15 Or Loading = 0 Then
                AddNum = -AddNum
            End If
            FrmMain.Cls
            FrmMain.CurrentX = FrmMain.ScaleWidth / 2 - 15
            FrmMain.CurrentY = FrmMain.ScaleHeight / 2
            FrmMain.Print "Processing Image " & String(Loading, ".")
            DoEvents
            For IntY = 1 To FrmMain.ScaleHeight
                Temp = GetPixel(PicMatrix.hDC, Int(PicMatrix.ScaleWidth / FrmMain.ScaleWidth * IntX), Int(PicMatrix.ScaleHeight / FrmMain.ScaleHeight * IntY))
                GetRgb Temp, R1, G1, B1
                Temp = Int((R1 + G1 + B1) / 3)
                IntBackGroundPic(IntX, IntY + 4) = Temp
                IntBackGroundPic(IntX, IntY + 4) = (IntBackGroundPic(IntX, IntY + 4) + 1) / 100 * 80 + 20
            Next
        Next
        FrmMain.Cls
        PicMatrix.Picture = LoadPicture("") 'Free up memory
    End If
    If BlnMask = True And IntMultipleColours = 1 Then
        PicMatrix.Picture = LoadPicture(StrImageFile)
        Font = "Arial"
        Font.Size = 12
        AddNum = 1
        For IntX = 1 To FrmMain.ScaleWidth
            Loading = Loading + AddNum
            If Loading = 15 Or Loading = 0 Then
                AddNum = -AddNum
            End If
            FrmMain.Cls
            FrmMain.CurrentX = FrmMain.ScaleWidth / 2 - 15
            FrmMain.CurrentY = FrmMain.ScaleHeight / 2
            FrmMain.Print "Processing Image " & String(Loading, ".")
            DoEvents
            For IntY = 1 To FrmMain.ScaleHeight
                Temp = GetPixel(PicMatrix.hDC, Int(PicMatrix.ScaleWidth / FrmMain.ScaleWidth * IntX), Int(PicMatrix.ScaleHeight / FrmMain.ScaleHeight * IntY))
                GetRgb Temp, R1, G1, B1
                Temp = Int((R1 + G1 + B1) / 3)
                If Temp < 128 Then
                    IntBackGroundPic(IntX, IntY + 4) = 0
                Else
                    IntBackGroundPic(IntX, IntY + 4) = 1
                End If
            Next
        Next
        FrmMain.Cls
        PicMatrix.Picture = LoadPicture("") 'Free up memory
    End If
    If BlnHybrid = True And IntMultipleColours = 1 Then
        PicMatrix.Picture = LoadPicture(StrImageFile)
        Font = "Arial"
        Font.Size = 12
        AddNum = 1
        For IntX = 1 To FrmMain.ScaleWidth
            Loading = Loading + AddNum
            If Loading = 15 Or Loading = 0 Then
                AddNum = -AddNum
            End If
            FrmMain.Cls
            FrmMain.CurrentX = FrmMain.ScaleWidth / 2 - 15
            FrmMain.CurrentY = FrmMain.ScaleHeight / 2
            FrmMain.Print "Processing Image " & String(Loading, ".")
            DoEvents
            For IntY = 1 To FrmMain.ScaleHeight
                Temp = GetPixel(PicMatrix.hDC, Int(PicMatrix.ScaleWidth / FrmMain.ScaleWidth * IntX), Int(PicMatrix.ScaleHeight / FrmMain.ScaleHeight * IntY))
                If Temp = RGB(0, 255, 0) Then
                    IntBackGroundPic(IntX, IntY + 4) = 0
                Else
                    GetRgb Temp, R1, G1, B1
                    Temp = Int((R1 + G1 + B1) / 3)
                    IntBackGroundPic(IntX, IntY + 4) = (Temp + 1) / 100 * 80 + 20
                End If
            Next
        Next
        FrmMain.Cls
        PicMatrix.Picture = LoadPicture("") 'Free up memory
    End If
    Font = "Matrix"   'Use the Matrix Font
    Font.Size = IntFontSize
    
    For DoRand = 1 To IntDropCols 'Create the starting IntDrops
        XR = Int(Rnd * FrmMain.ScaleWidth) + 1  'The IntX position
        YR = Int(Rnd * (FrmMain.ScaleHeight + 5)) + 1   'The IntY position
        IntLengthOfDrop(XR, YR) = Int(Rnd * IntMaxLength)      'The Length of the drop
        IntLeading(XR, YR) = 1 'Make it a IntLeading symbol
        IntLetter(XR, YR) = Int(Rnd * 43) + 65         'Set the letter/symbol
        If IntMultipleColours = 1 Then 'If multiple IntColours are enabled
            If IntWillFade = 1 Then
                IntColour(2, XR, YR) = 255
            End If
            If BlnUseBackGround = True Then
                IntColour(1, XR, YR) = IntBackGroundPic(XR, YR) + Rnd * 40 'Set the IntColour
            ElseIf BlnMask = True Then
                If IntBackGroundPic(XR, YR) = 1 Then
                    IntColour(1, XR, YR) = Rnd * 100 + 100
                Else
                    If Rnd * 20 < 1 Then
                        IntBackGroundPic(XR, YR) = 2
                        IntColour(1, XR, YR) = Rnd * 100 + 50
                    Else
                        IntColour(1, XR, YR) = Rnd * 100 + 100
                    End If
                End If
            ElseIf BlnHybrid = True Then
                If IntBackGroundPic(XR, YR) > 2 Then
                    IntColour(1, XR, YR) = IntBackGroundPic(XR, YR) + Rnd * 40
                Else
                    If Rnd * 20 < 1 Then
                        IntBackGroundPic(XR, YR) = 2
                        IntColour(1, XR, YR) = Rnd * 100 + 50
                    Else
                        IntColour(1, XR, YR) = Rnd * 100 + 100
                    End If
                End If
            Else
                IntColour(1, XR, YR) = Rnd * 100 + 100
            End If
        End If
        If IntDifferentFontSizes = 1 Then
            If Rnd * 10 < 4 Then
                IntFntSize(XR, YR) = 2 + Rnd * (IntFontSize - 2)
                IntColour(1, XR, YR) = IntColour(1, XR, YR) / IntFontSize * IntFntSize(XR, YR)
            Else
                IntFntSize(XR, YR) = IntFontSize
            End If
        End If
    Next
    TmrFrameRate.Enabled = True
    If IntSuperSpeed = 0 Then
        TmrMain.Enabled = True
        Exit Sub
    End If
    Do
        DoEvents
        FrmMain.WindowState = 2
        If IntMultipleColours = 0 Then
            Call OneIntColour
        Else
            Call MoreThanOneColour
        End If
        IntFrames = IntFrames + 1
    Loop
End Sub


Sub OneIntColour() 'The routine for drawing one IntColour
    Dim IntDrops As Integer
    For IntX = 1 To FrmMain.ScaleWidth
        For IntY = 1 To FrmMain.ScaleHeight + 5
            If IntLetter(IntX, IntY) <> 0 Then
                
                If IntLeading(IntX, IntY) = 1 Then 'Is it Leading
                    
                    If IntY <= FrmMain.ScaleHeight + 4 Then 'Is it smaller than the screen height                    If IntLengthOfDrop(IntX,IntY) > 0 Then 'Is there still IntDrops in this column
                        
                        If IntLengthOfDrop(IntX, IntY) > 0 Then 'Is there still IntDrops in this column
                            If IntLetter(IntX, IntY + 1) > 0 Then
                                Call Clear(IntX, IntY + 1)
                            End If
                            IntLengthOfDrop(IntX, IntY + 1) = IntLengthOfDrop(IntX, IntY) - 1
                            IntLeading(IntX, IntY + 1) = 2
                            IntLetter(IntX, IntY + 1) = Int(Rnd * 43) + 65
                            IntLeading(IntX, IntY) = 0
                            IntLngWaitLngBeforeClear(IntX, IntY) = IntMaxLngWait
                            FrmMain.CurrentX = IntX
                            FrmMain.CurrentY = IntY - 4
                            FrmMain.ForeColor = LngOneCol
                            FrmMain.Print Chr(IntLetter(IntX, IntY))
                            Call ShowHigh(IntX, IntY + 1)
                        Else    'End of Drop(Kill Letter/Symbol)
                            If IntLeading(IntX, IntY) = 1 Then
                                Call Clear(IntX, IntY)
                            End If
                        End If
                        
                    Else
                        Call Clear(IntX, IntY)
                    End If
                    
                End If
                
                If IntLeading(IntX, IntY) = 1 Or IntLeading(IntX, IntY) = 2 Then
                    IntLeading(IntX, IntY) = 1
                    IntDrops = IntDrops + 1
                End If
                
                If IntLngWaitLngBeforeClear(IntX, IntY) > 0 Then 'Is the Letter/Symbol dieing?
                    IntLngWaitLngBeforeClear(IntX, IntY) = IntLngWaitLngBeforeClear(IntX, IntY) - 1
                    If IntLngWaitLngBeforeClear(IntX, IntY) = 0 Then
                        Call Clear(IntX, IntY)
                    End If
                End If
                
            End If
        Next
    Next
    If IntDrops < IntDropCols Then
        Dim IntMakeNew As Integer
        For IntMakeNew = IntDrops To IntDropCols
            IntX = Int(Rnd * FrmMain.ScaleWidth) + 1
            If IntFromTop = 1 Then
                IntY = Int(Rnd * 5) + 1
            Else
                IntY = Int(Rnd * FrmMain.ScaleHeight) + 1
            End If
            If IntLetter(IntX, IntY) > 0 Then
                Call Clear(IntX, IntY)
            End If
            IntLengthOfDrop(IntX, IntY) = Int(Rnd * IntMaxLength)
            IntLeading(IntX, IntY) = 1
            IntLetter(IntX, IntY) = 64 + Int(Rnd * 26)
            Call ShowHigh(IntX, IntY)
        Next
    End If
End Sub

Sub MoreThanOneColour()
    Dim IntDrops As Integer
    Dim IntMakeNew As Integer
    For IntX = 1 To FrmMain.ScaleWidth
        For IntY = 1 To FrmMain.ScaleHeight + 5
            If IntLetter(IntX, IntY) <> 0 Then
                If IntLeading(IntX, IntY) = 1 Then 'Is it IntLeading
                    If IntY <= FrmMain.ScaleHeight + 4 Then 'Is it smaller than the screen height
                        If IntLengthOfDrop(IntX, IntY) > 0 Then 'Is there still IntDrops in this column
                            If IntLetter(IntX, IntY + 1) <> 0 Then
                                Call Clear(IntX, IntY + 1)
                            End If
                            IntLengthOfDrop(IntX, IntY + 1) = IntLengthOfDrop(IntX, IntY) - 1
                            IntLeading(IntX, IntY) = 0
                            IntLeading(IntX, IntY + 1) = 2
                            IntLetter(IntX, IntY + 1) = Int(Rnd * 43) + 65
                            If BlnUseBackGround = True Then
                                IntColour(1, IntX, IntY + 1) = IntBackGroundPic(IntX, IntY + 1) + Rnd * 40 'Set the IntColour
                            ElseIf BlnMask = True Then
                                If IntBackGroundPic(IntX, IntY) = 2 Then
                                    If IntBackGroundPic(IntX, IntY + 1) = 0 Then
                                        IntBackGroundPic(IntX, IntY + 1) = 2
                                    End If
                                    IntColour(1, IntX, IntY + 1) = Rnd * 100 + 50
                                    Call CheckFade(IntX, IntY + 1)
                                Else
                                    IntColour(1, IntX, IntY + 1) = Rnd * 100 + 100
                                    Call CheckFade(IntX, IntY + 1)
                                End If
                            ElseIf BlnHybrid = True Then
                                If IntBackGroundPic(IntX, IntY) = 2 Then
                                    If IntBackGroundPic(IntX, IntY + 1) = 0 Then
                                        IntBackGroundPic(IntX, IntY + 1) = 2
                                        IntColour(1, IntX, IntY + 1) = Rnd * 100 + 50
                                    Else
                                        IntColour(1, IntX, IntY + 1) = IntBackGroundPic(IntX, IntY + 1) + Rnd * 40 'Set the IntColour
                                    End If
                                    Call CheckFade(IntX, IntY + 1)
                                ElseIf IntBackGroundPic(IntX, IntY + 1) > 2 Then
                                    IntColour(1, IntX, IntY + 1) = IntBackGroundPic(IntX, IntY + 1) + Rnd * 40 'Set the IntColour
                                    Call CheckFade(IntX, IntY + 1)
                                End If
                            Else
                                IntColour(1, IntX, IntY + 1) = Rnd * 100 + 100
                                Call CheckFade(IntX, IntY + 1)
                            End If
                            If IntDifferentFontSizes = 1 Then
                                IntFntSize(IntX, IntY + 1) = IntFontSize - Rnd * (IntFontSize / 2)
                                IntColour(1, IntX, IntY + 1) = (IntColour(1, IntX, IntY + 1) / 2) / IntFontSize * IntFntSize(IntX, IntY + 1)
                            End If
                            IntLngWaitLngBeforeClear(IntX, IntY) = IntMaxLngWait / 2 + Rnd(IntMaxLngWait / 2)
                            Call ShowColor(IntX, IntY)
                            Call ShowHigh(IntX, IntY + 1)
                        Else    'End of Drop(Kill Letter/Symbol)
                            If IntLeading(IntX, IntY) = 1 Then
                                Call Clear(IntX, IntY)
                            End If
                        End If
                    Else
                        Call Clear(IntX, IntY)
                    End If
                ElseIf IntLeading(IntX, IntY) = 2 Then
                    IntLeading(IntX, IntY) = 1
                    IntDrops = IntDrops + 1
                End If
                If IntLngWaitLngBeforeClear(IntX, IntY) > 0 Then 'Is the Letter/Symbol dieing?
                    IntLngWaitLngBeforeClear(IntX, IntY) = IntLngWaitLngBeforeClear(IntX, IntY) - 1
                    If IntLngWaitLngBeforeClear(IntX, IntY) = 0 Or IntColour(1, IntX, IntY) = 0 Then
                        Call Clear(IntX, IntY)
                    ElseIf IntWillFade = 1 Then   'Is fading ativated
                        IntColour(1, IntX, IntY) = IntColour(1, IntX, IntY) - IntFadeSpeed
                        IntColour(2, IntX, IntY) = IntColour(2, IntX, IntY) - IntFadeSpeed * 2
                        If IntColour(1, IntX, IntY) < 0 Then
                            IntColour(1, IntX, IntY) = 0
                        End If
                        If IntColour(2, IntX, IntY) < 0 Then
                            IntColour(2, IntX, IntY) = 0
                        End If
                        If IntColour(1, IntX, IntY) = 0 Then
                            Call Clear(IntX, IntY)
                        ElseIf IntLeading(IntX, IntY) = 0 Then
                            Call ShowColor(IntX, IntY)
                        End If
                    End If
                End If
            End If
        Next
    Next
    If IntDrops < IntDropCols Then
        For IntMakeNew = IntDrops To IntDropCols
            IntX = Int(Rnd * FrmMain.ScaleWidth) + 1
            If IntFromTop = 1 Then
                IntY = Int(Rnd * 5) + 1
            Else
                IntY = Int(Rnd * FrmMain.ScaleHeight) + 1
            End If
            If IntLetter(IntX, IntY) > 0 Then
                Call Clear(IntX, IntY)
            End If
            IntLengthOfDrop(IntX, IntY) = Int(Rnd * IntMaxLength)
            IntLeading(IntX, IntY) = 1
            IntLetter(IntX, IntY) = 64 + Int(Rnd * 26)
            If BlnUseBackGround = True Then
                IntColour(1, IntX, IntY) = IntBackGroundPic(IntX, IntY) + Rnd * 40 'Set the IntColour
            ElseIf BlnMask = True Then
                If IntBackGroundPic(IntX, IntY) = 1 Then
                    IntColour(1, IntX, IntY) = Rnd * 100 + 100
                    Call CheckFade(IntX, IntY)
                Else
                    If Rnd * 20 < 1 Then
                        IntBackGroundPic(IntX, IntY) = 2
                        IntColour(1, IntX, IntY) = Rnd * 100 + 50
                        Call CheckFade(IntX, IntY)
                    Else
                        IntColour(1, IntX, IntY) = Rnd * 100 + 100
                        Call CheckFade(IntX, IntY)
                    End If
                End If
            ElseIf BlnHybrid = True Then
                If IntBackGroundPic(IntX, IntY) > 2 Then
                    IntColour(1, IntX, IntY) = IntBackGroundPic(IntX, IntY) + Rnd * 40 'Set the IntColour
                    Call CheckFade(IntX, IntY)
                Else
                    If Rnd * 20 < 1 Then
                        IntBackGroundPic(IntX, IntY) = 2
                        IntColour(1, IntX, IntY) = Rnd * 100 + 50
                        Call CheckFade(IntX, IntY)
                    Else
                        IntColour(1, IntX, IntY) = Rnd * 100 + 100
                        Call CheckFade(IntX, IntY)
                    End If
                End If
            Else
                IntColour(1, IntX, IntY) = Rnd * 100 + 100
                Call CheckFade(IntX, IntY)
            End If
            If IntDifferentFontSizes = 1 Then
                If Rnd * 10 < 4 Then
                    IntFntSize(IntX, IntY) = 2 + Rnd * (IntFontSize - 2)
                    IntColour(1, IntX, IntY) = IntColour(1, IntX, IntY) / IntFontSize * IntFntSize(IntX, IntY)
                Else
                    IntFntSize(IntX, IntY) = IntFontSize
                End If
            End If
            Call ShowHigh(IntX, IntY)
        Next
    End If
End Sub

Sub CheckFade(ByVal IntX As Integer, ByVal IntY As Integer)
    If IntWillFade = 1 Then
        IntColour(2, IntX, IntY) = 255
    End If
End Sub

Sub Clear(IntX, IntY) 'Clears a letter by redrawing it as black
    If IntDifferentFontSizes = 1 Then
        FrmMain.FontSize = IntFntSize(IntX, IntY)
    End If
    If BlnMask = True Or BlnHybrid = True Then
        If IntBackGroundPic(IntX, IntY) <> 0 Then
            FrmMain.ForeColor = vbBlack
            FrmMain.CurrentX = IntX
            FrmMain.CurrentY = IntY - 4
            FrmMain.Print Chr(IntLetter(IntX, IntY))
            IntLeading(IntX, IntY) = 0
            IntLngWaitLngBeforeClear(IntX, IntY) = 0
            IntLetter(IntX, IntY) = 0
            IntColour(1, IntX, IntY) = 0
            If IntBackGroundPic(IntX, IntY) = 2 Then
                IntBackGroundPic(IntX, IntY) = 0
            End If
        End If
    Else
        FrmMain.ForeColor = vbBlack
        FrmMain.CurrentX = IntX
        FrmMain.CurrentY = IntY - 4
        FrmMain.Print Chr(IntLetter(IntX, IntY))
        IntLeading(IntX, IntY) = 0
        IntLngWaitLngBeforeClear(IntX, IntY) = 0
        IntLetter(IntX, IntY) = 0
        IntColour(1, IntX, IntY) = 0
    End If
End Sub

Sub ShowHigh(IntX, IntY) 'Shows a highlighted letter
    If IntDifferentFontSizes = 1 Then
        FrmMain.FontSize = IntFntSize(IntX, IntY)
    End If
    If BlnMask = True Or BlnHybrid = True Then
        If IntBackGroundPic(IntX, IntY) <> 0 Then
            FrmMain.ForeColor = vbWhite
            FrmMain.CurrentX = IntX
            FrmMain.CurrentY = IntY - 4
            FrmMain.Print Chr(IntLetter(IntX, IntY))
        End If
    Else
        FrmMain.ForeColor = vbWhite
        FrmMain.CurrentX = IntX
        FrmMain.CurrentY = IntY - 4
        FrmMain.Print Chr(IntLetter(IntX, IntY))
    End If
End Sub

Sub ShowColor(IntX, IntY) 'Shows a Coloured letter
    If IntDifferentFontSizes = 1 Then
        FrmMain.FontSize = IntFntSize(IntX, IntY)
    End If
    If BlnMask = True Or BlnHybrid = True Then
        If IntBackGroundPic(IntX, IntY) <> 0 Then
            If IntReloadStyle = 0 Then
                If IntCodeColour = 0 Then
                    FrmMain.ForeColor = RGB(IntColour(1, IntX, IntY), 0, 0)
                ElseIf IntCodeColour = 1 Then
                    FrmMain.ForeColor = RGB(0, IntColour(1, IntX, IntY), 0)
                ElseIf IntCodeColour = 2 Then
                    FrmMain.ForeColor = RGB(0, 0, IntColour(1, IntX, IntY))
                End If
            Else
                If IntCodeColour = 0 Then
                    FrmMain.ForeColor = RGB(IntColour(2, IntX, IntY) + IntColour(1, IntX, IntY), IntColour(2, IntX, IntY), IntColour(2, IntX, IntY))
                ElseIf IntCodeColour = 1 Then
                    FrmMain.ForeColor = RGB(IntColour(2, IntX, IntY), IntColour(2, IntX, IntY) + IntColour(1, IntX, IntY), IntColour(2, IntX, IntY))
                ElseIf IntCodeColour = 2 Then
                    FrmMain.ForeColor = RGB(IntColour(2, IntX, IntY), IntColour(2, IntX, IntY), IntColour(2, IntX, IntY) + IntColour(1, IntX, IntY))
                End If
            End If
            FrmMain.CurrentX = IntX
            FrmMain.CurrentY = IntY - 4
            FrmMain.Print Chr(IntLetter(IntX, IntY))
        End If
    Else
        If IntReloadStyle = 0 Then
            If IntCodeColour = 0 Then
                FrmMain.ForeColor = RGB(IntColour(1, IntX, IntY), 0, 0)
            ElseIf IntCodeColour = 1 Then
                FrmMain.ForeColor = RGB(0, IntColour(1, IntX, IntY), 0)
            ElseIf IntCodeColour = 2 Then
                FrmMain.ForeColor = RGB(0, 0, IntColour(1, IntX, IntY))
            End If
        Else
            If IntCodeColour = 0 Then
                FrmMain.ForeColor = RGB(IntColour(2, IntX, IntY) + IntColour(1, IntX, IntY), IntColour(2, IntX, IntY), IntColour(2, IntX, IntY))
            ElseIf IntCodeColour = 1 Then
                FrmMain.ForeColor = RGB(IntColour(2, IntX, IntY), IntColour(2, IntX, IntY) + IntColour(1, IntX, IntY), IntColour(2, IntX, IntY))
            ElseIf IntCodeColour = 2 Then
                FrmMain.ForeColor = RGB(IntColour(2, IntX, IntY), IntColour(2, IntX, IntY), IntColour(2, IntX, IntY) + IntColour(1, IntX, IntY))
            End If
        End If
        FrmMain.CurrentX = IntX
        FrmMain.CurrentY = IntY - 4
        FrmMain.Print Chr(IntLetter(IntX, IntY))
    End If
End Sub

Sub ShowBlack(IntX, IntY) 'Shows a IntColoured letter
    If IntDifferentFontSizes = 1 Then
        FrmMain.FontSize = IntFntSize(IntX, IntY)
    End If
    If BlnMask = True Or BlnHybrid = True Then
        If IntBackGroundPic(IntX, IntY) <> 0 Then
            FrmMain.ForeColor = vbBlack
            FrmMain.CurrentX = IntX
            FrmMain.CurrentY = IntY - 4
            FrmMain.Print Chr(IntLetter(IntX, IntY))
        End If
    Else
        FrmMain.ForeColor = vbBlack
        FrmMain.CurrentX = IntX
        FrmMain.CurrentY = IntY - 4
        FrmMain.Print Chr(IntLetter(IntX, IntY))
    End If
End Sub

'#######################################################################
'Form Events
'#######################################################################

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If IntLastXPos = 0 And IntLastYPos = 0 Then
        IntLastXPos = X
        IntLastYPos = Y
    End If
    If Abs(X - IntLastXPos) > 20 Or Abs(Y - IntLastYPos) > 20 Then
        Call ExitProgram
    Else
        IntLastXPos = X
        IntLastYPos = Y
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call ExitProgram
End Sub

Private Sub Form_Terminate()
    Call ExitProgram
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ExitProgram
End Sub

Private Sub Form_Click()
    Call ExitProgram
End Sub

Private Sub Form_DblClick()
    Call ExitProgram
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call ExitProgram
End Sub

Private Sub PicMatrix_Click()
    Call ExitProgram
End Sub

Private Sub PicMatrix_DblClick()
    Call ExitProgram
End Sub

Private Sub PicMatrix_KeyPress(KeyAscii As Integer)
    Call ExitProgram
End Sub

Private Sub TmrFrameRate_Timer()
    If IntFrameRate = 0 Then
        IntFrameRate = IntFrames
    Else
        IntFrameRate = (IntFrameRate + IntFrames) / 2
    End If
    IntFrames = 0
End Sub

'#######################################################################
'Call Tracing Section
'#######################################################################

Private Sub TmrMain1_Timer()
    IntTextDone = IntTextDone + 1
    FrmMain.Cls
    FrmMain.CurrentX = 500
    FrmMain.CurrentY = 500
    IntAnim = 1 - IntAnim
    FrmMain.Print Mid$(StrStartText, 1, IntTextDone);
    FrmMain.ForeColor = RGB(0, 100 + (IntAnim * 150), 0)
    FrmMain.ForeColor = LngTraceCol
    If IntTextDone = Len(StrStartText) Then
        StrStartText = "Trace Program Running"
        IntTextDone = 0
        If IntSTextF = 1 Then
            TmrMain1.Enabled = False
            TmrMain2.Enabled = True
        End If
        IntSTextF = 1
        Call WaitFor(1)
    End If
    Call NewNewStrNumbers
End Sub

Private Sub TmrMain2_Timer()
    Dim IntDoPhone As Integer
    Dim IntNoPhone As Integer
    Dim IntDoClear As Integer
    Dim IntComplete As Integer
    Dim IntCheck As Integer
    Dim IntDoHor As Integer
    Dim IntDoVert As Integer
    Dim BlnExitMe As Boolean
    FrmMain.Cls
    For IntDoPhone = 1 To 11
        If BlnPhoneOn(IntDoPhone) = True Then
            FrmMain.CurrentX = LngXCoord(IntDoPhone)
            FrmMain.CurrentY = LngYCoord
            FrmMain.Print StrPhoneNo(IntDoPhone)
        End If
    Next
    LngWait = LngWait + 1
    If LngWait = 20 Then
        LngWait = 0
        BlnExitMe = False
        Do
            IntNoPhone = Int(Rnd * 11) + 1
            If BlnPhoneOn(IntNoPhone) = False Then
                BlnExitMe = True
                BlnPhoneOn(IntNoPhone) = True
                For IntDoClear = IntNoPhone To 60 Step 10
                    BlnCols(IntDoClear) = False
                Next
            End If
            IntComplete = 0
            For IntCheck = 1 To 11
                If BlnPhoneOn(IntCheck) = True Then
                    IntComplete = IntComplete + 1
                End If
            Next
            If IntComplete = 11 Then
                TmrMain2.Enabled = False
                Call Finish
            End If
        Loop Until BlnExitMe = True
    End If
    For IntDoHor = 1 To 60
        If BlnCols(IntDoHor) = True Then
            For IntDoVert = 30 To 1 Step -1
                FrmMain.CurrentX = IntXNums(IntDoHor)
                FrmMain.CurrentY = IntYNums(IntDoVert)
                FrmMain.ForeColor = RGB(0, 150 + Rnd * 100, 0)
                FrmMain.Print StrNumbers(IntDoHor, IntDoVert)
                StrNumbers(IntDoHor, IntDoVert) = StrNumbers(IntDoHor, IntDoVert - 1)
            Next
        End If
        StrNumbers(IntDoHor, 1) = Int(Rnd * 10)
    Next
    FrmMain.ForeColor = LngTraceCol
End Sub

Sub NewNewStrNumbers()
    Dim IntNewCol As Integer
    Dim IntVerts As Integer
    'Fills the Grid with random StrNumbers
    For IntNewCol = 1 To 60
        BlnCols(IntNewCol) = True
        For IntVerts = 1 To 30
            StrNumbers(IntNewCol, IntVerts) = Int(Rnd * 10)
        Next
    Next
End Sub

Sub Finish()
    Dim IntDoPhone As Integer
    For IntDoPhone = 1 To 11
        If BlnPhoneOn(IntDoPhone) = True Then
            FrmMain.CurrentX = LngXCoord(IntDoPhone)
            FrmMain.CurrentY = LngYCoord
            FrmMain.Print StrPhoneNo(IntDoPhone)
        End If
    Next
    FrmMain.CurrentX = 500
    FrmMain.CurrentY = 500
    FrmMain.Print "Trace Program: Completed "
    Call ClearUp
End Sub

Sub ClearUp()
    Dim IntPNo As Integer
    Call NewNewStrNumbers
    For IntPNo = 1 To 11
        BlnPhoneOn(IntPNo) = False
        If LngRanNum = 1 Then
            StrPhoneNo(IntPNo) = Int(Rnd * 9)
        Else
            StrPhoneNo(IntPNo) = Mid$(StrNumber, IntPNo, 1)
        End If
    Next
    Call WaitFor(30)
    StrStartText = "Call Trans opt: Rec " + Str$(Date) + " " + Str$(Time) + " Rec:Log> "
    IntTextDone = 0
    IntSTextF = 0
    TmrMain1.Enabled = True
End Sub

'#######################################################################
'Knock,Knock Neo... Section
'#######################################################################

Private Sub Tmrmain3_Timer()
    IntMatrixDone = IntMatrixDone + 1
    FrmMain.Cls
    FrmMain.CurrentY = 3
    FrmMain.CurrentX = 6
    Print Mid$(StrTxtMatrix(IntCurrentTxt), 1, IntMatrixDone);
    If IntMatrixDone = Len(StrTxtMatrix(IntCurrentTxt)) Then
        IntMatrixDone = 0
        IntCurrentTxt = IntCurrentTxt + 1
        If IntCurrentTxt = 5 Then
            TmrMain3.Enabled = False
            Call Doneall
        End If
        Call WaitFor(5)
        TmrMain3.Interval = IntTxtSpeed(IntCurrentTxt)
    End If
End Sub

Sub Doneall()
    TmrMain3.Enabled = False
    Call WaitFor(30)
    TmrMain3.Enabled = True
    IntCurrentTxt = 1
    IntMatrixDone = 0
    TmrMain3.Interval = IntTxtSpeed(IntCurrentTxt)
End Sub

Sub WaitFor(Interval)
    Dim LngBefore As Long
    LngBefore = Timer
    Do
        DoEvents
    Loop Until Timer - LngBefore > Interval
End Sub

Sub ExitProgram()
    ShowCursor (1)  'Make the cursor visible
    i = mciSendString("close video1", 0&, 0, 0)
    If IntActiveScreensaver = 0 Then
        SaveSetting "Kevin Pfister's Matrix", "Speed", "FrameRate", IntFrameRate
    End If
    End
End Sub

'#######################################################################
'Room
'#######################################################################

Private Sub TmrHallway_Timer()
    Dim DoRand As Integer
    Dim XR As Integer
    Dim YR As Integer
    Dim Temp As Long
    Dim Loading As Long
    Dim AddNum As Integer
    
    TmrHallway.Enabled = False
    
    If IntCodeColour = 0 Then
        LngOneCol = RGB(150, 0, 0)
    ElseIf IntCodeColour = 1 Then
        LngOneCol = RGB(0, 150, 0)
    ElseIf IntCodeColour = 2 Then
        LngOneCol = RGB(0, 0, 150)
    End If
    'Change the variable sizes to fit the screen size
    ReDim IntLengthOfDrop(FrmMain.ScaleWidth, FrmMain.ScaleHeight + 5) As Integer
    ReDim IntLeading(FrmMain.ScaleWidth, FrmMain.ScaleHeight + 5) As Integer
    ReDim IntLetter(FrmMain.ScaleWidth, FrmMain.ScaleHeight + 5) As Integer
    ReDim IntColour(1, FrmMain.ScaleWidth, FrmMain.ScaleHeight + 5) As Integer
    ReDim IntLngWaitLngBeforeClear(FrmMain.ScaleWidth, FrmMain.ScaleHeight + 5) As Integer
    ReDim IntBackGroundPic(FrmMain.ScaleWidth, FrmMain.ScaleHeight + 5) As Integer
    Font = "Matrix"   'Use the Matrix Font
    Font.Size = 12
    
    For DoRand = 1 To 50 'Create the starting IntDrops
        XR = Int(Rnd * FrmMain.ScaleWidth) + 1  'The IntX position
        YR = Int(Rnd * (FrmMain.ScaleHeight + 5)) + 1   'The IntY position
        IntLengthOfDrop(XR, YR) = Int(Rnd * IntMaxLength)      'The Length of the drop
        IntLeading(XR, YR) = 1 'Make it a IntLeading symbol
        IntLetter(XR, YR) = Int(Rnd * 43) + 65         'Set the letter/symbol
        IntColour(1, XR, YR) = Rnd * 100 + 150
    Next
    For IntX = 0 To 70
        For IntY = 0 To 80
            Loading = GetPixel(PicMatrix.hDC, IntX, IntY)
            If Loading = vbBlack Then
                MatrixPeople(IntX, IntY) = 0
            Else
                MatrixPeople(IntX, IntY) = 1
            End If
        Next
    Next
    PicMatrix.Picture = LoadPicture("")
    PicMatrix.Top = 0
    PicMatrix.Left = 0
    PicMatrix.Width = FrmMain.ScaleWidth
    PicMatrix.Height = FrmMain.ScaleHeight
    PicMatrix.Visible = True
    TempX = PicMatrix.ScaleWidth / 2
    TempY = PicMatrix.ScaleHeight / 2
    Dim A
    For A = 1 To 100
        Call MOColour
    Next
    PicMatrix.AutoRedraw = True
    Do
        DoEvents
        Call MOColour
        PicMatrix.Cls
        PicMatrix.Line (TempX - 255, TempY - 155)-(TempX + 255, TempY - 155), RGB(100, 100, 100)
        PicMatrix.Line (TempX + 255, TempY - 155)-(TempX + 255, TempY + 155), RGB(100, 100, 100)
        PicMatrix.Line (TempX + 255, TempY + 155)-(TempX - 255, TempY + 155), RGB(100, 100, 100)
        PicMatrix.Line (TempX - 255, TempY + 155)-(TempX - 255, TempY - 155), RGB(100, 100, 100)
        Call Hallway
        PicMatrix.Refresh
    Loop
End Sub

Sub Hallway()
    
    For IntX = 0 To 200 Step 2
        If IntX < 50 Then
            IntY = 0
        Else
            IntY = IntX - 50
        End If
        'Left Side
        StretchBlt PicMatrix.hDC, TempX - 250 + IntX, TempY - 150 + IntY / 2, 2, 300 - IntY, FrmMain.hDC, PicMatrix.ScaleWidth / 300 * IntX, 0, 4, TempY, vbSrcCopy
        'Right Side
        StretchBlt PicMatrix.hDC, TempX + 250 - IntX, TempY - 150 + IntY / 2, 2, 300 - IntY, FrmMain.hDC, PicMatrix.ScaleWidth / 300 * IntX, 0, 4, TempY, vbSrcCopy
    Next
    For IntX = 0 To 100 Step 2
        'Top
        StretchBlt PicMatrix.hDC, TempX - (200 - 2 * IntX), TempY - 150 + IntX, 400 - (4 * IntX), 2, FrmMain.hDC, 0, (PicMatrix.ScaleHeight / 200) * (150 - IntX), TempX, 4, vbSrcCopy
        'Bottom
        StretchBlt PicMatrix.hDC, TempX - (200 - 2 * IntX), TempY + 150 - IntX, 400 - (4 * IntX), 2, FrmMain.hDC, 0, (PicMatrix.ScaleHeight / 200) * (150 - IntX), TempX, 4, vbSrcCopy
    Next
    StretchBlt PicMatrix.hDC, TempX - 50, TempY - 75, 100, 150, FrmMain.hDC, 0, 0, PicMatrix.ScaleWidth / 4, PicMatrix.ScaleHeight / 3, vbSrcCopy
    For IntX = 0 To 65
        For IntY = 1 To 75
            If MatrixPeople(IntX, IntY) = 1 Then
                SetPixelV PicMatrix.hDC, TempX - 35 + IntX, TempY + IntY - 5, RGB(0, Rnd * 100 + 50, 0)
            End If
        Next
    Next
End Sub

Sub MOColour()
    Dim IntDrops As Integer
    Dim IntMakeNew As Integer
    For IntX = 1 To FrmMain.ScaleWidth
        For IntY = 1 To FrmMain.ScaleHeight + 5
            If IntLetter(IntX, IntY) <> 0 Then
                If IntLeading(IntX, IntY) = 1 Then 'Is it IntLeading
                    If IntY <= FrmMain.ScaleHeight + 4 Then 'Is it smaller than the screen height
                        If IntLengthOfDrop(IntX, IntY) > 0 Then 'Is there still IntDrops in this column
                            If IntLetter(IntX, IntY + 1) <> 0 Then
                                Call Clear1(IntX, IntY + 1)
                            End If
                            IntLengthOfDrop(IntX, IntY + 1) = IntLengthOfDrop(IntX, IntY) - 1
                            IntLeading(IntX, IntY) = 0
                            IntLeading(IntX, IntY + 1) = 2
                            IntLetter(IntX, IntY + 1) = Int(Rnd * 43) + 65
                            IntColour(1, IntX, IntY + 1) = Rnd * 100 + 150
                            IntLngWaitLngBeforeClear(IntX, IntY) = IntMaxLngWait
                            Call ShowColor1(IntX, IntY)
                            Call ShowHigh1(IntX, IntY + 1)
                        Else    'End of Drop(Kill Letter/Symbol)
                            If IntLeading(IntX, IntY) = 1 Then
                                Call Clear1(IntX, IntY)
                            End If
                        End If
                    Else
                        Call Clear1(IntX, IntY)
                    End If
                ElseIf IntLeading(IntX, IntY) = 2 Then
                    IntLeading(IntX, IntY) = 1
                    IntDrops = IntDrops + 1
                End If
                If IntLngWaitLngBeforeClear(IntX, IntY) > 0 Then 'Is the Letter/Symbol dieing?
                    IntLngWaitLngBeforeClear(IntX, IntY) = IntLngWaitLngBeforeClear(IntX, IntY) - 1
                    If IntLngWaitLngBeforeClear(IntX, IntY) = 0 Or IntColour(1, IntX, IntY) = 0 Then
                        Call Clear1(IntX, IntY)
                    End If
                End If
            End If
        Next
    Next
    If IntDrops < 50 Then
        For IntMakeNew = IntDrops To 50
            IntX = Int(Rnd * FrmMain.ScaleWidth) + 1
            If IntFromTop = 1 Then
                IntY = Int(Rnd * 5) + 1
            Else
                IntY = Int(Rnd * FrmMain.ScaleHeight) + 1
            End If
            If IntLetter(IntX, IntY) > 0 Then
                Call Clear1(IntX, IntY)
            End If
            IntLengthOfDrop(IntX, IntY) = Int(Rnd * 50)
            IntLeading(IntX, IntY) = 1
            IntLetter(IntX, IntY) = 64 + Int(Rnd * 26)
            IntColour(1, IntX, IntY) = Rnd * 100 + 150
            Call ShowHigh1(IntX, IntY)
        Next
    End If
End Sub

Sub Clear1(IntX, IntY) 'Clears a letter by redrawing it as black
    FrmMain.ForeColor = vbBlack
    FrmMain.CurrentX = IntX
    FrmMain.CurrentY = IntY - 4
    FrmMain.Print Chr(IntLetter(IntX, IntY))
    IntLeading(IntX, IntY) = 0
    IntLngWaitLngBeforeClear(IntX, IntY) = 0
    IntLetter(IntX, IntY) = 0
    IntColour(1, IntX, IntY) = 0
End Sub

Sub ShowHigh1(IntX, IntY) 'Shows a highlighted letter
    FrmMain.ForeColor = vbWhite
    FrmMain.CurrentX = IntX
    FrmMain.CurrentY = IntY - 4
    FrmMain.Print Chr(IntLetter(IntX, IntY))
End Sub

Sub ShowColor1(IntX, IntY) 'Shows a Coloured letter
    FrmMain.ForeColor = RGB(0, IntColour(1, IntX, IntY), 0)
    FrmMain.CurrentX = IntX
    FrmMain.CurrentY = IntY - 4
    FrmMain.Print Chr(IntLetter(IntX, IntY))
End Sub

Sub ShowBlack1(IntX, IntY) 'Shows a IntColoured letter
    FrmMain.ForeColor = vbBlack
    FrmMain.CurrentX = IntX
    FrmMain.CurrentY = IntY - 4
    FrmMain.Print Chr(IntLetter(IntX, IntY))
End Sub
