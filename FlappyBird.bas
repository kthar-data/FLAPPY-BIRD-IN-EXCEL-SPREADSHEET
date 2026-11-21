Attribute VB_Name = "FlappyBird"
Option Explicit

' =====================================================================
'  FLAPPY BIRD FOR EXCEL  -  v3, runs in its own fast window
'  (Windows desktop Excel 2010 or newer, 32 or 64 bit)
'
'  HOW TO RUN
'    1. Alt+F8 -> SetupFlappyBird -> Run     (builds a launcher sheet)
'    2. Click PLAY on that sheet             (or Alt+F8 -> PlayFlappyBird)
'    3. SPACE / UP ARROW / mouse click = flap
'       ESC or the window's X = quit
'
'  WHY THIS VERSION IS FAST
'    The game never touches the worksheet while you play. It draws into
'    an off-screen bitmap with GDI and copies that bitmap to a small
'    window of its own, so Excel's slow cell and shape repainting never
'    runs. Key presses are read straight from the message queue, so even
'    a very short tap is registered whatever the frame rate is.
' =====================================================================

' ---- Windows structures --------------------------------------------
Private Type GRect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type WinMsg
    hWnd As LongPtr
    Id As Long
    wP As LongPtr
    lP As LongPtr
    tm As Long
    ptX As Long
    ptY As Long
End Type

' ---- Windows API ---------------------------------------------------
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef lpCount As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFreq As Currency) As Long
Private Declare PtrSafe Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As LongPtr
Private Declare PtrSafe Function timeBeginPeriod Lib "winmm" (ByVal uPeriod As Long) As Long
Private Declare PtrSafe Function timeEndPeriod Lib "winmm" (ByVal uPeriod As Long) As Long

Private Declare PtrSafe Function CreateWindowExA Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, ByVal lpParam As LongPtr) As LongPtr
Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function SetFocus Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function AdjustWindowRect Lib "user32" (ByRef lpRect As GRect, ByVal dwStyle As Long, ByVal bMenu As Long) As Long
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function ValidateRect Lib "user32" (ByVal hWnd As LongPtr, ByVal lpRect As LongPtr) As Long
Private Declare PtrSafe Function FillRect Lib "user32" (ByVal hDC As LongPtr, ByRef lpRect As GRect, ByVal hBrush As LongPtr) As Long
Private Declare PtrSafe Function PeekMessageA Lib "user32" (ByRef lpMsg As WinMsg, ByVal hWnd As LongPtr, ByVal wMin As Long, ByVal wMax As Long, ByVal wRemove As Long) As Long
Private Declare PtrSafe Function TranslateMessage Lib "user32" (ByRef lpMsg As WinMsg) As Long
Private Declare PtrSafe Function DispatchMessageA Lib "user32" (ByRef lpMsg As WinMsg) As LongPtr

Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As LongPtr
Private Declare PtrSafe Function BitBlt Lib "gdi32" (ByVal hDestDC As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare PtrSafe Function StretchBlt Lib "gdi32" (ByVal hDestDC As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare PtrSafe Function SetStretchBltMode Lib "gdi32" (ByVal hDC As LongPtr, ByVal iMode As Long) As Long
Private Declare PtrSafe Function SetBkMode Lib "gdi32" (ByVal hDC As LongPtr, ByVal iBkMode As Long) As Long
Private Declare PtrSafe Function SetTextColor Lib "gdi32" (ByVal hDC As LongPtr, ByVal crColor As Long) As Long
Private Declare PtrSafe Function SetTextAlign Lib "gdi32" (ByVal hDC As LongPtr, ByVal wFlags As Long) As Long
Private Declare PtrSafe Function TextOutA Lib "gdi32" (ByVal hDC As LongPtr, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare PtrSafe Function CreateFontA Lib "gdi32" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As LongPtr

' ---- window / message constants ------------------------------------
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_SYSMENU As Long = &H80000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_VISIBLE As Long = &H10000000
Private Const SS_BLACKRECT As Long = &H4          ' plain rectangle, no text drawn
Private Const SS_NOTIFY As Long = &H100           ' lets the window receive mouse clicks
Private Const PM_REMOVE As Long = 1
Private Const WM_QUIT As Long = &H12
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const WM_CHAR As Long = &H102
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const VK_ESCAPE As Long = &H1B
Private Const VK_SPACE As Long = &H20
Private Const VK_UP As Long = &H26
Private Const SRCCOPY As Long = &HCC0020

' ---- screen / layout -----------------------------------------------
Private Const SCR_W As Long = 360                 ' game picture size, pixels
Private Const SCR_H As Long = 540
Private Const WIN_SCALE As Long = 1               ' 2 = double-size window (4K screens)
Private Const GROUND_Y As Long = 470              ' where the ground starts
Private Const HALF_W As Long = SCR_W \ 2

' ---- difficulty / feel  (pixels and seconds) -----------------------
Private Const GRAV As Double = 1500#              ' downward acceleration
Private Const FLAP_V As Double = -430#            ' upward speed right after a flap
Private Const MAX_FALL As Double = 620#           ' terminal falling speed
Private Const PIPE_SPEED As Double = 140#         ' how fast the world scrolls
Private Const PIPE_GAP As Long = 165              ' height of the opening
Private Const PIPE_SPACING As Double = 225#       ' distance between pipes
Private Const PIPE_W As Long = 60
Private Const PIPE_MARGIN As Long = 70            ' gap stays this far from top / ground
Private Const MAX_SHIFT As Long = 100             ' biggest gap move between two pipes
Private Const CAP_H As Long = 26
Private Const CAP_OVER As Long = 4                ' cap sticks out this far each side
Private Const MAXP As Long = 5

Private Const BIRD_X As Long = 90
Private Const BIRD_W As Long = 34
Private Const BIRD_H As Long = 24
Private Const BIRD_BASE_Y As Long = 215

Private Const PHYS_DT As Double = 1# / 120#       ' fixed physics step
Private Const FRAME_T As Double = 1# / 60#        ' target frame time

' ---- game states ---------------------------------------------------
Private Const ST_READY As Long = 0
Private Const ST_PLAY As Long = 1
Private Const ST_DEAD As Long = 2

' ---- brush indexes and fonts ---------------------------------------
Private Const B_SKY As Long = 0
Private Const B_CLOUD As Long = 1
Private Const B_PIPE As Long = 2
Private Const B_PLIGHT As Long = 3
Private Const B_PSHADE As Long = 4
Private Const B_DARK As Long = 5
Private Const B_GRASS As Long = 6
Private Const B_GRASS2 As Long = 7
Private Const B_SAND As Long = 8
Private Const B_SAND2 As Long = 9
Private Const B_BIRD As Long = 10
Private Const B_BEAK As Long = 11
Private Const B_WING As Long = 12
Private Const B_WHITE As Long = 13
Private Const NB As Long = 14

Private Const F_BIG As Long = 0
Private Const F_MID As Long = 1
Private Const F_SMALL As Long = 2

' ---- launcher sheet ------------------------------------------------
Private Const SHEET_NAME As String = "Flappy Bird"
Private Const CELL_BEST As String = "C6"

' ---- state ---------------------------------------------------------
Private mHwnd As LongPtr
Private mDC As LongPtr
Private mBmp As LongPtr
Private mOldBmp As LongPtr
Private mOldFont As LongPtr
Private hB(0 To NB - 1) As LongPtr
Private hF(0 To 2) As LongPtr
Private mFreq As Currency
Private mQuit As Boolean

Private mState As Long
Private mBirdY As Double
Private mVel As Double
Private mClock As Double
Private mDeadT As Double
Private mGroundX As Double
Private mCloudX As Double
Private mScore As Long
Private mBest As Long
Private mLastGY As Long

Private pX(1 To MAXP) As Double
Private pGY(1 To MAXP) As Long
Private pOn(1 To MAXP) As Boolean
Private pPass(1 To MAXP) As Boolean


' =====================================================================
'  PUBLIC ENTRY POINTS
' =====================================================================

' Builds a small launcher sheet with a PLAY button.
Public Sub SetupFlappyBird()
    Dim ws As Worksheet
    Dim btn As Object

    Set ws = GetSheet()
    ws.Activate
    Application.ScreenUpdating = False

    Do While ws.Shapes.Count > 0
        ws.Shapes(1).Delete
    Loop
    ws.Cells.Clear
    ws.Cells.ColumnWidth = 9
    ws.Cells.RowHeight = 15
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 26

    With ws.Range("B2")
        .Value = "FLAPPY BIRD"
        .Font.Size = 22
        .Font.Bold = True
        .Font.Color = RGB(84, 148, 36)
    End With
    ws.Range("B3").Value = "Click PLAY - the game opens in its own window."
    ws.Range("B4").Value = "SPACE / UP ARROW / mouse click = flap      ESC = quit"
    ws.Range("B6").Value = "Best score this session:"
    ws.Range("B6").Font.Bold = True
    ws.Range(CELL_BEST).Value = mBest
    ws.Range(CELL_BEST).HorizontalAlignment = xlLeft

    Set btn = ws.Buttons.Add(ws.Range("B8").Left, ws.Range("B8").Top, 170, 46)
    btn.Name = "btnPlay"
    btn.Caption = "PLAY"
    btn.Font.Bold = True
    btn.Font.Size = 18
    btn.OnAction = "PlayFlappyBird"

    ActiveWindow.DisplayGridlines = False
    ws.Range("A1").Select
    Application.ScreenUpdating = True
End Sub


Public Sub PlayFlappyBird()
    Dim ws As Worksheet
    Dim tFrame As Double, tPrev As Double, dt As Double, acc As Double
    Dim errNum As Long, errMsg As String
    Dim timerOn As Boolean

    ' Already running (window still open)? Do nothing.
    If mHwnd <> 0 Then
        If IsWindow(mHwnd) <> 0 Then Exit Sub
    End If

    On Error GoTo Fail
    ' Esc / Ctrl+Break become error 18, handled quietly below.
    Application.EnableCancelKey = xlErrorHandler

    Randomize
    mQuit = False
    QueryPerformanceFrequency mFreq
    timerOn = (timeBeginPeriod(1) = 0)      ' 1 ms sleep accuracy

    If Not OpenGameWindow() Then
        Err.Raise vbObjectError + 1, , "Could not create the game window."
    End If
    InitGdi
    ResetRun

    tPrev = NowSec()
    Do While Not mQuit
        If IsWindow(mHwnd) = 0 Then Exit Do

        tFrame = NowSec()
        dt = tFrame - tPrev
        tPrev = tFrame
        If dt > 0.1 Then dt = 0.1
        If dt < 0# Then dt = 0#

        PumpMessages

        acc = acc + dt
        Do While acc >= PHYS_DT
            StepGame PHYS_DT
            acc = acc - PHYS_DT
        Loop

        DrawScene
        PresentFrame

        ' Wait out the rest of the frame, still reading the keyboard.
        Do
            PumpMessages
            If mQuit Then Exit Do
            If IsWindow(mHwnd) = 0 Then Exit Do
            If NowSec() - tFrame >= FRAME_T Then Exit Do
            Sleep 1
        Loop
    Loop

Fail:
    errNum = Err.Number
    errMsg = Err.Description
    On Error Resume Next
    Application.EnableCancelKey = xlInterrupt
    CloseGame
    If timerOn Then timeEndPeriod 1

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    If Not ws Is Nothing Then ws.Range(CELL_BEST).Value = mBest

    If errNum <> 0 And errNum <> 18 Then
        MsgBox "Flappy Bird stopped: " & errMsg, vbExclamation
    End If
End Sub


' =====================================================================
'  INPUT
' =====================================================================

' Reads every waiting message. Key presses are taken from the queue
' (so a quick tap is never missed) and are not passed on to Excel.
Private Sub PumpMessages()
    Dim m As WinMsg

    Do While PeekMessageA(m, 0, 0, 0, PM_REMOVE) <> 0
        Select Case m.Id
            Case WM_QUIT
                mQuit = True
            Case WM_KEYDOWN
                If m.wP = VK_ESCAPE Then
                    mQuit = True
                ElseIf m.wP = VK_SPACE Or m.wP = VK_UP Then
                    ' bit 30 set = key auto-repeat; only count the first press
                    If (m.lP And &H40000000) = 0 Then Flap
                End If
            Case WM_KEYUP, WM_CHAR
                ' swallowed so Excel never types into a cell
            Case WM_LBUTTONDOWN
                If m.hWnd = mHwnd Then Flap
                DispatchMessageA m
            Case Else
                TranslateMessage m
                DispatchMessageA m
        End Select
    Loop
End Sub


Private Sub Flap()
    Select Case mState
        Case ST_READY
            StartRun
            mVel = FLAP_V
        Case ST_PLAY
            mVel = FLAP_V
        Case ST_DEAD
            If mDeadT > 0.6 Then ResetRun
    End Select
End Sub


' =====================================================================
'  GAME LOGIC
' =====================================================================

Private Sub ResetRun()
    Dim i As Long
    For i = 1 To MAXP
        pOn(i) = False
    Next i
    mState = ST_READY
    mScore = 0
    mBirdY = BIRD_BASE_Y
    mVel = 0#
    mDeadT = 0#
End Sub


Private Sub StartRun()
    Dim i As Long
    For i = 1 To MAXP
        pOn(i) = False
    Next i
    mScore = 0
    mLastGY = (GROUND_Y - PIPE_GAP) \ 2
    mState = ST_PLAY
    SpawnPipe
End Sub


Private Sub Die()
    mState = ST_DEAD
    mDeadT = 0#
    If mVel < 0# Then mVel = 0#
    If mScore > mBest Then mBest = mScore
End Sub


Private Sub StepGame(ByVal dt As Double)
    mClock = mClock + dt

    Select Case mState
        Case ST_READY
            ScrollWorld dt
            mBirdY = BIRD_BASE_Y + 6# * Sin(mClock * 5#)

        Case ST_PLAY
            ScrollWorld dt
            mVel = mVel + GRAV * dt
            If mVel > MAX_FALL Then mVel = MAX_FALL
            mBirdY = mBirdY + mVel * dt
            If mBirdY < 0# Then
                mBirdY = 0#
                mVel = 0#
            End If
            MovePipes dt
            CheckCollisions

        Case ST_DEAD
            mDeadT = mDeadT + dt
            mVel = mVel + GRAV * dt
            If mVel > MAX_FALL Then mVel = MAX_FALL
            mBirdY = mBirdY + mVel * dt
            If mBirdY + BIRD_H > GROUND_Y Then mBirdY = GROUND_Y - BIRD_H
    End Select
End Sub


Private Sub ScrollWorld(ByVal dt As Double)
    mGroundX = mGroundX + PIPE_SPEED * dt
    mGroundX = mGroundX - Int(mGroundX / 32#) * 32#
    mCloudX = mCloudX + 30# * dt
End Sub


Private Sub SpawnPipe()
    Dim i As Long
    Dim lo As Long, hi As Long, gy As Long

    lo = PIPE_MARGIN
    hi = GROUND_Y - PIPE_GAP - PIPE_MARGIN
    gy = mLastGY + CLng((Rnd * 2# - 1#) * MAX_SHIFT)
    If gy < lo Then gy = lo
    If gy > hi Then gy = hi

    For i = 1 To MAXP
        If Not pOn(i) Then
            pOn(i) = True
            pPass(i) = False
            pX(i) = SCR_W + 10
            pGY(i) = gy
            mLastGY = gy
            Exit Sub
        End If
    Next i
End Sub


Private Sub MovePipes(ByVal dt As Double)
    Dim i As Long
    Dim maxX As Double
    Dim found As Boolean

    For i = 1 To MAXP
        If pOn(i) Then
            pX(i) = pX(i) - PIPE_SPEED * dt
            If pX(i) + PIPE_W + CAP_OVER < 0# Then
                pOn(i) = False
            Else
                If Not found Or pX(i) > maxX Then
                    found = True
                    maxX = pX(i)
                End If
                If Not pPass(i) Then
                    If pX(i) + PIPE_W / 2# < BIRD_X + BIRD_W / 2# Then
                        pPass(i) = True
                        mScore = mScore + 1
                    End If
                End If
            End If
        End If
    Next i

    If Not found Then
        SpawnPipe
    ElseIf maxX <= SCR_W + 10 - PIPE_SPACING Then
        SpawnPipe
    End If
End Sub


Private Sub CheckCollisions()
    Dim i As Long
    Dim hx1 As Double, hx2 As Double, hy1 As Double, hy2 As Double
    Dim pLeft As Double
    Dim gy As Long, gb As Long
    Dim bodyX As Boolean, capX As Boolean

    If mBirdY + BIRD_H >= GROUND_Y Then
        mBirdY = GROUND_Y - BIRD_H
        Die
        Exit Sub
    End If

    ' Slightly smaller than the drawn bird, so near-misses feel fair.
    hx1 = BIRD_X + 5
    hx2 = BIRD_X + 31
    hy1 = mBirdY + 4#
    hy2 = mBirdY + 21#

    For i = 1 To MAXP
        If pOn(i) Then
            pLeft = pX(i)
            gy = pGY(i)
            gb = gy + PIPE_GAP
            bodyX = (hx2 > pLeft) And (hx1 < pLeft + PIPE_W)
            capX = (hx2 > pLeft - CAP_OVER) And (hx1 < pLeft + PIPE_W + CAP_OVER)

            If capX Then
                ' top pipe: body above its cap, then the cap itself
                If bodyX And hy1 < gy - CAP_H Then Die: Exit Sub
                If hy1 < gy And hy2 > gy - CAP_H Then Die: Exit Sub
                ' bottom pipe: cap first, then the body below it
                If hy2 > gb And hy1 < gb + CAP_H Then Die: Exit Sub
                If bodyX And hy2 > gb + CAP_H Then Die: Exit Sub
            End If
        End If
    Next i
End Sub


' =====================================================================
'  DRAWING  (all into an off-screen bitmap, then one copy to the window)
' =====================================================================

Private Sub DrawScene()
    Dim i As Long, x As Long, offs As Long
    Dim pLeft As Long, gy As Long

    ' Quick white flash the moment the bird crashes.
    If mState = ST_DEAD Then
        If mDeadT < 0.06 Then
            Box 0, 0, SCR_W, SCR_H, B_WHITE
            Exit Sub
        End If
    End If

    ' sky and clouds
    Box 0, 0, SCR_W, GROUND_Y, B_SKY
    DrawCloud Wrap(30# - mCloudX * 0.6, SCR_W + 140) - 70, 70
    DrawCloud Wrap(190# - mCloudX * 0.4, SCR_W + 140) - 70, 150
    DrawCloud Wrap(330# - mCloudX * 0.8, SCR_W + 140) - 70, 35

    ' pipes
    For i = 1 To MAXP
        If pOn(i) Then
            pLeft = CLng(pX(i))
            gy = pGY(i)
            Seg pLeft, -6, pLeft + PIPE_W, gy - CAP_H + 3
            Seg pLeft - CAP_OVER, gy - CAP_H, pLeft + PIPE_W + CAP_OVER, gy
            Seg pLeft - CAP_OVER, gy + PIPE_GAP, pLeft + PIPE_W + CAP_OVER, gy + PIPE_GAP + CAP_H
            Seg pLeft, gy + PIPE_GAP + CAP_H - 3, pLeft + PIPE_W, GROUND_Y + 6
        End If
    Next i

    ' ground
    Box 0, GROUND_Y, SCR_W, SCR_H, B_SAND
    Box 0, GROUND_Y, SCR_W, GROUND_Y + 4, B_DARK
    Box 0, GROUND_Y + 4, SCR_W, GROUND_Y + 20, B_GRASS
    offs = CLng(mGroundX)
    For x = -offs To SCR_W Step 32
        Box x, GROUND_Y + 4, x + 16, GROUND_Y + 20, B_GRASS2
    Next x
    Box 0, GROUND_Y + 20, SCR_W, GROUND_Y + 23, B_SAND2

    DrawBird

    ' text
    Select Case mState
        Case ST_READY
            Txt "GET READY", HALF_W, 110, F_MID
            Txt "SPACE or click to flap", HALF_W, 160, F_SMALL
            If mBest > 0 Then Txt "Best: " & mBest, HALF_W, 195, F_SMALL
        Case ST_PLAY
            Txt CStr(mScore), HALF_W, 28, F_BIG
        Case ST_DEAD
            Txt "GAME OVER", HALF_W, 90, F_MID
            If mDeadT > 0.3 Then
                Box 50, 150, 310, 270, B_DARK
                Box 53, 153, 307, 267, B_SAND
                Txt "SCORE  " & mScore, HALF_W, 168, F_MID
                Txt "BEST  " & mBest, HALF_W, 214, F_MID
            End If
            If mDeadT > 0.6 Then Txt "press SPACE to retry", HALF_W, 295, F_SMALL
    End Select
End Sub


Private Sub DrawBird()
    Dim py As Long, wy As Long

    py = CLng(mBirdY)
    If mState = ST_DEAD Then
        wy = 0
    Else
        wy = ((CLng(Int(mClock * 14#)) Mod 3) - 1) * 3     ' flapping wing
    End If

    Box BIRD_X + 2, py + 1, BIRD_X + 32, py + 23, B_DARK          ' outline
    Box BIRD_X + 4, py + 3, BIRD_X + 30, py + 21, B_BIRD          ' body
    Box BIRD_X + 20, py + 3, BIRD_X + 32, py + 15, B_DARK         ' eye
    Box BIRD_X + 21, py + 4, BIRD_X + 31, py + 14, B_WHITE
    Box BIRD_X + 26, py + 7, BIRD_X + 30, py + 12, B_DARK
    Box BIRD_X + 24, py + 14, BIRD_X + 37, py + 24, B_DARK        ' beak
    Box BIRD_X + 25, py + 15, BIRD_X + 36, py + 23, B_BEAK
    Box BIRD_X + 5, py + 11 + wy, BIRD_X + 19, py + 20 + wy, B_DARK   ' wing
    Box BIRD_X + 6, py + 12 + wy, BIRD_X + 18, py + 19 + wy, B_WING
End Sub


' One shaded pipe piece (used for both bodies and caps).
Private Sub Seg(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)
    Box x1, y1, x2, y2, B_DARK
    Box x1 + 3, y1 + 3, x2 - 3, y2 - 3, B_PIPE
    Box x1 + 6, y1 + 3, x1 + 14, y2 - 3, B_PLIGHT
    Box x2 - 14, y1 + 3, x2 - 3, y2 - 3, B_PSHADE
End Sub


Private Sub DrawCloud(ByVal x As Long, ByVal y As Long)
    Box x + 12, y, x + 52, y + 14, B_CLOUD
    Box x, y + 8, x + 66, y + 24, B_CLOUD
End Sub


Private Sub Box(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal b As Long)
    Dim rc As GRect
    rc.Left = x1
    rc.Top = y1
    rc.Right = x2
    rc.Bottom = y2
    FillRect mDC, rc, hB(b)
End Sub


' Centred text: white with a dark drop shadow.
Private Sub Txt(ByVal s As String, ByVal x As Long, ByVal y As Long, ByVal fnt As Long)
    SelectObject mDC, hF(fnt)
    SetTextColor mDC, RGB(84, 56, 71)
    TextOutA mDC, x + 2, y + 3, s, Len(s)
    SetTextColor mDC, RGB(255, 255, 255)
    TextOutA mDC, x, y, s, Len(s)
End Sub


' Copies the finished picture to the window in one go.
Private Sub PresentFrame()
    Dim hdcWin As LongPtr

    hdcWin = GetDC(mHwnd)
    If hdcWin = 0 Then Exit Sub
    If WIN_SCALE = 1 Then
        BitBlt hdcWin, 0, 0, SCR_W, SCR_H, mDC, 0, 0, SRCCOPY
    Else
        SetStretchBltMode hdcWin, 3
        StretchBlt hdcWin, 0, 0, SCR_W * WIN_SCALE, SCR_H * WIN_SCALE, mDC, 0, 0, SCR_W, SCR_H, SRCCOPY
    End If
    ReleaseDC mHwnd, hdcWin
    ValidateRect mHwnd, 0       ' tell Windows nothing needs repainting
End Sub


' =====================================================================
'  WINDOW AND GRAPHICS SET-UP / TEAR-DOWN
' =====================================================================

Private Function OpenGameWindow() As Boolean
    Dim rc As GRect
    Dim sty As Long
    Dim w As Long, h As Long, x As Long, y As Long

    sty = WS_CAPTION Or WS_SYSMENU Or WS_MINIMIZEBOX
    rc.Left = 0
    rc.Top = 0
    rc.Right = SCR_W * WIN_SCALE
    rc.Bottom = SCR_H * WIN_SCALE
    AdjustWindowRect rc, sty, 0
    w = rc.Right - rc.Left
    h = rc.Bottom - rc.Top
    x = (GetSystemMetrics(0) - w) \ 2
    y = (GetSystemMetrics(1) - h) \ 2
    If x < 0 Then x = 0
    If y < 0 Then y = 0

    ' A plain STATIC control used as a top-level window: it has no
    ' painting of its own (SS_BLACKRECT) and Windows closes it for us.
    mHwnd = CreateWindowExA(0, "STATIC", "Flappy Bird  -  SPACE / click to flap, ESC to quit", _
                            sty Or SS_BLACKRECT Or SS_NOTIFY Or WS_VISIBLE, _
                            x, y, w, h, 0, 0, GetModuleHandleA(vbNullString), 0)
    If mHwnd = 0 Then Exit Function

    SetForegroundWindow mHwnd
    SetFocus mHwnd
    OpenGameWindow = True
End Function


Private Sub InitGdi()
    Dim hdcWin As LongPtr
    Dim colors As Variant
    Dim i As Long

    hdcWin = GetDC(mHwnd)
    mDC = CreateCompatibleDC(hdcWin)
    mBmp = CreateCompatibleBitmap(hdcWin, SCR_W, SCR_H)
    ReleaseDC mHwnd, hdcWin
    mOldBmp = SelectObject(mDC, mBmp)

    colors = Array(RGB(112, 197, 206), RGB(236, 248, 250), RGB(115, 191, 46), _
                   RGB(172, 232, 100), RGB(84, 148, 36), RGB(84, 56, 71), _
                   RGB(140, 214, 64), RGB(115, 191, 46), RGB(222, 216, 149), _
                   RGB(203, 191, 107), RGB(250, 215, 40), RGB(240, 100, 30), _
                   RGB(255, 244, 160), RGB(255, 255, 255))
    For i = 0 To NB - 1
        hB(i) = CreateSolidBrush(CLng(colors(i)))
    Next i

    hF(F_BIG) = CreateFontA(-46, 0, 0, 0, 700, 0, 0, 0, 0, 0, 0, 4, 0, "Arial")
    hF(F_MID) = CreateFontA(-30, 0, 0, 0, 700, 0, 0, 0, 0, 0, 0, 4, 0, "Arial")
    hF(F_SMALL) = CreateFontA(-18, 0, 0, 0, 700, 0, 0, 0, 0, 0, 0, 4, 0, "Arial")
    mOldFont = SelectObject(mDC, hF(F_BIG))

    SetBkMode mDC, 1          ' transparent text background
    SetTextAlign mDC, 6       ' x is the centre of the text
End Sub


Private Sub CloseGame()
    Dim i As Long

    On Error Resume Next
    If mHwnd <> 0 Then
        If IsWindow(mHwnd) <> 0 Then DestroyWindow mHwnd
    End If
    mHwnd = 0

    If mDC <> 0 Then
        If mOldFont <> 0 Then SelectObject mDC, mOldFont
        If mOldBmp <> 0 Then SelectObject mDC, mOldBmp
    End If
    For i = 0 To 2
        If hF(i) <> 0 Then DeleteObject hF(i)
        hF(i) = 0
    Next i
    For i = 0 To NB - 1
        If hB(i) <> 0 Then DeleteObject hB(i)
        hB(i) = 0
    Next i
    If mBmp <> 0 Then DeleteObject mBmp
    If mDC <> 0 Then DeleteDC mDC
    mBmp = 0
    mDC = 0
    mOldBmp = 0
    mOldFont = 0
End Sub


' =====================================================================
'  HELPERS
' =====================================================================

' High-resolution clock in seconds.
Private Function NowSec() As Double
    Dim c As Currency
    If mFreq = 0 Then
        NowSec = Timer
    Else
        QueryPerformanceCounter c
        NowSec = CDbl(c) / CDbl(mFreq)
    End If
End Function


' v wrapped into the range 0 .. span-1
Private Function Wrap(ByVal v As Double, ByVal span As Long) As Long
    Dim r As Double
    r = v - Int(v / span) * span
    Wrap = CLng(Int(r))
End Function


Private Function GetSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = SHEET_NAME
    End If
    Set GetSheet = ws
End Function
