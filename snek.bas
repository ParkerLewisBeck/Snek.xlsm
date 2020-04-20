Attribute VB_Name = "snek"
Option Explicit
'The game boundaries are 34X72

'B2: BU35

'background color
'16737843

'snakecolor
'13408767


#If VBA7 Then
    '64 bit declares here
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
    Private Declare PtrSafe Function InvalidateRect Lib "user32" (ByVal hwnd As LongPtr, ByRef lpRect As Long, ByVal bErase As Long) As Long
    Private Declare PtrSafe Function UpdateWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function LockWindowUpdate Lib "user32" (ByVal hwndLock As LongPtr) As Long
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer
    Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long
#Else
    '32 bit declares here
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
    Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, ByRef lpRect As Long, ByVal bErase As Long) As Long
    Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer
    Private Declare Function timeGetTime Lib "winmm.dll" () As Long
#End If


Public Const KeyPressed As Integer = -32767
Private Const WM_SETREDRAW As Long = &HB&
Private Const WM_USER = &H400
Private Const EM_GETEVENTMASK = (WM_USER + 59)
Private Const EM_SETEVENTMASK = (WM_USER + 69)

Private Const SNAKECOLOR = 38
Private Const GROUNDCOLOR = 41
Public Coords() As Long
Public availableNumbers As Long


Sub ResetBoard()
    
    Sheets(1).Cells.Delete
    
    Dim i As Long
    'numbers across top row
    'For i = 1 To 73
    '    Cells(1, i).Value = i
    'Next
    
    Range("B2:BU35").Interior.ColorIndex = GROUNDCOLOR
    Range("B2:BU35").Font.Color = vbWhite
    Range("B2:BU35").HorizontalAlignment = xlCenter
    Range("B2:BU35").VerticalAlignment = xlCenter
    Columns("A:BV").ColumnWidth = 2.2
    
    Range("G6:H6").Interior.ColorIndex = SNAKECOLOR
    
    Range("A1").Select
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic

End Sub


Sub DrawSnake(player As Snake)

    Dim food As Variant
    food = Worksheets(1).Cells(player.x, player.y).Value
    
    Worksheets(1).Cells(player.x, player.y).Interior.ColorIndex = SNAKECOLOR
    
    If IsNumeric(food) And Not IsEmpty(food) And food <> "" Then
        
        player.Eat (food)
        availableNumbers = availableNumbers - 1
        Worksheets(1).Cells(player.x, player.y).Value = ""
        
    Else
        If player.PoopX <> 0 Then
            Worksheets(1).Cells(player.PoopX, player.PoopY).Interior.ColorIndex = GROUNDCOLOR
        End If
    End If
    

End Sub

Private Function ReadDirectionFromKey(player As Snake) As String
    ReadDirectionFromKey = player.Direction
    
    Select Case True

        Case GetAsyncKeyState(vbKeyRight):
            ReadDirectionFromKey = "right"

        Case GetAsyncKeyState(vbKeyLeft):
            ReadDirectionFromKey = "left"
            
        Case GetAsyncKeyState(vbKeyUp):
            ReadDirectionFromKey = "up"
        
        Case GetAsyncKeyState(vbKeyDown):
            ReadDirectionFromKey = "down"

    End Select

End Function

Sub Game()
    Dim hwnd As Long
    Dim i As Long
    Dim x As Long
    Dim y As Long
    
    Application.ScreenUpdating = True
      
    Dim player As Snake
    Set player = New Snake
    
    GenerateNumbers

    Dim lastFrameTime As Long
    lastFrameTime = timeGetTime
    
    Do
        DoEvents
        
        'All game code goes here.
        '*********************************
                
        'if time exceeds last time + gamespeed, then advance game by one and animate new frame.
        If timeGetTime - lastFrameTime > 50 Then
            DoEvents
            
            player.Direction = ReadDirectionFromKey(player)
            
            player.Slither
    
            DrawSnake player
            
            If player.Die Then Exit Do
            
            lastFrameTime = timeGetTime
        End If

        '*********************************
        
        'If availableNumbers < 4 Then GenerateNumbers   'not working at the moment.
        
    Loop
    
    MsgBox "Game over." & vbNewLine & "Your length: " & player.Length
    
    ResetBoard

End Sub

Sub GenerateNumbers()
    
    Dim i As Long
    Dim x As Long
    Dim y As Long
    Dim pellet As Long
    
    
    For i = 1 To 15
    
        pellet = Int((9 - 1 + 1) * Rnd + 1)
        x = Int((35 - 2 + 1) * Rnd + 2)
        y = Int((73 - 2 + 1) * Rnd + 2)
    
        Worksheets(1).Cells(x, y).Value = pellet
    
    Next
    
    availableNumbers = availableNumbers + 15
    
End Sub
