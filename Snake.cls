VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Snake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pLength As Long
Private pX As Long
Private pY As Long
Private pDirection As String
Private pCoords()
Private pPoopX As Long
Private pPoopY As Long
Private pDie As Boolean

''''''''''''''''''''''
' Die
''''''''''''''''''''''
Public Property Get Die() As Boolean
    Die = pDie
End Property
Public Property Let Die(Value As Boolean)
    pDie = Value
End Property


''''''''''''''''''''''
' Length
''''''''''''''''''''''
Public Property Get Length() As Long
    Length = pLength
End Property
Public Property Let Length(Value As Long)
    pLength = Value
End Property


''''''''''''''''''''''
' X value (head)
''''''''''''''''''''''
Public Property Get x() As Long
    x = pX
End Property
Public Property Let x(Value As Long)
    pX = Value
End Property



''''''''''''''''''''''
' Y value (head)
''''''''''''''''''''''
Public Property Get y() As Long
    y = pY
End Property
Public Property Let y(Value As Long)
    pY = Value
End Property



''''''''''''''''''''''
' X poop value (the spot just vacated by snake)
''''''''''''''''''''''
Public Property Get PoopX() As Long
    PoopX = pPoopX
End Property
Public Property Let PoopX(Value As Long)
    pPoopX = Value
End Property


''''''''''''''''''''''
' Y poop value (the spot just vacated by snake)
''''''''''''''''''''''
Public Property Get PoopY() As Long
    PoopY = pPoopY
End Property
Public Property Let PoopY(Value As Long)
    pPoopY = Value
End Property



''''''''''''''''''''''
' Direction
''''''''''''''''''''''
Public Property Get Direction() As String
    
    Direction = pDirection

End Property
Public Property Let Direction(Value As String)
    
    Select Case pDirection
    
        Case "up"
            If Value <> "down" Then pDirection = Value
            
        Case "down"
            If Value <> "up" Then pDirection = Value
            
        Case "left"
            If Value <> "right" Then pDirection = Value
            
        Case "right"
            If Value <> "left" Then pDirection = Value
            
    End Select
    
 
End Property

'Classes don't support public array members, so the coordinates are just a public array not attached to any class. Bad. Someone fix.
''''''''''''''''''''''
' Coordinates
''''''''''''''''''''''
'Public Property Get Coords() As Variant
'    Coords() = pCoords()
'End Property
'Public Property Let Coords(Value() As Variant)
'    pCoords = Value()
'End Property


Public Sub UpdateDirection()
    
    On Error Resume Next
    
    Me.Direction = GetText(CURRENTDIRECTORY & "\snek.txt")

End Sub


Public Sub Eat(pellets As Long)
    
    Me.Length = Me.Length + pellets * 2
    
    ReDim Preserve Coords(0 To Me.Length * 2 - 1)
    
    Coords(UBound(Coords) - 1) = Me.PoopX
    Coords(UBound(Coords)) = Me.PoopY
    
    
End Sub


Public Sub Slither()

    'I accidentally coded the whole game with x and y mixed up.
    'I'm not going back to fix it now, sorry.

    Select Case Me.Direction
        
        Case "up"
            
            Me.x = Me.x - 1
            
        Case "down"
            
            Me.x = Me.x + 1
            
        Case "left"
            
            Me.y = Me.y - 1
            
            
        Case "right"
            
            Me.y = Me.y + 1
            
    End Select

    'Capture the poop spot before we lose it, so we can color the ground
    Me.PoopX = Coords(UBound(Coords) - 1)
    Me.PoopY = Coords(UBound(Coords))

    'shift every point in the snake down a row
    Dim i As Long
    For i = UBound(Coords) To LBound(Coords) + 2 Step -1
        
        Coords(i) = Coords(i - 2)
                        
    Next i
    
    
    'capture new head position
    Coords(0) = Me.x
    Coords(1) = Me.y
    
    
    'check for collision
    For i = LBound(Coords) + 2 To UBound(Coords) Step 2
    
        If Me.x = Coords(i) And Me.y = Coords(i + 1) Then Me.Die = True
    
    Next i
    
    If Me.x = 1 Or Me.y = 1 Or Me.x = 36 Or Me.y = 74 Then Me.Die = True
        
    

End Sub

Private Sub Class_Initialize()
    Me.Length = 2
    Me.Die = False
    pDirection = "right"
    
    ReDim Coords(0 To Me.Length * 2 - 1)
        

    'Hardcoding the start position.
    'Could make a class constructor later to have this be variable
    Me.x = 6
    Me.y = 8
    
    Coords(0) = Me.x
    Coords(1) = Me.y
    Coords(2) = Me.x
    Coords(3) = Me.y - 1


End Sub

