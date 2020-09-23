VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   ScaleHeight     =   378
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   390
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrBrain 
      Interval        =   10
      Left            =   315
      Top             =   90
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type Ball
    CX As Single
    CY As Single
    vx As Single
    vy As Single
End Type

Const NUMBALLS As Integer = 10
Const PI As Single = 3.14159
Const IMPACT As Single = 4
Const FRICTION As Single = 0.95
Const RADIUS As Integer = 10
Const TRAILDIST As Integer = 20

Dim TableBalls(0 To NUMBALLS - 1) As Ball

Private Sub Form_Load()
    Dim i As Integer
    Dim X As Integer
    
    TableBalls(0).CX = 0.75 * Me.ScaleWidth '0 is cueball
    TableBalls(0).CY = Me.ScaleHeight \ 2
    
    Randomize (Rnd * Timer)
    
    X = Rnd * 0.5 * Me.ScaleWidth
    
    For i = 1 To NUMBALLS - 1
    
        With TableBalls(i)
            
            .CX = X
            .CY = (0.5 * TableBalls(0).CY) + (i * (RADIUS * 2))
        
        End With
        
    Next
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim R As Single
    Dim ang As Single
    
    With TableBalls(0)
        R = Sqr((X - .CX) ^ 2 + (Y - .CY) ^ 2)
    
        If R > TRAILDIST Then
            R = TRAILDIST
        End If
        
        If Not ((X - .CX) = 0) Then
            ang = Atn((Y - .CY) / (X - .CX))
        Else
            If Y > .CY Then
                ang = PI / 2
            Else
                ang = -PI / 2
            End If
        End If
        
        If X < .CX Then
            ang = ang + PI
        End If
        
        Me.Caption = R & " - " & (ang * 180 / PI)
        
        .vx = R * Cos(ang)
        .vy = R * Sin(ang)
        
    End With
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ang As Single
    Dim dx As Single
    Dim dy As Single
    
    With TableBalls(0)
        
        Me.Line (.CX, .CY)-(X, Y), vbGreen
        
    End With
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
    
        tmrBrain.Enabled = False
        
    Else
        
        tmrBrain.Enabled = True
        
    End If
End Sub

Private Sub tmrBrain_Timer()
    Dim ColDist As Single
    Dim i As Integer
    Dim j As Integer
    Dim ang As Single
    Dim R As Single
    Dim dx As Single
    Dim dy As Single
    
    Me.Picture = Nothing
    
    For i = 0 To NUMBALLS - 1
    
        With TableBalls(i)
            .CX = .CX + .vx
            .CY = .CY + .vy
            .vx = .vx * FRICTION
            .vy = .vy * FRICTION
            
            If (.CX < RADIUS) Then
                .vx = -.vx
                .CX = RADIUS
            ElseIf (.CX > (Me.ScaleWidth - 1 - RADIUS)) Then
                .vx = -.vx
                .CX = Me.ScaleWidth - 1 - RADIUS
            End If
            
            If (.CY < RADIUS) Then
                .vy = -.vy
                .CY = RADIUS
            ElseIf (.CY > (Me.ScaleHeight - 1 - RADIUS)) Then
                .vy = -.vy
                .CY = Me.ScaleHeight - 1 - RADIUS
            End If
            
            For j = 0 To NUMBALLS - 1
                
                ColDist = Sqr((TableBalls(j).CX - .CX) ^ 2 + (TableBalls(j).CY - .CY) ^ 2)
                
                If ColDist <= (2 * RADIUS) And (Not (i = j)) Then
                    
                    '===================================================
                    'Keep from going inside other balls
                    'Get the angle of collision
                    If Not ((TableBalls(j).CX - .CX) = 0) Then
                        ang = Atn((TableBalls(j).CY - .CY) / (TableBalls(j).CX - .CX))
                    Else '+ or - 90 degrees when change in x is 0
                        If TableBalls(j).CY > .CY Then
                            ang = PI / 2
                        Else
                            ang = -PI / 2
                        End If
                    End If
                    'Make sure it is the correct angle
                    'atn() only gives us a value from -90 to +90[pi/2]
                    If TableBalls(j).CX < .CX Then
                        ang = ang + PI
                    End If
                    
                    'If a ball is on or inside the other, push it to just outside
                    dx = -2.1 * RADIUS * Cos(ang)
                    dy = -2.1 * RADIUS * Sin(ang)
                    .CX = dx + TableBalls(j).CX
                    .CY = dy + TableBalls(j).CY
                    
                    'There is probably a better way to do this but
                    'this will get the corrected angle(from point of contact)
                    If Not ((TableBalls(j).CX - .CX) = 0) Then
                        ang = Atn((TableBalls(j).CY - .CY) / (TableBalls(j).CX - .CX))
                    Else
                        If TableBalls(j).CY > .CY Then
                            ang = PI / 2
                        Else
                            ang = -PI / 2
                        End If
                    End If
                    If TableBalls(j).CX < .CX Then
                        ang = ang + PI
                    End If
                    '===================================================
                    
                    '========================================
                    'Figure reflection angle
                    'normalize angle to 0 to 2PI
                    If ang > (2 * PI) Then
                        ang = (2 * PI) - ang
                    ElseIf ang < 0 Then
                        ang = ang + (2 * PI)
                    End If
                    
                    R = Sqr(.vx * .vx + .vy * .vy)
                    
                    dx = R * Cos(ang)
                    dy = R * Sin(ang)
                    
                    '========================================
                    
                    'Put the velocity of the 'attacking' ball into the target
                    TableBalls(j).vx = dx
                    TableBalls(j).vy = dy
                    
                    'Decrease velocity from the attacker
                    .vx = 0.2 * .vx
                    .vy = 0.2 * .vy
                End If
                
            Next
            
            If i = 0 Then
                Me.Circle (.CX, .CY), RADIUS, vbWhite
            Else
                Me.Circle (.CX, .CY), RADIUS, vbBlue
            End If
            
        End With
    
    Next
    
    Me.Refresh
End Sub
