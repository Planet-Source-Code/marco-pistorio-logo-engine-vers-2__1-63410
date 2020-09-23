VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Graph"
   ClientHeight    =   6810
   ClientLeft      =   2340
   ClientTop       =   1125
   ClientWidth     =   9015
   ControlBox      =   0   'False
   ForeColor       =   &H00008000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   240
   ScaleMode       =   0  'User
   ScaleWidth      =   640
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LATO As Single
Public ANGOLO As Single
Public INC As Single

Public LV As Single
Public DIAG As Single
Public LG As Single

Public MaxX As Integer
Public MaxY As Integer

Dim Turtle As New LogoEngine

Public Sub Sample1()
   
   Dim Counter As Integer
   Dim j As Integer
   
   Turtle.Init
   Turtle.SetZoom 2
   Turtle.Turtle_Goto -50, 0
  
   For Counter = 0 To 36
         
      Turtle.SetColor Int(7 * Rnd(1)) + 1
      Turtle.PenDown
      
      For j = 1 To 3
        Turtle.Forward 50
        Turtle.Right 120
      Next j

      Turtle.PenUp
      Turtle.Forward 10
      Turtle.Right 10
   
   Next

End Sub

Public Sub Sample2()
    
    Turtle.Init
    Turtle.SetZoom 2
    Turtle.Turtle_Goto -60, -70
    
    LATO = 5
    LV = 4
      
    SIERP LATO, LV

End Sub

Private Sub SIERP(LATO As Single, LV As Single)
    
   Dim j As Integer
   
   DIAG = LATO / Sqr(3)
  
   For j = 1 To 4
    UNLATO LV
    Turtle.Right 45
    Turtle.Forward DIAG
    Turtle.Right 45
   Next j

End Sub

Private Sub UNLATO(LV As Single)

 If LV = 0 Then Exit Sub
  
 Turtle.SetColor Int(7 * Rnd(1)) + 1
 
 UNLATO LV - 1
 Turtle.Right 45
 Turtle.Forward DIAG
 Turtle.Right 45
 
 UNLATO LV - 1
 Turtle.Left 90
 Turtle.Forward LATO
 Turtle.Left 90
 
 UNLATO LV - 1
 Turtle.Right 45
 Turtle.Forward DIAG
 Turtle.Right 45
 
 UNLATO LV - 1

End Sub

Public Sub Sample3()
       
   Turtle.Init
   Turtle.SetZoom 2
   Turtle.Turtle_Goto -40, -80
   
   LG = 160
   LV = 4
  
   NEVE LG, LV
   
End Sub

Private Sub NEVE(LG As Single, LV As Single)
    
   Dim j As Integer
   
   For j = 1 To 3
    LATO_N LG, LV
    Turtle.Right 120
  Next j

End Sub

Private Sub LATO_N(LG As Single, LV As Single)
 
 Turtle.SetColor Int(7 * Rnd(1)) + 1
 
 If LV = 0 Then
  Turtle.Forward LG
  Exit Sub
 End If
             
 LATO_N LG / 3, LV - 1
 Turtle.Left 60

 LATO_N LG / 3, LV - 1
 Turtle.Right 120

 LATO_N LG / 3, LV - 1
 Turtle.Left 60

 LATO_N LG / 3, LV - 1

End Sub

Public Sub Sample4()
      
   Turtle.Init
   Turtle.SetZoom 2
   Turtle.Turtle_Goto 0, -80
   LG = 100
   LV = 4
  
   Ramo LG, LV

End Sub

Private Sub Ramo(LG As Single, LV As Single)

 If LV = 0 Then Exit Sub
 
 Turtle.SetColor Int(7 * Rnd(1)) + 1
 
 Turtle.Forward LG
 
 Turtle.Left 30
 Ramo LG / 3, LV - 1

 Turtle.Right 40
 Ramo LG / 2, LV - 1

 Turtle.Right 50
 Ramo LG / 1.5, LV - 1
 
 Turtle.Left 60
 Turtle.Back LG

End Sub

Public Sub Sample5()
      
   Turtle.Init
   Turtle.SetZoom 2
   Turtle.Turtle_Goto 20, -40
   LV = 90
  
   LATO = 30
   ANGOLO = 2
   INC = 20
   
   SPIRALE LATO, ANGOLO, INC
  
End Sub

Private Sub SPIRALE(LATO As Single, ANGOLO As Single, INC As Single)

  If LV = 0 Then
    Exit Sub
  End If
  
  Turtle.SetColor Int(7 * Rnd(1)) + 1
  
  Turtle.Forward LATO
  Turtle.Right ANGOLO
  
  LV = LV - 1

  SPIRALE LATO, ANGOLO + INC, INC

End Sub

Public Sub Sample6()
   
   Dim i As Integer
   
   Turtle.Init
   
   For i = 1 To 100
      Turtle.SetColor Int(7 * Rnd(1)) + 1
      Poligono i, 8
      Turtle.Right 90 + i
   Next i
      
End Sub

Public Sub Sample7()
  Fractals
End Sub

Private Sub Form_Load()
  Set Turtle.frm = Me
  MaxX = 640
  MaxY = 480
End Sub

Private Sub Poligono(N As Integer, L As Single)
  
  Dim i As Integer
  
  For i = 1 To N
    Turtle.Forward L
    Turtle.Right 360 / N
  Next i
  
End Sub

Private Sub CurveW(L As Single, Liv As Integer)

Dim LTerzi As Single

  If Liv = 0 Then
     Turtle.Forward L
    Else
       LTerzi = L / 3
       CurveW LTerzi, Liv - 1
       Turtle.Left 90
       CurveW LTerzi, Liv - 1
       Turtle.Right 90
       CurveW LTerzi, Liv - 1
       Turtle.Right 90
       CurveW LTerzi, Liv - 1
       Turtle.Left 90
       CurveW LTerzi, Liv - 1
  End If

End Sub

Private Sub CurveC(L As Single)
  
  Dim R2 As Single
  
  R2 = 1 / Sqr(3)

  If L <= 6 Then
    Turtle.Forward L
    Else
      Turtle.Left 45
      CurveC (L * R2)
      Turtle.Right 90
      CurveC (L * R2)
      Turtle.Left 45
   End If
   
End Sub

Private Sub Hilbert(Lungh As Single, Liv As Integer, Ang As Single)
  
  Dim Ang90 As Single
  
  If Liv <> 0 Then
      Ang90 = Ang * 90
      
      Turtle.Left Ang90
      Hilbert Lungh, Liv - 1, -Ang
      Turtle.Forward Lungh
      Turtle.Right Ang90
      Hilbert Lungh, Liv - 1, Ang
      Turtle.Forward Lungh
      Hilbert Lungh, Liv - 1, Ang
      Turtle.Right Ang90
      Turtle.Forward Lungh
      Hilbert Lungh, Liv - 1, -Ang
      Turtle.Left Ang90
   End If
   
End Sub
  
Private Sub Dragon(L As Single, A As Single, Segno As Integer)
  
  Dim R2 As Single
  
  R2 = 1 / Sqr(3)

  If L <= 12 Then
    Turtle.SetDirection A
    Turtle.Forward L
    Else
      Dragon L * R2, A + Segno * 45, 1
      Dragon L * R2, A - Segno * 45, -1
   End If
   
End Sub

Private Sub Fractals()
  
  Turtle.Init
  
  Dim i As Integer
  
  For i = 1 To 10
    
    Select Case i Mod 4
      Case Is = 0:
        Turtle.SetColor 3
      Case Is = 1:
        Turtle.SetColor 1
      Case Is = 2:
        Turtle.SetColor 7
      Case Is = 3:
        Turtle.SetColor 4
    End Select
    
    Turtle.SetDirection 0
    Turtle.Turtle_Goto MaxX / 6, MaxY / 6
    CurveW MaxX / 3, 3
    
    Turtle.SetDirection 90
    Turtle.Turtle_Goto -MaxX * 0.4, -MaxY * 0.4
    Dragon MaxX * 0.65, 0, 1

    Turtle.SetDirection 90
    Turtle.Turtle_Goto -MaxX / 3, MaxY * 0.2
    CurveC MaxX * 0.8

    Turtle.SetDirection 90
    Turtle.Turtle_Goto MaxX / 3, -MaxY / 3
    Hilbert MaxX / 70, 4, 1
    
    Turtle.SetZoom (Turtle.GetZoom * 0.6)
  
  Next
  
End Sub

Public Sub Sample8()
  Prato
End Sub

Public Sub Star(Raggio As Single, CentroX As Single, CentroY As Single)
  
  Dim NumPetali As Integer
  Dim Alfa As Single
  Dim c As Integer
  
  NumPetali = Int(Raggio * 20) Mod 12 + 7
  Alfa = 360 / NumPetali
      
      For c = 1 To NumPetali
        Turtle.Forward Raggio
        Turtle.Right Alfa
        Turtle.Back Raggio
      Next c

  Turtle.SetColor 2
    
End Sub
  
Public Sub Flower(X As Single, DimFoglia As Single, ColorFiore As Integer)
  
  Dim A2 As Single
    
    If X <= DimFoglia Then
      Turtle.SetColor ColorFiore
      Star DimFoglia, Turtle.TurtleX, Turtle.TurtleY
     Else
      Turtle.SetColor 2
      A2 = Int(46 * Rnd(1))
      Turtle.Left A2
      Turtle.Forward X
      Flower X * 0.65, DimFoglia, ColorFiore
      Turtle.PenUp
      Turtle.Back X
      Turtle.PenDown
      Turtle.Right A2 * 2
      Turtle.Forward X
      Flower X * 0.65, DimFoglia, ColorFiore
      Turtle.PenUp
      Turtle.Back X
      Turtle.PenDown
      Turtle.Left A2
    End If
  
End Sub

Public Sub Prato()
  
  Dim i As Integer
  Dim j As Integer
  
  Dim LineaFiori As Integer
  
  Turtle.Init
  LineaFiori = 300
  
  Turtle.SetZoom 0.2
  
  For i = 1 To 4
    Turtle.SetZoom (Turtle.GetZoom * 1.4)
    For j = 1 To 5 * Rnd(1) + 3
      Turtle.SetDirection 90
      Turtle.Turtle_Goto MaxX * 0.93 * Rnd(1) - 320, -LineaFiori
      Flower 50, Int(8 * Rnd(1)) + 8, Int(7 * Rnd(1)) + 1
    Next j
  Next i
  
End Sub
