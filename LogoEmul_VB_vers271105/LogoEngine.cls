VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LogoEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'****************************************************
'* This class module emulates a small subset of     *
'* commands LOGO.                                   *
'*                                                  *                                                  *
'* Methods :                                        *
'*                                                  *
'* .Init           Initializations                  *
'* .SetColor x     Sets the actual color (0-7)      *
'* .SetDirection x Sets the actual direction (0-360)*
'* .SetZoom x      Sets the zoom factor to x        *
'* .TurtleGoto x,y Moves the turtle at x,y coord    *
'* .TurtleHome     Moves the turtle at center screen*
'* .TurtleX        Returns the actual turtle X-Coord*
'* .TurtleY        Returns the actual Turtle Y-Coord*
'* .GetZoom        Returns the actual zoom factor   *
'* .Forward x      Turtle Goes to forward for x units
'* .Back x         Turtle Goes to back for x units  *
'* .Left x         Turtle rotates by a x angle (+)  *
'* .Right x        Turtle rotates by a x angle (-)  *
'* .PenUp          Turtle's pen line it's not visible
'* .PenDown        Turtle's pen line it's visible   *
'*                                                  *
'****************************************************

Option Explicit

Public frm As VB.Form

Private XI As Single
Private XF As Single
Private YI As Single
Private YF As Single
Private AN As Single
Private Pen_Status As Integer
Private Colore As OLE_COLOR
Private zoom As Single

Const PIGRECO = 3.14159265

Public Sub Init()
  
  frm.ScaleMode = 0
  frm.Scale (-320, 240)-(320, -240)
  Turtle_Home
  Me.frm.Cls
  PenDown
  SetColor 7
  zoom = 1
  
End Sub

Public Sub SetColor(Col As Integer)
  Select Case Col
    Case Is = 0: Colore = vbBlack
    Case Is = 1: Colore = vbRed
    Case Is = 2: Colore = vbGreen
    Case Is = 3: Colore = vbYellow
    Case Is = 4: Colore = vbBlue
    Case Is = 5: Colore = vbMagenta
    Case Is = 6: Colore = vbCyan
    Case Is = 7: Colore = vbWhite
  End Select
End Sub

Public Sub Turtle_Goto(X As Single, Y As Single)
  XI = X * zoom
  YI = Y * zoom
End Sub

Public Sub Turtle_Home()
  Turtle_Goto 0, 0
  SetDirection 90
End Sub

Public Function TurtleX() As Single
  TurtleX = XI
End Function

Public Function TurtleY() As Single
  TurtleY = YI
End Function

Public Sub SetDirection(A As Single)
    
  Dim IC As Single
  
  IC = A * PIGRECO / 180
  AN = IC
  
End Sub

Public Sub SetZoom(z_fact As Single)
    zoom = z_fact
End Sub

Public Function GetZoom() As Single
    GetZoom = zoom
End Function

Private Sub BK(PR As Single)
   
    XF = XI - PR * Cos(AN) * zoom
    YF = YI - PR * Sin(AN) * zoom
    
    If Pen_Status = 1 Then
      frm.Line (XI, YI)-(XF, YF), Colore
    End If
    
    XI = XF: YI = YF

End Sub

Private Sub FD(PR As Single)
   
    XF = XI + PR * Cos(AN) * zoom
    YF = YI + PR * Sin(AN) * zoom
    
    If Pen_Status = 1 Then
      frm.Line (XI, YI)-(XF, YF), Colore
    End If
    
    XI = XF: YI = YF

End Sub

Private Sub LT(PR As Single)
    
    Dim IC As Single
    
    IC = PR * PIGRECO / 180
    AN = AN + IC

End Sub

Private Sub RT(PR As Single)
    
    Dim IC As Single
    
    IC = PR * PIGRECO / 180
    AN = AN - IC

End Sub

Private Sub PU()
    Pen_Status = 0
End Sub

Private Sub PD()
    Pen_Status = 1
End Sub

Public Sub Forward(PR As Single)
    FD PR
End Sub

Public Sub Back(PR As Single)
    BK PR
End Sub

Public Sub Right(PR As Single)
    RT PR
End Sub

Public Sub Left(PR As Single)
    LT PR
End Sub

Public Sub PenUp()
    PU
End Sub

Public Sub PenDown()
    PD
End Sub
