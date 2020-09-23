VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Menù "
   ClientHeight    =   5220
   ClientLeft      =   405
   ClientTop       =   405
   ClientWidth     =   1455
   ForeColor       =   &H0000C000&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   1455
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Sample8"
      Height          =   375
      Index           =   7
      Left            =   120
      TabIndex        =   9
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sample7"
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sample6"
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   1215
      Begin VB.Label Label1 
         Caption         =   "(C) LOGO EMULATOR BY M.PISTORIO 29/11/05"
         ForeColor       =   &H000000C0&
         Height          =   975
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sample5"
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sample4"
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sample3"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sample2"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sample1"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)
  
  Form1.Show
  
  Select Case Index
    Case Is = 0: Form1.Sample1
    Case Is = 1: Form1.Sample2
    Case Is = 2: Form1.Sample3
    Case Is = 3: Form1.Sample4
    Case Is = 4: Form1.Sample5
    Case Is = 5: Form1.Sample6
    Case Is = 6: Form1.Sample7
    Case Is = 7: Form1.Sample8
  End Select
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
End Sub
