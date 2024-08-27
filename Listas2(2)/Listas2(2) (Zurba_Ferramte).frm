VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   13290
   ScaleWidth      =   25080
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   735
      Index           =   4
      Left            =   7680
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   6480
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Index           =   3
      Left            =   7680
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Index           =   2
      Left            =   7680
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Index           =   1
      Left            =   7680
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Index           =   0
      Left            =   7680
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   3120
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   9480
      TabIndex        =   1
      Top             =   3240
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   10800
      TabIndex        =   0
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   975
      Index           =   4
      Left            =   4440
      TabIndex        =   11
      Top             =   2640
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   975
      Index           =   3
      Left            =   4440
      TabIndex        =   10
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   975
      Index           =   2
      Left            =   4440
      TabIndex        =   9
      Top             =   5880
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   975
      Index           =   1
      Left            =   4560
      TabIndex        =   8
      Top             =   7080
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   975
      Index           =   0
      Left            =   4440
      TabIndex        =   7
      Top             =   3720
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim A As Integer
Private Sub Command1_Click()

If List1.ListCount < 5 Then

    For A = 0 To 4
       
        If Val(Text1(A).Text) < 0 Or Val(Text1(A).Text) > 10 Then
            
            List1.AddItem "Escribiste mal la nota"
            
        Else
            
            List1.AddItem Label1(A).Caption & ": " & Text1(A).Text
            
        End If
        
    Next A

End If

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If Index >= 0 And Index <= 4 Then
        If KeyAscii >= 48 And KeyAscii <= 57 Then
            KeyAscii = KeyAscii
        ElseIf KeyAscii = 8 Then
            KeyAscii = KeyAscii
        Else
            KeyAscii = 0
        End If
    Else
        
    End If
    
End Sub
