VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   Caption         =   "Form1"
   ClientHeight    =   9810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19725
   LinkTopic       =   "Form1"
   ScaleHeight     =   9810
   ScaleWidth      =   19725
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   855
      Left            =   4200
      TabIndex        =   20
      Top             =   7680
      Width           =   2055
   End
   Begin VB.ListBox List2 
      Height          =   2400
      ItemData        =   "Listas2(2) (Zurba_Ferramte).frx":0000
      Left            =   2880
      List            =   "Listas2(2) (Zurba_Ferramte).frx":0002
      TabIndex        =   19
      Top             =   4920
      Width           =   5175
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Index           =   4
      Left            =   480
      TabIndex        =   18
      Top             =   7680
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Index           =   3
      Left            =   480
      TabIndex        =   17
      Top             =   6840
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Index           =   2
      Left            =   480
      TabIndex        =   16
      Top             =   6000
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Index           =   1
      Left            =   480
      TabIndex        =   15
      Top             =   5160
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Index           =   0
      Left            =   480
      TabIndex        =   14
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Calcular Promedio"
      Height          =   735
      Left            =   6720
      TabIndex        =   12
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   3240
      TabIndex        =   6
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   3240
      TabIndex        =   5
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   3240
      TabIndex        =   4
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   3240
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      ItemData        =   "Listas2(2) (Zurba_Ferramte).frx":0004
      Left            =   4440
      List            =   "Listas2(2) (Zurba_Ferramte).frx":0006
      TabIndex        =   1
      Top             =   240
      Width           =   6495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Dar notas"
      Height          =   735
      Left            =   4800
      TabIndex        =   0
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label2 
      Height          =   975
      Left            =   11160
      TabIndex        =   13
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Historia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   240
      TabIndex        =   11
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "C.ciudadana"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   240
      TabIndex        =   10
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Lengua"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Matematicas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Geografia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim notaMayor, notaMenor, suma As Integer
Dim promedio As Double
Dim A As Integer
Private Sub Command1_Click()
'------------------Dar notas-------------------
If List1.ListCount < 5 Then

    For A = 0 To 4
       
        If Val(Text1(A).Text) < 0 Or Val(Text1(A).Text) > 10 Then
            
            List1.AddItem "Escribiste mal la nota"
            
        ElseIf Text1(A).Text = "" Then
            
            List1.AddItem "escribi algo bobo"
            
        Else
            
            List1.AddItem Label1(A).Caption & ": " & Text1(A).Text
               
        End If
        
    
        
    Next A
    

End If
       

End Sub
Private Sub Command2_Click()
'-------------------------Calcular el promedio-----------------------
    notaMayor = Text1(0).Text
    notaMenor = Text1(0).Text

    For A = 0 To 4
    
        suma = suma + Text1(A).Text
        
        If Text1(A).Text > notaMayor Then
            notaMayor = Text1(A).Text
            
        ElseIf Text1(A).Text < notaMenor Then
            notaMenor = Text1(A).Text
            
        Else
            
        End If
        
    Next A
    
    promedio = suma / 5
    
    Label2.Caption = "la nota mayor: " & notaMayor & vbCrLf & "la nota menor: " & notaMenor & vbCrLf & "el promedio: " & promedio
    





End Sub
Private Sub Command3_Click()
Dim nombres(4) As String
       
If List2.ListCount < 5 Then
       
       For A = 0 To 4
             
             nombres(A) = Text2(A).Text
             
            If nombres(A) = "" Then
            
                List2.AddItem "escribi algo tonto"
                
            Else
                
                List2.AddItem "nombre: " & nombres(A)
            

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
Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)

        If Index >= 0 And Index <= 4 Then
           If KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Then
                KeyAscii = KeyAscii
           ElseIf KeyAscii = 32 Then
                KeyAscii = KeyAscii
           ElseIf KeyAscii = 8 Then
                KeyAscii = KeyAscii
           Else
                KeyAscii = 0
                
           End If
        End If
    

End Sub
