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
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   735
      Left            =   10920
      TabIndex        =   27
      Top             =   7680
      Width           =   3015
   End
   Begin VB.TextBox Text4 
      Height          =   975
      Left            =   10920
      TabIndex        =   26
      Top             =   6480
      Width           =   3015
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   615
      Left            =   20760
      TabIndex        =   25
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   18960
      TabIndex        =   24
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Modificar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10800
      TabIndex        =   23
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Dar notas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10800
      TabIndex        =   22
      Top             =   4200
      Width           =   2175
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      ItemData        =   "Listas2(2) (Zurba_Ferramte).frx":0000
      Left            =   13080
      List            =   "Listas2(2) (Zurba_Ferramte).frx":0002
      TabIndex        =   21
      Top             =   4200
      Width           =   8655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Introducir nombre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5520
      TabIndex        =   20
      Top             =   6960
      Width           =   2055
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      ItemData        =   "Listas2(2) (Zurba_Ferramte).frx":0004
      Left            =   3960
      List            =   "Listas2(2) (Zurba_Ferramte).frx":0006
      TabIndex        =   19
      Top             =   4560
      Width           =   5895
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   480
      TabIndex        =   18
      Top             =   7680
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   480
      TabIndex        =   17
      Top             =   6840
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   480
      TabIndex        =   16
      Top             =   6000
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   480
      TabIndex        =   15
      Top             =   5160
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   480
      TabIndex        =   14
      Top             =   4320
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Calcular Promedio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      ItemData        =   "Listas2(2) (Zurba_Ferramte).frx":0008
      Left            =   4440
      List            =   "Listas2(2) (Zurba_Ferramte).frx":000A
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
   Begin VB.Label Label3 
      Height          =   2055
      Left            =   15960
      TabIndex        =   28
      Top             =   6600
      Width           =   3135
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   11160
      TabIndex        =   13
      Top             =   240
      Width           =   3615
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
Dim coincidencias(7) As Integer

Private Sub Command1_Click()
'------------------Dar notas-------------------

List1.Clear

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
    notaMayor = CInt(Text1(0).Text)
    notaMenor = CInt(Text1(0).Text)
    suma = 0
    promedio = 0
    
    
    For A = 0 To 4
    
        suma = suma + Val(Text1(A).Text)
        
        If CInt(Text1(A).Text) > notaMayor Then
            
            notaMayor = CInt(Text1(A).Text)
        End If

        If CInt(Text1(A).Text) < notaMenor Then
            
            notaMenor = CInt(Text1(A).Text)
        
        End If

    Next A
    
    
    promedio = suma / 5
    
    Label2.Caption = "la nota mayor: " & notaMayor & vbCrLf & "la nota menor: " & notaMenor & vbCrLf & "el promedio: " & promedio
    





End Sub
Private Sub Command3_Click()
Dim nombres(4) As String
       
       List2.Clear
       
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

Private Sub Command4_Click()
    
    If List2.SelCount <> 0 Then
        List3.AddItem (List2.List(List2.ListIndex)) & ", obtuvo la nota de: " & promedio
    Else
        List3.AddItem "no hay alumno pibe"
    End If
    
End Sub

Private Sub Command5_Click()
    If List3.SelCount > 0 Then
        List3.List(List3.ListIndex) = (List2.List(List2.ListIndex)) & " - " & promedio
    End If
End Sub

Private Sub Command6_Click()
    If List3.SelCount > 0 Then
        List3.List(List3.ListIndex) = (List2.List(List2.ListIndex)) & " - " & Text3.Text
    End If
End Sub

Private Sub Command7_Click()
'Dim subA, subB As Integer
'
'    For subA = 0 To (List3.ListCount - 1)
'
'        For subB = 9 To Len(List3.List(List3.ListIndex))
'
'            If Mid(List3.List(List3.ListIndex, Len(Text4.Text), Len(Text4.Text))) = Text4.Text Then
'
'
'
'            End If
'
'        Next subB
'
'    Next subA
    
End Sub

Private Sub Text1_KeyPress(index As Integer, keyascii As Integer)

    If index >= 0 And index <= 4 Then
        If keyascii >= 48 And keyascii <= 57 Then
            keyascii = keyascii
        ElseIf keyascii = 8 Then
            keyascii = keyascii
        Else
            keyascii = 0
        End If
    Else

    End If
    
End Sub

Private Sub Text2_KeyPress(index As Integer, keyascii As Integer)

        If index >= 0 And index <= 4 Then
           If keyascii >= 65 And keyascii <= 90 Or keyascii >= 97 And keyascii <= 122 Then
                keyascii = keyascii
           ElseIf keyascii = 32 Then
                keyascii = keyascii
           ElseIf keyascii = 8 Then
                keyascii = keyascii
           Else
                keyascii = 0
                
           End If
        End If
    

End Sub

'-------------------------------------------------------------------------SUBRUTINAS DOWNS---------------------------------------------------------------------

Private Sub Text4_Change()
Dim subA, subB As Integer
    
    For subA = 0 To (List3.ListCount - 1)
        
        For subB = 9 To Len(List3.List(List3.ListIndex))
            
            If Mid(List3.List(List3.ListIndex), Len(Text4.Text), Len(Text4.Text)) = Text4.Text Then
                
                Label3.BackColor = RGB(255, 255, 255)
                
            Else
                
                Label3.BackColor = &H8000000F
                
            End If
            
        Next subB
        
    Next subA
End Sub
