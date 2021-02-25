VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "@2.000- JNG (Mastermind)"
   ClientHeight    =   4470
   ClientLeft      =   3150
   ClientTop       =   1950
   ClientWidth     =   4335
   BeginProperty Font 
      Name            =   "Benguiat Bk BT"
      Size            =   48
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "principal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "principal.frx":0442
   ScaleHeight     =   4470
   ScaleWidth      =   4335
   Begin VB.CommandButton Command2 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton Ju 
      Caption         =   "&Jugar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      ToolTipText     =   "aa"
      Top             =   3240
      Width           =   975
   End
   Begin VB.Shape Shape1 
      Height          =   1575
      Left            =   3120
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "o"
      Height          =   1215
      Left            =   3120
      TabIndex        =   8
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   3600
      TabIndex        =   5
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   3120
      TabIndex        =   4
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   3600
      TabIndex        =   3
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   3120
      TabIndex        =   2
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   3600
      TabIndex        =   1
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Index           =   0
      Left            =   3120
      TabIndex        =   0
      Top             =   1560
      Width           =   375
   End
   Begin VB.Menu menu1 
      Caption         =   "menu1"
      Visible         =   0   'False
      Begin VB.Menu mnuac 
         Caption         =   "&Acerca de"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private sol(0 To 3) As Double
Private solju(0 To 3) As Byte
Private yy, xx As Integer

Private Sub nuevo()
'form1.autoredraw=true
Form1.Cls
Form1.Refresh
Randomize
'genero la solucion
For k = 0 To 3
solju(k) = Int(Rnd * colores)
Next k
Label1(0).ForeColor = RGB(255, 0, 0)
Label1(1).ForeColor = RGB(0, 255, 0)
Label1(2).ForeColor = RGB(0, 0, 255)
Label1(3).ForeColor = RGB(255, 255, 0)
Label1(4).ForeColor = RGB(255, 0, 255)
Label1(5).ForeColor = RGB(0, 255, 255)
Label2.ForeColor = RGB(255, 0, 0)
Form1.FillStyle = solid
'pinta la solucion
'Form1.ForeColor = RGB(255, 0, 0)
'For xx = 1 To 4
'For yy = 1 To 10
'coordenadas de los puntos en pantalla
'(1,10)..(4,10)
'..............
'...(xx,yy)....
'..............
'(1,1)....(4,1)
'
'Circle (770 + ((xx - 1) * 360), 1310 + ((10 - yy) * 280)), 80
'Next yy
'Next xx
'coordenadas de los resultados por pantalla con yy=1..10
'yy = 10
'Circle (2200 + (0 * 120), 1310 + ((10 - yy) * 280)), 60
'Circle (2200 + (1 * 120), 1310 + ((10 - yy) * 280)), 60
'Circle (2200 + (2 * 120), 1310 + ((10 - yy) * 280)), 60
'Circle (2200 + (3 * 120), 1310 + ((10 - yy) * 280)), 60
Ju.Enabled = False
For j = (colores + 1) To 6
Label1(j - 1).Enabled = False
Next j
yy = 1
xx = 1
Label2.Tag = 0
For j = 0 To 3
sol(j) = 6
Next j
End Sub
Private Sub pintar()
For i = 1 To 4
Form1.ForeColor = Label1(solju(i - 1)).ForeColor
Form1.FillColor = Label1(solju(i - 1)).ForeColor
Circle (770 + ((i - 1) * 360), 800), 120
Next i
End Sub
Private Sub Command2_Click()
End
End Sub

Private Sub Form_load()
nuevo
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
 If y > (1310 + ((10 - yy) * 280) - 90) And y < (1310 + ((10 - yy) * 280) + 90) Then
    For i = 1 To 4
     If x > ((770 + ((i - 1) * 360) - 90)) And _
        x < ((770 + ((i - 1) * 360) + 90)) Then
        xx = i
        sol(i - 1) = Label2.Tag
        Form1.ForeColor = Label2.ForeColor
        Form1.FillColor = Label2.ForeColor
        Circle (770 + ((xx - 1) * 360), 1310 + ((10 - yy) * 280)), 80
        Exit For
      End If
    Next i
   End If
End If
If sol(0) <> "6" And sol(1) <> "6" And sol(2) <> "6" And sol(3) <> "6" Then
    Ju.Enabled = True
    Ju.SetFocus
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
Form1.PopupMenu menu1
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub Ju_Click()
Dim n, b As Byte
Dim bool(0 To 3) As Boolean ' para marcar solju(0..3)
Dim bool2(0 To 3) As Boolean ' para marcar sol(0..3)
n = 0
b = 0
For i = 0 To 3
bool(i) = False 'no pillada
bool2(i) = False 'no pillada
Next i
Ju.Enabled = False
Form1.SetFocus
' bool(i)->  solju(0,1,2,3) la solucion del juego
' bool2(j)->  sol(0,1,2,3) la solucion del usuario
'con esto saco las negras
For i = 0 To 3
    If sol(i) = solju(i) Then
    bool(i) = True
    bool2(i) = True
    n = n + 1
    End If
Next i
'con esto saco las blancas
For j = 0 To 3
    If bool2(j) = True Then
    GoTo 20
  Else
    For i = 0 To 3
        If bool(i) = False Then
           If solju(i) = sol(j) Then
            bool(i) = True
            bool2(j) = True
            b = b + 1
            Exit For
           End If
       End If
    Next i
  End If
20:
Next j
'tengo b--> blancas
'tengo n-->negras
'pinto las negras
Form1.ForeColor = vbBlack
Form1.FillColor = vbBlack
For i = 0 To (n - 1)
Circle (2200 + (i * 120), 1310 + ((10 - yy) * 280)), 60
Next i
Form1.ForeColor = vbWhite
Form1.FillColor = vbWhite
For i = n To (n + b - 1)
Circle (2200 + (i * 120), 1310 + ((10 - yy) * 280)), 60
Next i
If n = 4 Then
    'ha acertado
    pintar
    MsgBox "Muy bien ha acertado la combinacion", , "Fin juego"
    nuevo
Else
    yy = yy + 1
    For i = 0 To 3
        sol(i) = 6
    Next i
    If yy = 11 Then
        pintar
        MsgBox "Lo siento no ha acertado la combinacion", , "Fin juego"
        nuevo
    End If
End If
End Sub

Private Sub Label1_Click(Index As Integer)
Label2.ForeColor = Label1(Index).ForeColor
Label2.Tag = Index
End Sub

Private Sub mnuac_Click()
acerca2.Show (1)
End Sub
