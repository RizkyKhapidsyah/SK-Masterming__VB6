VERSION 5.00
Begin VB.Form Form2 
   Caption         =   " [4..6]"
   ClientHeight    =   1575
   ClientLeft      =   4065
   ClientTop       =   3735
   ClientWidth     =   2430
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   1575
   ScaleWidth      =   2430
   Begin VB.CommandButton Command1 
      Caption         =   "&Fin"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      MaxLength       =   2
      TabIndex        =   0
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Jugar"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Seleccion el numero de colores: de 4 a 6"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub
Private Sub Command2_Click()
k = 2
colores = Val(Text1.Text)
Load Form1
Form1.Show
Unload Me
'falta pasrle a principal q colores voy a usar
End Sub


Private Sub Form_load()
Command2.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If k = 1 Then End
End Sub

Private Sub Text1_Change()
If Text1.Text = "4" Or Text1 = "5" Or Text1 = "6" Then
Command2.Enabled = True
Command1.Enabled = False
Command2.SetFocus
colores = Val(Text1.Text)
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Text1.Text = ""
If KeyAscii <> 8 Then
        If Not IsNumeric("0" & Chr(KeyAscii)) Then
             KeyAscii = 0
        End If
 End If
 If KeyAscii = 13 Then KeyAscii = 0

 If Text1 = "4" Or Text1 = "5" Or Text1 = "6" Then
  Command1.Enabled = False
 Command2.SetFocus
 End If
End Sub
