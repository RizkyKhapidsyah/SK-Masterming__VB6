VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3345
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   6210
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "inicio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.Label Label4 
         Caption         =   "Forever Free"
         Height          =   255
         Left            =   4920
         TabIndex        =   4
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "@2.000 - By JNG"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         TabIndex        =   3
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Image Image3 
         Height          =   1320
         Left            =   360
         Picture         =   "inicio.frx":000C
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   1320
      End
      Begin VB.Image Image4 
         Height          =   1560
         Left            =   1320
         Picture         =   "inicio.frx":00AE
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   840
      End
      Begin VB.Line Line1 
         X1              =   5880
         X2              =   3360
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label2 
         Caption         =   "Version 2.0.b"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         TabIndex        =   2
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Mastermind"
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
         Left            =   3120
         TabIndex        =   1
         Top             =   360
         Width           =   2775
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   1200
         Picture         =   "inicio.frx":0150
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   2280
      End
      Begin VB.Image Image1 
         Height          =   840
         Index           =   0
         Left            =   840
         Picture         =   "inicio.frx":01F2
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub fin()
Unload Me
Form2.Show
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   fin
End Sub


Private Sub Form_load()
Load Form2
End Sub

Private Sub Frame1_Click()
   fin
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
fin
End Sub

Private Sub Image1_Click(Index As Integer)
fin
End Sub

Private Sub Image2_Click()
fin
End Sub

Private Sub Image3_Click()
fin
End Sub

Private Sub Image4_Click()
fin
End Sub

Private Sub Image5_Click()
fin
End Sub

