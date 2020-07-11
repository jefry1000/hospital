VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9120
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16290
   LinkTopic       =   "Form1"
   ScaleHeight     =   9120
   ScaleWidth      =   16290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "MEDICAMENTOS"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   3960
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CLIENTES"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   3120
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EMPLEADOS"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   2280
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MENU"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5520
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   10185
      Left            =   0
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   18300
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Show
Form1.Hide

End Sub

Private Sub Command2_Click()
Form3.Show
Form1.Hide

End Sub

Private Sub Command3_Click()
Form4.Show
Form1.Hide

End Sub
