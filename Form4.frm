VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   8640
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16095
   LinkTopic       =   "Form4"
   ScaleHeight     =   8640
   ScaleWidth      =   16095
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9000
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command7 
      Caption         =   "FOTO"
      Height          =   375
      Left            =   5760
      TabIndex        =   19
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "MODIFICAR"
      Height          =   735
      Left            =   11280
      TabIndex        =   18
      Top             =   6840
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "GUARDAR"
      Height          =   735
      Left            =   11280
      TabIndex        =   17
      Top             =   5880
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ELIMINAR"
      Height          =   735
      Left            =   11280
      TabIndex        =   16
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "NUEVO"
      Height          =   735
      Left            =   11280
      TabIndex        =   15
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<"
      Height          =   735
      Left            =   5760
      TabIndex        =   14
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   735
      Left            =   5760
      TabIndex        =   13
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2655
      Left            =   720
      TabIndex        =   12
      Top             =   4680
      Width           =   4815
      Begin VB.Image Image2 
         Height          =   2370
         Left            =   120
         Picture         =   "Form4.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   4620
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   6120
      Top             =   1320
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Form4.frx":E802
      OLEDBString     =   $"Form4.frx":E8AE
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "medicamentos"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text5 
      DataField       =   "fecha_de_vencimiento"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2520
      TabIndex        =   11
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      DataField       =   "cantidad"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   10
      Top             =   3120
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      DataField       =   "costo"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   9
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      DataField       =   "nombre"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   8
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      DataField       =   "Id"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "FOTO"
      DataField       =   "foto"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   8160
      Width           =   1815
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "fecha de vencimiento"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CANTIDAD"
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "COSTO"
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE"
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO DE MEDICINA"
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   855
      Left            =   5280
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
   Begin VB.Image Image1 
      Height          =   9210
      Left            =   0
      Picture         =   "Form4.frx":E95A
      Stretch         =   -1  'True
      Top             =   240
      Width           =   16140
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.Recordset.MoveNext

If Adodc1.Recordset.EOF Then
    Adodc1.Recordset.MoveFirst
End If

X = App.Path
Image2.Picture = LoadPicture(X & "\" & Label7.Caption)

End Sub

Private Sub Command2_Click()
Adodc1.Recordset.MovePrevious
 
 If Adodc1.Recordset.BOF Then
    Adodc1.Recordset.MoveLast
 End If
 X = App.Path
    Image2.Picture = LoadPicture(X & "\" & Label7.Caption)

End Sub

Private Sub Command3_Click()
Adodc1.Recordset.AddNew
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text1.SetFocus
Command3.Enabled = False
Command6.Enabled = False
Command5.Enabled = True
Label7.Caption = ""
Image2.Picture = LoadPicture(Label7.Caption)
Command7.Enabled = True

End Sub

Private Sub Command4_Click()
Adodc1.Recordset.Delete
Adodc1.Recordset.MoveFirst
 '''x = App.Path
 '''Image2.Picture = LoadPicture(x & "\" & Label6.Caption)

End Sub

Private Sub Command5_Click()
  Adodc1.Recordset.Update
 Adodc1.Recordset.MoveFirst
 X = App.Path
 Image2.Picture = LoadPicture(X & "\" & Label7.Caption)

 Command4.Enabled = True
 Command5.Enabled = False
 Command3.Enabled = True
 Command6.Enabled = True

End Sub

Private Sub Command6_Click()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Command4.Enabled = False
Command5.Enabled = True
Command6.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False

End Sub

Private Sub Command7_Click()
 CommonDialog1.ShowOpen
    Image2.Picture = LoadPicture(CommonDialog1.FileName)
    Label7.Caption = CommonDialog1.FileTitle
    
    If Label7.Caption = "" Then
        MsgBox ("debe seleccionar una imagen")
    Else
     Label7.Caption = CommonDialog1.FileTitle
    End If

End Sub

Private Sub Form_Load()
 X = App.Path
 Image2.Picture = LoadPicture(X & "\" & Label7.Caption)
 Command4.Enabled = True
 Command5.Enabled = False
 Text1.Enabled = False
 Text2.Enabled = False
 Text3.Enabled = False
 Text4.Enabled = False
 Text5.Enabled = False
 Command7.Enabled = False

End Sub
