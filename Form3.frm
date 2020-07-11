VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   8760
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15975
   LinkTopic       =   "Form3"
   ScaleHeight     =   8760
   ScaleWidth      =   15975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "REGRESAR"
      Height          =   855
      Left            =   7680
      TabIndex        =   19
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Caption         =   "FOTO"
      Height          =   255
      Left            =   5400
      TabIndex        =   18
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "MODIFICAR"
      Height          =   615
      Left            =   10080
      TabIndex        =   17
      Top             =   6600
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "GUARDAR"
      Height          =   615
      Left            =   10080
      TabIndex        =   16
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ELIMINAR"
      Height          =   615
      Left            =   10080
      TabIndex        =   15
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "NUEVO"
      Height          =   735
      Left            =   10080
      TabIndex        =   14
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "< "
      Height          =   855
      Left            =   5280
      TabIndex        =   13
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   735
      Left            =   5280
      TabIndex        =   12
      Top             =   4440
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9000
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   5760
      Top             =   480
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
      Connect         =   $"Form3.frx":0000
      OLEDBString     =   $"Form3.frx":00AC
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "clientes"
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
      DataField       =   "total a pagar"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2400
      TabIndex        =   10
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      DataField       =   "fecha de consulta"
      DataSource      =   "Adodc1"
      Height          =   525
      Left            =   2400
      TabIndex        =   9
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      DataField       =   "enfermedad"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2400
      TabIndex        =   8
      Text            =   " "
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      DataField       =   "nombre del paciente"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2400
      TabIndex        =   7
      Text            =   " "
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      DataField       =   "cui"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Text            =   " "
      Top             =   360
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3255
      Left            =   240
      TabIndex        =   5
      Top             =   4200
      Width           =   4695
      Begin VB.Image Image2 
         Height          =   2880
         Left            =   120
         Picture         =   "Form3.frx":0158
         Stretch         =   -1  'True
         Top             =   240
         Width           =   4380
      End
   End
   Begin VB.Label Label1 
      Caption         =   "cui"
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label6 
      Caption         =   "foto"
      DataField       =   "foto"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   8160
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "total a pagar"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "fecha de consulta"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "enfermedad"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "nombre del paciente"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   9645
      Left            =   0
      Picture         =   "Form3.frx":8766
      Stretch         =   -1  'True
      Top             =   120
      Width           =   16140
   End
End
Attribute VB_Name = "Form3"
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
Image2.Picture = LoadPicture(X & "\" & Label6.Caption)

End Sub

Private Sub Command2_Click()
Adodc1.Recordset.MovePrevious
 
 If Adodc1.Recordset.BOF Then
    Adodc1.Recordset.MoveLast
 End If
 X = App.Path
    Image2.Picture = LoadPicture(X & "\" & Label6.Caption)


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
Label6.Caption = ""
Image2.Picture = LoadPicture(Label6.Caption)
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
 Image2.Picture = LoadPicture(X & "\" & Label6.Caption)

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
    Label6.Caption = CommonDialog1.FileTitle
    
    If Label6.Caption = "" Then
        MsgBox ("debe seleccionar una imagen")
    Else
     Label6.Caption = CommonDialog1.FileTitle
    End If
   

End Sub

Private Sub Form_Load()
 X = App.Path
 Image2.Picture = LoadPicture(X & "\" & Label6.Caption)
 Command4.Enabled = True
 Command5.Enabled = False
 Text1.Enabled = False
 Text2.Enabled = False
 Text3.Enabled = False
 Text4.Enabled = False
 Text5.Enabled = False
 Command7.Enabled = False

End Sub
