VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   8745
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16065
   LinkTopic       =   "Form2"
   ScaleHeight     =   8745
   ScaleWidth      =   16065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "regresar"
      Height          =   615
      Left            =   9120
      TabIndex        =   19
      Top             =   7920
      Width           =   2055
   End
   Begin VB.CommandButton Command7 
      Caption         =   "foto"
      Height          =   495
      Left            =   5400
      TabIndex        =   18
      Top             =   5760
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9120
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   "MODIDIFICAR"
      Height          =   735
      Left            =   13680
      TabIndex        =   17
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "GUARDAR"
      Height          =   615
      Left            =   13680
      TabIndex        =   16
      Top             =   5760
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ELIMINAR"
      Height          =   615
      Left            =   13680
      TabIndex        =   15
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "NUEVO"
      Height          =   735
      Left            =   13680
      TabIndex        =   14
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<"
      Height          =   855
      Left            =   5400
      TabIndex        =   13
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   855
      Left            =   5400
      TabIndex        =   12
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "FOTOGRAFIA"
      Height          =   2535
      Left            =   240
      TabIndex        =   11
      Top             =   4560
      Width           =   4815
      Begin VB.Image Image2 
         Height          =   2445
         Left            =   0
         Picture         =   "Form2.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   4620
      End
   End
   Begin VB.TextBox Text5 
      DataField       =   "sueldo"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2160
      TabIndex        =   10
      Top             =   3720
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      DataField       =   "cargo"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2160
      TabIndex        =   9
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      DataField       =   "apellido"
      DataSource      =   "Adodc1"
      Height          =   525
      Left            =   2160
      TabIndex        =   8
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      DataField       =   "nombre"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2160
      TabIndex        =   7
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      DataField       =   "Id"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   840
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   6000
      Top             =   240
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1296
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
      Connect         =   $"Form2.frx":87CF
      OLEDBString     =   $"Form2.frx":887B
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "empleados"
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
   Begin VB.Label Label6 
      DataField       =   "foto"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   8160
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "sueldo"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "cargo"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "apellido"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "nombre"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "id"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   9465
      Left            =   0
      Picture         =   "Form2.frx":8927
      Stretch         =   -1  'True
      Top             =   120
      Width           =   16020
   End
End
Attribute VB_Name = "Form2"
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

Private Sub Command8_Click()
Form1.Show
Form2.Hide

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

Private Sub Label6_Click()
 MsgBox ("debe seleccionar una imagen")
    Else
     Label6.Caption = CommonDialog1.FileTitle
    End If
End Sub
