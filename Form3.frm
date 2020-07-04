VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   5805
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9855
   LinkTopic       =   "Form3"
   ScaleHeight     =   5805
   ScaleWidth      =   9855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "eliminar"
      Height          =   375
      Left            =   6120
      TabIndex        =   18
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "modificar"
      Height          =   375
      Left            =   4920
      TabIndex        =   17
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "guardar"
      Height          =   375
      Left            =   6120
      TabIndex        =   16
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "nuevo"
      Height          =   375
      Left            =   4920
      TabIndex        =   15
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">"
      Height          =   375
      Left            =   6120
      TabIndex        =   14
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      Height          =   375
      Left            =   5040
      TabIndex        =   13
      Top             =   3480
      Width           =   735
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   600
      Top             =   5160
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\user\Desktop\acces\zoologico.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\user\Desktop\acces\zoologico.mdb;Persist Security Info=False"
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
   Begin VB.TextBox Text6 
      DataField       =   "fecha_inicio"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   2520
      TabIndex        =   12
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      DataField       =   "sueldo"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   2520
      TabIndex        =   11
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      DataField       =   "cargo"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   2520
      TabIndex        =   10
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      DataField       =   "edad"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   2520
      TabIndex        =   9
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      DataField       =   "nombre_completo"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   2520
      TabIndex        =   8
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      DataField       =   "CUI"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   2520
      TabIndex        =   7
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "fecha de inicio"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "sueldo"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "cargo"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "edad"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "nombre"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "CUI"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      Caption         =   "empleados"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   6735
   End
   Begin VB.Image Image1 
      Height          =   5700
      Left            =   0
      Picture         =   "Form3.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9840
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Adodc1.Recordset.MovePrevious
    If Adodc1.Recordset.BOF Then
    Adodc1.Recordset.MoveLast
    
End If
    
End Sub

Private Sub Command2_Click()
    Adodc1.Recordset.MoveNext
    If Adodc1.Recordset.EOF Then
    Adodc1.Recordset.MoveFirst
    
End If
End Sub

Private Sub Command3_Click()
    Adodc1.Recordset.AddNew
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    Text5.Enabled = True
    Text6.Enabled = True
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = True
    Command5.Enabled = False
    Command6.Enabled = False
    Text1.SetFocus
End Sub

Private Sub Command4_Click()
    Adodc1.Recordset.Update
    Adodc1.Recordset.MoveFirst
End Sub

Private Sub Command5_Click()
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    Text5.Enabled = True
    Text6.Enabled = True
    comand1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = True
    Command5.Enabled = False
    Command6.Enabled = False
End Sub

Private Sub Command6_Click()
    Adodc1.Recordset.Delete
    Adodc1.Recordset.MoveFirst
    
End Sub

Private Sub Form_Load()
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    Text5.Enabled = False
    Text6.Enabled = False
    Text7.Enabled = False
    Text8.Enabled = False
    Command1.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = False
    Command4.Enabled = True
    Command5.Enabled = False
    Command6.Enabled = False
    
    
End Sub
