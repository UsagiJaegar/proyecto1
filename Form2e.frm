VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6150
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8445
   LinkTopic       =   "Form2"
   ScaleHeight     =   6150
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command7 
      Caption         =   "Foto"
      Height          =   375
      Left            =   6480
      TabIndex        =   26
      Top             =   4920
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5400
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   "modificar"
      Height          =   375
      Left            =   7440
      TabIndex        =   25
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "delete"
      Height          =   375
      Left            =   6720
      TabIndex        =   24
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "new"
      Height          =   375
      Left            =   6000
      TabIndex        =   23
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "save"
      Height          =   375
      Left            =   5280
      TabIndex        =   22
      Top             =   4200
      Width           =   615
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   240
      Top             =   5520
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\user\Desktop\proyecto1\zoologico.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\user\Desktop\proyecto1\zoologico.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "animales"
      Caption         =   ""
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
   Begin VB.CommandButton Command2 
      Caption         =   ">"
      Height          =   495
      Left            =   1320
      TabIndex        =   21
      Top             =   4800
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      Height          =   495
      Left            =   360
      TabIndex        =   20
      Top             =   4800
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      Caption         =   "fotografia"
      Height          =   1815
      Left            =   2640
      TabIndex        =   19
      Top             =   4080
      Width           =   2535
      Begin VB.Image Image2 
         DataSource      =   "Adodc2"
         Height          =   1455
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.TextBox Text8 
      DataField       =   "lugar_origen"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   6240
      TabIndex        =   17
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      DataField       =   "edad"
      DataSource      =   "Adodc2"
      Height          =   405
      Left            =   6240
      TabIndex        =   14
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      DataField       =   "tipo_alimentacion"
      DataSource      =   "Adodc2"
      Height          =   405
      Left            =   6240
      TabIndex        =   12
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      DataField       =   "peso"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   6240
      TabIndex        =   11
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      DataField       =   "cantidad"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      DataField       =   "especies"
      DataSource      =   "Adodc2"
      Height          =   405
      Left            =   2160
      TabIndex        =   6
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      DataField       =   "nombre"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      DataField       =   "codigo"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label11 
      DataField       =   "foto"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   6120
      TabIndex        =   18
      Top             =   5520
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "lugar de origen"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4440
      TabIndex        =   16
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   " lb"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7440
      TabIndex        =   15
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "edad"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4440
      TabIndex        =   13
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "alimentacion"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "peso"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4440
      TabIndex        =   9
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "cantidad"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "especies"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "nombre"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      Caption         =   "Animales"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   6975
   End
   Begin VB.Image Image1 
      Height          =   6060
      Left            =   0
      Picture         =   "Form2e.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8520
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Adodc2.Recordset.MovePrevious
    
    If Adodc2.Recordset.BOF Then
        Adodc2.Recordset.MoveLast
    End If
    
    x = App.Path
   Image2.Picture = LoadPicture(x & "\" & Label11.Caption)
        
End Sub

Private Sub Command2_Click()
    Adodc2.Recordset.MoveNext
    
    If Adodc2.Recordset.EOF Then
        Adodc2.Recordset.MoveFirst
    End If
    
    x = App.Path
    Image2.Picture = LoadPicture(x & "\" & Label11.Caption)
    
End Sub

Private Sub Command3_Click()
    FileCopy CommonDialog1.FileName, App.Path & "\\" & CommonDialog1.FileTitle
    Adodc2.Recordset.Update
    Adodc2.Recordset.MoveFirst
    x = App.Path
    Image2.Picture = LoadPicture(x & "\" & Label11.Caption)
    
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    Text5.Enabled = False
    Text6.Enabled = False
    Text7.Enabled = False
    Text8.Enabled = False
    Command3.Enabled = False
    Command1.Enabled = True
    Command2.Enabled = True
    Command4.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = True
    Command7.Enabled = False
    
End Sub

Private Sub Command4_Click()
    Adodc2.Recordset.AddNew
    
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    Text5.Enabled = True
    Text6.Enabled = True
    Text7.Enabled = True
    Text8.Enabled = True
    Command3.Enabled = True
    Command1.Enabled = False
    Command2.Enabled = False
    Command4.Enabled = False
    Command5.Enabled = False
    Command6.Enabled = False
    Command7.Enabled = True
    
    Text1.SetFocus
    
    Label11.Caption = ""
    Image2.Picture = LoadPicture(Label11.Caption)
    
End Sub

Private Sub Command5_Click()
    Adodc2.Recordset.Delete
    Adodc2.Recordset.MoveFirst
    x = App.Path
    Image2.Picture = LoadPicture(x & "\" & Label11.Caption)
End Sub

Private Sub Command6_Click()
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    Text5.Enabled = True
    Text6.Enabled = True
    Text7.Enabled = True
    Text8.Enabled = True
    Command3.Enabled = True
    Command1.Enabled = False
    Command2.Enabled = False
    Command4.Enabled = False
    Command5.Enabled = False
    Command6.Enabled = False
End Sub

Private Sub Command7_Click()
    CommonDialog1.ShowOpen
    Image2.Picture = LoadPicture(CommonDialog1.FileName)
    Label11.Caption = CommonDialog1.FileTitle
    
    If Label11.Caption = "" Then
        MsgBox ("seleccione una imagen")
    Else
         Label11.Caption = CommonDialog1.FileTitle
    End If
End Sub

Private Sub Form_Load()
    x = App.Path
    Image2.Picture = LoadPicture(x & "\" & Label11.Caption)
    
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
    Command5.Enabled = True
    Command6.Enabled = True
    Command7.Enabled = False
    
    

End Sub

