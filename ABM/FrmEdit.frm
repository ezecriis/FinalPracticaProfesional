VERSION 5.00
Begin VB.Form FrmEdit 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agregar registro"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4515
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   4515
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CmbSexo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "FrmEdit.frx":0000
      Left            =   1800
      List            =   "FrmEdit.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3480
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   0
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1800
      TabIndex        =   1
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1800
      TabIndex        =   2
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1800
      TabIndex        =   3
      Top             =   3000
      Width           =   2415
   End
   Begin VB.PictureBox cmdCancelar 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      Picture         =   "FrmEdit.frx":0023
      ScaleHeight     =   435
      ScaleWidth      =   1275
      TabIndex        =   7
      Top             =   4320
      Width           =   1335
   End
   Begin VB.PictureBox cmdGuardar 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      Picture         =   "FrmEdit.frx":05AD
      ScaleHeight     =   435
      ScaleWidth      =   1995
      TabIndex        =   6
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label lblID 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   270
      Left            =   1800
      TabIndex        =   15
      Top             =   720
      Width           =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Id de registro :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   14
      Top             =   720
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   13
      Top             =   1560
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   12
      Top             =   2040
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   3
      Left            =   360
      TabIndex        =   11
      Top             =   2520
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sexo"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   5
      Left            =   360
      TabIndex        =   10
      Top             =   3480
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Domicilio"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   4
      Left            =   360
      TabIndex        =   9
      Top             =   3000
      Width           =   630
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   3  'Dot
      Index           =   0
      X1              =   360
      X2              =   4200
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de alta :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   6
      Left            =   360
      TabIndex        =   8
      Top             =   240
      Width           =   1305
   End
   Begin VB.Label lblFecha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de alta"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   1800
      TabIndex        =   4
      Top             =   240
      Width           =   1365
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   3  'Dot
      Index           =   1
      X1              =   360
      X2              =   4200
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   3  'Dot
      Height          =   4935
      Left            =   120
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "FrmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Enum EACCION
    AGREGAR_REGISTRO = 0
    EDITAR_REGISTRO = 1
End Enum

Public IdRegistro
Public ACCION As EACCION



Private Sub cmdGuardar_Click()

On Error GoTo ErrorSub
    
    
    ' Valida el Nombre que no este vacio
    ''''''''''''''''''''''''''''''''
    If Trim(Text1(1)) = "" Then
        MsgBox "El Nombre de registro no puede estar vacio", vbCritical, "Datos incompletos"
        Text1(1).SetFocus
        Exit Sub
    
    ' Valida el Apellido
    ''''''''''''''''''''''''''''''''
    ElseIf Trim(Text1(2)) = "" Then
        MsgBox "El Apellido no puede estar vacio", vbCritical, "Datos incompletos"
        Text1(2).SetFocus
        Exit Sub
    
    ' Valida el Sexo
    ''''''''''''''''''''''''''''''''

    ElseIf Trim(CmbSexo.Text) = "" Then
        MsgBox "No se ha indicado el sexo", vbCritical, "Datos incompletos"
        CmbSexo.SetFocus
        Exit Sub

    End If

    'Agrega el registro
    '''''''''''''''''''''''''''''''
    
    Select Case ACCION
    Case EDITAR_REGISTRO
        cnn.Execute "UPDATE Personas set Nombre = '" & Text1(1) & _
                                         "', Apellido = '" & Text1(2) & _
                                         "', Telefono = '" & Text1(3) & _
                                         "', Direccion = '" & Text1(4) & _
                                         "', Sexo = '" & CmbSexo.ListIndex & _
                                         "' where Id = " & IdRegistro & ""
    Case AGREGAR_REGISTRO
        
        cnn.Execute "INSERT INTO Personas " & "(Nombre,Apellido,Telefono,Direccion,Sexo,FechaDeAlta) VALUES('" & _
                                 Text1(1) & "','" & _
                                 Text1(2) & "','" & _
                                 Text1(3) & "','" & _
                                 Text1(4) & "','" & _
                                 CmbSexo.ListIndex & "','" & _
                                 Format(Date, "dd/mm/yyyy") & "')"

    End Select
    
    rs.Requery 1
    
    Call CargarListView(FrmPrincipal.LV, rs)

    DoEvents
    Unload Me
    Set FrmEdit = Nothing
Exit Sub
ErrorSub:
MsgBox Err.Description

End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
       Unload Me
    End If
End Sub

