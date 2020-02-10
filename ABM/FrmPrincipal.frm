VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPrincipal 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registros"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   11250
   Icon            =   "FrmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   11250
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView LV 
      Height          =   5535
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   9763
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Apellido"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Teléfono"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Domicilio"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Sexo"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Fecha de alta"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.PictureBox ctxHookMenu1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2520
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   7
      Top             =   5760
      Width           =   1200
   End
   Begin VB.PictureBox cmdOpciones 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   120
      Picture         =   "FrmPrincipal.frx":038A
      ScaleHeight     =   555
      ScaleWidth      =   1155
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox cmdOpciones 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   120
      Picture         =   "FrmPrincipal.frx":0714
      ScaleHeight     =   555
      ScaleWidth      =   1155
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.PictureBox cmdOpciones 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   120
      Picture         =   "FrmPrincipal.frx":0C9E
      ScaleHeight     =   555
      ScaleWidth      =   1155
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.PictureBox cmdOpciones 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   120
      Picture         =   "FrmPrincipal.frx":1028
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   4
      Top             =   5040
      Width           =   1215
   End
   Begin VB.PictureBox cmdOpciones 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   120
      Picture         =   "FrmPrincipal.frx":15B2
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
   End
   Begin VB.PictureBox cmdOpciones 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   120
      Picture         =   "FrmPrincipal.frx":1B3C
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   6
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuImprimir 
         Caption         =   "Imprimir"
         Shortcut        =   ^P
      End
      Begin VB.Menu l1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuEdicion 
      Caption         =   "Edición"
      Begin VB.Menu mnuAgregar 
         Caption         =   "Agregar registro"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuEditarRegistro 
         Caption         =   "Editar Registro"
         Shortcut        =   ^E
      End
      Begin VB.Menu l5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEliminarReg 
         Caption         =   "Eliminar este registro"
      End
   End
   Begin VB.Menu mnuOpciones 
      Caption         =   "&Opciones"
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "FrmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' Botones de opción
''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOpciones_Click(Index As Integer)
    Select Case Index
        Case 0: Call Agregar
        Case 1: Call Editar
        Case 2: Call Eliminar
        Case 3: Unload Me
        Case 4: frmFilter.Show , Me
        Case 5: Call mnuImprimir_Click
    End Select
End Sub


'Abre el formulario para Editar el registro seleccionado en el ListView
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Editar()

    Dim i As Integer
    
    ' verifica que hay datos en el ListView y que hay uno seleccionado
    If (LV.ListItems.Count = 0) Then
       MsgBox "No hay ningún regisro para editar", vbInformation
       Exit Sub
    End If
    If (LV.SelectedItem Is Nothing) Then
       MsgBox "Debe seleccionar previamente un registro para poder editarlo", vbInformation
       Exit Sub
    End If
    
    With FrmEdit
        ' obtiene el elemento seleccionado
        .lblID = LV.SelectedItem.Text
        For i = 1 To 4
            .Text1(i).Text = LV.SelectedItem.ListSubItems(i).Text
        Next
        .CmbSexo = LV.SelectedItem.ListSubItems(5).Text
        .lblFecha = LV.SelectedItem.ListSubItems(6).Text
        .IdRegistro = LV.SelectedItem.Text
        .ACCION = EDITAR_REGISTRO
        
        .Show vbModal
    End With

End Sub

' Elimina el registro actual seleccionado
'''''''''''''''''''''''''''''''''''''''''''''

Private Sub Eliminar()

    
    
    If (LV.ListItems.Count = 0) Then
        MsgBox "No hay ningún registro para eliminar", vbInformation
        Exit Sub
    End If
    
    ' verifica que hay datos en el ListView y que hay uno seleccionado
    If (LV.SelectedItem Is Nothing) Then
        MsgBox "No hay registro seleccionado para eliminar", vbInformation
        Exit Sub
    End If
    
    
    With LV.SelectedItem
        ' pregunta
        If MsgBox("Se va a eliminar el registro : " & vbNewLine & _
                 String(50, "-") & vbNewLine & _
                 "ID: " & .Text & vbNewLine & _
                 "Nombre " & .ListSubItems(1).Text & vbNewLine & _
                 "Apellido: " & .ListSubItems(2).Text, _
                 vbExclamation + vbYesNo, "Eliminar") = vbYes Then
            ' Elimina
            cnn.Execute "delete from Personas where Id = " & .Text & ""
            ' refresca el recordset
            rs.Requery 1
            ' vuelve a cargar los datos en el ListView
            Call CargarListView(LV, rs)
        End If
    End With
End Sub


Sub Agregar()
    
    ' Acción
    FrmEdit.ACCION = AGREGAR_REGISTRO
    
    FrmEdit.lblFecha = Format(Date, "mm/dd/yyyy")
    ' Abre el Form
    FrmEdit.Show 1
End Sub

Sub Salir()
    Call Desconectar
    Unload Me
    End
End Sub


Private Sub Form_Load()
    ' Abre la conexión
    Call IniciarConexion
    ' carga el Recorset con todos los datos
    rs.Open "select * from Personas", cnn, adOpenStatic, adLockOptimistic
    ' llena el ListView
    Call CargarListView(LV, rs)

End Sub


Private Sub LV_DblClick()
    Call Editar
End Sub



Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Item As ListItem
    
    Set Item = LV.HitTest(x, y)
    
    If Not Item Is Nothing And Button = vbRightButton Then
       Item.Selected = True
       Me.PopupMenu mnuEdicion
    End If
End Sub

' menues
'''''''''''''''''''''''''''''

Private Sub mnuAgregar_Click()
    Call Agregar
End Sub

Private Sub mnuEditarRegistro_Click()
    Call Editar
End Sub

Private Sub mnuEliminarReg_Click()
    Call Eliminar
End Sub

Private Sub mnuImprimir_Click()
    Set DataReport1.DataSource = rs
    DataReport1.Show 1
End Sub

' salir

''''''''''''''''''''''''
Private Sub mnuSalir_Click()
    Call Salir
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim ret As VbMsgBoxResult
    
    ret = MsgBox("¿ Salir ?", vbInformation + vbYesNo)
    If ret = vbNo Then
        Cancel = True
    Else
        Call Salir
    End If
End Sub

