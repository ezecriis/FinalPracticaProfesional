VERSION 5.00
Begin VB.Form frmFilter 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Filtrar y ordenar"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox ChameleonBtn1 
      Height          =   375
      Left            =   5160
      Picture         =   "frmFilter.frx":0000
      ScaleHeight     =   315
      ScaleWidth      =   915
      TabIndex        =   8
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtSearch 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmFilter.frx":058A
      Left            =   3960
      List            =   "frmFilter.frx":0597
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Font Size"
      Top             =   240
      Width           =   2055
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmFilter.frx":05B1
      Left            =   1320
      List            =   "frmFilter.frx":05C1
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Font Size"
      Top             =   1080
      Width           =   2055
   End
   Begin VB.PictureBox CmdOrdenar 
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
      Index           =   0
      Left            =   3720
      Picture         =   "frmFilter.frx":05EA
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   3
      Top             =   960
      Width           =   375
   End
   Begin VB.PictureBox CmdOrdenar 
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
      Index           =   1
      Left            =   4200
      Picture         =   "frmFilter.frx":0B74
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   4
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filtrar"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   7
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Por el campo"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   1
      Left            =   2880
      TabIndex        =   6
      Top             =   240
      Width           =   930
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   3  'Dot
      Index           =   1
      X1              =   3600
      X2              =   3600
      Y1              =   1440
      Y2              =   960
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ordenar por"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   5
      Top             =   1080
      Width           =   840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   3  'Dot
      Index           =   3
      X1              =   360
      X2              =   6240
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   3  'Dot
      Height          =   1455
      Left            =   120
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "frmFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ChameleonBtn1_Click()
    Unload Me
End Sub

' Ordena en forma Ascendente y descendente el LV
''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CmdOrdenar_Click(Index As Integer)
    CmdOrdenar(0).Value = False
    CmdOrdenar(1).Value = False
    CmdOrdenar(Index).Value = True
    Call Filtrar
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
       Unload Me
    End If
End Sub

Private Sub Form_Load()
    With FrmPrincipal
        Me.Move (.Left + .LV.Left), _
                (.LV.Height + .LV.Top + .Top + 500)
    End With
    Call Filtrar
End Sub

Private Sub txtSearch_Change()
    Call Filtrar
End Sub

Private Sub Combo1_Click()
    Call Filtrar
End Sub

Private Sub Combo2_Click()
    Call Filtrar
End Sub

Public Sub Filtrar()
Dim Campo, OrderByCampo, Orden As String
Dim SQL As String

    If Combo1.ListIndex = -1 Then
        Combo1.ListIndex = 0
    End If
    If Combo2.ListIndex = -1 Then
        Combo2.ListIndex = 0
    End If
    If Combo1.ListIndex = 0 Then
        Campo = "Id"
    ElseIf Combo1.ListIndex = 1 Then
        Campo = "Nombre"
    ElseIf Combo1.ListIndex = 2 Then
        Campo = "Apellido"
    End If
    
    Select Case Combo2.ListIndex
        Case 0: OrderByCampo = "Id"
        Case 1: OrderByCampo = "Nombre"
        Case 2: OrderByCampo = "Apellido"
        Case 3: OrderByCampo = "FechaDeAlta"
    End Select

    If CmdOrdenar(0).Value Then Orden = "asc"
    If CmdOrdenar(1).Value Then Orden = "desc"

    ' si el recorset está abierto lo cierra
    If rs.State = adStateOpen Then
        rs.Close
    End If
    
    SQL = "SELECT * FROM Personas Where " & _
                         Campo & " like '" & txtSearch & _
                        "%' order by " & OrderByCampo & " " & Orden
    
    rs.Open SQL, cnn, adOpenStatic, adLockOptimistic
    
    Call CargarListView(FrmPrincipal.LV, rs)

End Sub



