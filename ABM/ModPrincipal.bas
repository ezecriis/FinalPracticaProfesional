Attribute VB_Name = "ModPrincipal"
Option Explicit

Public Declare Sub InitCommonControls Lib "comctl32" ()

' variables para la conexión y el recordset
''''''''''''''''''''''''''''''''''''''''''''
Public cnn As New ADODB.Connection
Public rs As New ADODB.Recordset

Public ObjItem As ListItem


Sub Main()
    On Error Resume Next
    Call InitCommonControls
    Err.Clear
    FrmPrincipal.Show
End Sub

' abre
Public Sub IniciarConexion()

    With cnn
        .CursorLocation = adUseClient
        .Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
              App.Path & "\datos.mdb" & ";Persist Security Info=False"
    End With

End Sub


Public Sub CargarListView(LV As ListView, rs As ADODB.Recordset)
    
    On Error GoTo ErrorSub
    
    Dim i As Integer
    'limpia el LV
    LV.ListItems.Clear
    
    ' si hay registros
    If rs.RecordCount > 0 Then
        
        ' recorre el recordset
        While Not rs.EOF
            ' añade los datos
            Set ObjItem = LV.ListItems.Add(, , rs(0))
                
           
           ObjItem.SubItems(1) = rs!Nombre
           ObjItem.SubItems(2) = rs!Apellido
           ObjItem.SubItems(3) = rs!Telefono
           ObjItem.SubItems(4) = rs!Direccion
           If Abs(rs!sexo) = 0 Then
              ObjItem.SubItems(5) = "Masculino"
           Else
              ObjItem.SubItems(5) = "Femenino"
           End If
           ObjItem.SubItems(6) = rs!FechaDeAlta
           
            ' siguiente registro
            rs.MoveNext
        Wend
        
    End If
    Call ForeColorColumn(&H8000&, 0, FrmPrincipal.LV)
    'Call ForeColorColumn(vbRed, 6, FrmPrincipal.LV)
    
    Exit Sub
    
ErrorSub:
    
    If Err.Number = 94 Then Resume Next
    
End Sub


' cierra
Sub Desconectar()
    On Local Error Resume Next
    rs.Close
    Set rs = Nothing
    cnn.Close
    Set cnn = Nothing
End Sub

