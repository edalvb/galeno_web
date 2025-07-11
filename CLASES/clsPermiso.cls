VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPermiso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim Varnn As Variant
Dim rsPermisos As ADODB.Recordset
Dim strConsulta As String
Dim collModulos As Collection
Dim collPermisos As Collection
Dim blAgregarPermiso As Boolean
Dim Comando As ADODB.Command
Dim Parametros As ADODB.Parameter

Public Sub addUno(ByRef Fuente As ListBox, ByRef Receptor As ListBox)
blAgregarPermiso = True
Agregar Fuente, Receptor

End Sub

Public Sub addTodo(Fuente As ListBox, Receptor As ListBox)
blAgregarPermiso = True
AgregarTodo Fuente, Receptor

End Sub

Public Sub delUno(Fuente As ListBox, Receptor As ListBox)
 blAgregarPermiso = False
 Agregar Receptor, Fuente
 
End Sub

Public Sub delTodo(Fuente As ListBox, Receptor As ListBox)
blAgregarPermiso = False
AgregarTodo Receptor, Fuente
End Sub

Public Sub CargarPermisos(Fuente As ListBox, Destino As ListBox, ConSglac As Connection, strUserId As String)
On Error GoTo errManEJAR
If strUserId <> "" Then
    If ConSglac.State = adStateOpen Then
    '    Set rsPermisos = New Recordset
        Fuente.Clear: Destino.Clear
        If Not (collModulos Is Nothing) Then
            Set collPermisos = Nothing
            Set collModulos = Nothing
        End If
        Set collPermisos = New Collection
        Set collModulos = New Collection
        ArmarConsultaModulos strUserId
            
        Set rsPermisos = ConSglac.Execute(strConsulta, , adCmdText)
        blAgregarPermiso = False
        cargarLista Fuente
        ArmarConsultaPermisos strUserId
        Set rsPermisos = ConSglac.Execute(strConsulta, , adCmdText)
        blAgregarPermiso = True
        cargarLista Destino
        On Error Resume Next
        rsPermisos.Close
        Set rsPermisos = Nothing
    End If
        
        
End If
Exit Sub
errManEJAR:

End Sub


Private Sub cargarLista(lstFuente As ListBox)
Dim i As Integer
'On Error GoTo errManEJAR
    If rsPermisos.State = adStateOpen Then
        
        
        Do While Not rsPermisos.EOF
            lstFuente.AddItem rsPermisos(1).Value, 0
            If blAgregarPermiso Then
                collPermisos.Add ("'" & rsPermisos(0).Value)
            Else
                collModulos.Add rsPermisos(0).Value
            End If
            rsPermisos.MoveNext
        
        Loop
        
        rsPermisos.Close
    
    End If
    
Exit Sub
errManEJAR:
End Sub

Private Sub Agregar(Fuente As ListBox, Receptor As ListBox)

If Fuente.ListIndex > -1 Then
    Receptor.AddItem Fuente.List(Fuente.ListIndex), 0
    
    If blAgregarPermiso Then
        collPermisos.Add collModulos.Item(collModulos.Count - (Fuente.ListIndex))
        collModulos.Remove (collModulos.Count - (Fuente.ListIndex))
    
    Else
        collModulos.Add collPermisos.Item(collPermisos.Count - (Fuente.ListIndex))
        collPermisos.Remove (collPermisos.Count - (Fuente.ListIndex))
    
    End If
    
    Fuente.RemoveItem Fuente.ListIndex
End If
End Sub

Private Sub AgregarTodo(Fuente As ListBox, Destino As ListBox)
Dim i As Integer
Dim n As Integer
Dim h As Integer
If blAgregarPermiso Then
    h = collModulos.Count
Else
    h = collPermisos.Count
End If
n = Fuente.ListCount
    For i = 0 To n - 1
        Destino.AddItem Fuente.List(i), 0
        
        If blAgregarPermiso Then
            collPermisos.Add collModulos.Item(h - (i))
            collModulos.Remove h - (i)
        Else
            collModulos.Add collPermisos.Item(h - (i))
            collPermisos.Remove h - (i)
        End If
        
        
      
    Next
    Fuente.Clear
End Sub

Private Sub ArmarConsultaModulos(UserId As String)
strConsulta = "SELECT modulos.CODMODULO, modulos.DESCRIPCION " & _
              "From modulos WHERE  (modulos.MOSTRAR)=True and  " & _
              "exists(select  HABILITACION.CODMODULO from HABILITACION " & _
              " where HABILITACION.id_usuario=" & UserId & _
              " and HABILITACION.CODMODULO=modulos.CODMODULO)=false;"

End Sub

Private Sub ArmarConsultaPermisos(UserId As String)
strConsulta = "SELECT HABILITACION.CODMODULO, MODULOS.DESCRIPCION" & _
            " FROM MODULOS INNER JOIN HABILITACION ON MODULOS.CODMODULO = " & _
            "HABILITACION.CODMODULO " & _
            "WHERE (((HABILITACION.ID_USUARIO)=" & UserId & "));"

End Sub


Public Sub Grabar(ConSglac As ADODB.Connection, Usuario As String, Fuente As ListBox, Destino As ListBox)
If ConSglac.State = adStateOpen Then
    EliminarPermiso ConSglac, Usuario
    AgregarPermiso ConSglac, Usuario
End If
  
End Sub


Private Sub EliminarPermiso(ConSglacE As ADODB.Connection, UserId As String)
Dim oModulos As Variant
Dim collAux As Collection
Set collAux = New Collection
For Each oModulos In collModulos
    If Left$(oModulos, 1) = "'" Then
        ConSglacE.Execute "delete from habilitacion where id_usuario=" & UserId & _
                          "and codmodulo='" & Right$(oModulos, (Len(oModulos) - 1)) _
                          & "'", , adCmdText
        collAux.Add Right$(oModulos, (Len(oModulos) - 1))
    Else
        collAux.Add oModulos
    End If

Next
Set collModulos = collAux
End Sub

Private Sub AgregarPermiso(ConSglacA As ADODB.Connection, usrId)
Dim oPermiso As Variant
Dim collAux As Collection
Set collAux = New Collection
For Each oPermiso In collPermisos
    If Left$(oPermiso, 1) <> "'" Then
        ConSglacA.Execute "insert into habilitacion (id_usuario,codmodulo) " & _
                  "values (" & usrId & " , '" & oPermiso & "')", , adCmdText
        collAux.Add ("'" & oPermiso)

    Else
        collAux.Add oPermiso
            
    End If
Next
Set collPermisos = collAux
End Sub

Private Sub Class_Initialize()
    Set collPermisos = New Collection
    Set collModulos = New Collection

End Sub




