VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsValidar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim rsPermiso As ADODB.Recordset
Dim Coneccion As ADODB.Connection
Dim intIndice As Integer
Dim i As Integer
Public Event ABMtollbar(intABM As Integer)
Public Event BUSCAR(enable As Boolean)
Dim strConsulta As String
Dim blBuscar As Boolean
Public Function Visualizar(Index As Integer, conexion As ADODB.Connection, usuario As String) As Boolean
Dim blVisualizar As Boolean
Set Coneccion = conexion

Select Case Index
    Case 1 'paciente
        strConsulta = strConsulta & "1.01.01' and id_usuario=" & usuario
        blVisualizar = TienePermiso(1)
    Case 2 'personal
        strConsulta = strConsulta & "1.01.02' and id_usuario=" & usuario
        blVisualizar = TienePermiso(2)
    Case 3 'lineas
        
        strConsulta = strConsulta & "1.01.03' and id_usuario=" & usuario
        blVisualizar = TienePermiso(3)
    Case 4 'insumos
        strConsulta = strConsulta & "1.01.04' and id_usuario=" & usuario
        blVisualizar = TienePermiso(4)
    Case 5 'doctor
        strConsulta = strConsulta & "1.01.05' and id_usuario=" & usuario
        blVisualizar = TienePermiso(5)
    Case 6 'area
        strConsulta = strConsulta & "1.01.06' and id_usuario=" & usuario
        blVisualizar = TienePermiso(6)
    Case 7 'especialidad
        strConsulta = strConsulta & "1.01.07' and id_usuario=" & usuario
        blVisualizar = TienePermiso(7)
    Case 8 'valores normales
        strConsulta = strConsulta & "1.01.08' and id_usuario=" & usuario
        blVisualizar = TienePermiso(8)
    Case 9 'instituto
        strConsulta = strConsulta & "1.01.09' and id_usuario=" & usuario
        blVisualizar = TienePermiso(9)
    Case 10 'aseguradora
        strConsulta = strConsulta & "1.01.10' and id_usuario=" & usuario
        blVisualizar = TienePermiso(10)
    Case 11 'usuarios
        strConsulta = strConsulta & "1.01.11' and id_usuario=" & usuario
        blVisualizar = TienePermiso(11)
    Case 12 'proveedores
        strConsulta = strConsulta & "1.01.12' and id_usuario=" & usuario
        blVisualizar = TienePermiso(12)
    Case 13 'ciudades
        strConsulta = strConsulta & "1.01.13' and id_usuario=" & usuario
        blVisualizar = TienePermiso(13)
    Case 14 'tabla de valores
        strConsulta = strConsulta & "1.01.14' and id_usuario=" & usuario
        blVisualizar = TienePermiso(14)

    Case 20 'pedidos
        strConsulta = strConsulta & "1.02.1' and id_usuario=" & usuario
        blVisualizar = TienePermiso(20)

    Case 15 'ATENCION AL PACIENTE
        strConsulta = strConsulta & "1.03.01' and id_usuario=" & usuario
        blVisualizar = TienePermiso(20)
   
    Case 16 'RESULTADO
        strConsulta = strConsulta & "1.03.02' and id_usuario=" & usuario
        blVisualizar = TienePermiso(20)

    Case 17
        strConsulta = strConsulta & "1.04.1' and id_usuario=" & usuario
        blVisualizar = TienePermiso(20)
        
    Case 18 ' ARANCELES
        strConsulta = strConsulta & "1.01.18' and id_usuario=" & usuario
        blVisualizar = TienePermiso(18)
    

End Select
RaiseEvent ABMtollbar(i)
If Not blVisualizar Then
    RaiseEvent BUSCAR(False)
Else
    RaiseEvent BUSCAR(True)
End If
Visualizar = blVisualizar
blBuscar = blVisualizar
End Function

Private Function TienePermiso(Indice As Integer) As Boolean
Dim STRAUX As String
'On Error GoTo errManejador

Set rsPermiso = Coneccion.Execute(strConsulta, , adCmdText)


If rsPermiso.BOF And rsPermiso.EOF Then
    TienePermiso = False
    i = 0
Else
    TienePermiso = True
    '001,010,011,100,101,110,111
    '1    2  , 3 ,4  ,5 , 6 , 7
    
    If CBool(rsPermiso.Fields("ALTA").Value) = True Then
        STRAUX = "1"
    End If
    
    If CBool(rsPermiso.Fields("BAJA").Value) = True Then
        STRAUX = STRAUX & "1"
    Else
        STRAUX = STRAUX & "0"
    End If
    
    If CBool(rsPermiso.Fields("MODIFICACION").Value) = True Then
        STRAUX = STRAUX & "1"
    Else
        STRAUX = STRAUX & "0"
    
    
    End If
    

    rsPermiso.Close
    i = CInt(STRAUX)

End If

    Set rsPermiso = Nothing
    Class_Initialize
Exit Function
errManejador:
TienePermiso = False
i = 0
Class_Initialize
End Function

Public Property Get PERMISOABM() As Integer
PERMISOABM = i
End Property


Private Sub Class_Initialize()
strConsulta = "select alta, baja, modificacion from habilitacion" & _
            " where codmodulo='"
End Sub

Public Property Get enable() As Boolean
enable = blBuscar
End Property


