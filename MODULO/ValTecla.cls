VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ValTecla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event Error(strDescrip As String)

Public Sub ValidarAscii(KeyAscii As Integer, Optional blNumero As Boolean = True)
'valida  el teclado en el caso que el campo sea numerico
If blNumero Then

      Select Case KeyAscii
        Case 1 To 7, 10 To 12, 14 To 47, 58 To 256
                KeyAscii = 0
      End Select
           
Else 'fecha
          Select Case KeyAscii
              Case 1 To 7, 10 To 12, 14 To 46, 58 To 256
                      KeyAscii = 0
            End Select
          
End If

End Sub

Public Function esFecha(strFecha As String) As Boolean
    esFecha = IsDate(strFecha)
    If Not esFecha Then RaiseEvent Error("La Fecha que introdujo, no es Valida" & vbCrLf & _
                              "Por Favor, Ingrese una Fecha Valida")
    
End Function

Public Sub habilitarToolbar(toolbrSglac As Toolbar, Optional blvalor As Boolean = True)
 Dim I As Integer
 On Error GoTo errManejador
 For I = 1 To toolbrSglac.Buttons.Count
    toolbrSglac.Buttons(I).Enabled = blvalor
 Next
Exit Sub
errManejador:
    
End Sub

