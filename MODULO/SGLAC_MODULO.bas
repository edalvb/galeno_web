Attribute VB_Name = "MODGALENO"

Public Declare Function AnimateWindow Lib "user32" (ByVal hwnd As Long, ByVal tiemp As Long, ByVal bandera As Long) As Long
Public Const MOUSEEVENTF_ABSOLUTE = &H8000 '  movimiento absoluto
Public Const MOUSEEVENTF_LEFTDOWN = &H2 '  botón izquierdo presionado
Public Const MOUSEEVENTF_LEFTUP = &H4 '  botón izquierdo soltado
Public Const MOUSEEVENTF_MOVE = &H1 '  movimiento del mouse
Public Const MOUSEEVENTF_RIGHTDOWN = &H8 '  botón derecho presionado
Public Const MOUSEEVENTF_RIGHTUP = &H10 '  botón derecho soltado
Public Const conHwndTopmost = -1
Public Const conHwndNoTopmost = -2
Public Const conSwpNoActivate = &H10
Public Const conSwpShowWindow = &H40
Public Const SWP_NOOWNERZORDER = &H200      '  No usar el orden Z del propietario
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
     Type OSVERSIONINFO
         dwOSVersionInfoSize As Long
         dwMajorVersion As Long
         dwMinorVersion As Long
         dwBuildNumber As Long
         dwPlatformId As Long
         szCSDVersion As String * 128   ' Maintenance string for PSS usage.
      End Type

Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2

Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
        (lpVersionInformation As OSVERSIONINFO) As Long

Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Cnn As ADODB.Connection
Public INTERCAMBIO As String
Public Campo As String
Public usuario As String
Public IdUsuario As String
Sub Main()
Dim frmLogo As frmODBCLogon
'Dim n As Long
'n = 1500
Set frmLogo = New frmODBCLogon
frmLogo.Show vbModal
If frmLogo.OK Then
  Set Cnn = frmLogo.Conect
  usuario = frmLogo.usuario
  IdUsuario = frmLogo.IdUsuario
  'AnimateWindow MDIGALENO.hwnd, n, &HA0000
  'MDIGALENO.Hide
  MDIGALENO.Show
Else
    End
End If

Unload frmLogo
ACTIVAR_ACCESOS
End Sub

Public Function OPTENER_CODIGO(Sql As String) As Integer
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    RS.Open Sql, Cnn
    If RS.BOF = False And RS.EOF = False Then
        OPTENER_CODIGO = RS(0).Value + 1
    Else
        OPTENER_CODIGO = 1
    End If
    If OPTENER_CODIGO = Empty Then
        OPTENER_CODIGO = 1
    End If
    RS.Close
End Function


Public Function FOCO(ARGUM As Long) As Long
Dim i As Long
i = SetFocus(ARGUM)
FOCO = i
End Function
Public Sub ACTIVAR_ACCESOS()
Dim RSACCESOS As ADODB.Recordset
Set RSACCESOS = New ADODB.Recordset
'On Error Resume Next
RSACCESOS.Open "SELECT SUBSTRING (CODMODULO, 1, 1) AS menu, SUBSTRING (CODMODULO, 3, 2) AS modulo, ID_USUARIO AS usuario From HABILITACION GROUP BY SUBSTRING (CODMODULO, 1, 1), SUBSTRING (CODMODULO, 3, 2), ID_USUARIO Having (ID_USUARIO = " & IdUsuario & " ) ORDER BY ID_USUARIO", Cnn
Do While Not RSACCESOS.EOF
    
   Select Case Val(RSACCESOS.Fields(0).Value)
   Case 1 'archivos
        Select Case Val(RSACCESOS.Fields(1).Value)
        Case 1 'movimientos
             MDIGALENO.mnuArchManten.Enabled = True
        Case 2 'atencion al paciente
             MDIGALENO.ordenser.Enabled = True
        Case 3
            MDIGALENO.serinter.Enabled = True
        End Select
   Case 2
        Select Case Val(RSACCESOS.Fields(1).Value)
            Case 1
                MDIGALENO.prnPacientes.Enabled = True
            Case 2
            Case 3
            Case 17
            Case 4
                MDIGALENO.prnLineas.Enabled = True
            Case 5
                MDIGALENO.prnInsumos.Enabled = True
            Case 6
                MDIGALENO.prnDoctor.Enabled = True
            Case 7
                MDIGALENO.prnArea.Enabled = True
            Case 8
                MDIGALENO.prnEspecialidad.Enabled = True
            Case 9
                MDIGALENO.prnInstituto.Enabled = True
            Case 10

            Case 11
                MDIGALENO.prnUsuarios.Enabled = True
            Case 12
                MDIGALENO.prnCiudad.Enabled = True
            Case 13
                MDIGALENO.prnUsuarios.Enabled = True
            Case 14
                MDIGALENO.prnValnorm.Enabled = True
            Case 15
                MDIGALENO.prnTabVal.Enabled = True
            Case 16
        End Select
   Case 3
        Select Case Val(RSACCESOS.Fields(1).Value)
            Case 1
                MDIGALENO.ConsultaPacientes.Enabled = True
            Case 2
            Case 3
                MDIGALENO.ConsultaLineas.Enabled = True
            Case 4
                MDIGALENO.ConsultaInsumos.Enabled = True
            Case 5
                MDIGALENO.ConsultaDoctor.Enabled = True
            Case 6
                MDIGALENO.ConsultaArea.Enabled = True
            Case 7
                MDIGALENO.ConsultaEspecialidad.Enabled = True
            Case 8
                MDIGALENO.ConsultaValNorm.Enabled = True
            Case 9
                MDIGALENO.ConsultaInstituto.Enabled = True
            Case 10
            Case 11
                MDIGALENO.ConsultaUsuarios.Enabled = True
            Case 12
                MDIGALENO.ConsultaProveedores.Enabled = True
            Case 13
                MDIGALENO.ConsultaCiudad.Enabled = True
            Case 14
                MDIGALENO.Consultatabval.Enabled = True
            Case 15
        End Select
   Case 4
        Select Case Val(RSACCESOS.Fields(1).Value)
            Case 1
            Case 2
        End Select
   Case 5
   End Select
   
    
    RSACCESOS.MoveNext
Loop
RSACCESOS.Close

Set RSACCESOS = Cnn.Execute("SELECT ADMINISTRADOR FROM USUARIOS WHERE ID_USUARIO=" & IdUsuario, , adCmdText)

If Not (RSACCESOS.EOF And RSACCESOS.BOF) Then
    
    MDIGALENO.mnuPermisos.Visible = CBool(RSACCESOS(0).Value)
End If

On Error Resume Next
RSACCESOS.Close
Set RSACCESOS = Nothing
End Sub
Public Function SysVersions32() As Long
 Dim v As OSVERSIONINFO, retval As Long
 Dim WindowsVersion As String, BuildVersion As String
 Dim PlatformName As String

    v.dwOSVersionInfoSize = Len(v)
    retval = GetVersionEx(v)


SysVersions32 = v.dwMajorVersion
End Function


