VERSION 5.00
Begin VB.Form frmODBCLogon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inicio de sesión en MEGASYS"
   ClientHeight    =   2565
   ClientLeft      =   2850
   ClientTop       =   1755
   ClientWidth     =   4650
   Icon            =   "frmODBCLogon.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboDSNList 
      Height          =   315
      Left            =   1905
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "Combo1"
      Top             =   1920
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   450
      Left            =   2820
      TabIndex        =   5
      Top             =   1845
      Width           =   1440
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   450
      Left            =   405
      TabIndex        =   4
      Top             =   1845
      Width           =   1440
   End
   Begin VB.Frame fraStep3 
      Caption         =   "Valores de conexión"
      ClipControls    =   0   'False
      Height          =   1170
      Index           =   0
      Left            =   60
      TabIndex        =   8
      Top             =   375
      Width           =   4455
      Begin VB.TextBox txtUID 
         Height          =   300
         Left            =   1560
         TabIndex        =   0
         Top             =   285
         Width           =   2715
      End
      Begin VB.TextBox txtPWD 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   630
         Width           =   2715
      End
      Begin VB.TextBox txtDatabase 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   780
         Visible         =   0   'False
         Width           =   2715
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "&Nombre del usuario:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   330
         Width           =   1410
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "&Contraseña:"
         Height          =   195
         Index           =   3
         Left            =   675
         TabIndex        =   6
         Top             =   660
         Width           =   990
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "&Empresa:"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   7
         Top             =   780
         Visible         =   0   'False
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmODBCLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public OK As Boolean
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function SQLCreateDataSource Lib "odbccp32.dll" (ByVal hwnd As Long, ByVal LPSTR As String) As Long
Private Declare Function SQLDataSources Lib "ODBC32.DLL" (ByVal henv&, ByVal fDirection%, ByVal szDSN$, ByVal cbDSNMax%, pcbDSN%, ByVal szDescription$, ByVal cbDescriptionMax%, pcbDescription%) As Integer
Private Declare Function SQLAllocEnv% Lib "ODBC32.DLL" (env&)
Const SQL_SUCCESS As Long = 0
Const SQL_FETCH_NEXT As Long = 1
Dim conexion As ADODB.Connection
Dim blExisteEcomipa As Boolean
Dim strCcon As String
Dim strUserName As String
Dim strPassword As String
Dim strUserId As String

Private Sub cmdCancel_Click()

OK = False
Me.Hide
On Error Resume Next
If conexion.State = adStateOpen Then
    conexion.Close
End If
Set conexion = Nothing
End Sub

Private Sub cmdOK_Click()
Dim i As Integer
Dim strError As String
Dim sConnect As String
Dim sADOConnect As String
Dim sDSN As String
sDSN = "DSN=GALENO" & ";"
sConnect = sConnect & "UID=" & txtUID.Text & ";"
sConnect = sConnect & "PWD=" & txtPWD.Text & ";"
sConnect = sConnect & "Database="
If Len(txtDatabase.Text) > 0 Then
    sConnect = sConnect & txtDatabase.Text & ";"
End If
sADOConnect = "PROVIDER=MSDataShape;Data PROVIDER=MSDASQL;" & sDSN & sConnect
strCcon = sADOConnect
strUserName = txtUID
strPassword = txtPWD
Set conexion = New ADODB.Connection
conexion.CursorLocation = adUseClient

'On Error GoTo errHandler
Call TestConexion
If OK Then
    If Not (Err.Number > 0) Then
        OK = True
        Me.Hide
    End If
Else
    MsgBox "Contraseña no valida", vbInformation
End If

Exit Sub
errHandler:
    MsgBox Err.Description, vbInformation
End Sub

Sub GetDSNsAndDrivers()
Dim i As Integer
Dim sDSNItem As String * 1024
Dim sDRVItem As String * 1024
Dim sDSN As String 'Nombre del DSN
Dim sDRV As String 'Controlador(driver) del DSN
Dim iDSNLen As Integer 'Longitud del DSN
Dim iDRVLen As Integer 'Longitud del Controlador(driver)
Dim lHenv As Long 'controlador de entorno
On Error Resume Next
cboDSNList.Clear
cboDSNList.AddItem "(Ninguno)"
'obtener los DSN
If SQLAllocEnv(lHenv) <> -1 Then
    Do Until i <> SQL_SUCCESS
        sDSNItem = Space$(1024)
        sDRVItem = Space$(1024)
        i = SQLDataSources(lHenv, SQL_FETCH_NEXT, sDSNItem, 1024, iDSNLen, sDRVItem, 1024, iDRVLen)
        sDSN = Left$(sDSNItem, iDSNLen)
        sDRV = Left$(sDRVItem, iDRVLen)
        If sDSN <> Space(iDSNLen) Then
            cboDSNList.AddItem sDSN
            'cboDrivers.AddItem sDRV
        End If
    Loop
End If
End Sub
Private Function ConfigDns() As Boolean
Dim lgResul As Long
MsgBox "Seleccione System Data Source y" & vbCrLf _
        & "No cambie el nombre del Data Source Name" & vbCrLf _
        & "que se encuentra por defecto", vbInformation
lgResul = SQLCreateDataSource(Me.hwnd, "General")
ConfigDns = CBool(lgResul)
End Function

Private Sub GetNomUser()
Dim lgResult As Long
Dim strName As String
Dim sizeBufer As Long
strName = Space$(255)
sizeBufer = Len(strName)
lgResult = GetUserName(strName, sizeBufer)
End Sub

Private Function ComprobarDSN() As Boolean
Dim strDsnName As String
Dim i As Integer
Dim blExiste As Boolean
GetNomUser
GetDSNsAndDrivers
For i = 0 To cboDSNList.ListCount - 1
    strDsnName = cboDSNList.List(i)
    If Trim$(strDsnName) = "GALENO" Then
        Exit Function
    End If
Next
i = MsgBox("No se encontró el DSN (Data Source Name) !" & vbCrLf _
& "Desea crear uno ahora", vbYesNo + vbInformation)
If i = vbYes Then
    If Not ConfigDns Then
        MsgBox "La configuracion del DSN " & vbCrLf _
        & "no culmino con exito; " & vbCrLf _
        & "Consulte con el Administrador del Sistema", vbInformation
        End
    Else
        ComprobarDSN
    End If
Else
    MsgBox "Consulte con su Administrador" & vbCrLf _
    & "o proveedor de software para configurar el DNS", vbInformation
    End
End If
End Function

Private Sub Form_Load()
ComprobarDSN
End Sub

Private Sub TestConexion()
Dim strConsulta As String
Dim rsAux As ADODB.Recordset
strConsulta = "select id_usuario from usuarios where nombre='" & strUserName & "' " & _
              "and clave='" & strPassword & "'"

conexion.Open strCcon + "G@LENO"
Set rsAux = conexion.Execute(strConsulta, , adCmdText)
If rsAux.EOF And rsAux.BOF Then
 OK = False
Else
 strUserId = rsAux.Fields(0).Value
 OK = True
End If

End Sub

Public Property Get usuario() As String
usuario = strUserName
End Property

Public Property Get IdUsuario() As String
IdUsuario = strUserId
End Property

Public Property Get Conect() As ADODB.Connection
Set Conect = conexion
End Property

