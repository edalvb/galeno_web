VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmPermisos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PERMISOS DEL USUARIO"
   ClientHeight    =   5745
   ClientLeft      =   1545
   ClientTop       =   435
   ClientWidth     =   8100
   Icon            =   "USUARIOS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8100
   Begin VB.CommandButton cmdABM 
      Caption         =   "Permiso de ABM"
      Height          =   480
      Left            =   5340
      TabIndex        =   11
      Top             =   5190
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Grabar"
      Height          =   480
      Left            =   6510
      TabIndex        =   10
      Top             =   5175
      Width           =   975
   End
   Begin VB.ListBox lstPermisos 
      Height          =   3180
      Left            =   4260
      TabIndex        =   9
      Top             =   1725
      Width           =   3135
   End
   Begin VB.CommandButton cmdAddDelPermiso 
      Height          =   315
      Index           =   3
      Left            =   3705
      Picture         =   "USUARIOS.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4110
      Width           =   360
   End
   Begin VB.CommandButton cmdAddDelPermiso 
      Height          =   315
      Index           =   2
      Left            =   3690
      Picture         =   "USUARIOS.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3405
      Width           =   360
   End
   Begin VB.CommandButton cmdAddDelPermiso 
      Height          =   315
      Index           =   1
      Left            =   3690
      Picture         =   "USUARIOS.frx":0ADE
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2730
      Width           =   360
   End
   Begin MSDataListLib.DataCombo dbCBoxUser 
      DataField       =   "id_usuario"
      Height          =   315
      Left            =   2175
      TabIndex        =   1
      Top             =   435
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "nombre"
      BoundColumn     =   "id_usuario"
      Text            =   "Seleccione el usuario"
   End
   Begin VB.CommandButton cmdAddDelPermiso 
      Height          =   315
      Index           =   0
      Left            =   3690
      Picture         =   "USUARIOS.frx":0E50
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2100
      Width           =   360
   End
   Begin VB.ListBox lstModulos 
      Height          =   3180
      ItemData        =   "USUARIOS.frx":1132
      Left            =   390
      List            =   "USUARIOS.frx":1134
      TabIndex        =   6
      Top             =   1725
      Width           =   3135
   End
   Begin VB.Label lblCaption 
      Caption         =   "Permisos del Usuario :"
      Height          =   330
      Index           =   2
      Left            =   4965
      TabIndex        =   8
      Top             =   1305
      Width           =   1830
   End
   Begin VB.Label lblCaption 
      Caption         =   "Permisos Disponibles :"
      Height          =   330
      Index           =   1
      Left            =   1005
      TabIndex        =   7
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label lblCaption 
      Caption         =   "Usuario del Sistema:"
      Height          =   330
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   450
      Width           =   1695
   End
End
Attribute VB_Name = "frmPermisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsUSUARIOS As Recordset
Dim PermUser As clsPermiso
Dim blCArgarListas As Boolean

Private Sub cmdABM_Click()
If Not (lstModulos.ListCount = 0 And lstPermisos.ListCount = 0) Then
    If dbCBoxUser.BoundText <> "" Then
        If blCArgarListas Then
            frmPermisoABM.Consultar dbCBoxUser.BoundText, Cnn
            frmPermisoABM.Show vbModal
        End If
    End If
End If
End Sub

Private Sub cmdAdd_Click()
'On Error Resume Next
PermUser.Grabar Cnn, dbCBoxUser.BoundText, lstModulos, lstPermisos
'ACTIVAR_ACCESOS
End Sub

Private Sub cmdAddDelPermiso_Click(Index As Integer)
Dim i As Integer
    Select Case Index
        Case 0
            i = lstModulos.ListIndex
            PermUser.addUno lstModulos, lstPermisos
            If lstModulos.ListCount = 0 Then
                setButton False
            Else
                If i > (lstModulos.ListCount - 1) Then
                    lstModulos.ListIndex = lstModulos.ListCount - 1
                Else
                    lstModulos.ListIndex = i
                End If
                cmdAddDelPermiso(2).Enabled = True
                cmdAddDelPermiso(3).Enabled = True
                
            End If
        Case 1
            PermUser.addTodo lstModulos, lstPermisos
            
            setButton False
        Case 2
            i = lstPermisos.ListIndex
            PermUser.delUno lstModulos, lstPermisos
            If lstPermisos.ListCount = 0 Then
                setButton True
            Else
                If i > (lstPermisos.ListCount - 1) Then
                    lstPermisos.ListIndex = lstPermisos.ListCount - 1
                Else
                    lstPermisos.ListIndex = i
                End If
                cmdAddDelPermiso(0).Enabled = True
                cmdAddDelPermiso(1).Enabled = True
                
            End If
        Case 3
            PermUser.delTodo lstModulos, lstPermisos
            setButton True
    End Select

End Sub

Private Sub dbCBoxUser_Change()

If blCArgarListas Then
    PermUser.CargarPermisos lstModulos, lstPermisos, Cnn, dbCBoxUser.BoundText
        
    
    If lstModulos.ListCount = 0 Then
        setButton False
    ElseIf lstPermisos.ListCount = 0 Then
        setButton True
    Else
        cmdAddDelPermiso(0).Enabled = True
        cmdAddDelPermiso(1).Enabled = True
        cmdAddDelPermiso(2).Enabled = True
        cmdAddDelPermiso(3).Enabled = True
    End If
End If

End Sub

Private Sub dbCBoxUser_Click(Area As Integer)
Debug.Print Area
End Sub

Private Sub dbCBoxUser_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then KeyCode = 0
End Sub

Private Sub dbCBoxUser_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then KeyAscii = 0

End Sub

Private Sub Form_Load()
blCArgarListas = False
Set rsUSUARIOS = New Recordset

rsUSUARIOS.Open "SELECT id_usuario, " & _
                "nombre from usuarios;", Cnn, adOpenStatic, _
                adLockReadOnly, adCmdText
If rsUSUARIOS.State = adStateOpen Then
    blCArgarListas = False
    Set dbCBoxUser.DataSource = rsUSUARIOS
    Set dbCBoxUser.RowSource = rsUSUARIOS
    Set PermUser = New clsPermiso
    'PermUser.CargarPermisos lstModulos, lstPermisos, Cnn, dbCBoxUser.BoundText
    blCArgarListas = True
    dbCBoxUser_Change
End If
End Sub

Private Sub setButton(blval As Boolean)
cmdAddDelPermiso(0).Enabled = blval
cmdAddDelPermiso(1).Enabled = blval
cmdAddDelPermiso(2).Enabled = Not blval
cmdAddDelPermiso(3).Enabled = Not blval
End Sub


Private Sub Form_Unload(Cancel As Integer)
Set PermUser = Nothing
End Sub

