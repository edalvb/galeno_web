VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmPermisoABM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PERMISOS DE ABM"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
   Icon            =   "frmPermisoABM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   4815
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   405
      Left            =   3765
      TabIndex        =   5
      Top             =   3690
      Width           =   960
   End
   Begin VB.CheckBox chkABM 
      Caption         =   "Modificar"
      DataField       =   "MODIFICACION"
      Height          =   375
      Index           =   2
      Left            =   3255
      TabIndex        =   3
      Top             =   3060
      Width           =   960
   End
   Begin VB.CheckBox chkABM 
      Caption         =   "Borrar"
      DataField       =   "BAJA"
      Height          =   375
      Index           =   1
      Left            =   1815
      TabIndex        =   2
      Top             =   3060
      Width           =   1215
   End
   Begin VB.CheckBox chkABM 
      Caption         =   "Agregar"
      DataField       =   "ALTA"
      Height          =   330
      Index           =   0
      Left            =   495
      TabIndex        =   1
      Top             =   3060
      Width           =   1215
   End
   Begin MSDataListLib.DataList dblstPermABM 
      DataField       =   "CODMODULO"
      Height          =   1815
      Left            =   195
      TabIndex        =   0
      Top             =   720
      Width           =   4350
      _ExtentX        =   7673
      _ExtentY        =   3201
      _Version        =   393216
      ListField       =   "DESCRIPCION"
      BoundColumn     =   "CODMODULO"
   End
   Begin VB.Label lblPermiso 
      Caption         =   "Permisos Disponibles"
      Height          =   300
      Left            =   1305
      TabIndex        =   4
      Top             =   330
      Width           =   1935
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   195
      X2              =   195
      Y1              =   2895
      Y2              =   3480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000004&
      Index           =   1
      X1              =   4575
      X2              =   4575
      Y1              =   2880
      Y2              =   3465
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000004&
      Index           =   2
      X1              =   210
      X2              =   4575
      Y1              =   3465
      Y2              =   3465
   End
   Begin VB.Line lnPermisos 
      X1              =   0
      X2              =   4365
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   195
      X2              =   4560
      Y1              =   2880
      Y2              =   2880
   End
End
Attribute VB_Name = "frmPermisoABM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPermisoLista As ADODB.Recordset
Dim rsPermisoCheck As ADODB.Recordset
Dim strConsultaLista As String
Dim strConsultaCheck As String
Dim blCambio As Boolean

Public Sub Consultar(UserName As String, ConSglac As ADODB.Connection)
'On Error GoTo errManejo

If ConSglac.State = adStateOpen Then
    Set rsPermisoLista = New ADODB.Recordset
    Set rsPermisoCheck = New ADODB.Recordset
    
    rsPermisoLista.CursorLocation = adUseClient
    rsPermisoCheck.CursorLocation = adUseClient
    strConsultaLista = "SELECT MODULOS.DESCRIPCION, HABILITACION.CODMODULO" & _
                    " FROM MODULOS INNER JOIN HABILITACION ON MODULOS.CODMODULO = HABILITACION.CODMODULO " & _
                    " WHERE (((HABILITACION.ID_USUARIO)=" & UserName & _
                    " ) AND ((MODULOS.TYPO)<>'MENU'));"
    
    strConsultaCheck = "SELECT HABILITACION.ID_USUARIO, HABILITACION.CODMODULO, " & _
                    " HABILITACION.ALTA, HABILITACION.BAJA, HABILITACION.MODIFICACION " & _
                    " From HABILITACION " & _
                    " WHERE (((HABILITACION.ID_USUARIO)=" & UserName & ")) AND EXISTS(SELECT MODULOS.CODMODULO " & _
                    " From MODULOS " & _
                    " WHERE MODULOS.CODMODULO=HABILITACION.CODMODULO AND MODULOS.TYPO<>'MENU')"
    rsPermisoLista.Open strConsultaLista, ConSglac, adOpenStatic, adLockOptimistic, adCmdText
    If Not (rsPermisoLista.EOF And rsPermisoLista.BOF) Then
        rsPermisoCheck.Open strConsultaCheck, ConSglac, adOpenStatic, adLockBatchOptimistic, adCmdText
        Set dblstPermABM.RowSource = rsPermisoLista
        'Set dblstPermABM.DataSource = rsPermisoCheck
        Set chkABM(0).DataSource = rsPermisoCheck
        Set chkABM(1).DataSource = rsPermisoCheck
        Set chkABM(2).DataSource = rsPermisoCheck
        
        dblstPermABM.BoundText = rsPermisoLista.Fields("CODMODULO").Value
        dblstPermABM_Click
        blCambio = False
    Else
        
        setButton False
    End If
Else
    MsgBox "La conección a la base de datos" & vbCrLf & _
            "se encuentra cerrada", vbCritical
    setButton False
End If
Exit Sub
errManejo:
MsgBox Err.Description
setButton False
End Sub

Private Sub chkABM_Click(Index As Integer)
blCambio = True
End Sub

Private Sub cmdGrabar_Click()
On Error GoTo errMan

rsPermisoCheck.UpdateBatch adAffectAllChapters
blCambio = False
Exit Sub
errMan:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical
blCambio = False
End Sub

Private Sub dblstPermABM_Click()
If dblstPermABM.SelectedItem > -1 Then
    rsPermisoCheck.Filter = "CODMODULO='" & dblstPermABM.BoundText & "'"
    If (rsPermisoCheck.EOF And rsPermisoCheck.BOF) Then
        chkABM(0).Enabled = False
        chkABM(1).Enabled = False
        chkABM(1).Enabled = False
    Else
        chkABM(0).Enabled = True
        chkABM(1).Enabled = True
        chkABM(1).Enabled = True
    End If
End If

End Sub

Private Sub dblstPermABM_KeyUp(KeyCode As Integer, Shift As Integer)
 
 Select Case KeyCode
    Case vbKeyUp, vbKeyPageUp
     
    Case vbKeyDown, vbKeyPageDown
     
  End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error Resume Next
Dim i As Integer
If blCambio Then
    i = MsgBox("Desea guardar los cambios", vbInformation + vbYesNoCancel)

    If i = vbYes Then
        rsPermisoCheck.Filter = ""
        rsPermisoCheck.MoveFirst
        rsPermisoCheck.UpdateBatch adAffectAllChapters
    
    ElseIf i = vbNo Then
        rsPermisoCheck.Filter = ""
        rsPermisoCheck.MoveFirst
        rsPermisoCheck.CancelBatch adAffectAllChapters
    Else
        Cancel = True
    End If
    
End If
If Not Cancel Then
On Error Resume Next
 rsPermisoCheck.Close
 rsPermisoLista.Close
 Set rsPermisoLista = Nothing
 Set rsPermisoCheck = Nothing
End If

End Sub

Private Sub setButton(blval)
Dim oCheck As CheckBox
For Each oCheck In chkABM
    oCheck.Enabled = blval
Next
cmdGrabar.Enabled = blval
End Sub
