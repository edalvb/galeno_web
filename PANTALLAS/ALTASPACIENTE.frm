VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form ALTASPACIENTE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ALTAS PACIENTE"
   ClientHeight    =   6150
   ClientLeft      =   5850
   ClientTop       =   5895
   ClientWidth     =   7815
   Icon            =   "ALTASPACIENTE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   7815
   Begin VB.TextBox txtFields 
      DataField       =   "OBSPACI"
      Height          =   285
      Index           =   4
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Frame Frame1FichaClinica 
      Height          =   4935
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   6735
      Begin VB.TextBox txtFields 
         DataField       =   "FE_CARGA"
         Height          =   285
         Index           =   0
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Caption         =   "Ficha Clinica"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   9
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "Número de Movimiento"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   8
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Ficha Clinica"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "Observación"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   4
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         Caption         =   "Fecha de Carga"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   3
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Caption         =   "Ficha Clinica"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSComctlLib.ImageList HistoriaClinicaIL 
      Left            =   6960
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ALTASPACIENTE.frx":030A
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ALTASPACIENTE.frx":041C
            Key             =   "Last"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ALTASPACIENTE.frx":052E
            Key             =   "First"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ALTASPACIENTE.frx":0640
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ALTASPACIENTE.frx":0752
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ALTASPACIENTE.frx":0C94
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ALTASPACIENTE.frx":11D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ALTASPACIENTE.frx":1718
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ALTASPACIENTE.frx":182A
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ALTASPACIENTE.frx":193C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ALTASPACIENTE.frx":1D8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ALTASPACIENTE.frx":21E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ALTASPACIENTE.frx":285A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ALTASPACIENTE.frx":2964
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ALTASPACIENTE.frx":2A76
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ALTASPACIENTE.frx":2D90
            Key             =   "Agregar"
            Object.Tag             =   "Agregar"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ALTASPACIENTE.frx":30AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ALTASPACIENTE.frx":35EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ALTASPACIENTE.frx":3B7A
            Key             =   "MODIFICAR"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ALTASPACIENTE.frx":42A6
            Key             =   "requery"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ALTASPACIENTE.frx":47FB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TlbConsulta 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   6
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1005
      ButtonWidth     =   1429
      ButtonHeight    =   953
      Style           =   1
      ImageList       =   "HistoriaClinicaIL"
      HotImageList    =   "HistoriaClinicaIL"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Agregar"
            Key             =   "Agregar"
            Object.ToolTipText     =   "Agregar"
            ImageIndex      =   16
            Object.Width           =   1e-4
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Modificar"
            Key             =   "Modificar"
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Actualizar"
            Key             =   "Actualizar"
            Object.ToolTipText     =   "Actualizar"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Grabar"
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar"
            Key             =   "Cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cerrar"
            Key             =   "Cerrar"
            Object.ToolTipText     =   "Cerrar"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "ALTASPACIENTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim WithEvents adoHISTOCLI As Recordset
Attribute adoHISTOCLI.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Dim tipdocu As String
Dim nrodocu As Long

Public Sub PASARIDENTIDAD(tipdoc As String, nrodoc As Long)
    tipdocu = tipdoc
    nrodocu = nrodoc
End Sub
Private Sub Form_Load()
  Set adoHISTOCLI = New Recordset
  adoHISTOCLI.Open "select FICHACLINICA,TIPDOC,NRODOC,FE_CARGA,OBSPACI,VIGENTE from FICHACLINICA Order by FICHACLINICA", Cnn, adOpenStatic, adLockOptimistic
  adoHISTOCLI.MoveFirst
  adoHISTOCLI.Find "tipdoc='" & tipdocu & "'"
  adoHISTOCLI.Find "nrodoc=" & nrodocu
  If Not adoHISTOCLI.EOF Then
    TlbConsulta.Buttons(1).Enabled = False
  Else
    TlbConsulta.Buttons(1).Enabled = True
    TlbConsulta.Buttons(3).Enabled = False
    TlbConsulta.Buttons(8).Enabled = False
    TlbConsulta.Buttons(6).Enabled = False
  End If
  
  
    Dim oText As TextBox
  'Enlaza los cuadros de texto con el proveedor de datos
  For Each oText In Me.TXTFIELDS
    Set oText.DataSource = adoHISTOCLI
  Next
  Set ChVigente.DataSource = adoHISTOCLI
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  lblStatus.Width = Me.Width - 1500
  cmdNext.Left = lblStatus.Width + 700
  cmdLast.Left = cmdNext.Left + 340
End Sub



Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub



Private Sub ADOHISTOCLI_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Aquí se coloca el código de validación
  'Se llama a este evento cuando ocurre la siguiente acción
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  With adoHISTOCLI
    .AddNew
    Me.TXTFIELDS(0) = OPTENER_CODIGO("SELECT FICHACLINICA FROM FICHACLINICA ORDER BY FICHACLINICA DESC")
    adoHISTOCLI.Fields("TIPDOC").Value = tipdocu
    adoHISTOCLI.Fields("NRODOC").Value = nrodocu
   
  End With
  TXTFIELDS(3).Text = Date + Time
  TXTFIELDS(4).Locked = False
  TlbConsulta.Buttons(1).Enabled = False
  TlbConsulta.Buttons(3).Enabled = True
  TlbConsulta.Buttons(8).Enabled = True
  TlbConsulta.Buttons(6).Enabled = True
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub


Private Sub cmdRefresh_Click()
  'Esto sólo es necesario en aplicaciones multiusuario
  On Error GoTo RefreshErr
  adoHISTOCLI.Requery
  adoHISTOCLI.MoveFirst
  adoHISTOCLI.Find "tipdoc='" & tipdocu & "'"
  adoHISTOCLI.Find "nrodoc=" & nrodocu
  If Not adoHISTOCLI.EOF Then
    TlbConsulta.Buttons(1).Enabled = False
  Else
    TlbConsulta.Buttons(1).Enabled = True
     TXTFIELDS(4).Locked = True
  End If
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
  On Error GoTo EditErr
  TXTFIELDS(4).Locked = False
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub
Private Sub cmdCancel_Click()
  On Error Resume Next

  adoHISTOCLI.CancelUpdate
  adoHISTOCLI.MoveFirst
  adoHISTOCLI.Find "tipdoc='" & tipdocu & "'"
  adoHISTOCLI.Find "nrodoc=" & nrodocu
  If Not adoHISTOCLI.EOF Then
    TlbConsulta.Buttons(1).Enabled = False
  Else
    TlbConsulta.Buttons(1).Enabled = True
    TlbConsulta.Buttons(3).Enabled = False
    TlbConsulta.Buttons(8).Enabled = False
    TlbConsulta.Buttons(6).Enabled = False
  End If
   TXTFIELDS(4).Locked = True
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  adoHISTOCLI.UpdateBatch adAffectAll
  TXTFIELDS(4).Locked = True
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub


Private Sub TlbConsulta_ButtonClick(ByVal Button As MSComctlLib.Button)
  'On Error Resume Next
    Select Case Button.Key
          
        Case "Agregar"
              Call cmdAdd_Click
        Case "Modificar"
             Call cmdEdit_Click
        Case "Actualizar"
            Call cmdRefresh_Click
        Case "Grabar"
            Call cmdUpdate_Click
        Case "Cancelar"
            Call cmdCancel_Click
        Case "Cerrar"
            Unload Me
    End Select
End Sub
