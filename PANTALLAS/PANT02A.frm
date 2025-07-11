VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm MDIGALENO 
   Appearance      =   0  'Flat
   BackColor       =   &H80000001&
   Caption         =   "MEG@SYS"
   ClientHeight    =   6285
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7635
   Icon            =   "PANT02A.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "PANT02A.frx":030A
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1440
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483647
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PANT02A.frx":34C6A
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PANT02A.frx":34D7C
            Key             =   "Last"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PANT02A.frx":34E8E
            Key             =   "First"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PANT02A.frx":34FA0
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PANT02A.frx":350B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PANT02A.frx":355F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PANT02A.frx":35B36
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PANT02A.frx":36078
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PANT02A.frx":3618A
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PANT02A.frx":3629C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PANT02A.frx":366EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PANT02A.frx":36B40
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PANT02A.frx":371BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PANT02A.frx":372C4
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PANT02A.frx":373D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PANT02A.frx":376F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PANT02A.frx":37A0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PANT02A.frx":37F4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PANT02A.frx":384DA
            Key             =   "MODIFICAR"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PANT02A.frx":38C06
            Key             =   "requery"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PANT02A.frx":3915B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Left            =   5940
      Top             =   1440
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   5865
      Visible         =   0   'False
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "PANT02A.frx":392F5
      OLEDropMode     =   1
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   27
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "First"
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   2
            Object.Width           =   589
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Previous"
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Next"
            Object.ToolTipText     =   "Adelante"
            ImageIndex      =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Last"
            Object.ToolTipText     =   "Ultimo"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cortar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copiar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Pegar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Search"
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "MODIFICAR"
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Actualizar"
            Object.ToolTipText     =   "Actualizar"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Add"
            Object.ToolTipText     =   "Agregar"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancel"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "Ayuda"
            ImageIndex      =   14
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Menu mnuArchivos 
      Caption         =   "&Archivos"
      WindowList      =   -1  'True
      Begin VB.Menu mnuArchManten 
         Caption         =   "&Mantenimiento"
      End
      Begin VB.Menu ordenser 
         Caption         =   "&Orden de Servicios"
      End
      Begin VB.Menu serinter 
         Caption         =   "&Servicios de Internacion"
      End
      Begin VB.Menu atenmed 
         Caption         =   "&Atencion Medica"
      End
      Begin VB.Menu liquidacion 
         Caption         =   "&Liquidaciones"
      End
      Begin VB.Menu mnuSepara 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuInformes 
      Caption         =   "&Informes"
      Begin VB.Menu prnPacientes 
         Caption         =   "Pacientes"
      End
      Begin VB.Menu ficha 
         Caption         =   "Historia Clinica"
      End
      Begin VB.Menu ordenesdeser 
         Caption         =   "Ordenes de Servicio"
      End
      Begin VB.Menu internacion 
         Caption         =   "Internaciones"
      End
      Begin VB.Menu prnLineas 
         Caption         =   "Líneas"
      End
      Begin VB.Menu prnInsumos 
         Caption         =   "Insumos"
      End
      Begin VB.Menu prnDoctor 
         Caption         =   "Doctor"
      End
      Begin VB.Menu prnArea 
         Caption         =   "Area"
      End
      Begin VB.Menu prnEspecialidad 
         Caption         =   "Especialidad"
      End
      Begin VB.Menu prnInstituto 
         Caption         =   "Institutos"
      End
      Begin VB.Menu prnCiudad 
         Caption         =   "Ciudades"
      End
      Begin VB.Menu prnUsuarios 
         Caption         =   "Usuarios"
      End
      Begin VB.Menu prnValnorm 
         Caption         =   "Valores Normales"
      End
      Begin VB.Menu prnTabVal 
         Caption         =   "Tabla de Valores"
      End
      Begin VB.Menu prnaranceles 
         Caption         =   "Aranceles"
      End
      Begin VB.Menu prnservidoct 
         Caption         =   "Servicios Prestados por Doctores"
      End
      Begin VB.Menu prnserviinst 
         Caption         =   "Servicios Prestados por Institutos"
      End
      Begin VB.Menu prnservipac 
         Caption         =   "Servicios Adquiridos por Pacientes"
      End
   End
   Begin VB.Menu mnuConsulta 
      Caption         =   "&Consulta"
      Begin VB.Menu ConsultaPacientes 
         Caption         =   "Pacientes"
      End
      Begin VB.Menu ConsultaLineas 
         Caption         =   "Lineas"
      End
      Begin VB.Menu ConsultaInsumos 
         Caption         =   "Insumos"
      End
      Begin VB.Menu ConsultaDoctor 
         Caption         =   "Doctor"
      End
      Begin VB.Menu ConsultaArea 
         Caption         =   "Area"
      End
      Begin VB.Menu ConsultaEspecialidad 
         Caption         =   "Especialidad"
      End
      Begin VB.Menu ConsultaValNorm 
         Caption         =   "Valores Normales"
      End
      Begin VB.Menu ConsultaInstituto 
         Caption         =   "Instituto"
      End
      Begin VB.Menu ConsultaUsuarios 
         Caption         =   "Usuarios"
      End
      Begin VB.Menu ConsultaProveedores 
         Caption         =   "Proveedores"
      End
      Begin VB.Menu ConsultaCiudad 
         Caption         =   "Ciudad"
      End
      Begin VB.Menu Consultatabval 
         Caption         =   "Tabla de Valores"
      End
      Begin VB.Menu Consultaarancel 
         Caption         =   "Aranceles"
      End
      Begin VB.Menu Consultaservidoctor 
         Caption         =   "Servicios Prestados por Doctores"
      End
      Begin VB.Menu Consultaserviins 
         Caption         =   "Servicios Prestados por Institutos"
      End
      Begin VB.Menu Consultaadquipacien 
         Caption         =   "Servicios Adquiridos por Pacientes"
      End
   End
   Begin VB.Menu mnuOlap 
      Caption         =   "Servicios &Olap"
      Begin VB.Menu olappacientes 
         Caption         =   "DW Pacientes"
      End
      Begin VB.Menu OlapInstitutos 
         Caption         =   "DW Institutos"
      End
      Begin VB.Menu Olapprocedimientos 
         Caption         =   "DW Analisis"
      End
   End
   Begin VB.Menu mnuPermisos 
      Caption         =   "&Permisos"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "MDIGALENO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub pacientes_Click()
MANTENIMIENTO.Show
End Sub

Private Sub Consultaadquipacien_Click()
Dim conTablas As frmConsult
Set conTablas = New frmConsult
Set conTablas.ActiveConection = Cnn
conTablas.Modal = True
conTablas.Caption = "Consulta de Servicios Adquiridos por Pacientes"
conTablas.Consulta = "SELECT ORDENSERVICIOCABECERA.FICHACLINICA, ORDENSERVICIOCABECERA.FEC_PED, ORDENSERVICIOCABECERA.NROMOV, ORDENSERVICIOCABECERA.TIPDOC, ORDENSERVICIOCABECERA.NRODOC, PERSONAS.NOMBRES, PERSONAS.APELLIDOS, PERSONAS.TITULO, PERSONAS.FECNAC, PERSONAS.GENERO, PERSONAS.TELEFONOS, ORDENSERVICIOCABECERA.CODDR, PERSONAS_2.NOMBRES AS NOMDOC, PERSONAS_2.APELLIDOS AS APEDOC, ORDENSERVICIOCABECERA.COD_ESPECIALIDAD, ESPECIALIDAD.NOMBRE, ORDENSERVICIOCABECERA.COD_EMPRESA, EMPRESA.NOMBRE AS NOMEMPR, ORDENSERVICIOCABECERA.NRO_SEGURO, ORDENSERVICIOCABECERA.IDE_ASEGURADO, ORDENSERVICIOCABECERA.TIPO_SEGURO, ORDENSERVICIOCABECERA.TIPOMOV, TIPOMOV.NOMTIPOMOV, ORDENSERVICIOCABECERA.RESPECOTIPDOC, ORDENSERVICIOCABECERA.RESPECONRODOC, PERSONAS_1.NOMBRES AS RESNOM, PERSONAS_1.APELLIDOS AS RESAPE, PERSONAS_1.APECASADA, PERSONAS_1.TITULO AS RESTITUL, PERSONAS_1.TELEFONOS AS RESTEL, ORDENSERVICIOCABECERA.CODINS, INSTITUTOS.INSTITUTO " _
                   & " FROM ORDENSERVICIOCABECERA INNER JOIN DOCTOR ON ORDENSERVICIOCABECERA.CODDR = DOCTOR.CODDR  INNER JOIN ESPECIALIDAD ON ORDENSERVICIOCABECERA.COD_ESPECIALIDAD = ESPECIALIDAD.COD_ESPECIALIDAD INNER JOIN EMPRESA ON ORDENSERVICIOCABECERA.COD_EMPRESA = EMPRESA.COD_EMPRESA INNER JOIN TIPOMOV ON ORDENSERVICIOCABECERA.TIPOMOV = TIPOMOV.TIPOMOV INNER JOIN PERSONAS ON ORDENSERVICIOCABECERA.TIPDOC = PERSONAS.TIPDOC AND ORDENSERVICIOCABECERA.NRODOC = PERSONAS.NRODOC INNER JOIN INSTITUTOS ON ORDENSERVICIOCABECERA.CODINS = INSTITUTOS.CODINS INNER JOIN PERSONAS PERSONAS_2 ON DOCTOR.TIPDOC = PERSONAS_2.TIPDOC AND DOCTOR.NRODOC = PERSONAS_2.NRODOC LEFT OUTER JOIN PERSONAS PERSONAS_1 ON ORDENSERVICIOCABECERA.RESPECOTIPDOC = PERSONAS_1.TIPDOC AND ORDENSERVICIOCABECERA.RESPECONRODOC = PERSONAS_1.NRODOC "
SetWindowPos conTablas.hwnd, conHwndTopmost, 120, 72, 500, 320, &H50
conTablas.Show
End Sub

Private Sub Consultaarancel_Click()
Dim conTablas As frmConsult
Set conTablas = New frmConsult
Set conTablas.ActiveConection = Cnn
conTablas.Modal = True
conTablas.Caption = "Consulta de Aranceles"
conTablas.Consulta = "SELECT A.CODSERVICIO,B.NOMSERVICIO, C.INSTITUTO,A.ARANCEL,A.COSTO FROM ARANCELES A , SERVICIO B, INSTITUTOS C Where A.CODSERVICIO = B.CODSERVICIO AND A.CODINS = C.CODINS"
SetWindowPos conTablas.hwnd, conHwndTopmost, 120, 72, 500, 320, &H50
conTablas.Show
End Sub

Private Sub ConsultaArea_Click()
Dim conTablas As frmConsult
Set conTablas = New frmConsult
Set conTablas.ActiveConection = Cnn
conTablas.Modal = True
conTablas.Caption = "Consulta de Areas"
conTablas.Consulta = "select * from area"
SetWindowPos conTablas.hwnd, conHwndTopmost, 120, 72, 500, 320, &H50
conTablas.Show
End Sub

Private Sub ConsultaAseguradora_Click()
Dim conTablas As frmConsult
Set conTablas = New frmConsult
Set conTablas.ActiveConection = Cnn
conTablas.Modal = True
conTablas.Caption = "Consulta de Aseguradoras"
conTablas.Consulta = "select * from aseguradora"
SetWindowPos conTablas.hwnd, conHwndTopmost, 120, 72, 500, 320, &H50
conTablas.Show
End Sub

Private Sub ConsultaCiudad_Click()
Dim conTablas As frmConsult
Set conTablas = New frmConsult
Set conTablas.ActiveConection = Cnn
conTablas.Modal = True
conTablas.Caption = "Consulta de Ciudades"
conTablas.Consulta = "SELECT * From dbo.Parametros WHERE (TIPO = 'ciudad')"
SetWindowPos conTablas.hwnd, conHwndTopmost, 120, 72, 500, 320, &H50
conTablas.Show
End Sub

Private Sub ConsultaDoctor_Click()
Dim conTablas As frmConsult
Set conTablas = New frmConsult
Set conTablas.ActiveConection = Cnn
conTablas.Modal = True
conTablas.Caption = "Consulta de Doctor"
conTablas.Consulta = "SELECT DOCTOR.CODDR, PERSONAS.TIPDOC, PERSONAS.NRODOC, PERSONAS.NOMBRES, PERSONAS.APELLIDOS, PERSONAS.TELEFONOS FROM PERSONAS INNER JOIN DOCTOR ON PERSONAS.TIPDOC = DOCTOR.TIPDOC AND PERSONAS.NRODOC = DOCTOR.NRODOC"
SetWindowPos conTablas.hwnd, conHwndTopmost, 120, 72, 500, 320, &H50
conTablas.Show
End Sub

Private Sub ConsultaEspecialidad_Click()
Dim conTablas As frmConsult
Const conSwpShowWindow = &H40
Set conTablas = New frmConsult
Set conTablas.ActiveConection = Cnn
conTablas.Modal = True
conTablas.Caption = "Consulta de Especialidades"
conTablas.Consulta = "select * from especialidad"
SetWindowPos conTablas.hwnd, conHwndTopmost, 120, 72, 500, 320, &H50
conTablas.Show
End Sub

Private Sub ConsultaInstituto_Click()
Dim conTablas As frmConsult
Set conTablas = New frmConsult
Set conTablas.ActiveConection = Cnn
conTablas.Modal = True
conTablas.Caption = "Consulta de Institutos"
conTablas.Consulta = "SELECT a.CODINS, a.INSTITUTO, a.DOMICILIO, b.DESCRIPCION, a.TELEF1, a.TELEF2 FROM dbo.INSTITUTOS a INNER JOIN dbo.PARAMETROS b ON a.CODCIU = b.CODIGO"
SetWindowPos conTablas.hwnd, conHwndTopmost, 120, 72, 500, 320, &H50
conTablas.Show
End Sub

Private Sub ConsultaInsumos_Click()
Dim conTablas As frmConsult
Set conTablas = New frmConsult
Set conTablas.ActiveConection = Cnn
conTablas.Modal = True
conTablas.Caption = "Consulta de Insumos"
conTablas.Consulta = "SELECT TOP 100 PERCENT a.COD_INSUMO, a.TIPO_INSUMO, a.NOM_INSUMO, b.NOM_PROVEEDOR, c.LIN_NOMBRE, a.COSTO, a.ARANCEL " _
                    & " FROM dbo.INSUMOS a INNER JOIN " _
                    & " dbo.PROVEEDORES b ON a.COD_PROVEEDOR = b.COD_PROVEEDOR INNER JOIN " _
                    & " dbo.LINEAS c ON a.COD_LINEA = c.COD_LINEA " _
                    & " ORDER BY a.COD_INSUMO"
SetWindowPos conTablas.hwnd, conHwndTopmost, 120, 72, 500, 320, &H50
conTablas.Show
End Sub

Private Sub ConsultaLineas_Click()
Dim conTablas As frmConsult
Set conTablas = New frmConsult
Set conTablas.ActiveConection = Cnn
conTablas.Modal = True
conTablas.Caption = "Consulta de Líneas"
conTablas.Consulta = "select * from lineas"
SetWindowPos conTablas.hwnd, conHwndTopmost, 120, 72, 500, 320, &H50
conTablas.Show
End Sub

Private Sub ConsultaMatUsados_Click()
Dim conTablas As frmConsult
Set conTablas = New frmConsult
Set conTablas.ActiveConection = Cnn
conTablas.Modal = True
conTablas.Caption = "Consulta de Insumos Usados"
conTablas.Consulta = "select * from usomaterial"
SetWindowPos conTablas.hwnd, conHwndTopmost, 120, 72, 500, 320, &H50
conTablas.Show

End Sub

Private Sub ConsultaPacientes_Click()
Dim conTablas As frmConsult
Set conTablas = New frmConsult
Set conTablas.ActiveConection = Cnn
conTablas.Modal = True
conTablas.Caption = "Consulta de Pacientes"
conTablas.Consulta = "SELECT     PERSONAS.TIPDOC, PERSONAS.NRODOC, PERSONAS.NOMBRES, PERSONAS.APELLIDOS, PERSONAS.FECNAC, PERSONAS.APECASADA, " _
                    & " PERSONAS.TELEFONOS , PERSONAS.DOM_PARTICULAR, PERSONAS.DOM_LABORAL, PERSONAS.OBSERVACION " _
                    & " FROM PERSONAS INNER JOIN FICHACLINICA ON PERSONAS.TIPDOC = FICHACLINICA.TIPDOC AND PERSONAS.NRODOC = FICHACLINICA.NRODOC "
SetWindowPos conTablas.hwnd, conHwndTopmost, 120, 72, 500, 320, &H50
conTablas.Show
End Sub

Private Sub ConsultaPersonal_Click()
Dim conTablas As frmConsult
Set conTablas = New frmConsult
Set conTablas.ActiveConection = Cnn
conTablas.Modal = True
conTablas.Caption = "Consulta de Personal"
conTablas.Consulta = "select * from personal"
SetWindowPos conTablas.hwnd, conHwndTopmost, 120, 72, 500, 320, &H50
conTablas.Show
End Sub

Private Sub ConsultaProveedores_Click()
Dim conTablas As frmConsult
Set conTablas = New frmConsult
Set conTablas.ActiveConection = Cnn
conTablas.Modal = True
conTablas.Caption = "Consulta de Proveedores"
conTablas.Consulta = "select * from proveedores"
SetWindowPos conTablas.hwnd, conHwndTopmost, 120, 72, 500, 320, &H50
conTablas.Show
End Sub

Private Sub Consultaservidoctor_Click()
Dim conTablas As frmConsult
Set conTablas = New frmConsult
Set conTablas.ActiveConection = Cnn
conTablas.Modal = True
conTablas.Caption = "Consulta de Servicios Prestados por Doctores"
conTablas.Consulta = "SELECT SERVICIOINTERNACIONDETALLE.FICHACLINICA, SERVICIOINTERNACIONDETALLE.FEC_PED, SERVICIOINTERNACIONDETALLE.NROMOV, SERVICIOINTERNACIONDETALLE.TIPDOC, SERVICIOINTERNACIONDETALLE.NRODOC, PERSONAS.NOMBRES, PERSONAS.APELLIDOS, PERSONAS.APECASADA, PERSONAS.TITULO, SERVICIOINTERNACIONDETALLE.CODSERVICIO, SERVICIOINTERNACIONDETALLE.TIPOSER, SERVICIOINTERNACIONDETALLE.CODSER, SERVICIOINTERNACIONDETALLE.NOMSERVICIO, SERVICIOINTERNACIONDETALLE.CANTIDAD, SERVICIOINTERNACIONDETALLE.ARANCEL, SERVICIOINTERNACIONDETALLE.COSTO, SERVICIOINTERNACIONDETALLE.CODDR, PERSONAS1.NOMBRES AS NOMDOC, PERSONAS1.APELLIDOS AS APEDOC, PERSONAS1.APECASADA AS CASDOC, PERSONAS1.TITULO AS TITDOC, DOCTOR.REGPRF, SERVICIOINTERNACIONDETALLE.HONORARIOS, SERVICIOINTERNACIONDETALLE.FECHA_MOVIMIENTO, SERVICIOINTERNACIONDETALLE.OBSERVACIONES " _
                    & " FROM PERSONAS PERSONAS1 INNER JOIN DOCTOR ON PERSONAS1.TIPDOC = DOCTOR.TIPDOC AND PERSONAS1.NRODOC = DOCTOR.NRODOC INNER JOIN SERVICIOINTERNACIONDETALLE INNER JOIN PERSONAS ON SERVICIOINTERNACIONDETALLE.TIPDOC = PERSONAS.TIPDOC AND SERVICIOINTERNACIONDETALLE.NRODOC = PERSONAS.NRODOC ON DOCTOR.CODDR = SERVICIOINTERNACIONDETALLE.CODDR "
SetWindowPos conTablas.hwnd, conHwndTopmost, 120, 72, 500, 320, &H50
conTablas.Show
End Sub

Private Sub Consultaserviins_Click()
Dim conTablas As frmConsult
Set conTablas = New frmConsult
Set conTablas.ActiveConection = Cnn
conTablas.Modal = True
conTablas.Caption = "Consulta de Servicios Prestados por Institutos"
conTablas.Consulta = " SELECT ORDENSERVICIOCABECERA.FICHACLINICA, ORDENSERVICIOCABECERA.FEC_PED, ORDENSERVICIOCABECERA.NROMOV, ORDENSERVICIOCABECERA.TIPDOC, ORDENSERVICIOCABECERA.NRODOC, PERSONAS.NOMBRES, PERSONAS.APELLIDOS, PERSONAS.TITULO, PERSONAS.FECNAC, PERSONAS.GENERO, PERSONAS.TELEFONOS, ORDENSERVICIOCABECERA.CODDR, PERSONAS_2.NOMBRES AS NOMDOC, PERSONAS_2.APELLIDOS AS APEDOC, ORDENSERVICIOCABECERA.COD_ESPECIALIDAD, ESPECIALIDAD.NOMBRE, ORDENSERVICIOCABECERA.COD_EMPRESA, EMPRESA.NOMBRE AS NOMEMPR, ORDENSERVICIOCABECERA.NRO_SEGURO, ORDENSERVICIOCABECERA.IDE_ASEGURADO, ORDENSERVICIOCABECERA.TIPO_SEGURO, ORDENSERVICIOCABECERA.TIPOMOV, TIPOMOV.NOMTIPOMOV, ORDENSERVICIOCABECERA.RESPECOTIPDOC, ORDENSERVICIOCABECERA.RESPECONRODOC, PERSONAS_1.NOMBRES AS RESNOM, PERSONAS_1.APELLIDOS AS RESAPE, PERSONAS_1.APECASADA, PERSONAS_1.TITULO AS RESTITUL, PERSONAS_1.TELEFONOS AS RESTEL, ORDENSERVICIOCABECERA.CODINS, INSTITUTOS.INSTITUTO  " _
                  & " FROM ORDENSERVICIOCABECERA INNER JOIN DOCTOR ON ORDENSERVICIOCABECERA.CODDR = DOCTOR.CODDR  INNER JOIN ESPECIALIDAD ON ORDENSERVICIOCABECERA.COD_ESPECIALIDAD = ESPECIALIDAD.COD_ESPECIALIDAD INNER JOIN EMPRESA ON ORDENSERVICIOCABECERA.COD_EMPRESA = EMPRESA.COD_EMPRESA INNER JOIN TIPOMOV ON ORDENSERVICIOCABECERA.TIPOMOV = TIPOMOV.TIPOMOV INNER JOIN PERSONAS ON ORDENSERVICIOCABECERA.TIPDOC = PERSONAS.TIPDOC AND ORDENSERVICIOCABECERA.NRODOC = PERSONAS.NRODOC INNER JOIN INSTITUTOS ON ORDENSERVICIOCABECERA.CODINS = INSTITUTOS.CODINS INNER JOIN PERSONAS PERSONAS_2 ON DOCTOR.TIPDOC = PERSONAS_2.TIPDOC AND DOCTOR.NRODOC = PERSONAS_2.NRODOC LEFT OUTER JOIN PERSONAS PERSONAS_1 ON ORDENSERVICIOCABECERA.RESPECOTIPDOC = PERSONAS_1.TIPDOC AND ORDENSERVICIOCABECERA.RESPECONRODOC = PERSONAS_1.NRODOC "
SetWindowPos conTablas.hwnd, conHwndTopmost, 120, 72, 500, 320, &H50
conTablas.Show
End Sub

Private Sub Consultatabval_Click()
Dim conTablas As frmConsult
Set conTablas = New frmConsult
Set conTablas.ActiveConection = Cnn
conTablas.Modal = True
conTablas.Caption = "Consulta de Tablas de Valores"
conTablas.Consulta = "SELECT TOP 100 PERCENT a.CODTAB, a.NOMBRE, a.CODITEM, b.ITEM_DESCRIPCION FROM dbo.TABVAL a INNER JOIN dbo.ITEM_HISTORIACLINICA_ESPECIALI b ON a.CODITEM = b.ITEM ORDER BY a.CODTAB"
SetWindowPos conTablas.hwnd, conHwndTopmost, 120, 72, 500, 320, &H50
conTablas.Show
End Sub

Private Sub ConsultaUsuarios_Click()
Dim conTablas As frmConsult
Set conTablas = New frmConsult
Set conTablas.ActiveConection = Cnn
conTablas.Modal = True
conTablas.Caption = "Consulta de Usuarios"
conTablas.Consulta = "SELECT ID_USUARIO, NOMBRE, CLAVE, ADMINISTRADOR FROM dbo.USUARIOS"
SetWindowPos conTablas.hwnd, conHwndTopmost, 120, 72, 500, 320, &H50
conTablas.Show
End Sub

Private Sub ConsultaValNorm_Click()
Dim conTablas As frmConsult
Set conTablas = New frmConsult
Set conTablas.ActiveConection = Cnn
conTablas.Modal = True
conTablas.Caption = "Consulta de Valores Normales"
conTablas.Consulta = "SELECT TOP 100 PERCENT a.CODSERVICIO, a.TIPOSER, b.NOMSERVICIO, a.SEXO, a.EDAD_INF,a.EDAD_SUP, a.LIM_INF, a.LIM_SUP, a.LIM_INFN, a.LIM_SUPN FROM dbo.VALNORM a INNER JOIN dbo.SERVICIO b ON a.CODSERVICIO = b.CODSERVICIO AND a.TIPOSER = b.TIPOSER ORDER BY a.CODSERVICIO"
SetWindowPos conTablas.hwnd, conHwndTopmost, 120, 72, 500, 320, &H50
conTablas.Show
End Sub
Private Sub ficha_Click()
Dim Consulta As frmConsult
Dim strSQL As String
Dim strOrden As String
Set Consulta = New frmConsult
Consulta.Modal = True
Consulta.Caption = "HISTORIA CLINICA"
Consulta.Top = 1050
Set Consulta.ActiveConection = Cnn
Consulta.Consulta = "SELECT     HISTORIACLINICACABECERA.FICHACLINICA, HISTORIACLINICACABECERA.FEC_PED, HISTORIACLINICACABECERA.NROMOV, " _
                   & " HISTORIACLINICACABECERA.TIPDOC, HISTORIACLINICACABECERA.NRODOC, PERSONAS.NOMBRES, PERSONAS.APELLIDOS, " _
                   & " PERSONAS.APECASADA, PERSONAS.TITULO, PERSONAS.FECNAC, PERSONAS.GENERO, HISTORIACLINICACABECERA.COD_DOCTOR, " _
                   & " PERSONAS_1.NOMBRES AS NOMDOC, PERSONAS_1.APELLIDOS AS APEDOC, PERSONAS_1.APECASADA AS CASDOC, " _
                   & " PERSONAS_1.TITULO AS TITULDOC, HISTORIACLINICACABECERA.MOTIVOINGRESO, HISTORIACLINICACABECERA.DATOS_PROPORCIONADOS_POR, " _
                   & " HISTORIACLINICACABECERA.INFTECCIR, HISTORIACLINICACABECERA.MEDICAMENTOS, HISTORIACLINICACABECERA.EVOLUCION, " _
                   & " HISTORIACLINICACABECERA.PROCEDIMIENTOS , HISTORIACLINICACABECERA.ANOTACION " _
                   & " FROM HISTORIACLINICACABECERA INNER JOIN " _
                   & " PERSONAS ON HISTORIACLINICACABECERA.TIPDOC = PERSONAS.TIPDOC AND " _
                   & " HISTORIACLINICACABECERA.NRODOC = PERSONAS.NRODOC LEFT OUTER JOIN " _
                   & " DOCTOR ON HISTORIACLINICACABECERA.TIPDOC = DOCTOR.TIPDOC AND " _
                   & " HISTORIACLINICACABECERA.NRODOC = DOCTOR.NRODOC LEFT OUTER JOIN " _
                   & " PERSONAS PERSONAS_1 ON DOCTOR.TIPDOC = PERSONAS_1.TIPDOC AND DOCTOR.NRODOC = PERSONAS_1.TIPDOC"
Consulta.Show vbModal
If Not Consulta.Descargado Then
        With DataGaleno.rsHISTORIACLINICA_Grouping
           If .State = adStateClosed Then
                    .Open 'colocar este codigo a todos los informes
                    .Filter = ""
                    .Requery
                    .MoveFirst
                    .Filter = Consulta.msGrid.Columns(0).Caption & " = " & (Consulta.msGrid.Columns(0).Value)
           
                    'strSQL = Consulta.GetFilterString
           End If
           'If strSQL <> "" Then
           '         .Filter = strSQL
           'Else
           '         .Filter = ""
           'End If
           strOrden = Consulta.GetOrder
           If strOrden <> "" Then
                    .Sort = strOrden
           End If
           rpt_historiaclinica.Refresh
           rpt_historiaclinica.Show
           .Filter = ""
           If .State = adStateOpen Then 'colocar este codigo a todos los informes
              .Close
              strSQL = ""
           End If
        End With
    Consulta.Cerrar
End If


End Sub

Private Sub hconsult_Click()
Dim Consulta As frmConsult
Dim strSQL As String
Dim strOrden As String
Set Consulta = New frmConsult
Consulta.Modal = True
Consulta.Caption = "Consulta de Informe de Indicaciones"
Consulta.Top = 1050
Set Consulta.ActiveConection = Cnn
Consulta.Consulta = "select * from paciente"
Consulta.Show vbModal
If Not Consulta.Descargado Then
        With DataEnvironment1.rshistconsulta_Grouping
           If .State = adStateClosed Then
                    .Open 'colocar este codigo a todos los informes
                    .Filter = ""
                    .Requery
                    .MoveFirst
                    .Filter = Consulta.msGrid.Columns(5).Caption & " = " & (Consulta.msGrid.Columns(5).Value)
           
                    'strSQL = Consulta.GetFilterString
           End If
           'If strSQL <> "" Then
           '         .Filter = strSQL
           'Else
           '         .Filter = ""
           'End If
           strOrden = Consulta.GetOrder
           If strOrden <> "" Then
                    .Sort = strOrden
           End If
                    DataReport19.Refresh
                    DataReport19.Show
                    .Filter = ""
                    
                    'DataEnvironment1.rspaciente.Close
            If .State = adStateOpen Then 'colocar este codigo a todos los informes
                    .Close
                    
                    strSQL = ""
            End If
                
        End With
   
    Consulta.Cerrar
End If
End Sub

Private Sub atenmed_Click()
fncBotones True, "ATENCIONMEDICA"
ATENCIONMEDICA.Show
ATENCIONMEDICA.buscarReg
End Sub

Private Sub internacion_Click()
Dim Consulta As frmConsult
Dim strSQL As String
Dim strOrden As String
Set Consulta = New frmConsult
Consulta.Modal = True
Consulta.Caption = "Servicios de Internacion"
Consulta.Top = 1050
Set Consulta.ActiveConection = Cnn
Consulta.Consulta = "SELECT SERVICIOINTERNACIONCABECERA.FICHACLINICA, SERVICIOINTERNACIONCABECERA.FEC_PED, SERVICIOINTERNACIONCABECERA.NROMOV, " _
                    & " SERVICIOINTERNACIONCABECERA.TIPDOC, SERVICIOINTERNACIONCABECERA.NRODOC, PERSONAS.NOMBRES, PERSONAS.APELLIDOS, " _
                    & " PERSONAS.APECASADA, PERSONAS.TITULO, PERSONAS.FECNAC, PERSONAS.GENERO, PERSONAS.TELEFONOS, " _
                    & " SERVICIOINTERNACIONCABECERA.FEC_INTERNACION, SERVICIOINTERNACIONCABECERA.COD_HABITACION, " _
                    & " SERVICIOINTERNACIONCABECERA.COD_CAMA, SERVICIOINTERNACIONCABECERA.TIPOMOV, TIPOMOV.NOMTIPOMOV, " _
                    & " SERVICIOINTERNACIONCABECERA.COD_TIPO " _
                    & " FROM         SERVICIOINTERNACIONCABECERA INNER JOIN " _
                    & " PERSONAS ON SERVICIOINTERNACIONCABECERA.TIPDOC = PERSONAS.TIPDOC AND " _
                    & " SERVICIOINTERNACIONCABECERA.NRODOC = PERSONAS.NRODOC INNER JOIN " _
                    & " TIPOMOV ON SERVICIOINTERNACIONCABECERA.TIPOMOV = TIPOMOV.TIPOMOV"
Consulta.Show vbModal
If Not Consulta.Descargado Then
        With DataGaleno.rsSERVICIOINTERNACION
           If .State = adStateClosed Then
                    .Open 'colocar este codigo a todos los informes
                    .Filter = ""
                    .Requery
                    .MoveFirst
                    '.Filter = Consulta.msGrid.Columns(0).Caption & " = " & (Consulta.msGrid.Columns(0).Value)
                    strSQL = Consulta.GetFilterString
           End If
           If strSQL <> "" Then
                    .Filter = strSQL
           Else
                    .Filter = ""
           End If
           strOrden = Consulta.GetOrder
           If strOrden <> "" Then
                    .Sort = strOrden
           End If
           rpt_serviciointernacion.Refresh
           rpt_serviciointernacion.Show
           .Filter = ""
           If .State = adStateOpen Then 'colocar este codigo a todos los informes
              .Close
              strSQL = ""
           End If
        End With
    Consulta.Cerrar
End If


End Sub

Private Sub liquidacion_Click()
LIQUIDACIONES.Show
End Sub

Private Sub MDIForm_Load()
fncBotones True, ""
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
On Error Resume Next
Cnn.Close
End Sub

Private Sub mnuArchManten_Click()
MANTENIMIENTO.Show
End Sub

Private Sub mnuindicaciones_Click()
Dim Consulta As frmConsult
Dim strSQL As String
Dim strOrden As String
Set Consulta = New frmConsult
Consulta.Modal = True
Consulta.Caption = "Consulta de Informe de Indicaciones"
Consulta.Top = 1050
Set Consulta.ActiveConection = Cnn
Consulta.Consulta = "select FECHA AS FEC_PED, NROMOV,CEDULA,RECNOMBRE,RECETARIO,INDICACION from recetario"
Consulta.Show vbModal
If Not Consulta.Descargado Then
    With DataEnvironment1.rsrecetario_Grouping
           If .State = adStateClosed Then
                    .Open 'colocar este codigo a todos los informes
                    '.Filter = ""
                    .Requery
                    .MoveFirst
                    strSQL = Consulta.msGrid.Columns(0).Caption & " =#" & (Consulta.msGrid.Columns(0).Text) & "# and " & Consulta.msGrid.Columns(1).Caption & " = " & (Consulta.msGrid.Columns(1).Text)
                    .Filter = strSQL
                    'strSQL = Consulta.GetFilterString
           End If
           If strSQL <> "" Then
                    .Filter = strSQL
           Else
                    .Filter = ""
           End If
           'strOrden = Consulta.GetOrder
           'If strOrden <> "" Then
           '         .Sort = strOrden
           'End If
                    DataReport17.Show
                    DataReport17.Refresh
                    'DataEnvironment1.rspaciente.Close
            If .State = adStateOpen Then 'colocar este codigo a todos los informes
                    .Close
                    strSQL = ""
            End If
    End With
Consulta.Cerrar
End If

End Sub

Private Sub mnuPermisos_Click()
frmPermisos.Show
End Sub

Private Sub mnuRemision_Click()
frmPedFact.Show
End Sub

Private Sub mnuSalir_Click()
On Error Resume Next
Cnn.Close
End
End Sub

Private Sub mnuUtilitario_Click()

End Sub
Private Sub movSistemaPersonal_Click()
frmMoviPerson.Show
End Sub

Private Sub OlapAnalisis_Click()
dw_analisis.Show
End Sub

Private Sub OlapInstitutos_Click()
dw_institutos.Show
End Sub

Private Sub pedidos_Click()
frmPEDIDOS.Show
End Sub

Private Sub prnInformes_Click()
Dim Consulta As frmConsult
Dim strSQL As String
Dim strOrden As String
Set Consulta = New frmConsult
Consulta.Modal = True
Consulta.Caption = "Consulta de Informe de Resultado"
Consulta.Top = 1050
Set Consulta.ActiveConection = Cnn
Consulta.Consulta = "select * from movana1"
Consulta.Show vbModal
If Not Consulta.Descargado Then
    With DataEnvironment1.rsRESULTADOS_Grouping
           .Open
           .Requery
           .MoveFirst
           '.Filter = Consulta.msGrid.Columns(0).Caption & " = " & (Consulta.msGrid.Columns(0).Value)
           '.Filter = Consulta.msGrid.Columns(1).Caption & " = " & (Consulta.msGrid.Columns(1).Value)
           strSQL = Consulta.GetFilterString
           If strSQL <> "" Then
            .Filter = strSQL
           End If
           strOrden = Consulta.GetOrder
           If strOrden <> "" Then
            .Sort = strOrden
           End If
            DataReport10.Show
            DataReport10.Refresh
            'DataEnvironment1.rsRESULTADOS_Grouping.Close
            If .State = adStateOpen Then 'colocar este codigo a todos los informes
                .Close
            End If
    End With
Consulta.Cerrar
End If
End Sub

Private Sub olappacientes_Click()
dw_pacientes.Show
End Sub

Private Sub Olapprocedimientos_Click()
dw_prodimientos.Show
End Sub

Private Sub ordenes_Click()

End Sub

Private Sub ordenesdeser_Click()
Dim Consulta As frmConsult
Dim strSQL As String
Dim strOrden As String
Set Consulta = New frmConsult
Consulta.Modal = True
Consulta.Caption = "ORDENES DE SERVICIOS"
Consulta.Top = 1050
Set Consulta.ActiveConection = Cnn
Consulta.Consulta = "SELECT ORDENSERVICIOCABECERA.FICHACLINICA, ORDENSERVICIOCABECERA.FEC_PED, ORDENSERVICIOCABECERA.NROMOV, ORDENSERVICIOCABECERA.TIPDOC, ORDENSERVICIOCABECERA.NRODOC, PERSONAS.NOMBRES, PERSONAS.APELLIDOS, PERSONAS.TITULO, PERSONAS.FECNAC, PERSONAS.GENERO, PERSONAS.TELEFONOS, ORDENSERVICIOCABECERA.CODDR, PERSONAS_2.NOMBRES AS NOMDOC, PERSONAS_2.APELLIDOS AS APEDOC, ORDENSERVICIOCABECERA.COD_ESPECIALIDAD, ESPECIALIDAD.NOMBRE," _
                    & " ORDENSERVICIOCABECERA.COD_EMPRESA, EMPRESA.NOMBRE AS NOMEMPR, ORDENSERVICIOCABECERA.NRO_SEGURO, ORDENSERVICIOCABECERA.IDE_ASEGURADO, ORDENSERVICIOCABECERA.TIPO_SEGURO, ORDENSERVICIOCABECERA.TIPOMOV, TIPOMOV.NOMTIPOMOV, ORDENSERVICIOCABECERA.RESPECOTIPDOC, " _
                    & " ORDENSERVICIOCABECERA.RESPECONRODOC, PERSONAS_1.NOMBRES AS RESNOM, PERSONAS_1.APELLIDOS AS RESAPE, PERSONAS_1.APECASADA, PERSONAS_1.TITULO AS RESTITUL, PERSONAS_1.TELEFONOS AS RESTEL, ORDENSERVICIOCABECERA.CODINS, INSTITUTOS.INSTITUTO " _
                    & " FROM ORDENSERVICIOCABECERA INNER JOIN DOCTOR ON ORDENSERVICIOCABECERA.CODDR = DOCTOR.CODDR INNER JOIN " _
                    & " ESPECIALIDAD ON ORDENSERVICIOCABECERA.COD_ESPECIALIDAD = ESPECIALIDAD.COD_ESPECIALIDAD INNER JOIN EMPRESA ON ORDENSERVICIOCABECERA.COD_EMPRESA = EMPRESA.COD_EMPRESA INNER JOIN TIPOMOV ON ORDENSERVICIOCABECERA.TIPOMOV = TIPOMOV.TIPOMOV INNER JOIN PERSONAS ON ORDENSERVICIOCABECERA.TIPDOC = PERSONAS.TIPDOC AND ORDENSERVICIOCABECERA.NRODOC = PERSONAS.NRODOC INNER JOIN INSTITUTOS ON ORDENSERVICIOCABECERA.CODINS = INSTITUTOS.CODINS " _
                    & " INNER JOIN PERSONAS PERSONAS_2 ON DOCTOR.TIPDOC = PERSONAS_2.TIPDOC AND DOCTOR.NRODOC = PERSONAS_2.NRODOC LEFT OUTER JOIN PERSONAS PERSONAS_1 ON ORDENSERVICIOCABECERA.RESPECOTIPDOC = PERSONAS_1.TIPDOC AND ORDENSERVICIOCABECERA.RESPECONRODOC = PERSONAS_1.NRODOC "
Consulta.Show vbModal
If Not Consulta.Descargado Then
        With DataGaleno.rsORDERNESSERVICIOS
           If .State = adStateClosed Then
                    .Open 'colocar este codigo a todos los informes
                    .Filter = ""
                    .Requery
                    .MoveFirst
                    '.Filter = Consulta.msGrid.Columns(0).Caption & " = " & (Consulta.msGrid.Columns(0).Value)
                    strSQL = Consulta.GetFilterString
           End If
           If strSQL <> "" Then
                    .Filter = strSQL
           Else
                    .Filter = ""
           End If
           strOrden = Consulta.GetOrder
           If strOrden <> "" Then
                    .Sort = strOrden
           End If
           rpt_ordenes.Refresh
           rpt_ordenes.Show
           .Filter = ""
           If .State = adStateOpen Then 'colocar este codigo a todos los informes
              .Close
              strSQL = ""
           End If
        End With
    Consulta.Cerrar
End If
End Sub

Private Sub ordenser_Click()
ORDENES.Show
End Sub

Private Sub prnaranceles_Click()
Dim Consulta As frmConsult
Dim strSQL As String
Dim strOrden As String
Set Consulta = New frmConsult
Consulta.Modal = True
Consulta.Caption = "Consulta de Informe de Aranceles"
Consulta.Top = 1050
Set Consulta.ActiveConection = Cnn
Consulta.Consulta = "SELECT A.CODSERVICIO,B.NOMSERVICIO, C.INSTITUTO,A.ARANCEL,A.COSTO FROM ARANCELES A , SERVICIO B, INSTITUTOS C Where A.CODSERVICIO = B.CODSERVICIO AND A.CODINS = C.CODINS"
Consulta.Show vbModal
If Not Consulta.Descargado Then
    With DataGaleno.rsarancel
           If .State = adStateClosed Then 'colocar este codigo a todos los informes
           .Open
           .Requery
           .MoveFirst
           '.Filter = Consulta.msGrid.Columns(0).Caption & " = " & (Consulta.msGrid.Columns(0).Value)
           '.Filter = Consulta.msGrid.Columns(1).Caption & " = " & (Consulta.msGrid.Columns(1).Value)
           strSQL = Consulta.GetFilterString
           End If
           If strSQL <> "" Then
                    .Filter = strSQL
           Else
                    .Filter = ""
           End If
           strOrden = Consulta.GetOrder
           If strOrden <> "" Then
            .Sort = strOrden
           End If
            rpt_arancel.Show
            rpt_arancel.Refresh
            'DataEnvironment1.rspaciente.Close
            If .State = adStateOpen Then 'colocar este codigo a todos los informes
                .Close
            End If
    End With
Consulta.Cerrar
End If
End Sub

Private Sub prnArea_Click()
Dim Consulta As frmConsult
Dim strSQL As String
Dim strOrden As String
Set Consulta = New frmConsult
Consulta.Modal = True
Consulta.Caption = "Consulta de Informe de Areas"
Consulta.Top = 1050
Set Consulta.ActiveConection = Cnn
Consulta.Consulta = "select * from area"
Consulta.Show vbModal
If Not Consulta.Descargado Then
    With DataGaleno.rsarea
           If .State = adStateClosed Then 'colocar este codigo a todos los informes
           .Open
           .Requery
           .MoveFirst
           '.Filter = Consulta.msGrid.Columns(0).Caption & " = " & (Consulta.msGrid.Columns(0).Value)
           '.Filter = Consulta.msGrid.Columns(1).Caption & " = " & (Consulta.msGrid.Columns(1).Value)
           strSQL = Consulta.GetFilterString
           End If
           If strSQL <> "" Then
                    .Filter = strSQL
           Else
                    .Filter = ""
           End If
           strOrden = Consulta.GetOrder
           If strOrden <> "" Then
            .Sort = strOrden
           End If
            rpt_área.Show
            rpt_área.Refresh
            'DataEnvironment1.rspaciente.Close
            If .State = adStateOpen Then 'colocar este codigo a todos los informes
                .Close
            End If
    End With
Consulta.Cerrar
End If
End Sub

Private Sub prnAsegurador_Click()
Dim Consulta As frmConsult
Dim strSQL As String
Dim strOrden As String
Set Consulta = New frmConsult
Consulta.Modal = True
Consulta.Caption = "Consulta de Informe de Aseguradora"
Consulta.Top = 1050
Set Consulta.ActiveConection = Cnn
Consulta.Consulta = "select * from aseguradora"
Consulta.Show vbModal
If Not Consulta.Descargado Then
    With DataEnvironment1.rsASEGURADORA
           If .State = adStateClosed Then 'colocar este codigo a todos los informes
           .Open
           .Requery
           .MoveFirst
           '.Filter = Consulta.msGrid.Columns(0).Caption & " = " & (Consulta.msGrid.Columns(0).Value)
           '.Filter = Consulta.msGrid.Columns(1).Caption & " = " & (Consulta.msGrid.Columns(1).Value)
           strSQL = Consulta.GetFilterString
           End If
           If strSQL <> "" Then
                    .Filter = strSQL
           Else
                    .Filter = ""
           End If
           strOrden = Consulta.GetOrder
           If strOrden <> "" Then
            .Sort = strOrden
           End If
            DataReport9.Show
            DataReport9.Refresh
            'DataEnvironment1.rspaciente.Close
            If .State = adStateOpen Then 'colocar este codigo a todos los informes
                .Close
            End If
    End With
Consulta.Cerrar
End If
End Sub

Private Sub prnCiudad_Click()
Dim Consulta As frmConsult
Dim strSQL As String
Dim strOrden As String
Set Consulta = New frmConsult
Consulta.Modal = True
Consulta.Caption = "Consulta de Informe de Ciudad"
Consulta.Top = 1050
Set Consulta.ActiveConection = Cnn
Consulta.Consulta = "SELECT * From dbo.Parametros WHERE (TIPO = 'ciudad')"
Consulta.Show vbModal
If Not Consulta.Descargado Then
    With DataGaleno.rsciudad
           If .State = adStateClosed Then  'colocar este codigo a todos los informes
           .Open
           .Requery
           .MoveFirst
           '.Filter = Consulta.msGrid.Columns(0).Caption & " = " & (Consulta.msGrid.Columns(0).Value)
           '.Filter = Consulta.msGrid.Columns(1).Caption & " = " & (Consulta.msGrid.Columns(1).Value)
           strSQL = Consulta.GetFilterString
           End If
           If strSQL <> "" Then
                    .Filter = strSQL
           Else
                    .Filter = ""
           End If
           strOrden = Consulta.GetOrder
           If strOrden <> "" Then
            .Sort = strOrden
           End If
            rpt_ciudad.Show
            rpt_ciudad.Refresh
            If .State = adStateOpen Then
               .Close
            End If
    End With
Consulta.Cerrar
End If
End Sub

Private Sub prnDoctor_Click()
Dim Consulta As frmConsult
Dim strSQL As String
Dim strOrden As String
Set Consulta = New frmConsult
Consulta.Modal = True
Consulta.Caption = "Consulta de Informe de Doctores"
Set Consulta.ActiveConection = Cnn
Consulta.Consulta = "SELECT DOCTOR.CODDR, PERSONAS.TIPDOC, PERSONAS.NRODOC, PERSONAS.NOMBRES, PERSONAS.APELLIDOS, PERSONAS.TELEFONOS FROM PERSONAS INNER JOIN DOCTOR ON PERSONAS.TIPDOC = DOCTOR.TIPDOC AND PERSONAS.NRODOC = DOCTOR.NRODOC"
Consulta.Top = 1050
Consulta.Show vbModal
If Not Consulta.Descargado Then
    With DataGaleno.rsdoctor
           If .State = adStateClosed Then
           .Open 'colocar este codigo a todos los informes
           .Requery
           .MoveFirst
           strSQL = Consulta.GetFilterString
           End If
           If strSQL <> "" Then
                    .Filter = strSQL
           Else
                    .Filter = ""
           End If
           strOrden = Consulta.GetOrder
           If strOrden <> "" Then
            .Sort = strOrden
           End If
            rpt_doctores.Show
            rpt_doctores.Refresh
           If .State = adStateOpen Then 'colocar este codigo a todos los informes
                .Close
           End If
    End With
Consulta.Cerrar
End If
End Sub

Private Sub prnEspecialidad_Click()
Dim Consulta As frmConsult
Dim strSQL As String
Dim strOrden As String
Set Consulta = New frmConsult
Consulta.Modal = True
Consulta.Caption = "Consulta de Informe de Especialidad"
Consulta.Top = 1050
Set Consulta.ActiveConection = Cnn
Consulta.Consulta = "select * from especialidad"
Consulta.Show vbModal
If Not Consulta.Descargado Then
    With DataGaleno.rsespecialidad
           If .State = adStateClosed Then  'colocar este codigo a todos los informes
           .Open
           .Requery
           .MoveFirst
           '.Filter = Consulta.msGrid.Columns(0).Caption & " = " & (Consulta.msGrid.Columns(0).Value)
           '.Filter = Consulta.msGrid.Columns(1).Caption & " = " & (Consulta.msGrid.Columns(1).Value)
           strSQL = Consulta.GetFilterString
           End If
           If strSQL <> "" Then
                    .Filter = strSQL
           Else
                    .Filter = ""
           End If
           strOrden = Consulta.GetOrder
           If strOrden <> "" Then
            .Sort = strOrden
           End If
            rpt_especialidad.Show
            rpt_especialidad.Refresh
            'DataEnvironment1.rspaciente.Close
            If .State = adStateOpen Then 'colocar este codigo a todos los informes
               .Close
            End If
    End With
Consulta.Cerrar
End If
End Sub

Private Sub prnInstituto_Click()
Dim Consulta As frmConsult
Dim strSQL As String
Dim strOrden As String
Set Consulta = New frmConsult
Consulta.Modal = True
Consulta.Caption = "Consulta de Informe de Instituto"
Consulta.Top = 1050
Set Consulta.ActiveConection = Cnn
Consulta.Consulta = "SELECT a.CODINS, a.INSTITUTO, a.DOMICILIO, b.DESCRIPCION, a.TELEF1, a.TELEF2 FROM dbo.INSTITUTOS a INNER JOIN dbo.PARAMETROS b ON a.CODCIU = b.CODIGO"
Consulta.Show vbModal
If Not Consulta.Descargado Then
    With DataGaleno.rsinstituto
           If .State = adStateClosed Then  'colocar este codigo a todos los informes
           .Open
           .Requery
           .MoveFirst
           '.Filter = Consulta.msGrid.Columns(0).Caption & " = " & (Consulta.msGrid.Columns(0).Value)
           '.Filter = Consulta.msGrid.Columns(1).Caption & " = " & (Consulta.msGrid.Columns(1).Value)
           strSQL = Consulta.GetFilterString
           End If
           If strSQL <> "" Then
                    .Filter = strSQL
           Else
                    .Filter = ""
           End If
           strOrden = Consulta.GetOrder
           If strOrden <> "" Then
            .Sort = strOrden
           End If
            rpt_institutos.Show
            rpt_institutos.Refresh
            'DataEnvironment1.rspaciente.Close
            If .State = adStateOpen Then 'colocar este codigo a todos los informes
                .Close
            End If
    End With
Consulta.Cerrar
End If
End Sub

Private Sub prnInsumos_Click()
Dim Consulta As frmConsult
Dim strSQL As String
Dim strOrden As String
Set Consulta = New frmConsult
Consulta.Modal = True
Consulta.Caption = "Consulta de Informe de Insumos"
Consulta.Top = 1050
Set Consulta.ActiveConection = Cnn
Consulta.Consulta = "SELECT TOP 100 PERCENT a.COD_INSUMO, a.TIPO_INSUMO, a.NOM_INSUMO, b.NOM_PROVEEDOR, c.LIN_NOMBRE, a.COSTO, a.ARANCEL " _
                    & " FROM dbo.INSUMOS a INNER JOIN " _
                    & " dbo.PROVEEDORES b ON a.COD_PROVEEDOR = b.COD_PROVEEDOR INNER JOIN " _
                    & " dbo.LINEAS c ON a.COD_LINEA = c.COD_LINEA " _
                    & " ORDER BY a.COD_INSUMO"
Consulta.Show vbModal
If Not Consulta.Descargado Then
    With DataGaleno.rsinsumos
           If .State = adStateClosed Then 'colocar este codigo a todos los informes
           .Open
           .Requery
           .MoveFirst
           '.Filter = Consulta.msGrid.Columns(0).Caption & " = " & (Consulta.msGrid.Columns(0).Value)
           '.Filter = Consulta.msGrid.Columns(1).Caption & " = " & (Consulta.msGrid.Columns(1).Value)
           strSQL = Consulta.GetFilterString
           End If
           If strSQL <> "" Then
                    .Filter = strSQL
           Else
                    .Filter = ""
           End If
           strOrden = Consulta.GetOrder
           If strOrden <> "" Then
            .Sort = strOrden
           End If
            rpt_insumos.Show
            rpt_insumos.Refresh
            'DataEnvironment1.rspaciente.Close
            If .State = adStateOpen Then 'colocar este codigo a todos los informes
                .Close
            End If
    End With
Consulta.Cerrar
End If
End Sub

Private Sub prnLineas_Click()
Dim Consulta As frmConsult
Dim strSQL As String
Dim strOrden As String
Set Consulta = New frmConsult
Consulta.Modal = True
Consulta.Caption = "Consulta de Informe de Líneas"
Consulta.Top = 1050
Set Consulta.ActiveConection = Cnn
Consulta.Consulta = "select * from lineas"
Consulta.Show vbModal
If Not Consulta.Descargado Then
    With DataGaleno.rslineas
           If .State = adStateClosed Then  'colocar este codigo a todos los informes
           .Open
           .Requery
           .MoveFirst
           '.Filter = Consulta.msGrid.Columns(0).Caption & " = " & (Consulta.msGrid.Columns(0).Value)
           '.Filter = Consulta.msGrid.Columns(1).Caption & " = " & (Consulta.msGrid.Columns(1).Value)
           strSQL = Consulta.GetFilterString
           End If
           If strSQL <> "" Then
                    .Filter = strSQL
           Else
                    .Filter = ""
           End If
           strOrden = Consulta.GetOrder
           If strOrden <> "" Then
            .Sort = strOrden
           End If
            rpt_lineas.Show
            rpt_lineas.Refresh
            'DataEnvironment1.rspaciente.Close
            If .State = adStateOpen Then 'colocar este codigo a todos los informes
                .Close
            End If
    End With
Consulta.Cerrar
End If
End Sub

Private Sub prnMatUsados_Click()
Dim Consulta As frmConsult
Dim strSQL As String
Dim strOrden As String
Set Consulta = New frmConsult
Consulta.Modal = True
Consulta.Caption = "Consulta de Informe de Materiales Usados"
Consulta.Top = 1050
Set Consulta.ActiveConection = Cnn
Consulta.Consulta = "select * from usomaterial"
Consulta.Show vbModal
If Not Consulta.Descargado Then
    With DataEnvironment1.rsUSOMATERIAL1
           If .State = adStateClosed Then
                    .Open 'colocar este codigo a todos los informes
                    .Requery
                    .MoveFirst
                    '.Filter = Consulta.msGrid.Columns(0).Caption & " = " & (Consulta.msGrid.Columns(0).Value)
                    '.Filter = Consulta.msGrid.Columns(1).Caption & " = " & (Consulta.msGrid.Columns(1).Value)
                    strSQL = Consulta.GetFilterString
           End If
           If strSQL <> "" Then
                    .Filter = strSQL
           Else
                    .Filter = ""
           End If
           strOrden = Consulta.GetOrder
           If strOrden <> "" Then
                    .Sort = strOrden
           End If
                    DataReport15.Show
                    'DataReport15.Refresh
                    'DataEnvironment1.rspersonal.Close
           
            If .State = adStateOpen Then 'colocar este codigo a todos los informes
                .Close
                strSQL = ""
            End If
           
    End With
Consulta.Cerrar
End If
End Sub

Private Sub prnPacientes_Click()
Dim Consulta As frmConsult
Dim strSQL As String
Dim strOrden As String
Set Consulta = New frmConsult
Consulta.Modal = True
Consulta.Caption = "Consulta de Informe de Pacientes"
Consulta.Top = 1050
Set Consulta.ActiveConection = Cnn
Consulta.Consulta = "SELECT     PERSONAS.TIPDOC, PERSONAS.NRODOC, PERSONAS.NOMBRES, PERSONAS.APELLIDOS, PERSONAS.FECNAC, PERSONAS.APECASADA, " _
                    & " PERSONAS.TELEFONOS , PERSONAS.DOM_PARTICULAR, PERSONAS.DOM_LABORAL, PERSONAS.OBSERVACION " _
                    & " FROM PERSONAS INNER JOIN FICHACLINICA ON PERSONAS.TIPDOC = FICHACLINICA.TIPDOC AND PERSONAS.NRODOC = FICHACLINICA.NRODOC "
Consulta.Show vbModal
If Not Consulta.Descargado Then
    With DataGaleno.rspacientes
           If .State = adStateClosed Then
                    .Open 'colocar este codigo a todos los informes
                    .Requery
                    .MoveFirst
                    '.Filter = Consulta.msGrid.Columns(0).Caption & " = " & (Consulta.msGrid.Columns(0).Value)
                    '.Filter = Consulta.msGrid.Columns(1).Caption & " = " & (Consulta.msGrid.Columns(1).Value)
                    strSQL = Consulta.GetFilterString
           End If
           If strSQL <> "" Then
                   .Filter = strSQL
           Else
                    .Filter = ""
           End If
           strOrden = Consulta.GetOrder
           If strOrden <> "" Then
                    .Sort = strOrden
           End If
           rpt_pacientes.Show
           rpt_pacientes.Refresh
           'DataEnvironment1.rspaciente.Close
           If .State = adStateOpen Then 'colocar este codigo a todos los informes
                    .Close
                    strSQL = ""
           End If
    End With
Consulta.Cerrar
End If
End Sub

Private Sub prnPedidos_Click()
Dim Consulta As frmConsult
Dim strSQL As String
Dim strOrden As String
Set Consulta = New frmConsult
Consulta.Modal = True
Consulta.Caption = "Consulta de Informe de Pedidos"
Consulta.Top = 1050
Set Consulta.ActiveConection = Cnn
Consulta.Consulta = "select * from pedIDOS"
Consulta.Show vbModal
If Not Consulta.Descargado Then
    With DataEnvironment1.rsPEDIDOS
           If .State = adStateClosed Then
                    .Open 'colocar este codigo a todos los informes
                    .Requery
                    .MoveFirst
                    '.Filter = Consulta.msGrid.Columns(0).Caption & " = " & (Consulta.msGrid.Columns(0).Value)
                    '.Filter = Consulta.msGrid.Columns(1).Caption & " = " & (Consulta.msGrid.Columns(1).Value)
                    strSQL = Consulta.GetFilterString
           End If
           If strSQL <> "" Then
                    .Filter = strSQL
           Else
                    .Filter = ""
           End If
           strOrden = Consulta.GetOrder
           If strOrden <> "" Then
                    .Sort = strOrden
           End If
           DataReport16.Show
           DataReport16.Refresh
                    'DataEnvironment1.rspersonal.Close
           If .State = adStateOpen Then 'colocar este codigo a todos los informes
                .Close
                strSQL = ""
            End If
           
    End With
Consulta.Cerrar
End If

End Sub

Private Sub prnservidoct_Click()
Dim Consulta As frmConsult
Dim strSQL As String
Dim strOrden As String
Set Consulta = New frmConsult
Consulta.Modal = True
Consulta.Caption = "Consulta de Informe de Servicios Prestados por Doctores"
Consulta.Top = 1050
Set Consulta.ActiveConection = Cnn
Consulta.Consulta = "SELECT SERVICIOINTERNACIONDETALLE.FICHACLINICA, SERVICIOINTERNACIONDETALLE.FEC_PED, SERVICIOINTERNACIONDETALLE.NROMOV, SERVICIOINTERNACIONDETALLE.TIPDOC, SERVICIOINTERNACIONDETALLE.NRODOC, PERSONAS.NOMBRES, PERSONAS.APELLIDOS, PERSONAS.APECASADA, PERSONAS.TITULO, SERVICIOINTERNACIONDETALLE.CODSERVICIO, SERVICIOINTERNACIONDETALLE.TIPOSER, SERVICIOINTERNACIONDETALLE.CODSER, SERVICIOINTERNACIONDETALLE.NOMSERVICIO, SERVICIOINTERNACIONDETALLE.CANTIDAD, SERVICIOINTERNACIONDETALLE.ARANCEL, SERVICIOINTERNACIONDETALLE.COSTO, SERVICIOINTERNACIONDETALLE.CODDR, PERSONAS1.NOMBRES AS NOMDOC, PERSONAS1.APELLIDOS AS APEDOC, PERSONAS1.APECASADA AS CASDOC, PERSONAS1.TITULO AS TITDOC, DOCTOR.REGPRF, SERVICIOINTERNACIONDETALLE.HONORARIOS, SERVICIOINTERNACIONDETALLE.FECHA_MOVIMIENTO, SERVICIOINTERNACIONDETALLE.OBSERVACIONES " _
                    & " FROM PERSONAS PERSONAS1 INNER JOIN DOCTOR ON PERSONAS1.TIPDOC = DOCTOR.TIPDOC AND PERSONAS1.NRODOC = DOCTOR.NRODOC INNER JOIN SERVICIOINTERNACIONDETALLE INNER JOIN PERSONAS ON SERVICIOINTERNACIONDETALLE.TIPDOC = PERSONAS.TIPDOC AND SERVICIOINTERNACIONDETALLE.NRODOC = PERSONAS.NRODOC ON DOCTOR.CODDR = SERVICIOINTERNACIONDETALLE.CODDR "
Consulta.Show vbModal
If Not Consulta.Descargado Then
    With DataGaleno.rsLIQ_DOCTOR_Grouping
           If .State = adStateClosed Then 'colocar este codigo a todos los informes
           .Open
           .Requery
           .MoveFirst
           '.Filter = Consulta.msGrid.Columns(0).Caption & " = " & (Consulta.msGrid.Columns(0).Value)
           '.Filter = Consulta.msGrid.Columns(1).Caption & " = " & (Consulta.msGrid.Columns(1).Value)
           strSQL = Consulta.GetFilterString
           End If
           If strSQL <> "" Then
                    .Filter = strSQL
           Else
                    .Filter = ""
           End If
           strOrden = Consulta.GetOrder
           If strOrden <> "" Then
            .Sort = strOrden
           End If
            rpt_liq_doctor.Show
            rpt_liq_doctor.Refresh
            'DataEnvironment1.rspaciente.Close
            If .State = adStateOpen Then 'colocar este codigo a todos los informes
                .Close
            End If
    End With
Consulta.Cerrar
End If
End Sub

Private Sub prnserviinst_Click()
Dim Consulta As frmConsult
Dim strSQL As String
Dim strOrden As String
Set Consulta = New frmConsult
Consulta.Modal = True
Consulta.Caption = "Consulta de Informe de Servicios Prestados por Institutos"
Consulta.Top = 1050
Set Consulta.ActiveConection = Cnn
Consulta.Consulta = " SELECT ORDENSERVICIOCABECERA.FICHACLINICA, ORDENSERVICIOCABECERA.FEC_PED, ORDENSERVICIOCABECERA.NROMOV, ORDENSERVICIOCABECERA.TIPDOC, ORDENSERVICIOCABECERA.NRODOC, PERSONAS.NOMBRES, PERSONAS.APELLIDOS, PERSONAS.TITULO, PERSONAS.FECNAC, PERSONAS.GENERO, PERSONAS.TELEFONOS, ORDENSERVICIOCABECERA.CODDR, PERSONAS_2.NOMBRES AS NOMDOC, PERSONAS_2.APELLIDOS AS APEDOC, ORDENSERVICIOCABECERA.COD_ESPECIALIDAD, ESPECIALIDAD.NOMBRE, ORDENSERVICIOCABECERA.COD_EMPRESA, EMPRESA.NOMBRE AS NOMEMPR, ORDENSERVICIOCABECERA.NRO_SEGURO, ORDENSERVICIOCABECERA.IDE_ASEGURADO, ORDENSERVICIOCABECERA.TIPO_SEGURO, ORDENSERVICIOCABECERA.TIPOMOV, TIPOMOV.NOMTIPOMOV, ORDENSERVICIOCABECERA.RESPECOTIPDOC, ORDENSERVICIOCABECERA.RESPECONRODOC, PERSONAS_1.NOMBRES AS RESNOM, PERSONAS_1.APELLIDOS AS RESAPE, PERSONAS_1.APECASADA, PERSONAS_1.TITULO AS RESTITUL, PERSONAS_1.TELEFONOS AS RESTEL, ORDENSERVICIOCABECERA.CODINS, INSTITUTOS.INSTITUTO  " _
                  & " FROM ORDENSERVICIOCABECERA INNER JOIN DOCTOR ON ORDENSERVICIOCABECERA.CODDR = DOCTOR.CODDR  INNER JOIN ESPECIALIDAD ON ORDENSERVICIOCABECERA.COD_ESPECIALIDAD = ESPECIALIDAD.COD_ESPECIALIDAD INNER JOIN EMPRESA ON ORDENSERVICIOCABECERA.COD_EMPRESA = EMPRESA.COD_EMPRESA INNER JOIN TIPOMOV ON ORDENSERVICIOCABECERA.TIPOMOV = TIPOMOV.TIPOMOV INNER JOIN PERSONAS ON ORDENSERVICIOCABECERA.TIPDOC = PERSONAS.TIPDOC AND ORDENSERVICIOCABECERA.NRODOC = PERSONAS.NRODOC INNER JOIN INSTITUTOS ON ORDENSERVICIOCABECERA.CODINS = INSTITUTOS.CODINS INNER JOIN PERSONAS PERSONAS_2 ON DOCTOR.TIPDOC = PERSONAS_2.TIPDOC AND DOCTOR.NRODOC = PERSONAS_2.NRODOC LEFT OUTER JOIN PERSONAS PERSONAS_1 ON ORDENSERVICIOCABECERA.RESPECOTIPDOC = PERSONAS_1.TIPDOC AND ORDENSERVICIOCABECERA.RESPECONRODOC = PERSONAS_1.NRODOC "
Consulta.Show vbModal
If Not Consulta.Descargado Then
    With DataGaleno.rsLIQ_INSTITUTOS_Grouping
           If .State = adStateClosed Then 'colocar este codigo a todos los informes
           .Open
           .Requery
           .MoveFirst
           '.Filter = Consulta.msGrid.Columns(0).Caption & " = " & (Consulta.msGrid.Columns(0).Value)
           '.Filter = Consulta.msGrid.Columns(1).Caption & " = " & (Consulta.msGrid.Columns(1).Value)
           strSQL = Consulta.GetFilterString
           End If
           If strSQL <> "" Then
                    .Filter = strSQL
           Else
                    .Filter = ""
           End If
           strOrden = Consulta.GetOrder
           If strOrden <> "" Then
            .Sort = strOrden
           End If
            rpt_liq_institutos.Show
            rpt_liq_institutos.Refresh
            'DataEnvironment1.rspaciente.Close
            If .State = adStateOpen Then 'colocar este codigo a todos los informes
                .Close
            End If
    End With
Consulta.Cerrar
End If
End Sub

Private Sub prnservipac_Click()
Dim Consulta As frmConsult
Dim strSQL As String
Dim strOrden As String
Set Consulta = New frmConsult
Consulta.Modal = True
Consulta.Caption = "Consulta de Informe de Servicios Adquiridos por Pacientes"
Consulta.Top = 1050
Set Consulta.ActiveConection = Cnn
Consulta.Consulta = "SELECT ORDENSERVICIOCABECERA.FICHACLINICA, ORDENSERVICIOCABECERA.FEC_PED, ORDENSERVICIOCABECERA.NROMOV, ORDENSERVICIOCABECERA.TIPDOC, ORDENSERVICIOCABECERA.NRODOC, PERSONAS.NOMBRES, PERSONAS.APELLIDOS, PERSONAS.TITULO, PERSONAS.FECNAC, PERSONAS.GENERO, PERSONAS.TELEFONOS, ORDENSERVICIOCABECERA.CODDR, PERSONAS_2.NOMBRES AS NOMDOC, PERSONAS_2.APELLIDOS AS APEDOC, ORDENSERVICIOCABECERA.COD_ESPECIALIDAD, ESPECIALIDAD.NOMBRE, ORDENSERVICIOCABECERA.COD_EMPRESA, EMPRESA.NOMBRE AS NOMEMPR, ORDENSERVICIOCABECERA.NRO_SEGURO, ORDENSERVICIOCABECERA.IDE_ASEGURADO, ORDENSERVICIOCABECERA.TIPO_SEGURO, ORDENSERVICIOCABECERA.TIPOMOV, TIPOMOV.NOMTIPOMOV, ORDENSERVICIOCABECERA.RESPECOTIPDOC, ORDENSERVICIOCABECERA.RESPECONRODOC, PERSONAS_1.NOMBRES AS RESNOM, PERSONAS_1.APELLIDOS AS RESAPE, PERSONAS_1.APECASADA, PERSONAS_1.TITULO AS RESTITUL, PERSONAS_1.TELEFONOS AS RESTEL, ORDENSERVICIOCABECERA.CODINS, INSTITUTOS.INSTITUTO " _
                   & " FROM ORDENSERVICIOCABECERA INNER JOIN DOCTOR ON ORDENSERVICIOCABECERA.CODDR = DOCTOR.CODDR  INNER JOIN ESPECIALIDAD ON ORDENSERVICIOCABECERA.COD_ESPECIALIDAD = ESPECIALIDAD.COD_ESPECIALIDAD INNER JOIN EMPRESA ON ORDENSERVICIOCABECERA.COD_EMPRESA = EMPRESA.COD_EMPRESA INNER JOIN TIPOMOV ON ORDENSERVICIOCABECERA.TIPOMOV = TIPOMOV.TIPOMOV INNER JOIN PERSONAS ON ORDENSERVICIOCABECERA.TIPDOC = PERSONAS.TIPDOC AND ORDENSERVICIOCABECERA.NRODOC = PERSONAS.NRODOC INNER JOIN INSTITUTOS ON ORDENSERVICIOCABECERA.CODINS = INSTITUTOS.CODINS INNER JOIN PERSONAS PERSONAS_2 ON DOCTOR.TIPDOC = PERSONAS_2.TIPDOC AND DOCTOR.NRODOC = PERSONAS_2.NRODOC LEFT OUTER JOIN PERSONAS PERSONAS_1 ON ORDENSERVICIOCABECERA.RESPECOTIPDOC = PERSONAS_1.TIPDOC AND ORDENSERVICIOCABECERA.RESPECONRODOC = PERSONAS_1.NRODOC "
Consulta.Show vbModal
If Not Consulta.Descargado Then
    With DataGaleno.rsLIQUIDACIONES_Grouping
           If .State = adStateClosed Then 'colocar este codigo a todos los informes
           .Open
           .Requery
           .MoveFirst
           '.Filter = Consulta.msGrid.Columns(0).Caption & " = " & (Consulta.msGrid.Columns(0).Value)
           '.Filter = Consulta.msGrid.Columns(1).Caption & " = " & (Consulta.msGrid.Columns(1).Value)
           strSQL = Consulta.GetFilterString
           End If
           If strSQL <> "" Then
                    .Filter = strSQL
           Else
                    .Filter = ""
           End If
           strOrden = Consulta.GetOrder
           If strOrden <> "" Then
            .Sort = strOrden
           End If
            rpt_liquidaciones.Show
            rpt_liquidaciones.Refresh
            'DataEnvironment1.rspaciente.Close
            If .State = adStateOpen Then 'colocar este codigo a todos los informes
                .Close
            End If
    End With
Consulta.Cerrar
End If
End Sub

Private Sub prnTabVal_Click()
Dim Consulta As frmConsult
Dim strSQL As String
Dim strOrden As String
Set Consulta = New frmConsult
Consulta.Modal = True
Consulta.Caption = "Consulta de Informe de Tabla de Valores"
Consulta.Top = 1050
Set Consulta.ActiveConection = Cnn
Consulta.Consulta = "SELECT TOP 100 PERCENT a.CODTAB, a.NOMBRE, a.CODITEM, b.ITEM_DESCRIPCION FROM dbo.TABVAL a INNER JOIN dbo.ITEM_HISTORIACLINICA_ESPECIALI b ON a.CODITEM = b.ITEM ORDER BY a.CODTAB"
Consulta.Show vbModal
If Not Consulta.Descargado Then
    With DataGaleno.rstab_valores
           If .State = adStateClosed Then
                .Open 'colocar este codigo a todos los informes
                .Requery
                .MoveFirst
                '.Filter = Consulta.msGrid.Columns(0).Caption & " = " & (Consulta.msGrid.Columns(0).Value)
                '.Filter = Consulta.msGrid.Columns(1).Caption & " = " & (Consulta.msGrid.Columns(1).Value)
                strSQL = Consulta.GetFilterString
           End If
           If strSQL <> "" Then
                    .Filter = strSQL
           Else
                    .Filter = ""
           End If
           strOrden = Consulta.GetOrder
           If strOrden <> "" Then
               .Sort = strOrden
           End If
                    rpt_tab_valores.Show
                    rpt_tab_valores.Refresh
                  
            If .State = adStateOpen Then 'colocar este codigo a todos los informes
                .Close
                strSQL = ""
            End If
    End With
Consulta.Cerrar
End If
End Sub

Private Sub prnUsuarios_Click()
Dim Consulta As frmConsult
Dim strSQL As String
Dim strOrden As String
Set Consulta = New frmConsult
Consulta.Modal = True
Consulta.Caption = "Consulta de Informe de Usuarios"
Consulta.Top = 1050
Set Consulta.ActiveConection = Cnn
Consulta.Consulta = "SELECT ID_USUARIO, NOMBRE, CLAVE, ADMINISTRADOR FROM dbo.USUARIOS"
Consulta.Show vbModal
If Not Consulta.Descargado Then
    With DataGaleno.rsusuarios
           If .State = adStateClosed Then
                .Open 'colocar este codigo a todos los informes
                .Requery
                .MoveFirst
                '.Filter = Consulta.msGrid.Columns(0).Caption & " = " & (Consulta.msGrid.Columns(0).Value)
                '.Filter = Consulta.msGrid.Columns(1).Caption & " = " & (Consulta.msGrid.Columns(1).Value)
                strSQL = Consulta.GetFilterString
           End If
           If strSQL <> "" Then
                    .Filter = strSQL
           Else
                    .Filter = ""
           End If
           strOrden = Consulta.GetOrder
           If strOrden <> "" Then
                .Sort = strOrden
           End If
           rpt_usuarios.Show
           rpt_usuarios.Refresh
           'DataEnvironment1.rspaciente.Close
            If .State = adStateOpen Then 'colocar este codigo a todos los informes
                .Close
                strSQL = ""
            End If
    End With
Consulta.Cerrar
End If
End Sub

Private Sub prnValnorm_Click()
Dim Consulta As frmConsult
Dim strSQL As String
Dim strOrden As String
Set Consulta = New frmConsult
Consulta.Modal = True
Consulta.Caption = "Consulta de Informe de Valores Normales"
Consulta.Top = 1050
Set Consulta.ActiveConection = Cnn
Consulta.Consulta = "SELECT TOP 100 PERCENT a.CODSERVICIO, a.TIPOSER, b.NOMSERVICIO, a.SEXO, a.EDAD_INF,a.EDAD_SUP, a.LIM_INF, a.LIM_SUP, a.LIM_INFN, a.LIM_SUPN FROM dbo.VALNORM a INNER JOIN dbo.SERVICIO b ON a.CODSERVICIO = b.CODSERVICIO AND a.TIPOSER = b.TIPOSER ORDER BY a.CODSERVICIO"
Consulta.Show vbModal
If Not Consulta.Descargado Then
    With DataGaleno.rsvalores_normales
           If .State = adStateClosed Then 'colocar el codigo siguiente a todos los informe
                .Open
           Else
            .Requery
           End If
           .MoveFirst
           '.Filter = Consulta.msGrid.Columns(0).Caption & " = " & (Consulta.msGrid.Columns(0).Value)
           '.Filter = Consulta.msGrid.Columns(1).Caption & " = " & (Consulta.msGrid.Columns(1).Value)
           strSQL = Consulta.GetFilterString
           If strSQL <> "" Then
                    .Filter = strSQL
           Else
                    .Filter = ""
           End If
           strOrden = Consulta.GetOrder
           If strOrden <> "" Then
                .Sort = strOrden
           End If
           rpt_valor_normal.Show
           rpt_valor_normal.Refresh
           If .State = adStateOpen Then 'colocar este codigo a todos los informes
                    .Close
                    strSQL = ""
           End If
    End With
Consulta.Cerrar
End If
End Sub

Private Sub resultados_Click()
Dim Consulta As frmConsult
Dim strSQL As String
Dim strOrden As String
Set Consulta = New frmConsult
Consulta.Modal = True
Consulta.Caption = "Consulta de Informe de Resultado"
Consulta.Top = 1050
Set Consulta.ActiveConection = Cnn
Consulta.Consulta = "SELECT MOVANA1.FEC_PED, MOVANA1.NROMOV, MOVANA1.NOMDOC, MOVANA1.ID_SEGURO, MOVANA1.NOMPAC, MOVANA1.EDAD, MOVANA1.SEXO, MOVANA1.NRODOC from movana1"
Consulta.Show vbModal
If Not Consulta.Descargado Then
    With DataEnvironment1.rsRESULTADOS_Grouping
           If .State = adStateClosed Then
                .Open 'colocar este codigo a todos los informes
                .Filter = ""
                .Requery
                .MoveFirst
                .Filter = Consulta.msGrid.Columns(0).Caption & " = " & (Consulta.msGrid.Columns(0).Value)
                .Filter = Consulta.msGrid.Columns(1).Caption & " = " & (Consulta.msGrid.Columns(1).Value)
                strSQL = Consulta.GetFilterString
           End If
           'If strSQL <> "" Then
           '         .Filter = strSQL
           'Else
           '         .Filter = ""
           'End If
           strOrden = Consulta.GetOrder
           If strOrden <> "" Then
                .Sort = strOrden
           End If
            DataReport10.Show
            DataReport10.Refresh
            'DataEnvironment1.rsRESULTADOS_Grouping.Close
            If .State = adStateOpen Then 'colocar este codigo a todos los informes
                .Close
                strSQL = ""
            End If
    End With
Consulta.Cerrar
End If
End Sub

Private Sub serinter_Click()
   SERVICIOS.Show
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
  On Error Resume Next
    Select Case Button.Key
        Case "Last"
            Select Case MDIGALENO.ActiveForm.Name
                Case "ORDENES"
                    ORDENES.UltimoReg
                Case "SERVICIOS"
                    SERVICIOS.UltimoReg
                Case "MANTENIMIENTO"
                    MANTENIMIENTO.prcUltimo
                Case "ATENCIONMEDICA"
                    ATENCIONMEDICA.UltimoReg
                Case "frmUSO_DE_MATERIALES"
                    frmUSO_DE_MATERIALES.prcLast
                Case "frmPedFact"
                    frmPedFact.ultimo
                Case "frmPEDIDOS"
                    frmPEDIDOS.prcLast
                Case "LIQUIDACIONES"
                    LIQUIDACIONES.UltimoReg
            End Select
        Case "Next"
            Select Case MDIGALENO.ActiveForm.Name
                Case "ORDENES"
                    ORDENES.SiguienteReg
                Case "SERVICIOS"
                    SERVICIOS.SiguienteReg
                Case "MANTENIMIENTO"
                    MANTENIMIENTO.prcSiguiente
               Case "ATENCIONMEDICA"
                    ATENCIONMEDICA.SiguienteReg
                Case "frmUSO_DE_MATERIALES"
                    frmUSO_DE_MATERIALES.prcNext
                Case "frmPedFact"
                    frmPedFact.Siguiente
                Case "frmPEDIDOS"
                    frmPEDIDOS.prcNext
                Case "LIQUIDACIONES"
                    LIQUIDACIONES.SiguienteReg
            End Select
        Case "Previous"
            Select Case MDIGALENO.ActiveForm.Name
                Case "ORDENES"
                    ORDENES.AnteriorReg
                 Case "SERVICIOS"
                    SERVICIOS.AnteriorReg
                Case "MANTENIMIENTO"
                    MANTENIMIENTO.prcAnterior
                Case "ATENCIONMEDICA"
                    ATENCIONMEDICA.AnteriorReg
                Case "frmUSO_DE_MATERIALES"
                    frmUSO_DE_MATERIALES.prcPrev
                Case "frmPedFact"
                    frmPedFact.Anterior
                Case "frmPEDIDOS"
                    frmPEDIDOS.prcPrevious
                Case "LIQUIDACIONES"
                    LIQUIDACIONES.AnteriorReg
                End Select
        Case "First"
            Select Case MDIGALENO.ActiveForm.Name
                Case "ORDENES"
                    ORDENES.PrimeroReg
                Case "SERVICIOS"
                    SERVICIOS.PrimeroReg
                Case "MANTENIMIENTO"
                    MANTENIMIENTO.prcPrimero
                Case "ATENCIONMEDICA"
                    ATENCIONMEDICA.PrimeroReg
                Case "frmUSO_DE_MATERIALES"
                    frmUSO_DE_MATERIALES.prcFirst
                Case "frmPedFact"
                    frmPedFact.Primero
                Case "frmPEDIDOS"
                    frmPEDIDOS.prcFirst
                Case "LIQUIDACIONES"
                    LIQUIDACIONES.PrimeroReg
            End Select
        Case "Search"
            Select Case MDIGALENO.ActiveForm.Name
                Case "ORDENES"
                    ORDENES.buscarReg
                Case "SERVICIOS"
                    SERVICIOS.buscarReg
                Case "MANTENIMIENTO"
                    MANTENIMIENTO.prcBuscar
                Case "ATENCIONMEDICA"
                    ATENCIONMEDICA.buscarReg
                Case "frmUSO_DE_MATERIALES"
                    frmUSO_DE_MATERIALES.prcBuscar
                Case "frmPedFact"
                    frmPedFact.BUSCAR
                Case "frmPEDIDOS"
                    frmPEDIDOS.prcayuda1
                Case "LIQUIDACIONES"
                    LIQUIDACIONES.buscarReg
            End Select
        Case "Add"
            Select Case MDIGALENO.ActiveForm.Name
                Case "ORDENES"
                    ORDENES.AgregarReg
                    If ORDENES.TAB_ORDENES.SelectedItem.Index <> 2 Then
                        fncBotones False, ""
                    End If
               Case "SERVICIOS"
                    SERVICIOS.AgregarReg
                    fncBotones False, ""
               Case "MANTENIMIENTO"
                    MANTENIMIENTO.prcAgregar
                    fncBotones False, ""
               Case "ATENCIONMEDICA"
                     fncBotones False, "ATENCIONMEDICA"
'                   ATENCIONMEDICA.AgregarReg
                Case "frmPedFact"
                    frmPedFact.Agregar
                Case "frmPEDIDOS"
                    frmPEDIDOS.prcAdd
                Case "LIQUIDACIONES"
                    fncBotones False, ""
                    LIQUIDACIONES.AgregarReg
                    
            End Select
        Case "Save"
            Select Case MDIGALENO.ActiveForm.Name
                Case "ORDENES"
                    ORDENES.ActualizarReg

               Case "SERVICIOS"
                    fncBotones True, ""
                    SERVICIOS.ActualizarReg
               Case "MANTENIMIENTO"
                    fncBotones True, ""
                    MANTENIMIENTO.prcGrabar
               Case "ATENCIONMEDICA"
                    fncBotones True, "ATENCIONMEDICA"
                    ATENCIONMEDICA.ActualizarReg
               Case "frmUSO_DE_MATERIALES"
                    frmUSO_DE_MATERIALES.prcUpdate
               Case "frmPedFact"
                    frmPedFact.Grabar
               Case "frmPEDIDOS"
                    frmPEDIDOS.prcUpdate
               Case "LIQUIDACIONES"
               fncBotones True, ""
                    LIQUIDACIONES.ActualizarReg
            End Select
        Case "Delete"
            Select Case MDIGALENO.ActiveForm.Name
                Case "ORDENES"
                    If ORDENES.TAB_ORDENES.SelectedItem.Index <> 2 Then
                        ORDENES.eliminarReg
                    End If
                Case "SERVICIOS"
                    SERVICIOS.eliminarReg
                Case "MANTENIMIENTO"
                    MANTENIMIENTO.prcEliminar
                Case "ATENCIONMEDICA"
                    ATENCIONMEDICA.eliminarReg
                Case "frmUSO_DE_MATERIALES"
                    frmUSO_DE_MATERIALES.prcDelete
                Case "frmPedFact"
                    frmPedFact.eliminar
                Case "frmPEDIDOS"
                    frmPEDIDOS.prcDelete
                Case "LIQUIDACIONES"
                    LIQUIDACIONES.eliminarReg
            End Select
        Case "Cancel"
            Select Case MDIGALENO.ActiveForm.Name
                Case "ORDENES"
                    fncBotones True, ""
                    ORDENES.CancelarReg
                Case "SERVICIOS"
                    fncBotones True, ""
                    SERVICIOS.CancelarReg
                Case "MANTENIMIENTO"
                    fncBotones True, ""
                    MANTENIMIENTO.prcCancelar
                Case "frmUSO_DE_MATERIALES"
                    frmUSO_DE_MATERIALES.prcCancel
                Case "ATENCIONMEDICA"
                    fncBotones True, "ATENCIONMEDICA"
                    ATENCIONMEDICA.CancelarReg
                Case "frmPedFact"
                    frmPedFact.CANCELAR
                Case "frmPEDIDOS"
                    frmPEDIDOS.prcCancel
                Case "LIQUIDACIONES"
                    fncBotones True, ""
                    LIQUIDACIONES.CancelarReg
                End Select
        Case "MODIFICAR"
            Select Case MDIGALENO.ActiveForm.Name
                'fncBotones (False)
                Case "ORDENES"
                    fncBotones False, ""
                    ORDENES.ModificarReg
                Case "SERVICIOS"
                    fncBotones False, ""
                    SERVICIOS.ModificarReg
                Case "MANTENIMIENTO"
                    fncBotones False, ""
                    MANTENIMIENTO.prcModificar
                Case "ATENCIONMEDICA"
                    fncBotones False, "ATENCIONMEDICA"
                    ATENCIONMEDICA.ModificarReg
                Case "frmUSO_DE_MATERIALES"
                    frmUSO_DE_MATERIALES.prcEdit
                Case "frmPedFact"
                    frmPedFact.Editar
                Case "frmPEDIDOS"
                    frmPEDIDOS.prcEdit
                Case "LIQUIDACIONES"
                    fncBotones False, ""
                    LIQUIDACIONES.ModificarReg
            End Select
        Case "Actualizar"
            Select Case MDIGALENO.ActiveForm.Name
                Case "ORDENES"
                    fncBotones True, ""
                    ORDENES.requeryReg
                Case "SERVICIOS"
                    fncBotones True, ""
                    SERVICIOS.requeryReg
                Case "MANTENIMIENTO"
                    fncBotones True, ""
                    MANTENIMIENTO.prcActualizar
                Case "ATENCIONMEDICA"
                    fncBotones True, "ATENCIONMEDICA"
                    ATENCIONMEDICA.requeryReg
                Case "frmUSO_DE_MATERIALES"
                    frmUSO_DE_MATERIALES.prcActualizar
                Case "frmPEDIDOS"
                    frmPEDIDOS.prcRefresh
                Case "LIQUIDACIONES"
                    LIQUIDACIONES.requeryReg
            End Select
        Case "Print"
            Select Case MDIGALENO.ActiveForm.Name
                Case "ORDENES"
                    Call ORDENES.imprimir
                Case "SERVICIOS"
                    Call SERVICIOS.imprimir
                Case "ATENCIONMEDICA"
                    Call ATENCIONMEDICA.imprimir
                Case "LIQUIDACIONES"
                    Call LIQUIDACIONES.imprimir

            End Select
            
        Case "Exit"
            Unload MDIGALENO
    End Select
End Sub

Public Function fncBotones(X As Boolean, MODULO As String)

    tbToolBar.Buttons("Cut").Enabled = X
    tbToolBar.Buttons("Copy").Enabled = X
    tbToolBar.Buttons("Paste").Enabled = X
    tbToolBar.Buttons("Search").Enabled = X
    tbToolBar.Buttons("MODIFICAR").Enabled = X
    tbToolBar.Buttons("Actualizar").Enabled = X
    If MODULO <> "ATENCIONMEDICA" Then
        tbToolBar.Buttons("Add").Enabled = X
        tbToolBar.Buttons("First").Enabled = X
        tbToolBar.Buttons("Previous").Enabled = X
        tbToolBar.Buttons("Next").Enabled = X
        tbToolBar.Buttons("Last").Enabled = X
        tbToolBar.Buttons("Delete").Enabled = X
    Else
        tbToolBar.Buttons("Add").Enabled = False
        tbToolBar.Buttons("First").Enabled = False
        tbToolBar.Buttons("Previous").Enabled = False
        tbToolBar.Buttons("Next").Enabled = False
        tbToolBar.Buttons("Last").Enabled = False
        tbToolBar.Buttons("Delete").Enabled = False
    End If
    'tbToolBar.Buttons("Save").Enabled = X
    tbToolBar.Buttons("Print").Enabled = X
    'tbToolBar.Buttons("Cancel").Enabled = X
    tbToolBar.Buttons("Exit").Enabled = X
End Function

