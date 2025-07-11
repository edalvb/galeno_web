VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form SERVICIOS 
   Caption         =   "SERVICIOS HOSPITALARIOS"
   ClientHeight    =   6525
   ClientLeft      =   1335
   ClientTop       =   2340
   ClientWidth     =   10875
   Icon            =   "SERVICIOINTERNACION.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6525
   ScaleWidth      =   10875
   Begin VB.Frame frmServiMe 
      BackColor       =   &H80000013&
      ForeColor       =   &H80000002&
      Height          =   6435
      Left            =   120
      TabIndex        =   10
      Top             =   360
      Visible         =   0   'False
      Width           =   11385
      Begin VB.CommandButton CmdAltaPaciente 
         BackColor       =   &H8000000D&
         Caption         =   "Ingresar a la Alta  del Paciente"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8160
         MaskColor       =   &H8000000D&
         TabIndex        =   39
         Top             =   1920
         UseMaskColor    =   -1  'True
         Width           =   1815
      End
      Begin VB.CommandButton cmdagregarserord 
         DragIcon        =   "SERVICIOINTERNACION.frx":030A
         Height          =   375
         Left            =   120
         Picture         =   "SERVICIOINTERNACION.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1920
         Width           =   495
      End
      Begin VB.CommandButton cmdagregarservicios 
         DragIcon        =   "SERVICIOINTERNACION.frx":0B8E
         Height          =   375
         Left            =   120
         Picture         =   "SERVICIOINTERNACION.frx":0FD0
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   2280
         Width           =   495
      End
      Begin VB.CommandButton CmdBuscarCama 
         DragIcon        =   "SERVICIOINTERNACION.frx":1412
         Height          =   375
         Left            =   7560
         Picture         =   "SERVICIOINTERNACION.frx":1854
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "COD_TIPO"
         Height          =   285
         Index           =   9
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   34
         Top             =   1440
         Width           =   945
      End
      Begin VB.CommandButton CmdAgregarIns 
         DragIcon        =   "SERVICIOINTERNACION.frx":1C96
         Height          =   375
         Left            =   3240
         Picture         =   "SERVICIOINTERNACION.frx":20D8
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   2280
         Width           =   495
      End
      Begin VB.CommandButton cmdagregarmed 
         DragIcon        =   "SERVICIOINTERNACION.frx":251A
         Height          =   375
         Left            =   3240
         Picture         =   "SERVICIOINTERNACION.frx":295C
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1920
         Width           =   495
      End
      Begin VB.CommandButton busfecha 
         DragIcon        =   "SERVICIOINTERNACION.frx":2D9E
         Height          =   375
         Left            =   120
         Picture         =   "SERVICIOINTERNACION.frx":31E0
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton CmdFichaClinica 
         BackColor       =   &H8000000D&
         Caption         =   "Ingresar a la Ficha Clinica del Paciente"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5760
         MaskColor       =   &H8000000D&
         TabIndex        =   26
         Top             =   1920
         UseMaskColor    =   -1  'True
         Width           =   2055
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "COD_CAMA"
         Height          =   285
         Index           =   8
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1440
         Width           =   1065
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "COD_HABITACION"
         Height          =   285
         Index           =   7
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1440
         Width           =   945
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "FEC_INTERNACION"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   9480
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   5
         Top             =   960
         Width           =   1065
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "FICHACLINICA"
         ForeColor       =   &H000000C0&
         Height          =   285
         Index           =   2
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "TIPDOC"
         Height          =   285
         Index           =   3
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "NRODOC"
         Height          =   285
         Index           =   4
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   600
         Width           =   1245
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "NROMOV"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Index           =   1
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "FEC_PED"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   3
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Index           =   0
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CmdbusHabitacion 
         DragIcon        =   "SERVICIOINTERNACION.frx":3622
         Height          =   375
         Left            =   4080
         Picture         =   "SERVICIOINTERNACION.frx":3A64
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "TIPOMOV"
         Height          =   285
         Index           =   6
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   960
         Width           =   945
      End
      Begin VB.CommandButton cmdBusTipo 
         DragIcon        =   "SERVICIOINTERNACION.frx":3EA6
         Height          =   375
         Left            =   120
         Picture         =   "SERVICIOINTERNACION.frx":42E8
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1320
         Width           =   495
      End
      Begin MSDataGridLib.DataGrid grd_servicios 
         Height          =   3495
         Left            =   120
         TabIndex        =   13
         Top             =   2880
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   6165
         _Version        =   393216
         AllowUpdate     =   -1  'True
         ForeColor       =   -2147483635
         HeadLines       =   1
         RowHeight       =   19
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "DETALLE DE SERVICIOS"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   15370
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   15370
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label LBLServicio 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Agregar Servicios de la Orden"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   15
         Left            =   720
         TabIndex        =   38
         Top             =   2040
         Width           =   2355
      End
      Begin VB.Label LBLServicio 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Movimiento"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   14
         Left            =   600
         TabIndex        =   33
         Top             =   960
         Width           =   1530
      End
      Begin VB.Label LBLServicio 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Agregar Insumos"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   12
         Left            =   3720
         TabIndex        =   31
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label LBLServicio 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Agregar Medicamentos"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   11
         Left            =   3720
         TabIndex        =   29
         Top             =   2040
         Width           =   1800
      End
      Begin VB.Label LBLServicio 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Código  de Cama"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   8
         Left            =   8160
         TabIndex        =   25
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label LBLServicio 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Código  de Habitacion"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   7
         Left            =   4680
         TabIndex        =   24
         Top             =   1440
         Width           =   1725
      End
      Begin VB.Label LBLServicio 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Internación"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   5
         Left            =   8040
         TabIndex        =   23
         Top             =   960
         Width           =   1410
      End
      Begin VB.Label LBLServicio 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ficha Clinica"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   2
         Left            =   8400
         TabIndex        =   22
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label LBLServicio 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo y Nro de Documento"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   1950
      End
      Begin VB.Label LBLServicio 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00D2BBA0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Index           =   10
         Left            =   6480
         TabIndex        =   20
         Top             =   600
         Width           =   1965
      End
      Begin VB.Label LBLServicio 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Numero de Orden"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   1
         Left            =   5040
         TabIndex        =   19
         Top             =   240
         Width           =   1410
      End
      Begin VB.Label LBLServicio 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Orden"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   18
         Top             =   240
         Width           =   1230
      End
      Begin VB.Label LBLServicio 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Habitación"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   6
         Left            =   720
         TabIndex        =   17
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label LBLServicio 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Agregar Otros Servicios"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   9
         Left            =   720
         TabIndex        =   16
         Top             =   2400
         Width           =   1875
      End
      Begin VB.Label LBLServicio 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre del Paciente"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   4
         Left            =   4800
         TabIndex        =   15
         Top             =   600
         Width           =   1740
      End
      Begin VB.Label LBLServiTIP 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00D2BBA0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Index           =   5
         Left            =   3480
         TabIndex        =   14
         Top             =   960
         Width           =   1965
      End
   End
   Begin MSComctlLib.TabStrip TAB_SERVICIOS 
      Height          =   6930
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   12224
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "INTERNACIONES"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label LBLServicio 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Código  de Tipo"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   13
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   1200
   End
End
Attribute VB_Name = "SERVICIOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As ADODB.Connection
Dim WithEvents adoServiMe As ADODB.Recordset
Attribute adoServiMe.VB_VarHelpID = -1
Dim WithEvents adoServiMeCa As ADODB.Recordset
Attribute adoServiMeCa.VB_VarHelpID = -1
Dim WithEvents adoServiMeDe As ADODB.Recordset
Attribute adoServiMeDe.VB_VarHelpID = -1
Dim VALIDAR As ValTecla
Dim WithEvents adoPrimaryRS As ADODB.Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim WithEvents CtrVisualizar As clsValidar
Attribute CtrVisualizar.VB_VarHelpID = -1
Dim WithEvents adoSENombres As ADODB.Recordset
Attribute adoSENombres.VB_VarHelpID = -1
Dim WithEvents adoHISTOCLI  As ADODB.Recordset
Attribute adoHISTOCLI.VB_VarHelpID = -1
Dim WithEvents adoSerTipoMov As ADODB.Recordset
Attribute adoSerTipoMov.VB_VarHelpID = -1
Public Sub imprimir()
        With DataGaleno.rsSERVICIOINTERNACION
           If .State = adStateClosed Then
                    .Open 'colocar este codigo a todos los informes
                    .Filter = ""
                    .Requery
                    .MoveFirst
                    .Filter = "FEC_PED" & " = '" & (TXTFIELDS(0).Text) & "' " & " And NROMOV" & " = " & (TXTFIELDS(1).Text)
                    'strSQL = Consulta.GetFilterString
           End If
           'If strSQL <> "" Then
           '         .Filter = strSQL
           'Else
           '         .Filter = ""
           'End If
           'strOrden = Consulta.GetOrder
           'If strOrden <> "" Then
           '         .Sort = strOrden
           'End If
           rpt_serviciointernacion.Refresh
           rpt_serviciointernacion.Show
           .Filter = ""
           If .State = adStateOpen Then 'colocar este codigo a todos los informes
              .Close
              strSQL = ""
           End If
        End With
End Sub

Private Sub BLOQUEAR_GRILLA(sino As Boolean)
Dim i As Integer
           If sino = True Then
                    For i = 0 To Me.grd_servicios.Columns.Count - 1
                        Me.grd_servicios.Columns(i).Locked = True
                        
                    Next
            Else
                    For i = 0 To Me.grd_servicios.Columns.Count - 1
                        If i < 4 Then
                            Me.grd_servicios.Columns(i).Locked = True
                        Else
                            Me.grd_servicios.Columns(i).Locked = False
                        End If
                    Next
            End If
End Sub
Private Sub AGREGAR_Click()
  On Error GoTo AddErr
  Cnn.BeginTrans
  Select Case Me.TAB_SERVICIOS.SelectedItem.Index
  Case 1
          With adoServiMe
              .AddNew
          End With
          Call BLOQUEAR_GRILLA(False)
          TXTFIELDS(0) = ""
          TXTFIELDS(5) = ""
          Me.TXTFIELDS(5).Locked = False
          SERVICIOS.grd_servicios.AllowUpdate = True
          SERVICIOS.grd_servicios.AllowDelete = True
          busfecha.SetFocus
        Exit Sub
  End Select
AddErr:
    MsgBox Err.Description
End Sub
Private Sub ACTUALIZAR_Click()
  On Error GoTo RefreshErr
  adoServiMe.Requery
  'Set adoServiMeDe = adoHISTOCLI("HISTORIACLINICADETALLE").UnderlyingValue
   Call TAB_SERVICIOS_Click
  Exit Sub
RefreshErr:
  MsgBox Err.Description
 
End Sub
Private Sub ANTERIOR_Click()
  On Error GoTo GoPrevError

  If Not adoServiMe.BOF Then adoServiMe.MovePrevious
  If adoServiMe.BOF And adoServiMe.RecordCount > 0 Then
    Beep
    adoServiMe.MoveFirst
  End If
  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub
Private Sub SIGUIENTE_Click()
  On Error GoTo GoNextError

  If Not adoServiMe.EOF Then adoServiMe.MoveNext
  If adoServiMe.EOF And adoServiMe.RecordCount > 0 Then
    Beep
    adoServiMe.MoveLast
  End If
  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub
Private Sub PRIMERO_Click()
  On Error GoTo GoFirstError

  adoServiMe.MoveFirst
  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub
Private Sub ULTIMO_Click()
  On Error GoTo GoLastError
  adoServiMe.MoveLast
  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub
Private Sub VOLVER_Click()
  Unload Me
End Sub

Private Sub GRABAR_Click()
'fncBotones (True)
 On Error GoTo UpdateErr
SendKeys "{TAB}", True
Select Case Me.TAB_SERVICIOS.SelectedItem.Index
Case 1
         adoServiMe.UpdateBatch
         
         If (grd_servicios.DataSource Is Nothing) Then
          
           Set grd_servicios.DataSource = adoServiMeDe
          
         End If
         
        adoServiMeDe.UpdateBatch
        Cnn.CommitTrans
        SERVICIOS.grd_servicios.AllowUpdate = False
        SERVICIOS.grd_servicios.AllowDelete = False
        Me.grd_servicios.Refresh
        Call BLOQUEAR_GRILLA(True)
         i = False
        Exit Sub
End Select
UpdateErr:
  MsgBox Err.Description
    
End Sub
Private Sub MODIFICAR_Click()
On Error GoTo EditErr
Cnn.BeginTrans
Select Case Me.TAB_SERVICIOS.SelectedItem.Index
Case 1
        Me.TXTFIELDS(5).Locked = False
        SERVICIOS.grd_servicios.AllowUpdate = True
        SERVICIOS.grd_servicios.AllowDelete = True
        Call BLOQUEAR_GRILLA(False)
End Select

Exit Sub
EditErr:
  MsgBox Err.Description
  
End Sub
Private Sub ELIMINAR_Click()
  Dim Respuesta As Integer
  On Error GoTo DeleteErr
  Respuesta = MsgBox("Esta Seguro que desea Eliminar el Registro", vbQuestion + vbYesNo, "Atencion")
  If Respuesta = 6 Then
    Select Case Me.TAB_SERVICIOS.SelectedItem.Index
    Case 1
        With adoServiMe
          .Delete
          .UpdateBatch adAffectAllChapters
          .MoveNext
          If .EOF Then .MoveLast
        End With
    End Select
  End If
Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub
Private Sub CANCELAR_Click()
Me.grd_servicios.AllowAddNew = False
Me.grd_servicios.AllowDelete = False
Me.grd_servicios.AllowUpdate = False
Call Cancel_Click
i = False
End Sub

Private Sub Cancel_Click()
  On Error Resume Next
  Cnn.RollbackTrans
  adoServiMe.CancelBatch
  adoServiMe_det.CancelBatch
  adoServiMe.Requery
   
  TAB_SERVICIOS_Click
  
 End Sub
Private Sub BUSCAR_Click()
  
 Dim frmAyuda As New frmConsult
 Set frmAyuda.ActiveConection = Cnn
 frmAyuda.Modal = True
 frmAyuda.Consulta = "SELECT * FROM SERVICIOINTERNACIONCABECERA"
 frmAyuda.Show vbModal
 If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
    
 End If
  If Not frmAyuda.Descargado Then
    With adoServiMe
           .MoveFirst
           .Find "fec_ped=" & frmAyuda.msGrid.Columns(1).Value
           .Find "nromov=" & frmAyuda.msGrid.Columns(2).Value
    End With
  End If
 frmAyuda.Cerrar
End Sub

Private Sub CmdAgregarIns_Click()
  On Error GoTo GoPrevError
  adoServiMe.UpdateBatch
  Dim frmAyuda As New frmConsult
  Set frmAyuda.ActiveConection = Cnn
  frmAyuda.Modal = True
  frmAyuda.Consulta = "SELECT COD_INSUMO AS CODSERVICIO, 'IN' AS TIPOSER, TIPO_INSUMO AS CODSER, NOM_INSUMO,ARANCEL,COSTO FROM INSUMOS"
  frmAyuda.Show vbModal
  If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
      With frmAyuda.msGrid
         Set adoServiMeDe = adoServiMe("SERVICIOINTERNACIONDETALLE").UnderlyingValue
         Call TAB_SERVICIOS_Click
         adoServiMeDe.Find "CODSERVICIO= " & .Columns(0).Text
        If adoServiMeDe.EOF Then
            adoServiMeDe.AddNew
            adoServiMeDe("CODSERVICIO").Value = .Columns(0).Value
            adoServiMeDe("TIPOSER").Value = .Columns(1).Value
            adoServiMeDe("CODSER").Value = .Columns(2).Value
            adoServiMeDe("NOMSERVICIO").Value = .Columns(3).Value
            'adoServiMeDe("COD_ENFERMERIA").Value = .Columns(4).Value
            adoServiMeDe("FECHA_MOVIMIENTO").Value = Date
            adoServiMeDe("ARANCEL").Value = .Columns(4).Value
            adoServiMeDe("COSTO").Value = .Columns(5).Value
            
            Set grd_servicios.DataSource = adoServiMeDe '("ChildCMD").UnderlyingValue
                'grd_servicios.Columns(6).Visible = False
                grd_servicios.Columns(10).Visible = False
                grd_servicios.Columns(11).Visible = False
                grd_servicios.Columns(12).Visible = False
                grd_servicios.Columns(13).Visible = False
                grd_servicios.Columns(14).Visible = False
                grd_servicios.Columns(15).Visible = False
                grd_servicios.Columns(16).Visible = False
             
        Else
            MsgBox ("El Servicio que ha elegido ya existe en el detalle por favor elija otro")
       End If
    End With
  End If
 frmAyuda.Cerrar
 Exit Sub
GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub cmdagregarmed_Click()
  On Error GoTo GoPrevError
  adoServiMe.UpdateBatch
  Dim frmAyuda As New frmConsult
  Set frmAyuda.ActiveConection = Cnn
  frmAyuda.Modal = True
  frmAyuda.Consulta = "SELECT ID_MEDICAMENTO AS CODSERVICIO, 'ME' AS TIPOSER, COD_MEDICAMENTO AS CODSER, NOM_MEDICAMENTO, ARANCEL,COSTO FROM MEDICAMENTOS"
  ',COD_ENFERMERIA,ARANCEL,COSTO  FROM SERVICIO"
  frmAyuda.Show vbModal
  If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
      With frmAyuda.msGrid
         Set adoServiMeDe = adoServiMe("SERVICIOINTERNACIONDETALLE").UnderlyingValue
         Call TAB_SERVICIOS_Click
         adoServiMeDe.Find "CODSERVICIO= " & .Columns(0).Text
        If adoServiMeDe.EOF Then
            adoServiMeDe.AddNew
            adoServiMeDe("CODSERVICIO").Value = .Columns(0).Value
            adoServiMeDe("TIPOSER").Value = .Columns(1).Value
            adoServiMeDe("CODSER").Value = .Columns(2).Value
            adoServiMeDe("NOMSERVICIO").Value = .Columns(3).Value
            'adoServiMeDe("COD_ENFERMERIA").Value = .Columns(4).Value
            adoServiMeDe("FECHA_MOVIMIENTO").Value = Date
            adoServiMeDe("ARANCEL").Value = .Columns(4).Value
            adoServiMeDe("COSTO").Value = .Columns(5).Value
            
            Set grd_servicios.DataSource = adoServiMeDe '("ChildCMD").UnderlyingValue
                'grd_servicios.Columns(6).Visible = False
                grd_servicios.Columns(10).Visible = False
                grd_servicios.Columns(11).Visible = False
                grd_servicios.Columns(12).Visible = False
                grd_servicios.Columns(13).Visible = False
                grd_servicios.Columns(14).Visible = False
                grd_servicios.Columns(15).Visible = False
                grd_servicios.Columns(16).Visible = False
             
        Else
            MsgBox ("El Servicio que ha elegido ya existe en el detalle por favor elija otro")
       End If
    End With
  End If
 frmAyuda.Cerrar
 Exit Sub
GoPrevError:
  MsgBox Err.Description
'Call CANCELAR_Click
  

End Sub

Private Sub cmdagregarserord_Click()
  On Error GoTo GoPrevError
  adoServiMe.UpdateBatch
  Dim frmAyuda As New frmConsult
  Set frmAyuda.ActiveConection = Cnn
  frmAyuda.Modal = True
  frmAyuda.Consulta = "SELECT CODSERVICIO, TIPOSER, CODSER, NOMSERVICIO, PRE_VENTA, PRE_COSTO FROM ORDENSERVICIODETALLE  WHERE ORDENSERVICIODETALLE.FICHACLINICA = " & TXTFIELDS(2).Text & " AND  ORDENSERVICIODETALLE.FEC_PED = '" & TXTFIELDS(0).Text & "' AND ORDENSERVICIODETALLE.NROMOV = " & TXTFIELDS(1).Text
  frmAyuda.Show vbModal
  If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
      With frmAyuda.msGrid
         Set adoServiMeDe = adoServiMe("SERVICIOINTERNACIONDETALLE").UnderlyingValue
         Call TAB_SERVICIOS_Click
         adoServiMeDe.Find "CODSERVICIO= " & .Columns(0).Text
        If adoServiMeDe.EOF Then
            adoServiMeDe.AddNew
            adoServiMeDe("CODSERVICIO").Value = .Columns(0).Value
            adoServiMeDe("TIPOSER").Value = .Columns(1).Value
            adoServiMeDe("CODSER").Value = .Columns(2).Value
            adoServiMeDe("NOMSERVICIO").Value = .Columns(3).Value
            adoServiMeDe("FECHA_MOVIMIENTO").Value = Date
            'adoServiMeDe("COD_ENFERMERIA").Value = .Columns(4).Value
            adoServiMeDe("ARANCEL").Value = .Columns(4).Value
            adoServiMeDe("COSTO").Value = .Columns(5).Value
            
            Set grd_servicios.DataSource = adoServiMeDe '("ChildCMD").UnderlyingValue
                'grd_servicios.Columns(6).Visible = False
                grd_servicios.Columns(10).Visible = False
                grd_servicios.Columns(11).Visible = False
                grd_servicios.Columns(12).Visible = False
                grd_servicios.Columns(13).Visible = False
                grd_servicios.Columns(14).Visible = False
                grd_servicios.Columns(15).Visible = False
                grd_servicios.Columns(16).Visible = False
                        
        Else
            MsgBox ("El Servicio que ha elegido ya existe en el detalle por favor elija otro")
       End If
    End With
  End If
 frmAyuda.Cerrar
 Exit Sub
GoPrevError:
  MsgBox Err.Description
'Call CANCELAR_Click
End Sub

Private Sub CmdAgregarServicios_Click()
  On Error GoTo GoPrevError
  adoServiMe.UpdateBatch
  Dim frmAyuda As New frmConsult
  Set frmAyuda.ActiveConection = Cnn
  frmAyuda.Modal = True
  frmAyuda.Consulta = "SELECT SERVICIO.CODSERVICIO,SERVICIO.TIPOSER,SERVICIO.CODSER,SERVICIO.NOMSERVICIO,ARANCELES.ARANCEL,ARANCELES.COSTO FROM SERVICIO,ARANCELES WHERE SERVICIO.CODSERVICIO=ARANCELES.CODSERVICIO AND  SERVICIO.TIPOSER=ARANCELES.TIPOSER"
  frmAyuda.Show vbModal
  If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
      With frmAyuda.msGrid
         Set adoServiMeDe = adoServiMe("SERVICIOINTERNACIONDETALLE").UnderlyingValue
         Call TAB_SERVICIOS_Click
         adoServiMeDe.Find "CODSERVICIO= " & .Columns(0).Text
        If adoServiMeDe.EOF Then
            adoServiMeDe.AddNew
            adoServiMeDe("CODSERVICIO").Value = .Columns(0).Value
            adoServiMeDe("TIPOSER").Value = .Columns(1).Value
            adoServiMeDe("CODSER").Value = .Columns(2).Value
            adoServiMeDe("NOMSERVICIO").Value = .Columns(3).Value
            adoServiMeDe("FECHA_MOVIMIENTO").Value = Date
            'adoServiMeDe("COD_ENFERMERIA").Value = .Columns(4).Value
            adoServiMeDe("ARANCEL").Value = .Columns(4).Value
            adoServiMeDe("COSTO").Value = .Columns(5).Value
            
            Set grd_servicios.DataSource = adoServiMeDe '("ChildCMD").UnderlyingValue
                'grd_servicios.Columns(6).Visible = False
                grd_servicios.Columns(10).Visible = False
                grd_servicios.Columns(11).Visible = False
                grd_servicios.Columns(12).Visible = False
                grd_servicios.Columns(13).Visible = False
                grd_servicios.Columns(14).Visible = False
                grd_servicios.Columns(15).Visible = False
                grd_servicios.Columns(16).Visible = False
                        
        Else
            MsgBox ("El Servicio que ha elegido ya existe en el detalle por favor elija otro")
       End If
    End With
  End If
 frmAyuda.Cerrar
 Exit Sub
GoPrevError:
  MsgBox Err.Description
'Call CANCELAR_Click
  
End Sub

Private Sub CmdAltaPaciente_Click()
      
End Sub

Private Sub CmdBuscarCama_Click()
    Dim frmAyuda As New frmConsult
    Set frmAyuda.ActiveConection = Cnn
    frmAyuda.Modal = True
    frmAyuda.Consulta = "SELECT COD_CAMA From HABITACIONES  WHERE  COD_TIPO = '" & Trim(TXTFIELDS(9).Text) & "' AND COD_HABITACION= '" & Trim(TXTFIELDS(7).Text) & "'" & " AND ESTADO='LIBRE'"
   'frmAyuda.Consulta = "SELECT COD_CAMA From HABITACIONES  WHERE  COD_TIPO = '" & Trim(TXTFIELDS(9).Text) & "'" AND COD_HABITACION = '" & Trim(TXTFIELDS(7).Text) & "'"

    frmAyuda.Show vbModal
    If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
      With frmAyuda.msGrid
          TXTFIELDS(8).Text = .Columns(0).Text
        ' TXTFIELDS(8).Text = .Columns(1).Text
          adoServiMe.Fields("COD_CAMA").Value = .Columns(0).Value
        ' adoServiMe.Fields("COD_CAMA").Value = .Columns(2).Value
      
    End With
 End If
 frmAyuda.Cerrar
 cmdagregarservicios.SetFocus
End Sub

Private Sub CmdbusHabitacion_Click()

    Dim frmAyuda As New frmConsult
    Set frmAyuda.ActiveConection = Cnn
    frmAyuda.Modal = True
   ' frmAyuda.Consulta = "SELECT COD_HABITACION From HABITACIONES  WHERE  COD_TIPO = Trim(TXTFIELDS(9).Text)  group by COD_HABITACION"
    frmAyuda.Consulta = "SELECT COD_HABITACION From HABITACIONES  WHERE  COD_TIPO = '" & Trim(TXTFIELDS(9).Text) & "'" & " AND ESTADO='LIBRE'" & " group by COD_HABITACION"
    'frmAyuda.Consulta = "SELECT COD_HABITACION FROM HABITACIONES WHERE  (TIPO = 'HABITACION')"
    frmAyuda.Show vbModal
    If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
      With frmAyuda.msGrid
         TXTFIELDS(7).Text = .Columns(0).Text
        ' TXTFIELDS(8).Text = .Columns(1).Text
         adoServiMe.Fields("COD_HABITACION").Value = .Columns(0).Value
        ' adoServiMe.Fields("COD_CAMA").Value = .Columns(2).Value
      
    End With
 End If
 frmAyuda.Cerrar
 CmdBuscarCama.SetFocus
End Sub

Private Sub cmdBusTipo_Click()
    Dim frmAyuda As New frmConsult
    Set frmAyuda.ActiveConection = Cnn
    frmAyuda.Modal = True
    'frmAyuda.Consulta = " SELECT CODIGO From Parametros  WHERE  (TIPO = 'HABITACION')"
    frmAyuda.Consulta = "SELECT COD_TIPO FROM HABITACIONES  group by COD_TIPO"
    frmAyuda.Show vbModal
    If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
      With frmAyuda.msGrid
           TXTFIELDS(7).Text = ""
           TXTFIELDS(8).Text = ""
           TXTFIELDS(9).Text = .Columns(0).Text
         ' TXTFIELDS(8).Text = .Columns(1).Text
           adoServiMe.Fields("COD_TIPO").Value = .Columns(0).Value
         ' adoServiMe.Fields("COD_CAMA").Value = .Columns(2).Value
           TXTFIELDS(7).Text = ""
           TXTFIELDS(8).Text = ""
      End With
   End If
   frmAyuda.Cerrar
   CmdbusHabitacion.SetFocus
End Sub



Private Sub busfecha_Click()
' On Error Resume Next
 Dim frmAyuda As New frmConsult
 Set frmAyuda.ActiveConection = Cnn
 frmAyuda.Modal = True
 frmAyuda.Consulta = "SELECT DISTINCT PERSONAS.NOMBRES,PERSONAS.APELLIDOS,ORDENSERVICIOCABECERA.FEC_PED," _
 & " ORDENSERVICIOCABECERA.NROMOV,ORDENSERVICIOCABECERA.FICHACLINICA,ORDENSERVICIOCABECERA.TIPDOC," _
 & " ORDENSERVICIOCABECERA.NRODOC,ORDENSERVICIOCABECERA.TIPOMOV,ORDENSERVICIOCABECERA.CODDR,ORDENSERVICIOCABECERA.CODINS FROM ORDENSERVICIOCABECERA,PERSONAS,ORDENSERVICIODETALLE " _
 & " WHERE ORDENSERVICIOCABECERA.TIPDOC=PERSONAS.TIPDOC AND ORDENSERVICIOCABECERA.NRODOC=PERSONAS.NRODOC AND ORDENSERVICIOCABECERA.fec_ped = ORDENSERVICIODETALLE.fec_ped and ORDENSERVICIOCABECERA.NROMOV =ORDENSERVICIODETALLE.nromov" _
 & " and (ORDENSERVICIODETALLE.TIPOSER='QUI' OR ORDENSERVICIODETALLE.TIPOSER='SAN') AND NOT EXISTS (SELECT fec_ped,nromov  From SERVICIOINTERNACIONCABECERA where " _
 & " SERVICIOINTERNACIONCABECERA.fec_ped = ORDENSERVICIOCABECERA.fec_ped and SERVICIOINTERNACIONCABECERA.NROMOV =ORDENSERVICIOCABECERA.nromov)"
 frmAyuda.Show vbModal
 If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
    With frmAyuda.msGrid
        adoServiMe.Fields("FEC_PED").Value = .Columns(2).Text
        adoServiMe.Fields("NROMOV").Value = .Columns(3).Text
        adoServiMe.Fields("FICHACLINICA").Value = .Columns(4).Text
        adoServiMe.Fields("TIPDOC").Value = .Columns(5).Text
        adoServiMe.Fields("NRODOC").Value = .Columns(6).Text
        adoServiMe.Fields("TIPOMOV").Value = .Columns(7).Text
        
        TXTFIELDS(0).Text = .Columns(2).Text
        TXTFIELDS(1).Text = .Columns(3).Text
        TXTFIELDS(2).Text = .Columns(4).Text
        TXTFIELDS(3).Text = .Columns(5).Text
        TXTFIELDS(4).Text = .Columns(6).Text
        TXTFIELDS(6).Text = .Columns(7).Text
            
        LBLServicio(10).Caption = .Columns(1).Text & ", " & .Columns(0).Text
        
        
       'If Not adoSerTipoMovState = 0 Then
             adoSerTipoMov.MoveFirst
            If Not TXTFIELDS(6).Text = "" Then
                adoSerTipoMov.Filter = "TIPOMOV=" & TXTFIELDS(6).Text
                LBLServiTIP(5).Caption = adoSerTipoMov.Fields("NOMTIPOMOV").Value
            Else
                LBLServiTIP(5).Caption = ""
                
            End If
         
       '  End If
        
        
    End With
 End If
 TXTFIELDS(5).SetFocus
frmAyuda.Cerrar
End Sub


Private Sub CmdFichaClinica_Click()
If MDIGALENO.tbToolBar.Buttons("Next").Enabled = True Then
    ATENCIONMEDICA.Show
    Call ATENCIONMEDICA.PASARIDENTIDAD(adoServiMe.Fields("FEC_PED").Value, adoServiMe.Fields("NROMOV").Value, adoServiMe.Fields("FICHACLINICA").Value, adoServiMe.Fields("TIPDOC").Value, adoServiMe.Fields("NRODOC").Value)
End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    Set CtrVisualizar = New clsValidar
    Set VALIDAR = New ValTecla
    Me.Height = 7500
    Me.Height = 7500
    Me.Left = 80
    Me.Top = 0
    Me.Width = 11800
    Set adoServiMe = New ADODB.Recordset
    Set adoServiMeCa = New ADODB.Recordset
    Set adoServiMeDe = New ADODB.Recordset
    Set adoSENombres = New ADODB.Recordset
    Set adoHISTOCLI = New ADODB.Recordset
    Set adoSerTipoMov = New ADODB.Recordset

    adoSENombres.Open "Select FICHACLINICA.FICHACLINICA,FICHACLINICA.TIPDOC,FICHACLINICA.NRODOC,NOMBRES,APELLIDOS FROM FICHACLINICA,PERSONAS WHERE FICHACLINICA.TIPDOC=PERSONAS.TIPDOC AND FICHACLINICA.NRODOC=PERSONAS.NRODOC  ", Cnn

    adoServiMe.Open "SHAPE {SELECT * FROM SERVICIOINTERNACIONCABECERA}  AS SERVICIOINTERNACIONCABECERA APPEND ({SELECT CODSERVICIO,TIPOSER,CODSER,NOMSERVICIO,CANTIDAD,ARANCEL,COSTO,CODDR,HONORARIOS,OBSERVACIONES,FECHA_MOVIMIENTO,COD_USUARIO,FICHACLINICA,FEC_PED,NROMOV," _
    & " TIPDOC,NRODOC FROM SERVICIOINTERNACIONDETALLE}  AS SERVICIOINTERNACIONDETALLE RELATE 'FICHACLINICA' TO 'FICHACLINICA','FEC_PED' TO 'FEC_PED','NROMOV' TO 'NROMOV','TIPDOC' TO 'TIPDOC','NRODOC' TO 'NRODOC') AS SERVICIOINTERNACIONDETALLE", Cnn, adOpenStatic, adLockBatchOptimistic
    
    If adoServiMeDe.State = 0 Then
        If adoServiMe.EOF = False And adoServiMe.BOF = False Then
            Set adoServiMeDe = adoServiMe("SERVICIOINTERNACIONDETALLE").UnderlyingValue
        End If
    End If
    Me.frmServiMe.Visible = CtrVisualizar.Visualizar(15, Cnn, IdUsuario)
   
    adoSerTipoMov.Open "Select * From TIPOMOV", Cnn
   
    Call TAB_SERVICIOS_Click

End Sub



Private Sub grd_servicios_DblClick()
On Error GoTo GoPrevError
Select Case grd_servicios.Col
Case 13
 Dim frmAyuda As New frmConsult
 Set frmAyuda.ActiveConection = Cnn
 frmAyuda.Modal = True
 frmAyuda.Consulta = "SELECT DOCTOR.CODDR,PERSONAS.APELLIDOS,PERSONAS.NOMBRES, PERSONAS.TIPDOC,PERSONAS.NRODOC,DOCTOR.REGPRF FROM DOCTOR  LEFT OUTER JOIN PERSONAS ON DOCTOR.TIPDOC = PERSONAS.TIPDOC AND DOCTOR.NRODOC = PERSONAS.NRODOC"
 frmAyuda.Show vbModal
 If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
    With frmAyuda.msGrid
         adoServiMeDe("CODDR").Value = .Columns(0).Value
    End With
 End If
End Select
frmAyuda.Cerrar
Exit Sub
GoPrevError:
    MsgBox Err.Description
End Sub

Private Sub grd_servicios_Error(ByVal DataError As Integer, Response As Integer)
 Response = 0
End Sub


Private Sub TAB_SERVICIOS_Click()
 'On Error Resume Next
 Select Case Me.TAB_SERVICIOS.SelectedItem.Index
 Case 1
    If MDIGALENO.tbToolBar.Buttons("MODIFICAR").Enabled = True And MDIGALENO.tbToolBar.Buttons("Add").Enabled = True And MDIGALENO.tbToolBar.Buttons("Delete").Enabled = True Then
          Dim oText As TextBox
          'Enlaza los cuadros de texto con el proveedor de datos
           For Each oText In Me.TXTFIELDS
             If oText.Index < 40 Then
                   Set oText.DataSource = adoServiMe
             End If
           Next
    End If
      If adoServiMe.EOF = False And adoServiMe.BOF = False Then
          Set adoServiMeDe = adoServiMe("SERVICIOINTERNACIONDETALLE").UnderlyingValue
      Else
          Set adoServiMeDe = Nothing
      End If
      Set grd_servicios.DataSource = adoServiMeDe '("ChildCMD").UnderlyingValue
      If MDIGALENO.tbToolBar.Buttons("MODIFICAR").Enabled = True And MDIGALENO.tbToolBar.Buttons("Add").Enabled = True And MDIGALENO.tbToolBar.Buttons("Delete").Enabled = True Then
        Call BLOQUEAR_GRILLA(True)
      Else
        Call BLOQUEAR_GRILLA(False)
      End If
      If Not (adoServiMe.EOF And adoServiMe.BOF) Then
       'grd_servicios.Columns(6).Visible = False
       grd_servicios.Columns(10).Visible = False
       grd_servicios.Columns(11).Visible = False
       grd_servicios.Columns(12).Visible = False
       grd_servicios.Columns(13).Visible = False
       grd_servicios.Columns(14).Visible = False
       grd_servicios.Columns(15).Visible = False
       grd_servicios.Columns(16).Visible = False

       
    End If

End Select
End Sub

Private Sub CtrVisualizar_ABMtollbar(intABM As Integer)
 Select Case intABM
    Case 0
        MDIGALENO.tbToolBar.Buttons("MODIFICAR").Enabled = False
        MDIGALENO.tbToolBar.Buttons("Add").Enabled = False
        MDIGALENO.tbToolBar.Buttons("Delete").Enabled = False
    
    Case 1 'modificar 001,010,011,100,101,110,111
        MDIGALENO.tbToolBar.Buttons("Add").Enabled = False
        MDIGALENO.tbToolBar.Buttons("Delete").Enabled = False
        MDIGALENO.tbToolBar.Buttons("MODIFICAR").Enabled = True
    
    Case 10 'baja
        MDIGALENO.tbToolBar.Buttons("Add").Enabled = False
        MDIGALENO.tbToolBar.Buttons("MODIFICAR").Enabled = False
        MDIGALENO.tbToolBar.Buttons("Delete").Enabled = True
    Case 11 'baja y modificacion
        MDIGALENO.tbToolBar.Buttons("Add").Enabled = False
        MDIGALENO.tbToolBar.Buttons("MODIFICAR").Enabled = True
        MDIGALENO.tbToolBar.Buttons("Delete").Enabled = True
      
    Case 100 'alta
        MDIGALENO.tbToolBar.Buttons("Delete").Enabled = False
        MDIGALENO.tbToolBar.Buttons("MODIFICAR").Enabled = False
        MDIGALENO.tbToolBar.Buttons("Add").Enabled = True

    Case 101 'alta y modificacion
        MDIGALENO.tbToolBar.Buttons("Delete").Enabled = False
        MDIGALENO.tbToolBar.Buttons("MODIFICAR").Enabled = True
        MDIGALENO.tbToolBar.Buttons("Add").Enabled = True
    
    Case 110 'alta y baja
        MDIGALENO.tbToolBar.Buttons("MODIFICAR").Enabled = False
        MDIGALENO.tbToolBar.Buttons("Add").Enabled = True
        MDIGALENO.tbToolBar.Buttons("Delete").Enabled = True
    
    Case 111 'abm
        MDIGALENO.tbToolBar.Buttons("MODIFICAR").Enabled = True
        MDIGALENO.tbToolBar.Buttons("Add").Enabled = True
        MDIGALENO.tbToolBar.Buttons("Delete").Enabled = True
    
End Select

End Sub
Private Sub Form_GotFocus()
  CtrVisualizar_ABMtollbar CtrVisualizar.PERMISOABM
End Sub

Private Sub Form_Unload(Cancel As Integer)
 On Error Resume Next
 VALIDAR.habilitarToolbar MDIGALENO.tbToolBar
 Screen.MousePointer = vbDefault
 Cnn.RollbackTrans
 adoPrimaryRS.Close
 adoPrimaryRS.Close
 Set VALIDAR = Nothing
End Sub

Public Sub UltimoReg()
 ULTIMO_Click
End Sub

Public Sub PrimeroReg()
  PRIMERO_Click
End Sub

Public Sub SiguienteReg()
 SIGUIENTE_Click
End Sub

Public Sub AnteriorReg()
ANTERIOR_Click
End Sub

Public Sub eliminarReg()
ELIMINAR_Click
End Sub

Public Sub AgregarReg()
AGREGAR_Click
End Sub
Public Sub ModificarReg()
MODIFICAR_Click
End Sub
Public Sub ActualizarReg()
GRABAR_Click
CtrVisualizar_ABMtollbar CtrVisualizar.PERMISOABM
End Sub
Public Sub requeryReg()
ACTUALIZAR_Click
End Sub
Public Sub buscarReg()
BUSCAR_Click
End Sub
Public Sub CancelarReg()

  On Error Resume Next
  adoServiMe.MoveFirst
  TAB_SERVICIOS_Click
  CANCELAR_Click
  CtrVisualizar_ABMtollbar CtrVisualizar.PERMISOABM

End Sub

Private Sub txtFields_Change(Index As Integer)
              
      If Not adoSENombres.State = 0 Then
            adoSENombres.MoveFirst
            If Not TXTFIELDS(3).Text = "" Then
                adoSENombres.Find "TIPDOC='" & Trim$(TXTFIELDS(3).Text) & "'"
            End If
        End If
        
        If Not adoSENombres.State = 0 Then
            adoSENombres.MoveFirst
            If Not TXTFIELDS(4).Text = "" Then
                adoSENombres.Find "NRODOC=" & Trim$(TXTFIELDS(4).Text)
                LBLServicio(10).Caption = adoSENombres.Fields("APELLIDOS").Value + ", " + adoSENombres.Fields("NOMBRES").Value
            Else
                LBLServicio(10).Caption = ""
            End If
        End If
                    
      'If Not adoSerTipoMovState = 0 Then
             adoSerTipoMov.MoveFirst
            If Not TXTFIELDS(6).Text = "" Then
                adoSerTipoMov.Filter = "TIPOMOV=" & TXTFIELDS(6).Text
                LBLServiTIP(5).Caption = adoSerTipoMov.Fields("NOMTIPOMOV").Value
            Else
                LBLServiTIP(5).Caption = ""
                
            End If
         
       '  End If
      
      
          
      '  If Not TXTFIELDS(5).Text = "" Then
      '     adoServiMe.Fields("FEC_INTERNACION").Value = TXTFIELDS(5).Text
           
      '  End If
         
      
End Sub

Private Sub TXTFIELDS_KeyPress(Index As Integer, KeyAscii As Integer)
 
Dim X As Boolean
      Select Case Index
            'Case 1, 58, 59, 80, 20, 21, 25, 26, 33, 35, 42 To 47, 58 'numerico
             '   VALIDAR.ValidarAscii KeyAscii
             Case 5  'fecha
                VALIDAR.ValidarAscii KeyAscii, False
                
             
                
      End Select

End Sub

Private Sub TXTFIELDS_Validate(Index As Integer, Cancel As Boolean)
'  SendKeys "{TAB}"
  Select Case Index
    Case 5 'fecha
        blCancel = Not VALIDAR.esFecha(TXTFIELDS(Index).Text)
        If blCancel = True Then
            MsgBox "DISCULPE FECHA INVALIDA - FORMATO  '99/99/9999'", vbExclamation
            TXTFIELDS(5) = ""
        Else
            If CDate(TXTFIELDS(5)) < CDate(TXTFIELDS(0)) Then
               MsgBox "DISCULPE LA FECHA DE INTERNACION, NO PUEDE SER MENOR QUE LA FECHA DE ORDEN ", vbExclamation
               blCancel = True
               TXTFIELDS(5).Text = ""
               TXTFIELDS(5).SetFocus
            End If
        End If
        Cancel = blCancel
        
        
    End Select
End Sub
