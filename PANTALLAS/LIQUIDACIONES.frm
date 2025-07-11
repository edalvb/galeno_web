VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form LIQUIDACIONES 
   Caption         =   "LIQUIDACIONES MEDICAS"
   ClientHeight    =   6840
   ClientLeft      =   1335
   ClientTop       =   2340
   ClientWidth     =   11745
   Icon            =   "LIQUIDACIONES.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6840
   ScaleWidth      =   11745
   Begin VB.Frame frmliquida 
      BackColor       =   &H80000013&
      ForeColor       =   &H80000002&
      Height          =   6075
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   11265
      Begin MSDataGridLib.DataGrid grd_liquidaciones 
         Height          =   2175
         Left            =   120
         Negotiate       =   -1  'True
         TabIndex        =   19
         Top             =   1920
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   3836
         _Version        =   393216
         AllowUpdate     =   0   'False
         ForeColor       =   -2147483635
         HeadLines       =   1
         RowHeight       =   19
         RowDividerStyle =   1
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
         Caption         =   "DETALLE DE LIQUIDACIONES"
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
               LCID            =   3082
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
               LCID            =   3082
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
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "FICHACLINICA"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
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
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TXTFIELDS 
         Alignment       =   1  'Right Justify
         DataField       =   "GRAVADA"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   14
         Left            =   8640
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   30
         Top             =   4560
         Width           =   2025
      End
      Begin VB.TextBox TXTFIELDS 
         Alignment       =   1  'Right Justify
         DataField       =   "IVA"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   13
         Left            =   8640
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   29
         Top             =   5280
         Width           =   2025
      End
      Begin VB.TextBox TXTFIELDS 
         Alignment       =   1  'Right Justify
         DataField       =   "TOTAL"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   12
         Left            =   8640
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   28
         Top             =   5640
         Width           =   2025
      End
      Begin VB.TextBox TXTFIELDS 
         Alignment       =   1  'Right Justify
         DataField       =   "DESCUENTO"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   8640
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   27
         Top             =   4920
         Width           =   2025
      End
      Begin VB.CommandButton busliqser 
         DragIcon        =   "LIQUIDACIONES.frx":0ABA
         Height          =   375
         Left            =   120
         Picture         =   "LIQUIDACIONES.frx":0EFC
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox TXTFIELDS 
         Alignment       =   1  'Right Justify
         DataField       =   "EXENTA"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   6480
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   20
         Top             =   4560
         Width           =   2145
      End
      Begin VB.CommandButton cmdBusTipo 
         DragIcon        =   "LIQUIDACIONES.frx":133E
         Height          =   375
         Left            =   120
         Picture         =   "LIQUIDACIONES.frx":1780
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "TIPO_FAC"
         Height          =   285
         Index           =   6
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   960
         Width           =   945
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "FEC_MOV"
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
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "NRODOC"
         Height          =   285
         Index           =   4
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   600
         Width           =   1245
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "TIPODOC"
         Height          =   285
         Index           =   3
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "NRO_MOV"
         ForeColor       =   &H000000C0&
         Height          =   285
         Index           =   2
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton busFICHA 
         DragIcon        =   "LIQUIDACIONES.frx":1BC2
         Height          =   375
         Left            =   120
         Picture         =   "LIQUIDACIONES.frx":2004
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "NRO_FAC"
         Height          =   285
         Index           =   9
         Left            =   7080
         MaxLength       =   10
         TabIndex        =   2
         Top             =   960
         Width           =   945
      End
      Begin VB.Label LBLliquida 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "GRAVADAS"
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
         Left            =   9240
         TabIndex        =   32
         Top             =   4200
         Width           =   960
      End
      Begin VB.Label LBLliquida 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "EXENTAS"
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
         Left            =   7200
         TabIndex        =   31
         Top             =   4200
         Width           =   720
      End
      Begin VB.Label LBLliquida 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Agregar Servicios"
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
         Left            =   840
         TabIndex        =   26
         Top             =   1560
         Width           =   1380
      End
      Begin VB.Label LBLliquida 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DESCUENTO"
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
         Left            =   5160
         TabIndex        =   24
         Top             =   4920
         Width           =   1020
      End
      Begin VB.Label LBLliquida 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL A PAGAR"
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
         Left            =   5160
         TabIndex        =   23
         Top             =   5640
         Width           =   1365
      End
      Begin VB.Label LBLliquida 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "IVA "
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
         Left            =   5160
         TabIndex        =   22
         Top             =   5280
         Width           =   345
      End
      Begin VB.Label LBLliquida 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SUB-TOTALES"
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
         Left            =   5160
         TabIndex        =   21
         Top             =   4560
         Width           =   1155
      End
      Begin VB.Label LBLliquida 
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
         Index           =   11
         Left            =   3480
         TabIndex        =   18
         Top             =   960
         Width           =   1365
      End
      Begin VB.Label LBLliquida 
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
         TabIndex        =   17
         Top             =   600
         Width           =   1740
      End
      Begin VB.Label LBLliquida 
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
         Index           =   0
         Left            =   720
         TabIndex        =   16
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label LBLliquida 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha y Numero de Movimiento"
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
         Left            =   3840
         TabIndex        =   15
         Top             =   240
         Width           =   2460
      End
      Begin VB.Label LBLliquida 
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
         TabIndex        =   14
         Top             =   600
         Width           =   1965
      End
      Begin VB.Label LBLliquida 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Número de Factura"
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
         Left            =   5280
         TabIndex        =   12
         Top             =   960
         Width           =   1485
      End
      Begin VB.Label LBLliquida 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Factura"
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
         Left            =   840
         TabIndex        =   11
         Top             =   960
         Width           =   1185
      End
      Begin VB.Label LBLliquida 
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
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   1950
      End
   End
   Begin MSComctlLib.TabStrip TAB_LIQUIDA 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   11668
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      MultiSelect     =   -1  'True
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "PACIENTES"
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
End
Attribute VB_Name = "LIQUIDACIONES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents adoliquida     As ADODB.Recordset
Attribute adoliquida.VB_VarHelpID = -1
Dim WithEvents adoliquida_det As ADODB.Recordset
Attribute adoliquida_det.VB_VarHelpID = -1
Dim WithEvents adoinstituto     As ADODB.Recordset
Attribute adoinstituto.VB_VarHelpID = -1
Dim WithEvents adoinstituto_det As ADODB.Recordset
Attribute adoinstituto_det.VB_VarHelpID = -1
Dim WithEvents adoNombres   As ADODB.Recordset
Attribute adoNombres.VB_VarHelpID = -1
Dim WithEvents adoFactura  As ADODB.Recordset
Attribute adoFactura.VB_VarHelpID = -1
Dim WithEvents CtrVisualizar As clsValidar
Attribute CtrVisualizar.VB_VarHelpID = -1
Dim WithEvents adoPrimaryRS As ADODB.Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim VALIDAR As ValTecla
Public Sub imprimir()
        With DataGaleno.rsfactura_liq_pac_Grouping
           If .State = adStateClosed Then
                    .Open 'colocar este codigo a todos los informes
                    .Filter = ""
                    .Requery
                    .MoveFirst
                    .Filter = "FEC_MOV" & " = " & (TXTFIELDS(1).Text) & " " & " And NRO_MOV" & " = " & (TXTFIELDS(2).Text)
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
           rpt_factura.Refresh
           rpt_factura.Show
           .Filter = ""
           If .State = adStateOpen Then 'colocar este codigo a todos los informes
              .Close
              strSQL = ""
           End If
        End With
End Sub

Private Sub busficha_Click()
 Dim frmAyuda As New frmConsult
 Set frmAyuda.ActiveConection = Cnn
 frmAyuda.Modal = True
 frmAyuda.Consulta = "SELECT FICHACLINICA.FICHACLINICA ,PERSONAS.TIPDOC,PERSONAS.NRODOC,PERSONAS.NOMBRES,PERSONAS.APELLIDOS FROM FICHACLINICA,PERSONAS WHERE FICHACLINICA.TIPDOC=PERSONAS.TIPDOC AND FICHACLINICA.NRODOC=PERSONAS.NRODOC"
 frmAyuda.Show vbModal
 If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
    With frmAyuda.msGrid
        adoliquida.Fields("FICHACLINICA").Value = .Columns(0).Value
        adoliquida.Fields("TIPODOC").Value = .Columns(1).Value
        adoliquida.Fields("NRODOC").Value = .Columns(2).Value
        LBLliquida(10).Caption = .Columns(4).Text & ", " & .Columns(3).Text
    End With
 End If
frmAyuda.Cerrar
End Sub

Private Sub busliqser_Click()
 On Error GoTo GoPrevError
 If adoliquida.Status <> 0 Then
   adoliquida.UpdateBatch
 End If
 Dim frmAyuda As New frmConsult
 Set frmAyuda.ActiveConection = Cnn
 frmAyuda.Modal = True
 frmAyuda.Consulta = "SELECT FEC_PED,NROMOV,CODSERVICIO, TIPOSER, CODSER, NOMSERVICIO, CANTIDAD, ARANCEL, COSTO" _
                     & " From SERVICIOINTERNACIONDETALLE where SERVICIOINTERNACIONDETALLE.fichaclinica= " & TXTFIELDS(0).Text _
                     & " AND NOT EXISTS (SELECT * FROM COBROS_DETALLE WHERE FECSER=SERVICIOINTERNACIONDETALLE.FEC_PED AND NROSER=SERVICIOINTERNACIONDETALLE.NROMOV AND COBROS_DETALLE.CODSERVICIO=SERVICIOINTERNACIONDETALLE.CODSERVICIO)" _
                     & " Union SELECT FEC_PED,NROMOV,CODSERVICIO, TIPOSER, CODSER," _
                     & " NOMSERVICIO, CANTIDAD, PRE_COSTO AS arancel, PRE_VENTA AS costo From ORDENSERVICIODETALLE where ORDENSERVICIODETALLE.fichaclinica= " & TXTFIELDS(0).Text _
                     & " AND NOT EXISTS (SELECT * FROM COBROS_DETALLE WHERE FECSER=ORDENSERVICIODETALLE.FEC_PED AND NROSER=ORDENSERVICIODETALLE.NROMOV AND COBROS_DETALLE.CODSERVICIO=ORDENSERVICIODETALLE.CODSERVICIO)"
 frmAyuda.Show vbModal
 If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
    With frmAyuda.msGrid
        Set adoliquida_det = adoliquida("cobros_detalle").UnderlyingValue
        'Call TAB_LIQUIDA_Click
        adoliquida_det.Find "CODSERVICIO= " & .Columns(2).Text
        If adoliquida_det.EOF Then
            adoliquida_det.AddNew
            adoliquida_det("ITEM").Value = grd_liquidaciones.VisibleRows
            adoliquida_det("FECSER").Value = .Columns(0).Value
            adoliquida_det("NROSER").Value = .Columns(1).Value
            adoliquida_det("CODSERVICIO").Value = .Columns(2).Value
            adoliquida_det("TIPOSER").Value = Trim(.Columns(3).Text)
            adoliquida_det("NOMSERVICIO").Value = .Columns(5).Value
            adoliquida_det("CANTIDAD").Value = .Columns(6).Value
            adoliquida_det("ARANCEL").Value = .Columns(7).Value
            TXTFIELDS(14).Text = Val(TXTFIELDS(14).Text) + (adoliquida_det("ARANCEL").Value * .Columns(6).Value)
            TXTFIELDS(13).Text = (Val(TXTFIELDS(14).Text) - Val(TXTFIELDS(11).Text)) * 0.1
            TXTFIELDS(12).Text = Val(TXTFIELDS(14).Text) - Val(TXTFIELDS(11).Text) + Val(TXTFIELDS(13).Text)
            adoliquida_det("COSTO").Value = .Columns(8).Value
            adoliquida_det("EXENTA").Value = "NO"
            Set grd_liquidaciones.DataSource = adoliquida_det '("ChildCMD").UnderlyingValue
            Call BLOQUEAR_GRILLA(True)
            'grd_liquidaciones.Columns(8).Visible = False
            If Not adoliquida.BOF Then
                    grd_liquidaciones.Columns(0).Visible = False
                    grd_liquidaciones.Columns(1).Visible = False
             End If
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

Private Sub cmdBusTipo_Click()
 Dim frmAyuda As New frmConsult
 Set frmAyuda.ActiveConection = Cnn
 frmAyuda.Modal = True
 frmAyuda.Consulta = "SELECT CODIGO,DESCRIPCION FROM PARAMETROS WHERE TIPO='FACTURA'"
 frmAyuda.Show vbModal
 If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
    With frmAyuda.msGrid
        adoliquida.Fields("TIPO_FAC").Value = .Columns(0).Value
        LBLliquida(11).Caption = .Columns(1).Text
    End With
    TXTFIELDS(9).SetFocus
 End If
frmAyuda.Cerrar
End Sub

Private Sub Cmdficli_Click()
Dim frmAyuda As New frmConsult
 Set frmAyuda.ActiveConection = Cnn
 frmAyuda.Modal = True
 frmAyuda.Consulta = "SELECT CODINS,INSTITUTO FROM INSTITUTOS"
 frmAyuda.Show vbModal
 If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
    With frmAyuda.msGrid
        adoinstituto.Fields("FICHACLINICA").Value = .Columns(0).Value
        adoinstituto.Fields("TIPODOC").Value = .Columns(1).Value
        adoinstituto.Fields("NRODOC").Value = .Columns(2).Value
        LBLliquida(10).Caption = .Columns(4).Text & ", " & .Columns(3).Text
    End With
 End If
frmAyuda.Cerrar
End Sub

Private Sub Form_Load()
    Set adoPrimaryRS = New ADODB.Recordset
    Set CtrVisualizar = New clsValidar
    Set VALIDAR = New ValTecla
    Me.Height = 7500
    Me.Height = 7500
    Me.Left = 80
    Me.Top = 0
    Me.Width = 11800
    Set adoliquida = New ADODB.Recordset
    Set adoliquida_det = New ADODB.Recordset
    Set adoNombres = New ADODB.Recordset
    Set adoFactura = New ADODB.Recordset
    adoliquida.Open "SHAPE {SELECT * FROM COBROS_CABECERA}  AS cobros_cabecera APPEND ({SELECT * FROM COBROS_DETALLE}  AS cobros_detalle RELATE 'FEC_MOV' TO 'FEC_MOV','NRO_MOV' TO 'NRO_MOV') AS cobros_detalle", Cnn, adOpenStatic, adLockBatchOptimistic
    adoNombres.Open "SELECT FICHACLINICA.FICHACLINICA ,PERSONAS.TIPDOC,PERSONAS.NRODOC,PERSONAS.NOMBRES,PERSONAS.APELLIDOS FROM FICHACLINICA,PERSONAS WHERE FICHACLINICA.TIPDOC=PERSONAS.TIPDOC AND FICHACLINICA.NRODOC=PERSONAS.NRODOC", Cnn
    adoFactura.Open "SELECT CODIGO,DESCRIPCION FROM PARAMETROS WHERE TIPO='FACTURA'", Cnn
        If adoliquida_det.State = 0 Then
          If adoliquida.EOF = False And adoliquida.BOF = False Then
            Set adoliquida_det = adoliquida("COBROS_DETALLE").UnderlyingValue
          End If
        End If
    Me.frmliquida.Visible = CtrVisualizar.Visualizar(15, Cnn, IdUsuario)
    Call TAB_LIQUIDA_Click
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

Private Sub grd_liquidaciones_BeforeDelete(Cancel As Integer)
        With adoliquida
            If adoliquida_det("EXENTA").Value = "SI" Then
                TXTFIELDS(5).Text = Val(TXTFIELDS(5).Text) - (adoliquida_det("ARANCEL").Value * adoliquida_det("CANTIDAD").Value)
            Else
                TXTFIELDS(14).Text = Val(TXTFIELDS(14).Text) - (adoliquida_det("ARANCEL").Value * adoliquida_det("CANTIDAD").Value)
                TXTFIELDS(13).Text = (Val(TXTFIELDS(14).Text) - Val(TXTFIELDS(11).Text)) * 0.1
            End If
            TXTFIELDS(13).SetFocus
            SendKeys "{TAB}", True
        End With
End Sub

Private Sub grd_liquidaciones_DblClick()
Select Case grd_liquidaciones.Col
Case 11
     Dim frmAyuda As New frmConsult
     Set frmAyuda.ActiveConection = Cnn
     frmAyuda.Modal = True
     frmAyuda.Consulta = "SELECT DESCRIPCION FROM PARAMETROS WHERE TIPO='EXENTA'"
     frmAyuda.Show vbModal
      If Not frmAyuda.Descargado Then
        With adoliquida
            If adoliquida_det("EXENTA").Value <> frmAyuda.msGrid.Columns(0).Text Then
                adoliquida_det("EXENTA").Value = frmAyuda.msGrid.Columns(0).Text
                If adoliquida_det("EXENTA").Value = "SI" Then
                    TXTFIELDS(5).Text = Val(TXTFIELDS(5).Text) + (adoliquida_det("ARANCEL").Value * adoliquida_det("CANTIDAD").Value)
                    TXTFIELDS(14).Text = Val(TXTFIELDS(14).Text) - (adoliquida_det("ARANCEL").Value * adoliquida_det("CANTIDAD").Value)
                    TXTFIELDS(13).Text = (Val(TXTFIELDS(14).Text) - Val(TXTFIELDS(11).Text)) * 0.1
                Else
                    TXTFIELDS(5).Text = Val(TXTFIELDS(5).Text) - (adoliquida_det("ARANCEL").Value * adoliquida_det("CANTIDAD").Value)
                    TXTFIELDS(14).Text = Val(TXTFIELDS(14).Text) + (adoliquida_det("ARANCEL").Value * adoliquida_det("CANTIDAD").Value)
                    TXTFIELDS(13).Text = (Val(TXTFIELDS(14).Text) - Val(TXTFIELDS(11).Text)) * 0.1
                End If
            End If
        End With
      End If
     frmAyuda.Cerrar
End Select
End Sub

Private Sub grd_liquidaciones_Scroll(Cancel As Integer)
            If Not adoliquida.BOF Then
                    grd_liquidaciones.Columns(0).Visible = False
                    grd_liquidaciones.Columns(1).Visible = False
            End If
End Sub

Private Sub TAB_LIQUIDA_Click()
On Error Resume Next
Select Case Me.TAB_LIQUIDA.SelectedItem.Index
Case 1

    Me.frmliquida.Visible = CtrVisualizar.Visualizar(15, Cnn, IdUsuario)
    If MDIGALENO.tbToolBar.Buttons("MODIFICAR").Enabled = True And MDIGALENO.tbToolBar.Buttons("Add").Enabled = True And MDIGALENO.tbToolBar.Buttons("Delete").Enabled = True Then
        Dim oText As TextBox
        'Enlaza los cuadros de texto con el proveedor de datos
        For Each oText In Me.TXTFIELDS
            If oText.Index < 40 Then
                Set oText.DataSource = adoliquida
            End If
        Next
        Call BLOQUEAR_GRILLA(True)
    Else
        Call BLOQUEAR_GRILLA(False)
    End If
    Set adoliquida_det = adoliquida("COBROS_DETALLE").UnderlyingValue
    Set grd_liquidaciones.DataSource = adoliquida_det '("ChildCMD").UnderlyingValue
    If Not adoliquida.BOF Then
        grd_liquidaciones.Columns(0).Visible = False
        grd_liquidaciones.Columns(1).Visible = False
    Else
        Set grd_liquidaciones.DataSource = Nothing
    End If
    

End Select
End Sub
Private Sub AGREGAR_Click()
  On Error GoTo AddErr
  Cnn.BeginTrans
  Select Case Me.TAB_LIQUIDA.SelectedItem.Index
  Case 1
          With adoliquida
              .AddNew
          End With
          grd_liquidaciones.AllowUpdate = True
          Call BLOQUEAR_GRILLA(False)
          TXTFIELDS(11).Locked = False
          TXTFIELDS(1) = Date
          TXTFIELDS(2) = OPTENER_CODIGO("SELECT NRO_MOV FROM cobros_CABECERA ORDER BY NRO_MOV DESC")
          busFICHA.SetFocus
        Exit Sub
  End Select
AddErr:
    MsgBox Err.Description
End Sub
Private Sub MODIFICAR_Click()
On Error GoTo EditErr
Cnn.BeginTrans
Select Case Me.TAB_LIQUIDA.SelectedItem.Index
Case 1
        LIQUIDACIONES.grd_liquidaciones.AllowUpdate = True
        LIQUIDACIONES.grd_liquidaciones.AllowDelete = True
        Call BLOQUEAR_GRILLA(True)
        TXTFIELDS(11).Locked = False
End Select
Exit Sub
EditErr:
  MsgBox Err.Description
End Sub
Private Sub CANCELAR_Click()
Me.grd_liquidaciones.AllowAddNew = False
Me.grd_liquidaciones.AllowDelete = False
Me.grd_liquidaciones.AllowUpdate = False
Call cmdCancel_Click
i = False
End Sub

Private Sub GRABAR_Click()
'fncBotones (True)
On Error GoTo UpdateErr
Select Case Me.TAB_LIQUIDA.SelectedItem.Index
Case 1
        If adoliquida_det.State = 1 Then
            TXTFIELDS(13).SetFocus
            TXTFIELDS(11).Locked = True
            SendKeys "{TAB}", True
            adoliquida.UpdateBatch
            adoliquida_det.UpdateBatch
            Cnn.CommitTrans
            Me.grd_liquidaciones.AllowUpdate = False
            Me.grd_liquidaciones.AllowDelete = False
            Call BLOQUEAR_GRILLA(True)
            i = False
        Else
             MsgBox ("No agrego ningun detalle a la Orden")
             Set grd_liquidaciones.DataSource = adoliquida_det '("ChildCMD").UnderlyingValue
             MDIGALENO.fncBotones False, "liquidaciones"
        End If
        Exit Sub
End Select
UpdateErr:
    MsgBox Err.Description
End Sub
Private Sub ELIMINAR_Click()
  Dim Respuesta As Integer
  On Error GoTo DeleteErr
  Respuesta = MsgBox("Esta Seguro que desea Eliminar el Registro", vbQuestion + vbYesNo, "Atencion")
  If Respuesta = 6 Then
    Cnn.BeginTrans
    Select Case Me.TAB_LIQUIDA.SelectedItem.Index
    Case 1
        With adoliquida
          .Delete
          .UpdateBatch
          Cnn.CommitTrans
          .Requery
          TXTFIELDS(11).Locked = True
          Call TAB_LIQUIDA_Click
        End With
    End Select
   End If
Exit Sub
DeleteErr:
  MsgBox Err.Description
  Cnn.RollbackTrans
  adoliquida.Requery

End Sub
Public Sub UltimoReg()
ULTIMO_Click
End Sub
Private Sub ULTIMO_Click()
  On Error GoTo GoLastError
  adoliquida.MoveLast
  Exit Sub
GoLastError:
  MsgBox Err.Description
End Sub
Public Sub SiguienteReg()
SIGUIENTE_Click
End Sub
Private Sub SIGUIENTE_Click()
  On Error GoTo GoNextError
  If Not adoliquida.EOF Then adoliquida.MoveNext
  If adoliquida.EOF And adoliquida.RecordCount > 0 Then
    Beep
    adoliquida.MoveLast
  End If
  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub
Public Sub ActualizarReg()
GRABAR_Click
'CtrVisualizar_ABMtollbar CtrVisualizar.PERMISOABM
End Sub

Private Sub ACTUALIZAR_Click()
  On Error GoTo RefreshErr
  adoliquida.Requery
  Set adoliquida_det = adoliquida("cobros_detalle").UnderlyingValue
  Call TAB_LIQUIDA_Click
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub
Public Sub AgregarReg()
AGREGAR_Click
End Sub

Public Sub AnteriorReg()
ANTERIOR_Click
End Sub

Private Sub ANTERIOR_Click()
  On Error GoTo GoPrevError
  If Not adoliquida.BOF Then adoliquida.MovePrevious
  If adoliquida.BOF And adoliquida.RecordCount > 0 Then
    Beep
    adoliquida.MoveFirst
  End If
  Exit Sub
GoPrevError:
  MsgBox Err.Description
End Sub

Public Sub buscarReg()
BUSCAR_Click
End Sub

Private Sub BUSCAR_Click()
 Dim frmAyuda As New frmConsult
 Set frmAyuda.ActiveConection = Cnn
 frmAyuda.Modal = True
' frmAyuda.Consulta = "SELECT * FROM MOVANA1"
 frmAyuda.Consulta = "SELECT * FROM COBROS_CABECERA"
 frmAyuda.Show vbModal
 If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
    
 End If
  If Not frmAyuda.Descargado Then
    With adoliquida
           .MoveFirst
           .Find "fec_mov=" & frmAyuda.msGrid.Columns(1).Value
           .Find "nro_mov=" & frmAyuda.msGrid.Columns(2).Value
    End With
  End If
 frmAyuda.Cerrar
End Sub

Public Sub CancelarReg()
TAB_LIQUIDA_Click
CANCELAR_Click
'CtrVisualizar_ABMtollbar CtrVisualizar.PERMISOABM
End Sub

Public Sub eliminarReg()
ELIMINAR_Click
End Sub

Public Sub ModificarReg()
MODIFICAR_Click
End Sub

Public Sub PrimeroReg()
PRIMERO_Click
End Sub

Private Sub PRIMERO_Click()
  On Error GoTo GoFirstError
  adoliquida.MoveFirst
  Exit Sub
GoFirstError:
  MsgBox Err.Description
End Sub
Public Sub requeryReg()
ACTUALIZAR_Click
End Sub
Private Sub BLOQUEAR_GRILLA(sino As Boolean)
Dim i As Integer
           If sino = True Then
                    For i = 0 To Me.grd_liquidaciones.Columns.Count - 1
                        Me.grd_liquidaciones.Columns(i).Locked = True
                      
                    Next
            Else
                    For i = 0 To Me.grd_liquidaciones.Columns.Count - 1
                        If i < Me.grd_liquidaciones.Columns.Count - 1 Then
                            Me.grd_liquidaciones.Columns(i).Locked = True
                        Else
                            Me.grd_liquidaciones.Columns(i).Locked = False
                        End If
                    Next
            End If
End Sub

Private Sub Form_GotFocus()
CtrVisualizar_ABMtollbar CtrVisualizar.PERMISOABM
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
Private Sub cmdCancel_Click()
  On Error Resume Next
  Cnn.RollbackTrans
  adoliquida.CancelBatch
  adoliquida_det.CancelBatch
  adoliquida.Requery
  'adoORDENES.MoveFirst
  TAB_LIQUIDA_Click
 End Sub

Private Sub txtFields_Change(Index As Integer)
        If Not adoNombres.State = 0 Then
             adoNombres.MoveFirst
             If Not TXTFIELDS(0).Text = "" Then
                adoNombres.Find "FICHACLINICA=" & TXTFIELDS(0).Text
                LBLliquida(10).Caption = adoNombres.Fields("APELLIDOS").Value + ", " + adoNombres.Fields("nombres").Value
             Else
                LBLliquida(10).Caption = ""
             End If
        End If
        If Not adoFactura.State = 0 Then
            adoFactura.MoveFirst
            If Not TXTFIELDS(6).Text = "" Then
                adoFactura.Find "CODIGO='" & TXTFIELDS(6).Text & "'"
                LBLliquida(11).Caption = adoFactura.Fields("DESCRIPCION").Value
            Else
                LBLliquida(11).Caption = ""
            End If
        End If
End Sub

Private Sub TXTFIELDS_Validate(Index As Integer, Cancel As Boolean)
          '  TXTFIELDS(11).Text = Val(TXTFIELDS(11).Text)
             TXTFIELDS(13).Text = (Val(TXTFIELDS(14).Text) - Val(TXTFIELDS(11).Text)) * 0.1
             TXTFIELDS(12).Text = Val(TXTFIELDS(14).Text) + Val(TXTFIELDS(5).Text) - Val(TXTFIELDS(11).Text) + Val(TXTFIELDS(13).Text)
End Sub
