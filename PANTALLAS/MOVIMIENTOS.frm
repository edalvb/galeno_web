VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form ORDENES 
   AutoRedraw      =   -1  'True
   Caption         =   "ORDENES DE SERVICIOS"
   ClientHeight    =   7815
   ClientLeft      =   285
   ClientTop       =   1620
   ClientWidth     =   11595
   ClipControls    =   0   'False
   Icon            =   "MOVIMIENTOS.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7815
   ScaleWidth      =   11595
   Visible         =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame2 
      Height          =   6435
      Left            =   120
      TabIndex        =   24
      Top             =   480
      Visible         =   0   'False
      Width           =   11265
      Begin MSDataGridLib.DataGrid grd_ordenes 
         Height          =   2655
         Left            =   240
         TabIndex        =   36
         Top             =   3720
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   4683
         _Version        =   393216
         AllowUpdate     =   0   'False
         ForeColor       =   -2147483642
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
      Begin VB.CommandButton busespecialidad 
         DragIcon        =   "MOVIMIENTOS.frx":030A
         Height          =   375
         Left            =   5640
         Picture         =   "MOVIMIENTOS.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton busficha 
         DragIcon        =   "MOVIMIENTOS.frx":0B8E
         Height          =   375
         Left            =   6480
         Picture         =   "MOVIMIENTOS.frx":0FD0
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdbusres 
         DragIcon        =   "MOVIMIENTOS.frx":1412
         Height          =   375
         Left            =   3360
         Picture         =   "MOVIMIENTOS.frx":1854
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2880
         Width           =   495
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "RESPECOTIPDOC"
         Height          =   285
         Index           =   14
         Left            =   3960
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   15
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "RESPECONRODOC"
         Height          =   285
         Index           =   15
         Left            =   4800
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   22
         Top             =   2880
         Width           =   1365
      End
      Begin VB.CommandButton bustipomov 
         DragIcon        =   "MOVIMIENTOS.frx":1C96
         Height          =   375
         Left            =   120
         Picture         =   "MOVIMIENTOS.frx":20D8
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "TIPOMOV"
         Height          =   285
         Index           =   13
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   2520
         Width           =   945
      End
      Begin VB.CommandButton busanalisis 
         DragIcon        =   "MOVIMIENTOS.frx":251A
         Height          =   375
         Left            =   240
         Picture         =   "MOVIMIENTOS.frx":295C
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   3240
         Width           =   495
      End
      Begin VB.CommandButton businstitut 
         DragIcon        =   "MOVIMIENTOS.frx":2D9E
         Height          =   375
         Left            =   120
         Picture         =   "MOVIMIENTOS.frx":31E0
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2040
         Width           =   495
      End
      Begin VB.CommandButton busaseg 
         DragIcon        =   "MOVIMIENTOS.frx":3622
         Height          =   375
         Left            =   120
         Picture         =   "MOVIMIENTOS.frx":3A64
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton busdoc 
         DragIcon        =   "MOVIMIENTOS.frx":3EA6
         Height          =   375
         Left            =   120
         Picture         =   "MOVIMIENTOS.frx":42E8
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "FEC_PED"
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
         Index           =   1
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   1335
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
         Index           =   2
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "CODDR"
         Height          =   285
         Index           =   6
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "TIPO_SEGURO"
         Height          =   285
         Index           =   11
         Left            =   9000
         TabIndex        =   11
         Top             =   960
         Width           =   1200
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "COD_EMPRESA"
         Height          =   285
         Index           =   8
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1005
         Width           =   705
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "NRO_SEGURO"
         Height          =   285
         Index           =   9
         Left            =   3915
         TabIndex        =   9
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "IDE_ASEGURADO"
         Height          =   285
         Index           =   10
         Left            =   6360
         TabIndex        =   10
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "NRODOC"
         Height          =   285
         Index           =   5
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   600
         Width           =   1245
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "TIPDOC"
         Height          =   285
         Index           =   4
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "FICHACLINICA"
         ForeColor       =   &H000000C0&
         Height          =   285
         Index           =   3
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   270
         Width           =   1215
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "CODINS"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd/MM/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   2400
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   2160
         Width           =   945
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "COD_ESPECIALIDAD"
         Height          =   285
         Index           =   7
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1560
         Width           =   675
      End
      Begin VB.Label LBLORDEN 
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
         Index           =   6
         Left            =   3360
         TabIndex        =   47
         Top             =   2520
         Width           =   1485
      End
      Begin VB.Label LBLORDEN 
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
         Left            =   3360
         TabIndex        =   46
         Top             =   2160
         Width           =   1965
      End
      Begin VB.Label LBLORDEN 
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
         Index           =   7
         Left            =   6240
         TabIndex        =   45
         Top             =   2880
         Width           =   1965
      End
      Begin VB.Label lblLabels 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Responsable Tipo y Nro de Documento"
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
         Left            =   240
         TabIndex        =   44
         Top             =   3000
         Width           =   2970
      End
      Begin VB.Label lblLabels 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Seguro"
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
         Left            =   7800
         TabIndex        =   43
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label lblLabels 
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
         Index           =   5
         Left            =   4440
         TabIndex        =   42
         Top             =   600
         Width           =   1740
      End
      Begin VB.Label LBLORDEN 
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
         Index           =   4
         Left            =   8040
         TabIndex        =   41
         Top             =   1560
         Width           =   1965
      End
      Begin VB.Label LBLORDEN 
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
         Index           =   3
         Left            =   2160
         TabIndex        =   40
         Top             =   1680
         Width           =   1965
      End
      Begin VB.Label LBLORDEN 
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
         Index           =   2
         Left            =   360
         TabIndex        =   39
         Top             =   1320
         Width           =   1965
      End
      Begin VB.Label lblLabels 
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
         Index           =   1
         Left            =   840
         TabIndex        =   38
         Top             =   3360
         Width           =   1380
      End
      Begin VB.Label lblLabels 
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
         Index           =   0
         Left            =   720
         TabIndex        =   37
         Top             =   2520
         Width           =   1530
      End
      Begin VB.Label lblLabels 
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
         Index           =   28
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   1230
      End
      Begin VB.Label lblLabels 
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
         Index           =   27
         Left            =   3225
         TabIndex        =   34
         Top             =   255
         Width           =   1410
      End
      Begin VB.Label lblLabels 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor"
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
         Index           =   26
         Left            =   720
         TabIndex        =   33
         Top             =   1695
         Width           =   510
      End
      Begin VB.Label lblLabels 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "IDE  Asegurado"
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
         Index           =   25
         Left            =   5160
         TabIndex        =   32
         Top             =   960
         Width           =   1185
      End
      Begin VB.Label lblLabels 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa"
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
         Index           =   24
         Left            =   720
         TabIndex        =   31
         Top             =   1020
         Width           =   660
      End
      Begin VB.Label lblLabels 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Numero de Seguro:"
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
         Index           =   23
         Left            =   2280
         TabIndex        =   30
         Top             =   960
         Width           =   1635
      End
      Begin VB.Label LBLORDEN 
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
         Index           =   1
         Left            =   6480
         TabIndex        =   29
         Top             =   600
         Width           =   1965
      End
      Begin VB.Label lblLabels 
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
         Index           =   18
         Left            =   90
         TabIndex        =   28
         Top             =   660
         Width           =   1950
      End
      Begin VB.Label lblLabels 
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
         Index           =   17
         Left            =   7080
         TabIndex        =   27
         Top             =   300
         Width           =   1050
      End
      Begin VB.Label lblLabels 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo de Instituto"
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
         Index           =   16
         Left            =   720
         TabIndex        =   26
         Top             =   2145
         Width           =   1500
      End
      Begin VB.Label lblLabels 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Especialidad"
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
         Left            =   6960
         TabIndex        =   25
         Top             =   1560
         Width           =   960
      End
   End
   Begin MSComctlLib.TabStrip TAB_ORDENES 
      Height          =   6930
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   12224
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "EMISION DE ORDENES"
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
Attribute VB_Name = "ORDENES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents adoORDENES     As ADODB.Recordset
Attribute adoORDENES.VB_VarHelpID = -1
Dim WithEvents adoORDENES_CAB As ADODB.Recordset
Attribute adoORDENES_CAB.VB_VarHelpID = -1
Dim WithEvents adoORDENES_det As ADODB.Recordset
Attribute adoORDENES_det.VB_VarHelpID = -1
Dim WithEvents adoFCNombres   As ADODB.Recordset
Attribute adoFCNombres.VB_VarHelpID = -1
Dim WithEvents adoFCempresas  As ADODB.Recordset
Attribute adoFCempresas.VB_VarHelpID = -1
Dim WithEvents adoFCdoctores  As ADODB.Recordset
Attribute adoFCdoctores.VB_VarHelpID = -1
Dim WithEvents adoFCespecialidad As ADODB.Recordset
Attribute adoFCespecialidad.VB_VarHelpID = -1
Dim WithEvents adoFCinstituto As ADODB.Recordset
Attribute adoFCinstituto.VB_VarHelpID = -1
Dim WithEvents adoFCtipomov As ADODB.Recordset
Attribute adoFCtipomov.VB_VarHelpID = -1
Dim WithEvents adoFCrespoeco As ADODB.Recordset
Attribute adoFCrespoeco.VB_VarHelpID = -1
Dim strAuxiliar As String
Public Sub imprimir()
        With DataGaleno.rsORDERNESSERVICIOS
           If .State = adStateClosed Then
                    .Open 'colocar este codigo a todos los informes
                    .Filter = ""
                    .Requery
                    .MoveFirst
                    .Filter = "FEC_PED" & " = '" & (TXTFIELDS(1).Text) & "' " & " And NROMOV" & " = " & (TXTFIELDS(2).Text)
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
           rpt_ordenes.Refresh
           rpt_ordenes.Show
           .Filter = ""
           If .State = adStateOpen Then 'colocar este codigo a todos los informes
              .Close
              strSQL = ""
           End If
        End With
End Sub
Private Sub AGREGAR_Click()
  On Error GoTo AddErr
  Cnn.BeginTrans
  Select Case Me.TAB_ORDENES.SelectedItem.Index
  Case 1
          With adoORDENES
              .AddNew
          End With
          grd_ordenes.AllowUpdate = True
          grd_ordenes.AllowDelete = True
          Call BLOQUEAR_GRILLA(False)
          LBLORDEN(1).Caption = ""
          LBLORDEN(2).Caption = ""
          LBLORDEN(3).Caption = ""
          LBLORDEN(4).Caption = ""
          LBLORDEN(5).Caption = ""
          LBLORDEN(6).Caption = ""
          LBLORDEN(7).Caption = ""
          TXTFIELDS(1) = Date
          TXTFIELDS(2) = OPTENER_CODIGO("SELECT NROMOV FROM ORDENSERVICIOCABECERA ORDER BY NROMOV DESC")
          busFICHA.SetFocus
        Exit Sub
  End Select
AddErr:
    MsgBox Err.Description
End Sub
Private Sub ACTUALIZAR_Click()
  On Error GoTo RefreshErr
  adoORDENES.Requery
  Set adoORDENES_det = adoORDENES("ORDENSERVICIODETALLE").UnderlyingValue
  Call TAB_ORDENES_Click
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub
Private Sub MODIFICAR_Click()
On Error GoTo EditErr
Cnn.BeginTrans
Select Case Me.TAB_ORDENES.SelectedItem.Index
Case 1
        ordenes.grd_ordenes.AllowUpdate = True
        ordenes.grd_ordenes.AllowDelete = True
        Call BLOQUEAR_GRILLA(False)
End Select
Exit Sub
EditErr:
  MsgBox Err.Description
End Sub

Private Sub ANTERIOR_Click()
  On Error GoTo GoPrevError

  If Not adoORDENES.BOF Then adoORDENES.MovePrevious
  If adoORDENES.BOF And adoORDENES.RecordCount > 0 Then
    Beep
    adoORDENES.MoveFirst
  End If
  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub busanalisis_Click()
 On Error GoTo GoPrevError
 If adoORDENES.Status <> 0 Then
   adoORDENES.UpdateBatch
 End If
 Dim frmAyuda As New frmConsult
 Set frmAyuda.ActiveConection = Cnn
 frmAyuda.Modal = True
 frmAyuda.Consulta = "SELECT SERVICIO.CODSERVICIO, SERVICIO.TIPOSER, CODSER, NOMSERVICIO, NOMANALC,UNIMED,CODANALS,ARANCELES.ARANCEL,ARANCELES.COSTO FROM SERVICIO,ARANCELES WHERE SERVICIO.CODSERVICIO=ARANCELES.CODSERVICIO AND SERVICIO.TIPOSER=ARANCELES.TIPOSER AND ARANCELES.CODDR=" & TXTFIELDS(6).Text & " AND ARANCELES.CODINS=" & TXTFIELDS(12).Text
 frmAyuda.msGrid.Columns(8).Visible = False
 frmAyuda.Show vbModal
 If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
    With frmAyuda.msGrid
        Set adoORDENES_det = adoORDENES("ORDENSERVICIODETALLE").UnderlyingValue
        Call TAB_ORDENES_Click
        adoORDENES_det.Find "CODSERVICIO= " & .Columns(0).Text
        If adoORDENES_det.EOF Then
            adoORDENES_det.AddNew
            adoORDENES_det("CODSERVICIO").Value = .Columns(0).Value
            adoORDENES_det("TIPOSER").Value = .Columns(1).Value
            adoORDENES_det("CODSER").Value = .Columns(2).Value
            adoORDENES_det("NOMSERVICIO").Value = .Columns(3).Value
            adoORDENES_det("FEC_PRO").Value = TXTFIELDS(1).Text
            adoORDENES_det("FEC_ENT").Value = TXTFIELDS(1).Text
            adoORDENES_det("PRE_VENTA").Value = .Columns(7).Value
            adoORDENES_det("PRE_COSTO").Value = .Columns(8).Value
            adoORDENES_det("URGENTE").Value = "N"
            Set grd_ordenes.DataSource = adoORDENES_det '("ChildCMD").UnderlyingValue
            Call BLOQUEAR_GRILLA(False)
            grd_ordenes.Columns(8).Visible = False
            grd_ordenes.Columns(9).Visible = False
            grd_ordenes.Columns(10).Visible = False
            grd_ordenes.Columns(11).Visible = False
            grd_ordenes.Columns(12).Visible = False
            grd_ordenes.Columns(13).Visible = False
        Else
            MsgBox ("El Servicio que ha elegido ya existe en el detalle por favor elija otro")
       End If
    End With
  End If
 frmAyuda.Cerrar
 Exit Sub
GoPrevError:
  MsgBox Err.Description, vbCritical
'Call CANCELAR_Click
End Sub

Private Sub busaseg_Click()
 Dim frmAyuda As New frmConsult
 Set frmAyuda.ActiveConection = Cnn
 frmAyuda.Modal = True
 frmAyuda.Consulta = "SELECT * FROM EMPRESA"
 frmAyuda.Show vbModal
 If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
    With frmAyuda.msGrid
        adoORDENES.Fields("COD_EMPRESA").Value = .Columns(0).Value
        LBLORDEN(2).Caption = .Columns(1).Text
      End With
  End If
 frmAyuda.Cerrar
 busdoc.SetFocus
 End Sub

Private Sub BUSCAR_Click()
 Dim frmAyuda As New frmConsult
 Set frmAyuda.ActiveConection = Cnn
 frmAyuda.Modal = True
' frmAyuda.Consulta = "SELECT * FROM MOVANA1"
 frmAyuda.Consulta = "SELECT * FROM ORDENSERVICIOCABECERA"
 frmAyuda.Show vbModal
 If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
    
 End If
  If Not frmAyuda.Descargado Then
    With adoORDENES
           .MoveFirst
           .Find "fec_ped=" & frmAyuda.msGrid.Columns(1).Value
           .Find "nromov=" & frmAyuda.msGrid.Columns(2).Value
    End With
  End If
 frmAyuda.Cerrar
End Sub

Private Sub BUSCAR_MOV_Click()
 Dim frmAyuda As New frmConsult
 Set frmAyuda.ActiveConection = Cnn
 frmAyuda.Modal = True
 frmAyuda.Consulta = "SELECT * FROM ordenserviciocabecera "
 frmAyuda.Show vbModal
 If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
    With frmAyuda.msGrid
        TXTFIELDS(11).Text = .Columns(0).Text
        TXTFIELDS(10).Text = .Columns(1).Text
        TXTFIELDS(8).Text = .Columns(7).Text
        TXTFIELDS(9).Text = .Columns(6).Text
        TXTFIELDS(7).Locked = False
        
    End With
 End If
frmAyuda.Cerrar
End Sub

Private Sub buscarservicio_Click()
Dim frmAyuda As New frmConsult
 Set frmAyuda.ActiveConection = Cnn
 frmAyuda.Modal = True
 frmAyuda.Consulta = "SELECT * FROM tiposervicio"
 
 frmAyuda.Show vbModal
 If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
    With frmAyuda.msGrid
        TXTFIELDS(7).Text = TXTFIELDS(7).Text + .Columns(1).Text
        
    End With
 End If
frmAyuda.Cerrar

End Sub

Private Sub busdoc_Click()
  
 Dim frmAyuda As New frmConsult
 Set frmAyuda.ActiveConection = Cnn
 frmAyuda.Modal = True
 frmAyuda.Consulta = "SELECT DOCTOR.CODDR,PERSONAS.APELLIDOS,PERSONAS.NOMBRES, PERSONAS.TIPDOC,PERSONAS.NRODOC,DOCTOR.REGPRF FROM DOCTOR  LEFT OUTER JOIN PERSONAS ON DOCTOR.TIPDOC = PERSONAS.TIPDOC AND DOCTOR.NRODOC = PERSONAS.NRODOC"
 frmAyuda.Show vbModal
 If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
    With frmAyuda.msGrid
        adoORDENES.Fields("CODDR").Value = .Columns(0).Value
        LBLORDEN(3).Caption = .Columns(1).Text & ", " & .Columns(2).Text
        adoFCespecialidad.MoveFirst
        adoFCespecialidad.Find "CODDR=" & adoORDENES.Fields("CODDR").Value
        adoORDENES.Fields("COD_ESPECIALIDAD").Value = adoFCespecialidad.Fields("COD_ESPECIALIDAD").Value
        LBLORDEN(4).Caption = adoFCespecialidad.Fields("nombre").Value
              
    End With
 End If
frmAyuda.Cerrar
businstitut.SetFocus

End Sub

Private Sub busespecialidad_Click()
 Dim frmAyuda As New frmConsult
 Set frmAyuda.ActiveConection = Cnn
 frmAyuda.Modal = True
 frmAyuda.Consulta = "SELECT ESPECIALIDAD.COD_ESPECIALIDAD, ESPECIALIDAD.NOMBRE FROM ESPECIALIDAD,DOCTORESPECIALIDAD WHERE DOCTORESPECIALIDAD.COD_ESPECIALIDAD=ESPECIALIDAD.COD_ESPECIALIDAD AND DOCTORESPECIALIDAD.CODDR=" & TXTFIELDS(6).Text
 frmAyuda.Show vbModal
 If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
    With frmAyuda.msGrid
         adoORDENES.Fields("COD_ESPECIALIDAD").Value = .Columns(0).Value
         LBLORDEN(4).Caption = .Columns(1).Text
    End With
 End If
frmAyuda.Cerrar
businstitut.SetFocus

End Sub


Private Sub busficha_Click()
 Dim frmAyuda As New frmConsult
 Set frmAyuda.ActiveConection = Cnn
 frmAyuda.Modal = True
 frmAyuda.Consulta = "SELECT FICHACLINICA,PERSONAS.TIPDOC,PERSONAS.NRODOC,PERSONAS.NOMBRES,PERSONAS.APELLIDOS FROM FICHACLINICA,PERSONAS WHERE FICHACLINICA.TIPDOC=PERSONAS.TIPDOC AND FICHACLINICA.NRODOC=PERSONAS.NRODOC"
 frmAyuda.Show vbModal
 If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
    With frmAyuda.msGrid
        adoORDENES.Fields("FICHACLINICA").Value = .Columns(0).Value
        adoORDENES.Fields("TIPDOC").Value = .Columns(1).Value
        adoORDENES.Fields("NRODOC").Value = .Columns(2).Value
        TXTFIELDS(3).Text = .Columns(0).Text
        TXTFIELDS(4).Text = .Columns(1).Text
        TXTFIELDS(5).Text = .Columns(2).Text
        LBLORDEN(1).Caption = .Columns(4).Text & ", " & .Columns(3).Text
    End With
 End If
frmAyuda.Cerrar
busaseg.SetFocus

End Sub

Private Sub businstitut_Click()
 Dim frmAyuda As New frmConsult
 Set frmAyuda.ActiveConection = Cnn
 frmAyuda.Modal = True
 frmAyuda.Consulta = "SELECT CODINS,INSTITUTO,TELEF1,TELEF2 FROM institutos"
 frmAyuda.Show vbModal
 If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
    With frmAyuda.msGrid
        adoORDENES.Fields("CODINS").Value = .Columns(0).Value
        LBLORDEN(5).Caption = .Columns(1).Text
    End With
 End If
frmAyuda.Cerrar
bustipomov.SetFocus

End Sub

Private Sub CANCELAR_Click()
Me.grd_ordenes.AllowAddNew = False
Me.grd_ordenes.AllowDelete = False
Me.grd_ordenes.AllowUpdate = False
Call cmdCancel_Click
i = False
End Sub
Private Sub BUSMOV_Click()
  
 Dim frmAyuda As New frmConsult
 Set frmAyuda.ActiveConection = Cnn
 frmAyuda.Modal = True
 frmAyuda.Consulta = "SELECT * FROM ordenserviciocabecera"
 frmAyuda.Show vbModal
 If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
    With frmAyuda.msGrid
        TXTFIELDS(0).Text = .Columns(0).Text
        TXTFIELDS(1).Text = .Columns(1).Text
        TXTFIELDS(2).Text = .Columns(7).Text
        TXTFIELDS(3).Text = .Columns(6).Text
        TXTFIELDS(4).Locked = False
        TXTFIELDS(5).Locked = False
    End With
 End If
frmAyuda.Cerrar

End Sub

Private Sub bustipomov_Click()
 Dim frmAyuda As New frmConsult
 Set frmAyuda.ActiveConection = Cnn
 frmAyuda.Modal = True
 frmAyuda.Consulta = "SELECT * FROM TIPOMOV"
 frmAyuda.Show vbModal
 If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
    With frmAyuda.msGrid
    
        adoORDENES.Fields("TIPOMOV").Value = .Columns(0).Value
        LBLORDEN(6).Caption = .Columns(1).Text
                        
    End With
 End If
frmAyuda.Cerrar
cmdbusres.SetFocus
End Sub
Private Sub cmdbusres_Click()
 Dim frmAyuda As New frmConsult
 Set frmAyuda.ActiveConection = Cnn
 frmAyuda.Modal = True
 frmAyuda.Consulta = "SELECT TIPDOC ,NRODOC,NOMBRES ,APELLIDOS FROM PERSONAS"
 frmAyuda.Show vbModal
 If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
    With frmAyuda.msGrid
        strAuxiliar = .Columns(1).Text
        adoORDENES.Fields("RESPECOTIPDOC").Value = .Columns(0).Value
        adoORDENES.Fields("RESPECONRODOC").Value = .Columns(1).Value
        LBLORDEN(7).Caption = .Columns(3).Text & ", " & .Columns(2).Text
    End With
 End If
frmAyuda.Cerrar
busanalisis.SetFocus
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
Private Sub Form_Load()
Set CtrVisualizar = New clsValidar
Set VALIDAR = New ValTecla
Me.Height = 7500
Me.Left = 80
Me.Top = 0
Me.Width = 11800
Set adoORDENES = New ADODB.Recordset
Set adoORDENES_CAB = New ADODB.Recordset
Set adoORDENES_det = New ADODB.Recordset
Set adoFCNombres = New ADODB.Recordset
Set adoFCempresas = New ADODB.Recordset
Set adoFCdoctores = New ADODB.Recordset
Set adoFCespecialidad = New ADODB.Recordset
Set adoFCinstituto = New ADODB.Recordset
Set adoFCtipomov = New ADODB.Recordset
Set adoFCrespoeco = New ADODB.Recordset
adoFCNombres.Open "Select FICHACLINICA.FICHACLINICA,FICHACLINICA.TIPDOC,FICHACLINICA.NRODOC,NOMBRES,APELLIDOS FROM FICHACLINICA,PERSONAS WHERE FICHACLINICA.TIPDOC=PERSONAS.TIPDOC AND FICHACLINICA.NRODOC=PERSONAS.NRODOC  ", Cnn
adoFCempresas.Open "Select cod_empresa,nombre from empresa", Cnn
adoFCdoctores.Open "Select DOCTOR.CODDR,DOCTOR.TIPDOC,DOCTOR.NRODOC,NOMBRES,APELLIDOS FROM DOCTOR,PERSONAS WHERE DOCTOR.TIPDOC=PERSONAS.TIPDOC AND DOCTOR.NRODOC=PERSONAS.NRODOC ", Cnn
adoFCespecialidad.Open "Select * FROM DOCTORESPECIALIDAD", Cnn
'adoFCespecialidad.Open "Select ESPECIALIDAD.COD_ESPECIALIDAD,ESPECIALIDAD.NOMBRE FROM ESPECIALIDAD,DOCTOR  WHERE ESPECIALIDAD.COD_ESPECIALIDAD=DOCTOR.COD_ESPECIALIDAD ", Cnn
adoFCinstituto.Open "Select CODINS,INSTITUTO FROM INSTITUTOS", Cnn
adoFCtipomov.Open "Select TIPOMOV,NOMTIPOMOV FROM TIPOMOV", Cnn
'adoFCrespoeco.Open "Select * from PERSONAS", Cnn
adoFCrespoeco.Open "SELECT TIPDOC ,NRODOC,NOMBRES ,APELLIDOS FROM PERSONAS", Cnn
adoORDENES.Open "SHAPE {SELECT * FROM ORDENSERVICIOCABECERA}  AS ORDENSERVICIOCABECERA APPEND ({SELECT CODSERVICIO, TIPOSER,CODSER,NOMSERVICIO, URGENTE, FEC_PRO, FEC_ENT,PRE_VENTA,PRE_COSTO, FEC_PED, NROMOV, TIPDOC, NRODOC ,FICHACLINICA FROM ORDENSERVICIODETALLE} AS ORDENSERVICIODETALLE RELATE 'FICHACLINICA' TO 'FICHACLINICA','FEC_PED' TO 'FEC_PED','NROMOV' TO 'NROMOV','TIPDOC' TO 'TIPDOC','NRODOC' TO 'NRODOC') AS ORDENSERVICIODETALLE", Cnn, adOpenStatic, adLockBatchOptimistic
Me.Frame2.Visible = CtrVisualizar.Visualizar(15, Cnn, IdUsuario)
Call TAB_ORDENES_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
 
 adoORDENES.Close
 adoORDENES_CAB.Close
 adoORDENES_det.Close
 adoFCNombres.Close
 adoFCempresas.Close
 adoFCdoctores.Close
 adoFCespecialidad.Close
 adoFCinstituto.Close
 adoFCtipomov.Close
 adoFCrespoeco.Close
 Cnn.RollbackTrans
 MDIGALENO.fncBotones True, ""
 VALIDAR.habilitarToolbar MDIGALENO.tbToolBar
End Sub
Private Sub SIGUIENTE_Click()
  On Error GoTo GoNextError

  If Not adoORDENES.EOF Then adoORDENES.MoveNext
  If adoORDENES.EOF And adoORDENES.RecordCount > 0 Then
    Beep
    adoORDENES.MoveLast
  End If
  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub
Private Sub PRIMERO_Click()
  On Error GoTo GoFirstError

  adoORDENES.MoveFirst
  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub TAB_ORDENES_Click()
On Error Resume Next
Select Case Me.TAB_ORDENES.SelectedItem.Index
Case 1
    If MDIGALENO.tbToolBar.Buttons("MODIFICAR").Enabled = True And MDIGALENO.tbToolBar.Buttons("Add").Enabled = True And MDIGALENO.tbToolBar.Buttons("Delete").Enabled = True Then
        Dim oText As TextBox
        'Enlaza los cuadros de texto con el proveedor de datos
        For Each oText In Me.TXTFIELDS
            If oText.Index < 40 Then
                Set oText.DataSource = adoORDENES
            End If
        Next

        Call BLOQUEAR_GRILLA(True)
    Else
        Call BLOQUEAR_GRILLA(False)
    End If
    Set adoORDENES_det = adoORDENES("ORDENSERVICIODETALLE").UnderlyingValue
    If TXTFIELDS(8).Text = "" Then
           LBLORDEN(2).Caption = ""
    End If
    Set grd_ordenes.DataSource = adoORDENES_det '("ChildCMD").UnderlyingValue
    If Not adoORDENES.BOF Then
       grd_ordenes.Columns(8).Visible = False
       grd_ordenes.Columns(9).Visible = False
       grd_ordenes.Columns(10).Visible = False
       grd_ordenes.Columns(11).Visible = False
       grd_ordenes.Columns(12).Visible = False
       grd_ordenes.Columns(13).Visible = False
       
    End If

End Select
End Sub
Private Sub ULTIMO_Click()
  On Error GoTo GoLastError
  adoORDENES.MoveLast
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
Select Case Me.TAB_ORDENES.SelectedItem.Index
Case 1
        If adoORDENES_det.EOF = False Then
            SendKeys "{TAB}", True
            adoORDENES.UpdateBatch
            adoORDENES_det.UpdateBatch
            Cnn.CommitTrans
            Me.grd_ordenes.AllowUpdate = False
            Me.grd_ordenes.AllowDelete = False
            Call BLOQUEAR_GRILLA(True)
            i = False
            MDIGALENO.fncBotones True, ""
        Else
            MsgBox "No agrego ningun detalle a la Orden", vbExclamation
            Set grd_ordenes.DataSource = adoORDENES_det '("ChildCMD").UnderlyingValue
            MDIGALENO.fncBotones False, "ORDENES"
        End If
        Exit Sub
End Select
UpdateErr:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub ELIMINAR_Click()
  Dim Respuesta As Integer
  On Error GoTo DeleteErr
  Respuesta = MsgBox("Esta Seguro que desea Eliminar el Registro", vbQuestion + vbYesNo, "Atencion")
  If Respuesta = 6 Then
    Cnn.BeginTrans
    Select Case Me.TAB_ORDENES.SelectedItem.Index
    Case 1
        With adoORDENES
          .Delete
          .UpdateBatch
          Cnn.CommitTrans
          .Requery
          Call TAB_ORDENES_Click
        End With
    End Select
   End If
Exit Sub
DeleteErr:
  MsgBox Err.Description
'  Cnn.RollbackTrans
End Sub
Private Sub cmdCancel_Click()
  On Error Resume Next
  Cnn.RollbackTrans
  adoORDENES.CancelBatch
  adoORDENES_det.CancelBatch
  adoORDENES.Requery
  'adoORDENES.MoveFirst
  TAB_ORDENES_Click
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
'CtrVisualizar_ABMtollbar CtrVisualizar.PERMISOABM
End Sub
Public Sub requeryReg()
ACTUALIZAR_Click
End Sub
Public Sub buscarReg()
BUSCAR_Click
End Sub
Public Sub CancelarReg()
adoORDENES.MoveFirst
TAB_ORDENES_Click
CANCELAR_Click
'CtrVisualizar_ABMtollbar CtrVisualizar.PERMISOABM
End Sub
Private Sub BLOQUEAR_GRILLA(sino As Boolean)
Dim i As Integer
           If sino = True Then
                    For i = 0 To Me.grd_ordenes.Columns.Count - 1
                        Me.grd_ordenes.Columns(i).Locked = True
                        
                    Next
            Else
                    For i = 0 To Me.grd_ordenes.Columns.Count - 1
                        If i < 4 Then
                            Me.grd_ordenes.Columns(i).Locked = True
                        Else
                            Me.grd_ordenes.Columns(i).Locked = False
                        End If
                    Next
            End If
End Sub

Private Sub txtFields_Change(Index As Integer)
        If Not adoFCNombres.State = 0 Then
             adoFCNombres.MoveFirst
             If Not TXTFIELDS(3).Text = "" Then
                adoFCNombres.Find "FICHACLINICA=" & TXTFIELDS(3).Text
                LBLORDEN(1).Caption = adoFCNombres.Fields("APELLIDOS").Value + ", " + adoFCNombres.Fields("nombres").Value
             Else
                LBLORDEN(1).Caption = ""
             End If
        End If
        If Not adoFCempresas.State = 0 Then
            adoFCempresas.MoveFirst
            If Not TXTFIELDS(8).Text = "" Then
                adoFCempresas.Find "cod_empresa=" & TXTFIELDS(8).Text
                LBLORDEN(2).Caption = adoFCempresas.Fields("nombre").Value
            Else
                LBLORDEN(2).Caption = ""
            End If
        End If
        If Not adoFCdoctores.State = 0 Then
            adoFCdoctores.MoveFirst
            If Not TXTFIELDS(6).Text = "" Then
                adoFCdoctores.Find "CODDR=" & TXTFIELDS(6).Text
                LBLORDEN(3).Caption = adoFCdoctores.Fields("APELLIDOS").Value + "," + adoFCdoctores.Fields("nombres").Value
            Else
                LBLORDEN(3).Caption = ""
            End If
        End If
        If Not adoFCespecialidad.State = 0 Then
            adoFCespecialidad.MoveFirst
            If Not TXTFIELDS(7).Text = "" Then
               adoFCespecialidad.Find "COD_ESPECIALIDAD=" & TXTFIELDS(7).Text
               LBLORDEN(4).Caption = adoFCespecialidad.Fields("nombre").Value
            Else
                LBLORDEN(4).Caption = ""
            End If
        End If
        If Not adoFCtipomov.State = 0 Then
             adoFCinstituto.MoveFirst
             If Not TXTFIELDS(12).Text = "" Then
                 adoFCinstituto.Find "CODINS=" & TXTFIELDS(12).Text
                 LBLORDEN(5).Caption = adoFCinstituto.Fields("Instituto").Value
             Else
                 LBLORDEN(5).Caption = ""
             End If
        End If
        If Not adoFCtipomov.State = 0 Then
            adoFCtipomov.MoveFirst
            If Not TXTFIELDS(13).Text = "" Then
                adoFCtipomov.Find "TIPOMOV=" & TXTFIELDS(13).Text
                LBLORDEN(6).Caption = adoFCtipomov.Fields("NOMTIPOMOV").Value
            Else
                LBLORDEN(6).Caption = ""
            End If
        End If
        
        If Not adoFCrespoeco.State = 0 Then
            adoFCrespoeco.MoveFirst
            If Not TXTFIELDS(14).Text = "" Then
                adoFCrespoeco.Find "TIPDOC='" & Trim$(TXTFIELDS(14).Text) & "'"
            End If
        End If
        If Not adoFCrespoeco.State = 0 Then
            adoFCrespoeco.MoveFirst
            If Not TXTFIELDS(15).Text = "" Then
                adoFCrespoeco.Find "NRODOC=" & Trim$(TXTFIELDS(15).Text)
                LBLORDEN(7).Caption = adoFCrespoeco.Fields("APELLIDOS").Value + ", " + adoFCrespoeco.Fields("NOMBRES").Value
            Else
                LBLORDEN(7).Caption = ""
            End If
        End If
        
End Sub
