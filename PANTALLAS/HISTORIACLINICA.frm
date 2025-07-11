VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form ATENCIONMEDICA 
   Caption         =   "ATENCION MEDICA"
   ClientHeight    =   7215
   ClientLeft      =   225
   ClientTop       =   510
   ClientWidth     =   11565
   Icon            =   "HISTORIACLINICA.frx":0000
   MDIChild        =   -1  'True
   ScaleHeight     =   52692.58
   ScaleMode       =   0  'User
   ScaleWidth      =   6867.039
   Begin VB.Frame frmevolucion 
      Height          =   4815
      Left            =   120
      TabIndex        =   36
      Top             =   2040
      Width           =   11175
      Begin VB.CommandButton cmd_imp_med 
         DragIcon        =   "HISTORIACLINICA.frx":0ABA
         Height          =   615
         Left            =   5400
         Picture         =   "HISTORIACLINICA.frx":0EFC
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   2760
         Width           =   495
      End
      Begin VB.CommandButton cmd_imp_pro 
         DragIcon        =   "HISTORIACLINICA.frx":1566
         Height          =   615
         Left            =   10440
         Picture         =   "HISTORIACLINICA.frx":19A8
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   480
         Width           =   495
      End
      Begin VB.CommandButton cmd_imp_diag 
         DragIcon        =   "HISTORIACLINICA.frx":2012
         Height          =   615
         Left            =   5400
         Picture         =   "HISTORIACLINICA.frx":2454
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "EVOLUCION"
         Height          =   1815
         Index           =   9
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   42
         Top             =   480
         Width           =   5175
      End
      Begin VB.CommandButton buscie10 
         DragIcon        =   "HISTORIACLINICA.frx":2ABE
         Height          =   375
         Left            =   1920
         Picture         =   "HISTORIACLINICA.frx":2F00
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   120
         Width           =   495
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "MEDICAMENTOS"
         Height          =   1575
         Index           =   8
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   40
         Top             =   2760
         Width           =   5175
      End
      Begin VB.CommandButton busmed 
         DragIcon        =   "HISTORIACLINICA.frx":3342
         Height          =   375
         Left            =   1920
         Picture         =   "HISTORIACLINICA.frx":3784
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "PROCEDIMIENTOS"
         Height          =   2775
         Index           =   12
         Left            =   5880
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   38
         Top             =   480
         Width           =   4575
      End
      Begin VB.CommandButton buspro 
         DragIcon        =   "HISTORIACLINICA.frx":3BC6
         Height          =   375
         Left            =   7680
         Picture         =   "HISTORIACLINICA.frx":4008
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblLabels 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Diagnostico"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Index           =   11
         Left            =   360
         TabIndex        =   45
         Top             =   120
         Width           =   1185
      End
      Begin VB.Label lblLabels 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Medicamentos"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Index           =   8
         Left            =   240
         TabIndex        =   44
         Top             =   2400
         Width           =   1485
      End
      Begin VB.Label lblLabels 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Procedimientos"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Index           =   14
         Left            =   5880
         TabIndex        =   43
         Top             =   120
         Width           =   1530
      End
   End
   Begin VB.Frame frmenfermeria 
      Height          =   4815
      Left            =   120
      TabIndex        =   32
      Top             =   2040
      Width           =   11175
      Begin MSDataGridLib.DataGrid grdsigvit 
         Height          =   3375
         Left            =   5640
         Negotiate       =   -1  'True
         TabIndex        =   46
         Top             =   600
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   5953
         _Version        =   393216
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         AllowDelete     =   -1  'True
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
      Begin VB.CommandButton busenferm 
         DragIcon        =   "HISTORIACLINICA.frx":444A
         Height          =   375
         Left            =   360
         Picture         =   "HISTORIACLINICA.frx":488C
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   120
         Width           =   495
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "ANOTACION"
         Height          =   3375
         Index           =   16
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   33
         Top             =   600
         Width           =   5175
      End
      Begin VB.Label lblLabels 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha/Hora/Medicacion/Notas"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Index           =   17
         Left            =   960
         TabIndex        =   35
         Top             =   240
         Width           =   3060
      End
   End
   Begin VB.Frame frmhistocli 
      Height          =   6555
      Left            =   120
      TabIndex        =   11
      Top             =   360
      Width           =   11265
      Begin VB.CommandButton cmDbuscie10 
         DragIcon        =   "HISTORIACLINICA.frx":4CCE
         Height          =   375
         Left            =   3480
         Picture         =   "HISTORIACLINICA.frx":5110
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "INFTECCIR"
         Height          =   885
         Index           =   0
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   2040
         Width           =   5175
      End
      Begin MSDataGridLib.DataGrid GRDHISTIPO 
         Height          =   2775
         Left            =   120
         Negotiate       =   -1  'True
         TabIndex        =   29
         Top             =   3480
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   4895
         _Version        =   393216
         BorderStyle     =   0
         ColumnHeaders   =   0   'False
         ForeColor       =   8388608
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   1
         WrapCellPointer =   -1  'True
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
      Begin MSDataGridLib.DataGrid GRDHISTOITEM 
         CausesValidation=   0   'False
         Height          =   2775
         Left            =   3720
         Negotiate       =   -1  'True
         TabIndex        =   12
         Top             =   3480
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   4895
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         BorderStyle     =   0
         Enabled         =   -1  'True
         ColumnHeaders   =   0   'False
         ForeColor       =   8388608
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   1
         WrapCellPointer =   -1  'True
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         Appearance      =   0  'Flat
         DataField       =   "HISTORIA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2805
         Index           =   15
         Left            =   7200
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   26
         Top             =   3480
         Width           =   3855
      End
      Begin VB.CommandButton busINFTECNICA 
         DragIcon        =   "HISTORIACLINICA.frx":5552
         Height          =   315
         Left            =   2280
         Picture         =   "HISTORIACLINICA.frx":5994
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "FICHACLINICA"
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   1455
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
         Index           =   2
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "NROMOV"
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
         Index           =   3
         Left            =   6720
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "COD_DOCTOR"
         Height          =   285
         Index           =   7
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1365
         Width           =   825
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "MOTIVOINGRESO"
         Height          =   525
         Index           =   10
         Left            =   6000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1080
         Width           =   4215
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "NRODOC"
         Height          =   285
         Index           =   5
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   600
         Width           =   1245
      End
      Begin VB.TextBox TXTFIELDS 
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
         Index           =   6
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "TIPDOC"
         Height          =   285
         Index           =   4
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   600
         Width           =   675
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "DATOS_PROPORCIONADOS_POR"
         Height          =   525
         Index           =   11
         Left            =   5760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   2520
         Width           =   4215
      End
      Begin VB.CommandButton busitem_histocli 
         DragIcon        =   "HISTORIACLINICA.frx":5DD6
         Height          =   375
         Left            =   120
         Picture         =   "HISTORIACLINICA.frx":6218
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label lblLabels 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CIE 10"
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
         Index           =   10
         Left            =   2880
         TabIndex        =   48
         Top             =   1800
         Width           =   510
      End
      Begin VB.Label lblLabels 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "INFORMACION TECNICA"
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
         Left            =   120
         TabIndex        =   31
         Top             =   1800
         Width           =   2070
      End
      Begin VB.Label LBLHISTOCLI 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00D2BBA0&
         Caption         =   "ANAMNESIS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Index           =   3
         Left            =   3720
         TabIndex        =   28
         Top             =   3120
         Width           =   1440
      End
      Begin VB.Label LBLHISTOCLI 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00D2BBA0&
         Caption         =   "HISTORIA CLINICA"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Index           =   0
         Left            =   7755
         TabIndex        =   27
         Top             =   3120
         Width           =   2265
      End
      Begin VB.Label lblLabels 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ficha Clinica Nro"
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
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1380
      End
      Begin VB.Label lblLabels 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha y Numero de Orden"
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
         Left            =   3240
         TabIndex        =   24
         Top             =   255
         Width           =   2040
      End
      Begin VB.Label LBLHISTOCLI 
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
         Left            =   1920
         TabIndex        =   23
         Top             =   960
         Width           =   1635
      End
      Begin VB.Label lblLabels 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Medico Tratante "
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
         Left            =   120
         TabIndex        =   22
         Top             =   1380
         Width           =   1305
      End
      Begin VB.Label LBLHISTOCLI 
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
         Left            =   2520
         TabIndex        =   21
         Top             =   1365
         Width           =   1245
      End
      Begin VB.Label lblLabels 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Numero de Documento"
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
         Left            =   2520
         TabIndex        =   20
         Top             =   600
         Width           =   1785
      End
      Begin VB.Label lblLabels 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Nacimiento"
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
         Left            =   6000
         TabIndex        =   19
         Top             =   660
         Width           =   1620
      End
      Begin VB.Label lblLabels 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo de Ingreso"
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
         Left            =   6000
         TabIndex        =   18
         Top             =   840
         Width           =   1380
      End
      Begin VB.Label lblLabels 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Documento"
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
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   1485
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
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   1620
      End
      Begin VB.Label lblLabels 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Datos Proporcionados por"
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
         Index           =   4
         Left            =   5760
         TabIndex        =   15
         Top             =   2280
         Width           =   1980
      End
      Begin VB.Label lblLabels 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Elegir Especialidad"
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
         Left            =   720
         TabIndex        =   14
         Top             =   3240
         Width           =   1470
      End
   End
   Begin MSComctlLib.TabStrip Tab_histocli 
      Height          =   7050
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   12435
      MultiRow        =   -1  'True
      MultiSelect     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "HISTORIA CLINICA"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "HOJA DE EVOLUCION"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ANOTACIONES DE ENFERMERIA"
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
Attribute VB_Name = "ATENCIONMEDICA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents adoHISTOCLI    As ADODB.Recordset
Attribute adoHISTOCLI.VB_VarHelpID = -1
Dim WithEvents adoHistoCliDet As ADODB.Recordset
Attribute adoHistoCliDet.VB_VarHelpID = -1
Dim WithEvents adoHistotipo   As ADODB.Recordset
Attribute adoHistotipo.VB_VarHelpID = -1
Dim WithEvents adoHistoitem   As ADODB.Recordset
Attribute adoHistoitem.VB_VarHelpID = -1
Dim WithEvents adoHistodoc    As ADODB.Recordset
Attribute adoHistodoc.VB_VarHelpID = -1
Dim WithEvents adoHistoper    As ADODB.Recordset
Attribute adoHistoper.VB_VarHelpID = -1
Dim WithEvents adoHistosig    As ADODB.Recordset
Attribute adoHistosig.VB_VarHelpID = -1
Dim WithEvents adoesp         As ADODB.Recordset
Attribute adoesp.VB_VarHelpID = -1
Dim VALIDAR                   As ValTecla
Dim WithEvents CtrVisualizar  As clsValidar
Attribute CtrVisualizar.VB_VarHelpID = -1
Dim WFECPED                   As Date
Dim WNROORD                   As Long
Dim WFICHACLI                 As Long
Dim tipdocu                   As String
Dim nrodocu                   As Long
Dim cod_especialidad          As Integer
Dim NOM_ESPECIALIDAD          As String
Dim AÑOS                      As Date
Dim AÑOS1                     As Date
Dim AÑOS2                     As Long
Public Sub imprimir()
        With DataGaleno.rsHISTORIACLINICA_Grouping
           If .State = adStateClosed Then
                    .Open 'colocar este codigo a todos los informes
                    .Filter = ""
                    .Requery
                    .MoveFirst
                    .Filter = "FICHACLINICA" & " = " & (TXTFIELDS(1).Text)
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
           rpt_historiaclinica.Refresh
           rpt_historiaclinica.Show
           .Filter = ""
           If .State = adStateOpen Then 'colocar este codigo a todos los informes
              .Close
              strSQL = ""
           End If
        End With
End Sub

Private Sub CARGAR_item(cod_esp As Integer)
On Error GoTo cargarerr
Dim ADOitem As ADODB.Recordset
Set ADOitem = New ADODB.Recordset
ADOitem.Open "select COD_ESPECIALIDAD,TIPO from ITEM_HISTORIACLINICA_ESPECIALI where cod_especialidad=" & cod_esp & "group by cod_especialidad,tipo", Cnn
Do While Not ADOitem.EOF
    If Not adoHistoCliDet.EOF Then
        adoHistoCliDet.MoveFirst
        adoHistoCliDet.Find "cod_especialidad=" & ADOitem("COD_ESPECIALIDAD").Value
        adoHistoCliDet.Find "TIPO='" & ADOitem("TIPO").Value & "'"
    End If
    If adoHistoCliDet.EOF Then
       adoHistoCliDet.AddNew
       adoHistoCliDet("cod_especialidad").Value = ADOitem("COD_ESPECIALIDAD").Value
       adoHistoCliDet("TIPO").Value = ADOitem("TIPO").Value
    End If
    ADOitem.MoveNext
Loop
ADOitem.Close
Exit Sub
cargarerr:
MsgBox cargarerr.Description
End Sub
Private Sub AGREGAR_Click()
  On Error GoTo AddErr
  Cnn.BeginTrans
  Select Case Me.Tab_histocli.SelectedItem.Index
  Case 1
          Call BLOQUEO(False)
          With adoHISTOCLI
              cod_especialidad = 0
              .AddNew
          End With
          
          LBLHISTOCLI(1).Caption = ""
          LBLHISTOCLI(2).Caption = ""
                   
          'TXTFIELDS(3) = OPTENER_CODIGO("SELECT NROMOV FROM ORDENSERVICIOCABECERA ORDER BY NROMOV DESC")
          'buscahistocli.SetFocus
   
          Tab_histocli_Click
        Exit Sub
  End Select
AddErr:
    MsgBox Err.Description
End Sub
Private Sub ACTUALIZAR_Click()
  On Error GoTo RefreshErr
  adoHISTOCLI.Requery
  If adoHISTOCLI.EOF = False And adoHISTOCLI.BOF = False Then
     Set adoHistoCliDet = adoHISTOCLI("HistoCliDet").UnderlyingValue
  End If
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub
Private Sub ANTERIOR_Click()
  On Error GoTo GoPrevError
  If Not adoHISTOCLI.BOF Then adoHISTOCLI.MovePrevious
  If adoHISTOCLI.BOF And adoHISTOCLI.RecordCount > 0 Then
    Beep
    adoHISTOCLI.MoveFirst
  End If
  Exit Sub
GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub adoHISTOCLI_RecordChangeComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
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

Private Sub buscahistocli_Click()
 Dim frmAyuda As New frmConsult
 Set frmAyuda.ActiveConection = Cnn
 frmAyuda.Modal = True
 frmAyuda.Consulta = "SELECT FICHACLINICA.FICHACLINICA,FICHACLINICA.TIPDOC,FICHACLINICA.NRODOC ,PERSONAS.NOMBRES,PERSONAS.APELLIDOS,PERSONAS.FECNAC FROM FICHACLINICA,PERSONAS WHERE FICHACLINICA.TIPDOC=PERSONAS.TIPDOC AND FICHACLINICA.NRODOC=PERSONAS.NRODOC"
 'frmAyuda.Consulta = "SELECT FICHACLINICA,PERSONAS.TIPDOC,PERSONAS.NRODOC,PERSONAS.NOMBRES,PERSONAS.APELLIDOS FROM FICHACLINICA,PERSONAS WHERE FICHACLINICA.TIPDOC=PERSONAS.TIPDOC AND FICHACLINICA.NRODOC=PERSONAS.NRODOC"
 frmAyuda.Show vbModal
 If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
    With frmAyuda.msGrid
        TXTFIELDS(1).Text = .Columns(0).Text
        TXTFIELDS(4).Text = .Columns(1).Text
        TXTFIELDS(5).Text = .Columns(2).Text
        AÑOS = Date
        AÑOS1 = .Columns(5).Text
        TXTFIELDS(6).Text = .Columns(5).Text
        LBLHISTOCLI(1).Caption = .Columns(4).Text & ", " & .Columns(3).Text
    End With
 End If
frmAyuda.Cerrar
End Sub

Private Sub buscie10_Click()
 Dim frmAyuda As New frmConsult
 Set frmAyuda.ActiveConection = Cnn
 frmAyuda.Modal = True
 frmAyuda.Consulta = "SELECT COD_ENF AS COD,DESCRIPCION FROM CIE10_4 "
 frmAyuda.Show vbModal
 If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("MODIFICAR").Enabled = False Then
    With frmAyuda.msGrid
            If TXTFIELDS(9) = "" Then
                TXTFIELDS(9).Text = TXTFIELDS(9).Text + Str(Day(Date)) + Str(Month(Date)) + Str(Year(Date)) + Str(Hour(Time())) + ":" + Str(Minute(Time())) + "   " + .Columns(1).Text
            Else
                TXTFIELDS(9).SetFocus
                For i = 1 To Len(TXTFIELDS(9)) / 2
                    SendKeys "{DOWN}", True
                Next
                SendKeys "{END}", True
                SendKeys "{ENTER}", True
                TXTFIELDS(9).SetFocus
                TXTFIELDS(9).Text = TXTFIELDS(9).Text + Str(Day(Date)) + Str(Month(Date)) + Str(Year(Date)) + Str(Hour(Time())) + ":" + Str(Minute(Time())) + "   " + .Columns(1).Text
            End If
   End With
 End If
frmAyuda.Cerrar
End Sub

Private Sub busdoctor_Click()

End Sub

Private Sub busenferm_Click()
 Dim frmAyuda As New frmConsult
 Set frmAyuda.ActiveConection = Cnn
 frmAyuda.Modal = True
 frmAyuda.Consulta = "SELECT NOM_MEDICAMENTO,PRESENTACION,VIA,DOSIS,HORARIO FROM MEDICAMENTOS"
 frmAyuda.Show vbModal
 If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("MODIFICAR").Enabled = False Then
    With frmAyuda.msGrid
            If TXTFIELDS(16) = "" Then
                TXTFIELDS(16).Text = TXTFIELDS(16).Text + Str(Day(Date)) + Str(Month(Date)) + Str(Year(Date)) + Str(Hour(Time())) + ":" + Str(Minute(Time())) + "   " + .Columns(1).Text
            Else
                TXTFIELDS(16).SetFocus
                For i = 1 To Len(TXTFIELDS(16)) / 2
                    SendKeys "{DOWN}", True
                Next
                SendKeys "{END}", True
                SendKeys "{ENTER}", True
                TXTFIELDS(16).SetFocus
                TXTFIELDS(16).Text = TXTFIELDS(16).Text + Str(Day(Date)) + Str(Month(Date)) + Str(Year(Date)) + Str(Hour(Time())) + ":" + Str(Minute(Time())) + "   " + .Columns(1).Text
            End If
   End With
 End If
frmAyuda.Cerrar
End Sub

Private Sub busfecha_Click()
   Dim frmAyuda As New frmConsult
   Set frmAyuda.ActiveConection = Cnn
   frmAyuda.Modal = True
   frmAyuda.Consulta = "SELECT FICHACLINICA,FEC_PED,NROMOV FROM ORDENSERVICIOCABECERA WHERE ORDENSERVICIOCABECERA.FICHACLINICA=" & TXTFIELDS(1).Text & " AND NOT EXISTS (SELECT FEC_PED,NROMOV FROM HISTORIACLINICACABECERA WHERE ORDENSERVICIOCABECERA.FEC_PED=HISTORIACLINICACABECERA.FEC_PED AND ORDENSERVICIOCABECERA.NROMOV=HISTORIACLINICACABECERA.NROMOV)"
  'frmAyuda.Consulta = "SELECT FICHACLINICA,PERSONAS.TIPDOC,PERSONAS.NRODOC,PERSONAS.NOMBRES,PERSONAS.APELLIDOS FROM FICHACLINICA,PERSONAS WHERE FICHACLINICA.TIPDOC=PERSONAS.TIPDOC AND FICHACLINICA.NRODOC=PERSONAS.NRODOC"
   frmAyuda.Show vbModal
   If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
     With frmAyuda.msGrid
        TXTFIELDS(1).Text = .Columns(0).Text
        TXTFIELDS(2).Text = .Columns(1).Text
        TXTFIELDS(3).Text = .Columns(2).Text
        'LBLORDEN(1).Caption = .Columns(4).Text & ", " & .Columns(3).Text
     End With
  End If
  frmAyuda.Cerrar
End Sub
Private Sub SIGUIENTE_Click()
  On Error GoTo GoNextError
  If Not adoHISTOCLI.EOF Then adoHISTOCLI.MoveNext
  If adoHISTOCLI.EOF And adoHISTOCLI.RecordCount > 0 Then
    Beep
    adoHISTOCLI.MoveLast
  End If
  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub
Private Sub PRIMERO_Click()
  On Error GoTo GoFirstError
  adoHISTOCLI.MoveFirst
  Exit Sub
GoFirstError:
  MsgBox Err.Description
End Sub
Private Sub ULTIMO_Click()
  On Error GoTo GoLastError
  adoHISTOCLI.MoveLast
  Exit Sub
GoLastError:
  MsgBox Err.Description
End Sub
Private Sub VOLVER_Click()
  cmdCancel_Click
  Unload Me
End Sub
Private Sub GRABAR_Click()
On Error GoTo UpdateErr
Select Case Me.Tab_histocli.SelectedItem.Index
Case 1, 2, 3
         SendKeys "{TAB}", True
         adoHISTOCLI.UpdateBatch adAffectAllChapters
         adoHistoCliDet.UpdateBatch adAffectAllChapters
         i = False
         Call BLOQUEO(True)
         Cnn.Execute "DELETE FROM HISTORIACLINICADETALLE WHERE HISTORIA LIKE ''"
         Cnn.CommitTrans
         Exit Sub
End Select
UpdateErr:
    MsgBox Err.Description
End Sub
Private Sub MODIFICAR_Click()
On Error GoTo EditErr
Select Case Me.Tab_histocli.SelectedItem.Index
Case 1, 2
        Call BLOQUEO(False)
        ATENCIONMEDICA.GRDHISTOITEM.AllowUpdate = True
        ATENCIONMEDICA.GRDHISTOITEM.AllowDelete = True
        Cnn.BeginTrans
Case 3
        Call BLOQUEO(False)
        Cnn.BeginTrans
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
    Select Case Me.Tab_histocli.SelectedItem.Index
    Case 1
        With adoHISTOCLI
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

Private Sub busINFTECNICA_Click()
 Dim frmAyuda As New frmConsult
 Set frmAyuda.ActiveConection = Cnn
 frmAyuda.Modal = True
 frmAyuda.Consulta = "SELECT * FROM PARAMETROS WHERE TIPO='CIRUGIA'"
 frmAyuda.msGrid.Columns(0).Visible = False
 frmAyuda.Show vbModal
 If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("MODIFICAR").Enabled = False Then
    With frmAyuda.msGrid
            If TXTFIELDS(0) = "" Then
                TXTFIELDS(0).Text = TXTFIELDS(0).Text + .Columns(2).Text
            Else
                TXTFIELDS(0).SetFocus
                For i = 1 To Len(TXTFIELDS(0)) / 2
                    SendKeys "{DOWN}", True
                Next
                SendKeys "{END}", True
                SendKeys "{ENTER}", True
                TXTFIELDS(0).SetFocus
                TXTFIELDS(0).Text = TXTFIELDS(0).Text + .Columns(2).Text
            End If
   End With
 End If
frmAyuda.Cerrar
End Sub

Private Sub busitem_histocli_Click()
If MDIGALENO.tbToolBar.Buttons("MODIFICAR").Enabled = False Then
    Dim frmAyuda As New frmConsult
    Set frmAyuda.ActiveConection = Cnn
    adoHISTOCLI.UpdateBatch
    frmAyuda.Modal = True
    frmAyuda.Consulta = "select * from especialidad"
    frmAyuda.Show vbModal
    If Not frmAyuda.Descargado Then
        With frmAyuda.msGrid
            Set adoHistoCliDet = adoHISTOCLI("HistoCliDet").UnderlyingValue
            adoHistoCliDet.Find "COD_ESPECIALIDAD= " & .Columns(0).Text
            cod_especialidad = .Columns(0).Value
            LBLHISTOCLI(3).Caption = "ANAMNESIS" & " " & .Columns(1).Text
            CARGAR_item (.Columns(0).Value)
            Set TXTFIELDS(15).DataSource = adoHistoCliDet
            Tab_histocli_Click
        End With
    End If
frmAyuda.Cerrar
End If
End Sub

Private Sub busmed_Click()
 Dim frmAyuda As New frmConsult
 Set frmAyuda.ActiveConection = Cnn
 frmAyuda.Modal = True
 frmAyuda.Consulta = "SELECT NOM_MEDICAMENTO,PRESENTACION,VIA,DOSIS,HORARIO FROM MEDICAMENTOS"
 frmAyuda.Show vbModal
 If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("MODIFICAR").Enabled = False Then
    With frmAyuda.msGrid
            If TXTFIELDS(8) = "" Then
                TXTFIELDS(8).Text = TXTFIELDS(8).Text + Str(Day(Date)) + Str(Month(Date)) + Str(Year(Date)) + Str(Hour(Time())) + ":" + Str(Minute(Time())) + "   " + .Columns(1).Text
            Else
                TXTFIELDS(8).SetFocus
                For i = 1 To Len(TXTFIELDS(12)) / 2
                    SendKeys "{DOWN}", True
                Next
                SendKeys "{END}", True
                SendKeys "{ENTER}", True
                TXTFIELDS(8).SetFocus
                TXTFIELDS(8).Text = TXTFIELDS(8).Text + Str(Day(Date)) + Str(Month(Date)) + Str(Year(Date)) + Str(Hour(Time())) + ":" + Str(Minute(Time())) + "   " + .Columns(1).Text
            End If
   End With
 End If
frmAyuda.Cerrar
End Sub

Private Sub buspro_Click()
 Dim frmAyuda As New frmConsult
 Set frmAyuda.ActiveConection = Cnn
 frmAyuda.Modal = True
 frmAyuda.Consulta = "SELECT COD_PROCEDIMIENTO AS COD , NOM_PROCEDIMIENTO FROM PROCEDIMIENTOS"
 frmAyuda.Show vbModal
 If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("MODIFICAR").Enabled = False Then
    With frmAyuda.msGrid
            If TXTFIELDS(12) = "" Then
                TXTFIELDS(12).Text = TXTFIELDS(12).Text + Str(Day(Date)) + Str(Month(Date)) + Str(Year(Date)) + Str(Hour(Time())) + ":" + Str(Minute(Time())) + "   " + .Columns(1).Text
            Else
                TXTFIELDS(12).SetFocus
                For i = 1 To Len(TXTFIELDS(12)) / 2
                    SendKeys "{DOWN}", True
                Next
                SendKeys "{END}", True
                SendKeys "{ENTER}", True
                TXTFIELDS(12).SetFocus
                TXTFIELDS(12).Text = TXTFIELDS(12).Text + Str(Day(Date)) + Str(Month(Date)) + Str(Year(Date)) + Str(Hour(Time())) + ":" + Str(Minute(Time())) + "   " + .Columns(1).Text
            End If
   End With
 End If
frmAyuda.Cerrar
End Sub

Private Sub cmdbusmed_Click()
 Dim frmAyuda As New frmConsult
 Set frmAyuda.ActiveConection = Cnn
 frmAyuda.Modal = True
 frmAyuda.Consulta = "SELECT DOCTOR.CODDR,DOCTOR.REGPRF,PERSONAS.TIPDOC +', '+ STR(PERSONAS.NRODOC) as IDENTIDAD, PERSONAS.APELLIDOS+', '+PERSONAS.NOMBRES AS NOMDOC FROM DOCTOR  LEFT OUTER JOIN PERSONAS ON DOCTOR.TIPDOC = PERSONAS.TIPDOC AND DOCTOR.NRODOC = PERSONAS.NRODOC"
 frmAyuda.Show vbModal
 If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("MODIFICAR").Enabled = False Then
    With frmAyuda.msGrid
            If TXTFIELDS(9) = "" Then
                TXTFIELDS(9).Text = TXTFIELDS(9).Text + "Dr. " + .Columns(3).Text + " Registro Nro. " + .Columns(1).Text
            Else
                TXTFIELDS(9).SetFocus
                For i = 1 To Len(TXTFIELDS(9)) / 2
                    SendKeys "{DOWN}", True
                Next
                SendKeys "{END}", True
                SendKeys "{ENTER}", True
                TXTFIELDS(9).SetFocus
                TXTFIELDS(9).Text = TXTFIELDS(9).Text + "Dr. " + .Columns(3).Text + " Registro Nro. " + .Columns(1).Text
            End If
   End With
 End If
frmAyuda.Cerrar
End Sub

Private Sub cmd_imp_diag_Click()
If MDIGALENO.tbToolBar.Buttons("MODIFICAR").Enabled = True Then
    With DataGaleno.rstexto
         If .State = adStateClosed Then
                  Cnn.Execute "INSERT INTO texto (TEXTO,FICHACLINICA,TIPDOC,NRODOC,NOMPACIENTE) VALUES ( '" & TXTFIELDS(9).SelText & "', " & TXTFIELDS(1).Text & ", '" & TXTFIELDS(4).Text & "' , " & TXTFIELDS(5).Text & " , '" & Trim(LBLHISTOCLI(1).Caption) & "' )"
                  .Open 'colocar este codigo a todos los informes
                  .Requery
        End If
        rpt_diagnostico.Show
        rpt_diagnostico.Refresh
        If .State = adStateOpen Then 'colocar este codigo a todos los informes
                  .Close
                  Cnn.Execute "DELETE FROM texto"
        End If
    End With
End If
End Sub

Private Sub cmd_imp_med_Click()
If MDIGALENO.tbToolBar.Buttons("MODIFICAR").Enabled = True Then
    With DataGaleno.rstexto
         If .State = adStateClosed Then
                  Cnn.Execute "INSERT INTO texto (TEXTO,FICHACLINICA,TIPDOC,NRODOC,NOMPACIENTE) VALUES ( '" & TXTFIELDS(8).SelText & "', " & TXTFIELDS(1).Text & ", '" & TXTFIELDS(4).Text & "' , " & TXTFIELDS(5).Text & " , '" & Trim(LBLHISTOCLI(1).Caption) & "' )"
                  .Open 'colocar este codigo a todos los informes
                  .Requery
        End If
        rpt_indicaciones.Show
        rpt_indicaciones.Refresh
        If .State = adStateOpen Then 'colocar este codigo a todos los informes
                  .Close
                  Cnn.Execute "DELETE FROM texto"
        End If
    End With
End If
End Sub

Private Sub cmd_imp_pro_Click()
If MDIGALENO.tbToolBar.Buttons("MODIFICAR").Enabled = True Then
    With DataGaleno.rstexto
         If .State = adStateClosed Then
                  Cnn.Execute "INSERT INTO texto (TEXTO,FICHACLINICA,TIPDOC,NRODOC,NOMPACIENTE) VALUES ( '" & TXTFIELDS(12).SelText & "', " & TXTFIELDS(1).Text & ", '" & TXTFIELDS(4).Text & "' , " & TXTFIELDS(5).Text & " , '" & Trim(LBLHISTOCLI(1).Caption) & "' )"
                  .Open 'colocar este codigo a todos los informes
                  .Requery
        End If
        rpt_procedimientos.Show
        rpt_procedimientos.Refresh
        If .State = adStateOpen Then 'colocar este codigo a todos los informes
                  .Close
                  Cnn.Execute "DELETE FROM texto"
        End If
    End With
End If
End Sub

Private Sub cmDbuscie10_Click()
 Dim frmAyuda As New frmConsult
 Set frmAyuda.ActiveConection = Cnn
 frmAyuda.Modal = True
 frmAyuda.Consulta = "SELECT COD_ENF AS COD,DESCRIPCION FROM CIE10_4 "
 frmAyuda.Show vbModal
 If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("MODIFICAR").Enabled = False Then
    With frmAyuda.msGrid
            If TXTFIELDS(0) = "" Then
                TXTFIELDS(0).Text = TXTFIELDS(0).Text + .Columns(1).Text
            Else
                TXTFIELDS(0).SetFocus
                For i = 1 To Len(TXTFIELDS(0)) / 2
                    SendKeys "{DOWN}", True
                Next
                SendKeys "{END}", True
                SendKeys "{ENTER}", True
                TXTFIELDS(0).SetFocus
                TXTFIELDS(0).Text = TXTFIELDS(0).Text + .Columns(1).Text
            End If
   End With
 End If
frmAyuda.Cerrar

End Sub

Private Sub Form_Load()
    Set VALIDAR = New ValTecla
    Me.Height = 7500
    Me.Height = 7500
    Me.Left = 80
    Me.Top = 0
    Me.Width = 11800
    Set adoHISTOCLI = New ADODB.Recordset
    Set adoHistoCliDet = New ADODB.Recordset
    Set adoHistotipo = New ADODB.Recordset
    Set adoHistoitem = New ADODB.Recordset
    Set adoHistodoc = New ADODB.Recordset
    Set adoHistoper = New ADODB.Recordset
    Set adoHistosig = New ADODB.Recordset
    Set adoesp = New ADODB.Recordset
    adoHistoper.Open "SELECT TIPDOC,NRODOC,NOMBRES,APELLIDOS,FECNAC FROM PERSONAS", Cnn
    adoHistodoc.Open "select CODDR,DOCTOR.TIPDOC,DOCTOR.TIPDOC,PERSONAS.APELLIDOS+' ,'+PERSONAS.NOMBRES AS NOMDOC from doctor,PERSONAS WHERE DOCTOR.TIPDOC=PERSONAS.TIPDOC AND DOCTOR.NRODOC=PERSONAS.NRODOC", Cnn
    adoHISTOCLI.Open "SHAPE {SELECT * FROM HISTORIACLINICACABECERA}  AS HistoCliCab APPEND ({SELECT * FROM HISTORIACLINICADETALLE}  AS HistoCliDet RELATE 'FICHACLINICA' TO 'FICHACLINICA','FEC_PED' TO 'FEC_PED','NROMOV' TO 'NROMOV','TIPDOC' TO 'TIPDOC','NRODOC' TO 'NRODOC') AS HistoCliDet", Cnn, adOpenStatic, adLockOptimistic
    adoHistosig.Open "select * from signosvitales", Cnn, adOpenStatic, adLockOptimistic
    adoesp.Open "select * from especialidad", Cnn
    If adoHistoCliDet.State = 0 Then
      If adoHISTOCLI.EOF = False And adoHISTOCLI.BOF = False Then
        Set adoHistoCliDet = adoHISTOCLI("HistoCliDet").UnderlyingValue
      End If
    End If
    Call BLOQUEO(True)
    Call Tab_histocli_Click
End Sub
Public Sub PASARIDENTIDAD(FEC_PED As Date, NROMOV As Long, FICHACLINICA As Long, tipdoc As String, nrodoc As Long)
    WFECPED = FEC_PED
    WNROORD = NROMOV
    WFICHACLI = FICHACLINICA
    tipdocu = tipdoc
    nrodocu = nrodoc
           With adoHISTOCLI
                  On Error Resume Next
                  .MoveFirst
                  .Find "fec_ped=" & WFECPED
                  .Find "nromov=" & WNROORD
                   Dim oText As TextBox
                   For Each oText In Me.TXTFIELDS
                    If oText.Index < 40 Then
                       Set oText.DataSource = adoHISTOCLI
                    End If
                   Next


                  If .EOF Then
                        MDIGALENO.fncBotones False, "ATENCIONMEDICA"
                        AGREGAR_Click
                        TXTFIELDS(1).Text = WFICHACLI
                        TXTFIELDS(2).Text = WFECPED
                        TXTFIELDS(3).Text = WNROORD
                        TXTFIELDS(4).Text = tipdocu
                        TXTFIELDS(5).Text = nrodocu
                        'TXTFIELDS(7).Text = Val(Mid(frmAyuda.msGrid.Columns(5).Text, 1))
                        'LBLHISTOCLI(1).Caption = frmAyuda.msGrid.Columns(2).Text

                  End If
                  Call txtFields_Change(1)
                  TXTFIELDS(6).Text = frmAyuda.msGrid.Columns(6).Text
           End With
        frmAyuda.Cerrar
        Me.frmhistocli.Visible = True
    
End Sub

Private Sub grdhistipo_Click()
 On Error Resume Next
  'If MDIGALENO.tbToolBar.Buttons("MODIFICAR").Enabled = False Then
    With GRDHISTIPO
                adoHistoitem.Filter = ""
                adoHistoitem.MoveFirst
                adoHistoitem.Filter = "TIPO='" & Trim$(.Columns(0).Text) & "'"
                adoHistoCliDet.MoveFirst
                adoHistoCliDet.Find "TIPO='" & Me.GRDHISTIPO.Columns("TIPO").Text & "'"
    End With
  'End If
End Sub

Private Sub GRDHISTOITEM_DblClick()
 Dim frmAyuda As New frmConsult
 Set frmAyuda.ActiveConection = Cnn
 frmAyuda.Modal = True
 frmAyuda.Consulta = "SELECT VALOR FROM VALORES WHERE cod_especialidad=" & cod_especialidad & "and TIPO LIKE '" & Me.GRDHISTIPO.Columns("TIPO").Text & "'" & "and cod_item=" & Me.GRDHISTOITEM.Columns("ITEM").Value
 W_COUNT = OPTENER_CODIGO("SELECT COUNT(VALOR) FROM VALORES WHERE cod_especialidad=" & cod_especialidad & "and TIPO LIKE '" & Me.GRDHISTIPO.Columns("TIPO").Text & "'" & "and cod_item=" & Me.GRDHISTOITEM.Columns("ITEM").Value)
 If Not W_COUNT < 2 Then
    frmAyuda.Show vbModal
    If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("MODIFICAR").Enabled = False Then
        With frmAyuda
            adoHistoCliDet.MoveFirst
            adoHistoCliDet.Find "TIPO='" & Me.GRDHISTIPO.Columns("TIPO").Text & "'"
            If TXTFIELDS(15) = "" Then
                TXTFIELDS(15).Text = GRDHISTOITEM.Columns(2).Text & " " & .Seleccion
            Else
                TXTFIELDS(15).SetFocus
                For i = 1 To Len(TXTFIELDS(15)) / 2
                    SendKeys "{DOWN}", True
                Next
                SendKeys "{END}", True
                SendKeys "{ENTER}", True
                TXTFIELDS(15).SetFocus
                TXTFIELDS(15).Text = TXTFIELDS(15).Text + GRDHISTOITEM.Columns(2).Text & " " + .Seleccion
                
            End If
        End With
    End If
    frmAyuda.Cerrar
 Else
    If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("MODIFICAR").Enabled = False Then
        adoHistoCliDet.MoveFirst
        adoHistoCliDet.Find "TIPO='" & Me.GRDHISTIPO.Columns("TIPO").Text & "'"
        TXTFIELDS(15).SetFocus
        For i = 1 To Len(TXTFIELDS(15)) / 2
            SendKeys "{DOWN}", True
        Next
        SendKeys "{END}", True
        SendKeys "{ENTER}", True
        TXTFIELDS(15).SetFocus
        TXTFIELDS(15).Text = TXTFIELDS(15).Text + GRDHISTOITEM.Columns(2).Text
    End If
 End If
 TXTFIELDS(15).SetFocus
End Sub

Private Sub grdsigvit_DblClick()
If MDIGALENO.tbToolBar.Buttons("MODIFICAR").Enabled = False Then
    adoHistosig.AddNew
    adoHistosig.Fields(0).Value = TXTFIELDS(1).Text
    adoHistosig.Fields(1).Value = TXTFIELDS(2).Text
    adoHistosig.Fields(2).Value = TXTFIELDS(3).Text
    adoHistosig.Fields(3).Value = TXTFIELDS(4).Text
    adoHistosig.Fields(4).Value = TXTFIELDS(5).Text
    adoHistosig.Fields(5).Value = Date + Time
    Set grdsigvit.DataSource = adoHistosig
    grdsigvit.Columns(0).Visible = False
    grdsigvit.Columns(1).Visible = False
    grdsigvit.Columns(2).Visible = False
    grdsigvit.Columns(3).Visible = False
    grdsigvit.Columns(4).Visible = False
End If

End Sub

Private Sub Tab_histocli_Click()
  On Error Resume Next
  Select Case Me.Tab_histocli.SelectedItem.Index
  Case 1
                frmenfermeria.Visible = False
                frmevolucion.Visible = False
                Set TXTFIELDS(15).DataSource = adoHistoCliDet
                If adoHistoCliDet.State = 1 Then
                    If adoHISTOCLI.EOF = False And adoHISTOCLI.BOF = False Then
                        If adoHistotipo.State = 1 Then
                            adoHistotipo.Close
                        End If
                        adoHistotipo.Open "select TIPO from HISTORIACLINICADETALLE where FEC_PED= '" & TXTFIELDS(2).Text & "' AND NROMOV= '" & TXTFIELDS(3).Text & "' AND  cod_especialidad='" & Str(cod_especialidad) + "' group by tipo", Cnn
                        Set GRDHISTIPO.DataSource = adoHistotipo
                        If adoHistoitem.State = 1 Then
                            adoHistoitem.Close
                        End If
                        adoHistoitem.Open "select TIPO,ITEM,ITEM_DESCRIPCION from ITEM_HISTORIACLINICA_ESPECIALI where cod_especialidad='" & Str(cod_especialidad) + "'", Cnn
                        Set GRDHISTOITEM.DataSource = adoHistoitem
                        GRDHISTOITEM.Columns(0).Visible = False
                        GRDHISTOITEM.Columns(1).Visible = False
                        adoHistosig.Filter = adoHistosig.Fields(0).Name & "=" & TXTFIELDS(1).Text & _
                                           " AND " & adoHistosig.Fields(1).Name & "=" & TXTFIELDS(2).Text & _
                                           " AND " & adoHistosig.Fields(2).Name & "=" & TXTFIELDS(3).Text & _
                                           " AND " & adoHistosig.Fields(3).Name & "= '" & TXTFIELDS(4).Text & "'" & _
                                           " AND " & adoHistosig.Fields(4).Name & "=" & TXTFIELDS(5).Text
                        Set grdsigvit.DataSource = adoHistosig
                        grdsigvit.Columns(0).Visible = False
                        grdsigvit.Columns(1).Visible = False
                        grdsigvit.Columns(2).Visible = False
                        grdsigvit.Columns(3).Visible = False
                        grdsigvit.Columns(4).Visible = False
                        GRDHISTIPO.Columns("TIPO").Width = 5000
                        GRDHISTOITEM.Columns(2).Width = 5000
                        grdhistipo_Click
                   End If
                End If
  Case 2
                frmenfermeria.Visible = False
                frmevolucion.Visible = True
  Case 3
                frmenfermeria.Visible = True
                frmevolucion.Visible = False
                adoHistosig.Filter = ""
                adoHistosig.Filter = adoHistosig.Fields(0).Name & "=" & TXTFIELDS(1).Text & _
                                     " AND " & adoHistosig.Fields(1).Name & "=" & TXTFIELDS(2).Text & _
                                     " AND " & adoHistosig.Fields(2).Name & "=" & TXTFIELDS(3).Text & _
                                     " AND " & adoHistosig.Fields(3).Name & "= '" & TXTFIELDS(4).Text & "'" & _
                                     " AND " & adoHistosig.Fields(4).Name & "=" & TXTFIELDS(5).Text
                Set grdsigvit.DataSource = adoHistosig
                grdsigvit.Columns(0).Visible = False
                grdsigvit.Columns(1).Visible = False
                grdsigvit.Columns(2).Visible = False
                grdsigvit.Columns(3).Visible = False
                grdsigvit.Columns(4).Visible = False

  End Select

End Sub
Private Sub CANCELAR_Click()
Me.GRDHISTOITEM.AllowAddNew = False
Me.GRDHISTOITEM.AllowDelete = False
Me.GRDHISTOITEM.AllowUpdate = False
Call cmdCancel_Click
i = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
 On Error Resume Next
 Call cmdCancel_Click
 MDIGALENO.tbToolBar.Buttons("Add").Enabled = True
End Sub
Private Sub cmdCancel_Click()
  On Error Resume Next
  Cnn.RollbackTrans
  adoHISTOCLI.CancelUpdate
  MDIGALENO.fncBotones True, ""
  Unload Me
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
End Sub
Public Sub requeryReg()
ACTUALIZAR_Click
End Sub
Public Sub buscarReg()
BUSCAR_Click
End Sub
Public Sub CancelarReg()
CANCELAR_Click
End Sub

Private Sub txtFields_Change(Index As Integer)

Select Case Index
Case 1, 3
      If adoHistoCliDet.State = 1 Then
           If Not adoHistoCliDet.EOF And Not adoHistoCliDet.BOF Then
               cod_especialidad = adoHistoCliDet.Fields("COD_ESPECIALIDAD").Value
               adoHistodoc.MoveFirst
               adoHistodoc.Find "coddr=" & Val(TXTFIELDS(7).Text)
               If Not adoHistodoc.EOF And Not adoHistodoc.BOF Then
                    LBLHISTOCLI(2).Caption = adoHistodoc.Fields("nomdoc").Value
               End If
           Else
               cod_especialidad = 0
           End If
           adoesp.Filter = "cod_especialidad=" & Str(cod_especialidad)
           NOM_ESPECIALIDAD = IIf(adoesp.EOF = True, " ", adoesp.Fields(1))
           LBLHISTOCLI(3).Caption = "ANAMNESIS" & " " & NOM_ESPECIALIDAD
            If adoHISTOCLI.EOF = False And adoHISTOCLI.BOF = False Then
                If adoHistotipo.State = 1 Then
                    adoHistotipo.Close
                End If
                If adoHistoCliDet.EOF = False And adoHistoCliDet.BOF = False Then
                    adoHistotipo.Open "select TIPO from HISTORIACLINICADETALLE where FEC_PED= '" & adoHistoCliDet.Fields(1).Value & "' AND NROMOV= '" & adoHistoCliDet.Fields(2).Value & "' AND  cod_especialidad='" & Str(cod_especialidad) + "' group by tipo", Cnn
                    Set GRDHISTIPO.DataSource = adoHistotipo
                End If
                If adoHistoitem.State = 1 Then
                    adoHistoitem.Close
                End If
                adoHistoitem.Open "select TIPO,ITEM,ITEM_DESCRIPCION from ITEM_HISTORIACLINICA_ESPECIALI where cod_especialidad='" & Str(cod_especialidad) + "'", Cnn
                Set GRDHISTOITEM.DataSource = adoHistoitem
                GRDHISTOITEM.Columns(0).Visible = False
                GRDHISTOITEM.Columns(1).Visible = False
                grdhistipo_Click
           End If
        End If
 Case 5
        adoHistoper.MoveFirst
        adoHistoper.Find "TIPDOC='" & TXTFIELDS(4).Text & "'"
        adoHistoper.Find "NRODOC=" & TXTFIELDS(5).Text
        If Not adoHistoper.EOF And Not adoHistoper.BOF Then
               LBLHISTOCLI(1).Caption = adoHistoper.Fields("APELLIDOS").Value & ", " & adoHistoper.Fields("NOMBRES").Value
        End If
 Case 7
        adoHistodoc.MoveFirst
        adoHistodoc.Find "coddr=" & Val(TXTFIELDS(7).Text)
        If Not adoHistodoc.EOF And Not adoHistodoc.BOF Then
               LBLHISTOCLI(2).Caption = adoHistodoc.Fields("nomdoc").Value
        End If

End Select
End Sub
Private Sub BUSCAR_Click()
 Select Case Me.Tab_histocli.SelectedItem.Index
  Case 1
        On Error Resume Next
        Me.frmhistocli.Visible = False
        Dim frmAyuda As New frmConsult
        Set frmAyuda.ActiveConection = Cnn
        frmAyuda.Modal = True
        frmAyuda.Consulta = "select * from tipomov "
        frmAyuda.Show vbModal
        w_tipomov = frmAyuda.msGrid.Columns(0).Value
        frmAyuda.Modal = True
        frmAyuda.Consulta = "SELECT DISTINCT ORDENSERVICIOCABECERA.FICHACLINICA,PERSONAS.TIPDOC+' '+STR(PERSONAS.NRODOC) AS IDENTIDAD,RTRIM(PERSONAS.APELLIDOS)+' ,'+RTRIM(PERSONAS.NOMBRES) AS PACIENTES,ORDENSERVICIOCABECERA.FEC_PED,ORDENSERVICIOCABECERA.NROMOV,CODDR,FECNAC FROM ORDENSERVICIOCABECERA,ORDENSERVICIODETALLE ,PERSONAS WHERE TIPOMOV=" & "'" & w_tipomov & "'" & " AND PERSONAS.TIPDOC=ORDENSERVICIOCABECERA.TIPDOC AND  PERSONAS.NRODOC=ORDENSERVICIOCABECERA.NRODOC" & _
        " AND ORDENSERVICIOCABECERA.FICHACLINICA=ORDENSERVICIODETALLE.FICHACLINICA AND ORDENSERVICIOCABECERA.FEC_PED=ORDENSERVICIODETALLE.FEC_PED AND ORDENSERVICIOCABECERA.NROMOV=ORDENSERVICIODETALLE.NROMOV AND (ORDENSERVICIODETALLE.TIPOSER='CO' OR ORDENSERVICIODETALLE.TIPOSER='QUI' OR ORDENSERVICIODETALLE.TIPOSER='SAN')"
        frmAyuda.Show vbModal
        If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
        End If
         If Not frmAyuda.Descargado Then
           With adoHISTOCLI
                  On Error Resume Next
                  .MoveFirst
                  .Find "fec_ped=" & frmAyuda.msGrid.Columns(3).Value
                  .Find "nromov=" & frmAyuda.msGrid.Columns(4).Value
                   Dim oText As TextBox
                   For Each oText In Me.TXTFIELDS
                    If oText.Index < 40 Then
                       Set oText.DataSource = adoHISTOCLI
                    End If
                   Next
                   Set TXTFIELDS(15).DataSource = adoHistoCliDet

                  If .EOF Then
                        MDIGALENO.fncBotones False, "ATENCIONMEDICA"
                        AGREGAR_Click
                        TXTFIELDS(1).Text = frmAyuda.msGrid.Columns(0).Text
                        TXTFIELDS(2).Text = frmAyuda.msGrid.Columns(3).Text
                        TXTFIELDS(3).Text = frmAyuda.msGrid.Columns(4).Text
                        TXTFIELDS(4).Text = Mid(frmAyuda.msGrid.Columns(1).Text, 1, 3)
                        TXTFIELDS(5).Text = Val(Mid(frmAyuda.msGrid.Columns(1).Text, 4))
                        TXTFIELDS(7).Text = Val(Mid(frmAyuda.msGrid.Columns(5).Text, 1))
                        LBLHISTOCLI(1).Caption = frmAyuda.msGrid.Columns(2).Text

                  End If
                  Call txtFields_Change(1)
                  TXTFIELDS(6).Text = frmAyuda.msGrid.Columns(6).Text
           End With
         End If
        frmAyuda.Cerrar
        Me.frmhistocli.Visible = True


    End Select

End Sub
Private Sub BLOQUEO(X As Boolean)
    Dim oText As TextBox
    For Each oText In Me.TXTFIELDS
           oText.Locked = X
    Next
    If MDIGALENO.tbToolBar.Buttons("MODIFICAR").Enabled = False Then
       Me.grdsigvit.AllowDelete = True
       Me.grdsigvit.AllowUpdate = True
    Else
       Me.grdsigvit.AllowDelete = False
       Me.grdsigvit.AllowUpdate = False
    End If
End Sub

