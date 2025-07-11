VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MANTENIMIENTO 
   BackColor       =   &H00808080&
   Caption         =   "MEG@SYS"
   ClientHeight    =   7020
   ClientLeft      =   585
   ClientTop       =   1140
   ClientWidth     =   10575
   Icon            =   "MANTENIMIENTO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Palette         =   "MANTENIMIENTO.frx":030A
   Picture         =   "MANTENIMIENTO.frx":087A
   ScaleHeight     =   10783.05
   ScaleMode       =   0  'User
   ScaleWidth      =   2.11504e5
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Frame Frame20Item 
      Height          =   3615
      Left            =   240
      TabIndex        =   343
      Top             =   960
      Width           =   10215
      Begin VB.CommandButton CmdTipoAnamn 
         Height          =   380
         Left            =   3960
         Picture         =   "MANTENIMIENTO.frx":0CED
         Style           =   1  'Graphical
         TabIndex        =   354
         ToolTipText     =   "Ayuda de TITULO"
         Top             =   1440
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.CommandButton CmdAyudaEspe 
         Height          =   380
         Left            =   3960
         Picture         =   "MANTENIMIENTO.frx":112F
         Style           =   1  'Graphical
         TabIndex        =   352
         ToolTipText     =   "Ayuda de Código de Líneas"
         Top             =   960
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "TIPO"
         Height          =   285
         Index           =   171
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   347
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "COD_ESPECIALIDAD"
         Height          =   285
         Index           =   170
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   346
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "ITEM"
         Height          =   285
         Index           =   172
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   345
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "ITEM_DESCRIPCION"
         Height          =   285
         Index           =   173
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   344
         Top             =   2400
         Width           =   7575
      End
      Begin VB.Label LblManIteTipo 
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
         Index           =   0
         Left            =   4560
         TabIndex        =   355
         Top             =   1440
         Width           =   1965
      End
      Begin VB.Label LBLManEspIte 
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
         Left            =   4560
         TabIndex        =   353
         Top             =   960
         Width           =   1965
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
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
         Height          =   225
         Index           =   115
         Left            =   1920
         TabIndex        =   351
         Top             =   1440
         Width           =   360
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Item"
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
         Height          =   225
         Index           =   85
         Left            =   1920
         TabIndex        =   350
         Top             =   1920
         Width           =   360
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Item Descripcion"
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
         Height          =   225
         Index           =   84
         Left            =   840
         TabIndex        =   349
         Top             =   2400
         Width           =   1395
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cod. Especialidad"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   63
         Left            =   840
         TabIndex        =   348
         Top             =   960
         Width           =   1410
      End
   End
   Begin VB.Frame Frame19Medicamentos 
      Height          =   5175
      Left            =   240
      TabIndex        =   264
      Top             =   840
      Width           =   10215
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "COSTO"
         Height          =   285
         Index           =   159
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   281
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "CAN_EXISTENCIA"
         Height          =   285
         Index           =   158
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   280
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "HORARIO"
         Height          =   285
         Index           =   157
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   279
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "DOSIS"
         Height          =   285
         Index           =   156
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   278
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "ARANCEL"
         Height          =   285
         Index           =   160
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   277
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "PRESENTACION"
         Height          =   285
         Index           =   154
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   270
         Top             =   2040
         Width           =   2175
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "NOM_MEDICAMENTO"
         Height          =   285
         Index           =   153
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   269
         Top             =   1680
         Width           =   4815
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "COD_MEDICAMENTO"
         Height          =   285
         Index           =   152
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   268
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "ID_MEDICAMENTO"
         Height          =   285
         Index           =   150
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   267
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "COD_FABRICANTE"
         Height          =   285
         Index           =   151
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   266
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "VIA"
         Height          =   285
         Index           =   155
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   265
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dosis"
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
         Height          =   225
         Index           =   119
         Left            =   3000
         TabIndex        =   286
         Top             =   2760
         Width           =   465
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Costo"
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
         Height          =   225
         Index           =   118
         Left            =   2880
         TabIndex        =   285
         Top             =   3840
         Width           =   465
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cantidad en Existencia"
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
         Height          =   225
         Index           =   117
         Left            =   1560
         TabIndex        =   284
         Top             =   3480
         Width           =   1845
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Horario"
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
         Height          =   225
         Index           =   116
         Left            =   2760
         TabIndex        =   283
         Top             =   3120
         Width           =   645
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Arancel"
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
         Height          =   225
         Index           =   114
         Left            =   2760
         TabIndex        =   282
         Top             =   4200
         Width           =   660
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Id Medicamento"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   113
         Left            =   2160
         TabIndex        =   276
         Top             =   600
         Width           =   1275
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Presentación"
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
         Height          =   225
         Index           =   112
         Left            =   2280
         TabIndex        =   275
         Top             =   2040
         Width           =   1080
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nombre Medicamento"
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
         Height          =   225
         Index           =   111
         Left            =   1680
         TabIndex        =   274
         Top             =   1680
         Width           =   1770
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cod. Medicamento"
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
         Height          =   225
         Index           =   110
         Left            =   1920
         TabIndex        =   273
         Top             =   1320
         Width           =   1485
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cod. Fabricante"
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
         Height          =   225
         Index           =   109
         Left            =   2160
         TabIndex        =   272
         Top             =   960
         Width           =   1275
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Via"
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
         Height          =   225
         Index           =   108
         Left            =   3120
         TabIndex        =   271
         Top             =   2400
         Width           =   285
      End
   End
   Begin VB.Frame frame5INSUMO 
      Height          =   4695
      Left            =   240
      TabIndex        =   175
      Top             =   840
      Visible         =   0   'False
      Width           =   10125
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "ARANCEL"
         Height          =   285
         Index           =   52
         Left            =   4200
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   263
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "COSTO"
         Height          =   285
         Index           =   51
         Left            =   4200
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   262
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "TIPO_INSUMO"
         Height          =   285
         Index           =   53
         Left            =   4200
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   259
         Top             =   960
         Width           =   4815
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "COD_INSUMO"
         Height          =   285
         Index           =   54
         Left            =   4200
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   184
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "NOM_INSUMO"
         Height          =   285
         Index           =   55
         Left            =   4200
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   183
         Top             =   1440
         Width           =   4815
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "COD_PROVEEDOR"
         Height          =   285
         Index           =   56
         Left            =   4200
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   182
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "COD_LINEA"
         Height          =   285
         Index           =   57
         Left            =   4200
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   181
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "STOCK_MINIMO"
         Height          =   285
         Index           =   58
         Left            =   4200
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   180
         Top             =   3360
         Width           =   1020
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "UNIDADES_EXISTENCIA"
         Height          =   285
         Index           =   59
         Left            =   4200
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   179
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CheckBox chkFields 
         DataField       =   "SUSPENDIDO"
         Height          =   285
         Index           =   9
         Left            =   4200
         TabIndex        =   178
         Top             =   4080
         Width           =   450
      End
      Begin VB.CommandButton cmdAyudaLinea 
         Height          =   380
         Left            =   5160
         Picture         =   "MANTENIMIENTO.frx":1571
         Style           =   1  'Graphical
         TabIndex        =   177
         ToolTipText     =   "Ayuda de Código de Líneas"
         Top             =   2280
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.CommandButton cmdAyudaProv 
         Height          =   380
         Left            =   5160
         Picture         =   "MANTENIMIENTO.frx":19B3
         Style           =   1  'Graphical
         TabIndex        =   176
         ToolTipText     =   "Ayuda de Código de Proveedores"
         Top             =   1800
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Arancel"
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
         Height          =   225
         Index           =   107
         Left            =   3525
         TabIndex        =   261
         Top             =   3000
         Width           =   660
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Costo"
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
         Height          =   225
         Index           =   106
         Left            =   3720
         TabIndex        =   260
         Top             =   2640
         Width           =   465
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tipo Insumo"
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
         Height          =   225
         Index           =   105
         Left            =   3120
         TabIndex        =   258
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label LBLManLiIns 
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
         Left            =   5640
         TabIndex        =   193
         Top             =   2280
         Width           =   1965
      End
      Begin VB.Label LBLManProIns 
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
         Left            =   5640
         TabIndex        =   192
         Top             =   1800
         Width           =   1965
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nombre del Insumo"
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
         Height          =   225
         Index           =   21
         Left            =   2520
         TabIndex        =   190
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Código del Proveedor"
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
         Height          =   225
         Index           =   22
         Left            =   2400
         TabIndex        =   189
         Top             =   1800
         Width           =   1725
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Código de Línea"
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
         Height          =   225
         Index           =   23
         Left            =   2880
         TabIndex        =   188
         Top             =   2280
         Width           =   1305
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Stok Minimo"
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
         Height          =   225
         Index           =   24
         Left            =   3120
         TabIndex        =   187
         Top             =   3360
         Width           =   1065
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Unidades en Existencia"
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
         Height          =   225
         Index           =   26
         Left            =   2280
         TabIndex        =   186
         Top             =   3720
         Width           =   1860
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Suspendido"
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
         Height          =   225
         Index           =   29
         Left            =   3120
         TabIndex        =   185
         Top             =   4080
         Width           =   930
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Código del Insumo"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   20
         Left            =   2700
         TabIndex        =   191
         Top             =   480
         Width           =   1500
      End
   End
   Begin VB.Frame Frame18Aranceles 
      Height          =   4335
      Left            =   240
      TabIndex        =   243
      Top             =   840
      Width           =   10215
      Begin VB.CommandButton cmdbusdoc 
         DragIcon        =   "MANTENIMIENTO.frx":1DF5
         Height          =   375
         Left            =   5880
         Picture         =   "MANTENIMIENTO.frx":2237
         Style           =   1  'Graphical
         TabIndex        =   330
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton CmdAranIns 
         Height          =   380
         Left            =   5880
         Picture         =   "MANTENIMIENTO.frx":2679
         Style           =   1  'Graphical
         TabIndex        =   257
         ToolTipText     =   "Ayuda de Tipo de Servicio"
         Top             =   2040
         Width           =   540
      End
      Begin VB.CommandButton CmdAyuSer 
         Height          =   380
         Left            =   5880
         Picture         =   "MANTENIMIENTO.frx":2ABB
         Style           =   1  'Graphical
         TabIndex        =   256
         ToolTipText     =   "Ayuda de Tipo de Servicio"
         Top             =   840
         Width           =   540
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "COSTO"
         Height          =   285
         Index           =   135
         Left            =   4560
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   255
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "CODSERVICIO"
         Height          =   285
         Index           =   130
         Left            =   4560
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   248
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "TIPOSER"
         Height          =   285
         Index           =   131
         Left            =   4560
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   247
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "CODDR"
         Height          =   285
         Index           =   132
         Left            =   4560
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   246
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "CODINS"
         Height          =   285
         Index           =   133
         Left            =   4560
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   245
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "ARANCEL"
         Height          =   285
         Index           =   134
         Left            =   4560
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   244
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label LBLManAraTiSer 
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
         Index           =   0
         Left            =   6480
         TabIndex        =   333
         Top             =   1200
         Width           =   1965
      End
      Begin VB.Label LBLInsti 
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
         TabIndex        =   332
         Top             =   2040
         Width           =   1965
      End
      Begin VB.Label LBLAraDoc 
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
         Index           =   0
         Left            =   6480
         TabIndex        =   331
         Top             =   1560
         Width           =   1965
      End
      Begin VB.Label LBLManAraSer 
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
         TabIndex        =   329
         Top             =   840
         Width           =   1965
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Costo"
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
         Index           =   104
         Left            =   3840
         TabIndex        =   254
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Código del Servicio"
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
         Index           =   98
         Left            =   2880
         TabIndex        =   253
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo de Servicio"
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
         Index           =   96
         Left            =   3240
         TabIndex        =   252
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   " Doctor"
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
         Index           =   95
         Left            =   3120
         TabIndex        =   251
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Instituto"
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
         Index           =   94
         Left            =   2880
         TabIndex        =   250
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Arancel"
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
         Index           =   93
         Left            =   3840
         TabIndex        =   249
         Top             =   2400
         Width           =   735
      End
   End
   Begin VB.Frame Frame9PARAMETROS 
      Height          =   3255
      Left            =   240
      TabIndex        =   131
      Top             =   840
      Visible         =   0   'False
      Width           =   10215
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "DESCRIPCION"
         Height          =   285
         Index           =   32
         Left            =   4080
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   173
         Top             =   1440
         Width           =   5055
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "CODIGO"
         Height          =   285
         Index           =   31
         Left            =   4080
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   135
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "TIPO"
         Height          =   285
         Index           =   30
         Left            =   4080
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   134
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Descripción"
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
         Index           =   12
         Left            =   3120
         TabIndex        =   172
         Top             =   1440
         Width           =   990
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Código"
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
         Index           =   25
         Left            =   3480
         TabIndex        =   133
         Top             =   960
         Width           =   630
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo   "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   255
         Index           =   13
         Left            =   3720
         TabIndex        =   132
         Top             =   480
         Width           =   450
      End
   End
   Begin VB.Frame Frame8ESPECIALIDAD 
      Height          =   3105
      Left            =   240
      TabIndex        =   86
      Top             =   840
      Visible         =   0   'False
      Width           =   10335
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "COD_ESPECIALIDAD"
         Height          =   285
         Index           =   38
         Left            =   4560
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   31
         Top             =   600
         Width           =   1125
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "NOMBRE"
         Height          =   285
         Index           =   39
         Left            =   4560
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   129
         Top             =   1080
         Width           =   3990
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Código de Especialidad"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   34
         Left            =   2700
         TabIndex        =   88
         Top             =   600
         Width           =   1845
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
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
         Height          =   225
         Index           =   31
         Left            =   3900
         TabIndex        =   87
         Top             =   1080
         Width           =   645
      End
   End
   Begin VB.Frame Frame7AREA 
      Height          =   3165
      Left            =   240
      TabIndex        =   83
      Top             =   840
      Visible         =   0   'False
      Width           =   10365
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "DESCRIPCION"
         Height          =   285
         Index           =   37
         Left            =   3840
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   128
         Top             =   1320
         Width           =   4695
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "CODAREA"
         Height          =   285
         Index           =   36
         Left            =   3840
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   127
         Top             =   720
         Width           =   555
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
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
         Height          =   225
         Index           =   33
         Left            =   2760
         TabIndex        =   85
         Top             =   1320
         Width           =   990
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Código de Area"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   32
         Left            =   2520
         TabIndex        =   84
         Top             =   720
         Width           =   1245
      End
   End
   Begin VB.Frame Frame4LINEA 
      Height          =   3315
      Left            =   240
      TabIndex        =   54
      Top             =   840
      Visible         =   0   'False
      Width           =   10200
      Begin VB.PictureBox picButtons 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   315
         ScaleHeight     =   300
         ScaleWidth      =   5775
         TabIndex        =   67
         Top             =   2880
         Width           =   5775
      End
      Begin VB.PictureBox picStatBox 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   270
         ScaleHeight     =   300
         ScaleWidth      =   5775
         TabIndex        =   66
         Top             =   2880
         Width           =   5775
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "LIN_NOMBRE"
         Height          =   285
         Index           =   141
         Left            =   3720
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   25
         Top             =   960
         Width           =   4695
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "COD_LINEA"
         Height          =   285
         Index           =   140
         Left            =   3720
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   24
         Top             =   480
         Width           =   765
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "DESCRIPCION"
         Height          =   285
         Index           =   142
         Left            =   3720
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   26
         Top             =   1440
         Width           =   4695
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
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
         Height          =   225
         Index           =   19
         Left            =   2700
         TabIndex        =   65
         Top             =   1440
         Width           =   990
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nombre de Linea"
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
         Height          =   225
         Index           =   18
         Left            =   2340
         TabIndex        =   60
         Top             =   960
         Width           =   1380
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Código de Línea"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   16
         Left            =   2400
         TabIndex        =   57
         Top             =   480
         Width           =   1305
      End
   End
   Begin VB.Frame Frame1PERSONA 
      Height          =   6135
      Left            =   240
      TabIndex        =   40
      Top             =   720
      Width           =   10335
      Begin VB.TextBox TXTFIELDS 
         BackColor       =   &H00FFFFFF&
         DataField       =   "FECNAC"
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
         Index           =   20
         Left            =   6000
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   328
         Text            =   " "
         Top             =   1680
         Width           =   1065
      End
      Begin VB.TextBox TXTFIELDS 
         BackColor       =   &H00FFFFFF&
         DataField       =   "CARGO"
         Height          =   285
         Index           =   12
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   200
         Text            =   " "
         Top             =   3600
         Width           =   1065
      End
      Begin VB.CommandButton CmdAyudaCargo 
         Height          =   375
         Left            =   2760
         Picture         =   "MANTENIMIENTO.frx":2EFD
         Style           =   1  'Graphical
         TabIndex        =   199
         ToolTipText     =   "Ayuda de CODIGO DE CARGO"
         Top             =   3600
         Width           =   540
      End
      Begin VB.CommandButton CmdAyudaTipoPerso 
         Height          =   375
         Left            =   6720
         Picture         =   "MANTENIMIENTO.frx":333F
         Style           =   1  'Graphical
         TabIndex        =   198
         ToolTipText     =   "Ayuda de TIPO DE PERSONA"
         Top             =   3120
         Width           =   540
      End
      Begin VB.TextBox TXTFIELDS 
         BackColor       =   &H00FFFFFF&
         DataField       =   "TIPOPERSONA"
         Height          =   285
         Index           =   11
         Left            =   5640
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   197
         Text            =   " "
         Top             =   3120
         Width           =   1065
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "RADIO"
         Height          =   285
         Index           =   18
         Left            =   5400
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   15
         Top             =   5040
         Width           =   1455
      End
      Begin VB.TextBox TXTFIELDS 
         BackColor       =   &H00FFFFFF&
         DataField       =   "MOVIL"
         Height          =   285
         Index           =   17
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   13
         Text            =   " "
         Top             =   5040
         Width           =   1905
      End
      Begin VB.TextBox TXTFIELDS 
         BackColor       =   &H00FFFFFF&
         DataField       =   "nombres"
         Height          =   285
         Index           =   2
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   2
         Top             =   600
         Width           =   4335
      End
      Begin VB.CommandButton cmdFicha 
         BackColor       =   &H8000000D&
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
         Height          =   495
         Left            =   8160
         MaskColor       =   &H8000000D&
         TabIndex        =   154
         Top             =   3600
         UseMaskColor    =   -1  'True
         Width           =   1935
      End
      Begin MSComDlg.CommonDialog Comdiapers 
         Left            =   9000
         Top             =   2760
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "Imágenes (*.jpg;*.bmp)|*.jpg;*.bmp"
         FilterIndex     =   1
         Flags           =   8192
         InitDir         =   "D:\GALENO\PICTURE\FOTO"
      End
      Begin VB.CommandButton CMDFOTOPERS 
         BackColor       =   &H8000000A&
         Caption         =   "Agregar Fotografia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   7320
         Style           =   1  'Graphical
         TabIndex        =   153
         Top             =   1560
         Width           =   1890
      End
      Begin VB.CommandButton cmdayudaciudadpar 
         Height          =   375
         Left            =   2760
         Picture         =   "MANTENIMIENTO.frx":3781
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Ayuda de CODIGO DE CIUDAD"
         Top             =   3120
         Width           =   540
      End
      Begin VB.CommandButton cmdayudaciudadlab 
         Height          =   380
         Left            =   2760
         Picture         =   "MANTENIMIENTO.frx":3BC3
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Ayuda de CODIGO DE CIUDAD"
         Top             =   4320
         Width           =   540
      End
      Begin VB.CommandButton cmdayudaecivil 
         Height          =   380
         Left            =   6120
         Picture         =   "MANTENIMIENTO.frx":4005
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Ayuda de ESTADO CIVIL"
         Top             =   2160
         Width           =   540
      End
      Begin VB.CommandButton cmdayudagenero 
         Height          =   380
         Left            =   2760
         Picture         =   "MANTENIMIENTO.frx":4447
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Ayuda de GENERO"
         Top             =   2040
         Width           =   540
      End
      Begin VB.CommandButton Cmdayudatitulo 
         Height          =   380
         Left            =   2760
         Picture         =   "MANTENIMIENTO.frx":4889
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Ayuda de TITULO"
         Top             =   1680
         Width           =   540
      End
      Begin VB.CommandButton Cmdayudatipdoc 
         Height          =   380
         Left            =   3000
         Picture         =   "MANTENIMIENTO.frx":4CCB
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Ayuda de TIPO DE DOCUMENTO"
         Top             =   240
         Width           =   540
      End
      Begin VB.TextBox TXTFIELDS 
         BackColor       =   &H00FFFFFF&
         DataField       =   "E_MAIL"
         Height          =   285
         Index           =   16
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   23
         Text            =   " "
         Top             =   4680
         Width           =   5745
      End
      Begin VB.TextBox TXTFIELDS 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COD_CIUDADL"
         Height          =   285
         Index           =   14
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   14
         Text            =   " "
         Top             =   4320
         Width           =   1065
      End
      Begin VB.TextBox TXTFIELDS 
         BackColor       =   &H00FFFFFF&
         DataField       =   "DOM_LABORAL"
         Height          =   285
         Index           =   13
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   11
         Text            =   " "
         Top             =   3960
         Width           =   5265
      End
      Begin VB.TextBox TXTFIELDS 
         BackColor       =   &H00FFFFFF&
         DataField       =   "COD_CIUDADP"
         Height          =   285
         Index           =   10
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   10
         Text            =   " "
         Top             =   3120
         Width           =   1065
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "GENERO"
         Height          =   285
         Index           =   6
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   6
         Top             =   2040
         Width           =   1050
      End
      Begin VB.TextBox TXTFIELDS 
         BackColor       =   &H00FFFFFF&
         DataField       =   "NRODOC"
         Height          =   285
         Index           =   1
         Left            =   5880
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   1
         Text            =   " "
         Top             =   240
         Width           =   1305
      End
      Begin VB.TextBox TXTFIELDS 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ESTCIV"
         Height          =   285
         Index           =   7
         Left            =   5040
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   7
         Text            =   " "
         Top             =   2160
         Width           =   1065
      End
      Begin VB.TextBox TXTFIELDS 
         BackColor       =   &H00FFFFFF&
         DataField       =   "OBSERVACION"
         Height          =   675
         Index           =   19
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   200
         TabIndex        =   16
         Text            =   " "
         Top             =   5400
         Width           =   7125
      End
      Begin VB.TextBox TXTFIELDS 
         BackColor       =   &H00FFFFFF&
         DataField       =   "DOM_PARTICULAR"
         Height          =   285
         Index           =   9
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   9
         Text            =   " "
         Top             =   2760
         Width           =   6225
      End
      Begin VB.TextBox TXTFIELDS 
         BackColor       =   &H00FFFFFF&
         DataField       =   "apellidos"
         Height          =   285
         Index           =   3
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   3
         Text            =   " "
         Top             =   960
         Width           =   4320
      End
      Begin VB.TextBox TXTFIELDS 
         BackColor       =   &H00FFFFFF&
         DataField       =   "TITULO"
         Height          =   285
         Index           =   5
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   5
         Text            =   " "
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox TXTFIELDS 
         BackColor       =   &H00FFFFFF&
         DataField       =   "RUC"
         Height          =   285
         Index           =   8
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   8
         Text            =   " "
         Top             =   2400
         Width           =   1410
      End
      Begin VB.TextBox TXTFIELDS 
         BackColor       =   &H00FFFFFF&
         DataField       =   "APECASADA"
         Height          =   255
         Index           =   4
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   4
         Text            =   " "
         Top             =   1320
         Width           =   4155
      End
      Begin VB.TextBox TXTFIELDS 
         BackColor       =   &H00FFFFFF&
         DataField       =   "TELEFONOS"
         Height          =   285
         Index           =   15
         Left            =   4440
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   12
         Text            =   " "
         Top             =   4320
         Width           =   3570
      End
      Begin VB.TextBox TXTFIELDS 
         BackColor       =   &H00FFFFFF&
         DataField       =   "tipdoc"
         Height          =   285
         Index           =   0
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   0
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Cargo"
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
         Height          =   225
         Left            =   1080
         TabIndex        =   201
         Top             =   3600
         Width           =   510
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Persona"
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
         Height          =   225
         Left            =   4320
         TabIndex        =   196
         Top             =   3120
         Width           =   1305
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Radio Mensaje"
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
         Height          =   225
         Index           =   8
         Left            =   4140
         TabIndex        =   168
         Top             =   5040
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Telef. Celular"
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
         Height          =   225
         Left            =   480
         TabIndex        =   159
         Top             =   5040
         Width           =   1140
      End
      Begin VB.Image PICPER 
         Height          =   1290
         Left            =   7320
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1875
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Correo Electronico"
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
         Height          =   225
         Left            =   120
         TabIndex        =   152
         Top             =   4680
         Width           =   1560
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad"
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
         Height          =   225
         Left            =   1080
         TabIndex        =   151
         Top             =   4320
         Width           =   570
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Dirección Laboral"
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
         Height          =   225
         Left            =   120
         TabIndex        =   150
         Top             =   3960
         Width           =   1485
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad"
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
         Height          =   225
         Left            =   1080
         TabIndex        =   149
         Top             =   3120
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Genero"
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
         Height          =   225
         Left            =   1080
         TabIndex        =   117
         Top             =   2040
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Apellido de Casada"
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
         Height          =   225
         Left            =   120
         TabIndex        =   43
         Top             =   1320
         Width           =   1530
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Nacimiento"
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
         Height          =   225
         Left            =   4080
         TabIndex        =   53
         Top             =   1680
         Width           =   1680
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Estado Civil"
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
         Height          =   225
         Left            =   4080
         TabIndex        =   52
         Top             =   2160
         Width           =   960
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
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
         Height          =   225
         Left            =   600
         TabIndex        =   49
         Top             =   5520
         Width           =   1020
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Dirección Part."
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
         Height          =   225
         Left            =   360
         TabIndex        =   48
         Top             =   2760
         Width           =   1260
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Telefonos"
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
         Height          =   225
         Left            =   3600
         TabIndex        =   47
         Top             =   4320
         Width           =   795
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Apellidos"
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
         Height          =   225
         Left            =   180
         TabIndex        =   46
         Top             =   960
         Width           =   765
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Titulo"
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
         Height          =   225
         Left            =   1200
         TabIndex        =   45
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nombres"
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
         Height          =   225
         Left            =   240
         TabIndex        =   44
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "R U C"
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
         Height          =   225
         Left            =   1200
         TabIndex        =   42
         Top             =   2400
         Width           =   480
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Documento"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   195
         TabIndex        =   41
         Top             =   240
         Width           =   1545
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Numero de Documento"
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
         Height          =   225
         Left            =   3960
         TabIndex        =   130
         Top             =   240
         Width           =   1845
      End
   End
   Begin VB.Frame Frame3PROVEEDOR 
      Height          =   4170
      Left            =   240
      TabIndex        =   109
      Top             =   840
      Visible         =   0   'False
      Width           =   10335
      Begin VB.CommandButton cmdAyudaProvee 
         Height          =   380
         Left            =   4560
         Picture         =   "MANTENIMIENTO.frx":510D
         Style           =   1  'Graphical
         TabIndex        =   137
         ToolTipText     =   "Ayuda de Código de Ciudad"
         Top             =   1920
         Width           =   540
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "COD_PROVEEDOR"
         Height          =   285
         Index           =   64
         Left            =   3600
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   118
         Top             =   480
         Width           =   1140
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "NOM_PROVEEDOR"
         Height          =   285
         Index           =   65
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   119
         Top             =   840
         Width           =   3615
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "CONTACTO"
         Height          =   285
         Index           =   66
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   120
         Top             =   1230
         Width           =   3615
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "DIRECCION"
         Height          =   285
         Index           =   67
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   121
         Top             =   1590
         Width           =   5175
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "CIUDAD"
         Height          =   285
         Index           =   68
         Left            =   3600
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   122
         Top             =   1920
         Width           =   885
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "TELEFONO"
         Height          =   285
         Index           =   69
         Left            =   3600
         MaxLength       =   24
         TabIndex        =   123
         Top             =   2280
         Width           =   1245
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "FAX"
         Height          =   285
         Index           =   70
         Left            =   3600
         MaxLength       =   24
         TabIndex        =   124
         Top             =   2640
         Width           =   1275
      End
      Begin VB.Label LBLManCiuProv 
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
         Index           =   0
         Left            =   5160
         TabIndex        =   174
         Top             =   1920
         Width           =   1965
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Código del Proveedor"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   103
         Left            =   1860
         TabIndex        =   116
         Top             =   480
         Width           =   1725
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nombre del Proveedor"
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
         Height          =   225
         Index           =   102
         Left            =   1800
         TabIndex        =   115
         Top             =   840
         Width           =   1800
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nombre del Contacto"
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
         Height          =   225
         Index           =   101
         Left            =   1920
         TabIndex        =   114
         Top             =   1215
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
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
         Height          =   225
         Index           =   100
         Left            =   2820
         TabIndex        =   113
         Top             =   1590
         Width           =   810
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ciudad"
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
         Height          =   225
         Index           =   99
         Left            =   3060
         TabIndex        =   112
         Top             =   1935
         Width           =   570
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Teléfono"
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
         Height          =   225
         Index           =   97
         Left            =   2880
         TabIndex        =   111
         Top             =   2280
         Width           =   705
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fax"
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
         Height          =   225
         Index           =   91
         Left            =   3300
         TabIndex        =   110
         Top             =   2640
         Width           =   285
      End
   End
   Begin VB.Frame Frame11INSTITUTO 
      Height          =   3600
      Left            =   240
      TabIndex        =   98
      Top             =   840
      Visible         =   0   'False
      Width           =   10380
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "PAGINA WEB"
         Height          =   285
         Index           =   89
         Left            =   2880
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   64
         Top             =   2880
         Width           =   4635
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "CORREO ELECTRONICO"
         Height          =   285
         Index           =   88
         Left            =   2880
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   63
         Top             =   2520
         Width           =   4635
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "TELEF1"
         Height          =   285
         Index           =   86
         Left            =   2880
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   61
         Top             =   1800
         Width           =   2115
      End
      Begin VB.CommandButton cmdayudaInst 
         Height          =   380
         Left            =   3840
         Picture         =   "MANTENIMIENTO.frx":554F
         Style           =   1  'Graphical
         TabIndex        =   136
         ToolTipText     =   "Ayuda de Código de Ciudad"
         Top             =   1440
         Width           =   540
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "CODINS"
         Height          =   285
         Index           =   82
         Left            =   2880
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   55
         Top             =   360
         Width           =   1350
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "INSTITUTO"
         Height          =   285
         Index           =   83
         Left            =   2880
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   56
         Top             =   720
         Width           =   3375
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "DOMICILIO"
         Height          =   285
         Index           =   84
         Left            =   2880
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   58
         Top             =   1080
         Width           =   6495
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "TELEF2"
         Height          =   285
         Index           =   87
         Left            =   2880
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   62
         Top             =   2160
         Width           =   2115
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "CODCIU"
         Height          =   285
         Index           =   85
         Left            =   2880
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   59
         Top             =   1440
         Width           =   945
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Página Web"
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
         Height          =   225
         Index           =   15
         Left            =   1800
         TabIndex        =   195
         Top             =   2880
         Width           =   990
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Correo Electronico"
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
         Height          =   225
         Index           =   14
         Left            =   1320
         TabIndex        =   194
         Top             =   2520
         Width           =   1560
      End
      Begin VB.Label LBLManCiuIns 
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
         Left            =   4320
         TabIndex        =   171
         Top             =   1440
         Width           =   1965
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Teléfono 1"
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
         Height          =   225
         Index           =   73
         Left            =   2040
         TabIndex        =   104
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Teléfono 2"
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
         Height          =   225
         Index           =   72
         Left            =   2040
         TabIndex        =   103
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nombre del Instituto"
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
         Height          =   225
         Index           =   71
         Left            =   1200
         TabIndex        =   102
         Top             =   720
         Width           =   1665
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Código del Instituto"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   70
         Left            =   1260
         TabIndex        =   101
         Top             =   360
         Width           =   1590
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
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
         Height          =   225
         Index           =   69
         Left            =   2040
         TabIndex        =   100
         Top             =   1080
         Width           =   810
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ciudad"
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
         Height          =   225
         Index           =   67
         Left            =   2280
         TabIndex        =   99
         Top             =   1440
         Width           =   570
      End
   End
   Begin VB.Frame Frame10VALORNORMAL 
      Height          =   3975
      Left            =   360
      TabIndex        =   89
      Top             =   840
      Visible         =   0   'False
      Width           =   10245
      Begin VB.CommandButton CmdAyuGene 
         Height          =   380
         Left            =   5640
         Picture         =   "MANTENIMIENTO.frx":5991
         Style           =   1  'Graphical
         TabIndex        =   326
         ToolTipText     =   "Ayuda de TITULO"
         Top             =   840
         Width           =   540
      End
      Begin VB.CommandButton cmdAyudaAnalisis 
         Height          =   380
         Left            =   5640
         Picture         =   "MANTENIMIENTO.frx":5DD3
         Style           =   1  'Graphical
         TabIndex        =   146
         Top             =   525
         Width           =   540
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "SEXO"
         Height          =   285
         Index           =   41
         Left            =   4560
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   33
         Top             =   870
         Width           =   1080
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "EDAD_INF"
         Height          =   285
         Index           =   42
         Left            =   4560
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   34
         Top             =   1200
         Width           =   330
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "EDAD_SUP"
         Height          =   285
         Index           =   43
         Left            =   4560
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   35
         Top             =   1530
         Width           =   345
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "LIM_INF"
         Height          =   285
         Index           =   44
         Left            =   4560
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   36
         Top             =   1860
         Width           =   855
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "LIM_SUP"
         Height          =   285
         Index           =   45
         Left            =   4560
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   37
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "LIM_INFN"
         Height          =   285
         Index           =   46
         Left            =   4560
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   38
         Top             =   2475
         Width           =   855
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "LIM_SUPN"
         Height          =   285
         Index           =   47
         Left            =   4560
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   39
         Top             =   2805
         Width           =   855
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "CODSERVICIO"
         Height          =   285
         Index           =   40
         Left            =   4560
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   32
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label LBLManGENE 
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
         Left            =   6240
         TabIndex        =   327
         Top             =   840
         Width           =   1965
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Código del Análisis"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   55
         Left            =   2940
         TabIndex        =   94
         Top             =   600
         Width           =   1590
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Años"
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
         Height          =   225
         Index           =   30
         Left            =   5040
         TabIndex        =   145
         Top             =   1560
         Width           =   420
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Años"
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
         Height          =   225
         Index           =   28
         Left            =   5040
         TabIndex        =   144
         Top             =   1245
         Width           =   420
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Límite Superior Normal"
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
         Height          =   225
         Index           =   79
         Left            =   2580
         TabIndex        =   97
         Top             =   2850
         Width           =   1980
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Límite Inferior Normal"
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
         Height          =   225
         Index           =   78
         Left            =   2700
         TabIndex        =   96
         Top             =   2520
         Width           =   1890
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Edad Inferior"
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
         Height          =   225
         Index           =   59
         Left            =   3480
         TabIndex        =   95
         Top             =   1200
         Width           =   1065
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sexo"
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
         Height          =   225
         Index           =   42
         Left            =   4140
         TabIndex        =   93
         Top             =   870
         Width           =   390
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Límite Inferior"
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
         Height          =   225
         Index           =   41
         Left            =   3300
         TabIndex        =   92
         Top             =   1860
         Width           =   1230
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Edad Superior"
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
         Height          =   225
         Index           =   40
         Left            =   3420
         TabIndex        =   91
         Top             =   1545
         Width           =   1155
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Límite Superior"
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
         Height          =   225
         Index           =   39
         Left            =   3180
         TabIndex        =   90
         Top             =   2175
         Width           =   1320
      End
   End
   Begin VB.Frame Frame6DOCTOR 
      Height          =   5940
      Left            =   360
      TabIndex        =   68
      Top             =   840
      Visible         =   0   'False
      Width           =   10275
      Begin VB.CommandButton cmdAyudaEspec 
         Height          =   380
         Left            =   960
         MouseIcon       =   "MANTENIMIENTO.frx":6215
         Picture         =   "MANTENIMIENTO.frx":6657
         Style           =   1  'Graphical
         TabIndex        =   108
         ToolTipText     =   "Ayuda de Código de Especialidad"
         Top             =   4080
         Width           =   540
      End
      Begin MSDataGridLib.DataGrid grd_Especialidad 
         Height          =   975
         Left            =   1080
         TabIndex        =   170
         Top             =   4560
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   1720
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
         BackColor       =   13810592
         HeadLines       =   1
         RowHeight       =   18
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ESPECIALIDADES DEL DOCTOR"
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
      Begin VB.CommandButton busficha 
         DragIcon        =   "MANTENIMIENTO.frx":6A99
         Height          =   375
         Left            =   5880
         Picture         =   "MANTENIMIENTO.frx":6EDB
         Style           =   1  'Graphical
         TabIndex        =   158
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdAyudaDocCiu2 
         Height          =   380
         Left            =   4200
         Picture         =   "MANTENIMIENTO.frx":731D
         Style           =   1  'Graphical
         TabIndex        =   143
         ToolTipText     =   "Ayuda de Código de Ciudad"
         Top             =   3240
         Width           =   540
      End
      Begin VB.CommandButton cmdAyudaDocCiu1 
         Height          =   380
         Left            =   4200
         Picture         =   "MANTENIMIENTO.frx":775F
         Style           =   1  'Graphical
         TabIndex        =   142
         ToolTipText     =   "Ayuda de Código de Ciudad"
         Top             =   2520
         Width           =   540
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "CODDR"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   15370
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   22
         Left            =   3240
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   27
         Top             =   360
         Width           =   1110
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "REGPRF"
         Height          =   285
         Index           =   23
         Left            =   3120
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   28
         Top             =   3720
         Width           =   1395
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "TIPDOC"
         Height          =   285
         Index           =   24
         Left            =   3240
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   29
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "NRODOC"
         Height          =   285
         Index           =   25
         Left            =   4320
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   30
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label LBLManDoc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00D2BBA0&
         DataField       =   "RADIO"
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
         Index           =   12
         Left            =   7560
         TabIndex        =   169
         Top             =   3840
         Width           =   1965
      End
      Begin VB.Label LBLManDoc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00D2BBA0&
         DataField       =   "COD_CIUDADL"
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
         Left            =   3240
         TabIndex        =   167
         Top             =   3240
         Width           =   885
      End
      Begin VB.Label LBLManDoc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00D2BBA0&
         DataField       =   "DOM_LABORAL"
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
         Left            =   3240
         TabIndex        =   166
         Top             =   2880
         Width           =   1965
      End
      Begin VB.Label LBLManDoc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00D2BBA0&
         DataField       =   "COD_CIUDADP"
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
         Index           =   9
         Left            =   3240
         TabIndex        =   165
         Top             =   2520
         Width           =   885
      End
      Begin VB.Label LBLManDoc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00D2BBA0&
         DataField       =   "DOM_PARTICULAR"
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
         Index           =   8
         Left            =   3240
         TabIndex        =   164
         Top             =   2160
         Width           =   1965
      End
      Begin VB.Label LBLManDoc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00D2BBA0&
         DataField       =   "MOVIL"
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
         Left            =   7680
         TabIndex        =   163
         Top             =   1920
         Width           =   1965
      End
      Begin VB.Label LBLManDoc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00D2BBA0&
         DataField       =   "TELEFONOS"
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
         Left            =   3240
         TabIndex        =   162
         Top             =   1800
         Width           =   1965
      End
      Begin VB.Label LBLManDoc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00D2BBA0&
         DataField       =   "E_MAIL"
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
         Left            =   7680
         TabIndex        =   161
         Top             =   1440
         Width           =   1965
      End
      Begin VB.Label LBLManDoc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00D2BBA0&
         DataField       =   "RUC"
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
         Left            =   3240
         TabIndex        =   160
         Top             =   1440
         Width           =   1965
      End
      Begin VB.Label LBLManDoc 
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
         Left            =   4680
         TabIndex        =   157
         Top             =   3240
         Width           =   1965
      End
      Begin VB.Label LBLManDoc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00D2BBA0&
         Caption         =   " "
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
         Left            =   3240
         TabIndex        =   156
         Top             =   1080
         Width           =   45
      End
      Begin VB.Label LBLManDoc 
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
         Left            =   4680
         TabIndex        =   155
         Top             =   2520
         Width           =   1965
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ciudad"
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
         Height          =   225
         Index           =   36
         Left            =   2520
         TabIndex        =   126
         Top             =   3240
         Width           =   570
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ciudad"
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
         Height          =   225
         Index           =   35
         Left            =   2520
         TabIndex        =   125
         Top             =   2520
         Width           =   570
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dirección Laboral"
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
         Height          =   225
         Index           =   54
         Left            =   1560
         TabIndex        =   82
         Top             =   2880
         Width           =   1485
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Código del Doctor"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   53
         Left            =   1740
         TabIndex        =   81
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Registro  Profesional"
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
         Height          =   225
         Index           =   52
         Left            =   1260
         TabIndex        =   80
         Top             =   3720
         Width           =   1755
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nombre del Doctor"
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
         Height          =   225
         Index           =   51
         Left            =   1620
         TabIndex        =   79
         Top             =   1080
         Width           =   1530
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tipo y Número de Documento"
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
         Height          =   225
         Index           =   50
         Left            =   780
         TabIndex        =   78
         Top             =   720
         Width           =   2385
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Agregar Especialidades:"
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
         Height          =   225
         Index           =   49
         Left            =   1800
         TabIndex        =   77
         Top             =   4200
         Width           =   1995
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "R.U.C."
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
         Height          =   225
         Index           =   48
         Left            =   2580
         TabIndex        =   76
         Top             =   1440
         Width           =   525
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Radio Mensaje"
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
         Height          =   225
         Index           =   47
         Left            =   6240
         TabIndex        =   75
         Top             =   3840
         Width           =   1200
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Número de Celular"
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
         Height          =   225
         Index           =   46
         Left            =   6000
         TabIndex        =   74
         Top             =   1920
         Width           =   1560
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "E-Mail"
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
         Height          =   225
         Index           =   45
         Left            =   7080
         TabIndex        =   73
         Top             =   1440
         Width           =   540
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Teléfonos"
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
         Height          =   225
         Index           =   44
         Left            =   2340
         TabIndex        =   72
         Top             =   1800
         Width           =   795
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dirección Particular"
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
         Height          =   225
         Index           =   43
         Left            =   1440
         TabIndex        =   69
         Top             =   2160
         Width           =   1710
      End
   End
   Begin VB.Frame Frame2Empresas 
      Height          =   4215
      Left            =   240
      TabIndex        =   228
      Top             =   840
      Width           =   10335
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "COD_EMPRESA"
         Height          =   285
         Index           =   120
         Left            =   2880
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   235
         Top             =   720
         Width           =   1080
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "COD_CIUDAD"
         Height          =   285
         Index           =   123
         Left            =   2880
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   234
         Top             =   1800
         Width           =   900
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "DIRECCION"
         Height          =   285
         Index           =   122
         Left            =   2880
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   233
         Top             =   1440
         Width           =   6135
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "TELEFONO"
         Height          =   285
         Index           =   124
         Left            =   2880
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   232
         Top             =   2160
         Width           =   3720
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "NOMBRE"
         Height          =   285
         Index           =   121
         Left            =   2880
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   231
         Top             =   1080
         Width           =   6135
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "RUC"
         Height          =   285
         Index           =   125
         Left            =   2880
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   230
         Top             =   2520
         Width           =   3720
      End
      Begin VB.CommandButton CmdAyudaCiud 
         Height          =   380
         Left            =   3840
         Picture         =   "MANTENIMIENTO.frx":7BA1
         Style           =   1  'Graphical
         TabIndex        =   229
         ToolTipText     =   "Ayuda de Código de Ciudad"
         Top             =   1800
         Width           =   540
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ruc"
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
         Height          =   225
         Index           =   90
         Left            =   2520
         TabIndex        =   242
         Top             =   2520
         Width           =   330
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
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
         Height          =   225
         Index           =   88
         Left            =   2040
         TabIndex        =   241
         Top             =   1440
         Width           =   810
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ciudad "
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
         Height          =   225
         Index           =   87
         Left            =   2280
         TabIndex        =   240
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nombre "
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
         Height          =   225
         Index           =   86
         Left            =   2160
         TabIndex        =   239
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Código de Empresa "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   83
         Left            =   1320
         TabIndex        =   238
         Top             =   720
         Width           =   1590
      End
      Begin VB.Label LBLManCiuEmp 
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
         Left            =   4320
         TabIndex        =   237
         Top             =   1800
         Width           =   1965
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Teléfono"
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
         Height          =   225
         Index           =   82
         Left            =   2160
         TabIndex        =   236
         Top             =   2160
         Width           =   705
      End
   End
   Begin VB.Frame Frame17TipoServicio 
      Height          =   3495
      Left            =   240
      TabIndex        =   222
      Top             =   840
      Width           =   10335
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "NOMTIPOSER"
         Height          =   285
         Index           =   76
         Left            =   3000
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   224
         Top             =   1320
         Width           =   6015
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "TIPOSER"
         Height          =   285
         Index           =   75
         Left            =   3000
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   223
         Top             =   840
         Width           =   540
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nombre del Tipo de Servicio"
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
         Height          =   225
         Index           =   81
         Left            =   720
         TabIndex        =   226
         Top             =   1320
         Width           =   2280
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tipo Servicio"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   80
         Left            =   1920
         TabIndex        =   225
         Top             =   840
         Width           =   1080
      End
   End
   Begin VB.Frame Frame15Habitaciones 
      Height          =   3855
      Left            =   360
      TabIndex        =   208
      Top             =   840
      Width           =   10215
      Begin VB.CommandButton CmdEstadoCama 
         DragIcon        =   "MANTENIMIENTO.frx":7FE3
         Height          =   375
         Left            =   5280
         Picture         =   "MANTENIMIENTO.frx":8425
         Style           =   1  'Graphical
         TabIndex        =   221
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "ESTADO"
         Height          =   285
         Index           =   93
         Left            =   3720
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   220
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton CmdAyudaHabitacion 
         DragIcon        =   "MANTENIMIENTO.frx":8867
         Height          =   375
         Left            =   4920
         Picture         =   "MANTENIMIENTO.frx":8CA9
         Style           =   1  'Graphical
         TabIndex        =   219
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "COD_CAMA"
         Height          =   285
         Index           =   91
         Left            =   3720
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   218
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "COD_TIPO"
         Height          =   285
         Index           =   90
         Left            =   3720
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   212
         Top             =   480
         Width           =   1140
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "COD_RUB"
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   94
         Left            =   3720
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   211
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton CmdAyudaTipoHab 
         DragIcon        =   "MANTENIMIENTO.frx":90EB
         Height          =   375
         Left            =   4920
         Picture         =   "MANTENIMIENTO.frx":952D
         Style           =   1  'Graphical
         TabIndex        =   210
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "COD_HABITACION"
         Height          =   285
         Index           =   92
         Left            =   3720
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   209
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Código Cama"
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
         Height          =   225
         Index           =   5
         Left            =   2520
         TabIndex        =   217
         Top             =   1200
         Width           =   1080
      End
      Begin VB.Label l 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Código de Tipo"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   2505
         TabIndex        =   216
         Top             =   495
         Width           =   1200
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Rubro"
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
         Height          =   225
         Index           =   4
         Left            =   3120
         TabIndex        =   215
         Top             =   1920
         Width           =   510
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Estado"
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
         Height          =   225
         Index           =   3
         Left            =   3120
         TabIndex        =   214
         Top             =   1560
         Width           =   525
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Código Habitación"
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
         Height          =   225
         Index           =   2
         Left            =   2160
         TabIndex        =   213
         Top             =   840
         Width           =   1485
      End
   End
   Begin VB.Frame Frame13USUARIO 
      Height          =   3585
      Left            =   240
      TabIndex        =   105
      Top             =   720
      Visible         =   0   'False
      Width           =   10335
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "NRODOC"
         Height          =   285
         Index           =   62
         Left            =   4920
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   204
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "TIPDOC"
         Height          =   285
         Index           =   61
         Left            =   3960
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   203
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton CmdBusUsua 
         DragIcon        =   "MANTENIMIENTO.frx":996F
         Height          =   375
         Left            =   6360
         Picture         =   "MANTENIMIENTO.frx":9DB1
         Style           =   1  'Graphical
         TabIndex        =   202
         Top             =   960
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Administrador"
         DataField       =   "ADMINISTRADOR"
         Height          =   225
         Left            =   3960
         TabIndex        =   148
         Top             =   2400
         Width           =   1470
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "CLAVE"
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   63
         Left            =   3960
         Locked          =   -1  'True
         MaxLength       =   40
         PasswordChar    =   "*"
         TabIndex        =   51
         Top             =   1920
         Width           =   2715
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "ID_USUARIO"
         Height          =   285
         Index           =   60
         Left            =   3960
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   50
         Top             =   480
         Width           =   1020
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tipo y Número de Documento"
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
         Height          =   225
         Index           =   1
         Left            =   1560
         TabIndex        =   207
         Top             =   960
         Width           =   2385
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nombre del Usuario"
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
         Height          =   225
         Index           =   0
         Left            =   2280
         TabIndex        =   206
         Top             =   1440
         Width           =   1620
      End
      Begin VB.Label LBLManUsu 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00D2BBA0&
         Caption         =   " "
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
         Left            =   3960
         TabIndex        =   205
         Top             =   1440
         Width           =   525
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Clave"
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
         Height          =   225
         Index           =   92
         Left            =   3480
         TabIndex        =   107
         Top             =   1920
         Width           =   450
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Código del Usuario"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   89
         Left            =   2400
         TabIndex        =   106
         Top             =   495
         Width           =   1545
      End
   End
   Begin VB.Frame Frame14TABLAVALOR 
      Height          =   3555
      Left            =   240
      TabIndex        =   138
      Top             =   840
      Visible         =   0   'False
      Width           =   10230
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "CODITEM"
         Height          =   285
         Index           =   80
         Left            =   4800
         MaxLength       =   4
         TabIndex        =   71
         Top             =   1080
         Width           =   780
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "NOMBRE"
         Height          =   285
         Index           =   81
         Left            =   4800
         MaxLength       =   30
         TabIndex        =   141
         Top             =   1560
         Width           =   3375
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "CODTAB"
         Height          =   285
         Index           =   79
         Left            =   4800
         MaxLength       =   4
         TabIndex        =   70
         Top             =   600
         Width           =   780
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Código de Item"
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
         Height          =   225
         Index           =   37
         Left            =   3600
         TabIndex        =   147
         Top             =   1080
         Width           =   1200
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nombre de la Tabla de Valores"
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
         Height          =   225
         Index           =   27
         Left            =   2280
         TabIndex        =   140
         Top             =   1560
         Width           =   2475
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Código de Tabla de Valores"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   11
         Left            =   2520
         TabIndex        =   139
         Top             =   600
         Width           =   2205
      End
   End
   Begin VB.Frame Frame12Procedimientos 
      Height          =   5055
      Left            =   240
      TabIndex        =   334
      Top             =   840
      Width           =   10215
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "TIPO"
         Height          =   285
         Index           =   169
         Left            =   5640
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   341
         Top             =   3600
         Width           =   375
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "COD_PROCEDIMIENTO"
         Height          =   285
         Index           =   165
         Left            =   480
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   337
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "NOM_PROCEDIMIENTO"
         Height          =   885
         Index           =   166
         Left            =   480
         Locked          =   -1  'True
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   336
         Top             =   1920
         Width           =   9615
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "GALENAJE"
         Height          =   285
         Index           =   167
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   335
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo"
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
         Index           =   62
         Left            =   5160
         TabIndex        =   342
         Top             =   3600
         Width           =   390
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Codigo de Procedimiento "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   255
         Index           =   61
         Left            =   480
         TabIndex        =   340
         Top             =   720
         Width           =   2085
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Nombre del Procedimiento"
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
         Index           =   38
         Left            =   480
         TabIndex        =   339
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Galenaje"
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
         Index           =   10
         Left            =   480
         TabIndex        =   338
         Top             =   3480
         Width           =   750
      End
   End
   Begin VB.Frame Frame16Servicios 
      Height          =   6015
      Left            =   240
      TabIndex        =   287
      Top             =   720
      Width           =   10335
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "CODSERVICIO"
         Height          =   285
         Index           =   101
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   306
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "TIPOSER"
         Height          =   285
         Index           =   102
         Left            =   2160
         MaxLength       =   3
         TabIndex        =   305
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "CODSER"
         Height          =   285
         Index           =   103
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   304
         Top             =   1440
         Width           =   3735
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "NOMSERVICIO"
         Height          =   285
         Index           =   104
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   303
         Top             =   1800
         Width           =   6135
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "COD_PROCEDIMIENTO"
         Height          =   285
         Index           =   105
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   302
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "NOMANALC"
         Height          =   285
         Index           =   106
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   301
         Top             =   2520
         Width           =   6135
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "NIVEL"
         Height          =   285
         Index           =   107
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   300
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "UNIMED"
         Height          =   285
         Index           =   108
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   299
         Top             =   3240
         Width           =   3735
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "CODANALS"
         Height          =   285
         Index           =   109
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   298
         Top             =   3600
         Width           =   1215
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "CODTAB"
         Height          =   285
         Index           =   110
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   297
         Top             =   3960
         Width           =   3735
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "CODAREA"
         Height          =   285
         Index           =   113
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   296
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "MATERIAL"
         Height          =   285
         Index           =   111
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   295
         Top             =   4320
         Width           =   3735
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "TIEMP_ENTN"
         Height          =   285
         Index           =   114
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   294
         Top             =   4680
         Width           =   1215
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "TIEMPO_ENTU"
         Height          =   285
         Index           =   115
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   293
         Top             =   5040
         Width           =   1215
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "METODO"
         Height          =   285
         Index           =   112
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   292
         Top             =   5400
         Width           =   3735
      End
      Begin VB.TextBox TXTFIELDS 
         DataField       =   "TITULO"
         Height          =   285
         Index           =   116
         Left            =   6960
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   291
         Top             =   5400
         Width           =   255
      End
      Begin VB.CommandButton CmdAyudaTipoSer 
         Height          =   380
         Left            =   2640
         Picture         =   "MANTENIMIENTO.frx":A1F3
         Style           =   1  'Graphical
         TabIndex        =   290
         ToolTipText     =   "Ayuda de Tipo de Servicio"
         Top             =   720
         Width           =   540
      End
      Begin VB.CommandButton CmdAyudaArea 
         Height          =   380
         Left            =   2640
         Picture         =   "MANTENIMIENTO.frx":A635
         Style           =   1  'Graphical
         TabIndex        =   289
         ToolTipText     =   "Ayuda de Tipo de Servicio"
         Top             =   1080
         Width           =   540
      End
      Begin VB.CommandButton CmdAyudaProcedimiento 
         Height          =   380
         Left            =   3120
         Picture         =   "MANTENIMIENTO.frx":AA77
         Style           =   1  'Graphical
         TabIndex        =   288
         ToolTipText     =   "Ayuda de Tipo de Servicio"
         Top             =   2160
         Width           =   540
      End
      Begin VB.Label LBLServiArea 
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
         Index           =   0
         Left            =   3120
         TabIndex        =   325
         Top             =   1080
         Width           =   2085
      End
      Begin VB.Label LBLServiProce 
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
         Left            =   3600
         TabIndex        =   324
         Top             =   2160
         Width           =   2085
      End
      Begin VB.Label LBLServiTiSer 
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
         Index           =   0
         Left            =   3120
         TabIndex        =   323
         Top             =   720
         Width           =   2085
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Método"
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
         Index           =   76
         Left            =   1560
         TabIndex        =   322
         Top             =   5400
         Width           =   630
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Titulo"
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
         Index           =   77
         Left            =   6360
         TabIndex        =   321
         Top             =   5400
         Width           =   510
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Código del Servicio"
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
         Index           =   6
         Left            =   480
         TabIndex        =   320
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo de Servicio"
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
         Index           =   7
         Left            =   720
         TabIndex        =   319
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Código Servicio"
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
         Index           =   9
         Left            =   840
         TabIndex        =   318
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Nombre Servicio"
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
         Index           =   17
         Left            =   600
         TabIndex        =   317
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Código de Procedimiento"
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
         Index           =   56
         Left            =   120
         TabIndex        =   316
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Nombre  Analisis"
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
         Index           =   57
         Left            =   720
         TabIndex        =   315
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Nivel"
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
         Index           =   58
         Left            =   1680
         TabIndex        =   314
         Top             =   2880
         Width           =   510
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Unidad de Medida"
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
         Index           =   60
         Left            =   720
         TabIndex        =   313
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Código de Análisis "
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
         Index           =   64
         Left            =   600
         TabIndex        =   312
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Código Tab"
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
         Index           =   65
         Left            =   1200
         TabIndex        =   311
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Código de Area"
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
         Index           =   66
         Left            =   720
         TabIndex        =   310
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Material"
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
         Index           =   68
         Left            =   1440
         TabIndex        =   309
         Top             =   4320
         Width           =   750
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Tiempo Ent."
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
         Index           =   74
         Left            =   1200
         TabIndex        =   308
         Top             =   4680
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Tiempo"
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
         Index           =   75
         Left            =   1560
         TabIndex        =   307
         Top             =   5040
         Width           =   630
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   10545
      Left            =   0
      TabIndex        =   227
      Top             =   0
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   18600
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      MultiSelect     =   -1  'True
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   20
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "PERSONAS"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "EMPRESAS"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "LINEAS"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "INSUMOS"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "DOCTOR"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "AREA"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ESPECIALIDAD"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "VALORES NORMALES"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "INSTITUTOS"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "PROCEDIMIENTOS"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab11 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "USUARIOS"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab12 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "PROVEEDOR"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab13 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "PARAMETROS"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab14 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "TABLA DE VALORES"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab15 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "HABITACIONES"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab16 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "SERVICIOS"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab17 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "TIPO DE SERVICIOS"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab18 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ARANCELES"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab19 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "MEDICAMENTOS"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab20 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ITEM HISTORIA CLI"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "MANTENIMIENTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Control As Integer
Dim bUltfoco As Boolean
Dim bPrifoco As Boolean
Dim bActualizar As Boolean
Dim VALIDAR As ValTecla
Dim mvBookMark  As Variant
Dim mbDataChanged  As Boolean
Dim mbAddNewFlag As Boolean
Dim mbEditFlag As Boolean
Dim blCancel As Boolean
Dim WithEvents CtrVisualizar As clsValidar
Attribute CtrVisualizar.VB_VarHelpID = -1
Dim WithEvents adoPrimaryRS As ADODB.Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim WithEvents adoManCiuDoc As ADODB.Recordset
Attribute adoManCiuDoc.VB_VarHelpID = -1
Dim WithEvents adoManCiuIns As ADODB.Recordset
Attribute adoManCiuIns.VB_VarHelpID = -1
Dim WithEvents adoManEspecialidad As ADODB.Recordset
Attribute adoManEspecialidad.VB_VarHelpID = -1
Dim WithEvents adoManDoctor As ADODB.Recordset
Attribute adoManDoctor.VB_VarHelpID = -1
Dim WithEvents adoManPersona As ADODB.Recordset
Attribute adoManPersona.VB_VarHelpID = -1
Dim WithEvents adoManCiuProv As ADODB.Recordset
Attribute adoManCiuProv.VB_VarHelpID = -1
Dim WithEvents adoManProIns As ADODB.Recordset
Attribute adoManProIns.VB_VarHelpID = -1
Dim WithEvents adoManLinIns As ADODB.Recordset
Attribute adoManLinIns.VB_VarHelpID = -1
Dim WithEvents adoManNomUsu As ADODB.Recordset
Attribute adoManNomUsu.VB_VarHelpID = -1
Dim WithEvents adoManTipHab As ADODB.Recordset
Attribute adoManTipHab.VB_VarHelpID = -1
Dim WithEvents adoManTipoSer As ADODB.Recordset
Attribute adoManTipoSer.VB_VarHelpID = -1
Dim WithEvents adoManCiuEmp As ADODB.Recordset
Attribute adoManCiuEmp.VB_VarHelpID = -1
Dim WithEvents adoServiProce As ADODB.Recordset
Attribute adoServiProce.VB_VarHelpID = -1
Dim WithEvents adoServiArea As ADODB.Recordset
Attribute adoServiArea.VB_VarHelpID = -1
Dim WithEvents adoValNorGen As ADODB.Recordset
Attribute adoValNorGen.VB_VarHelpID = -1
Dim WithEvents adoSerAran As ADODB.Recordset
Attribute adoSerAran.VB_VarHelpID = -1
Dim WithEvents adoAranDoc As ADODB.Recordset
Attribute adoAranDoc.VB_VarHelpID = -1
Dim WithEvents adoAranIns As ADODB.Recordset
Attribute adoAranIns.VB_VarHelpID = -1
Dim WithEvents adoAranTiser As ADODB.Recordset
Attribute adoAranTiser.VB_VarHelpID = -1
Dim WithEvents adoManEspeItem As ADODB.Recordset
Attribute adoManEspeItem.VB_VarHelpID = -1
Dim WithEvents adoManTipAnam As ADODB.Recordset
Attribute adoManTipAnam.VB_VarHelpID = -1


Private Sub CmdAgregar_Click()
BLOQUEO (False)
Cnn.BeginTrans
prcCaso5y14Agregar (True)
Select Case Me.TabStrip1.SelectedItem.Index

Case 1
    Me.TXTFIELDS(0).SetFocus
    
Case 2
    
     Me.TXTFIELDS(120).Locked = True
  

Case 3
    Me.TXTFIELDS(140).Locked = True
    Me.TXTFIELDS(141).SetFocus
Case 4
    Me.TXTFIELDS(53).SetFocus
    
Case 5

    Me.TXTFIELDS(22).Locked = True
    Me.TXTFIELDS(24).Locked = True
    Me.TXTFIELDS(25).Locked = True
    
Case 6
    
    Me.TXTFIELDS(36).SetFocus
Case 7

    Me.TXTFIELDS(38).Locked = True
    Me.TXTFIELDS(39).SetFocus
Case 8
    cmdAyudaAnalisis.Visible = True
    Me.cmdAyudaAnalisis.SetFocus
Case 9
    Me.TXTFIELDS(82).Locked = True
    Me.TXTFIELDS(83).SetFocus
    
Case 10
    Me.TXTFIELDS(165).Locked = True
    Me.TXTFIELDS(166).SetFocus

Case 11
     Me.TXTFIELDS(60).Locked = True
    Me.TXTFIELDS(61).SetFocus
    LBLManUsu(1) = ""
Case 12
    Me.TXTFIELDS(64).Locked = True
    Me.TXTFIELDS(65).SetFocus
Case 13
    TXTFIELDS(30).Enabled = True
    TXTFIELDS(31).Enabled = True
    Me.TXTFIELDS(30).SetFocus
     
Case 14
    
    TXTFIELDS(79).Enabled = True
    TXTFIELDS(80).Enabled = True
    Me.TXTFIELDS(79).SetFocus
Case 15
    'Me.TXTFIELDS(26).SetFocus
    
Case 16
     TXTFIELDS(101).Locked = True
     Me.TXTFIELDS(102).SetFocus
     
Case 17
     TXTFIELDS(75).Enabled = True
     Me.TXTFIELDS(75).SetFocus
     
Case 18
     TXTFIELDS(130).Locked = True
     TXTFIELDS(131).Locked = True
     TXTFIELDS(130).SetFocus
    
     
Case 19

      Me.TXTFIELDS(150).Locked = True
      Me.TXTFIELDS(151).SetFocus

Case 20
     


End Select
 'On Error Resume Next
  
  
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    
    End If
    .AddNew
   End With
     
   
   Set PICPER.Picture = Nothing
    Select Case Me.TabStrip1.SelectedItem.Index
    
    Case 2
       Me.TXTFIELDS(120) = OPTENER_CODIGO("SELECT COD_EMPRESA FROM EMPRESA ORDER BY COD_EMPRESA DESC")
    
    Case 3
       Me.TXTFIELDS(140) = OPTENER_CODIGO("SELECT COD_LINEA FROM LINEAS ORDER BY COD_LINEA DESC")
       
    Case 4
       Me.TXTFIELDS(54) = OPTENER_CODIGO("SELECT COD_INSUMO FROM INSUMOS ORDER BY COD_INSUMO DESC")
        
    Case 5
        Me.TXTFIELDS(22) = OPTENER_CODIGO("SELECT CODDR FROM DOCTOR ORDER BY CODDR DESC")
           LBLManDoc(0) = ""
           LBLManDoc(1) = ""
           LBLManDoc(2) = ""
        '  LBLManDoc(3) = ""
           LBLManDoc(4) = ""
           LBLManDoc(5) = ""
           LBLManDoc(6) = ""
           LBLManDoc(7) = ""
           LBLManDoc(8) = ""
           LBLManDoc(9) = ""
           LBLManDoc(10) = ""
           LBLManDoc(11) = ""
           LBLManDoc(12) = ""
    Case 7
        Me.TXTFIELDS(38) = OPTENER_CODIGO("SELECT COD_ESPECIALIDAD FROM ESPECIALIDAD  ORDER BY COD_ESPECIALIDAD DESC")
    Case 9
        Me.TXTFIELDS(82) = OPTENER_CODIGO("SELECT CODINS FROM INSTITUTOS ORDER BY CODINS DESC")
    
    Case 10
      
        Me.TXTFIELDS(165) = OPTENER_CODIGO("SELECT COD_PROCEDIMIENTO FROM PROCEDIMIENTOS ORDER BY COD_PROCEDIMIENTO DESC")
       
      
    Case 11
        Me.TXTFIELDS(60) = OPTENER_CODIGO("SELECT ID_USUARIO FROM USUARIOS ORDER BY ID_USUARIO DESC")
        
    Case 12
       Me.TXTFIELDS(64) = OPTENER_CODIGO("SELECT COD_PROVEEDOR FROM PROVEEDORES ORDER BY COD_PROVEEDOR DESC")
    'Case 13
       ' Me.TXTFIELDS(77) = OPTENER_CODIGO("SELECT CODCIUDAD FROM CIUDADES ORDER BY CODCIUDAD DESC")
        
    Case 16
    
       Me.TXTFIELDS(101) = OPTENER_CODIGO("SELECT CODSERVICIO FROM SERVICIO ORDER BY CODSERVICIO DESC")
       
       
    Case 19
       Me.TXTFIELDS(150) = OPTENER_CODIGO("SELECT ID_MEDICAMENTO FROM MEDICAMENTOS ORDER BY ID_MEDICAMENTO DESC")
        
   End Select
  
  Exit Sub

End Sub
Private Sub cmdBuscar_Click()
Dim frmAyuda As New frmConsult
'Dialog.Show
    Select Case Me.TabStrip1.SelectedItem.Index
         Case 1 'personas
                 
            Set frmAyuda.ActiveConection = Cnn
            frmAyuda.Modal = True
            frmAyuda.Consulta = "SELECT * FROM personas"
            frmAyuda.Show vbModal
            If Not frmAyuda.Descargado Then
            
                    With adoPrimaryRS
                            .MoveFirst
                            .Find "tipdoc='" & frmAyuda.msGrid.Columns(0).Value & "'"
                            .Find "nrodoc=" & frmAyuda.msGrid.Columns(1).Value
                    End With
                       On Error Resume Next ' La direccion tiene que estar en el path
                       Set PICPER.Picture = LoadPicture(adoPrimaryRS.Fields("FOTO").Value, vbLPLarge, vbLPColor)

           End If
           frmAyuda.Cerrar
           
        Case 2  'Empresas
            
           Set frmAyuda.ActiveConection = Cnn
           frmAyuda.Modal = True
           frmAyuda.Consulta = "SELECT * FROM EMPRESA"
           frmAyuda.Show vbModal
           If Not frmAyuda.Descargado Then
                With adoPrimaryRS
                        .MoveFirst
                        .Find "COD_EMPRESA=" & frmAyuda.msGrid.Columns(0).Value
                 End With
           End If
           frmAyuda.Cerrar
        
        Case 3   'LINEAS
         
           Set frmAyuda.ActiveConection = Cnn
           frmAyuda.Modal = True
           frmAyuda.Consulta = "SELECT * FROM LINEAS"
           frmAyuda.Show vbModal
           If Not frmAyuda.Descargado Then
                With adoPrimaryRS
                        .MoveFirst
                        .Find "COD_LINEA=" & frmAyuda.msGrid.Columns(0).Value
                 End With
           End If
           frmAyuda.Cerrar
        
        Case 4
         
           Set frmAyuda.ActiveConection = Cnn
           frmAyuda.Modal = True
           frmAyuda.Consulta = "SELECT * FROM INSUMOS"
           frmAyuda.Show vbModal
           If Not frmAyuda.Descargado Then
                With adoPrimaryRS
                        .MoveFirst
                        .Find "COD_INSUMO=" & frmAyuda.msGrid.Columns(0).Value
                End With
          End If
          frmAyuda.Cerrar
        
        Case 5
         
           Set frmAyuda.ActiveConection = Cnn
           frmAyuda.Modal = True
           frmAyuda.Consulta = "SELECT * FROM DOCTOR"
           frmAyuda.Show vbModal
           If Not frmAyuda.Descargado Then
                With adoPrimaryRS
                        .MoveFirst
                        .Find "CODDR=" & frmAyuda.msGrid.Columns(0).Value
                End With
         
           End If
           frmAyuda.Cerrar
        
        Case 6
         
          Set frmAyuda.ActiveConection = Cnn
          frmAyuda.Modal = True
          frmAyuda.Consulta = "SELECT * FROM AREA"
          frmAyuda.Show vbModal
          If Not frmAyuda.Descargado Then
                With adoPrimaryRS
                        .MoveFirst
                        .Find "codarea='" & frmAyuda.msGrid.Columns(0).Value & "'"
                End With
         
          End If
          frmAyuda.Cerrar
        
        Case 7
         
         Set frmAyuda.ActiveConection = Cnn
         frmAyuda.Modal = True
         frmAyuda.Consulta = "SELECT * FROM ESPECIALIDAD"
         frmAyuda.Show vbModal
         If Not frmAyuda.Descargado Then
                With adoPrimaryRS
                        .MoveFirst
                        .Find "COD_ESPECIALIDAD=" & frmAyuda.msGrid.Columns(0).Value
                End With
         End If
         frmAyuda.Cerrar
        
        Case 8
         
         Set frmAyuda.ActiveConection = Cnn
         frmAyuda.Modal = True
         frmAyuda.Consulta = "SELECT * FROM VALNORM"
         frmAyuda.Show vbModal
         If Not frmAyuda.Descargado Then
                With adoPrimaryRS
                        .MoveFirst
                        .Find "CODSERVICIO=" & frmAyuda.msGrid.Columns(0).Value
                End With
         End If
         frmAyuda.Cerrar
        
        Case 9
         
         Set frmAyuda.ActiveConection = Cnn
         frmAyuda.Modal = True
         frmAyuda.Consulta = "SELECT * FROM INSTITUTOS"
         frmAyuda.Show vbModal
         If Not frmAyuda.Descargado Then
                With adoPrimaryRS
                        .MoveFirst
                        .Find "CODINS=" & frmAyuda.msGrid.Columns(0).Value
                End With
         
         End If
         frmAyuda.Cerrar
        
       
        Case 10
         
         Set frmAyuda.ActiveConection = Cnn
         frmAyuda.Modal = True
         frmAyuda.Consulta = "SELECT * FROM PROCEDIMIENTOS"
         frmAyuda.Show vbModal
         If Not frmAyuda.Descargado Then
                With adoPrimaryRS
                        .MoveFirst
                        .Find "COD_PROCEDIMIENTO=" & frmAyuda.msGrid.Columns(0).Value
                End With
         
         End If
         frmAyuda.Cerrar
        
        
       
       
        
        Case 11
         
         Set frmAyuda.ActiveConection = Cnn
         frmAyuda.Modal = True
         frmAyuda.Consulta = "SELECT * FROM USUARIOS"
         frmAyuda.Show vbModal
         If Not frmAyuda.Descargado Then
                With adoPrimaryRS
                        .MoveFirst
                        .Find "ID_USUARIO=" & frmAyuda.msGrid.Columns(0).Value
                End With
         End If
         frmAyuda.Cerrar
        
        Case 12
          Set frmAyuda.ActiveConection = Cnn
          frmAyuda.Modal = True
          frmAyuda.Consulta = "SELECT * FROM PROVEEDORES"
          frmAyuda.Show vbModal
          If Not frmAyuda.Descargado Then
                With adoPrimaryRS
                        .MoveFirst
                        .Find "COD_PROVEEDOR=" & frmAyuda.msGrid.Columns(0).Value
                End With
          End If
          frmAyuda.Cerrar
        
       Case 13 'Parametros
                 
         Set frmAyuda.ActiveConection = Cnn
         frmAyuda.Modal = True
         frmAyuda.Consulta = "SELECT * FROM PARAMETROS"
         frmAyuda.Show vbModal
         If Not frmAyuda.Descargado Then
            
                    With adoPrimaryRS
                            .MoveFirst
                            .Find "tipo='" & frmAyuda.msGrid.Columns(0).Value & "'"
                            .Find "codigo='" & frmAyuda.msGrid.Columns(1).Value & "'"
                           
                    End With
           
        End If
        frmAyuda.Cerrar
       
        
      Case 14
          Set frmAyuda.ActiveConection = Cnn
          frmAyuda.Modal = True
          frmAyuda.Consulta = "SELECT * FROM TABVAL"
          frmAyuda.Show vbModal
          If Not frmAyuda.Descargado Then
                With adoPrimaryRS
                        .MoveFirst
                        .Find "CODTAB='" & frmAyuda.msGrid.Columns(0).Value & "'"
                        .Find "CODITEM=" & frmAyuda.msGrid.Columns(1).Value
                                           
                End With
          End If
          frmAyuda.Cerrar
          
          
     Case 15
          Set frmAyuda.ActiveConection = Cnn
          frmAyuda.Modal = True
          frmAyuda.Consulta = "SELECT * FROM HABITACIONES"
          frmAyuda.Show vbModal
          If Not frmAyuda.Descargado Then
                With adoPrimaryRS
                        .MoveFirst
                        .Find "COD_TIPO='" & frmAyuda.msGrid.Columns(0).Value & "'"
                        .Find "COD_HABITACION=" & frmAyuda.msGrid.Columns(1).Value
                                           
                End With
          End If
          frmAyuda.Cerrar
          
     Case 16
          Set frmAyuda.ActiveConection = Cnn
          frmAyuda.Modal = True
          frmAyuda.Consulta = "SELECT * FROM SERVICIO"
          frmAyuda.Show vbModal
          If Not frmAyuda.Descargado Then
                With adoPrimaryRS
                        .MoveFirst
                        .Find "CODSERVICIO='" & frmAyuda.msGrid.Columns(0).Value & "'"
                        .Find "TIPOSER='" & frmAyuda.msGrid.Columns(1).Value & "'"
                                           
                End With
          End If
          frmAyuda.Cerrar
          
          
    Case 17
      
          Set frmAyuda.ActiveConection = Cnn
          frmAyuda.Modal = True
          frmAyuda.Consulta = "SELECT * FROM TIPOSERVICIO"
          frmAyuda.Show vbModal
          If Not frmAyuda.Descargado Then
                With adoPrimaryRS
                        .MoveFirst
                        .Find "TIPOSER='" & frmAyuda.msGrid.Columns(0).Value & "'"
                        .Find "NOMTIPOSER='" & frmAyuda.msGrid.Columns(1).Value & "'"
                                           
                End With
          End If
          frmAyuda.Cerrar
        
        
     Case 18
      
          Set frmAyuda.ActiveConection = Cnn
          frmAyuda.Modal = True
          frmAyuda.Consulta = "SELECT * FROM ARANCELES"
          frmAyuda.Show vbModal
          If Not frmAyuda.Descargado Then
                With adoPrimaryRS
                        .MoveFirst
                        .Find "CODSERVICIO='" & frmAyuda.msGrid.Columns(0).Value & "'"
                        .Find "TIPOSER='" & frmAyuda.msGrid.Columns(1).Value & "'"
                        .Find "CODDR='" & frmAyuda.msGrid.Columns(2).Value & "'"
                        .Find "CODINS='" & frmAyuda.msGrid.Columns(3).Value & "'"
                                           
                                           
                End With
          End If
          frmAyuda.Cerrar
           
      Case 19
      
          Set frmAyuda.ActiveConection = Cnn
          frmAyuda.Modal = True
          frmAyuda.Consulta = "SELECT * FROM MEDICAMENTOS"
          frmAyuda.Show vbModal
          If Not frmAyuda.Descargado Then
                With adoPrimaryRS
                        .MoveFirst
                        .Find "ID_MEDICAMENTO='" & frmAyuda.msGrid.Columns(0).Value & "'"
                        
                                           
                End With
          End If
          frmAyuda.Cerrar
          
          
     Case 20
          
          Set frmAyuda.ActiveConection = Cnn
          frmAyuda.Modal = True
          frmAyuda.Consulta = "SELECT * FROM ITEM_HISTORIACLINICA_ESPECIALI"
          frmAyuda.Show vbModal
          If Not frmAyuda.Descargado Then
                With adoPrimaryRS
                        .MoveFirst
                        .Find "COD_ESPECIALIDAD=" & frmAyuda.msGrid.Columns(0).Value
                        .Find "TIPO='" & frmAyuda.msGrid.Columns(1).Value & "'"
                        .Find "ITEM='" & frmAyuda.msGrid.Columns(2).Value & "'"
                End With
          End If
          frmAyuda.Cerrar
         
          
        
        
          
    End Select
    Set frmAyuda = Nothing
End Sub


Private Sub boxfechan_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub


Private Sub busficha_Click()
        Dim Consulta As frmConsult
        Set Consulta = New frmConsult
        Consulta.Modal = True
        Consulta.Caption = "Consulta de TIPO DE DOCUMENTO"
        Set Consulta.ActiveConection = Cnn
        Consulta.Consulta = "SELECT TIPDOC,NRODOC,NOMBRES,APELLIDOS,RUC,MOVIL,E_MAIL,TELEFONOS,DOM_PARTICULAR,COD_CIUDADP,DOM_LABORAL,COD_CIUDADL,RADIO FROM PERSONAS WHERE TITULO='DR' and NOT EXISTS (SELECT tipdoc ,nrodoc  From Doctor where tipdoc=personas.tipdoc and nrodoc=personas.nrodoc)"
        Consulta.Show vbModal
        If Not Consulta.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
           
           With Consulta.msGrid
                
                adoPrimaryRS(1).Value = .Columns(0).Text
                adoPrimaryRS(2).Value = .Columns(1).Text
                LBLManDoc(0).Caption = .Columns(3).Text & ", " & .Columns(2).Text
                LBLManDoc(4).Caption = .Columns(4).Text
                LBLManDoc(5).Caption = .Columns(6).Text
                LBLManDoc(6).Caption = .Columns(7).Text
                LBLManDoc(7).Caption = .Columns(5).Text
                LBLManDoc(8).Caption = .Columns(8).Text
                LBLManDoc(9).Caption = .Columns(9).Text
                LBLManDoc(10).Caption = .Columns(10).Text
                LBLManDoc(11).Caption = .Columns(11).Text
                LBLManDoc(12).Caption = .Columns(12).Text
                
                
          End With
                    
           Consulta.Cerrar
           TXTFIELDS(23).SetFocus
        End If
        
         
  '       Exit Sub
'GoPrevError:
 ' MsgBox Err.Description
        
   
 End Sub
 
Private Sub CmdAranIns_Click()
  Dim frmAyuda As New frmConsult
 Set frmAyuda.ActiveConection = Cnn
 frmAyuda.Modal = True
 frmAyuda.Consulta = "SELECT CODINS,INSTITUTO,TELEF1,TELEF2 FROM institutos"
 frmAyuda.Show vbModal
 If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
    With frmAyuda.msGrid
        adoPrimaryRS.Fields("CODINS").Value = .Columns(0).Value
        LBLInsti(1).Caption = .Columns(1).Text
    End With
 End If
 frmAyuda.Cerrar
  ' bustipomov.SetFocus
End Sub

Private Sub cmdAyudaAnalisis_Click()
        Dim Consulta As frmConsult
        Set Consulta = New frmConsult
        Consulta.Modal = True
        Consulta.Caption = "Consulta de Valores Normales"
        'Consulta.StartUpPosition = 2
        'SetWindowPos conTablas.hwnd, conHwndTopmost, 120, 72, 500, 320, &H50
        Set Consulta.ActiveConection = Cnn
        Consulta.Consulta = "select * from valnorm"
        Consulta.Show vbModal
        If Not Consulta.Descargado Then
            TXTFIELDS(40).Text = Consulta.Seleccion
            Consulta.Cerrar
        End If
        TXTFIELDS(41).SetFocus

End Sub

Private Sub CmdAyudaArea_Click()
  Dim Consulta As frmConsult
        Set Consulta = New frmConsult
        Consulta.Modal = True
        Consulta.Caption = "Consulta de AREAS MEDICAS"
        Set Consulta.ActiveConection = Cnn
        Consulta.Consulta = "select * from AREA "
        Consulta.Show vbModal
        If Not Consulta.Descargado Then
            adoPrimaryRS.Fields("CODAREA").Value = Trim$(Consulta.Seleccion)
            LBLServiArea(0).Caption = Consulta.msGrid.Columns(1).Value
           
            adoPrimaryRS.Fields("CODSER").Value = adoPrimaryRS.Fields("CODAREA").Value & Trim$(TXTFIELDS(101).Text)
            Consulta.Cerrar
        End If
        TXTFIELDS(111).SetFocus
End Sub


Private Sub CmdAyudaCargo_Click()
        Dim Consulta As frmConsult
        Set Consulta = New frmConsult
        Consulta.Modal = True
        Consulta.Caption = "Consulta de Cargo"
        Set Consulta.ActiveConection = Cnn
        Consulta.Consulta = "select CODIGO,DESCRIPCION from PARAMETROS WHERE TIPO='CARGO'"
        Consulta.Show vbModal
        If Not Consulta.Descargado Then
            adoPrimaryRS.Fields("CARGO").Value = Trim$(Consulta.Seleccion)
            'TXTFIELDS(12) = Trim$(Consulta.Seleccion)
            Consulta.Cerrar
        End If
        TXTFIELDS(13).SetFocus
End Sub

Private Sub CmdAyudaCiud_Click()
       Dim Consulta As frmConsult
        Set Consulta = New frmConsult
        Consulta.Modal = True
        Consulta.Caption = "Consulta de CIUDAD"
        Set Consulta.ActiveConection = Cnn
        Consulta.Consulta = "select CODIGO,DESCRIPCION from PARAMETROS WHERE TIPO='CIUDAD'"
        Consulta.Show vbModal
        If Not Consulta.Descargado Then
            adoPrimaryRS.Fields("COD_CIUDAD").Value = Trim$(Consulta.Seleccion)
            'TXTFIELDS(14) = Trim$(Consulta.Seleccion)
            Consulta.Cerrar
        End If
        TXTFIELDS(124).SetFocus
End Sub

Private Sub cmdayudaciudadlab_Click()
        Dim Consulta As frmConsult
        Set Consulta = New frmConsult
        Consulta.Modal = True
        Consulta.Caption = "Consulta de CIUDAD"
        Set Consulta.ActiveConection = Cnn
        Consulta.Consulta = "select CODIGO,DESCRIPCION from PARAMETROS WHERE TIPO='CIUDAD'"
        Consulta.Show vbModal
        If Not Consulta.Descargado Then
            adoPrimaryRS.Fields("COD_CIUDADL").Value = Trim$(Consulta.Seleccion)
            'TXTFIELDS(14) = Trim$(Consulta.Seleccion)
            Consulta.Cerrar
        End If
        TXTFIELDS(15).SetFocus
End Sub

Private Sub cmdayudaciudadpar_Click()
        Dim Consulta As frmConsult
        Set Consulta = New frmConsult
        Consulta.Modal = True
        Consulta.Caption = "Consulta de CIUDAD"
        Set Consulta.ActiveConection = Cnn
        Consulta.Consulta = "select CODIGO,DESCRIPCION from PARAMETROS WHERE TIPO='CIUDAD'"
        Consulta.Show vbModal
        If Not Consulta.Descargado Then
            adoPrimaryRS.Fields("COD_CIUDADP").Value = Trim$(Consulta.Seleccion)
            'TXTFIELDS(10) = Trim$(Consulta.Seleccion)
            Consulta.Cerrar
        End If
        TXTFIELDS(11).SetFocus
End Sub


Private Sub cmdayudaecivil_Click()
  Dim Consulta As frmConsult
        Set Consulta = New frmConsult
        Consulta.Modal = True
        Consulta.Caption = "Consulta de ESTADO CIVIL"
        Set Consulta.ActiveConection = Cnn
        Consulta.Consulta = "select CODIGO,DESCRIPCION from PARAMETROS WHERE TIPO='ESTADO'"
        Consulta.Show vbModal
        If Not Consulta.Descargado Then
            adoPrimaryRS.Fields("ESTCIV").Value = Trim$(Consulta.Seleccion)
            'TXTFIELDS(7) = Trim$(Consulta.Seleccion)
            Consulta.Cerrar
        End If
        TXTFIELDS(8).SetFocus
End Sub

Private Sub CmdAyudaEspe_Click()
      
        Dim Consulta As frmConsult
        Set Consulta = New frmConsult
        Consulta.Modal = True
        Set Consulta.ActiveConection = Cnn
        Consulta.Consulta = "select * from ESPECIALIDAD"
        Consulta.Show vbModal
        If Not Consulta.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
           With Consulta.msGrid
              If Not adoManEspeItem.EOF Then
                  ' "CODIGO='" & TXTFIELDS(41).Text & "'"
                   adoPrimaryRS.Fields("COD_ESPECIALIDAD").Value = Trim$(Consulta.Seleccion)
                   'LBLManEspIte(1).Caption = adoManEspeItem.Fields("NOMBRE").Value
              End If
           End With
         
           Consulta.Cerrar
        End If
        TXTFIELDS(171).SetFocus
End Sub

Private Sub cmdAyudaEspec_Click()
  If TXTFIELDS(24) <> "" Then
    
       On Error GoTo GoPrevError
       Dim Consulta As frmConsult
       Set Consulta = New frmConsult
       Set Consulta.ActiveConection = Cnn
       Consulta.Modal = True
       Consulta.Caption = "Consulta de Especialidad"
       Consulta.Consulta = "select * from especialidad"
       Consulta.Show vbModal
       If Not Consulta.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
          With Consulta.msGrid
            If Not adoManEspecialidad.EOF Then
              adoManEspecialidad.MoveFirst
              adoManEspecialidad.Find "CODDR= " & Me.TXTFIELDS(22).Text
              adoManEspecialidad.Find "COD_ESPECIALIDAD=" & .Columns(0).Text
            End If
           If adoManEspecialidad.EOF Then
              adoManEspecialidad.AddNew
              adoManEspecialidad("CODDR").Value = Me.TXTFIELDS(22).Text
              adoManEspecialidad("COD_ESPECIALIDAD").Value = .Columns(0).Value
              adoManEspecialidad("NOMBRE").Value = .Columns(1).Value
              Set grd_Especialidad.DataSource = adoManEspecialidad '("ChildCMD").UnderlyingValue
              grd_Especialidad.Columns(2).Visible = False
           Else
             MsgBox ("La Especialidad ya existe para el Código de Doctor")
           End If
      
        End With
    
   End If
 
   Consulta.Cerrar
  
End If
  Exit Sub
GoPrevError:
  MsgBox Err.Description
End Sub
Private Sub cmdayudagenero_Click()
        Dim Consulta As frmConsult
        Set Consulta = New frmConsult
        Consulta.Modal = True
        Consulta.Caption = "Consulta de GENERO"
        Set Consulta.ActiveConection = Cnn
        Consulta.Consulta = "select CODIGO,DESCRIPCION from PARAMETROS WHERE TIPO='GENERO'"
        Consulta.Show vbModal
        If Not Consulta.Descargado Then
            adoPrimaryRS.Fields("GENERO").Value = Trim$(Consulta.Seleccion)
            'TXTFIELDS(6).Text = Trim$(Consulta.Seleccion)
            Consulta.Cerrar
        End If
        TXTFIELDS(7).SetFocus
End Sub

Private Sub CmdAyudaHabitacion_Click()
       Dim Consulta As frmConsult
        Set Consulta = New frmConsult
        Consulta.Modal = True
        Consulta.Caption = "Consulta de TIPO DE Habitación"
        Set Consulta.ActiveConection = Cnn
        Consulta.Consulta = "select CODIGO,DESCRIPCION from PARAMETROS WHERE TIPO='CAMA' "
        
       ' AND NOT EXISTS (SELECT COD_HABITACION,COD_CAMA FROM HABITACIONES WHERE COD_HABITACION=PARAMETROS.  "
       ' Consulta.Consulta = "SELECT TIPDOC,NRODOC,NOMBRES FROM PERSONAS WHERE TITULO='DR' and NOT EXISTS (SELECT tipdoc ,nrodoc  From Doctor where tipdoc=personas.tipdoc and nrodoc=personas.nrodoc)"
        
        Consulta.Show vbModal
        If Not Consulta.Descargado Then
            adoPrimaryRS.Fields("COD_CAMA").Value = Trim$(Consulta.Seleccion)
            adoPrimaryRS.Fields("COD_HABITACION").Value = Trim$(Consulta.Seleccion)
            Consulta.Cerrar
        End If
         'CmdEstadoCama.SetFocus
         TXTFIELDS(93).SetFocus
End Sub

Private Sub CmdAyudaProcedimiento_Click()
       Dim Consulta As frmConsult
       Set Consulta = New frmConsult
       Consulta.Modal = True
       Consulta.Caption = "Consulta de PROCEDIMIENTOS"
       Set Consulta.ActiveConection = Cnn
       Consulta.Consulta = "select * from PROCEDIMIENTOS "
       Consulta.Show vbModal
       If Not Consulta.Descargado Then
           adoPrimaryRS.Fields("COD_PROCEDIMIENTO").Value = Trim$(Consulta.Seleccion)
           LBLServiProce(1).Caption = Consulta.msGrid.Columns(1).Value
           Consulta.Cerrar
       End If
       TXTFIELDS(106).SetFocus
End Sub

Private Sub CmdAyudaTipoHab_Click()
   Dim Consulta As frmConsult
        Set Consulta = New frmConsult
        Consulta.Modal = True
        Consulta.Caption = "Consulta de TIPO DE Habitación"
        Set Consulta.ActiveConection = Cnn
        Consulta.Consulta = "select CODIGO,DESCRIPCION from PARAMETROS WHERE TIPO='TIPOHABITA'"
        Consulta.Show vbModal
        If Not Consulta.Descargado Then
            adoPrimaryRS.Fields("COD_TIPO").Value = Trim$(Consulta.Seleccion)
            
            Consulta.Cerrar
        End If
         'CmdAyudaHabitacion.SetFocus
         TXTFIELDS(91).SetFocus
End Sub

Private Sub CmdAyudaTipoPerso_Click()
        Dim Consulta As frmConsult
        Set Consulta = New frmConsult
        Consulta.Modal = True
        Consulta.Caption = "Consulta de TIPO DE Habitación"
        Set Consulta.ActiveConection = Cnn
        Consulta.Consulta = "select CODIGO,DESCRIPCION from PARAMETROS WHERE TIPO='TIPOPERS'"
        Consulta.Show vbModal
        If Not Consulta.Descargado Then
            adoPrimaryRS.Fields("TIPOPERSONA").Value = Trim$(Consulta.Seleccion)
            'TXTFIELDS(11) = Trim$(Consulta.Seleccion)
            Consulta.Cerrar
        End If
        TXTFIELDS(12).SetFocus
End Sub

Private Sub CmdAyudaTipoSer_Click()
        Dim Consulta As frmConsult
        Set Consulta = New frmConsult
        Consulta.Modal = True
        Consulta.Caption = "Consulta de TIPO DE SERVICIO"
        Set Consulta.ActiveConection = Cnn
        Consulta.Consulta = "select * from TIPOSERVICIO "
        Consulta.Show vbModal
        If Not Consulta.Descargado Then
            adoPrimaryRS.Fields("TIPOSER").Value = Trim$(Consulta.Seleccion)
             LBLServiTiSer(0).Caption = Consulta.msGrid.Columns(1).Value
             
           ' TXTFIELDS(5) = Trim$(Consulta.Seleccion)
            Consulta.Cerrar
        End If
        TXTFIELDS(113).SetFocus
End Sub

Private Sub Cmdayudatitulo_Click()
        Dim Consulta As frmConsult
        Set Consulta = New frmConsult
        Consulta.Modal = True
        Consulta.Caption = "Consulta de TITULO"
        Set Consulta.ActiveConection = Cnn
        Consulta.Consulta = "select CODIGO,DESCRIPCION from PARAMETROS WHERE TIPO='TITULO'"
        Consulta.Show vbModal
        If Not Consulta.Descargado Then
            adoPrimaryRS.Fields("TITULO").Value = Trim$(Consulta.Seleccion)
           ' TXTFIELDS(5) = Trim$(Consulta.Seleccion)
            Consulta.Cerrar
        End If
        TXTFIELDS(5).SetFocus
End Sub
Private Sub cmdayudaInst_Click()
        Dim Consulta As frmConsult
        Set Consulta = New frmConsult
        Consulta.Modal = True
        Consulta.Caption = "Consulta de Ciudades"
        Set Consulta.ActiveConection = Cnn
        Consulta.Consulta = "select CODIGO,DESCRIPCION from PARAMETROS WHERE TIPO='CIUDAD'"
        Consulta.Show vbModal
        If Not Consulta.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
            With Consulta.msGrid
               If Not adoManCiuIns.EOF Then
                  adoPrimaryRS.Fields("CODCIU").Value = Trim$(Consulta.Seleccion)
                  LBLManCiuIns(5).Caption = adoManCiuIns.Fields("Descripcion").Value
               End If
            End With
           Consulta.Cerrar
        End If
        TXTFIELDS(86).SetFocus
End Sub
Private Sub cmdAyudaLinea_Click()
        Dim Consulta As frmConsult
        Set Consulta = New frmConsult
        Consulta.Modal = True
        Set Consulta.ActiveConection = Cnn
        Consulta.Consulta = "select * from lineas"
        Consulta.Show vbModal
        If Not Consulta.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
           With Consulta.msGrid
              If Not adoManLinIns.EOF Then
                 '  adoManLinIns.Filter = "COD_LINEA=" & txtfields(57).Text
                   adoPrimaryRS.Fields("COD_LINEA").Value = Trim$(Consulta.Seleccion)
                   LBLManLiIns(1).Caption = adoManLinIns.Fields("LIN_NOMBRE").Value
              End If
           End With
         
            Consulta.Cerrar
        End If
        TXTFIELDS(51).SetFocus
        
End Sub
Private Sub cmdAyudaProv_Click()
        Dim Consulta As frmConsult
        Set Consulta = New frmConsult
        Consulta.Modal = True
        Consulta.Caption = "Consulta de Proveedores"
        Set Consulta.ActiveConection = Cnn
        Consulta.Consulta = "select * from proveedores"
        Consulta.Show vbModal
        If Not Consulta.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
           With Consulta.msGrid
              If Not adoManProIns.EOF Then
                 adoPrimaryRS.Fields("COD_PROVEEDOR").Value = Trim$(Consulta.Seleccion)
                 LBLManProIns(1).Caption = adoManProIns.Fields("NOM_PROVEEDOR").Value
              End If
           End With
           Consulta.Cerrar
        End If
    
        TXTFIELDS(57).SetFocus
        
End Sub
Public Sub CmdCancelar_Click()
 On Error Resume Next
  BLOQUEO (True)
  'adoPrimaryRS.CancelBatch
  Cnn.RollbackTrans
  adoPrimaryRS.CancelUpdate
  adoPrimaryRS.Requery
  
  Select Case Me.TabStrip1.SelectedItem.Index
      Case 5
         adoManEspecialidad.Requery
  End Select
  
  
  'adoPrimaryRS.MoveFirst
  mbDataChanged = False
  CtrVisualizar_ABMtollbar CtrVisualizar.PERMISOABM
  CtrVisualizar_BUSCAR CtrVisualizar.enable
Exit Sub

RefreshErr:
  MsgBox Err.Description
End Sub
Private Sub cmdAnterior_Click()
  On Error GoTo GoPrevError
  If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious
  If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
    'ha sobrepasado el final; vuelva atrás
    adoPrimaryRS.MoveFirst
  End If
  'muestra el registro actual
  mbDataChanged = False
  Exit Sub
GoPrevError:
  MsgBox Err.Description
End Sub
Private Sub cmdSiguiente_Click()
  '  On Error GoTo GoNextError
    If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext
    
    If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
        Beep
         'ha sobrepasado el final; vuelva atrás
        adoPrimaryRS.MoveLast
      
    End If
    'muestra el registro actual
    mbDataChanged = False
    Exit Sub

'GoNextError:
 ' MsgBox Err.Description

End Sub
Private Sub cmdPrimero_Click()
    On Error GoTo GoFirstError
    adoPrimaryRS.MoveFirst
    mbDataChanged = False
    Exit Sub
GoFirstError:
  MsgBox Err.Description
End Sub
Private Sub cmdModificar_Click()
    BLOQUEO (False)
    prcCaso5y14Agregar (False)
    On Error GoTo EditErr
    Select Case Me.TabStrip1.SelectedItem.Index
    Case 1
        Me.TXTFIELDS(1).SetFocus
        
    Case 2
        
        Me.TXTFIELDS(120).Locked = True
        Me.TXTFIELDS(121).SetFocus
   
    Case 3
    
        Me.TXTFIELDS(140).Locked = True
        Me.TXTFIELDS(141).SetFocus
    Case 4
        Me.TXTFIELDS(54).Locked = True
        Me.TXTFIELDS(53).SetFocus
    Case 5
        Me.TXTFIELDS(22).Locked = True
        Me.TXTFIELDS(23).SetFocus
    Case 6
        Me.TXTFIELDS(36).Locked = True
        Me.TXTFIELDS(37).SetFocus
    Case 7
        Me.TXTFIELDS(38).Locked = True
        Me.TXTFIELDS(39).SetFocus
    Case 8
        Me.TXTFIELDS(40).Locked = True
        Me.TXTFIELDS(41).SetFocus
    Case 9
        
        Me.TXTFIELDS(82).Locked = True
        Me.TXTFIELDS(83).SetFocus
        
    Case 10
    
        Me.TXTFIELDS(165).Locked = True
        Me.TXTFIELDS(166).SetFocus
        
    Case 11
        
        Me.TXTFIELDS(60).Locked = True
        Me.TXTFIELDS(61).SetFocus
    Case 12
        
        Me.TXTFIELDS(64).Locked = True
        Me.TXTFIELDS(65).SetFocus
    Case 13
          
        TXTFIELDS(30).Locked = True
        TXTFIELDS(31).Locked = True
        Me.TXTFIELDS(32).SetFocus
        
    Case 14
        TXTFIELDS(79).Locked = True
        TXTFIELDS(80).Locked = True
        Me.TXTFIELDS(81).SetFocus
        
    Case 16
    
        TXTFIELDS(101).Locked = True
        TXTFIELDS(102).Locked = True
        Me.TXTFIELDS(102).SetFocus
        
    Case 17
    
        TXTFIELDS(75).Locked = True
        TXTFIELDS(76).Enabled = True
        Me.TXTFIELDS(76).SetFocus
        
    Case 18
    
        TXTFIELDS(130).Enabled = True
        TXTFIELDS(131).Enabled = True
        TXTFIELDS(132).SetFocus
        
    Case 19
    
        Me.TXTFIELDS(150).Locked = True
        Me.TXTFIELDS(151).SetFocus
        
    Case 20
    
        Me.TXTFIELDS(170).Locked = True
        Me.TXTFIELDS(171).Locked = True
        Me.TXTFIELDS(172).SetFocus
    
  
  
End Select
Cnn.BeginTrans
Exit Sub
EditErr:
   MsgBox Err.Description
End Sub
Private Sub cmdUltimo_Click()
    On Error GoTo GoLastError
    adoPrimaryRS.MoveLast
    mbDataChanged = False
    Exit Sub
GoLastError:
  MsgBox Err.Description
End Sub
Private Sub Cmdayudatipdoc_Click()
        Dim Consulta As frmConsult
        Set Consulta = New frmConsult
        Consulta.Modal = True
        Consulta.Caption = "Consulta de TIPO DE DOCUMENTO"
        Set Consulta.ActiveConection = Cnn
        Consulta.Consulta = "select CODIGO,DESCRIPCION from PARAMETROS WHERE TIPO='TIPODOC'"
        Consulta.Show vbModal
        If Not Consulta.Descargado Then
            adoPrimaryRS.Fields("TIPDOC").Value = Trim$(Consulta.Seleccion)
            'TXTFIELDS(0) = Trim$(Consulta.Seleccion)
            Consulta.Cerrar
        End If
        TXTFIELDS(1).SetFocus
End Sub

Private Sub CmdAyuGene_Click()
    Dim Consulta As frmConsult
    Set Consulta = New frmConsult
    Consulta.Modal = True
    Consulta.Caption = "Consulta de GENERO"
    Set Consulta.ActiveConection = Cnn
    Consulta.Consulta = "select CODIGO,DESCRIPCION from PARAMETROS WHERE TIPO='GENERO'"
    Consulta.Show vbModal
    If Not Consulta.Descargado Then
         adoPrimaryRS.Fields("SEXO").Value = Trim$(Consulta.Seleccion)
          LBLManGENE(1).Caption = Consulta.msGrid.Columns(1).Value
         'TXTFIELDS(41).Text = Trim$(Consulta.Seleccion)
         Consulta.Cerrar
    End If
    TXTFIELDS(42).SetFocus
End Sub

Private Sub CmdAyuSer_Click()
    Dim Consulta As frmConsult
    Set Consulta = New frmConsult
    Consulta.Modal = True
    Consulta.Caption = "Consulta de SERVICIO"
    Set Consulta.ActiveConection = Cnn
    Consulta.Consulta = "select  CODSERVICIO,TIPOSER,CODSER,NOMSERVICIO from SERVICIO "
    
    Consulta.Show vbModal
    If Not Consulta.Descargado Then
         adoPrimaryRS.Fields("CODSERVICIO").Value = Trim$(Consulta.Seleccion)
         adoPrimaryRS.Fields("TIPOSER").Value = Consulta.msGrid.Columns(1).Value
         LBLManAraSer(1).Caption = Consulta.msGrid.Columns(3).Value
         
         adoAranTiser.MoveFirst
         adoAranTiser.Filter = "TIPOSER='" & Trim(TXTFIELDS(131).Text) & "'"
         LBLManAraTiSer(0).Caption = adoAranTiser.Fields("NOMTIPOSER").Value
                  
         ' TXTFIELDS(5) = Trim$(Consulta.Seleccion)
         Consulta.Cerrar
    End If
    'TXTFIELDS(103).SetFocus
    
    
End Sub

Private Sub Command2_Click()
  Dim Consulta As frmConsult
        Set Consulta = New frmConsult
        Consulta.Modal = True
        Consulta.Caption = "Consulta de Proveedores"
        Set Consulta.ActiveConection = Cnn
        Consulta.Consulta = "select * from proveedores"
        Consulta.Show vbModal
        If Not Consulta.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
           With Consulta.msGrid
              If Not adoManProIns.EOF Then
                 adoPrimaryRS.Fields("COD_PROVEEDOR").Value = Trim$(Consulta.Seleccion)
                 LBLManProIns(1).Caption = adoManProIns.Fields("NOM_PROVEEDOR").Value
              End If
           End With
           Consulta.Cerrar
        End If
    
        TXTFIELDS(57).SetFocus

End Sub
Private Sub CmdEliminar_Click()
Dim Respuesta As Integer
 On Error GoTo GoLastError
 Respuesta = MsgBox("Esta Seguro que desea Eliminar el Registro", vbQuestion + vbYesNo, "Atencion")
If Respuesta = vbYes Then
    adoPrimaryRS.Delete
    adoPrimaryRS.MoveNext
    If adoPrimaryRS.EOF Then
        adoPrimaryRS.MoveFirst
    End If
    adoPrimaryRS.UpdateBatch
    adoPrimaryRS.Update
End If
Exit Sub
GoLastError:
Respuesta = MsgBox(Err.Description, vbExclamation, "Atencion")
adoPrimaryRS.CancelBatch
'adoPrimaryRS.CancelUpdate
End Sub
Private Sub cmdAyudaProvee_Click()
        Dim Consulta As frmConsult
        Set Consulta = New frmConsult
        Consulta.Modal = True
        Consulta.Caption = "Consulta de Ciudades"
        Set Consulta.ActiveConection = Cnn
        Consulta.Consulta = "select CODIGO,DESCRIPCION from PARAMETROS WHERE TIPO='CIUDAD'"
        Consulta.Show vbModal
        If Not Consulta.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
            With Consulta.msGrid
               If Not adoManCiuProv.EOF Then
                   adoPrimaryRS.Fields("CIUDAD").Value = Trim$(Consulta.Seleccion)
                   LBLManCiuProv(0).Caption = adoManCiuProv.Fields("Descripcion").Value
               End If
            End With
           Consulta.Cerrar
        End If
        TXTFIELDS(69).SetFocus
        
End Sub

Private Sub cmdbusdoc_Click()
  Dim frmAyuda As New frmConsult
  Set frmAyuda.ActiveConection = Cnn
  frmAyuda.Modal = True
  frmAyuda.Consulta = "SELECT DOCTOR.CODDR,PERSONAS.APELLIDOS,PERSONAS.NOMBRES, PERSONAS.TIPDOC,PERSONAS.NRODOC,DOCTOR.REGPRF FROM DOCTOR  LEFT OUTER JOIN PERSONAS ON DOCTOR.TIPDOC = PERSONAS.TIPDOC AND DOCTOR.NRODOC = PERSONAS.NRODOC"
  frmAyuda.Show vbModal
 If Not frmAyuda.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
    With frmAyuda.msGrid
        adoPrimaryRS.Fields("CODDR").Value = .Columns(0).Value
        LBLAraDoc(0).Caption = .Columns(1).Text & ", " & .Columns(2).Text
                      
    End With
 End If
frmAyuda.Cerrar
'businstitut.SetFocus
End Sub

Private Sub CmdBusUsua_Click()
   Dim Consulta As frmConsult
        Set Consulta = New frmConsult
        Consulta.Modal = True
        Consulta.Caption = "Consulta de TIPO DE DOCUMENTO"
        Set Consulta.ActiveConection = Cnn
        Consulta.Consulta = "SELECT TIPDOC,NRODOC,NOMBRES,APELLIDOS FROM PERSONAS WHERE TIPOPERSONA='001' and NOT EXISTS (SELECT tipdoc ,nrodoc  From USUARIOS  where tipdoc=personas.tipdoc and nrodoc=personas.nrodoc)"
        Consulta.Show vbModal
        If Not Consulta.Descargado And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
           
           With Consulta.msGrid
                adoPrimaryRS(0).Value = TXTFIELDS(60)
                adoPrimaryRS(1).Value = .Columns(0).Text
                adoPrimaryRS(2).Value = .Columns(1).Text
                
                LBLManUsu(1).Caption = .Columns(3).Text & ", " & .Columns(2).Text
                               
          End With
                    
           Consulta.Cerrar
           TXTFIELDS(63).SetFocus
        End If
End Sub

Private Sub CmdEstadoCama_Click()
    Dim Consulta As frmConsult
        Set Consulta = New frmConsult
        Consulta.Modal = True
        Consulta.Caption = "Consulta de Estado Cama"
        Set Consulta.ActiveConection = Cnn
        Consulta.Consulta = "select CODIGO,DESCRIPCION from PARAMETROS WHERE TIPO='ESTADOCAMA'"
        Consulta.Show vbModal
        If Not Consulta.Descargado Then
            adoPrimaryRS.Fields("ESTADO").Value = Trim$(Consulta.Seleccion)
            Consulta.Cerrar
        End If
        TXTFIELDS(94).SetFocus
End Sub

Private Sub cmdFicha_Click()
 Call FICHACLINICA.PASARIDENTIDAD(adoPrimaryRS.Fields("TIPDOC").Value, adoPrimaryRS.Fields("NRODOC").Value)
 FICHACLINICA.Show vbModal

End Sub

Private Sub CMDFOTOPERS_Click()
   
   If MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
     Me.Comdiapers.ShowOpen
     Set PICPER.Picture = LoadPicture(Me.Comdiapers.FileName, vbLPLarge, vbLPColor)
     adoPrimaryRS.Fields("FOTO").Value = Me.Comdiapers.FileName
   End If
  
End Sub

Private Sub Command1_Click()
       Dim Consulta As frmConsult
        Set Consulta = New frmConsult
        Consulta.Modal = True
        Consulta.Caption = "Consulta de TITULO"
        Set Consulta.ActiveConection = Cnn
        Consulta.Consulta = "select CODIGO,DESCRIPCION from PARAMETROS WHERE TIPO='TITULO'"
        Consulta.Show vbModal
        If Not Consulta.Descargado Then
            adoPrimaryRS.Fields("TITULO").Value = Trim$(Consulta.Seleccion)
           ' TXTFIELDS(5) = Trim$(Consulta.Seleccion)
            Consulta.Cerrar
        End If
        TXTFIELDS(5).SetFocus
End Sub

Private Sub CmdTipoAnamn_Click()
        Dim Consulta As frmConsult
        Set Consulta = New frmConsult
        Consulta.Modal = True
        Consulta.Caption = "Consulta de Tipo "
        Set Consulta.ActiveConection = Cnn
        Consulta.Consulta = "select CODIGO,DESCRIPCION from PARAMETROS WHERE TIPO='ANAMNESIS'"
        Consulta.Show vbModal
        If Not Consulta.Descargado Then
           With Consulta.msGrid
              If Not adoManTipAnam.EOF Then
                 adoPrimaryRS.Fields("TIPO").Value = Trim$(Consulta.Seleccion)
                ' TXTFIELDS(5) = Trim$(Consulta.Seleccion)
                 LblManIteTipo(0).Caption = .Columns(1).Text
                 Consulta.Cerrar
              End If
          End With
        End If
        TXTFIELDS(172).SetFocus
    
        
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

Private Sub CtrVisualizar_BUSCAR(enable As Boolean)
    MDIGALENO.tbToolBar.Buttons("Search").Enabled = enable
End Sub

Private Sub Form_Activate()
    Call TabStrip1_Click
End Sub

Private Sub Form_GotFocus()
    CtrVisualizar_ABMtollbar CtrVisualizar.PERMISOABM
End Sub

Private Sub Form_Load()

    Dim i As Integer
    Set adoPrimaryRS = New ADODB.Recordset
    Set CtrVisualizar = New clsValidar
    Set adoManCiuDoc = New ADODB.Recordset
    Set adoManEspecialidad = New ADODB.Recordset
    Set adoManDoctor = New ADODB.Recordset
    Set adoManCiuIns = New ADODB.Recordset
    Set adoManPersona = New ADODB.Recordset
    Set adoManCiuProv = New ADODB.Recordset
    Set adoManProIns = New ADODB.Recordset
    Set adoManLinIns = New ADODB.Recordset
    Set adoManNomUsu = New ADODB.Recordset
    Set adoManTipHab = New ADODB.Recordset
    Set adoManTipoSer = New ADODB.Recordset
    Set adoManCiuEmp = New ADODB.Recordset
    Set adoServiProce = New ADODB.Recordset
    Set adoServiArea = New ADODB.Recordset
    Set adoValNorGen = New ADODB.Recordset
    Set adoSerAran = New ADODB.Recordset
    Set adoAranDoc = New ADODB.Recordset
    Set adoAranIns = New ADODB.Recordset
    Set adoAranTiser = New ADODB.Recordset
    Set adoManEspeItem = New ADODB.Recordset
    Set adoManTipAnam = New ADODB.Recordset

          
    Me.Height = 7500
    Me.Left = 500
    Me.Top = 0
    Set VALIDAR = New ValTecla
    Me.Width = 11000
    Frame1PERSONA.Visible = CtrVisualizar.Visualizar(1, Cnn, IdUsuario)
    
 End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    On Error GoTo errMan
    Dim intRespuesta As Integer
    'hacer el codigo para controlar si es que el recorset se debe actualizar
    If bActualizar Then  'verifica SI ES QUE EL RECORDSET NECESITA ACTUALIZARSE
                         'ANTES DE CERRAR EL FORMULARIO
        intRespuesta = MsgBox("DESEA GUARDAR LOS CAMBIOS REALIZADOS", vbYesNoCancel, vbInformation)
        If intRespuesta = vbNo Then
            On Error Resume Next
            adoPrimaryRS.CancelUpdate
        Else
            If intRespuesta = vbYes Then
                adoPrimaryRS.Update
            Else
                Cancel = True
            End If
         End If
    End If
Exit Sub
errMan:
        Cancel = True
        MsgBox Err.Description
        bActualizar = True
        
End Sub
Private Sub cmdGrabar_Click()
   ' On Error GoTo UpdateErr
    BLOQUEO (True)
    
    SendKeys "{TAB}", True
    adoPrimaryRS.UpdateBatch
    Select Case Me.TabStrip1.SelectedItem.Index
      Case 5
         adoManEspecialidad.UpdateBatch
    End Select
        Cnn.CommitTrans
        If TabStrip1.SelectedItem.Index = 11 Then
            adoPrimaryRS.Requery
        End If
    Exit Sub
'UpdateErr:
 '    MsgBox Err.Description
End Sub

Private Sub LBLManDoc_Change(Index As Integer)
If LBLManDoc(9).Caption = "" Then
    LBLManDoc(1).Caption = ""
Else
    adoManCiuDoc.MoveFirst
    adoManCiuDoc.Find "CODIGO='" & Trim$(LBLManDoc(9).Caption) & "'"
    If Not adoManCiuDoc.EOF And Not adoManCiuDoc.BOF Then
        LBLManDoc(1).Caption = adoManCiuDoc.Fields("Descripcion").Value
    End If
End If
If LBLManDoc(11).Caption = "" Then
     LBLManDoc(2).Caption = ""
Else
   adoManCiuDoc.MoveFirst
   adoManCiuDoc.Find "CODIGO=" & Trim$(LBLManDoc(11).Caption)
   If Not adoManCiuDoc.EOF And Not adoManCiuDoc.BOF Then
        LBLManDoc(2).Caption = adoManCiuDoc.Fields("Descripcion").Value
   End If
End If
 
End Sub

Private Sub TabStrip1_Click()
prcCaso5y14Agregar (False)
Select Case Me.TabStrip1.SelectedItem.Index
Case 1
            On Error Resume Next
            adoPrimaryRS.CancelUpdate
            
            Frame1PERSONA.Visible = CtrVisualizar.Visualizar(1, Cnn, IdUsuario)
            Frame2Empresas.Visible = False
            Frame3PROVEEDOR.Visible = False
            Frame4LINEA.Visible = False
            frame5INSUMO.Visible = False
            Frame6DOCTOR.Visible = False
            Frame7AREA.Visible = False
            Frame8ESPECIALIDAD.Visible = False
            Frame9PARAMETROS.Visible = False
            Frame10VALORNORMAL.Visible = False
            Frame11INSTITUTO.Visible = False
            Frame12Procedimientos.Visible = False
            Frame13USUARIO.Visible = False
            Frame14TABLAVALOR.Visible = False
            Frame15Habitaciones.Visible = False
            Frame16Servicios.Visible = False
            Frame17TipoServicio.Visible = False
            Frame18Aranceles.Visible = False
            Frame19Medicamentos.Visible = False
            Frame20Item.Visible = False
            
            adoPrimaryRS.Close
            adoPrimaryRS.Open "select * from PERSONAS ORDER BY APELLIDOS", Cnn, adOpenStatic, adLockBatchOptimistic, adCmdText
            Dim oText As TextBox
            'Enlaza los cuadros de texto con el proveedor de datos
            For Each oText In Me.TXTFIELDS
                If (oText.Index >= 0 And oText.Index <= 20) Then
                    Set oText.DataSource = adoPrimaryRS
                End If
            Next
            
            'Set Me.boxfechan.DataSource = adoPrimaryRS
            'Set Me.PICPER.DataSource = adoPrimaryRS
            Set PICPER.Picture = LoadPicture(adoPrimaryRS.Fields("FOTO").Value, vbLPLarge, vbLPColor)
            mbDataChanged = False
            Cmdayudatipdoc.SetFocus
            
            
            
 Case 2
 
            On Error Resume Next
            adoPrimaryRS.CancelUpdate
            
            Frame2Empresas.Visible = CtrVisualizar.Visualizar(2, Cnn, IdUsuario)
            Frame1PERSONA.Visible = False
            Frame3PROVEEDOR.Visible = False
            Frame4LINEA.Visible = False
            frame5INSUMO.Visible = False
            Frame6DOCTOR.Visible = False
            Frame7AREA.Visible = False
            Frame8ESPECIALIDAD.Visible = False
            Frame9PARAMETROS.Visible = False
            Frame10VALORNORMAL.Visible = False
            Frame11INSTITUTO.Visible = False
            Frame12Procedimientos.Visible = False
            Frame13USUARIO.Visible = False
            Frame14TABLAVALOR.Visible = False
            Frame15Habitaciones.Visible = False
            Frame16Servicios.Visible = False
            Frame17TipoServicio.Visible = False
            Frame18Aranceles.Visible = False
            Frame19Medicamentos.Visible = False
            Frame20Item.Visible = False
            
            adoPrimaryRS.Close
            adoPrimaryRS.Open "select * FROM EMPRESA", Cnn, adOpenStatic, adLockBatchOptimistic, adCmdText
            
            adoManCiuEmp.Open "Select CODIGO,DESCRIPCION FROM PARAMETROS", Cnn
              
            'Dim oText As TextBox
            'Enlaza los cuadros de texto con el proveedor de datos
            For Each oText In Me.TXTFIELDS
                If (oText.Index > 119 And oText.Index < 126) Then
                    Set oText.DataSource = adoPrimaryRS
                End If
            Next
            
                                   
           'adoManCiuEmp.Find "CODIGO=" & TXTFIELDS(123).Text
           'LBLManCiuEmp(0).Caption = adoManCiuEmp.Fields("Descripcion").Value
            
Case 3
            
            On Error Resume Next
            adoPrimaryRS.CancelUpdate
            
            Frame4LINEA.Visible = CtrVisualizar.Visualizar(3, Cnn, IdUsuario)
            Frame1PERSONA.Visible = False
            Frame2Empresas.Visible = False
            Frame3PROVEEDOR.Visible = False
            frame5INSUMO.Visible = False
            Frame6DOCTOR.Visible = False
            Frame7AREA.Visible = False
            Frame8ESPECIALIDAD.Visible = False
            Frame9PARAMETROS.Visible = False
            Frame10VALORNORMAL.Visible = False
            Frame11INSTITUTO.Visible = False
            Frame12Procedimientos.Visible = False
            Frame13USUARIO.Visible = False
            Frame14TABLAVALOR.Visible = False
            Frame15Habitaciones.Visible = False
            Frame16Servicios.Visible = False
            Frame17TipoServicio.Visible = False
            Frame18Aranceles.Visible = False
            Frame19Medicamentos.Visible = False
            Frame20Item.Visible = False
            
            adoPrimaryRS.Close
            adoPrimaryRS.Open "select COD_LINEA,LIN_NOMBRE,DESCRIPCION from LINEAS Order by COD_LINEA", Cnn, adOpenStatic, adLockBatchOptimistic, adCmdText
            'Dim oText As TextBox
            'Enlaza los cuadros de texto con el proveedor de datos
            For Each oText In Me.TXTFIELDS
                If oText.Index > 139 And oText.Index < 143 Then
                    Set oText.DataSource = adoPrimaryRS
                End If
            Next
            mbDataChanged = False
Case 4
            
            On Error Resume Next
            adoPrimaryRS.CancelUpdate
            
            frame5INSUMO.Visible = CtrVisualizar.Visualizar(4, Cnn, IdUsuario)
            Frame1PERSONA.Visible = False
            Frame2Empresas.Visible = False
            Frame3PROVEEDOR.Visible = False
            Frame4LINEA.Visible = False
            Frame6DOCTOR.Visible = False
            Frame7AREA.Visible = False
            Frame8ESPECIALIDAD.Visible = False
            Frame9PARAMETROS.Visible = False
            Frame10VALORNORMAL.Visible = False
            Frame11INSTITUTO.Visible = False
            Frame12Procedimientos.Visible = False
            Frame13USUARIO.Visible = False
            Frame14TABLAVALOR.Visible = False
            Frame15Habitaciones.Visible = False
            Frame16Servicios.Visible = False
            Frame17TipoServicio.Visible = False
            Frame18Aranceles.Visible = False
            Frame19Medicamentos.Visible = False
            Frame20Item.Visible = False
            
            cmdAyudaLinea.Visible = False
            cmdAyudaProv.Visible = False
            
            adoPrimaryRS.Close
            adoPrimaryRS.Open "select * FROM INSUMOS", Cnn, adOpenStatic, adLockBatchOptimistic, adCmdText
            adoManProIns.Open "select * FROM PROVEEDORES", Cnn
            adoManLinIns.Open "Select * from LINEAS", Cnn
            'Dim oText As TextBox
            'Enlaza los cuadros de texto con el proveedor de datos
            For Each oText In Me.TXTFIELDS
                If oText.Index > 50 And oText.Index < 60 Then
                  Set oText.DataSource = adoPrimaryRS
                End If
            Next
            
            Set chkFields(9).DataSource = adoPrimaryRS
                 
Case 5
            On Error Resume Next
            adoPrimaryRS.CancelUpdate
            
            Frame6DOCTOR.Visible = CtrVisualizar.Visualizar(5, Cnn, IdUsuario)
            Frame1PERSONA.Visible = False
            Frame2Empresas.Visible = False
            Frame3PROVEEDOR.Visible = False
            Frame4LINEA.Visible = False
            frame5INSUMO.Visible = False
            Frame7AREA.Visible = False
            Frame8ESPECIALIDAD.Visible = False
            Frame9PARAMETROS.Visible = False
            Frame10VALORNORMAL.Visible = False
            Frame11INSTITUTO.Visible = False
            Frame12Procedimientos.Visible = False
            Frame13USUARIO.Visible = False
            Frame14TABLAVALOR.Visible = False
            Frame15Habitaciones.Visible = False
            Frame16Servicios.Visible = False
            Frame17TipoServicio.Visible = False
            Frame18Aranceles.Visible = False
            Frame19Medicamentos.Visible = False
            Frame20Item.Visible = False
            
            cmdAyudaDocCiu1.Visible = False
            cmdAyudaDocCiu2.Visible = False
            
            adoPrimaryRS.Close
            adoPrimaryRS.Open "select * FROM DOCTOR", Cnn, adOpenStatic, adLockBatchOptimistic
            adoManCiuDoc.Open "Select CODIGO,DESCRIPCION FROM PARAMETROS", Cnn
            adoManEspecialidad.Open "Select  COD_ESPECIALIDAD,NOMBRE,CODDR FROM DOCTORESPECIALIDAD", Cnn, adOpenStatic, adLockBatchOptimistic, adCmdText
            adoManDoctor.Open "Select DOCTOR.CODDR,DOCTOR.TIPDOC,DOCTOR.NRODOC,PERSONAS.NOMBRES,PERSONAS.APELLIDOS,PERSONAS.RUC,PERSONAS.E_MAIL,PERSONAS.TELEFONOS,PERSONAS.MOVIL,PERSONAS.DOM_PARTICULAR,PERSONAS.COD_CIUDADP,PERSONAS.DOM_LABORAL,PERSONAS.COD_CIUDADL,PERSONAS.RADIO FROM DOCTOR,PERSONAS WHERE DOCTOR.TIPDOC=PERSONAS.TIPDOC AND DOCTOR.NRODOC=PERSONAS.NRODOC ", Cnn
            adoManPersona.Open "select * from personas where titulo='DR'", Cnn
            Dim i As Integer
            'Dim oText As TextBox
            'Enlaza los cuadros de texto con el proveedor de datos
            For Each oText In Me.TXTFIELDS
                If oText.Index > 21 And oText.Index < 27 Then
                    Set oText.DataSource = adoPrimaryRS
                End If
            Next
            
            Set grd_Especialidad.DataSource = adoManEspecialidad
            grd_Especialidad.Columns(2).Visible = False
            
            For i = 3 To 13
                Set LBLManDoc(i).DataSource = adoManDoctor
            Next

Case 6
            
            On Error Resume Next
            adoPrimaryRS.CancelUpdate
            
            Frame7AREA.Visible = CtrVisualizar.Visualizar(6, Cnn, IdUsuario)
            Frame1PERSONA.Visible = False
            Frame2Empresas.Visible = False
            Frame3PROVEEDOR.Visible = False
            Frame4LINEA.Visible = False
            frame5INSUMO.Visible = False
            Frame6DOCTOR.Visible = False
            Frame8ESPECIALIDAD.Visible = False
            Frame9PARAMETROS.Visible = False
            Frame10VALORNORMAL.Visible = False
            Frame11INSTITUTO.Visible = False
            Frame12Procedimientos.Visible = False
            Frame13USUARIO.Visible = False
            Frame14TABLAVALOR.Visible = False
            Frame15Habitaciones.Visible = False
            Frame16Servicios.Visible = False
            Frame17TipoServicio.Visible = False
            Frame18Aranceles.Visible = False
            Frame19Medicamentos.Visible = False
            Frame20Item.Visible = False
                        
            adoPrimaryRS.Close
            adoPrimaryRS.Open "select * FROM AREA", Cnn, adOpenStatic, adLockBatchOptimistic, adCmdText
            'Dim oText As TextBox
            'Enlaza los cuadros de texto con el proveedor de datos
            For Each oText In Me.TXTFIELDS
                If oText.Index > 35 And oText.Index < 38 Then
                    Set oText.DataSource = adoPrimaryRS
                End If
            Next
Case 7
            On Error Resume Next
            adoPrimaryRS.CancelUpdate
            
            Frame8ESPECIALIDAD.Visible = CtrVisualizar.Visualizar(7, Cnn, IdUsuario)
            Frame1PERSONA.Visible = False
            Frame2Empresas.Visible = False
            Frame3PROVEEDOR.Visible = False
            Frame4LINEA.Visible = False
            frame5INSUMO.Visible = False
            Frame6DOCTOR.Visible = False
            Frame7AREA.Visible = False
            Frame9PARAMETROS.Visible = False
            Frame10VALORNORMAL.Visible = False
            Frame11INSTITUTO.Visible = False
            Frame12Procedimientos.Visible = False
            Frame13USUARIO.Visible = False
            Frame14TABLAVALOR.Visible = False
            Frame15Habitaciones.Visible = False
            Frame16Servicios.Visible = False
            Frame17TipoServicio.Visible = False
            Frame18Aranceles.Visible = False
            Frame19Medicamentos.Visible = False
            Frame20Item.Visible = False
                       
            adoPrimaryRS.Close
            adoPrimaryRS.Open "select * FROM ESPECIALIDAD", Cnn, adOpenStatic, adLockBatchOptimistic, adCmdText
            'Dim oText As TextBox
            'Enlaza los cuadros de texto con el proveedor de datos
            For Each oText In Me.TXTFIELDS
                If oText.Index > 37 And oText.Index < 40 Then
                    Set oText.DataSource = adoPrimaryRS
                End If
            Next
            
Case 8
            On Error Resume Next
            adoPrimaryRS.CancelUpdate
            
            Frame10VALORNORMAL.Visible = CtrVisualizar.Visualizar(8, Cnn, IdUsuario)
            Frame1PERSONA.Visible = False
            Frame2Empresas.Visible = False
            Frame3PROVEEDOR.Visible = False
            Frame4LINEA.Visible = False
            frame5INSUMO.Visible = False
            Frame6DOCTOR.Visible = False
            Frame7AREA.Visible = False
            Frame8ESPECIALIDAD.Visible = False
            Frame9PARAMETROS.Visible = False
            Frame11INSTITUTO.Visible = False
            Frame12Procedimientos.Visible = False
            Frame13USUARIO.Visible = False
            Frame14TABLAVALOR.Visible = False
            Frame15Habitaciones.Visible = False
            Frame16Servicios.Visible = False
            Frame17TipoServicio.Visible = False
            Frame18Aranceles.Visible = False
            Frame19Medicamentos.Visible = False
            Frame20Item.Visible = False
                        
            cmdAyudaAnalisis.Visible = False
            
            adoPrimaryRS.Close
            adoPrimaryRS.Open "select * FROM VALNORM", Cnn, adOpenStatic, adLockOptimistic, adCmdText
            adoValNorGen.Open "Select CODIGO,DESCRIPCION FROM PARAMETROS", Cnn
            
            'Dim oText As TextBox
            'Enlaza los cuadros de texto con el proveedor de datos
            For Each oText In Me.TXTFIELDS
                If oText.Index > 39 And oText.Index < 55 Then
                    Set oText.DataSource = adoPrimaryRS
                End If
            Next

Case 9
            On Error Resume Next
            adoPrimaryRS.CancelUpdate
            
            Frame11INSTITUTO.Visible = CtrVisualizar.Visualizar(9, Cnn, IdUsuario)
            Frame1PERSONA.Visible = False
            Frame2Empresas.Visible = False
            Frame3PROVEEDOR.Visible = False
            Frame4LINEA.Visible = False
            frame5INSUMO.Visible = False
            Frame6DOCTOR.Visible = False
            Frame7AREA.Visible = False
            Frame8ESPECIALIDAD.Visible = False
            Frame9PARAMETROS.Visible = False
            Frame10VALORNORMAL.Visible = False
            Frame12Procedimientos.Visible = False
            Frame13USUARIO.Visible = False
            Frame14TABLAVALOR.Visible = False
            Frame15Habitaciones.Visible = False
            Frame16Servicios.Visible = False
            Frame17TipoServicio.Visible = False
            Frame18Aranceles.Visible = False
            Frame19Medicamentos.Visible = False
            Frame20Item.Visible = False
            
            cmdayudaInst.Visible = False
            adoPrimaryRS.Close
            adoPrimaryRS.Open "select * FROM INSTITUTOS", Cnn, adOpenStatic, adLockBatchOptimistic, adCmdText
            adoManCiuIns.Open "Select CODIGO,DESCRIPCION FROM PARAMETROS", Cnn
            'Dim oText As TextBox
            'Enlaza los cuadros de texto con el proveedor de datos
            For Each oText In Me.TXTFIELDS
                If oText.Index > 81 And oText.Index < 90 Then
                    Set oText.DataSource = adoPrimaryRS
                End If
            Next
            
            adoManCiuIns.Find "CODIGO=" & TXTFIELDS(85).Text
            LBLManCiuIns(5).Caption = adoManCiuIns.Fields("Descripcion").Value
            'Set LBLManCiuIns(5).DataSource = adoManCiuIns

Case 10

            On Error Resume Next
            adoPrimaryRS.CancelUpdate
            
            Frame12Procedimientos.Visible = True
            Frame1PERSONA.Visible = False
            Frame2Empresas.Visible = False
            Frame3PROVEEDOR.Visible = False
            Frame4LINEA.Visible = False
            frame5INSUMO.Visible = False
            Frame6DOCTOR.Visible = False
            Frame7AREA.Visible = False
            Frame8ESPECIALIDAD.Visible = False
            Frame9PARAMETROS.Visible = False
            Frame10VALORNORMAL.Visible = False
            Frame11INSTITUTO.Visible = False
            Frame13USUARIO.Visible = False
            Frame14TABLAVALOR.Visible = False
            Frame15Habitaciones.Visible = False
            Frame16Servicios.Visible = False
            Frame17TipoServicio.Visible = False
            Frame18Aranceles.Visible = False
            Frame19Medicamentos.Visible = False
            Frame20Item.Visible = False
            
            cmdayudaInst.Visible = False
            adoPrimaryRS.Close
            adoPrimaryRS.Open "select * FROM PROCEDIMIENTOS", Cnn, adOpenStatic, adLockBatchOptimistic, adCmdText

            'Dim oText As TextBox
            'Enlaza los cuadros de texto con el proveedor de datos
            For Each oText In Me.TXTFIELDS
                If oText.Index > 164 And oText.Index < 170 Then
                    Set oText.DataSource = adoPrimaryRS
                End If
            Next

Case 11

            On Error Resume Next
            adoPrimaryRS.CancelUpdate
            
            Frame13USUARIO.Visible = CtrVisualizar.Visualizar(11, Cnn, IdUsuario)
            Frame1PERSONA.Visible = False
            Frame2Empresas.Visible = False
            Frame3PROVEEDOR.Visible = False
            Frame4LINEA.Visible = False
            frame5INSUMO.Visible = False
            Frame6DOCTOR.Visible = False
            Frame7AREA.Visible = False
            Frame8ESPECIALIDAD.Visible = False
            Frame9PARAMETROS.Visible = False
            Frame10VALORNORMAL.Visible = False
            Frame11INSTITUTO.Visible = False
            Frame12Procedimientos.Visible = False
            Frame14TABLAVALOR.Visible = False
            Frame15Habitaciones.Visible = False
            Frame16Servicios.Visible = False
            Frame17TipoServicio.Visible = False
            Frame18Aranceles.Visible = False
            Frame19Medicamentos.Visible = False
            
            Frame20Item.Visible = False
            adoPrimaryRS.Close
            adoPrimaryRS.Open "select * FROM USUARIOS", Cnn, adOpenStatic, adLockBatchOptimistic, adCmdText
            adoManNomUsu.Open "SELECT TIPDOC,NRODOC,NOMBRES,APELLIDOS FROM PERSONAS WHERE TIPOPERSONA='001'", Cnn
            'Dim oText As TextBox
            'Enlaza los cuadros de texto con el proveedor de datos
            For Each oText In Me.TXTFIELDS
                If oText.Index > 59 And oText.Index < 64 Then
                    Set oText.DataSource = adoPrimaryRS
                End If
            Next
            Set Check1.DataSource = adoPrimaryRS
            
Case 12
            
            On Error Resume Next
            adoPrimaryRS.CancelUpdate
            
            Frame3PROVEEDOR.Visible = CtrVisualizar.Visualizar(12, Cnn, IdUsuario)
            Frame1PERSONA.Visible = False
            Frame2Empresas.Visible = False
            Frame4LINEA.Visible = False
            frame5INSUMO.Visible = False
            Frame6DOCTOR.Visible = False
            Frame7AREA.Visible = False
            Frame8ESPECIALIDAD.Visible = False
            Frame9PARAMETROS.Visible = False
            Frame10VALORNORMAL.Visible = False
            Frame11INSTITUTO.Visible = False
            Frame12Procedimientos.Visible = False
            Frame13USUARIO.Visible = False
            Frame14TABLAVALOR.Visible = False
            Frame15Habitaciones.Visible = False
            Frame16Servicios.Visible = False
            Frame17TipoServicio.Visible = False
            Frame18Aranceles.Visible = False
            Frame19Medicamentos.Visible = False
            Frame20Item.Visible = False
            
            cmdAyudaProvee.Visible = False
            adoPrimaryRS.Close
            adoPrimaryRS.Open "select * FROM PROVEEDORES", Cnn, adOpenStatic, adLockBatchOptimistic, adCmdText
            adoManCiuProv.Open "Select CODIGO,DESCRIPCION FROM PARAMETROS", Cnn
            'Dim oText As TextBox
            'Enlaza los cuadros de texto con el proveedor de datos
            For Each oText In Me.TXTFIELDS
                If oText.Index > 63 And oText.Index < 71 Then
                    Set oText.DataSource = adoPrimaryRS
                End If
            Next
Case 13

            On Error Resume Next
            adoPrimaryRS.CancelUpdate
            
            Frame9PARAMETROS.Visible = CtrVisualizar.Visualizar(13, Cnn, IdUsuario)
            Frame1PERSONA.Visible = False
            Frame2Empresas.Visible = False
            Frame3PROVEEDOR.Visible = False
            Frame4LINEA.Visible = False
            frame5INSUMO.Visible = False
            Frame6DOCTOR.Visible = False
            Frame7AREA.Visible = False
            Frame8ESPECIALIDAD.Visible = False
            Frame10VALORNORMAL.Visible = False
            Frame11INSTITUTO.Visible = False
            Frame12Procedimientos.Visible = False
            Frame13USUARIO.Visible = False
            Frame14TABLAVALOR.Visible = False
            Frame15Habitaciones.Visible = False
            Frame16Servicios.Visible = False
            Frame17TipoServicio.Visible = False
            Frame18Aranceles.Visible = False
            Frame19Medicamentos.Visible = False
            Frame20Item.Visible = False
                        
            adoPrimaryRS.Close
            adoPrimaryRS.Open "select * FROM PARAMETROS", Cnn, adOpenStatic, adLockBatchOptimistic, adCmdText
            'Dim oText As TextBox
            'Enlaza los cuadros de texto con el proveedor de datos
            For Each oText In Me.TXTFIELDS
                If oText.Index > 29 And oText.Index < 33 Then
                    Set oText.DataSource = adoPrimaryRS
                End If
            Next
            
Case 14

            On Error Resume Next
            adoPrimaryRS.CancelUpdate

            Frame14TABLAVALOR.Visible = CtrVisualizar.Visualizar(14, Cnn, IdUsuario)
            Frame1PERSONA.Visible = False
            Frame2Empresas.Visible = False
            Frame3PROVEEDOR.Visible = False
            Frame4LINEA.Visible = False
            frame5INSUMO.Visible = False
            Frame6DOCTOR.Visible = False
            Frame7AREA.Visible = False
            Frame8ESPECIALIDAD.Visible = False
            Frame9PARAMETROS.Visible = False
            Frame10VALORNORMAL.Visible = False
            Frame11INSTITUTO.Visible = False
            Frame12Procedimientos.Visible = False
            Frame13USUARIO.Visible = False
            Frame15Habitaciones.Visible = False
            Frame16Servicios.Visible = False
            Frame17TipoServicio.Visible = False
            Frame18Aranceles.Visible = False
            Frame19Medicamentos.Visible = False
            Frame20Item.Visible = False
                        
            adoPrimaryRS.Close
            adoPrimaryRS.Open "select * FROM tabval", Cnn, adOpenStatic, adLockBatchOptimistic, adCmdText
            'Dim oText As TextBox
            'Enlaza los cuadros de texto con el proveedor de datos
            For Each oText In Me.TXTFIELDS
                If oText.Index > 78 And oText.Index < 82 Then
                    Set oText.DataSource = adoPrimaryRS
                End If
            Next
            
            
 Case 15
            
            
            On Error Resume Next
            adoPrimaryRS.CancelUpdate
            
            Frame15Habitaciones.Visible = CtrVisualizar.Visualizar(15, Cnn, IdUsuario)
            Frame1PERSONA.Visible = False
            Frame2Empresas.Visible = False
            Frame3PROVEEDOR.Visible = False
            Frame4LINEA.Visible = False
            frame5INSUMO.Visible = False
            Frame6DOCTOR.Visible = False
            Frame7AREA.Visible = False
            Frame8ESPECIALIDAD.Visible = False
            Frame9PARAMETROS.Visible = False
            Frame10VALORNORMAL.Visible = False
            Frame11INSTITUTO.Visible = False
            Frame12Procedimientos.Visible = False
            Frame13USUARIO.Visible = False
            Frame14TABLAVALOR.Visible = False
            Frame16Servicios.Visible = False
            Frame17TipoServicio.Visible = False
            Frame18Aranceles.Visible = False
            Frame19Medicamentos.Visible = False
            Frame20Item.Visible = False
                        
            adoPrimaryRS.Close
            adoPrimaryRS.Open "select * FROM Habitaciones", Cnn, adOpenStatic, adLockBatchOptimistic, adCmdText
            adoManTipHab.Open "Select CODIGO,DESCRIPCION FROM PARAMETROS", Cnn
            'Dim oText As TextBox
            'Enlaza los cuadros de texto con el proveedor de datos
            For Each oText In Me.TXTFIELDS
                If oText.Index > 89 And oText.Index < 95 Then
                    Set oText.DataSource = adoPrimaryRS
                End If
            Next
            
            
  Case 16
            
            
            On Error Resume Next
            adoPrimaryRS.CancelUpdate
            
            Frame16Servicios.Visible = CtrVisualizar.Visualizar(16, Cnn, IdUsuario)
            Frame1PERSONA.Visible = False
            Frame2Empresas.Visible = False
            Frame3PROVEEDOR.Visible = False
            Frame4LINEA.Visible = False
            frame5INSUMO.Visible = False
            Frame6DOCTOR.Visible = False
            Frame7AREA.Visible = False
            Frame8ESPECIALIDAD.Visible = False
            Frame9PARAMETROS.Visible = False
            Frame10VALORNORMAL.Visible = False
            Frame11INSTITUTO.Visible = False
            Frame12Procedimientos.Visible = False
            Frame13USUARIO.Visible = False
            Frame14TABLAVALOR.Visible = False
            Frame15Habitaciones.Visible = False
            Frame17TipoServicio.Visible = False
            Frame18Aranceles.Visible = False
            Frame19Medicamentos.Visible = False
            Frame20Item.Visible = False
                       
            adoPrimaryRS.Close
            adoPrimaryRS.Open "select * FROM Servicio", Cnn, adOpenStatic, adLockBatchOptimistic, adCmdText
            adoManTipoSer.Open "Select * From TIPOSERVICIO", Cnn
            adoServiProce.Open "Select * from PROCEDIMIENTOS", Cnn
            adoServiArea.Open "Select * from AREA", Cnn
            
            'Dim oText As TextBox
            'Enlaza los cuadros de texto con el proveedor de datos
            For Each oText In Me.TXTFIELDS
                If oText.Index > 100 And oText.Index < 117 Then
                    Set oText.DataSource = adoPrimaryRS
                End If
            Next
            
            
 Case 17
            
            On Error Resume Next
            adoPrimaryRS.CancelUpdate
            
            Frame17TipoServicio.Visible = CtrVisualizar.Visualizar(17, Cnn, IdUsuario)
            Frame1PERSONA.Visible = False
            Frame2Empresas.Visible = False
            Frame3PROVEEDOR.Visible = False
            Frame4LINEA.Visible = False
            frame5INSUMO.Visible = False
            Frame6DOCTOR.Visible = False
            Frame7AREA.Visible = False
            Frame8ESPECIALIDAD.Visible = False
            Frame9PARAMETROS.Visible = False
            Frame10VALORNORMAL.Visible = False
            Frame11INSTITUTO.Visible = False
            Frame12Procedimientos.Visible = False
            Frame13USUARIO.Visible = False
            Frame14TABLAVALOR.Visible = False
            Frame15Habitaciones.Visible = False
            Frame16Servicios.Visible = False
            Frame18Aranceles.Visible = False
            Frame19Medicamentos.Visible = False
            Frame20Item.Visible = False
                        
            adoPrimaryRS.Close
            adoPrimaryRS.Open "select * FROM TIPOSERVICIO", Cnn, adOpenStatic, adLockBatchOptimistic, adCmdText
           
            'Dim oText As TextBox
            'Enlaza los cuadros de texto con el proveedor de datos
            For Each oText In Me.TXTFIELDS
                If oText.Index > 74 And oText.Index < 77 Then
                    Set oText.DataSource = adoPrimaryRS
                End If
            Next
            
            
Case 18
            
            On Error Resume Next
            adoPrimaryRS.CancelUpdate
                      
           '  Frame18Aranceles.Visible = CtrVisualizar.Visualizar(18, Cnn, IdUsuario)
            Frame18Aranceles.Visible = True
            Frame1PERSONA.Visible = False
            Frame2Empresas.Visible = False
            Frame3PROVEEDOR.Visible = False
            Frame4LINEA.Visible = False
            frame5INSUMO.Visible = False
            Frame6DOCTOR.Visible = False
            Frame7AREA.Visible = False
            Frame8ESPECIALIDAD.Visible = False
            Frame9PARAMETROS.Visible = False
            Frame10VALORNORMAL.Visible = False
            Frame11INSTITUTO.Visible = False
            Frame12Procedimientos.Visible = False
            Frame13USUARIO.Visible = False
            Frame14TABLAVALOR.Visible = False
            Frame15Habitaciones.Visible = False
            Frame16Servicios.Visible = False
            Frame17TipoServicio.Visible = False
            Frame19Medicamentos.Visible = False
            Frame20Item.Visible = False
                        
            adoPrimaryRS.Close
            adoPrimaryRS.Open "select * FROM ARANCELES", Cnn, adOpenStatic, adLockBatchOptimistic, adCmdText
            adoSerAran.Open "select * FROM Servicio", Cnn
            adoAranDoc.Open "SELECT DOCTOR.CODDR,PERSONAS.APELLIDOS,PERSONAS.NOMBRES, PERSONAS.TIPDOC,PERSONAS.NRODOC,DOCTOR.REGPRF FROM DOCTOR  LEFT OUTER JOIN PERSONAS ON DOCTOR.TIPDOC = PERSONAS.TIPDOC AND DOCTOR.NRODOC = PERSONAS.NRODOC", Cnn
            adoAranIns.Open "SELECT CODINS,INSTITUTO FROM institutos", Cnn
            adoAranTiser.Open "select * FROM TIPOSERVICIO", Cnn
            
            'Dim oText As TextBox
            'Enlaza los cuadros de texto con el proveedor de datos
            For Each oText In Me.TXTFIELDS
                If oText.Index > 129 And oText.Index < 136 Then
                    Set oText.DataSource = adoPrimaryRS
                End If
            Next
            
   Case 19
            
            On Error Resume Next
            adoPrimaryRS.CancelUpdate
                      
            'Frame18Aranceles.Visible = CtrVisualizar.Visualizar(18, Cnn, IdUsuario)
            Frame19Medicamentos.Visible = True
            Frame1PERSONA.Visible = False
            Frame2Empresas.Visible = False
            Frame3PROVEEDOR.Visible = False
            Frame4LINEA.Visible = False
            frame5INSUMO.Visible = False
            Frame6DOCTOR.Visible = False
            Frame7AREA.Visible = False
            Frame8ESPECIALIDAD.Visible = False
            Frame9PARAMETROS.Visible = False
            Frame10VALORNORMAL.Visible = False
            Frame11INSTITUTO.Visible = False
            Frame12Procedimientos.Visible = False
            Frame13USUARIO.Visible = False
            Frame14TABLAVALOR.Visible = False
            Frame15Habitaciones.Visible = False
            Frame16Servicios.Visible = False
            Frame17TipoServicio.Visible = False
            Frame18Aranceles.Visible = False
            Frame20Item.Visible = False
            
                
            adoPrimaryRS.Close
            adoPrimaryRS.Open "select * FROM MEDICAMENTOS", Cnn, adOpenStatic, adLockBatchOptimistic, adCmdText
           
            'Dim oText As TextBox
            'Enlaza los cuadros de texto con el proveedor de datos
            For Each oText In Me.TXTFIELDS
                If oText.Index > 149 And oText.Index < 161 Then
                    Set oText.DataSource = adoPrimaryRS
                End If
            Next
                      
   Case 20
                      
            On Error Resume Next
            adoPrimaryRS.CancelUpdate
                      
            Frame20Item.Visible = CtrVisualizar.Visualizar(20, Cnn, IdUsuario)
            Frame1PERSONA.Visible = False
            Frame2Empresas.Visible = False
            Frame3PROVEEDOR.Visible = False
            Frame4LINEA.Visible = False
            frame5INSUMO.Visible = False
            Frame6DOCTOR.Visible = False
            Frame7AREA.Visible = False
            Frame8ESPECIALIDAD.Visible = False
            Frame9PARAMETROS.Visible = False
            Frame10VALORNORMAL.Visible = False
            Frame11INSTITUTO.Visible = False
            Frame12Procedimientos.Visible = False
            Frame13USUARIO.Visible = False
            Frame14TABLAVALOR.Visible = False
            Frame15Habitaciones.Visible = False
            Frame16Servicios.Visible = False
            Frame17TipoServicio.Visible = False
            Frame18Aranceles.Visible = False
            Frame19Medicamentos.Visible = False
                
            adoPrimaryRS.Close
            adoPrimaryRS.Open "select * FROM ITEM_HISTORIACLINICA_ESPECIALI", Cnn, adOpenStatic, adLockBatchOptimistic, adCmdText
            adoManEspeItem.Open "Select * from ESPECIALIDAD", Cnn
            adoManTipAnam.Open "Select CODIGO,DESCRIPCION FROM PARAMETROS WHERE TIPO='ANAMNESIS'", Cnn
            'Dim oText As TextBox
            'Enlaza los cuadros de texto con el proveedor de datos
    
                                                   
            
            For Each oText In Me.TXTFIELDS
                If oText.Index > 169 And oText.Index < 174 Then
                    Set oText.DataSource = adoPrimaryRS
                End If
            Next
        
End Select

End Sub
Private Sub Form_Resize()
  On Error Resume Next
  TabStrip1.Width = Me.ScaleWidth - TabStrip1.Left
  TabStrip1.Height = Me.ScaleHeight - 100
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
Private Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  Dim bCancel As Boolean
  Select Case adReason
  Case adRsnAddNew
        bActualizar = True
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
            bActualizar = False
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
        bActualizar = False
  Case adRsnUpdate
        bActualizar = False
  End Select
  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub txtFields_Change(Index As Integer)
On Error Resume Next
Select Case Me.TabStrip1.SelectedItem.Index
           
      
    Case 2
                   
          If Not adoManCiuEmp.State = 0 Then
             adoManCiuEmp.MoveFirst
             If Not TXTFIELDS(123).Text = "" Then
                 adoManCiuEmp.Filter = "CODIGO=" & Trim(TXTFIELDS(123).Text)
                 LBLManCiuEmp(1).Caption = adoManCiuEmp.Fields("Descripcion").Value
             Else
                 LBLManCiuEmp(1).Caption = ""
                 
             End If
          
         End If
        

    Case 4
    
         If Not adoManProIns.State = 0 Then
             adoManProIns.MoveFirst
            If Not TXTFIELDS(56).Text = "" Then
                adoManProIns.Filter = "COD_PROVEEDOR=" & TXTFIELDS(56).Text
                LBLManProIns(1).Caption = adoManProIns.Fields("NOM_PROVEEDOR").Value
            Else
                LBLManProIns(1).Caption = ""
                
            End If
         
         End If
         
                 
          If Not adoManLinIns.State = 0 Then
             adoManLinIns.MoveFirst
            If Not TXTFIELDS(57).Text = "" Then
                adoManLinIns.Filter = "COD_LINEA=" & TXTFIELDS(57).Text
                LBLManLiIns(1).Caption = adoManLinIns.Fields("LIN_NOMBRE").Value
            Else
                LBLManLiIns(1).Caption = ""
                
            End If
         End If
         
    Case 5
       If Not TXTFIELDS(22).Text = "" And Index = 22 Then
          adoManEspecialidad.Filter = "CODDR=" & TXTFIELDS(22).Text

           If Not adoManDoctor.BOF And Not MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
               adoManDoctor.MoveFirst
               adoManDoctor.Find "CODDR=" & TXTFIELDS(22).Text
               LBLManDoc(0).Caption = adoManDoctor.Fields("APELLIDOS").Value + "," + adoManDoctor.Fields("nombres").Value
               adoManEspecialidad.Filter = Nothing
               If (IsNull(adoManDoctor.Fields("RUC").Value)) Then
                   LBLManDoc(4).Caption = ""
               Else
                   LBLManDoc(4).Caption = adoManDoctor.Fields("RUC").Value
               End If
               
               If (IsNull(adoManDoctor.Fields("E_MAIL").Value)) Then
                   LBLManDoc(5).Caption = ""
               Else
                   LBLManDoc(5).Caption = adoManDoctor.Fields("E_MAIL").Value
               End If
               
               If (IsNull(adoManDoctor.Fields("TELEFONOS").Value)) Then
                   LBLManDoc(6).Caption = ""
               Else
                   LBLManDoc(6).Caption = adoManDoctor.Fields("TELEFONOS").Value
               End If
               
               If (IsNull(adoManDoctor.Fields("MOVIL").Value)) Then
                   LBLManDoc(7).Caption = ""
               Else
                   LBLManDoc(7).Caption = adoManDoctor.Fields("MOVIL").Value
               End If
                         
                                      
               If (IsNull(adoManDoctor.Fields("DOM_PARTICULAR").Value)) Then
                   LBLManDoc(8).Caption = ""
               Else
                   LBLManDoc(8).Caption = adoManDoctor.Fields("DOM_PARTICULAR").Value
               End If
                                      
               If (IsNull(adoManDoctor.Fields("COD_CIUDADP").Value)) Then
                  LBLManDoc(9).Caption = ""
                  LBLManDoc(1).Caption = ""
               Else
                   LBLManDoc(9).Caption = adoManDoctor.Fields("COD_CIUDADP").Value
                   adoManCiuDoc.MoveFirst
                   adoManCiuDoc.Find "CODIGO='" & Trim$(LBLManDoc(9).Caption) & "'"
                   LBLManDoc(1).Caption = adoManCiuDoc.Fields("Descripcion").Value
               End If
                                   
               If (IsNull(adoManDoctor.Fields("DOM_LABORAL").Value)) Then
                   LBLManDoc(10).Caption = ""
               Else
                   LBLManDoc(10).Caption = adoManDoctor.Fields("DOM_LABORAL").Value
               End If
                                  
               If (IsNull(adoManDoctor.Fields("COD_CIUDADL").Value)) Then
                   LBLManDoc(11).Caption = ""
               Else
                   LBLManDoc(11).Caption = adoManDoctor.Fields("COD_CIUDADL").Value
                   adoManCiuDoc.MoveFirst
                   adoManCiuDoc.Find "CODIGO=" & Trim$(LBLManDoc(11).Caption)
                   LBLManDoc(2).Caption = adoManCiuDoc.Fields("Descripcion").Value
                   
               End If
               
               If (IsNull(adoManDoctor.Fields("RADIO").Value)) Then
                   LBLManDoc(12).Caption = ""
               Else
                   LBLManDoc(12).Caption = adoManDoctor.Fields("RADIO").Value
                
               End If
               
           End If
         
         End If
  
    Case 8
        If Not adoValNorGen.State = 0 Then
           adoValNorGen.MoveFirst
           If Not TXTFIELDS(41).Text = "" Then
              adoValNorGen.Filter = "CODIGO='" & TXTFIELDS(41).Text & "'"
       
              LBLManGENE(1).Caption = adoValNorGen.Fields("Descripcion").Value
           Else
              LBLManGENE(1).Caption = ""
                
           End If
         
        End If
         
    Case 9
    
     
        If Not adoManCiuIns.State = 0 Then
           adoManCiuIns.MoveFirst
           If Not TXTFIELDS(85).Text = "" Then
               adoManCiuIns.Filter = "CODIGO=" & TXTFIELDS(85).Text
       
               LBLManCiuIns(5).Caption = adoManCiuIns.Fields("Descripcion").Value
           Else
                LBLManCiuIns(5).Caption = ""
                
           End If
          
        End If
      
    
   Case 11
              
     If Not adoManNomUsu.State = 0 Then
            adoManNomUsu.MoveFirst
            If Not TXTFIELDS(61).Text = "" Then
                adoManNomUsu.Find "TIPDOC='" & Trim$(TXTFIELDS(61).Text) & "'"
            End If
     End If
        If Not adoManNomUsu.State = 0 Then
            adoManNomUsu.MoveFirst
            If Not TXTFIELDS(62).Text = "" Then
                adoManNomUsu.Find "NRODOC=" & Trim$(TXTFIELDS(62).Text)
                LBLManUsu(1).Caption = adoManNomUsu.Fields("APELLIDOS").Value + ", " + adoManNomUsu.Fields("NOMBRES").Value
            Else
                LBLManUsu(1).Caption = ""
            End If
        End If
       
                       
   Case 12
  
       If Not adoManCiuProv.State = 0 Then
         adoManCiuProv.MoveFirst
         If Not TXTFIELDS(68).Text = "" Then
             adoManCiuProv.Filter = "CODIGO=" & TXTFIELDS(68).Text
       
             LBLManCiuProv(0).Caption = adoManCiuProv.Fields("Descripcion").Value
         Else
             LBLManCiuProv(0).Caption = ""
                
         End If
         
      End If
      
   Case 16
  
       If Not adoManTipoSer.State = 0 Then
         adoManTipoSer.MoveFirst
         If Not TXTFIELDS(102).Text = "" Then
             adoManTipoSer.Filter = "TIPOSER='" & Trim$(TXTFIELDS(102).Text) & "'"
             LBLServiTiSer(0).Caption = adoManTipoSer.Fields("NOMTIPOSER").Value
         Else
             LBLServiTiSer(0).Caption = ""
                
         End If
         
       End If
      
      
       If Not adoServiProce.State = 0 Then
         adoServiProce.MoveFirst
         If Not TXTFIELDS(105).Text = "" Then
             adoServiProce.Filter = "COD_PROCEDIMIENTO=" & Trim$(TXTFIELDS(105).Text)
             LBLServiProce(1).Caption = adoServiProce.Fields("NOM_PROCEDIMIENTO").Value
         Else
             LBLServiProce(1).Caption = ""
                
         End If
         
       End If
        
       If Not adoServiArea.State = 0 Then
         adoServiArea.MoveFirst
         If Not TXTFIELDS(113).Text = "" Then
             adoServiArea.Filter = "CODAREA='" & Trim$(TXTFIELDS(113).Text) & "'"
             LBLServiArea(0).Caption = adoServiArea.Fields("DESCRIPCION").Value
         Else
             LBLServiArea(0).Caption = ""
                
         End If
         
       End If
       
       
   Case 18
    
       If Not adoSerAran.State = 0 Then
         adoSerAran.MoveFirst
         If Not TXTFIELDS(130).Text = "" Then
             adoSerAran.Filter = "CODSERVICIO=" & Trim$(TXTFIELDS(130).Text)
             LBLManAraSer(1).Caption = adoSerAran.Fields("NOMSERVICIO").Value
         Else
             LBLManAraSer(1).Caption = ""
                
         End If
         
       End If
    
        
        If Not adoAranTiser.State = 0 Then
         adoAranTiser.MoveFirst
         If Not TXTFIELDS(131).Text = "" Then
             adoAranTiser.Filter = "TIPOSER='" & Trim(TXTFIELDS(131).Text) & "'"
             LBLManAraTiSer(0).Caption = adoAranTiser.Fields("NOMTIPOSER").Value
         Else
             LBLManAraTiSer(0).Caption = ""
                
         End If
         
       End If
        
            
       If Not adoAranDoc.State = 0 Then
         adoAranDoc.MoveFirst
         If Not TXTFIELDS(132).Text = "" Then
             adoAranDoc.Filter = "CODDR=" & Trim$(TXTFIELDS(132).Text)
             LBLAraDoc(0).Caption = adoAranDoc.Fields("APELLIDOS").Value + "," + adoAranDoc.Fields("nombres").Value
         Else
             LBLAraDoc(0).Caption = ""
                
         End If
         
       End If
               
                      
       If Not adoAranIns.State = 0 Then
         adoAranIns.MoveFirst
         If Not TXTFIELDS(133).Text = "" Then
             adoAranIns.Filter = "CODINS=" & Trim$(TXTFIELDS(133).Text)
             LBLInsti(1).Caption = adoAranIns.Fields("INSTITUTO").Value
         Else
             LBLInsti(1).Caption = ""
                
         End If
         
       End If
               
    Case 20
    
        If Not adoManEspeItem.State = 0 Then
           adoManEspeItem.MoveFirst
          If Not TXTFIELDS(170).Text = "" Then
               adoManEspeItem.Filter = "COD_ESPECIALIDAD=" & Trim$(TXTFIELDS(170).Text)
               LBLManEspIte(1).Caption = adoManEspeItem.Fields("NOMBRE").Value
          Else
               LBLManEspIte(1).Caption = ""
                
          End If
         
        End If
        
     
        If Not adoManTipAnam.State = 0 Then
            adoManTipAnam.MoveFirst
          If Not TXTFIELDS(171).Text = "" Then
              adoManTipAnam.Filter = "CODIGO='" & (TXTFIELDS(171).Text) & "'"
              LblManIteTipo(0).Caption = adoManTipAnam.Fields("DESCRIPCION").Value
          Else
              LblManIteTipo(0).Caption = ""
                
          End If
         
        End If
        

        
    End Select
             
End Sub


Private Sub txtfields_GotFocus(Index As Integer)
        If Index = 1 Then
            bPrifoco = True
        Else
            bPrifoco = False
        End If

        'If Index = 91 And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
         '      cmdAyudaNRODOC.Visible = True
       ' Else
        '       cmdAyudaNRODOC.Visible = False
        'End If

        
        If Index = 33 And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
               cmdAyudaDocCiu1.Visible = True
        Else
               cmdAyudaDocCiu1.Visible = False
        End If

        If Index = 35 And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
                cmdAyudaDocCiu2.Visible = True
        Else
                cmdAyudaDocCiu2.Visible = False
        End If
        
        
        If Index = 41 And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
                CmdAyuGene.Visible = True
        Else
                CmdAyuGene.Visible = False
        End If
        
        
        If Index = 56 And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
                cmdAyudaProv.Visible = True
        Else
              cmdAyudaProv.Visible = False
        End If
        

        If Index = 57 And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
              cmdAyudaLinea.Visible = True
        Else
              cmdAyudaLinea.Visible = False
        End If
        
        
        
        
        If Index = 68 And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
                cmdAyudaProvee.Visible = True
        Else
                cmdAyudaProvee.Visible = False
        End If


        If Index = 85 And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
                cmdayudaInst.Visible = True
        Else
                cmdayudaInst.Visible = False
        End If
        
        
        
        If Index = 90 And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
                CmdAyudaTipoHab.Visible = True
        Else
                CmdAyudaTipoHab.Visible = False
        End If
        
        If Index = 91 And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
                CmdAyudaHabitacion.Visible = True
        Else
                CmdAyudaHabitacion.Visible = False
        End If
        
        If Index = 93 And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
                CmdEstadoCama.Visible = True
        Else
                CmdEstadoCama.Visible = False
        End If
             
        If Index = 102 And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
                CmdAyudaTipoSer.Visible = True
        Else
                CmdAyudaTipoSer.Visible = False
        End If
             
        If Index = 105 And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
                CmdAyudaProcedimiento.Visible = True
        Else
                CmdAyudaProcedimiento.Visible = False
        End If
        
        If Index = 113 And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
                CmdAyudaArea.Visible = True
        Else
                CmdAyudaArea.Visible = False
        End If
        
        If Index = 123 And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
                CmdAyudaCiud.Visible = True
        Else
                CmdAyudaCiud.Visible = False
        End If
        
        If Index = 130 And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
                CmdAyuSer.Visible = True
        Else
                CmdAyuSer.Visible = False
        End If
       
        If Index = 132 And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
                cmdbusdoc.Visible = True
        Else
                cmdbusdoc.Visible = False
        End If
        
        
        
        If Index = 133 And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
                CmdAranIns.Visible = True
        Else
                CmdAranIns.Visible = False
        End If
        
        
        If Index = 133 And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
                CmdAranIns.Visible = True
        Else
                CmdAranIns.Visible = False
        End If
        
        If Index = 170 And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
                CmdAyudaEspe.Visible = True
        Else
                CmdAyudaEspe.Visible = False
        End If
        
        
        If Index = 171 And MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
                CmdTipoAnamn.Visible = True
        Else
                CmdTipoAnamn.Visible = False
        End If
        
       
        
End Sub
Private Sub TXTFIELDS_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
Dim X As Boolean
Select Case Index
      
        Case 0
        Case 5
        Case 6
        Case 7
        'Case 10
        'Case 14
        
    '     Case 15
     '        If MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
      '             Dim ciudadp As Integer
       '            ciudadp = OPTENER_CODIGO("SELECT COUNT(COD_CIUDADP) FROM PARAMETROS WHERE TIPO=" & txtfields(15).Text & " ")
        '                                                               "select CODIGO,DESCRIPCION from PARAMETROS WHERE TIPO='CIUDAD'"
         '                   If txtfields(15).Text = "" Then
          '                      MsgBox "DISCULPE EL CODIGO ESTA VACIO, POR FAVOR INGRESE NUEVAMENTE", vbExclamation
           '                   blCancel = True
            '                Else
             '                     If ciudadp < 2 Then
              '                          MsgBox "DISCULPE EL CODIGO NO EXISTE, POR FAVOR INGRESE NUEVAMENTE", vbExclamation
               '                         blCancel = True
                '                        txtfields(15).SelStart = 0
                 ''                       txtfields(15).SelLength = Len(txtfields(15).Text)
                   '                 End If
        '
                       '     End If
         '    End If
           
          ' Case 17
           '  If MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
            '       Dim ciudadl As Integer
             '      ciudadl = OPTENER_CODIGO("SELECT COUNT(COD_CIUDADL) FROM CIUDADES WHERE COD_CIUDADP=" & txtfields(17).Text & " ")
              '              If txtfields(15).Text = "" Then
               '                 MsgBox "DISCULPE EL CODIGO ESTA VACIO, POR FAVOR INGRESE NUEVAMENTE", vbExclamation
                '              blCancel = True
                 '           Else
                  '                If ciudadp < 2 Then
                   '                     MsgBox "DISCULPE EL CODIGO NO EXISTE, POR FAVOR INGRESE NUEVAMENTE", vbExclamation
                    '                    blCancel = True
                     '                   txtfields(17).SelStart = 0
                      '                  txtfields(17).SelLength = Len(txtfields(17).Text)
                       '             End If
        '
         '                   End If
          '   End If
            
            Case 24
           
            Case 25
                                       
            Case 26
            If MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
                    Dim espe As Integer
                    espe = OPTENER_CODIGO("SELECT COUNT(CODESP) FROM ESPECIALIDAD WHERE CODESP=  " & TXTFIELDS(26).Text & "")
                            If TXTFIELDS(26).Text = "" Then
                                 MsgBox "DISCULPE EL CODIGO ESTA VACIO, POR FAVOR INGRESE NUEVAMENTE", vbExclamation
                                blCancel = True
                            Else
                                   If espe < 2 Then
                                        MsgBox "DISCULPE EL CODIGO NO EXISTE, POR FAVOR INGRESE NUEVAMENTE", vbExclamation
                                        blCancel = True
                                        TXTFIELDS(26).SelStart = 0
                                        TXTFIELDS(26).SelLength = Len(TXTFIELDS(26).Text)
                                    End If
                                    
                         End If
                End If
        Case 33
                 If MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
                    Dim ciudaddr As Integer
                    ciudaddr = OPTENER_CODIGO("SELECT COUNT(CODCIUDAD) FROM CIUDADES WHERE CODCIUDAD=" & TXTFIELDS(33).Text & "")
                    If TXTFIELDS(33).Text = "" Then
                                 MsgBox "DISCULPE EL CODIGO ESTA VACIO, POR FAVOR INGRESE NUEVAMENTE", vbExclamation
                                blCancel = True
                            Else
                                   If ciudaddr < 2 Then
                                        MsgBox "DISCULPE EL CODIGO NO EXISTE, POR FAVOR INGRESE NUEVAMENTE", vbExclamation
                                        blCancel = True
                                        TXTFIELDS(33).SelStart = 0
                                        TXTFIELDS(33).SelLength = Len(TXTFIELDS(33).Text)
                                    End If
                                    
                         End If
           End If
        Case 35
                 If MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
                    Dim ciudaddra As Integer
                    ciudaddra = OPTENER_CODIGO("SELECT COUNT(CODCIUDAD) FROM CIUDADES WHERE CODCIUDAD=" & TXTFIELDS(35).Text & "")
                    If TXTFIELDS(35).Text = "" Then
                                 MsgBox "DISCULPE EL CODIGO ESTA VACIO, POR FAVOR INGRESE NUEVAMENTE", vbExclamation
                                blCancel = True
                            Else
                                   If ciudaddra < 2 Then
                                        MsgBox "DISCULPE EL CODIGO NO EXISTE, POR FAVOR INGRESE NUEVAMENTE", vbExclamation
                                        blCancel = True
                                        TXTFIELDS(35).SelStart = 0
                                        TXTFIELDS(35).SelLength = Len(TXTFIELDS(35).Text)
                                    End If
                                    
                         End If
                 End If
           
            Case 36
            If MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
             TXTFIELDS(36).Text = UCase(TXTFIELDS(36).Text)
                    
                    Dim i As Integer
                    i = OPTENER_CODIGO("SELECT COUNT(CODAREA) FROM AREA WHERE CODAREA= '" & TXTFIELDS(36).Text & "'")
                            If TXTFIELDS(36) = "" Then
                                    MsgBox "DISCULPE EL CODIGO ESTA VACIO, POR FAVOR INGRESE NUEVAMENTE", vbExclamation
                                    blCancel = True
                           End If
                            If i > 1 Then
                                         MsgBox "DISCULPE EL CODIGO YA EXISTE, POR FAVOR INGRESE NUEVAMENTE", vbExclamation
                                        blCancel = True
                                        TXTFIELDS(36).SelStart = 0
                                        TXTFIELDS(36).SelLength = Len(TXTFIELDS(36))
                            End If
           End If
            Case 40
            If MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
                    Dim anali As Integer
                    anali = OPTENER_CODIGO("SELECT COUNT(CODIGO) FROM ANALISIS WHERE CODIGO=  " & TXTFIELDS(40).Text & "")
                            If TXTFIELDS(40).Text = "" Then
                                 MsgBox "DISCULPE EL CODIGO ESTA VACIO, POR FAVOR INGRESE NUEVAMENTE", vbExclamation
                                blCancel = True
                            Else
                                   If anali < 2 Then
                                        MsgBox "DISCULPE EL CODIGO NO EXISTE, POR FAVOR INGRESE NUEVAMENTE", vbExclamation
                                        blCancel = True
                                        TXTFIELDS(40).SelStart = 0
                                        TXTFIELDS(40).SelLength = Len(TXTFIELDS(40).Text)
                                    End If
                                    
                         End If
           End If
           
       
        
         Case 56
            
              If MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
                    Dim provee As Integer
                    provee = OPTENER_CODIGO("SELECT COUNT(COD_PROVEEDOR) FROM PROVEEDORES WHERE COD_PROVEEDOR=  " & TXTFIELDS(56).Text & "")
                            If TXTFIELDS(56).Text = "" Then
                                 MsgBox "DISCULPE EL CODIGO ESTA VACIO, POR FAVOR INGRESE NUEVAMENTE", vbExclamation
                                blCancel = True
                            Else
                                   If provee < 2 Then
                                        MsgBox "DISCULPE EL CODIGO NO EXISTE, POR FAVOR INGRESE NUEVAMENTE", vbExclamation
                                        blCancel = True
                                        TXTFIELDS(56).SelStart = 0
                                        TXTFIELDS(56).SelLength = Len(TXTFIELDS(56).Text)
                                    End If
                                    
                         End If
              End If
           
            Case 57
               If MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
                    Dim linea As Integer
                    linea = OPTENER_CODIGO("SELECT COUNT(COD_LINEA) FROM LINEAS WHERE COD_LINEA=  " & TXTFIELDS(57).Text & "")
                            If TXTFIELDS(57).Text = "" Then
                                 MsgBox "DISCULPE EL CODIGO ESTA VACIO, POR FAVOR INGRESE NUEVAMENTE", vbExclamation
                                blCancel = True
                            Else
                                   If linea < 2 Then
                                        MsgBox "DISCULPE EL CODIGO NO EXISTE, POR FAVOR INGRESE NUEVAMENTE", vbExclamation
                                        blCancel = True
                                        TXTFIELDS(57).SelStart = 0
                                        TXTFIELDS(57).SelLength = Len(TXTFIELDS(57).Text)
                                    End If
                                    
                         End If
            End If
                       
      '  Case 58
       '          If MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
        '            Dim ciudadinst As Integer
         '           ciudadinst = OPTENER_CODIGO("SELECT COUNT(CODCIUDAD) FROM CIUDADES WHERE CODCIUDAD=" & TXTFIELDS(58).Text & "")
          '          If TXTFIELDS(58).Text = "" Then
           '                      MsgBox "DISCULPE EL CODIGO ESTA VACIO, POR FAVOR INGRESE NUEVAMENTE", vbExclamation
            '                    blCancel = True
          '   '               Else
           '                        If ciudadinst < 2 Then
           '                            MsgBox "DISCULPE EL CODIGO NO EXISTE, POR FAVOR INGRESE NUEVAMENTE", vbExclamation
           '                             blCancel = True
           '                             TXTFIELDS(58).SelStart = 0
           '                             TXTFIELDS(58).SelLength = Len(TXTFIELDS(58).Text)
           '                         End If
           '
           '              End If
           'End If
        Case 68
                 If MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
                    Dim ciudadprov As Integer
                    ciudadprov = OPTENER_CODIGO("SELECT COUNT(CODCIUDAD) FROM CIUDADES WHERE CODCIUDAD=" & TXTFIELDS(68).Text & "")
                    If TXTFIELDS(68).Text = "" Then
                                 MsgBox "DISCULPE EL CODIGO ESTA VACIO, POR FAVOR INGRESE NUEVAMENTE", vbExclamation
                                blCancel = True
                            Else
                                   If ciudadprov < 2 Then
                                        MsgBox "DISCULPE EL CODIGO NO EXISTE, POR FAVOR INGRESE NUEVAMENTE", vbExclamation
                                        blCancel = True
                                        TXTFIELDS(68).SelStart = 0
                                        TXTFIELDS(68).SelLength = Len(TXTFIELDS(68).Text)
                                    End If
                                    
                         End If
           End If
           
        

        'Case 79
         '   If MDIGALENO.tbToolBar.Buttons("Next").Enabled = False Then
          '        txtfields(79).Text = UCase(txtfields(79).Text)
         
           '         Dim A As Integer
            '        A = OPTENER_CODIGO("SELECT COUNT(CODTAB) FROM TABVAL WHERE CODTAB= '" & txtfields(79).Text & "'")
             '               If A > 1 Then
              '                   MsgBox "DISCULPE EL CODIGO YA EXISTE, POR FAVOR INGRESE NUEVAMENTE", vbExclamation
               '                     blCancel = True
                '                    txtfields(79).SelStart = 0
                 '                   txtfields(79).SelLength = Len(txtfields(79))
                  '
                   '         Else
                    '            If txtfields(79) = "" Then
                     '                MsgBox "DISCULPE EL CODIGO ESTA VACIO, POR FAVOR INGRESE NUEVAMENTE", vbExclamation
                      '               blCancel = True
                       '         End If
                        '    End If
           'End If
                       
    End Select

    SendKeys "{TAB}"
    KeyAscii = 0

Else
       Select Case Index
            Case 1, 25, 26, 33, 35, 42 To 47, 58, 59, 80, 94, 101, 105, 107, 109, 114, 115, 158, 159, 160, 167, 168, 170, 172 'numerico
                VALIDAR.ValidarAscii KeyAscii
            'Case 8  'fecha
             '   VALIDAR.ValidarAscii KeyAscii, False
       End Select
End If

End Sub
Private Sub txtFields_LostFocus(Index As Integer)
    TXTFIELDS(Index).Text = UCase(TXTFIELDS(Index).Text)
End Sub
Private Sub BLOQUEO(X As Boolean)
    Dim oText As TextBox
    For Each oText In Me.TXTFIELDS
           oText.Locked = X
    Next
   
End Sub
Public Sub prcBuscar()
    Call cmdBuscar_Click
End Sub
Public Sub prcPrimero()
    Call cmdPrimero_Click
cambiarImg
End Sub
Public Sub prcAnterior()
    Call cmdAnterior_Click
    cambiarImg
End Sub
Public Sub prcSiguiente()
    Call cmdSiguiente_Click
    cambiarImg
End Sub
Public Sub prcUltimo()
    Call cmdUltimo_Click
    cambiarImg
End Sub
Public Sub prcAgregar()
    Call CmdAgregar_Click
End Sub
Public Sub prcModificar()
    Call cmdModificar_Click
End Sub
Public Sub prcCancelar()
    Call CmdCancelar_Click
    cambiarImg
End Sub
Public Sub prcEliminar()
    Call CmdEliminar_Click
    cambiarImg
End Sub
Public Sub prcGrabar()
    Call cmdGrabar_Click
End Sub
Public Sub prcActualizar()
    On Error Resume Next
    adoPrimaryRS.Requery
End Sub
Private Sub cambiarImg()
On Error Resume Next
If Frame1PERSONA.Visible = True Then
    If Not (adoPrimaryRS.BOF And adoPrimaryRS.EOF) Then
        If adoPrimaryRS.Fields("FOTO").Value <> "" Then
           Set PICPER.Picture = LoadPicture(adoPrimaryRS.Fields("FOTO").Value, vbLPLarge, vbLPColor)
        Else
            Set PICPER.Picture = Nothing
        End If
    End If
End If
End Sub


Private Sub TXTFIELDS_Validate(Index As Integer, Cancel As Boolean)
'     SendKeys "{TAB}"
    'Select Case Index
        'Case 7, 8 '5, 6, 84 'fecha
              '  blCancel = Not VALIDAR.esFecha(TXTFIELDS(Index).Text)
              '  If blCancel = True Then
               '        MsgBox "DISCULPE SOLO SE ACEPTAN FECHAS DEL FORMATO '99/99/99' O '99/99/9999'", vbExclamation
               ' End If
       ' End Select
     'Cancel = blCancel
     'blCancel = False
End Sub
Private Sub prcCaso5y14Agregar(X As Boolean)
Select Case Me.TabStrip1.SelectedItem.Index
Case 1
    'If X = True And MDIGALENO.tbToolBar.Buttons("Add").Enabled = True Then
        TXTFIELDS(0).Locked = True
        TXTFIELDS(5).Locked = True
        TXTFIELDS(6).Locked = True
        TXTFIELDS(7).Locked = True
        TXTFIELDS(10).Locked = True
        TXTFIELDS(11).Locked = True
        TXTFIELDS(12).Locked = True
        TXTFIELDS(14).Locked = True
    'End If
    
Case 2
        TXTFIELDS(123).Locked = True
        
    
Case 4
       TXTFIELDS(56).Locked = True
       TXTFIELDS(57).Locked = True
  
Case 8
        TXTFIELDS(40).Locked = True
        TXTFIELDS(41).Locked = True



Case 9
        TXTFIELDS(85).Locked = True
        
        
Case 11

        TXTFIELDS(61).Locked = True
        TXTFIELDS(62).Locked = True
    
Case 12
        TXTFIELDS(68).Locked = True
        
Case 15
        TXTFIELDS(90).Locked = True
        TXTFIELDS(91).Locked = True
        TXTFIELDS(92).Locked = True
        TXTFIELDS(93).Locked = True
Case 16
        TXTFIELDS(102).Locked = True
        TXTFIELDS(105).Locked = True
        TXTFIELDS(113).Locked = True
        
Case 18
        TXTFIELDS(130).Locked = True
        TXTFIELDS(131).Locked = True
        TXTFIELDS(132).Locked = True
        TXTFIELDS(133).Locked = True
    
End Select
End Sub

