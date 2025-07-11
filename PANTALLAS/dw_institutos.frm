VERSION 5.00
Object = "{0002E540-0000-0000-C000-000000000046}#1.0#0"; "MSOWC.DLL"
Begin VB.Form dw_institutos 
   Caption         =   "SERVICIOS OLAP DE GALENO POR INSTITUTOS"
   ClientHeight    =   6795
   ClientLeft      =   1065
   ClientTop       =   2340
   ClientWidth     =   8535
   Icon            =   "dw_institutos.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   8535
   Begin OWC.PivotTable DW_INST 
      Bindings        =   "dw_institutos.frx":030A
      Height          =   2625
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7770
      XMLData         =   $"dw_institutos.frx":0323
   End
End
Attribute VB_Name = "dw_institutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'dw_institutos.Height = 6990
'dw_institutos.Width = 9100
'dw_institutos.Top = 0
'dw_institutos.Left = 1500
'DW_INST.MaxHeight = 432
'DW_INST.MaxWidth = 600
'DW_INST.Top = 0
'DW_INST.Left = 0
End Sub


