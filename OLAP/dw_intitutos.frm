VERSION 5.00
Object = "{0002E540-0000-0000-C000-000000000046}#1.0#0"; "MSOWC.DLL"
Begin VB.Form dw_Institutos 
   Caption         =   "SGLAC 1.0"
   ClientHeight    =   6165
   ClientLeft      =   1065
   ClientTop       =   2340
   ClientWidth     =   8535
   Icon            =   "dw_intitutos.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   6165
   ScaleWidth      =   8535
   Begin OWC.PivotTable PivotTable1 
      Bindings        =   "dw_intitutos.frx":030A
      Height          =   5760
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7770
      XMLData         =   $"dw_intitutos.frx":0329
   End
End
Attribute VB_Name = "dw_Institutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
dw_Institutos.Height = 6990
dw_Institutos.Width = 9100
dw_Institutos.Top = 0
dw_Institutos.Left = 1500
PivotTable1.MaxHeight = 432
PivotTable1.MaxWidth = 600
PivotTable1.Top = 0
PivotTable1.Left = 0
End Sub

Private Sub Picture1_Click()

End Sub
