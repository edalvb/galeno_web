VERSION 5.00
Begin VB.Form EDITOR 
   Caption         =   "EDITOR DE ANTECEDENTES"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.CommandButton CANCELAR 
      Caption         =   "CANCELAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton ACEPTAR 
      Caption         =   "ACEPTAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox MEMO 
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "EDITOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Cerrar()
    Unload Me
End Sub

Public Property Get Descargado() As Boolean
Descargado = blUnload
End Property

Private Sub ACEPTAR_Click()
Me.Visible = False
End Sub

Private Sub CANCELAR_Click()
Me.MEMO.Text = ""
Me.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
 blUnload = True
End Sub
