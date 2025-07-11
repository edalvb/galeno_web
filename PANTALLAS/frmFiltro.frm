VERSION 5.00
Begin VB.Form frmFiltro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filtro"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5010
   Icon            =   "frmFiltro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   390
      Left            =   3705
      TabIndex        =   4
      Top             =   1140
      Width           =   915
   End
   Begin VB.TextBox txtParam 
      Height          =   315
      Index           =   1
      Left            =   3120
      TabIndex        =   2
      Top             =   645
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.ComboBox cmb 
      Height          =   315
      Left            =   1515
      TabIndex        =   1
      Text            =   "="
      Top             =   645
      Width           =   1455
   End
   Begin VB.TextBox txtParam 
      Height          =   315
      Index           =   0
      Left            =   195
      TabIndex        =   0
      Top             =   660
      Width           =   1215
   End
   Begin VB.Label lblCaption 
      Caption         =   "Primer Valor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   7
      Top             =   375
      Width           =   1140
   End
   Begin VB.Label lblCaption 
      Caption         =   "Segundo Valor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   2
      Left            =   3135
      TabIndex        =   6
      Top             =   375
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label lblCaption 
      Caption         =   "Campo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   255
      Index           =   3
      Left            =   210
      TabIndex        =   5
      Top             =   660
      Width           =   1140
   End
   Begin VB.Label lblCaption 
      Caption         =   "Operador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   1
      Left            =   1755
      TabIndex        =   3
      Top             =   375
      Width           =   915
   End
End
Attribute VB_Name = "frmFiltro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intTipoDato As Integer
Dim strCampo As String 'el campo Sql
Dim WithEvents ValidarTecla As ValTecla
Attribute ValidarTecla.VB_VarHelpID = -1
Public Event CloseFilter(strFiltro)
Const Fecha = 2
Const Numero = 1
Const Cadena = 0
Private Sub cmb_Click()
If cmb.ListIndex = 6 Then
    txtParam(0).Left = 195
    
    lblCaption(0).Left = 180
    
    txtParam(1).Visible = True
    lblCaption(2).Visible = True
Else
    txtParam(0).Left = txtParam(1).Left
        
    lblCaption(0).Left = lblCaption(2).Left
      
    lblCaption(2).Visible = False
    txtParam(1).Visible = False
End If
End Sub

Private Sub cmb_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmdAceptar_Click()
Dim bCancel As Boolean
bCancel = False
    If cmb.ListIndex = 6 Then
            
        Select Case intTipoDato
            Case Cadena
                If txtParam(0) > txtParam(1) Then bCancel = True
            Case Numero
                If Val(txtParam(0)) > Val(txtParam(1)) Then bCancel = True
            Case Fecha
                If IsDate(txtParam(0)) Then
                    If IsDate(txtParam(1)) Then
                        If CDate(txtParam(0)) > CDate(txtParam(1)) Then bCancel = True
                    Else
                        bCancel = True
                    End If
                Else
                    bCancel = True
                End If
        End Select

         If bCancel Then
            MsgBox "El Primer Valor debe ser menor que el Segundo Valor" & vbCrLf _
                & "Por favor, reingrese el Primer Valor", vbInformation
           txtParam(0).SetFocus
           txtParam(0).SelStart = 0
           txtParam(0).SelLength = Len(txtParam(0))
         End If

    End If

If Not bCancel Then
    GetFilterString
    If SysVersions32 >= 4 Then
        AnimateWindow Me.hwnd, 200, &H90000
    End If
    Unload Me
    
End If

End Sub

Private Sub Form_Load()
cmb.AddItem "Igual", 0
cmb.AddItem "Mayor que", 1
cmb.AddItem "Menor que", 2
cmb.AddItem "Mayor o Igual", 3
cmb.AddItem "Menor o Igual", 4
cmb.AddItem "Distinto", 5
cmb.AddItem "Entre", 6
cmb.AddItem "Que suene a", 7
cmb.ListIndex = 0
Set ValidarTecla = New ValTecla
    txtParam(0).Left = txtParam(1).Left
        
    lblCaption(0).Left = lblCaption(2).Left

End Sub

Public Property Let TipoDato(ByVal vNewValue As Integer)
intTipoDato = vNewValue
End Property

Private Sub GetFilterString()
Dim strSQL As String 'cadena que contiene la sentencia Sql

Select Case intTipoDato
    Case Cadena 'STRING
        strSQL = ArmarString()
    Case Numero 'NUMERICO
        strSQL = ArmarNumero()
    Case Fecha 'FECHA
        strSQL = ArmarFecha()
End Select
'Debug.Print strSQL

RaiseEvent CloseFilter(strSQL)
End Sub

Private Function ArmarString() As String

Select Case cmb.ListIndex
    Case 0 'igual
        ArmarString = strCampo & " = '" & txtParam(0).Text & "'"
    Case 1 'mayor
        ArmarString = strCampo & " > '" & txtParam(0).Text & "'"
    Case 2 'menor
        ArmarString = strCampo & " < '" & txtParam(0).Text & "'"
    Case 3 'mayor igual
        ArmarString = strCampo & " >= '" & txtParam(0).Text & "'"
    Case 4 'menor igual
        ArmarString = strCampo & " <= '" & txtParam(0).Text & "'"
    Case 5 'distinto
        ArmarString = strCampo & " <> '" & txtParam(0).Text & "'"
    Case 6 'entre
        ArmarString = strCampo & " >= '" & txtParam(0).Text & "' and " & strCampo & " <= '" & txtParam(1).Text & "'"
    Case 7 'like
        ArmarString = strCampo & " like '%" & txtParam(0).Text & "%'"
    Case Else 'igual
        ArmarString = strCampo & " = '" & txtParam(0).Text & "'"
    
End Select
End Function

Private Function ArmarNumero() As String
Select Case cmb.ListIndex
    Case 0 'igual
        ArmarNumero = strCampo & " = " & txtParam(0).Text
    Case 1 'mayor
        ArmarNumero = strCampo & " > " & txtParam(0).Text
    Case 2 'menor
        ArmarNumero = strCampo & " < " & txtParam(0).Text
    Case 3 'mayor igual
        ArmarNumero = strCampo & " >= " & txtParam(0).Text
    Case 4 'menor igual
        ArmarNumero = strCampo & " <= " & txtParam(0).Text
    Case 5 'distinto
        ArmarNumero = strCampo & " <> " & txtParam(0).Text
    Case 6 'entre
        ArmarNumero = strCampo & " >= " & txtParam(0).Text & " and " & strCampo & " <= " & txtParam(1).Text
    Case 7 'like
        ArmarNumero = strCampo & " like " & txtParam(0).Text
    Case Else 'igual
        ArmarNumero = strCampo & " = " & txtParam(0).Text
End Select

End Function

Private Function ArmarFecha() As String
Select Case cmb.ListIndex
    Case 0 'igual
        ArmarFecha = strCampo & " = #" & txtParam(0).Text & "#"
    Case 1 'mayor
        ArmarFecha = strCampo & " > #" & txtParam(0).Text & "#"
    Case 2 'menor
        ArmarFecha = strCampo & " < #" & txtParam(0).Text & "#"
    Case 3 'mayor igual
        ArmarFecha = strCampo & " >= #" & txtParam(0).Text & "#"
    Case 4 'menor igual
        ArmarFecha = strCampo & " <= #" & txtParam(0).Text & "#"
    Case 5 'distinto
        ArmarFecha = strCampo & " <> #" & txtParam(0).Text & "#"
    Case 6 'entre
        ArmarFecha = strCampo & " >= #" & txtParam(0).Text & "# and " & strCampo & " <= #" & txtParam(1).Text & "#"
    Case 7 'like
        ArmarFecha = strCampo & " like #" & txtParam(0).Text & "#"
    Case Else 'igual
        ArmarFecha = strCampo & " = #" & txtParam(0).Text & "#"
    
End Select

End Function

Public Property Let Campo(ByVal vNewValue As String)
strCampo = vNewValue
lblCaption(3).Caption = strCampo
End Property

Private Sub txtParam_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
Else
    With ValidarTecla
        Select Case intTipoDato
            Case 1 'numero
                .ValidarAscii KeyAscii
            Case 2 'fecha
                .ValidarAscii KeyAscii, False
        End Select
    End With
End If
End Sub

Private Sub txtParam_Validate(Index As Integer, Cancel As Boolean)
If intTipoDato = 2 Then
    Cancel = Not ValidarTecla.esFecha(txtParam(Index).Text)
    If Cancel Then txtParam(Index).SelStart = 1: txtParam(Index).SelLength = Len(txtParam(Index).Text)
End If

End Sub

Private Sub ValidarTecla_ERROR(strDescrip As String)
MsgBox strDescrip, vbInformation
End Sub
