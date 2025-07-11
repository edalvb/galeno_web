VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmConsult 
   Caption         =   "CONSULTA"
   ClientHeight    =   2490
   ClientLeft      =   2430
   ClientTop       =   1785
   ClientWidth     =   7755
   Icon            =   "FrmCon.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2490
   ScaleWidth      =   7755
   Begin MSDataGridLib.DataGrid msGrid 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   5530
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
         Weight          =   400
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
         SizeMode        =   1
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   960
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCon.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCon.frx":041C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCon.frx":0526
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCon.frx":06F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCon.frx":08BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCon.frx":09CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCon.frx":0AE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCon.frx":0BF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCon.frx":0D0A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TlbConsulta 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "BuscarPrev"
            Object.ToolTipText     =   "Previo"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Buscar"
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "BuscarSig"
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Filtrar"
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Pegar"
            Object.ToolTipText     =   "Pegar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Primero"
            Object.ToolTipText     =   "Ir al Primero"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Anterior"
            Object.ToolTipText     =   "Ir al Anterior"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Siguiente"
            Object.ToolTipText     =   "Ir al Siguiente"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ultimo"
            Object.ToolTipText     =   "Ir al Ultimo"
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu mnuFiltro 
      Caption         =   "&Filtro Avanzado"
      Begin VB.Menu mnuFMultiple 
         Caption         =   " Filtro &Multiple"
      End
      Begin VB.Menu mnuFiltroOr 
         Caption         =   "Filtro &Or"
      End
   End
End
Attribute VB_Name = "frmConsult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conexion As ADODB.Connection
Dim strCadenaIngresada As String
Dim STRSELECCION As String
Dim strOrder As String
Dim rsConsulta As ADODB.Recordset
Dim strFiltroAvanzado As String
Dim strFiltro As String
Dim strSqlSentencia As String
Dim strOrden As String
Dim blUnload As Boolean
Dim WithEvents Filtro As frmFiltro
Attribute Filtro.VB_VarHelpID = -1

Public Property Set ActiveConection(cl As ADODB.Connection)
'falta la rurina de control de error
Set conexion = cl

End Property

Public Property Let Consulta(Sql As String)
Dim strMensaje As String
'guarda la cadena que sera utilizado para ordenar
strSqlSentencia = Sql
'Se inicia el recordeset en caso que no lo este
If rsConsulta Is Nothing Then
    Set rsConsulta = New Recordset
    rsConsulta.CursorLocation = adUseClient
Else
 'en caso de estar iniciado comprueba si esta abierto
 'el recordset, en caso que lo este lo cierra puesto que lo habrira más abajo
    If rsConsulta.State = adStateOpen Then
        rsConsulta.Close
    End If
End If

On Error GoTo errControl
'si se envio la sentencia SQL y se asigno la conexion a ConConsulta, habre el recordset
If Not (Sql = "") Then
  If Not (conexion Is Nothing) Then
        rsConsulta.Open Sql & strOrden, conexion, adOpenStatic, adLockReadOnly, adCmdText
        Set msGrid.DataSource = rsConsulta
        Call Alinear
  Else
    'deberia dipararse un error hacia el usuario
    MsgBox "Primero defina la propiedad ActiveConection " & _
    vbCrLf & "y luego la propiedad consulta del objeto Consulta", vbInformation, "BUSCAR"
  End If
End If
Call TamanoCelda
Exit Property
errControl:

If Not (conexion Is Nothing) = True Then
    Call ErrorObjeto
Else
    MsgBox Err.Description & "" & Err.Number, vbCritical, "Error"
    Err.Raise Number:=vbObjectError + 1, Source:="Consulta1", Description:="No se ha establecido la conexion antes de la consulta"
End If
End Property

Public Property Get Seleccion() As String
'propiedad de lectura
Seleccion = STRSELECCION
End Property

Private Sub filtrar(strFiltrar As String)
 On Error Resume Next
If RTrim$(strFiltrar) <> "" Then
      strFiltro = ArmarCadenaConsulta(strFiltrar, True)
      rsConsulta.Filter = strFiltro
Else
    strFiltro = ""
    rsConsulta.Filter = ""
End If
msGrid.HoldFields
Set msGrid.DataSource = rsConsulta
Exit Sub
End Sub

Private Sub Filtro_CloseFilter(strFiltro As Variant)
Dim strFiltTemp As String
strFiltTemp = strFiltroAvanzado

If mnuFiltroOr.Checked Then
    If strFiltroAvanzado = "" Then
        strFiltroAvanzado = strFiltro
    Else
        strFiltroAvanzado = strFiltroAvanzado & " or " & strFiltro
    End If
Else
    If strFiltroAvanzado = "" Then
        strFiltroAvanzado = strFiltro
    Else
        strFiltroAvanzado = strFiltroAvanzado & " and " & strFiltro
    End If
End If

Debug.Print strFiltroAvanzado
On Error Resume Next
rsConsulta.Filter = strFiltroAvanzado
If Err.Number > 0 Then
    strFiltroAvanzado = strFiltTemp
    MsgBox "El valor que introdujo, no es correcto", vbInformation
    rsConsulta.Filter = strFiltroAvanzado
End If
    
Set Filtro = Nothing

End Sub

Private Sub Form_Initialize()
On Error GoTo RenovarErr

RenovarErr:
 'MsgBox "una mierd"

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'On Error GoTo RenovarErr

  
If Not (rsConsulta.EOF And rsConsulta.BOF) Then
    If KeyAscii = 13 And TlbConsulta.Buttons(7).Visible Then
        STRSELECCION = STRSELECCION + " " + msGrid.Columns(0).Value
        Call Animar
        'Me.Visible = False
    End If
Else
    Unload Me
End If

'RenovarErr:
' MsgBox Err.Number
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'Al presionar la tecla ESC se descarga el formulario consulta
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub Form_Resize()
On Error GoTo RenovarErr

RenovarErr:
'  MsgBox Err.Number
'Centra la grilla en el formulario Consulta
'On Error Resume Next
'msGrid.Left = (Me.ScaleWidth / 2) - (msGrid.Width / 2)
'msGrid.Top = (Me.ScaleHeight / 2) - (msGrid.Height / 2)
msGrid.Height = Me.ScaleHeight - (TlbConsulta.Height + 100)
End Sub

Private Sub TamanoCelda()
'Cambia el tamaño de las celdas del grid
On Error Resume Next
Dim i As Integer
Dim lngLongitud As Long
For i = 0 To msGrid.Columns.Count - 1
    
    lngLongitud = rsConsulta(i).DefinedSize * 120
    If lngLongitud < 1750 Then
        msGrid.Columns(i).Width = 1750
    Else
        msGrid.Columns(i).Width = lngLongitud
    End If
Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo RenovarErr

RenovarErr:
'  MsgBox Err.Number
'On Error Resume Next
    Set msGrid.DataSource = Nothing
    If Not (rsConsulta Is Nothing) Then
        If rsConsulta.State = adStateOpen Then
            rsConsulta.Close 'cierra el recorset
            Set rsConsulta = Nothing
        End If
    End If
    blUnload = True
End Sub

Private Sub mnuFiltroOr_Click()
mnuFiltroOr.Checked = Not mnuFiltroOr.Checked
End Sub

Private Sub mnuFMultiple_Click()
mnuFMultiple.Checked = Not mnuFMultiple.Checked
If Not mnuFMultiple.Checked Then
    strFiltroAvanzado = ""
    rsConsulta.Filter = ""
    strFiltro = ""
    msGrid.HoldFields
    Set msGrid.DataSource = rsConsulta
End If
End Sub

Private Sub msGrid_Click()

'MsgBox msGrid.Col & " " & msGrid.Row
'If msGrid.Col = -1 And msGrid.Row > 0 Then
'        STRSELECCION = STRSELECCION + " " + msGrid.SelBookmarks
'End If
End Sub

Private Sub msGrid_DblClick()
'MsgBox msGrid.Col & " " & msGrid.Row
'If msGrid.Col = -1 And msGrid.Row > 0 Then
'        STRSELECCION = STRSELECCION + " " + msGrid.Columns(0).Value
'        Me.Visible = False
'End If

End Sub

Private Sub msGrid_Error(ByVal DataError As Integer, Response As Integer)
Response = 0
End Sub

Private Sub msGrid_HeadClick(ByVal ColIndex As Integer)
Static A As Boolean
Static lastCol As Integer

    If lastCol = ColIndex Then
        A = Not A
    Else
        A = False
    End If
    If Not A Then
        strOrder = rsConsulta(ColIndex).Name & " ASC"
        rsConsulta.Sort = strOrder
    Else
        strOrder = rsConsulta(ColIndex).Name & " DESC"
        rsConsulta.Sort = strOrder
    End If
       lastCol = ColIndex
       msGrid.ClearSelCols
       msGrid.Col = ColIndex
End Sub

Private Sub ErrorObjeto()
Dim errColeccion As ADODB.Errors
Dim errCon As ADODB.Error
Dim strMensaje  As String
'falta tratar y encapsular el error y disparar hacia el lado del cliente "err.Raise"
Set errColeccion = conexion.Errors
For Each errCon In errColeccion
    strMensaje = errCon.Description & vbCrLf _
    & errCon.Source & vbCrLf & errCon.Source
Next
MsgBox strMensaje, vbCritical

End Sub

Private Sub msGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'Controla las teclas presionadas
If KeyCode = vbKeyF5 Then 'Filtrar
    If mnuFMultiple.Checked Then
        Call FiltroAvanzado
    Else
        strCadenaIngresada = InputBox("Ingrese el valor de la columna " & msGrid.Columns(msGrid.Col).Caption & " ha ser filtrada", "Filtrar")
        filtrar (strCadenaIngresada)
    End If

Else
   If Not (rsConsulta.EOF And rsConsulta.BOF) Then
        Select Case KeyCode
            Case vbKeyF6 'Buscar registro anterior (Se activa si el usuario presiono F7)
              If TlbConsulta.Buttons(1).Visible Then
                BuscarAnterior (strCadenaIngresada)
              End If
            
            Case vbKeyF7 'Buscar registro
                strCadenaIngresada = InputBox("Ingrese el valor de la columna " & msGrid.Columns(msGrid.Col).Caption & " ha ser buscado", "BUSCAR")
                BUSCAR (strCadenaIngresada)
            
            Case vbKeyF8 'Buscar siguiente registro (Se activa si el usuario presiono F7)
              If TlbConsulta.Buttons(3).Visible Then
                BuscarSiguiente (strCadenaIngresada)
              End If
            
            Case vbKeyHome 'Ir al primer registro
                rsConsulta.MoveFirst
                            
            Case vbKeyEnd 'Ir al ultimo registro
                rsConsulta.MoveLast
        
        End Select
    End If
End If
End Sub

Private Sub TlbConsulta_ButtonClick(ByVal Button As MSComctlLib.Button)
'Verifica que boton se presiono en la barra de herramientas del formulario consulta

    If rsConsulta.EOF And rsConsulta.BOF Then
        'MsgBox "No existe Registros", vbInformation, "BUSCAR"
        Call filtrar("")
        strFiltro = ""
    Else
        Select Case Button.Key
             Case "Buscar"
                 strCadenaIngresada = InputBox("Ingrese el valor de la columna " & msGrid.Columns(msGrid.Col).Caption & " ha ser buscado", "Buscar")
                 BUSCAR (strCadenaIngresada)
             Case "BuscarPrev"
                 BuscarAnterior (strCadenaIngresada)
             Case "BuscarSig"
                 BuscarSiguiente (strCadenaIngresada)
             Case "Filtrar"
                 If mnuFMultiple.Checked Then
                    Call FiltroAvanzado
                 Else
                    strCadenaIngresada = InputBox("Ingrese el valor de la columna " & msGrid.Columns(msGrid.Col).Caption & " ha ser filtrada", "Filtrar")
                    filtrar (strCadenaIngresada)
                 End If
             Case "Pegar"
                 Dim i As Integer ' Counter
                 Dim intCount As Integer
                 intCount = msGrid.SelBookmarks.Count - 1
                 For i = 0 To intCount
                      STRSELECCION = STRSELECCION + " " + msGrid.Columns(0).CellText(msGrid.SelBookmarks(i))
                       ' Realiza operaciones aquí. Si las operaciones deben ser
                       ' canceladas, sale del bucle y usa la matriz para
                       ' deshacer los cambios.
                 Next i
                 'STRSELECCION = STRSELECCION + " " + msGrid.Columns(0).Value
                 Call Animar
                 Me.Visible = False
             Case "Primero"
                 rsConsulta.MoveFirst
             Case "Anterior"
                 If Not (rsConsulta.BOF) Then rsConsulta.MovePrevious
                 If rsConsulta.BOF Then rsConsulta.MoveFirst
             Case "Siguiente"
                 If Not (rsConsulta.EOF) Then rsConsulta.MoveNext
                 If rsConsulta.EOF Then rsConsulta.MoveLast
             Case "Ultimo"
                 rsConsulta.MoveLast
        End Select
    End If

Exit Sub
End Sub

Private Sub BUSCAR(strBuscar As String)
Dim Res As String
Dim Marcador As Double
ValidarDatIngresado strBuscar
'On Error Resume Next
If strBuscar <> "" Then
    If Not (rsConsulta.EOF And rsConsulta.BOF) Then
       Marcador = rsConsulta.Bookmark
        rsConsulta.MoveFirst
        Res = ArmarCadenaConsulta(strBuscar)
        
        If strBuscar <> "" Then
           
            rsConsulta.Find Res, , adSearchForward
        End If
        If Not (rsConsulta.EOF) Then
            If rsConsulta(msGrid.Col).Value = strBuscar Then
                 TlbConsulta.Buttons(3).Visible = True
                 TlbConsulta.Buttons(1).Visible = True
                ' rsConsulta.Bookmark = Marcador
            End If
        Else
            rsConsulta.Bookmark = Marcador
            If Len(strBuscar) = 1 Then
                If Asc(strBuscar) >= 48 And Asc(strBuscar) <= 57 Then
                    Res = "El Número"
                Else
                    Res = "La Letra"
                End If
            Else
                If Asc(strBuscar) >= 48 And Asc(strBuscar) <= 57 Then
                    Res = "El Número"
                Else
                    Res = "La Palabra"
                End If
            End If
            TlbConsulta.Buttons(1).Visible = False
            TlbConsulta.Buttons(3).Visible = False
            Res = "" & Res & " " & strBuscar & " No existe dentro" & vbCrLf & "de la Columna " & _
            msGrid.Columns(msGrid.Col).Caption
            MsgBox Res, 64, "Resultado de la busqueda"
            'rsConsulta.MoveFirst
            strBuscar = ""
        End If
    End If
End If
Exit Sub
End Sub

Private Sub BuscarSiguiente(strBuscar As String)
Dim Tem As String
Dim Marcador As Double
On Error Resume Next
If strBuscar <> "" Then
    
    If Not rsConsulta.EOF Then
        'rsConsulta.MoveNext
        Marcador = rsConsulta.Bookmark
        Tem = strBuscar
        If Tem <> "" Then
          rsConsulta.Find Tem, 1, adSearchForward
          
        End If
    End If
    If rsConsulta.EOF Or Marcador = rsConsulta.Bookmark Then
        MsgBox "No existen más registros", vbInformation, Me.Caption
        'Call BuscarAnterior(strBuscar)
        TlbConsulta.Buttons(1).Visible = False
        TlbConsulta.Buttons(3).Visible = False
        strBuscar = ""
    End If
End If
Exit Sub
End Sub

Private Sub BuscarAnterior(strBuscar As String)
Dim Tem As String
Dim Marcador As Double
On Error Resume Next
If strBuscar <> "" Then
    If Not (rsConsulta.BOF) Then
    Tem = ArmarCadenaConsulta(strBuscar)
        'rsConsulta.MovePrevious
        If Tem <> "" Then
            
            Marcador = rsConsulta.Bookmark
            rsConsulta.Find Tem, -1, adSearchBackward
        
        End If
    End If
    If rsConsulta.BOF Or Marcador = rsConsulta.Bookmark Then
        MsgBox "No existen mas registros", vbInformation, Me.Caption
    '    Call BuscarSiguiente(strBuscar)
        TlbConsulta.Buttons(1).Visible = False
        TlbConsulta.Buttons(3).Visible = False
        strBuscar = ""
    End If
End If
Exit Sub
End Sub

Private Function ArmarCadenaConsulta(Cadena As String, Optional filtrar As Boolean = False) As String
Select Case rsConsulta(msGrid.Col).Type
    Case DataTypeEnum.adDate, DataTypeEnum.adDBDate, DataTypeEnum.adDBTimeStamp 'SI ES FECHA
        If IsDate(Cadena) Then
            ArmarCadenaConsulta = msGrid.Columns(msGrid.Col).DataField & _
                                                    " = #" & Trim$(Cadena) & "#"
       Else
            ArmarCadenaConsulta = ""
       End If
   
   Case DataTypeEnum.adChapter, DataTypeEnum.adChar, _
            DataTypeEnum.adLongVarChar, DataTypeEnum.adLongVarWChar, _
            DataTypeEnum.adVarWChar, DataTypeEnum.adWChar  'SI ES CADENA
            ArmarCadenaConsulta = msGrid.Columns(msGrid.Col).DataField & " like '%" & Cadena & "%'"
    Case Else 'SI ES NUMERICO
        If IsNumeric(Cadena) Then
            ArmarCadenaConsulta = msGrid.Columns(msGrid.Col).DataField & _
                                                      " = " & Cadena
        Else
            ArmarCadenaConsulta = ""
        End If
End Select
    
End Function

Public Property Let Orden(ByVal vNewValue As String)
'Especifica la columna que sera ordenada de entrada
If vNewValue <> "" Then
    strOrden = " order by " & vNewValue
End If
End Property

Private Sub Alinear()
Dim i As Long
On Error GoTo errManejador
For i = 0 To msGrid.Columns.Count - 1
    If rsConsulta(i).Type = adNumeric Then
        msGrid.Columns(i).Alignment = dbgCenter
    End If
Next
Exit Sub
errManejador:
MsgBox Err.Description & "" & Err.Number, vbCritical, "Error"
End Sub

Public Sub Cerrar()
    Unload Me
End Sub

Public Property Let Modal(ByVal vNewValue As Boolean)
TlbConsulta.Buttons(6).Visible = vNewValue
TlbConsulta.Buttons(7).Visible = vNewValue
End Property

Public Property Get Descargado() As Boolean
Descargado = blUnload
End Property

Private Sub ValidarDatIngresado(Cadena As String)

Select Case rsConsulta(msGrid.Col).Type
               
        Case DataTypeEnum.adChapter, DataTypeEnum.adVarWChar, DataTypeEnum.adChar
        Case DataTypeEnum.adDate, DataTypeEnum.adDBDate
                If Not IsDate(Cadena) Then Cadena = ""
        Case DataTypeEnum.adNumeric, DataTypeEnum.adBoolean, DataTypeEnum.adBigInt, DataTypeEnum.adDecimal, DataTypeEnum.adDouble, DataTypeEnum.adInteger, DataTypeEnum.adVarNumeric
                If Not IsNumeric(Cadena) Then Cadena = ""
        Case Else
                Cadena = ""
End Select
End Sub
Private Sub FiltroAvanzado()
Dim Pos As RECT
On Error Resume Next
GetWindowRect Me.hwnd, Pos
Set Filtro = New frmFiltro
Load Filtro
ArmarCadenaConsulta "" 'BUSCA EL TIPO
Filtro.TipoDato = GetType()
Filtro.Campo = rsConsulta(msGrid.Col).Name 'PASA EL CAMPO
SetWindowPos Me.hwnd, conHwndNoTopmost, Pos.Left, Pos.Top, (Me.Width / Screen.TwipsPerPixelX), (Me.Height / Screen.TwipsPerPixelY), conSwpNoActivate Or conSwpShowWindow
Filtro.Caption = Filtro.Caption & "-" & rsConsulta(msGrid.Col).Name
Filtro.Show vbModal
End Sub

Private Function GetType()
Select Case rsConsulta(msGrid.Col).Type
    Case DataTypeEnum.adDate, DataTypeEnum.adDBDate, DataTypeEnum.adDBTimeStamp 'SI ES FECHA
            GetType = 2
    Case DataTypeEnum.adChapter, DataTypeEnum.adChar, _
            DataTypeEnum.adLongVarChar, DataTypeEnum.adLongVarWChar, _
            DataTypeEnum.adVarWChar, DataTypeEnum.adWChar  'SI ES CADENA
            GetType = 0
    Case Else 'SI ES NUMERICO
            GetType = 1
End Select

End Function

Public Property Get GetFilterString() As String
If strFiltroAvanzado = "" Then
    GetFilterString = strFiltro
Else
    GetFilterString = strFiltroAvanzado
End If
End Property

Public Property Get GetOrder() As String
GetOrder = strOrder
End Property

Private Sub Animar()

    'If SysVersions32 >= 4 Then
    '    AnimateWindow Me.hwnd, 40, &H90000
    'End If
     Me.Visible = False

End Sub
