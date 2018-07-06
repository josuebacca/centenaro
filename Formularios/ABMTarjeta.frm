VERSION 5.00
Begin VB.Form ABMTarjeta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos de la Tarjeta..."
   ClientHeight    =   3465
   ClientLeft      =   2700
   ClientTop       =   2625
   ClientWidth     =   5520
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8,25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ABMTarjeta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboTipo 
      Height          =   315
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   690
      Width           =   1545
   End
   Begin VB.TextBox txtContrasena 
      Height          =   330
      Left            =   1020
      MaxLength       =   50
      TabIndex        =   6
      Top             =   2475
      Width           =   4275
   End
   Begin VB.TextBox txtNroCta 
      Height          =   330
      Left            =   1020
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1755
      Width           =   4275
   End
   Begin VB.TextBox txtUsuario 
      Height          =   330
      Left            =   1020
      MaxLength       =   50
      TabIndex        =   5
      Top             =   2115
      Width           =   4275
   End
   Begin VB.TextBox txtTelefono 
      Height          =   330
      Left            =   1020
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1395
      Width           =   4275
   End
   Begin VB.TextBox txtDescri 
      Height          =   330
      Left            =   1020
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1035
      Width           =   4275
   End
   Begin VB.TextBox txtID 
      Height          =   330
      Left            =   1020
      TabIndex        =   0
      Top             =   225
      Width           =   930
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   345
      Left            =   3990
      TabIndex        =   8
      Top             =   3045
      Width           =   1300
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   2640
      TabIndex        =   7
      Top             =   3045
      Width           =   1300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tipo:"
      Height          =   195
      Index           =   6
      Left            =   105
      TabIndex        =   15
      Top             =   750
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Contrase�a:"
      Height          =   195
      Index           =   5
      Left            =   105
      TabIndex        =   14
      Top             =   2535
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nro. Cta:"
      Height          =   195
      Index           =   4
      Left            =   105
      TabIndex        =   13
      Top             =   1815
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Usuario:"
      Height          =   195
      Index           =   3
      Left            =   105
      TabIndex        =   12
      Top             =   2175
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tel�fono:"
      Height          =   195
      Index           =   2
      Left            =   105
      TabIndex        =   11
      Top             =   1455
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Descripci�n:"
      Height          =   195
      Index           =   1
      Left            =   105
      TabIndex        =   10
      Top             =   1110
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Id.:"
      Height          =   195
      Index           =   0
      Left            =   105
      TabIndex        =   9
      Top             =   270
      Width           =   270
   End
End
Attribute VB_Name = "ABMTarjeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'parametros para la configuraci�n de la ventana de datos
Dim vFieldID As String
Dim vStringSQL As String
Dim vFormLlama As Form
Public vMode As Integer
Dim vListView As ListView
Dim vDesFieldID As String


'constantes para funcionalidad de uso del formulario
Const cSugerirID = True 'si es True si sugiere un identificador cuando deja el campo en blanco
Const cTabla = "TARJETA"
Const cCampoID = "TAR_CODIGO"
Const cDesRegistro = "Tarjeta"

Function ActualizarListaBase(pMode As Integer)
    On Error GoTo moco
    Dim rec As ADODB.Recordset
    Dim cSQL As String
    Dim I As Integer
    Dim auxListItem As ListItem
    Dim IndiceCampoID As Integer
    Dim OrdenCampo As Integer
    Dim f As ADODB.Field
    Set rec = New ADODB.Recordset
    
    'armo la cadena a ejecutar
    If InStr(1, vStringSQL, "WHERE") = 0 Then
        cSQL = vStringSQL & " WHERE " & cCampoID & " = " & txtID.Text
    Else
        cSQL = vStringSQL & " AND " & cCampoID & " = " & txtID.Text
    End If
    
    If pMode = 4 Then
        vListView.ListItems.Remove vListView.SelectedItem.Index
        Exit Function
    End If
    
    rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If (rec.BOF And rec.EOF) = 0 Then
        If rec.EOF = False Then
        
'            'busco el indce del campo identificador
            OrdenCampo = 0
            IndiceCampoID = 0
            For Each f In rec.Fields
                OrdenCampo = OrdenCampo + 1
                If UCase(f.Name) = UCase(vDesFieldID) Then
                    IndiceCampoID = OrdenCampo - 1
                End If
            Next f
        
            'recorro la coleci�n de campos a actualizar
            For I = 0 To rec.Fields.Count - 1
                If I = 0 Then
                    Select Case pMode
                        Case 1
                            Set auxListItem = vListView.ListItems.Add(, "'" & rec.Fields(IndiceCampoID) & "'", CStr(IIf(IsNull(rec.Fields(I)), "", rec.Fields(I))), 1)
                            auxListItem.Icon = 1
                            auxListItem.SmallIcon = 1
                            
                        Case 2
                            Set auxListItem = vListView.SelectedItem
                            auxListItem.Text = rec.Fields(I)
                    End Select
                Else
                    auxListItem.SubItems(I) = IIf(IsNull(rec.Fields(I)), "", rec.Fields(I))
                End If
            Next I
        End If
    End If
    Exit Function
moco:
    If Err.Number = 35613 Then
        Call Menu.mnuContextABM_Click(4)
    End If
End Function

Function SetMode(pMode As Integer)

    'Configura los controles del form segun el parametro pMode
    'Parametro: pMode indica el modo en que se utilizar� este form
    '  pMode  =             1> Indica nuevo registro
    '                       2> Editar registro existente
    '                       3> Mostrar dato del registro existente
    '                       4> Eliminar registro existente
    
    
    Select Case pMode
        Case 1, 2
            AcCtrl TxtDescri
            AcCtrl txtTelefono
            AcCtrl txtNroCta
            AcCtrl txtUsuario
            AcCtrl txtContrasena
            
        Case 3, 4
            DesacCtrl TxtDescri
            DesacCtrl txtTelefono
            DesacCtrl txtNroCta
            DesacCtrl txtUsuario
            DesacCtrl txtContrasena
    End Select
    
    
    Select Case pMode
        Case 1
            cmdAceptar.Enabled = False
            Me.Caption = "Nueva " & cDesRegistro & "..."
            txtID_LostFocus
            DesacCtrl txtID
            
        Case 2
            cmdAceptar.Enabled = False
            Me.Caption = "Editando " & cDesRegistro & "..."
            DesacCtrl txtID
            DesacCtrl cboTipo
        Case 3
            cmdAceptar.Visible = False
            Me.Caption = "Datos de la " & cDesRegistro & "..."
            DesacCtrl txtID
            DesacCtrl cboTipo
        Case 4
            cmdAceptar.Enabled = True
            Me.Caption = "Eliminando " & cDesRegistro & " ..."
            DesacCtrl txtID
            DesacCtrl cboTipo
    End Select
End Function

Public Function SetWindow(pWindow As Form, pSQL As String, pMode As Integer, pListview As ListView, pDesID As String)
    
    Set vFormLlama = pWindow 'Objeto ventana que que llama a la ventana de datos
    vStringSQL = pSQL 'string utilizado para argar la lista base
    vMode = pMode  'modo en que se utilizar� la ventana de datos
    Set vListView = pListview 'objeto listview que se est� editando
    vDesFieldID = pDesID 'nombre del campo identificador
    
    'valor del campo identificador de registro seleccionado (0 si es un reg. nuevo)
    If vMode <> 1 Then
        If vListView.SelectedItem.Selected = True Then
            vFieldID = vListView.SelectedItem.Key
        Else
            vFieldID = 0
        End If
    Else
        vFieldID = 0
    End If

End Function


Function Validar(pMode As Integer) As Boolean

    If pMode <> 4 Then
        Validar = False
        If txtID.Text = "" Then
            Beep
            MsgBox "Falta informaci�n." & Chr(13) & _
                             "Ingrese la Identificaci�n de la " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtID.SetFocus
            Exit Function
        
        ElseIf cboTipo.ListIndex = -1 Then
            Beep
            MsgBox "Falta informaci�n." & Chr(13) & _
                             "Ingrese el Tipo de Tarjeta antes de aceptar.", vbCritical + vbOKOnly, App.Title
            TxtDescri.SetFocus
            Exit Function
            
        ElseIf TxtDescri.Text = "" Then
            Beep
            MsgBox "Falta informaci�n." & Chr(13) & _
                             "Ingrese la descripci�n de la " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            TxtDescri.SetFocus
            Exit Function
            
'        ElseIf txtTelefono.Text = "" Then
'            Beep
'            MsgBox "Falta informaci�n." & Chr(13) & _
'                             "Ingrese el tel�fono que identifica a la " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
'            txtDescri.SetFocus
'            Exit Function
        End If
    End If
    
    Validar = True
    
End Function

Private Sub cmdAceptar_Click()

    Dim cSQL As String
    
    If Validar(vMode) = True Then
    
        Screen.MousePointer = vbHourglass
    
        DBConn.BeginTrans
        Select Case vMode
            Case 1 'nuevo
            
                cSQL = "INSERT INTO " & cTabla
                cSQL = cSQL & "  (TAR_CODIGO, TTA_CODIGO, TAR_DESCRI, TAR_TELEFONO, TAR_NROCTA, TAR_USUARIO, TAR_PASWOR) "
                cSQL = cSQL & "VALUES  ( "
                cSQL = cSQL & XN(txtID.Text) & ", "
                cSQL = cSQL & cboTipo.ItemData(cboTipo.ListIndex) & ", "
                cSQL = cSQL & XS(TxtDescri.Text) & ", "
                cSQL = cSQL & XS(txtTelefono.Text) & ", "
                cSQL = cSQL & XS(txtNroCta.Text) & ", "
                cSQL = cSQL & XS(txtUsuario.Text) & ", "
                cSQL = cSQL & XS(txtContrasena.Text) & ") "

            Case 2 'editar
                
                cSQL = "UPDATE " & cTabla & " SET "
                cSQL = cSQL & "  TAR_DESCRI = " & XS(TxtDescri.Text)
                cSQL = cSQL & " ,TAR_TELEFONO = " & XS(txtTelefono.Text)
                cSQL = cSQL & " ,TAR_NROCTA = " & XS(txtNroCta.Text)
                cSQL = cSQL & " ,TAR_USUARIO = " & XS(txtUsuario.Text)
                cSQL = cSQL & " ,TAR_PASWOR = " & XS(txtContrasena.Text)
                cSQL = cSQL & " WHERE TAR_CODIGO  = " & XN(txtID.Text)
            
            Case 4 'eliminar
            
                cSQL = "DELETE FROM " & cTabla & " WHERE TAR_CODIGO  = " & XN(txtID.Text)
            
        End Select
        
        DBConn.Execute cSQL
        DBConn.CommitTrans
        On Error GoTo 0
        
        If mOrigen = True Then
            'actualizo la lista base
            ActualizarListaBase vMode
        End If
        Screen.MousePointer = vbDefault
        Unload Me
    End If
    Exit Sub
    
ErrorTran:
    
    DBConn.RollbackTrans
    Screen.MousePointer = vbDefault
    
    'manejo el error
    ManejoDeErrores DBConn.ErrorNative
    
End Sub


Private Sub cmdAyuda_Click()
    Call WinHelp(Me.hWnd, App.Path & "\help\AYUDA.HLP", cdlHelpContext, 12)
End Sub

Private Sub cmdCerrar_Click()

    Unload Me
    
End Sub

Private Sub Form_Activate()
    'hizo click en una columna no correcta
    If vMode = 2 And vFieldID = "0" Then
        Unload Me
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
    
    
    If KeyAscii = 27 Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    Dim cSQL As String
    Dim hSQL As String
    Dim rec As ADODB.Recordset
    Set rec = New ADODB.Recordset
    
    'Me.Top = vFormLlama.Top + 1500
    'Me.Left = vFormLlama.Left + 1000
    
    cboTipo.Clear
    cSQL = "SELECT * FROM TIPO_TARJETA order by TTA_DESCRI"
    rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If (rec.BOF And rec.EOF) = 0 Then
       Do While rec.EOF = False
          cboTipo.AddItem Trim(rec!TTA_DESCRI)
          cboTipo.ItemData(cboTipo.NewIndex) = rec!TTA_CODIGO
          rec.MoveNext
       Loop
       cboTipo.ListIndex = 0
    End If
    rec.Close
    
    If vMode <> 1 Then
        If vFieldID <> "0" Then
            cSQL = "SELECT * FROM " & cTabla & "  WHERE TAR_CODIGO = " & Mid(vFieldID, 2, Len(vFieldID) - 2)
            rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
            If (rec.BOF And rec.EOF) = 0 Then
                'si encontr� el registro muestro los datos
                txtID.Text = rec!TAR_CODIGO
                BuscaCodigoProxItemData rec!TTA_CODIGO, cboTipo
                
                TxtDescri.Text = Trim(ChkNull(rec!TAR_DESCRI))
                txtTelefono.Text = Trim(ChkNull(rec!TAR_TELEFONO))
                txtNroCta.Text = Trim(ChkNull(rec!TAR_NROCTA))
                txtUsuario.Text = Trim(ChkNull(rec!TAR_USUARIO))
                txtContrasena.Text = Trim(ChkNull(rec!TAR_PASWOR))
            Else
                Beep
                MsgBox "Imposible encontrar el registro seleccionado.", vbCritical + vbOKOnly, App.Title
            End If
        End If
    End If
    
    'establesco funcionalidad del form de datos
    SetMode vMode
End Sub

Private Sub txtContrasena_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtContrasena_GotFocus()
    seltxt
End Sub

Private Sub txtContrasena_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtdescri_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtdescri_GotFocus()
    SelecTexto TxtDescri
End Sub

Private Sub txtdescri_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtID_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtID_GotFocus()
    seltxt
End Sub

Private Sub txtID_LostFocus()

    Dim cSQL As String
    Dim rec As ADODB.Recordset
    Set rec = New ADODB.Recordset
    
    If vMode = 1 Then ' si se esta usando en modo de nuevo registro
        If txtID.Text = "" Then
            If cSugerirID = True Then
                cSQL = "SELECT MAX(" & cCampoID & ") FROM " & cTabla
                rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
                If (rec.BOF And rec.EOF) = 0 Then
                    If rec.Fields(0) > 0 Then
                        txtID.Text = rec.Fields(0) + 1
                    Else
                        txtID.Text = 1
                    End If
                End If
            End If
        Else
            'verifico que no sea clave repetida
            cSQL = "SELECT COUNT(*) FROM " & cTabla & " WHERE " & cCampoID & " = " & txtID.Text
            rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
            If (rec.BOF And rec.EOF) = 0 Then
                If rec.Fields(0) > 0 Then
                    Beep
                    MsgBox "C�digo de " & cDesRegistro & " repetido." & Chr(13) & _
                                     "El c�digo ingresado Pertenece a otro registro de " & cDesRegistro & ".", vbCritical + vbOKOnly, App.Title
                    txtID.Text = ""
                    txtID.SetFocus
                End If
            End If
        End If
    End If
End Sub

Private Sub txtNroCta_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtNroCta_GotFocus()
    seltxt
End Sub

Private Sub txtNroCta_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtTelefono_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtTelefono_GotFocus()
    seltxt
End Sub

Private Sub txtTelefono_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtUsuario_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtUsuario_GotFocus()
    seltxt
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub