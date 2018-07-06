VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ABMClientes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos del Cliente..."
   ClientHeight    =   5910
   ClientLeft      =   2700
   ClientTop       =   2625
   ClientWidth     =   4845
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ABMClientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkctacte 
      Caption         =   "Bloquear Cta Cte"
      Height          =   375
      Left            =   2760
      TabIndex        =   37
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox txtlimitectacte 
      Height          =   315
      Left            =   1185
      MaxLength       =   10
      TabIndex        =   7
      Top             =   2040
      Width           =   1400
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   330
      Left            =   900
      TabIndex        =   35
      Top             =   5460
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox txtNroDoc 
      Height          =   315
      Left            =   3675
      MaxLength       =   9
      TabIndex        =   6
      Top             =   1695
      Width           =   1005
   End
   Begin VB.TextBox txtCodPostal 
      Height          =   315
      Left            =   1185
      MaxLength       =   10
      TabIndex        =   12
      Top             =   3450
      Width           =   1230
   End
   Begin VB.TextBox txtObserva 
      Height          =   570
      Left            =   1185
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   4800
      Width           =   3495
   End
   Begin VB.TextBox txtIngresosBrutos 
      Height          =   315
      Left            =   3675
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1320
      Width           =   1005
   End
   Begin VB.ComboBox cboIva 
      Height          =   315
      ItemData        =   "ABMClientes.frx":000C
      Left            =   1185
      List            =   "ABMClientes.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   975
      Width           =   3495
   End
   Begin VB.TextBox txtDomicilio 
      Height          =   315
      Left            =   1185
      MaxLength       =   50
      TabIndex        =   11
      Top             =   3105
      Width           =   3495
   End
   Begin VB.TextBox txtMail 
      Height          =   315
      Left            =   1185
      MaxLength       =   50
      TabIndex        =   15
      Top             =   4455
      Width           =   3495
   End
   Begin VB.TextBox txtFax 
      Height          =   315
      Left            =   1185
      MaxLength       =   30
      TabIndex        =   14
      Top             =   4125
      Width           =   3495
   End
   Begin VB.TextBox txtTelefono 
      Height          =   315
      Left            =   1185
      MaxLength       =   30
      TabIndex        =   13
      Top             =   3795
      Width           =   3495
   End
   Begin VB.ComboBox cboLocalidad 
      Height          =   315
      ItemData        =   "ABMClientes.frx":0010
      Left            =   1185
      List            =   "ABMClientes.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2760
      Width           =   3495
   End
   Begin VB.ComboBox cboProvincia 
      Height          =   315
      ItemData        =   "ABMClientes.frx":0014
      Left            =   1185
      List            =   "ABMClientes.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2415
      Width           =   3495
   End
   Begin VB.ComboBox cboPais 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "ABMClientes.frx":0018
      Left            =   3615
      List            =   "ABMClientes.frx":001A
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   -15
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton cmdAyuda 
      Height          =   315
      Left            =   240
      Picture         =   "ABMClientes.frx":001C
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5460
      Width           =   330
   End
   Begin VB.TextBox txtNombre 
      Height          =   315
      Left            =   1185
      MaxLength       =   50
      TabIndex        =   1
      Top             =   630
      Width           =   3495
   End
   Begin VB.TextBox txtID 
      Height          =   315
      Left            =   1185
      TabIndex        =   0
      Top             =   285
      Width           =   720
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   345
      Left            =   3420
      TabIndex        =   18
      Top             =   5460
      Width           =   1300
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   2070
      TabIndex        =   17
      Top             =   5460
      Width           =   1300
   End
   Begin MSMask.MaskEdBox txtCuit 
      Height          =   315
      Left            =   1185
      TabIndex        =   3
      Top             =   1320
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   13
      Mask            =   "##-########-#"
      PromptChar      =   "_"
   End
   Begin MSComCtl2.DTPicker Fecha 
      Height          =   315
      Left            =   1185
      TabIndex        =   5
      Top             =   1680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   52494337
      CurrentDate     =   41098
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cta. Cte:"
      Height          =   195
      Index           =   14
      Left            =   135
      TabIndex        =   36
      Top             =   2085
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Provincia:"
      Height          =   195
      Index           =   13
      Left            =   135
      TabIndex        =   34
      Top             =   2475
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nro. Doc.:"
      Height          =   195
      Index           =   3
      Left            =   2745
      TabIndex        =   33
      Top             =   1710
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código Postal:"
      Height          =   195
      Index           =   2
      Left            =   135
      TabIndex        =   32
      Top             =   3510
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Observación:"
      Height          =   195
      Index           =   12
      Left            =   135
      TabIndex        =   31
      Top             =   4800
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "F. Nacimiento:"
      Height          =   195
      Index           =   2
      Left            =   135
      TabIndex        =   30
      Top             =   1710
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ing. Brutos:"
      Height          =   195
      Index           =   11
      Left            =   2745
      TabIndex        =   29
      Top             =   1380
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "C.U.I.T.:"
      Height          =   195
      Index           =   10
      Left            =   135
      TabIndex        =   28
      Top             =   1380
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cond. I.V.A.:"
      Height          =   195
      Index           =   9
      Left            =   135
      TabIndex        =   27
      Top             =   1020
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Domicilio:"
      Height          =   195
      Index           =   8
      Left            =   135
      TabIndex        =   26
      Top             =   3150
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "e-mail:"
      Height          =   195
      Index           =   7
      Left            =   135
      TabIndex        =   25
      Top             =   4500
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fax:"
      Height          =   195
      Index           =   6
      Left            =   135
      TabIndex        =   24
      Top             =   4170
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Teléfono:"
      Height          =   195
      Index           =   5
      Left            =   135
      TabIndex        =   23
      Top             =   3840
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Localidad:"
      Height          =   195
      Index           =   4
      Left            =   135
      TabIndex        =   22
      Top             =   2805
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nombre:"
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   20
      Top             =   675
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Id.:"
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   19
      Top             =   315
      Width           =   270
   End
End
Attribute VB_Name = "ABMClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'parametros para la configuración de la ventana de datos
Public vFieldID As String
Dim vStringSQL As String
Dim vFormLlama As Form
Public vMode As Integer
Dim vListView As ListView
Dim vDesFieldID As String
Dim Pais As String
Dim Provincia As String


'constantes para funcionalidad de uso del formulario
Const cSugerirID = True 'si es True si sugiere un identificador cuando deja el campo en blanco
Const cTabla = "CLIENTE"
Const cCampoID = "CLI_CODIGO"
Const cDesRegistro = "Cliente"

Function ActualizarListaBase(pMode As Integer)
    On Error GoTo moco
    Dim rec As ADODB.Recordset
    Dim cSQL As String
    Dim i As Integer
    Dim auxListItem As ListItem
    Dim IndiceCampoID As Integer
    Dim OrdenCampo As Integer
    Dim f As ADODB.Field
    Set rec = New ADODB.Recordset
    
    'armo la cadena a ejecutar
    If InStr(1, vStringSQL, "WHERE") = 0 Then
        cSQL = vStringSQL & " WHERE " & cCampoID & " = " & txtId.Text
    Else
        cSQL = vStringSQL & " AND " & cCampoID & " = " & txtId.Text
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
        
            'recorro la coleción de campos a actualizar
            For i = 0 To rec.Fields.Count - 1
                If i = 0 Then
                    Select Case pMode
                        Case 1
                            Set auxListItem = vListView.ListItems.Add(, "'" & rec.Fields(IndiceCampoID) & "'", CStr(IIf(IsNull(rec.Fields(i)), "", rec.Fields(i))), 1)
                            auxListItem.Icon = 1
                            auxListItem.SmallIcon = 1
                            
                        Case 2
                            Set auxListItem = vListView.SelectedItem
                            auxListItem.Text = rec.Fields(i)
                    End Select
                Else
                    auxListItem.SubItems(i) = IIf(IsNull(rec.Fields(i)), "", rec.Fields(i))
                End If
            Next i
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
    'Parametro: pMode indica el modo en que se utilizará este form
    '  pMode  =             1> Indica nuevo registro
    '                       2> Editar registro existente
    '                       3> Mostrar dato del registro existente
    '                       4> Eliminar registro existente
    
    
    Select Case pMode
        Case 1, 2
            AcCtrlx txtNombre
            AcCtrlx cboIva
            AcCtrlx txtCuit
            AcCtrlx txtIngresosBrutos
            AcCtrlx txtNroDoc
            AcCtrlx Fecha
            'AcCtrlx cboPais
            AcCtrlx cboProvincia
            AcCtrlx cboLocalidad
            AcCtrlx txtDomicilio
            AcCtrlx txtTelefono
            AcCtrlx txtFax
            AcCtrlx txtCodPostal
            AcCtrlx txtMail
            AcCtrlx txtObserva
        Case 3, 4
            DesacCtrlx txtNombre
            DesacCtrlx cboIva
            DesacCtrlx txtCuit
            DesacCtrlx txtIngresosBrutos
            DesacCtrlx txtNroDoc
            DesacCtrlx Fecha
            'DesacCtrlx cboPais
            DesacCtrlx cboProvincia
            DesacCtrlx cboLocalidad
            DesacCtrlx txtDomicilio
            DesacCtrlx txtTelefono
            DesacCtrlx txtFax
            DesacCtrlx txtCodPostal
            DesacCtrlx txtMail
            DesacCtrlx txtObserva
    End Select
    
    Select Case pMode
        Case 1
            cmdAceptar.Enabled = False
            Me.Caption = "Nuevo " & cDesRegistro
            txtID_LostFocus
            DesacCtrl txtId
        Case 2
            cmdAceptar.Enabled = False
            Me.Caption = "Editando " & cDesRegistro
            DesacCtrl txtId
        Case 3
            cmdAceptar.Visible = False
            Me.Caption = "Datos del " & cDesRegistro
            DesacCtrl txtId
        Case 4
            cmdAceptar.Enabled = True
            Me.Caption = "Eliminando " & cDesRegistro
            DesacCtrl txtId
    End Select
    
End Function

Public Function SetWindow(pWindow As Form, pSQL As String, pMode As Integer, pListview As ListView, pDesID As String)
    
    Set vFormLlama = pWindow 'Objeto ventana que que llama a la ventana de datos
    vStringSQL = pSQL 'string utilizado para argar la lista base
    vMode = pMode  'modo en que se utilizará la ventana de datos
    Set vListView = pListview 'objeto listview que se está editando
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
        If txtId.Text = "" Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese la Identificación del  " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtId.SetFocus
            Exit Function
        ElseIf txtNombre.Text = "" Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese el Nombre del " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtNombre.SetFocus
            Exit Function
        
        ElseIf cboPais.ListIndex = -1 Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese el Paí del " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            cboPais.SetFocus
            Exit Function
            
        ElseIf cboProvincia.ListIndex = -1 Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese la Provincia del " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            cboPais.SetFocus
            Exit Function
        
        ElseIf cboLocalidad.ListIndex = -1 Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese la Localidad del " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            cboProvincia.SetFocus
            Exit Function
        End If
    End If
    
    Validar = True
    
End Function

Private Sub cboCanal_Click()
    cmdAceptar.Enabled = True
End Sub

Private Sub cboIva_Click()
    cmdAceptar.Enabled = True
End Sub

Private Sub cboLocalidad_Click()
    cmdAceptar.Enabled = True
End Sub

Private Sub cboPais_Click()
    cmdAceptar.Enabled = True
End Sub

Private Sub cboPais_LostFocus()
    If vMode = 2 And Pais = cboPais.Text Then
        Exit Sub
    End If
    Set Rec1 = New ADODB.Recordset
    cboProvincia.Clear
    sql = "SELECT PRO_CODIGO,PRO_DESCRI"
    sql = sql & " FROM PROVINCIA "
    sql = sql & " WHERE PAI_CODIGO=" & cboPais.ItemData(cboPais.ListIndex)
    'sql = sql & " AND PRO_CODIGO=1" 'CORDOBA
    sql = sql & " ORDER BY PRO_DESCRI"
    
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If (Rec1.BOF And Rec1.EOF) = 0 Then
       Do While Rec1.EOF = False
          cboProvincia.AddItem Trim(Rec1!PRO_DESCRI)
          cboProvincia.ItemData(cboProvincia.NewIndex) = Rec1!PRO_CODIGO
          Rec1.MoveNext
       Loop
       cboProvincia.ListIndex = cboProvincia.ListIndex + 1
       BuscaProx "CORDOBA", cboProvincia
    Else
       MsgBox "No hay cargado Provincia para ese País.", vbOKOnly + vbCritical, TIT_MSGBOX
    End If
    Rec1.Close
    cboProvincia_LostFocus
End Sub

Private Sub cboProvincia_Click()
    cmdAceptar.Enabled = True
End Sub

Private Sub cboProvincia_LostFocus()
    If vMode = 2 And Provincia = cboProvincia.Text Then
        Exit Sub
    End If
    Set Rec1 = New ADODB.Recordset
    cboLocalidad.Clear
    sql = "SELECT LOC_CODIGO,LOC_DESCRI FROM LOCALIDAD"
    sql = sql & " WHERE PAI_CODIGO=" & cboPais.ItemData(cboPais.ListIndex)
    sql = sql & " AND PRO_CODIGO=" & cboProvincia.ItemData(cboProvincia.ListIndex)
    sql = sql & " ORDER BY LOC_DESCRI "
    
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If (Rec1.BOF And Rec1.EOF) = 0 Then
       Do While Rec1.EOF = False
          cboLocalidad.AddItem Trim(Rec1!LOC_DESCRI)
          cboLocalidad.ItemData(cboLocalidad.NewIndex) = Rec1!LOC_CODIGO
          Rec1.MoveNext
       Loop
       cboLocalidad.ListIndex = cboLocalidad.ListIndex + 1
       BuscaProx "PILAR", cboLocalidad
    Else
       MsgBox "No hay cargada Localidad para esta Provincia.", vbOKOnly + vbCritical, TIT_MSGBOX
    End If
    Rec1.Close
    'BuscaProx "CORDOBA", cboLocalidad
End Sub

Private Sub chkctacte_Click()
    cmdAceptar.Enabled = True
End Sub

Private Sub cmdAceptar_Click()
' Resume Error
    Dim cSQL As String
    
    If Validar(vMode) = True Then
        
        On Error GoTo ErrorTran
        
        Screen.MousePointer = vbHourglass
    
        DBConn.BeginTrans
        If IsNull(Fecha.Value) Then
            Fecha.Value = Date
        End If
        Select Case vMode
            Case 1 'nuevo
                
                cSQL = "INSERT INTO " & cTabla
                cSQL = cSQL & "     (CLI_CODIGO, CLI_RAZSOC, CLI_DOMICI, CLI_CUIT,"
                cSQL = cSQL & " CLI_INGBRU, CLI_CUMPLE, IVA_CODIGO, CLI_NRODOC,"
                cSQL = cSQL & " CLI_TELEFONO, CLI_MAIL, CLI_FAX, CLI_CODPOS,"
                cSQL = cSQL & " LOC_CODIGO, PRO_CODIGO, PAI_CODIGO, CLI_OBSERVA, CLI_CTACTE, CLI_BLOCKEADO) "
                cSQL = cSQL & " VALUES "
                cSQL = cSQL & "     (" & XN(txtId.Text) & ", " & XS(txtNombre.Text) & ", "
                cSQL = cSQL & XS(txtDomicilio.Text) & ", " & XS(txtCuit.Text) & ", "
                cSQL = cSQL & XS(txtIngresosBrutos.Text) & ", "
                cSQL = cSQL & XD(Fecha.Value) & ", "
                cSQL = cSQL & cboIva.ItemData(cboIva.ListIndex) & ", "
                cSQL = cSQL & XN(txtNroDoc.Text) & ", "
                cSQL = cSQL & XS(txtTelefono.Text) & ", "
                cSQL = cSQL & XS(txtMail.Text) & ", " & XS(txtFax.Text) & ", "
                cSQL = cSQL & XS(txtCodPostal.Text) & ", "
                cSQL = cSQL & cboLocalidad.ItemData(cboLocalidad.ListIndex) & ", "
                cSQL = cSQL & cboProvincia.ItemData(cboProvincia.ListIndex) & ", "
                cSQL = cSQL & cboPais.ItemData(cboPais.ListIndex) & ","
                cSQL = cSQL & XS(Trim(txtObserva.Text)) & ","
                cSQL = cSQL & XN(Trim(txtlimitectacte.Text)) & ","
                'CLI_CTACTE BLOQUEADO
                If chkctacte.Value = Checked Then
                    cSQL = cSQL & 1 & ")"
                Else
                    cSQL = cSQL & 0 & ")"
                End If
                
                
            Case 2 'editar
                
                cSQL = "UPDATE " & cTabla & " SET "
                cSQL = cSQL & "  CLI_RAZSOC=" & XS(txtNombre.Text)
                cSQL = cSQL & " ,CLI_DOMICI=" & XS(txtDomicilio.Text)
                cSQL = cSQL & " ,CLI_CUIT=" & XS(txtCuit.Text)
                cSQL = cSQL & " ,CLI_INGBRU=" & XS(txtIngresosBrutos.Text)
                cSQL = cSQL & " ,CLI_CUMPLE=" & XD(Fecha.Value)
                cSQL = cSQL & " ,IVA_CODIGO=" & cboIva.ItemData(cboIva.ListIndex)
                cSQL = cSQL & " ,CLI_TELEFONO=" & XS(txtTelefono.Text)
                cSQL = cSQL & " ,CLI_MAIL=" & XS(txtMail.Text)
                cSQL = cSQL & " ,CLI_FAX=" & XS(txtFax.Text)
                cSQL = cSQL & " ,CLI_CODPOS=" & XS(txtCodPostal.Text)
                cSQL = cSQL & " ,LOC_CODIGO=" & cboLocalidad.ItemData(cboLocalidad.ListIndex)
                cSQL = cSQL & " ,PRO_CODIGO=" & cboProvincia.ItemData(cboProvincia.ListIndex)
                cSQL = cSQL & " ,PAI_CODIGO=" & cboPais.ItemData(cboPais.ListIndex)
                cSQL = cSQL & " ,CLI_OBSERVA=" & XS(Trim(txtObserva.Text))
                cSQL = cSQL & " ,CLI_NRODOC=" & XN(txtNroDoc.Text)
                cSQL = cSQL & " ,CLI_CTACTE=" & XN(txtlimitectacte.Text)
                If chkctacte.Value = Checked Then
                    cSQL = cSQL & " ,CLI_BLOCKEADO=1"
                Else
                    cSQL = cSQL & " ,CLI_BLOCKEADO=0"
                End If
                cSQL = cSQL & " WHERE CLI_CODIGO  = " & XN(txtId.Text)
            
            Case 4 'eliminar
                cSQL = "DELETE FROM " & cTabla & " WHERE CLI_CODIGO  = " & XN(txtId.Text)
                
        End Select
        
        DBConn.Execute cSQL
        DBConn.CommitTrans
        'On Error GoTo 0
        
        'actualizo la lista base
        ActualizarListaBase vMode
        
        Screen.MousePointer = vbDefault
        Unload Me
    End If
    Exit Sub
    
ErrorTran:
    
    DBConn.RollbackTrans
    Screen.MousePointer = vbDefault
 '   Resume Error
    'manejo el error
    'ManejoDeErrores DBConn.ErrorNative
    MsgBox Err.Description, vbCritical
    
End Sub


Private Sub cmdAyuda_Click()
    Call WinHelp(Me.hWnd, App.Path & "\help\AYUDA.HLP", cdlHelpContext, 12)
End Sub

Private Sub cmdCerrar_Click()

    Unload Me
    
End Sub

Private Sub Command1_Click()
    Dim X As Integer
    X = 2
    sql = "SELECT * FROM XX"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            sql = "INSERT INTO CLIENTE (CLI_CODIGO,CLI_RAZSOC,"
            sql = sql & " CLI_DOMICI,CLI_TELEFONO,CLI_FAX,CLI_MAIL,CLI_CUMPLE,"
            sql = sql & " IVA_CODIGO,PAI_CODIGO,PRO_CODIGO,LOC_CODIGO,CLI_NRODOC) VALUES ("
            sql = sql & X & ","
            sql = sql & "'" & Trim(rec!apellido) & " " & Trim(rec!Nombre) & "',"
            sql = sql & XS(rec!DIRECCION) & ","
            sql = sql & XS(rec!te) & ","
            sql = sql & XS(rec!cel) & ","
            sql = sql & XS(rec!mail) & ","
            sql = sql & XDQ(ChkNull(rec!nacimiento)) & ",2,1,1,"
            sql = sql & buscaloc(Trim(rec!CIUDAD)) & ","
            sql = sql & XN(rec!DNI) & ")"
            DBConn.Execute sql
            X = X + 1
            rec.MoveNext
        Loop
    End If
End Sub

Private Function buscaloc(mlocdescri As String) As Integer
    Select Case mlocdescri
        Case "PILAR"
            buscaloc = 1
        Case "RIO SEGUNDO", "RIO II", "RIO 2", "RIO II  CBA"
            buscaloc = 2
        Case "COSTA SACATE"
            buscaloc = 6
        Case "LAGUNA LARGA"
            buscaloc = 5
        Case "LAGUNILLA"
            buscaloc = 10
        Case "ONCATIVO"
            buscaloc = 20
        Case "VILLA DEL ROSARIO"
            buscaloc = 7
        Case "TOLEDO"
            buscaloc = 3
        Case "LOZADA", "LOSADA"
            buscaloc = 4
        Case "MATORRALES"
            buscaloc = 17
        Case "DESPEÑADEROS"
            buscaloc = 9
        Case "IMPIRA"
            buscaloc = 21
        Case "CAPILLA DE LOS REMEDIOS"
            buscaloc = 25
        Case "CARLOS PAZ"
            buscaloc = 11
        Case "MINA CLAVERO"
            buscaloc = 12
        Case "CORDOBA"
            buscaloc = 13
        Case "VILLA DEL TOTORAL"
            buscaloc = 14
        Case "COSME SUD"
            buscaloc = 15
        Case "JAMES CRAIK"
            buscaloc = 19
        Case "PIQUILLIN"
            buscaloc = 18
        Case "LAS JUNTURAS"
            buscaloc = 22
        Case "CALCHIN OESTE"
            buscaloc = 24
        Case "RINCON"
            buscaloc = 8
        Case Else
            buscaloc = 1
    End Select
End Function

Private Sub Form_Activate()
    'hizo click en una columna no correcta
    If vMode = 2 And vFieldID = "0" Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Initialize()
    'MsgBox "Initialize"
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
    'CARGO COMBO CONDICIN IVA
    Call CargoComboBox(cboIva, "CONDICION_IVA", "IVA_CODIGO", "IVA_DESCRI")
    If cboIva.ListCount > 0 Then
        cboIva.ListIndex = 0
    End If
    
    'cargo el combo de PAIS
    cboPais.Clear
    cSQL = "SELECT * FROM PAIS WHERE PAI_CODIGO=1 ORDER BY PAI_DESCRI"
    rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If (rec.BOF And rec.EOF) = 0 Then
       Do While rec.EOF = False
          cboPais.AddItem Trim(rec!PAI_DESCRI)
          cboPais.ItemData(cboPais.NewIndex) = rec!PAI_CODIGO
          rec.MoveNext
       Loop
       cboPais.ListIndex = cboPais.ListIndex + 1
    End If
    rec.Close
    cboPais_LostFocus
    
    Pais = ""
    Provincia = ""
    If vMode <> 1 Then
        If vFieldID <> "0" Then
            cSQL = "SELECT * FROM " & cTabla & "  WHERE CLI_CODIGO = " & Mid(vFieldID, 2, Len(vFieldID) - 2)
            rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
            If (rec.BOF And rec.EOF) = 0 Then
                'si encontró el registro muestro los datos
                txtId.Text = rec!CLI_CODIGO
                txtNombre.Text = rec!CLI_RAZSOC
                
                Call BuscaCodigoProxItemData(rec!IVA_CODIGO, cboIva)
                txtCuit.Text = ChkNull(rec!CLI_CUIT)
                txtIngresosBrutos.Text = ChkNull(rec!CLI_INGBRU)
                Fecha.Value = ChkNull(rec!CLI_CUMPLE)
                
                Call BuscaCodigoProxItemData(CInt(rec!PAI_CODIGO), cboPais)
                cboPais_LostFocus
                Pais = cboPais.Text
                
                Call BuscaCodigoProxItemData(CInt(rec!PRO_CODIGO), cboProvincia)
                cboProvincia_LostFocus
                Provincia = cboProvincia.Text
                
                txtNroDoc.Text = ChkNull(rec!CLI_NRODOC)
                Call BuscaCodigoProxItemData(CInt(rec!LOC_CODIGO), cboLocalidad)
                txtDomicilio.Text = ChkNull(rec!CLI_DOMICI)
                txtTelefono.Text = ChkNull(rec!CLI_TELEFONO)
                txtFax.Text = ChkNull(rec!CLI_FAX)
                txtCodPostal.Text = ChkNull(rec!CLI_CODPOS)
                txtMail.Text = ChkNull(rec!CLI_MAIL)
                txtObserva.Text = Trim(ChkNull(rec!CLI_OBSERVA))
                txtlimitectacte.Text = Chk0(rec!CLI_CTACTE)
                
                If txtlimitectacte.Text <> "" Or txtlimitectacte <> 0 Then
                    txtlimitectacte.Text = Valido_Importe2(txtlimitectacte)
                End If
                If Chk0(rec!CLI_BLOCKEADO) = 1 Then
                    chkctacte.Value = Checked
                Else
                    chkctacte.Value = Unchecked
                End If
            Else
                Beep
                MsgBox "Imposible encontrar el registro seleccionado.", vbCritical + vbOKOnly, App.Title
            End If
        End If
    End If
    'Me.Top = 0
    If usuario = "A" Then
        txtlimitectacte.Enabled = True
    Else
        txtlimitectacte.Enabled = False
    End If
    
    Centrar_pantalla Me
    
    'establesco funcionalidad del form de datos
    SetMode vMode
    
'' actualizar FCL_TOTALACT
'    sql = "SELECT * FROM FACTURA_CLIENTE WHERE FCL_FECHA > " & XDQ("01/01/2017")
'    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    If rec.EOF = False Then
'        Do While rec.EOF = False
'           sql = "UPDATE FACTURA_CLIENTE SET "
'           sql = sql & " FCL_TOTALACT=" & XN(Chk0(rec!FCL_TOTAL))
'           sql = sql & " WHERE TCO_CODIGO = " & rec!TCO_CODIGO
'           sql = sql & " AND FCL_SUCURSAL = " & rec!FCL_SUCURSAL
'           sql = sql & " AND FCL_NUMERO = " & rec!FCL_NUMERO
'           DBConn.Execute sql
'           rec.MoveNext
'        Loop
'    End If
'    rec.Close
    
    
'    'function para buscar cuits repetidos
'    sql = "SELECT CLI_CUIT, COUNT (*) AS CANT "
'    sql = sql & " FROM CLIENTE "
'    sql = sql & " GROUP BY CLI_CUIT"
'    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    DBConn.Execute "DELETE FROM TMP_CUITS"
'    If Rec1.EOF = False Then
'        Do While Rec1.EOF = False
'            If Rec1!CANT > 1 Then
'                'MsgBox "Cliente: " & Rec1!CLI_CUIT & " CANTIDAD: " & Rec1!CANT, vbInformation, TIT_MSGBOX
'                sql = "INSERT INTO TMP_CUITS (CLI_CUIT,CLI_CANT)" ',CLI_RAZSOC)"
'                sql = sql & " VALUES("
'                sql = sql & Chk0(Rec1!CLI_CUIT) & ","
'                sql = sql & XN(Rec1!CANT) & ")"
'                'sql = sql & XS(Rec1!CLI_RAZSOC) & ")"
'                DBConn.Execute sql
'            End If
'            Rec1.MoveNext
'        Loop
'    End If
'    Rec1.Close

''ACTUALIZAR TMP_CUIS
'    Dim Codigo As Integer
'    Dim Nombre As String
'    sql = "SELECT CLI_CUIT FROM TMP_CUITS"
'    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    If Rec1.EOF = False Then
'        Do While Rec1.EOF = False
'            If Rec1!CLI_CUIT <> "0" Then
'                Codigo = BuscoCliente(Rec1!CLI_CUIT)
'                Nombre = Trim(buscorazsoc(Rec1!CLI_CUIT))
'                sql = "UPDATE TMP_CUITS SET"
'                sql = sql & " CLI_CODIGO = " & Codigo
'                sql = sql & " ,CLI_RAZSOC = " & XS(Nombre)
'                sql = sql & " WHERE CLI_CUIT = " & XS(Rec1!CLI_CUIT)
'                DBConn.Execute sql
'            End If
'            Rec1.MoveNext
'        Loop
'    End If
'    Rec1.Close

'ACTUALIZO FACTURA_CLIENTE, RECIBO_CLIENTE Y BORRO CLIENTES
'PRIMERO BUSCO LOS CODIGOS DE CLIENTES POR CUIT Y QUE NO ESTAN EN TMP_CUITS
'    sql = "SELECT CLI_CUIT,CLI_CODIGO FROM TMP_CUITS"
'    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    If Rec1.EOF = False Then
'        Do While Rec1.EOF = False
'            If Rec1!CLI_CUIT <> "0" Then
'                actualizacion Rec1!CLI_CUIT, Rec1!CLI_CODIGO
'
'            End If
'            Rec1.MoveNext
'        Loop
'    End If
'    Rec1.Close

    
End Sub
'Private Function actualizacion(cuit As String, Cli As Integer)
'
'    sql = "SELECT CLI_CODIGO FROM CLIENTE WHERE CLI_CUIT=" & XS(cuit)
'    sql = sql & " AND CLI_CODIGO <> " & Cli
'    sql = sql & " ORDER BY CLI_CODIGO"
'
'    Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    If Rec2.EOF = False Then
'        Do While Rec2.EOF = False
'            'actualizo facturas del cliente
'            sql = "UPDATE FACTURA_CLIENTE SET"
'            sql = sql & " CLI_CODIGO=" & Cli
'            sql = sql & " WHERE CLI_CODIGO = " & Rec2!CLI_CODIGO
'            DBConn.Execute sql
'
'            'actualizo recibos del cliente
'            sql = "UPDATE RECIBO_CLIENTE SET"
'            sql = sql & " CLI_CODIGO=" & Cli
'            sql = sql & " WHERE CLI_CODIGO = " & Rec2!CLI_CODIGO
'            DBConn.Execute sql
'
'            'elimino clientes
'            sql = "DELETE FROM CLIENTE "
'            sql = sql & " WHERE CLI_CODIGO = " & Rec2!CLI_CODIGO
'            DBConn.Execute sql
'            Rec2.MoveNext
'        Loop
'    End If
'    Rec2.Close
'
'
'End Function
'
'
'Private Function BuscoCliente(cuit As String) As Integer
''    BuscoCliente = 0
''    sql = "SELECT TOP 1 CLI_CODIGO FROM CLIENTE WHERE CLI_CUIT=" & XS(cuit) & " ORDER BY CLI_CODIGO"
''    Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
''    If Rec2.EOF = False Then
''        BuscoCliente = Rec2!CLI_CODIGO
''    End If
''    Rec2.Close
'
'
'End Function
'Private Function buscorazsoc(cuit As String) As String
''    buscorazsoc = ""
''    sql = "SELECT TOP 1 CLI_RAZSOC FROM CLIENTE WHERE CLI_CUIT=" & XS(cuit) & " ORDER BY CLI_CODIGO"
''    Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
''    If Rec2.EOF = False Then
''        buscorazsoc = Rec2!CLI_RAZSOC
''    End If
''    Rec2.Close
'
'
'End Function


Private Sub txtCodPostal_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtCodPostal_GotFocus()
    SelecTexto txtCodPostal
End Sub

Private Sub txtCodPostal_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtCuit_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtCuit_GotFocus()
    SelecTexto txtCuit
End Sub

Private Sub txtCuit_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCuit_KeyUp(KeyCode As Integer, Shift As Integer)
    If Len(Trim(txtCuit.ClipText)) = 12 Then
      txtCuit.SelStart = 12
  End If
End Sub

Private Sub txtCuit_LostFocus()
    If txtCuit.Text <> "" Then
        'rutina de validación de CUIT
        If Not ValidoCuit(txtCuit) Then
            txtCuit.SetFocus
            Exit Sub
        End If
        If vMode = 1 Then
            buscocuit
        End If
        
    End If
End Sub
Private Function buscocuit()
    sql = "SELECT CLI_RAZSOC,CLI_CUIT FROM CLIENTE WHERE CLI_CUIT=" & XS(txtCuit.Text)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        MsgBox "El CUIT " & Format(txtCuit.Text, "##-########-#") & " ya ha sido ingresado en el Sistema para el cliente: " & rec!CLI_RAZSOC & "", vbExclamation, TIT_MSGBOX
        txtCuit.Text = ""
        txtCuit.SetFocus
    End If
    rec.Close
End Function
Private Sub txtDomicilio_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtDomicilio_GotFocus()
    SelecTexto txtDomicilio
End Sub

Private Sub txtDomicilio_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtFax_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtFax_GotFocus()
    SelecTexto txtFax
End Sub

Private Sub txtFax_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtIngresosBrutos_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtIngresosBrutos_GotFocus()
    SelecTexto txtIngresosBrutos
End Sub

Private Sub txtIngresosBrutos_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtlimitectacte_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtlimitectacte_GotFocus()
    SelecTexto txtlimitectacte
    
End Sub

Private Sub txtlimitectacte_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtlimitectacte, KeyAscii)
End Sub

Private Sub txtlimitectacte_LostFocus()
    If txtlimitectacte.Text <> "" Or txtlimitectacte <> "0" Then
        txtlimitectacte.Text = Valido_Importe2(txtlimitectacte)
    End If
End Sub

Private Sub txtMail_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtMail_GotFocus()
    SelecTexto txtMail
End Sub

Private Sub txtNombre_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtNombre_GotFocus()
    seltxt
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
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
        If txtId.Text = "" Then
            If cSugerirID = True Then
                cSQL = "SELECT MAX(" & cCampoID & ") FROM " & cTabla
                'cSQL = cSQL & " WHERE PAI_CODIGO = " & cboPais.ItemData(cboPais.ListIndex)
                rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
                If (rec.BOF And rec.EOF) = 0 Then
                    If rec.Fields(0) > 0 Then
                        txtId.Text = rec.Fields(0) + 1
                    Else
                        txtId.Text = 1
                    End If
                End If
            End If
        Else
            'verifico que no sea clave repetida
            cSQL = "SELECT COUNT(*) FROM " & cTabla & " WHERE " & cCampoID & " = " & XN(txtId.Text)
            'cSQL = cSQL & " AND PAI_CODIGO = " & cboPais.ItemData(cboPais.ListIndex)
            rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
            If (rec.BOF And rec.EOF) = 0 Then
                If rec.Fields(0) > 0 Then
                    Beep
                    MsgBox "Código de " & cDesRegistro & " repetido." & Chr(13) & _
                                     "El código ingresado Pertenece a otro registro de " & cDesRegistro & ".", vbCritical + vbOKOnly, App.Title
                    txtId.Text = ""
                    txtId.SetFocus
                End If
            End If
        End If
    End If
End Sub

Private Sub txtNroDoc_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtNroDoc_GotFocus()
    SelecTexto txtNroDoc
End Sub

Private Sub txtNroDoc_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtObserva_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtObserva_GotFocus()
    SelecTexto txtObserva
End Sub

Private Sub txtTelefono_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtTelefono_GotFocus()
    SelecTexto txtTelefono
End Sub

Private Sub txtTelefono_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub
