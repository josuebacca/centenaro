VERSION 5.00
Begin VB.Form frmFiltro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filtro Búsqueda"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3885
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFiltro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   3885
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cbmCerrarFiltro 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   2700
      TabIndex        =   2
      Top             =   1080
      Width           =   1110
   End
   Begin VB.CommandButton cmdAceptarFiltro 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1080
      Width           =   1110
   End
   Begin VB.TextBox txtBusqueda 
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   600
      Width           =   3690
   End
   Begin VB.Label Label1 
      Caption         =   "Filtro de Búsqueda"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   105
      Width           =   3690
   End
End
Attribute VB_Name = "frmFiltro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbmCerrarFiltro_Click()
    Unload Me
End Sub

Private Sub cmdAceptarFiltro_Click()
    Dim auxListView As ListView
    Screen.MousePointer = vbHourglass
    With auxDllActiva
        Set auxListView = .FormBase.lstvLista
        If txtBusqueda.Text <> "" Then
            If .Caption = "Actualización de Localidades" Then
                .sql = "SELECT L.LOC_DESCRI, L.LOC_CODIGO, P.PRO_DESCRI, P.PRO_CODIGO, PA.PAI_DESCRI, P.PAI_CODIGO"
                .sql = .sql & " FROM LOCALIDAD L, PROVINCIA P, PAIS PA"
                .sql = .sql & " WHERE P.PAI_CODIGO=PA.PAI_CODIGO"
                .sql = .sql & " AND L.PAI_CODIGO=PA.PAI_CODIGO"
                .sql = .sql & " AND L.PRO_CODIGO=P.PRO_CODIGO"
                .sql = .sql & " AND L.LOC_DESCRI LIKE " & XS("%" & txtBusqueda.Text & "%")
            End If
            If .Caption = "Actualización de Clientes" Then
                .sql = "SELECT CLI_RAZSOC, CLI_CODIGO, CLI_DOMICI, CLI_TELEFONO, CLI_MAIL"
                .sql = .sql & " FROM CLIENTE"
                .sql = .sql & " WHERE CLI_RAZSOC LIKE " & XS(txtBusqueda.Text & "%")
            End If
            If .Caption = "Actualización de Proveedores" Then
                .sql = "SELECT P.PROV_RAZSOC, P.PROV_CODIGO, T.TPR_DESCRI, T.TPR_CODIGO"
                .sql = .sql & " FROM PROVEEDOR P, TIPO_PROVEEDOR T"
                .sql = .sql & " WHERE P.TPR_CODIGO=T.TPR_CODIGO"
                .sql = .sql & " AND P.PROV_RAZSOC LIKE " & XS("%" & txtBusqueda.Text & "%")
            End If
            If .Caption = "Actualización de Productos" Then
                .sql = "SELECT P.PTO_DESCRI, P.PTO_PRECTO,P.PTO_IVA, P.PTO_CODIGO, R.RUB_DESCRI, L.LNA_DESCRI, M.MAR_DESCRI" & _
                       " FROM PRODUCTO P, RUBROS R, LINEAS L, MARCAS M" & _
                       " WHERE R.LNA_CODIGO=L.LNA_CODIGO AND P.LNA_CODIGO=L.LNA_CODIGO" & _
                       " AND P.RUB_CODIGO=R.RUB_CODIGO" & _
                       " AND M.MAR_CODIGO=P.MAR_CODIGO"
                .sql = .sql & " AND P.PTO_DESCRI like " & XS(txtBusqueda.Text & "%")
            End If
            If .Caption = "Actualización de Líneas" Then
                .sql = "SELECT LNA_DESCRI, LNA_CODIGO FROM LINEAS"
                .sql = .sql & " WHERE LNA_DESCRI LIKE " & XS(txtBusqueda.Text & "%")
            End If
            If .Caption = "Actualización de Rubros" Then
                .sql = "SELECT R.RUB_DESCRI, R.RUB_CODIGO, L.LNA_DESCRI, L.LNA_CODIGO" & _
                       " FROM RUBROS R, LINEAS L" & _
                       " WHERE R.LNA_CODIGO=L.LNA_CODIGO" & _
                       " AND R.RUB_DESCRI LIKE " & XS(txtBusqueda.Text & "%")
            End If
        Else
            If .Caption = "Actualización de Localidades" Then
                .sql = "SELECT L.LOC_DESCRI, L.LOC_CODIGO, P.PRO_DESCRI, P.PRO_CODIGO, PA.PAI_DESCRI, P.PAI_CODIGO"
                .sql = .sql & " FROM LOCALIDAD L, PROVINCIA P, PAIS PA"
                .sql = .sql & " WHERE P.PAI_CODIGO=PA.PAI_CODIGO"
                .sql = .sql & " AND L.PAI_CODIGO=PA.PAI_CODIGO"
                .sql = .sql & " AND L.PRO_CODIGO=P.PRO_CODIGO"
            End If
            If .Caption = "Actualización de Clientes" Then
                .sql = "SELECT CLI_RAZSOC, CLI_CODIGO, CLI_DOMICI, CLI_TELEFONO, CLI_MAIL"
                .sql = .sql & " FROM CLIENTE"
            End If
            If .Caption = "Actualización de Proveedores" Then
                .sql = "SELECT P.PROV_RAZSOC, P.PROV_CODIGO, T.TPR_DESCRI, T.TPR_CODIGO"
                .sql = .sql & " FROM PROVEEDOR P, TIPO_PROVEEDOR T"
                .sql = .sql & " WHERE P.TPR_CODIGO=T.TPR_CODIGO"
            End If
            If .Caption = "Actualización de Productos" Then
                .sql = "SELECT P.PTO_DESCRI, P.PTO_PRECTO,P.PTO_IVA, P.PTO_CODIGO, R.RUB_DESCRI, L.LNA_DESCRI, M.MAR_DESCRI" & _
                       " FROM PRODUCTO P, RUBROS R, LINEAS L, MARCAS M" & _
                       " WHERE R.LNA_CODIGO=L.LNA_CODIGO AND P.LNA_CODIGO=L.LNA_CODIGO" & _
                       " AND P.RUB_CODIGO=R.RUB_CODIGO" & _
                       " AND M.MAR_CODIGO=P.MAR_CODIGO"
            End If
            If .Caption = "Actualización de Líneas" Then
                .sql = "SELECT LNA_DESCRI, LNA_CODIGO FROM LINEAS"
            End If
            If .Caption = "Actualización de Rubros" Then
                .sql = "SELECT R.RUB_DESCRI, R.RUB_CODIGO, L.LNA_DESCRI, L.LNA_CODIGO" & _
                       " FROM RUBROS R, LINEAS L" & _
                       " WHERE R.LNA_CODIGO=L.LNA_CODIGO"
            End If
        End If
        CargarListView .FormBase, auxListView, .sql, .FieldID, .HeaderSQL, .FormBase.ImgLstLista
        .FormBase.sBarEstado.Panels(1).Text = auxListView.ListItems.Count & " Registro(s)"
    End With
    Screen.MousePointer = vbDefault
    
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAceptarFiltro_Click
    End If
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub txtBusqueda_GotFocus()
    SelecTexto txtBusqueda
End Sub

Private Sub txtBusqueda_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub
