VERSION 5.00
Begin VB.Form frmAjusteStock 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ajuste de Stock - Productos"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   4950
      TabIndex        =   10
      Top             =   2250
      Width           =   960
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Guardar"
      Height          =   450
      Left            =   3000
      TabIndex        =   8
      Top             =   2250
      Width           =   960
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2145
      Left            =   45
      TabIndex        =   11
      Top             =   30
      Width           =   5880
      Begin VB.Frame fraTanque 
         Height          =   1095
         Left            =   1320
         TabIndex        =   19
         Top             =   960
         Visible         =   0   'False
         Width           =   3375
         Begin VB.TextBox txtSFisSis2 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1800
            TabIndex        =   6
            Top             =   360
            Width           =   1410
         End
         Begin VB.TextBox txtSFReal2 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1800
            MaxLength       =   15
            TabIndex        =   7
            Top             =   690
            Width           =   1410
         End
         Begin VB.TextBox txtSFisSis1 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   4
            Top             =   360
            Width           =   1410
         End
         Begin VB.TextBox txtSFReal1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            MaxLength       =   15
            TabIndex        =   5
            Top             =   690
            Width           =   1410
         End
         Begin VB.Label lbltanque2 
            AutoSize        =   -1  'True
            Caption         =   "Tanque 2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1800
            TabIndex        =   21
            Top             =   120
            Width           =   780
         End
         Begin VB.Label lbltanque1 
            AutoSize        =   -1  'True
            Caption         =   "Tanque 1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   360
            TabIndex        =   20
            Top             =   120
            Width           =   780
         End
      End
      Begin VB.TextBox txtCodInt 
         Height          =   345
         Left            =   4785
         TabIndex        =   18
         Top             =   930
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.TextBox txtdescri 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1395
         TabIndex        =   1
         Top             =   555
         Width           =   4320
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   210
         TabIndex        =   0
         Top             =   555
         Width           =   1170
      End
      Begin VB.TextBox txtStockFisicoReal 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1395
         TabIndex        =   3
         Top             =   1575
         Width           =   1170
      End
      Begin VB.TextBox txtStockFisicoSis 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1395
         TabIndex        =   2
         Top             =   1245
         Width           =   1170
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   270
         TabIndex        =   17
         Top             =   330
         Width           =   555
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Producto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1455
         TabIndex        =   16
         Top             =   330
         Width           =   1725
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Stock Fisico"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   210
         TabIndex        =   15
         Top             =   975
         Width           =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Real:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   930
         TabIndex        =   14
         Top             =   1620
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sistema:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   630
         TabIndex        =   13
         Top             =   1305
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   450
      Left            =   3975
      TabIndex        =   9
      Top             =   2250
      Width           =   960
   End
   Begin VB.Label lblEstado 
      AutoSize        =   -1  'True
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   90
      TabIndex        =   12
      Top             =   2340
      Width           =   660
   End
End
Attribute VB_Name = "frmAjusteStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGrabar_Click()
    Dim tanque1 As Integer
    Dim tanque2 As Integer
    Select Case txtcodigo.Text
        Case 1
            tanque1 = 1
            tanque2 = 2
        Case 3
            tanque1 = 3
            tanque2 = 4
    End Select
    
    If txtcodigo.Text = "" Then
        MsgBox "Falta Ingresar el Producto", vbCritical, TIT_MSGBOX
        txtcodigo.SetFocus
        Exit Sub
    End If
    If txtcodigo.Text = 1 Or txtcodigo.Text = 3 Then 'BUSCO STOCKS DE TANQUES
        If txtSFReal1.Text = "" Or txtSFReal2.Text = "" Then
            MsgBox "El Stock Real no puede estar en blanco", vbCritical, TIT_MSGBOX
            txtSFReal1.SetFocus
        End If
    Else
        If txtStockFisicoReal.Text = "" Then
            MsgBox "El Stock Real no puede estar en blanco", vbCritical, TIT_MSGBOX
            txtStockFisicoReal.SetFocus
            Exit Sub
        End If
    End If
    If MsgBox("Confirma ajuste de Stock", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        If txtcodigo.Text = 1 Or txtcodigo.Text = 3 Then
            lblEstado.Caption = "Actualizando..."
            'tanque 1
            sql = "UPDATE PRODUCTO_DETALLE"
            sql = sql & " SET PDT_CANTIDAD=" & XN(txtSFReal1.Text)
            sql = sql & " WHERE PDT_CODIGO=" & tanque1
            sql = sql & " AND PTO_CODIGO=" & XN(txtcodigo.Text)
            DBConn.Execute sql
            
            'tanque 2
            sql = "UPDATE PRODUCTO_DETALLE"
            sql = sql & " SET PDT_CANTIDAD=" & XN(txtSFReal2.Text)
            sql = sql & " WHERE PDT_CODIGO=" & tanque2
            sql = sql & " AND PTO_CODIGO=" & XN(txtcodigo.Text)
            DBConn.Execute sql
            fraTanque.Visible = True
        Else
        
        
        
        
            lblEstado.Caption = "Actualizando..."
            sql = "UPDATE STOCK"
            sql = sql & " SET DST_STKFIS=" & XN(txtStockFisicoReal.Text)
            sql = sql & " WHERE STK_CODIGO=" & XN(Sucursal)
            sql = sql & " AND PTO_CODIGO=" & XN(txtCodInt.Text)
            DBConn.Execute sql
        End If
        lblEstado.Caption = ""
        CmdNuevo_Click
    End If
End Sub

Private Sub CmdNuevo_Click()
    txtcodigo.Text = ""
    txtStockFisicoReal.Text = ""
    txtStockFisicoSis.Text = ""
    txtcodigo.SetFocus
    fraTanque.Visible = False
    lbltanque1.Caption = "Tanque 1"
    lbltanque2.Caption = "Tanque 2"
End Sub

Private Sub CmdSalir_Click()
    Set frmAjusteStock = Nothing
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then MySendKeys Chr(9)
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    'Call Centrar_pantalla(Me)
    Me.Left = 0
    Me.Top = 0
    sql = "SELECT SUC_CODIGO, SUC_DESCRI "
    sql = sql & " FROM SUCURSAL R "
    sql = sql & " WHERE SUC_CODIGO = " & XN(Sucursal)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Frame1.Caption = "Ajuste de Stock Sucursal  - " & Trim(rec!SUC_DESCRI)
    End If
    rec.Close
    lblEstado.Caption = ""
End Sub

Private Sub TxtCodigo_Change()
    If txtcodigo.Text = "" Then
        txtcodigo.Text = ""
        txtdescri.Text = ""
        txtCodInt.Text = ""
        txtStockFisicoReal.Text = ""
        txtStockFisicoSis.Text = ""
        CmdGrabar.Enabled = False
    End If
End Sub

Private Sub TxtCodigo_GotFocus()
    SelecTexto txtcodigo
End Sub

Private Sub txtcodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarProducto "CODIGO"
        txtcodigo.SetFocus
    End If
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub TxtCodigo_LostFocus()
    If txtcodigo.Text <> "" Then
        Set rec = New ADODB.Recordset
        sql = " SELECT P.PTO_DESCRI, P.PTO_CODIGO, S.DST_STKFIS"
        sql = sql & " FROM PRODUCTO P, STOCK S"
        sql = sql & " WHERE"
        sql = sql & " P.PTO_CODIGO=S.PTO_CODIGO"
        sql = sql & " AND S.STK_CODIGO=" & XN(Sucursal)
        If IsNumeric(txtcodigo.Text) Then
            sql = sql & " AND P.PTO_CODIGO =" & XN(txtcodigo.Text) & " OR P.PTO_CODBARRAS=" & XS(txtcodigo.Text)
        Else
            sql = sql & " AND P.PTO_CODBARRAS=" & XS(txtcodigo.Text)
        End If
        sql = sql & " ORDER BY P.PTO_CODIGO"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtdescri.Text = Trim(rec!PTO_DESCRI)
            txtCodInt.Text = rec!PTO_CODIGO
            txtStockFisicoSis.Text = Chk0(rec!DST_STKFIS)
            CmdGrabar.Enabled = True
        Else
            MsgBox "El Código no existe", vbExclamation, TIT_MSGBOX
            txtcodigo.SetFocus
            CmdGrabar.Enabled = False
        End If
        rec.Close
        If txtcodigo.Text = 1 Or txtcodigo.Text = 3 Then 'BUSCO STOCKS DE TANQUES
            fraTanque.Visible = True
            If txtcodigo.Text = 3 Then 'GASOIL
                lbltanque1.Caption = "Tanque 3"
                lbltanque2.Caption = "Tanque 4"
            Else
                lbltanque1.Caption = "Tanque 1"
                lbltanque2.Caption = "Tanque 2"
            End If
                        
            Set rec = New ADODB.Recordset
            sql = " SELECT P.PTO_CODIGO,PD.PDT_CODIGO, PD.PDT_CANTIDAD"
            sql = sql & " FROM PRODUCTO P, PRODUCTO_DETALLE PD"
            sql = sql & " WHERE"
            sql = sql & " P.PTO_CODIGO=PD.PTO_CODIGO"
            sql = sql & " AND P.PTO_CODIGO =" & XN(txtcodigo.Text)
            sql = sql & " ORDER BY PD.PDT_CODIGO"
            
            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If rec.EOF = False Then
                Do While rec.EOF = False
                    If rec!PDT_CODIGO = 1 Or rec!PDT_CODIGO = 3 Then 'TANQUE 1 NAFTA O 3 DE GASOIL
                        txtSFisSis1.Text = Chk0(rec!PDT_CANTIDAD)
                        txtSFReal1.Text = Chk0(rec!PDT_CANTIDAD)
                    Else 'TANQUE 2 NAFTA O 4 DE GASOIL
                        txtSFisSis2.Text = Chk0(rec!PDT_CANTIDAD)
                        txtSFReal2.Text = Chk0(rec!PDT_CANTIDAD)
                    End If
                    rec.MoveNext
                
                    
                    
                Loop
                CmdGrabar.Enabled = True
                
            Else
                MsgBox "El Código no existe", vbExclamation, TIT_MSGBOX
                txtcodigo.SetFocus
                CmdGrabar.Enabled = False
        End If
            
        Else
            fraTanque.Visible = False
        End If
    End If
End Sub

Private Sub txtdescri_Change()
    If txtdescri.Text = "" Then
        txtcodigo.Text = ""
    End If
End Sub

Private Sub txtdescri_GotFocus()
    SelecTexto txtdescri
End Sub

Private Sub txtdescri_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarProducto "CODIGO"
        txtdescri.SetFocus
    End If
End Sub

Private Sub txtdescri_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtDescri_LostFocus()
   If txtcodigo.Text = "" And txtdescri.Text <> "" Then
        Set Rec1 = New ADODB.Recordset
        Screen.MousePointer = vbHourglass
        sql = "SELECT P.PTO_CODIGO,P.PTO_DESCRI,P.PTO_CODBARRAS,S.DST_STKFIS"
        sql = sql & " FROM PRODUCTO P, STOCK S"
        sql = sql & " WHERE P.PTO_DESCRI LIKE '" & txtdescri.Text & "%'"
        sql = sql & " AND P.PTO_CODIGO=S.PTO_CODIGO"
        sql = sql & " AND S.STK_CODIGO=" & XN(Sucursal)
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            If Rec1.RecordCount > 1 Then
                'grdGrilla.SetFocus
                BuscarProducto "CADENA", Trim(txtdescri.Text)
                txtdescri.SetFocus
            Else
                txtcodigo.Text = Trim(ChkNull(Rec1!PTO_CODBARRAS))
                txtdescri.Text = Trim(Rec1!PTO_DESCRI)
                txtCodInt.Text = Trim(Rec1!PTO_CODIGO)
                txtStockFisicoSis.Text = Chk0(Rec1!DST_STKFIS)
                CmdGrabar.Enabled = True
            End If
        Else
                MsgBox "No se ha encontrado el Producto", vbExclamation, TIT_MSGBOX
                txtdescri.Text = ""
        End If
        Rec1.Close
        Screen.MousePointer = vbNormal
    ElseIf txtcodigo.Text = "" And txtdescri.Text = "" Then
        CmdGrabar.Enabled = False
    End If
End Sub

Private Sub txtSFReal1_GotFocus()
    SelecTexto txtSFReal1
End Sub

Private Sub txtSFReal1_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtSFReal1, KeyAscii)
End Sub

Private Sub txtSFReal2_gotfocus()
    SelecTexto txtSFReal2
End Sub

Private Sub txtSFReal2_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtSFReal2, KeyAscii)
End Sub

Private Sub txtStockFisicoReal_GotFocus()
    SelecTexto txtStockFisicoReal
End Sub

Private Sub txtStockFisicoReal_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Public Sub BuscarProducto(mQuien As String, Optional mCadena As String)
    Dim cSQL As String
    Dim hSQL As String
    Dim B As CBusqueda
    Dim I, posicion As Integer
    Dim cadena As String
    
    Set B = New CBusqueda
    With B
        'Set .Conn = DBConn
        cSQL = "SELECT PTO_DESCRI, PTO_CODIGO"
        cSQL = cSQL & " FROM PRODUCTO"
        If mQuien = "CADENA" Then
            cSQL = cSQL & " WHERE"
            cSQL = cSQL & " PTO_DESCRI LIKE '" & Trim(mCadena) & "%'"
        End If
        
        hSQL = "Descripción, Código"
        .sql = cSQL
        .Headers = hSQL
        .Field = "PTO_DESCRI"
        campo1 = .Field
        .Field = "PTO_CODIGO"
        campo2 = .Field
        .OrderBy = "PTO_DESCRI"
        camponumerico = False
        .Titulo = "Busqueda de Productos :"
        .MaxRecords = 1
        .Show
        ' utilizar la coleccion de datos devueltos
        If .ResultFields.Count > 0 Then
                txtcodigo.Text = .ResultFields(2)
                TxtCodigo_LostFocus
        End If
    End With
    Set B = Nothing
End Sub

