VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAnulaDocumentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anulación de ...."
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10005
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
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   10005
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   480
      Left            =   9030
      Picture         =   "frmAnulaDocumentos.frx":0000
      TabIndex        =   9
      Top             =   5520
      Width           =   870
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   480
      Left            =   7260
      Picture         =   "frmAnulaDocumentos.frx":030A
      TabIndex        =   7
      Top             =   5520
      Width           =   870
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   480
      Left            =   8145
      Picture         =   "frmAnulaDocumentos.frx":0614
      TabIndex        =   8
      Top             =   5520
      Width           =   870
   End
   Begin MSFlexGridLib.MSFlexGrid GrdModulos 
      Height          =   3750
      Left            =   45
      TabIndex        =   6
      Top             =   1455
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   6615
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      RowHeightMin    =   280
      BackColorSel    =   16761024
      AllowBigSelection=   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame frameBuscar 
      Caption         =   "xxx..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   75
      TabIndex        =   10
      Top             =   30
      Width           =   9825
      Begin VB.TextBox txtDesCli 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3075
         MaxLength       =   50
         TabIndex        =   1
         Tag             =   "Descripción"
         Top             =   270
         Width           =   4155
      End
      Begin VB.TextBox txtCliente 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2280
         MaxLength       =   40
         TabIndex        =   0
         Top             =   270
         Width           =   765
      End
      Begin VB.CommandButton CmdBuscAprox 
         Caption         =   "Buscar"
         Height          =   420
         Left            =   7680
         MaskColor       =   &H000000FF&
         TabIndex        =   5
         Top             =   840
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.ComboBox cboDocumento 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   615
         Width           =   3630
      End
      Begin MSComCtl2.DTPicker FechaDesde 
         Height          =   315
         Left            =   2280
         TabIndex        =   3
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   54657025
         CurrentDate     =   41098
      End
      Begin MSComCtl2.DTPicker FechaHasta 
         Height          =   315
         Left            =   5760
         TabIndex        =   4
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   54657025
         CurrentDate     =   41098
      End
      Begin VB.Label lblFechaHasta 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   4815
         TabIndex        =   14
         Top             =   1020
         Width           =   960
      End
      Begin VB.Label lblFechaDesde 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   1185
         TabIndex        =   13
         Top             =   1020
         Width           =   990
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   1185
         TabIndex        =   12
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   1185
         TabIndex        =   11
         Top             =   645
         Width           =   360
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Estado"
      Height          =   195
      Left            =   4395
      TabIndex        =   19
      Top             =   5250
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Anulado"
      Height          =   195
      Left            =   5055
      TabIndex        =   18
      Top             =   5850
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Pendiente"
      Height          =   195
      Left            =   5055
      TabIndex        =   17
      Top             =   5640
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Definitivo"
      Height          =   195
      Left            =   5055
      TabIndex        =   16
      Top             =   5445
      Width           =   675
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   150
      Left            =   4380
      Top             =   5880
      Width           =   540
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      Height          =   150
      Left            =   4380
      Top             =   5685
      Width           =   540
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   150
      Left            =   4380
      Top             =   5490
      Width           =   540
   End
   Begin VB.Label lblEstado 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   210
      TabIndex        =   15
      Top             =   5535
      Width           =   660
   End
End
Attribute VB_Name = "frmAnulaDocumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public TipodeAnulacion As Integer
Dim i As Integer

Private Sub CmdBuscAprox_Click()
    Select Case TipodeAnulacion
        Case 1 'Presupuestos
            'GrdModulos.Rows = 1
            'BuscoPedidos
        Case 2 'REMITOS
            'GrdModulos.Rows = 1
            'BuscoRemitos
        Case 3 'FACTURAS
            GrdModulos.Rows = 1
            BuscoFacturas
        Case 4 'RECIBOS
            GrdModulos.Rows = 1
            BuscoRecibos
        Case 5 'ORDEN COMPRA
            'GrdModulos.Rows = 1
            'BuscoOrdenesCompra
        Case 6 'REMITOS
            'GrdModulos.Rows = 1
            'BuscoRemitosProveedor
        Case 7 'FACTURAS
            'GrdModulos.Rows = 1
            'BuscoFacturasProveedor
        Case 8 'ORDEN DE PAGO
            'GrdModulos.Rows = 1
            'BuscoOrdenPago
        Case 9 'ORDEN DE PAGO
            'GrdModulos.Rows = 1
            'BuscoGastosGrales
        Case 10 'NOTA DE CREDITO
            'GrdModulos.Rows = 1
            'BuscoNC
            
    End Select
End Sub

Private Sub BuscoFacturas()
    lblEstado.Caption = "Buscando Facturas..."
    Screen.MousePointer = vbHourglass
    'poner sucursal
    sql = "SELECT DISTINCT FC.FCL_NUMERO,FC.FCL_SUCURSAL,FC.FCL_FECHA, FC.EST_CODIGO, E.EST_DESCRI,"
    sql = sql & " C.CLI_CODIGO, C.CLI_RAZSOC, TC.TCO_ABREVIA, FC.TCO_CODIGO"
    sql = sql & " FROM FACTURA_CLIENTE FC, CLIENTE C,"
    sql = sql & " TIPO_COMPROBANTE TC, ESTADO_DOCUMENTO E"
    sql = sql & " WHERE"
    sql = sql & " FC.TCO_CODIGO=TC.TCO_CODIGO"
    sql = sql & " AND FC.EST_CODIGO=E.EST_CODIGO"
    sql = sql & " AND FC.CLI_CODIGO=C.CLI_CODIGO"
    If txtCliente.Text <> "" Then
        sql = sql & " AND FC.CLI_CODIGO=" & XN(txtCliente.Text)
    End If
    If FechaDesde.Value <> "" Then
        sql = sql & " AND FC.FCL_FECHA>=" & XDQ(FechaDesde.Value)
    End If
    If FechaHasta.Value <> "" Then
        sql = sql & " AND FC.FCL_FECHA<=" & XDQ(FechaHasta.Value)
    End If
    If cboDocumento.List(cboDocumento.ListIndex) <> "(Todos)" Then
        sql = sql & " AND FC.TCO_CODIGO=" & XN(cboDocumento.ItemData(cboDocumento.ListIndex))
    End If
    sql = sql & " ORDER BY FC.FCL_FECHA,FC.FCL_NUMERO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            GrdModulos.AddItem rec!TCO_ABREVIA & Chr(9) & Format(rec!FCL_SUCURSAL, "0000") & "-" & Format(rec!FCL_NUMERO, "00000000") & Chr(9) & rec!FCL_FECHA _
                            & Chr(9) & rec!CLI_RAZSOC & Chr(9) & rec!EST_DESCRI & Chr(9) & rec!EST_CODIGO _
                            & Chr(9) & rec!EST_CODIGO & Chr(9) & rec!TCO_CODIGO & Chr(9) & rec!CLI_CODIGO
                                                        
            If rec!EST_CODIGO = 2 Then
                Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.Rows - 1, vbRed)
            ElseIf rec!EST_CODIGO = 1 Then
                Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.Rows - 1, vbBlue)
            End If
            rec.MoveNext
        Loop
        GrdModulos.SetFocus
    Else
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        MsgBox "No se encontraron Facturas...", vbExclamation, TIT_MSGBOX
    End If
    rec.Close
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdGrabar_Click()
    If MsgBox("¿Confirma Anular?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    
    On Error GoTo SeClavo
    lblEstado.Caption = "Actualizando..."
    Screen.MousePointer = vbHourglass
    DBConn.BeginTrans
        
    Select Case TipodeAnulacion
        Case 1 'Presupuestos
            'ActualizoPedido
        Case 2 'REMITOS
            'ActualizoRemito
        Case 3 'FACTURAS
            ActualizoFactura
        Case 4 'RECIBOS
            ActualizoRecibo
        Case 5 'ORDEN COMPRA
            'ActualizoOrdenCompra
        Case 6 'REMITO PROVEEDOR
            'ActualizoRemitoProveedor
        Case 7 'FACTURAS PROVEEDOR
            'ActualizoFacturaProveedor
        Case 8 'ORDEN PAGO
            'ActualizoOrdenPago
        Case 9 'ORDEN PAGO
            'ActualizoGastoGeneral
        Case 10 'NOTA DE CREDITO
            'ActualizoNotaCredito
    End Select
    
    DBConn.CommitTrans
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    CmdNuevo_Click
    Exit Sub

SeClavo:
    DBConn.RollbackTrans
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub ActualizoFactura()
    For i = 1 To GrdModulos.Rows - 1
        If GrdModulos.TextMatrix(i, 5) <> GrdModulos.TextMatrix(i, 6) Then 'PREGUNTA SI HUBO CAMBIO
            Set Rec2 = New ADODB.Recordset
            sql = "SELECT FCL_TCO_CODIGO FROM FACTURAS_NOTA_CREDITO_CLIENTE"
            sql = sql & " WHERE"
            sql = sql & " FCL_TCO_CODIGO=" & XN(GrdModulos.TextMatrix(i, 7))
            sql = sql & " AND FCL_NUMERO=" & XN(Right(GrdModulos.TextMatrix(i, 1), 8))
            sql = sql & " AND FCL_SUCURSAL=" & XN(Left(GrdModulos.TextMatrix(i, 1), 4))
            Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
            
            If Rec2.EOF = True Then
                sql = "UPDATE FACTURA_CLIENTE"
                sql = sql & " SET EST_CODIGO=" & XN(GrdModulos.TextMatrix(i, 6))
                sql = sql & " WHERE"
                sql = sql & " TCO_CODIGO=" & XN(GrdModulos.TextMatrix(i, 7))
                sql = sql & " AND FCL_NUMERO=" & XN(Right(GrdModulos.TextMatrix(i, 1), 8))
                sql = sql & " AND FCL_SUCURSAL=" & XN(Left(GrdModulos.TextMatrix(i, 1), 4))
                DBConn.Execute sql
                
                'VUELVO ATRAS EL STOCK
                sql = "SELECT PTO_CODIGO, DFC_CANTIDAD"
                sql = sql & " FROM DETALLE_FACTURA_CLIENTE"
                sql = sql & " WHERE"
                sql = sql & " TCO_CODIGO=" & XN(GrdModulos.TextMatrix(i, 7))
                sql = sql & " AND FCL_NUMERO=" & XN(Right(GrdModulos.TextMatrix(i, 1), 8))
                sql = sql & " AND FCL_SUCURSAL=" & XN(Left(GrdModulos.TextMatrix(i, 1), 4))
                rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
                If rec.EOF = False Then
                    Do While rec.EOF = False
                        sql = "UPDATE STOCK SET"
                        sql = sql & " DST_STKFIS = DST_STKFIS + " & XN(rec!DFC_CANTIDAD)
                        sql = sql & " WHERE STK_CODIGO = " & XN(Left(GrdModulos.TextMatrix(i, 1), 4))
                        sql = sql & " AND PTO_CODIGO = " & XN(rec!PTO_CODIGO)
                        DBConn.Execute sql
                        rec.MoveNext
                    Loop
                End If
                rec.Close
            Else
                MsgBox "La Factura número: " & GrdModulos.TextMatrix(i, 1) & ", no puede ser ANULADA" & Chr(13) & _
                                           " por estar relacionada con una Nota de Crédito", vbCritical, TIT_MSGBOX
                GrdModulos_dblClick
            End If
            If Rec2.State = 1 Then Rec2.Close
        End If
    Next
End Sub

Private Sub CambiColoryEstado(Estado As Boolean)
    cboDocumento.Enabled = Estado
    If Estado = False Then
        cboDocumento.BackColor = &H8000000F
    Else
        cboDocumento.BackColor = &H80000005
    End If
End Sub

Private Sub CmdSalir_Click()
    Set frmAnulaDocumentos = Nothing
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then MySendKeys Chr(9)
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
     Set rec = New ADODB.Recordset
     Set Rec2 = New ADODB.Recordset
     
    Me.Left = 0
    Me.Top = 0
    
    Select Case TipodeAnulacion
        Case 3 'FACTURAS
            frmAnulaDocumentos.Caption = "Anular Facturas"
            frameBuscar.Caption = "Buscar Facturas por..."
            'CARGO COMBO FACTURA
            LlenarComboFactura
            ConfiguroGrillaFactura
            Call CambiColoryEstado(True)
            
        Case 4 'RECIBOS
            frmAnulaDocumentos.Caption = "Anular Recibos"
            frameBuscar.Caption = "Buscar Recibos por..."
            'CARGO COMBO RECIBO
            LlenarComboRecibo
            ConfiguroGrillaRecibo
            Call CambiColoryEstado(True)
            
        Case 5 'NOTA DE CREDITO
            frmAnulaDocumentos.Caption = "Anular Nota de Crédito"
            frameBuscar.Caption = "Buscar Nota de Crédito por..."
            'CARGO COMBO NOTA DE CREDITO
            LlenarComboNotaCredito
            ConfiguroGrillaNotaDC
            Call CambiColoryEstado(True)
            
        Case 6 'NOTA DE DEBITO
            frmAnulaDocumentos.Caption = "Anular Nota de Débito"
            frameBuscar.Caption = "Buscar Nota de Débito por..."
            'CARGO COMBO NOTA DE DEBITO
            LlenarComboNotaDebito
            ConfiguroGrillaNotaDC
            Call CambiColoryEstado(True)
            
    End Select
    lblEstado.Caption = ""
End Sub

Private Sub LlenarComboNotaDebito()
    sql = "SELECT * FROM TIPO_COMPROBANTE"
    sql = sql & " WHERE TCO_DESCRI LIKE 'NOTA DE DEB%'"
    sql = sql & " ORDER BY TCO_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboDocumento.AddItem "(Todos)"
        Do While rec.EOF = False
            cboDocumento.AddItem rec!TCO_DESCRI
            cboDocumento.ItemData(cboDocumento.NewIndex) = rec!TCO_CODIGO
            rec.MoveNext
        Loop
        cboDocumento.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub LlenarComboNotaCredito()
    sql = "SELECT * FROM TIPO_COMPROBANTE"
    sql = sql & " WHERE TCO_DESCRI LIKE 'NOTA DE CRED%'"
    sql = sql & " ORDER BY TCO_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboDocumento.AddItem "(Todos)"
        Do While rec.EOF = False
            cboDocumento.AddItem rec!TCO_DESCRI
            cboDocumento.ItemData(cboDocumento.NewIndex) = rec!TCO_CODIGO
            rec.MoveNext
        Loop
        cboDocumento.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub ConfiguroGrillaRecibo()
    If TipodeAnulacion = 8 Then
        GrdModulos.FormatString = "^Tipo Rec|^Número|^Fecha|Importe|Proveedor|^Estado|codigo estado|" _
                                & "codigo estado que cambio|TIPO RECIBO|COD CLIENTE|REPRESENTADA"
    Else
        GrdModulos.FormatString = "^Tipo Rec|^Número|^Fecha|Importe|Cliente|^Estado|codigo estado|" _
                                & "codigo estado que cambio|TIPO RECIBO|COD CLIENTE|REPRESENTADA"
    End If
    GrdModulos.ColWidth(0) = 1000 'TIPO_RECIBO
    GrdModulos.ColWidth(1) = 1300 'NRO RECIBO
    GrdModulos.ColWidth(2) = 1100 'FECHA_RECIBO
    GrdModulos.ColWidth(3) = 1100 'IMPORTE
    GrdModulos.ColWidth(4) = 3000 'CLIENTE
    GrdModulos.ColWidth(5) = 2000 'ESTADO
    GrdModulos.ColWidth(6) = 0    'CODIGO ESTADO
    GrdModulos.ColWidth(7) = 0    'CODIGO ESTADO QUE CAMBIO
    GrdModulos.ColWidth(8) = 0    'TIPO RECIBO (TCO_CODIGO)
    GrdModulos.ColWidth(9) = 0    'CODIGO CLIENTE
    GrdModulos.ColWidth(10) = 0    'REPRESENTADA
    GrdModulos.Cols = 11
    GrdModulos.Rows = 2
    
End Sub

Private Sub ConfiguroGrillaNotaDC()
    GrdModulos.FormatString = "^Tipo|^Número|^Fecha|Cliente|^Estado|codigo estado|" _
                            & "codigo estado QUE CAMBIO|TIPO Nota credito|COD CLIENTE|"
                            
    GrdModulos.ColWidth(0) = 1000 'TIPO_NOTA
    GrdModulos.ColWidth(1) = 1300 'NRO NOTA
    GrdModulos.ColWidth(2) = 1200 'FECHA
    GrdModulos.ColWidth(3) = 3900 'CLIENTE
    GrdModulos.ColWidth(4) = 2000 'ESTADO
    GrdModulos.ColWidth(5) = 0    'CODIGO ESTADO
    GrdModulos.ColWidth(6) = 0    'CODIGO ESTADO QUE CAMBIO
    GrdModulos.ColWidth(7) = 0    'TIPO nota credito (TCO_CODIGO)
    GrdModulos.ColWidth(8) = 0    'CODIGO CLIENTE
    GrdModulos.Cols = 9
    GrdModulos.Rows = 2
End Sub

Private Sub ConfiguroGrillaFactura()
    GrdModulos.FormatString = "^Tipo|^Número|^Fecha|Cliente|^Estado|codigo estado|" _
                            & "codigo estado QUE CAMBIO|TIPO FACTURA|COD CLIENTE"
                                                        
    GrdModulos.ColWidth(0) = 1000 'TIPO_NOTA
    GrdModulos.ColWidth(1) = 1300 'NRO FACTURA
    GrdModulos.ColWidth(2) = 1200 'FECHA_FACTURA
    GrdModulos.ColWidth(3) = 3900 'CLIENTE
    GrdModulos.ColWidth(4) = 2000 'ESTADO
    GrdModulos.ColWidth(5) = 0    'CODIGO ESTADO
    GrdModulos.ColWidth(6) = 0    'CODIGO ESTADO QUE CAMBIO
    GrdModulos.ColWidth(7) = 0    'TIPO FACTURA (TCO_CODIGO)
    GrdModulos.ColWidth(8) = 0    'CODIGO CLIENTE
    GrdModulos.Cols = 9
    GrdModulos.Rows = 2
End Sub

Private Sub LlenarComboFactura()
    sql = "SELECT * FROM TIPO_COMPROBANTE"
    sql = sql & " WHERE TCO_DESCRI LIKE 'FAC%'"
    sql = sql & " ORDER BY TCO_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboDocumento.AddItem "(Todos)"
        Do While rec.EOF = False
            cboDocumento.AddItem rec!TCO_DESCRI
            cboDocumento.ItemData(cboDocumento.NewIndex) = rec!TCO_CODIGO
            rec.MoveNext
        Loop
        cboDocumento.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub LlenarComboRecibo()
    sql = "SELECT * FROM TIPO_COMPROBANTE"
    sql = sql & " WHERE TCO_DESCRI LIKE 'RECIB%'"
    sql = sql & " ORDER BY TCO_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboDocumento.AddItem "(Todos)"
        Do While rec.EOF = False
            cboDocumento.AddItem rec!TCO_DESCRI
            cboDocumento.ItemData(cboDocumento.NewIndex) = rec!TCO_CODIGO
            rec.MoveNext
        Loop
        cboDocumento.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub GrdModulos_dblClick()
    If GrdModulos.Rows > 1 Then
        Select Case TipodeAnulacion
            Case 3 'FACTURAS
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = 2 Then
                    MsgBox "No se puede cambiar el estado a la Factura" & Chr(13) & _
                           "el mimo ya fue Anulado", vbExclamation, TIT_MSGBOX
                    
                    Exit Sub
                End If
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 3 Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 2
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 4) = "ANULADO"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbRed)
                     
                ElseIf GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 2 Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 3
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 4) = "DEFINITIVO"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbBlack)
                End If
                
            Case 4 'RECIBOS
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 2 Then
                    MsgBox "No se puede cambiar el estado al Recibo" & Chr(13) & _
                           ",el mismo ya fue Anulado", vbExclamation, TIT_MSGBOX
                    
                    Exit Sub
                End If
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 7) = 3 Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 7) = 2
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = "ANULADO"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbRed)
                     
                ElseIf GrdModulos.TextMatrix(GrdModulos.RowSel, 7) = 2 Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 7) = 3
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = "DEFINITIVO"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbBlack)
                End If
            
            Case 5 'NOTA DE CREDITO
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = 2 Then
                    MsgBox "No se puede cambiar el estado a la Nota de Crédito" & Chr(13) & _
                           "la misma ya fue Anulada", vbExclamation, TIT_MSGBOX
                    
                    Exit Sub
                End If
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 3 Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 2
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 4) = "ANULADO"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbRed)
                     
                ElseIf GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 2 Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 3
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 4) = "DEFINITIVO"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbBlack)
                End If
                
            Case 6 'NOTA DE DEBITO
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = 2 Then
                    MsgBox "No se puede cambiar el estado a la Nota de Débito" & Chr(13) & _
                           "la misma ya fue Anulada", vbExclamation, TIT_MSGBOX
                    
                    Exit Sub
                End If
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 3 Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 2
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 4) = "ANULADO"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbRed)
                     
                ElseIf GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 2 Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 3
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 4) = "DEFINITIVO"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbBlack)
                End If
        End Select
    End If
End Sub

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then GrdModulos_dblClick
End Sub

Private Sub CmdNuevo_Click()
    txtCliente.Text = ""
    txtDesCli.Text = ""
    FechaDesde.Value = ""
    FechaHasta.Value = ""
    cboDocumento.ListIndex = 0
    GrdModulos.Rows = 1
    GrdModulos.Rows = 2
    txtCliente.SetFocus
End Sub

Private Sub txtCliente_Change()
    If txtCliente.Text = "" Then
        txtDesCli.Text = ""
    End If
End Sub

Private Sub txtCliente_GotFocus()
    SelecTexto txtCliente
End Sub

Private Sub txtCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarClientes txtCliente, "CODIGO"
    End If
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCliente_LostFocus()
    If txtCliente.Text <> "" Then
        Set rec = New ADODB.Recordset
        sql = "SELECT C.CLI_CODIGO,C.CLI_RAZSOC"
        sql = sql & " FROM CLIENTE C"
        sql = sql & " WHERE"
        sql = sql & " CLI_CODIGO =" & XN(txtCliente.Text)
        If rec.State = 1 Then rec.Close
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtDesCli.Text = rec!CLI_RAZSOC
        Else
            MsgBox "El Cliente no existe", vbExclamation, TIT_MSGBOX
            txtCliente.SetFocus
        End If
        If rec.State = 1 Then rec.Close
    End If
End Sub

Private Sub txtDesCli_Change()
    If txtDesCli.Text = "" Then
        txtCliente.Text = ""
    End If
End Sub

Private Sub txtDesCli_GotFocus()
    SelecTexto txtDesCli
End Sub

Private Sub txtDesCli_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtDesCli_LostFocus()
    If txtCliente.Text = "" And txtDesCli.Text <> "" Then
        Set rec = New ADODB.Recordset
        sql = "SELECT C.CLI_CODIGO,C.CLI_RAZSOC"
        sql = sql & " FROM CLIENTE C"
        sql = sql & " WHERE"
        sql = sql & " CLI_RAZSOC LIKE '" & XN(Trim(txtDesCli.Text)) & "%'"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If rec.RecordCount > 1 Then
                BuscarClientes txtCliente, "CADENA", Trim(txtDesCli.Text)
                If rec.State = 1 Then rec.Close
                txtDesCli.SetFocus
            Else
                txtCliente.Text = rec!CLI_CODIGO
                txtDesCli.Text = rec!CLI_RAZSOC
            End If
        Else
            MsgBox "No se encontro el Cliente", vbExclamation, TIT_MSGBOX
            txtCliente.SetFocus
        End If
        If rec.State = 1 Then rec.Close
    End If
End Sub

Public Sub BuscarClientes(Txt As Control, mQuien As String, Optional mCadena As String)
    Dim cSQL As String
    Dim hSQL As String
    Dim B As CBusqueda
    Dim i, posicion As Integer
    Dim cadena As String
    
    Set B = New CBusqueda
    With B
        cSQL = "SELECT CLI_RAZSOC, CLI_CODIGO"
        cSQL = cSQL & " FROM CLIENTE C"
        If mQuien = "CADENA" Then
            cSQL = cSQL & " WHERE CLI_RAZSOC LIKE '" & Trim(mCadena) & "%'"
        End If
        
        hSQL = "Nombre, Código"
        .sql = cSQL
        .Headers = hSQL
        .Field = "CLI_RAZSOC"
        campo1 = .Field
        .Field = "CLI_CODIGO"
        campo2 = .Field
        .OrderBy = "CLI_RAZSOC"
        camponumerico = False
        .Titulo = "Busqueda de Clientes :"
        .MaxRecords = 1
        .Show

        ' utilizar la coleccion de datos devueltos
        If .ResultFields.Count > 0 Then
            txtCliente.Text = .ResultFields(2)
            txtCliente_LostFocus
        End If
    End With
    
    Set B = Nothing
End Sub

Private Sub BuscoRecibos()
    lblEstado.Caption = "Buscando Recibos..."
    Screen.MousePointer = vbHourglass
    
    sql = "SELECT RC.REC_NUMERO, RC.REC_SUCURSAL, RC.REC_FECHA,RC.REC_TOTAL,"
    sql = sql & " RC.TCO_CODIGO, TC.TCO_ABREVIA,RC.CLI_CODIGO,"
    sql = sql & " C.CLI_RAZSOC, E.EST_DESCRI, RC.EST_CODIGO"
    sql = sql & " FROM RECIBO_CLIENTE RC, CLIENTE C, ESTADO_DOCUMENTO E, TIPO_COMPROBANTE TC"
    sql = sql & " WHERE"
    sql = sql & " RC.TCO_CODIGO=TC.TCO_CODIGO"
    sql = sql & " AND RC.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & " AND RC.EST_CODIGO=E.EST_CODIGO"
    If txtCliente.Text <> "" Then sql = sql & " AND RC.CLI_CODIGO=" & XN(txtCliente)
    'If txtVendedor.Text <> "" Then sql = sql & " AND RC.VEN_CODIGO=" & XN(txtVendedor)
    If Not IsNull(FechaDesde) Then sql = sql & " AND RC.REC_FECHA>=" & XDQ(FechaDesde)
    If Not IsNull(FechaHasta) Then sql = sql & " AND RC.REC_FECHA<=" & XDQ(FechaHasta)
    'If chkTipo.Value = Checked Then sql = sql & " AND RC.TCO_CODIGO=" & XN(cboDocumento.ItemData(cboDocumento.ListIndex))
    sql = sql & " ORDER BY RC.REC_SUCURSAL, RC.REC_NUMERO"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            GrdModulos.AddItem rec!TCO_ABREVIA & Chr(9) & Format(rec!REC_SUCURSAL, "0000") & "-" & Format(rec!REC_NUMERO, "00000000") _
                               & Chr(9) & rec!REC_FECHA & Chr(9) & Format(rec!REC_TOTAL, "#0.00") & Chr(9) & rec!CLI_RAZSOC _
                               & Chr(9) & rec!EST_DESCRI & Chr(9) & rec!EST_CODIGO _
                               & Chr(9) & rec!EST_CODIGO & Chr(9) & rec!TCO_CODIGO _
                               & Chr(9) & rec!CLI_CODIGO & Chr(9) & ""
                               
            If rec!EST_CODIGO = 2 Then
                Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.Rows - 1, vbRed)
            ElseIf rec!EST_CODIGO = 1 Then
                Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.Rows - 1, vbBlue)
            End If
            rec.MoveNext
        Loop
        GrdModulos.SetFocus
    Else
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        MsgBox "No se encontraron Recibos...", vbExclamation, TIT_MSGBOX
        'chkCliente.SetFocus
    End If
    rec.Close
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
End Sub

Private Sub ActualizoRecibo()
    Dim SaldoFactura As String
    SaldoFactura = "0"
    
    For i = 1 To GrdModulos.Rows - 1
        If GrdModulos.TextMatrix(i, 6) <> GrdModulos.TextMatrix(i, 7) Then
            Set rec = New ADODB.Recordset
            
            sql = "UPDATE RECIBO_CLIENTE"
            sql = sql & " SET EST_CODIGO=" & XN(GrdModulos.TextMatrix(i, 7))
            sql = sql & " WHERE"
            sql = sql & " TCO_CODIGO=" & XN(GrdModulos.TextMatrix(i, 8))
            sql = sql & " AND REC_NUMERO=" & XN(Right(GrdModulos.TextMatrix(i, 1), 8))
            sql = sql & " AND REC_SUCURSAL=" & XN(Left(GrdModulos.TextMatrix(i, 1), 4))
            DBConn.Execute sql
            
            'ACTUALIZO EL SALDO DE LAS FACTURAS
            sql = "SELECT FR.FCL_TCO_CODIGO, FR.FCL_NUMERO, FR.FCL_SUCURSAL, FR.FCL_FECHA,"
            sql = sql & " FR.REC_IMPORTE,FC.FCL_SALDO"
            sql = sql & " FROM FACTURAS_RECIBO_CLIENTE FR, FACTURA_CLIENTE FC"
            sql = sql & " WHERE"
            sql = sql & " FR.TCO_CODIGO=" & XN(GrdModulos.TextMatrix(i, 8))
            sql = sql & " AND FR.REC_NUMERO=" & XN(Right(GrdModulos.TextMatrix(i, 1), 8))
            sql = sql & " AND FR.REC_SUCURSAL=" & XN(Left(GrdModulos.TextMatrix(i, 1), 4))
            sql = sql & " AND FR.FCL_TCO_CODIGO=FC.TCO_CODIGO"
            sql = sql & " AND FR.FCL_NUMERO=FC.FCL_NUMERO"
            sql = sql & " AND FR.FCL_SUCURSAL=FC.FCL_SUCURSAL"
            
            
            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If rec.EOF = False Then
                Do While rec.EOF = False
                    SaldoFactura = CDbl(rec!REC_IMPORTE) + CDbl(rec!FCL_SALDO)
                    sql = "UPDATE FACTURA_CLIENTE"
                    sql = sql & " SET FCL_SALDO=" & XN(SaldoFactura)
                    sql = sql & " WHERE"
                    sql = sql & " TCO_CODIGO=" & XN(rec!FCL_TCO_CODIGO)
                    sql = sql & " AND FCL_NUMERO=" & XN(rec!FCL_NUMERO)
                    sql = sql & " AND FCL_SUCURSAL=" & XN(rec!FCL_SUCURSAL)
                    DBConn.Execute sql
                    SaldoFactura = "0"
                    rec.MoveNext
                Loop
            End If
            rec.Close
            
            'ACTUALIZO EL SALDO DE EL DINERO A CTA DEL CLIENTE
            sql = "SELECT RS.TCO_CODIGO, RS.REC_NUMERO, RS.REC_SUCURSAL,"
            sql = sql & " RS.REC_FECHA, RS.REC_SALDO, DR.DRE_COMIMP"
            sql = sql & " FROM DETALLE_RECIBO_CLIENTE DR, RECIBO_CLIENTE_SALDO RS"
            sql = sql & " WHERE"
            sql = sql & " RS.TCO_CODIGO=DR.DRE_TCO_CODIGO"
            sql = sql & " AND RS.REC_NUMERO=DR.DRE_COMNUMERO"
            sql = sql & " AND RS.REC_SUCURSAL=DR.DRE_COMSUCURSAL"
            sql = sql & " AND DR.TCO_CODIGO=" & XN(GrdModulos.TextMatrix(i, 8))
            sql = sql & " AND DR.REC_NUMERO=" & XN(Right(GrdModulos.TextMatrix(i, 1), 8))
            sql = sql & " AND DR.REC_SUCURSAL=" & XN(Left(GrdModulos.TextMatrix(i, 1), 4))
            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            
            If rec.EOF = False Then
                Do While rec.EOF = False
                    SaldoFactura = CDbl(rec!DRE_COMIMP) + CDbl(rec!REC_SALDO)
                    sql = "UPDATE RECIBO_CLIENTE_SALDO"
                    sql = sql & " SET REC_SALDO=" & XN(SaldoFactura)
                    sql = sql & " WHERE"
                    sql = sql & " TCO_CODIGO=" & XN(rec!TCO_CODIGO)
                    sql = sql & " AND REC_NUMERO=" & XN(rec!REC_NUMERO)
                    sql = sql & " AND REC_SUCURSAL=" & XN(rec!REC_SUCURSAL)
                    DBConn.Execute sql
                    
                    SaldoFactura = "0"
                    rec.MoveNext
                Loop
            End If
            rec.Close
            
            ''ACTUALIZO LA CTA-CTE
            'DBConn.Execute QuitoCtaCteCliente(GrdModulos.TextMatrix(i, 9), GrdModulos.TextMatrix(i, 8), _
                                              Right(GrdModulos.TextMatrix(i, 1), 8), Left(GrdModulos.TextMatrix(i, 1), 4))
        End If
    Next
End Sub

