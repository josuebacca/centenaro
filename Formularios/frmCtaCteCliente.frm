VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCtaCteCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cta-Cte Clientes"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7995
   ControlBox      =   0   'False
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
   ScaleHeight     =   7755
   ScaleWidth      =   7995
   Begin VB.TextBox txtintaux 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3600
      MaxLength       =   40
      TabIndex        =   51
      Text            =   "0,00"
      Top             =   5160
      Width           =   1320
   End
   Begin VB.Frame Frame4 
      Caption         =   "Saldos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   50
      TabIndex        =   42
      Top             =   5160
      Width           =   7890
      Begin VB.TextBox txtintereses 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3600
         MaxLength       =   40
         TabIndex        =   48
         Text            =   "0,00"
         Top             =   360
         Width           =   1320
      End
      Begin VB.TextBox txtdebe 
         Enabled         =   0   'False
         Height          =   315
         Left            =   840
         MaxLength       =   40
         TabIndex        =   47
         Text            =   "0,00"
         Top             =   360
         Width           =   1320
      End
      Begin VB.TextBox txtsaldototal 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6360
         MaxLength       =   40
         TabIndex        =   43
         Text            =   "0,00"
         Top             =   360
         Width           =   1320
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Intereses:"
         Height          =   195
         Left            =   2760
         TabIndex        =   46
         Top             =   360
         Width           =   750
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Total:"
         Height          =   195
         Left            =   5400
         TabIndex        =   45
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Debe:"
         Height          =   195
         Left            =   360
         TabIndex        =   44
         Top             =   390
         Width           =   435
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Pagos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   50
      TabIndex        =   36
      Top             =   1920
      Width           =   7890
      Begin VB.TextBox txtinttot 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6120
         MaxLength       =   40
         TabIndex        =   50
         Text            =   "0,00"
         Top             =   2880
         Width           =   1320
      End
      Begin VB.TextBox txtchetot 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4680
         MaxLength       =   40
         TabIndex        =   49
         Text            =   "0,00"
         Top             =   2880
         Width           =   1320
      End
      Begin VB.CommandButton cmdquitarTodos 
         Height          =   375
         Left            =   7200
         Picture         =   "frmCtaCteCliente.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Eliminar todos los registros"
         Top             =   480
         Width           =   495
      End
      Begin VB.CommandButton cmdQuitar 
         Height          =   375
         Left            =   6720
         Picture         =   "frmCtaCteCliente.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Eliminar registro"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtchemonto 
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
         Left            =   4920
         MaxLength       =   50
         TabIndex        =   10
         Tag             =   "Descripción"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtcheban 
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
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   8
         Tag             =   "Descripción"
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtchenro 
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
         Left            =   120
         MaxLength       =   50
         TabIndex        =   7
         Tag             =   "Descripción"
         Top             =   480
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid GrillaCheques 
         Height          =   2010
         Left            =   120
         TabIndex        =   37
         Top             =   840
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   3545
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorSel    =   8388736
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
      End
      Begin MSComCtl2.DTPicker fechacob 
         Height          =   315
         Left            =   3480
         TabIndex        =   9
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   51970049
         CurrentDate     =   41098
      End
      Begin VB.CommandButton cmdAgregar 
         Height          =   375
         Left            =   6240
         Picture         =   "frmCtaCteCliente.frx":0714
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Agregar registro"
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Monto"
         Height          =   255
         Left            =   4920
         TabIndex        =   41
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Fecha Cobro"
         Height          =   255
         Left            =   3480
         TabIndex        =   40
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Banco"
         Height          =   255
         Left            =   1560
         TabIndex        =   39
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Numero Cheque"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   50
      TabIndex        =   32
      Top             =   1200
      Width           =   7890
      Begin VB.CheckBox chkinteres 
         Caption         =   "Aplica Interes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   0
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.TextBox txtporcinteres 
         Height          =   315
         Left            =   6120
         MaxLength       =   40
         TabIndex        =   6
         Text            =   "0,12"
         Top             =   240
         Width           =   480
      End
      Begin MSComCtl2.DTPicker fechainteres 
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   51970049
         CurrentDate     =   41098
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta la fecha:"
         Height          =   255
         Left            =   360
         TabIndex        =   35
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Porcentaje de Interes:"
         Height          =   255
         Left            =   4440
         TabIndex        =   34
         Top             =   270
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "%"
         Height          =   255
         Left            =   6720
         TabIndex        =   33
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Movimientos.."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   4965
      TabIndex        =   28
      Top             =   6045
      Width           =   2970
      Begin VB.OptionButton optSaldosHistoricos 
         Caption         =   "Saldos Historicos"
         Height          =   225
         Left            =   1365
         TabIndex        =   16
         Top             =   315
         Width           =   1545
      End
      Begin VB.OptionButton optSaldos 
         Caption         =   "Saldos"
         Height          =   225
         Left            =   105
         TabIndex        =   14
         Top             =   315
         Width           =   990
      End
      Begin VB.OptionButton optPendiente 
         Caption         =   "Pendientes"
         Height          =   225
         Left            =   105
         TabIndex        =   15
         Top             =   660
         Width           =   1155
      End
      Begin VB.OptionButton optTodo 
         Caption         =   "Todos"
         Height          =   195
         Left            =   1365
         TabIndex        =   17
         Top             =   660
         Value           =   -1  'True
         Width           =   1500
      End
   End
   Begin VB.Frame FrameImpresora 
      Caption         =   "Impresora"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   45
      TabIndex        =   25
      Top             =   6045
      Width           =   4920
      Begin VB.CommandButton CBImpresora 
         Caption         =   "Configurar Impresora"
         Height          =   345
         Left            =   195
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   600
         Width           =   1755
      End
      Begin VB.OptionButton optPantalla 
         Caption         =   "Pantalla"
         Height          =   195
         Left            =   1020
         TabIndex        =   21
         Top             =   315
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton optImpresora 
         Caption         =   "Impresora"
         Height          =   195
         Left            =   2085
         TabIndex        =   22
         Top             =   315
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   225
         TabIndex        =   26
         Top             =   300
         Width           =   600
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   405
      Left            =   6945
      TabIndex        =   20
      Top             =   7245
      Width           =   975
   End
   Begin VB.Frame frameBuscar 
      Caption         =   "Ver..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   50
      TabIndex        =   23
      Top             =   0
      Width           =   7890
      Begin VB.TextBox txtCliente 
         Height          =   315
         Left            =   1620
         MaxLength       =   40
         TabIndex        =   0
         Top             =   345
         Width           =   720
      End
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
         Left            =   2370
         MaxLength       =   50
         TabIndex        =   1
         Tag             =   "Descripción"
         Top             =   345
         Width           =   4575
      End
      Begin MSComCtl2.DTPicker FechaDesde 
         Height          =   315
         Left            =   1620
         TabIndex        =   2
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   51970049
         CurrentDate     =   41098
      End
      Begin MSComCtl2.DTPicker FechaHasta 
         Height          =   315
         Left            =   5520
         TabIndex        =   3
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   51970049
         CurrentDate     =   41098
      End
      Begin VB.Label lblFechaHasta 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   4605
         TabIndex        =   30
         Top             =   735
         Width           =   960
      End
      Begin VB.Label lblFechaDesde 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   525
         TabIndex        =   29
         Top             =   720
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
         Left            =   525
         TabIndex        =   24
         Top             =   390
         Width           =   555
      End
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   3885
      Top             =   7005
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   3390
      Top             =   6945
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "&Listar"
      Height          =   405
      Left            =   4980
      TabIndex        =   18
      Top             =   7245
      Width           =   960
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   405
      Left            =   5955
      TabIndex        =   19
      Top             =   7245
      Width           =   975
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
      Left            =   120
      TabIndex        =   31
      Top             =   7155
      Width           =   660
   End
End
Attribute VB_Name = "frmCtaCteCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Saldo As Double
Dim Cliente As Integer
Dim Orden As Integer
Dim interes As Double

Private Sub CBImpresora_Click()
    CDImpresora.PrinterDefault = True
    CDImpresora.ShowPrinter
    FrameImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
End Sub

Private Sub BuscarCtaCTeClientes()
       
    sql = "DELETE FROM CTA_CTE_CLIENTE"
    DBConn.Execute sql
    
    If optPendiente.Value = True Or optSaldos.Value = True Then
        
        'FACTURAS PENDIENTES
        sql = "INSERT INTO CTA_CTE_CLIENTE (CLI_CODIGO,TCO_CODIGO,COM_NUMERO,COM_SUCURSAL,"
        sql = sql & " COM_FECHA,COM_IMPORTE,COM_IMP_DEBE,COM_IMP_HABER,CTA_CTE_DH,COM_NUMEROTXT)"
        sql = sql & " SELECT F.CLI_CODIGO,F.TCO_CODIGO,F.FCL_NUMERO,F.FCL_SUCURSAL,"
        sql = sql & " F.FCL_FECHA,F.FCL_TOTAL,F.FCL_SALDO,0 AS HABER,'D' AS DEBE,FCL_NUMEROTXT"
        sql = sql & " FROM SALDO_FACTURAS_CLIENTE_V F"
        sql = sql & " WHERE"
        sql = sql & " F.EST_CODIGO=3"
        sql = sql & " AND F.FCL_SALDO > 0"
        If txtCliente.Text <> "" Then
            sql = sql & " AND F.CLI_CODIGO=" & XN(txtCliente.Text)
        End If
        If FechaDesde.Value <> "" Then
            sql = sql & " AND F.FCL_FECHA>=" & XDQ(FechaDesde.Value)
        End If
        If FechaHasta.Value <> "" Then
            sql = sql & " AND F.FCL_FECHA<=" & XDQ(FechaHasta.Value)
        End If
        DBConn.Execute sql

        'NOTA DEBITOS CLIENTE PENDIENTES
        sql = "INSERT INTO CTA_CTE_CLIENTE (CLI_CODIGO,TCO_CODIGO,COM_NUMERO,COM_SUCURSAL,"
        sql = sql & " COM_FECHA,COM_IMPORTE,COM_IMP_DEBE,COM_IMP_HABER,CTA_CTE_DH,"
        sql = sql & " COM_NUMEROTXT)"
        sql = sql & " SELECT DISTINCT N.CLI_CODIGO,N.TCO_CODIGO,N.NDC_NUMERO,N.NDC_SUCURSAL,"
        sql = sql & " N.NDC_FECHA,N.NDC_TOTAL,N.NDC_SALDO,0 AS HABER,'D' AS DEBE,N.NDC_NUMEROTXT"
        sql = sql & " FROM NOTA_DEBITO_CLIENTE N"
        sql = sql & " WHERE N.EST_CODIGO=3"
        sql = sql & " AND N.NDC_SALDO > 0"
        If txtCliente.Text <> "" Then
            sql = sql & " AND N.CLI_CODIGO=" & XN(txtCliente.Text)
        End If
        If FechaDesde.Value <> "" Then
            sql = sql & " AND N.NDC_FECHA>=" & XDQ(FechaDesde.Value)
        End If
        If FechaHasta.Value <> "" Then
            sql = sql & " AND N.NDC_FECHA<=" & XDQ(FechaHasta.Value)
        End If
        DBConn.Execute sql
        
        'NOTA CREDITO CLIENTE PENDIENTES
        sql = "INSERT INTO CTA_CTE_CLIENTE (CLI_CODIGO,TCO_CODIGO,COM_NUMERO,COM_SUCURSAL,"
        sql = sql & " COM_FECHA,COM_IMPORTE,COM_IMP_DEBE,COM_IMP_HABER,CTA_CTE_DH,"
        sql = sql & " COM_NUMEROTXT)"
        sql = sql & " SELECT DISTINCT N.CLI_CODIGO,N.TCO_CODIGO,N.NCC_NUMERO,N.NCC_SUCURSAL,"
        sql = sql & " N.NCC_FECHA,N.NCC_TOTAL,0 AS DEBE,NCC_SALDO,'C' AS CREDITO,N.NCC_NUMEROTXT"
        sql = sql & " FROM NOTA_CREDITO_CLIENTE N"
        sql = sql & " WHERE N.EST_CODIGO=3"
        sql = sql & " AND N.NCC_SALDO > 0"
        If txtCliente.Text <> "" Then
            sql = sql & " AND N.CLI_CODIGO=" & XN(txtCliente.Text)
        End If
        If FechaDesde.Value <> "" Then
            sql = sql & " AND N.NCC_FECHA>=" & XDQ(FechaDesde.Value)
        End If
        If FechaHasta.Value <> "" Then
            sql = sql & " AND N.NCC_FECHA<=" & XDQ(FechaHasta.Value)
        End If
        DBConn.Execute sql
        
        'TODOS LOS RECIBOS CON SALDOS
        sql = "INSERT INTO CTA_CTE_CLIENTE (CLI_CODIGO,TCO_CODIGO,COM_NUMERO,COM_SUCURSAL,"
        sql = sql & " COM_FECHA,COM_IMPORTE,COM_IMP_DEBE,COM_IMP_HABER,CTA_CTE_DH,"
        sql = sql & " COM_NUMEROTXT)"
        sql = sql & " SELECT R.CLI_CODIGO,R.TCO_CODIGO,R.REC_NUMERO,R.REC_SUCURSAL,R.REC_FECHA,"
        sql = sql & " S.REC_SALDO AS TOTAL,0 AS DEBE,S.REC_SALDO AS HABER,'C' AS CREDITO,R.REC_NUMEROTXT"
        sql = sql & " FROM RECIBO_CLIENTE R , RECIBO_CLIENTE_SALDO S"
        sql = sql & " WHERE R.EST_CODIGO=3"
        sql = sql & " AND R.TCO_CODIGO=S.TCO_CODIGO"
        sql = sql & " AND R.REC_SUCURSAL=S.REC_SUCURSAL"
        sql = sql & " AND R.REC_NUMERO=S.REC_NUMERO"
        sql = sql & " AND S.REC_SALDO > 0"
        If txtCliente.Text <> "" Then
            sql = sql & " AND R.CLI_CODIGO=" & XN(txtCliente.Text)
        End If
        If FechaDesde.Value <> "" Then
            sql = sql & " AND R.REC_FECHA >= " & XDQ(FechaDesde.Value)
        End If
        If FechaHasta.Value <> "" Then
            sql = sql & " AND R.REC_FECHA <= " & XDQ(FechaHasta.Value)
        End If
        DBConn.Execute sql
    End If
    
    If optTodo.Value = True Or optSaldosHistoricos.Value = True Then
        If chkinteres.Value = Checked Then
            calculointeres fechainteres.Value, FechaHasta.Value
        Else
            interes = 0
        End If
        'ACTUALIZAR TOTAL DE FACTURAS PENDIENTES
        actualizototal
        
        'TODAS LAS FACTURAS
        sql = "INSERT INTO CTA_CTE_CLIENTE (CLI_CODIGO,TCO_CODIGO,COM_NUMERO,COM_SUCURSAL,"
        sql = sql & " COM_FECHA,COM_IMPORTE,COM_IMP_DEBE,COM_IMP_HABER,CTA_CTE_DH,COM_NUMEROTXT)"
        sql = sql & " SELECT F.CLI_CODIGO,F.TCO_CODIGO,F.FCL_NUMERO,F.FCL_SUCURSAL,"
        sql = sql & " F.FCL_FECHA,F.FCL_TOTAL,F.FCL_TOTALACT,0 AS HABER,'D' AS DEBE,FCL_NUMEROTXT"
        sql = sql & " FROM FACTURA_CLIENTE F"
        sql = sql & " WHERE F.EST_CODIGO=3"
        sql = sql & " AND FPG_CODIGO=2"
        If txtCliente.Text <> "" Then
            sql = sql & " AND F.CLI_CODIGO=" & XN(txtCliente.Text)
        End If
        If FechaDesde.Value <> "" Then
            sql = sql & " AND F.FCL_FECHA >= " & XDQ(FechaDesde.Value)
        End If
        If FechaHasta.Value <> "" Then
            sql = sql & " AND F.FCL_FECHA <= " & XDQ(FechaHasta.Value)
        End If
        DBConn.Execute sql
    
        'TODAS LAS NOTAS DEBITOS CLIENTE
        sql = "INSERT INTO CTA_CTE_CLIENTE (CLI_CODIGO,TCO_CODIGO,COM_NUMERO,COM_SUCURSAL,"
        sql = sql & " COM_FECHA,COM_IMPORTE,COM_IMP_DEBE,COM_IMP_HABER,CTA_CTE_DH,"
        sql = sql & " COM_NUMEROTXT)"
        sql = sql & " SELECT DISTINCT N.CLI_CODIGO,N.TCO_CODIGO,N.NDC_NUMERO,N.NDC_SUCURSAL,"
        sql = sql & " N.NDC_FECHA,N.NDC_TOTAL,N.NDC_TOTAL,0 AS HABER,'D' AS DEBE,N.NDC_NUMEROTXT"
        sql = sql & " FROM NOTA_DEBITO_CLIENTE N"
        sql = sql & " WHERE N.EST_CODIGO=3"
        If txtCliente.Text <> "" Then
            sql = sql & " AND N.CLI_CODIGO=" & XN(txtCliente.Text)
        End If
        If FechaDesde.Value <> "" Then
            sql = sql & " AND N.NDC_FECHA >= " & XDQ(FechaDesde.Value)
        End If
        If FechaHasta.Value <> "" Then
            sql = sql & " AND N.NDC_FECHA <= " & XDQ(FechaHasta.Value)
        End If
        DBConn.Execute sql
        
        'TODAS LAS NOTAS CREDITO CLIENTE
        sql = "INSERT INTO CTA_CTE_CLIENTE (CLI_CODIGO,TCO_CODIGO,COM_NUMERO,COM_SUCURSAL,"
        sql = sql & " COM_FECHA,COM_IMPORTE,COM_IMP_DEBE,COM_IMP_HABER,CTA_CTE_DH,"
        sql = sql & " COM_NUMEROTXT)"
        sql = sql & " SELECT DISTINCT N.CLI_CODIGO,N.TCO_CODIGO,N.NCC_NUMERO,N.NCC_SUCURSAL,"
        sql = sql & " N.NCC_FECHA,N.NCC_TOTAL,0 AS DEBE,NCC_TOTAL,'C' AS CREDITO,N.NCC_NUMEROTXT"
        sql = sql & " FROM NOTA_CREDITO_CLIENTE N"
        sql = sql & " WHERE N.EST_CODIGO=3"
        If txtCliente.Text <> "" Then
            sql = sql & " AND N.CLI_CODIGO=" & XN(txtCliente.Text)
        End If
        If FechaDesde.Value <> "" Then
            sql = sql & " AND N.NCC_FECHA >= " & XDQ(FechaDesde.Value)
        End If
        If FechaHasta.Value <> "" Then
            sql = sql & " AND N.NCC_FECHA <= " & XDQ(FechaHasta.Value)
        End If
        DBConn.Execute sql
        
        'TODOS LOS RECIBOS
        sql = "INSERT INTO CTA_CTE_CLIENTE (CLI_CODIGO,TCO_CODIGO,COM_NUMERO,COM_SUCURSAL,"
        sql = sql & " COM_FECHA,COM_IMPORTE,COM_IMP_DEBE,COM_IMP_HABER,CTA_CTE_DH,"
        sql = sql & " COM_NUMEROTXT)"
        sql = sql & " SELECT DISTINCT R.CLI_CODIGO,R.TCO_CODIGO,R.REC_NUMERO,R.REC_SUCURSAL,"
        sql = sql & " R.REC_FECHA,R.REC_TOTAL,0 AS DEBE,REC_TOTAL,'C' AS CREDITO,R.REC_NUMEROTXT"
        sql = sql & " FROM RECIBO_CLIENTE R"
        sql = sql & " WHERE R.EST_CODIGO=3"
        If txtCliente.Text <> "" Then
            sql = sql & " AND R.CLI_CODIGO=" & XN(txtCliente.Text)
        End If
        If FechaDesde.Value <> "" Then
            sql = sql & " AND R.REC_FECHA >= " & XDQ(FechaDesde.Value)
        End If
        If FechaHasta.Value <> "" Then
            sql = sql & " AND R.REC_FECHA <= " & XDQ(FechaHasta.Value)
        End If
        DBConn.Execute sql
        
        'ACTUALIZO INTERES
        sql = "UPDATE CTA_CTE_CLIENTE SET CTA_CTE_INTERES=" & XN(CStr(interes))
        DBConn.Execute sql
                
        'TODOS LOS RECIBOS CON SALDOS
'        sql = " SELECT R.CLI_CODIGO,R.TCO_CODIGO,R.REC_NUMERO,R.REC_SUCURSAL,"
'        sql = sql & " R.REC_FECHA,(R.REC_TOTAL+S.REC_SALDO) AS TOTAL,R.REC_NUMEROTXT"
'        sql = sql & " FROM RECIBO_CLIENTE R , RECIBO_CLIENTE_SALDO S"
'        sql = sql & " WHERE R.EST_CODIGO=3"
'        sql = sql & " AND R.TCO_CODIGO=S.TCO_CODIGO"
'        sql = sql & " AND R.REC_SUCURSAL=S.REC_SUCURSAL"
'        sql = sql & " AND R.REC_NUMERO=S.REC_NUMERO"
'        If txtCliente.Text <> "" Then
'            sql = sql & " AND R.CLI_CODIGO=" & XN(txtCliente.Text)
'        End If
'        If FechaDesde.value <> "" Then
'            sql = sql & " AND R.REC_FECHA >= " & XDQ(FechaDesde.value)
'        End If
'        If FechaHasta.value <> "" Then
'            sql = sql & " AND R.REC_FECHA <= " & XDQ(FechaHasta.value)
'        End If
'        Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'        If Rec.EOF = False Then
'            Do While Rec.EOF = False
'                sql = "DELETE FROM CTA_CTE_CLIENTE"
'                sql = sql & " WHERE"
'                sql = sql & " CLI_CODIGO=" & XN(Rec!CLI_CODIGO)
'                sql = sql & " AND TCO_CODIGO=" & XN(Rec!TCO_CODIGO)
'                sql = sql & " AND COM_NUMERO=" & XN(Rec!REC_NUMERO)
'                sql = sql & " AND COM_SUCURSAL=" & XN(Rec!REC_SUCURSAL)
'                DBConn.Execute sql
'
'                sql = "INSERT INTO CTA_CTE_CLIENTE (CLI_CODIGO,TCO_CODIGO,COM_NUMERO,COM_SUCURSAL,"
'                sql = sql & " COM_FECHA,COM_IMPORTE,COM_IMP_DEBE,COM_IMP_HABER,CTA_CTE_DH,"
'                sql = sql & " COM_NUMEROTXT)"
'                sql = sql & " VALUES ("
'                sql = sql & XN(Rec!CLI_CODIGO) & ","
'                sql = sql & XN(Rec!TCO_CODIGO) & ","
'                sql = sql & XN(Rec!REC_NUMERO) & ","
'                sql = sql & XN(Rec!REC_SUCURSAL) & ","
'                sql = sql & XDQ(Rec!REC_FECHA) & ","
'                sql = sql & XN(Rec!TOTAL) & ","
'                sql = sql & XN("0") & ","
'                sql = sql & XN(Rec!TOTAL) & ","
'                sql = sql & XS("C") & ","
'                sql = sql & XS(Rec!REC_NUMEROTXT) & ")"
'                DBConn.Execute sql
                
'                Rec.MoveNext
'            Loop
'        End If
'        Rec.Close
    End If
    If optSaldos.Value = True Or optSaldosHistoricos.Value = True Then
        BuscaSaldosGeneral
    Else
        BuscaSaldosDetalle
    End If
End Sub
Private Function actualizototal()
    Dim TIPO_FAC As Integer
    Dim FACTURA As String
    Dim NUEVO_TOTAL As String
    Dim i As Integer
    sql = "SELECT * FROM "
    sql = sql & " FACTURA_CLIENTE F, DETALLE_FACTURA_CLIENTE DF, PRODUCTO P"
    sql = sql & " WHERE F.EST_CODIGO=3"
    sql = sql & " AND F.FPG_CODIGO=2"
    sql = sql & " AND FPG_CODIGO=2"
    sql = sql & " AND F.FCL_SUCURSAL=DF.FCL_SUCURSAL"
    sql = sql & " AND F.FCL_NUMERO=DF.FCL_NUMERO"
    sql = sql & " AND F.TCO_CODIGO=DF.TCO_CODIGO"
    sql = sql & " AND DF.PTO_CODIGO=P.PTO_CODIGO"
    If txtCliente.Text <> "" Then
        sql = sql & " AND F.CLI_CODIGO=" & XN(txtCliente.Text)
    End If
    If FechaDesde.Value <> "" Then
        sql = sql & " AND F.FCL_FECHA >= " & XDQ(FechaDesde.Value)
    End If
    If FechaHasta.Value <> "" Then
        sql = sql & " AND F.FCL_FECHA <= " & XDQ(FechaHasta.Value)
    End If
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        FACTURA = 0
        Do While rec.EOF = False
            If FACTURA <> rec!FCL_NUMERO Then
                FACTURA = rec!FCL_NUMERO
                NUEVO_TOTAL = rec!FCL_TOTAL
                TIPO_FAC = rec!TCO_CODIGO
            
                If rec!PTO_PRECTO > rec!DFC_PRECIO Then
                    NUEVO_TOTAL = rec!DFC_CANTIDAD * rec!PTO_PRECTO
                End If
            Else
                If rec!PTO_PRECTO > rec!DFC_PRECIO Then
                    NUEVO_TOTAL = NUEVO_TOTAL + (rec!DFC_CANTIDAD * rec!PTO_PRECTO)
                End If
            End If
            sql = "UPDATE FACTURA_CLIENTE SET FCL_TOTALACT=" & XN(NUEVO_TOTAL)
            sql = sql & " WHERE "
            sql = sql & " TCO_CODIGO=" & TIPO_FAC
            sql = sql & " AND FCL_NUMERO=" & FACTURA
            sql = sql & " AND FCL_SUCURSAL=" & XN(rec!FCL_SUCURSAL)
            DBConn.Execute sql

            rec.MoveNext
        Loop
    End If
    rec.Close
End Function
Private Sub calculointeres(fecha_interes As Date, fecha_hasta As Date)
    Dim MES As Integer ' mes consultado
    Dim anio As Integer ' año consultado
    Dim messig As Integer ' mes siguiente o el que le aplico interes
    Dim aniosig As Integer 'mes siguiente o el que le aplico interes
    Dim fechacomienzoint As Date 'fecha desde la cual aplico el interes
    Dim dias As Integer
    
    MES = Month(fecha_hasta)
    anio = Year(fecha_hasta)
    
    If MES = 12 Then
        messig = 1
        aniosig = anio + 1
    Else
        messig = MES + 1
        aniosig = Year(fecha_hasta)
    End If
    fechacomienzoint = "10/" & messig & "/" & aniosig
    If fecha_interes > fechacomienzoint Then
        dias = fecha_interes - fechacomienzoint
        interes = dias * txtporcinteres.Text
    Else
        interes = 0
    End If
    
    
End Sub
Private Sub BuscaSaldosDetalle()
    'CONFIGURO EL SALDO
    sql = "SELECT * FROM CTA_CTE_CLIENTE"
    sql = sql & " ORDER BY CLI_CODIGO,COM_FECHA"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = False Then
        Cliente = rec!CLI_CODIGO
        Saldo = 0
        Orden = 1
        Do While rec.EOF = False
            If rec!CTA_CTE_DH = "D" Then
                Saldo = Saldo + CDbl(Chk0(rec!COM_IMP_DEBE))
            Else
                Saldo = Saldo - CDbl(Chk0(rec!COM_IMP_HABER))
            End If
            sql = "UPDATE CTA_CTE_CLIENTE SET CTA_CTA_SALDO=" & XN(CStr(Saldo))
            sql = sql & " ,CTA_CTE_ORDEN=" & XN(CStr(Orden))
            sql = sql & " ,COM_NUMEROTXT=" & XS(Format(rec!COM_NUMERO, "00000000"))
            sql = sql & " WHERE CLI_CODIGO=" & XN(rec!CLI_CODIGO)
            sql = sql & " AND TCO_CODIGO=" & XN(rec!TCO_CODIGO)
            sql = sql & " AND COM_NUMERO=" & XN(rec!COM_NUMERO)
            sql = sql & " AND COM_SUCURSAL=" & XN(rec!COM_SUCURSAL)
            DBConn.Execute sql
            
            Orden = Orden + 1
            rec.MoveNext
            If rec.EOF = False Then
                'SI NO VA DETALLADO POR REPRESENTADA
                If Cliente <> rec!CLI_CODIGO Then
                    Cliente = rec!CLI_CODIGO
                    Saldo = 0
                    Orden = 1
                End If
            End If
        Loop
    End If
    rec.Close
End Sub

Private Sub BuscaSaldosGeneral()
    'CONFIGURO EL SALDO
    sql = "SELECT SUM(COM_IMP_DEBE) AS DEBE ,SUM(COM_IMP_HABER)AS HABER "
    sql = sql & " ,CLI_CODIGO"
    sql = sql & " FROM CTA_CTE_CLIENTE"
    sql = sql & " GROUP BY CLI_CODIGO"
    sql = sql & " ORDER BY CLI_CODIGO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = False Then
        Saldo = 0
        Do While rec.EOF = False
             Saldo = CDbl(rec!DEBE) - CDbl(rec!HABER)
             sql = "DELETE FROM CTA_CTE_CLIENTE"
             sql = sql & " WHERE CLI_CODIGO=" & XN(rec!CLI_CODIGO)
             DBConn.Execute sql
             
            'If Saldo > 0 Then
                sql = "INSERT INTO CTA_CTE_CLIENTE (CLI_CODIGO,TCO_CODIGO,COM_NUMERO,"
                sql = sql & "COM_SUCURSAL,CTA_CTE_SALDOFINAL)"
                sql = sql & " VALUES ("
                sql = sql & XN(rec!CLI_CODIGO) & ","
                sql = sql & XN("1") & ","
                sql = sql & XN("1") & ","
                sql = sql & XN("1") & ","
                sql = sql & XN(CStr(Saldo)) & ")"
                DBConn.Execute sql
            'End If
            Saldo = 0
            rec.MoveNext
        Loop
    End If
    rec.Close
End Sub

Private Sub chkinteres_Click()
    If chkinteres.Value = Unchecked Then
        fechainteres.Value = ""
        txtporcinteres.Text = "0,00"
    Else
        fechainteres.Value = Date
        txtporcinteres.Text = "0,12"
    End If
End Sub

Private Sub CmdAgregar_Click()
    Dim montointeres As Double
    Dim porc_capital As Double
    'calcular interes
    calculointeres fechacob.Value, FechaHasta.Value
    
    porc_capital = (CDbl(txtchemonto.Text) / txtdebe.Text)
    'interes = interes * capital
    
    'interes = interes - interes * porc_capital
    If fechacob.Value <> "" And txtchemonto.Text <> "" Then
        GrillaCheques.AddItem txtchenro.Text & Chr(9) & _
                              txtcheban.Text & Chr(9) & _
                              Format(fechacob.Value, "dd/mm/yyyy") & Chr(9) & _
                              txtchemonto.Text & Chr(9) & _
                              Format(interes, "#,##0.00")
        sumatotales
    End If
    
End Sub
Private Function sumatotales()
    Dim suma_interes As Double
    Dim i As Integer
    suma_interes = 0
    If GrillaCheques.Rows > 1 Then
        For i = 1 To GrillaCheques.Rows - 1
            suma_interes = suma_interes + CDbl(GrillaCheques.TextMatrix(i, 4))
        Next
    End If
    txtinttot.Text = suma_interes
    txtinttot.Text = Valido_Importe2(txtinttot.Text)
    'txtintereses.Text = (CDbl(txtdebe.Text) * (interes + suma_interes)) / 100
    'txtintereses.Text = Valido_Importe2(txtintereses)
    'txtsaldototal.Text = CDbl(txtdebe) + CDbl(txtintereses)
    'txtsaldototal.Text = Valido_Importe2(txtsaldototal)
End Function

Private Sub cmdListar_Click()
    On Error GoTo CLAVOSE
    
    Screen.MousePointer = vbHourglass
    lblestado.Caption = "Buscando..."
    'LLENO LA TABLA CTA_CTE_CLIENTE
    BuscarCtaCTeClientes
    
    DBConn.Execute "DELETE FROM TMP_RESUMEN_CUENTA_BANCO"
        
    Rep.WindowState = crptMaximized
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""

    If FechaDesde.Value <> "" And FechaHasta.Value <> "" Then
        Rep.Formulas(0) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & FechaHasta.Value & "'"
    ElseIf FechaDesde.Value <> "" And FechaHasta.Value = "" Then
        Rep.Formulas(0) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & Date & "'"
    ElseIf FechaDesde.Value = "" And FechaHasta.Value <> "" Then
        Rep.Formulas(0) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & FechaHasta.Value & "'"
    ElseIf FechaDesde.Value = "" And FechaHasta.Value = "" Then
        Rep.Formulas(0) = "FECHA='" & " Al: " & Date & "'"
    End If
    
    Rep.WindowTitle = "CTA-CTE de Clientes..."
    If optPendiente.Value = True Or optTodo.Value = True Then
        Rep.ReportFileName = DRIVE & DirReport & "ctacte_clientes.rpt"
    Else
        Rep.ReportFileName = DRIVE & DirReport & "ctacte_clientes_Saldos.rpt"
    End If
    If optPantalla.Value = True Then
        Rep.Destination = crptToWindow
    ElseIf optImpresora.Value = True Then
        Rep.Destination = crptToPrinter
    End If
    Rep.Action = 1
     
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    Screen.MousePointer = vbNormal
    lblestado.Caption = ""
    Exit Sub
    
CLAVOSE:
    If rec.State = 1 Then rec.Close
    Screen.MousePointer = vbNormal
    lblestado.Caption = ""
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub CmdNuevo_Click()
    txtCliente.Text = ""
    lblestado.Caption = ""
    FechaDesde.Value = ""
    FechaHasta.Value = ""
    optTodo.Value = True
    
    GrillaCheques.Rows = 1
    txtchenro.Text = ""
    txtcheban.Text = ""
    fechacob.Value = ""
    txtchemonto.Text = ""
    
    txtdebe = "0,00"
    txtintereses = "0,00"
    txtsaldototal = "0,00"
End Sub

Private Sub cmdQuitar_Click()
    If GrillaCheques.Rows > 2 Then
        GrillaCheques.RemoveItem (GrillaCheques.RowSel)
    Else
        GrillaCheques.Rows = 1
    End If
    sumatotales
End Sub

Private Sub cmdquitarTodos_Click()
    GrillaCheques.Rows = 1
    txtinttot.Text = "0,00"
End Sub

Private Sub CmdSalir_Click()
    Set frmCtaCteCliente = Nothing
    Unload Me
End Sub

Private Sub fechacob_GotFocus()
    seltxt
End Sub

Private Sub FechaHasta_LostFocus()
'    If GrillaCheques.Rows = 1 And FechaDesde.Value <> "" And FechaHasta.Value <> "" Then
'        BuscarCtaCTeClientes
'        rec.Open "SELECT SUM(COM_IMP_DEBE) AS DEBE,CTA_CTE_INTERES FROM CTA_CTE_CLIENTE GROUP BY CTA_CTE_INTERES", DBConn, adOpenStatic, adLockOptimistic
'        If rec.EOF = False Then
'            txtdebe.Text = Chk0(rec!DEBE)
'            txtdebe.Text = Valido_Importe2(txtdebe)
'            txtintereses = CDbl(txtdebe) * Chk0(rec!CTA_CTE_INTERES) / 100
'            txtintereses = Valido_Importe2(txtintereses)
'            txtintaux.Text = Chk0(rec!CTA_CTE_INTERES)
'            txtintaux.Text = Valido_Importe2(txtintaux.Text)
'            txtsaldototal = CDbl(txtdebe) + CDbl(txtintereses)
'            txtsaldototal = Valido_Importe2(txtsaldototal)
'        End If
'        rec.Close
'    End If
End Sub

Private Sub fechainteres_Click()
    If IsNull(fechainteres.Value) = True Then
        chkinteres.Value = Unchecked
        txtporcinteres.Text = "0,00"
    Else
        chkinteres.Value = Checked
        txtporcinteres.Text = "0,12"
    End If
End Sub

Private Sub fechainteres_LostFocus()
    FechaHasta_LostFocus
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then MySendKeys Chr(9)
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset

    Me.Left = 0
    Me.Top = 0
    FrameImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
    lblestado.Caption = ""
    fechainteres = Date
    preparogrilla
End Sub
Private Function preparogrilla()
'GRILLA CHEQUES
    GrillaCheques.FormatString = "Cheque Nro|Banco|Fecha Cobro|Importe|Interes"
    GrillaCheques.ColWidth(0) = 1300   'Cheque Nro
    GrillaCheques.ColWidth(1) = 2100   'Banco
    GrillaCheques.ColWidth(2) = 1300   'Fecha Cobro
    GrillaCheques.ColWidth(3) = 1300   'Importe
    GrillaCheques.ColWidth(4) = 1300   'Interes
    GrillaCheques.Rows = 1
End Function

Private Sub txtcheban_GotFocus()
    seltxt
End Sub

Private Sub txtcheban_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtchemonto_GotFocus()
    seltxt
End Sub

Private Sub txtchemonto_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtchemonto, KeyAscii)
End Sub

Private Sub txtchemonto_LostFocus()
    txtchemonto.Text = VALIDO_IMPORTE(txtchemonto.Text)
End Sub

Private Sub txtchenro_GotFocus()
    seltxt
End Sub

Private Sub txtchenro_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
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
        BuscarClientes "", "CODIGO"
    End If
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCliente_LostFocus()
    If txtCliente.Text <> "" Then
        Set rec = New ADODB.Recordset
        sql = "SELECT CLI_CODIGO, CLI_RAZSOC"
        sql = sql & " FROM CLIENTE"
        sql = sql & " WHERE "
        If txtCliente.Text <> "" Then
            sql = sql & " CLI_CODIGO=" & XN(txtCliente)
        End If
        If rec.State = 1 Then rec.Close
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtDesCli.Text = Trim(rec!CLI_RAZSOC)
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

Private Sub txtDesCli_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarClientes "", "CODIGO"
    End If
End Sub

Private Sub txtDesCli_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtDesCli_LostFocus()
    If txtCliente.Text = "" And txtDesCli.Text <> "" Then
        Set rec = New ADODB.Recordset
        sql = "SELECT CLI_CODIGO, CLI_RAZSOC"
        sql = sql & " FROM CLIENTE"
        sql = sql & " WHERE "
        sql = sql & " CLI_RAZSOC LIKE '%" & Trim(txtDesCli) & "%'"
        If rec.State = 1 Then rec.Close
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If rec.RecordCount > 1 Then
                BuscarClientes "", "CADENA", Trim(txtDesCli.Text)
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

Private Function BuscoCliente(Cli As String) As String
    sql = "SELECT CLI_CODIGO, CLI_RAZSOC"
    sql = sql & " FROM CLIENTE"
    sql = sql & " WHERE "
    If txtCliente.Text <> "" Then
        sql = sql & " CLI_CODIGO=" & XN(Cli)
    Else
        sql = sql & " CLI_RAZSOC LIKE '" & Cli & "%'"
    End If
    BuscoCliente = sql
End Function

Public Sub BuscarClientes(Txt As String, mQuien As String, Optional mCadena As String)
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



Private Sub txtporcinteres_GotFocus()
    seltxt
End Sub
