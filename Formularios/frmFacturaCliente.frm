VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{AFD24A52-2823-4FBD-B75D-C282C11E1D98}#1.0#0"; "IFEpson.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFacturaCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Factura de Clientes..."
   ClientHeight    =   7200
   ClientLeft      =   300
   ClientTop       =   1365
   ClientWidth     =   10680
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   10680
   Begin VB.Frame fracheque 
      Height          =   1845
      Left            =   2880
      TabIndex        =   134
      Top             =   3000
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton cmdcrerrarcheque 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   3450
         TabIndex        =   139
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtchenumero 
         Height          =   315
         Left            =   1665
         TabIndex        =   136
         Top             =   1005
         Width           =   2505
      End
      Begin VB.TextBox txtchebanco 
         Height          =   315
         Left            =   1665
         TabIndex        =   135
         Top             =   645
         Width           =   2505
      End
      Begin VB.CommandButton cmdaceptarcheque 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2100
         TabIndex        =   137
         Top             =   1320
         Width           =   1305
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero:"
         Height          =   315
         Left            =   405
         TabIndex        =   141
         Top             =   1005
         Width           =   1215
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Banco:"
         Height          =   315
         Left            =   405
         TabIndex        =   140
         Top             =   645
         Width           =   1215
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "Datos Cheque"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   30
         TabIndex        =   138
         Top             =   120
         Width           =   4845
      End
   End
   Begin VB.Frame fraFiscal 
      Caption         =   "Valores Fiscales"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2430
      Left            =   7200
      TabIndex        =   93
      Top             =   3840
      Visible         =   0   'False
      Width           =   3225
      Begin VB.TextBox txtNetoFiscal 
         Height          =   345
         Left            =   735
         TabIndex        =   96
         Top             =   750
         Width           =   1815
      End
      Begin VB.TextBox txtIvaFiscal 
         Height          =   345
         Left            =   735
         TabIndex        =   95
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtTotalFiscal 
         Height          =   345
         Left            =   735
         TabIndex        =   94
         Top             =   1650
         Width           =   1815
      End
      Begin VB.Label Label32 
         Caption         =   "neto"
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   99
         Top             =   810
         Width           =   855
      End
      Begin VB.Label Label32 
         Caption         =   "iva"
         Height          =   255
         Index           =   1
         Left            =   210
         TabIndex        =   98
         Top             =   1230
         Width           =   855
      End
      Begin VB.Label Label32 
         Caption         =   "total"
         Height          =   255
         Index           =   2
         Left            =   210
         TabIndex        =   97
         Top             =   1710
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdNC 
      Caption         =   "Nota de Credito"
      Enabled         =   0   'False
      Height          =   330
      Left            =   5160
      TabIndex        =   129
      Top             =   6825
      Width           =   1335
   End
   Begin VB.Frame fraTarjeta 
      Height          =   3285
      Left            =   2280
      TabIndex        =   59
      Top             =   1680
      Width           =   4935
      Begin VB.CommandButton cmdAceptoTarjeta 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2220
         TabIndex        =   68
         Top             =   2760
         Width           =   1425
      End
      Begin VB.TextBox txtLote 
         Height          =   315
         Left            =   1665
         TabIndex        =   64
         Top             =   1605
         Width           =   2505
      End
      Begin VB.TextBox txtCupon 
         Height          =   315
         Left            =   1665
         TabIndex        =   65
         Top             =   1965
         Width           =   2505
      End
      Begin VB.ComboBox cboPlan 
         Height          =   315
         ItemData        =   "frmFacturaCliente.frx":0000
         Left            =   1665
         List            =   "frmFacturaCliente.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   1245
         Width           =   2505
      End
      Begin VB.ComboBox cboTarjeta 
         Height          =   315
         ItemData        =   "frmFacturaCliente.frx":0004
         Left            =   1665
         List            =   "frmFacturaCliente.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   855
         Width           =   2505
      End
      Begin VB.TextBox txtTar_Autorizacion 
         Height          =   315
         Left            =   1665
         MaxLength       =   30
         TabIndex        =   66
         Top             =   2325
         Width           =   2505
      End
      Begin VB.CommandButton cmdCerrarTarjeta 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   3690
         TabIndex        =   70
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton cmdAltaTarjeta 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4260
         TabIndex        =   61
         ToolTipText     =   "Alta de Tarjeta"
         Top             =   870
         Width           =   480
      End
      Begin VB.CommandButton cmdAltaPlan 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4260
         TabIndex        =   60
         ToolTipText     =   "Alta de Plan"
         Top             =   1260
         Width           =   480
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "Datos Tarjeta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   30
         TabIndex        =   74
         Top             =   120
         Width           =   4845
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lote:"
         Height          =   315
         Left            =   405
         TabIndex        =   73
         Top             =   1605
         Width           =   1215
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cupón:"
         Height          =   315
         Left            =   405
         TabIndex        =   72
         Top             =   1965
         Width           =   1215
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Plan:"
         Height          =   315
         Left            =   405
         TabIndex        =   71
         Top             =   1245
         Width           =   1215
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tarjeta:"
         Height          =   315
         Left            =   405
         TabIndex        =   69
         Top             =   855
         Width           =   1215
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Autorización:"
         Height          =   315
         Left            =   405
         TabIndex        =   67
         Top             =   2325
         Width           =   1215
      End
   End
   Begin VB.Frame fraPagos 
      Height          =   5175
      Left            =   5520
      TabIndex        =   75
      Top             =   1440
      Width           =   4935
      Begin VB.TextBox txtImportePago 
         Height          =   315
         Left            =   1470
         TabIndex        =   83
         Top             =   1815
         Width           =   1245
      End
      Begin VB.CommandButton cmdAceptarPagos 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2160
         TabIndex        =   84
         Top             =   4695
         Width           =   1425
      End
      Begin VB.CommandButton cmdBorroFila 
         Caption         =   "Borrar Fila"
         Height          =   375
         Left            =   90
         TabIndex        =   81
         Top             =   4695
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         Height          =   795
         Left            =   120
         TabIndex        =   78
         Top             =   570
         Width           =   4695
         Begin VB.TextBox txtTotalPagos 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   375
            Left            =   3120
            TabIndex        =   79
            Top             =   300
            Width           =   1515
         End
         Begin VB.Label Label35 
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "T O T A L"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   90
            TabIndex        =   80
            Top             =   300
            Width           =   3015
         End
      End
      Begin VB.TextBox txtGrabar 
         Height          =   285
         Left            =   3540
         TabIndex        =   77
         Top             =   1920
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdCerrarPagos 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   3630
         TabIndex        =   76
         Top             =   4695
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid grdPagos 
         Height          =   2445
         Left            =   120
         TabIndex        =   85
         Top             =   2190
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   4313
         _Version        =   393216
         Rows            =   1
         Cols            =   15
         FixedCols       =   0
         ForeColorSel    =   12632064
         ScrollTrack     =   -1  'True
         FocusRect       =   2
         HighLight       =   2
         SelectionMode   =   1
         FormatString    =   $"frmFacturaCliente.frx":0008
      End
      Begin VB.ComboBox cboFormaPago 
         Height          =   315
         ItemData        =   "frmFacturaCliente.frx":000E
         Left            =   1470
         List            =   "frmFacturaCliente.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   82
         Top             =   1470
         Width           =   3330
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Importe:"
         Height          =   330
         Left            =   120
         TabIndex        =   88
         Top             =   1815
         Width           =   1320
      End
      Begin VB.Label Label40 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "Forma de Pago"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   45
         TabIndex        =   87
         Top             =   120
         Width           =   4845
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Forma Pago"
         Height          =   330
         Left            =   120
         TabIndex        =   86
         Top             =   1470
         Width           =   1320
      End
   End
   Begin VB.TextBox mProvincia 
      Height          =   285
      Left            =   5205
      TabIndex        =   91
      Top             =   6345
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox mLocalidad 
      Height          =   285
      Left            =   5220
      TabIndex        =   90
      Top             =   6510
      Visible         =   0   'False
      Width           =   1335
   End
   Begin EPSON_Impresora_Fiscal.PrinterFiscal pf 
      Left            =   1170
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   450
      Left            =   2850
      TabIndex        =   14
      Top             =   6705
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   330
      Left            =   9735
      TabIndex        =   16
      Top             =   6825
      Width           =   870
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   7920
      TabIndex        =   13
      Top             =   6825
      Width           =   870
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   8835
      TabIndex        =   15
      Top             =   6825
      Width           =   870
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   6735
      Left            =   15
      TabIndex        =   31
      Top             =   30
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   512
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Datos"
      TabPicture(0)   =   "frmFacturaCliente.frx":0012
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label12"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblblockeado"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "FrameFactura"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "FrameCliente"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtObservaciones"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "frmFacturaCliente.frx":002E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GrdModulos"
      Tab(1).Control(1)=   "frameBuscar"
      Tab(1).ControlCount=   2
      Begin VB.TextBox txtObservaciones 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1350
         MaxLength       =   60
         TabIndex        =   11
         Top             =   6360
         Width           =   9090
      End
      Begin VB.Frame FrameCliente 
         Caption         =   "Cliente..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   105
         TabIndex        =   45
         Top             =   345
         Width           =   5475
         Begin VB.TextBox txtIngBrutos 
            Enabled         =   0   'False
            Height          =   315
            Left            =   3855
            TabIndex        =   118
            Top             =   990
            Width           =   1350
         End
         Begin VB.TextBox txtNRO_DOCUMENTO 
            Enabled         =   0   'False
            Height          =   315
            Left            =   780
            TabIndex        =   117
            Top             =   990
            Width           =   1770
         End
         Begin VB.CommandButton cmdModificarCli 
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   12
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4545
            TabIndex        =   107
            ToolTipText     =   "Modificar Datos Cliente"
            Top             =   300
            Width           =   330
         End
         Begin VB.TextBox txtTelefono 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4920
            TabIndex        =   105
            Top             =   960
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.CommandButton cmdNuevoCli 
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   12
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4200
            TabIndex        =   104
            Top             =   300
            Width           =   330
         End
         Begin VB.CommandButton cmdbuscaComp 
            Height          =   315
            Left            =   4890
            Picture         =   "frmFacturaCliente.frx":004A
            Style           =   1  'Graphical
            TabIndex        =   102
            ToolTipText     =   "Buscar Cliente"
            Top             =   300
            Width           =   330
         End
         Begin VB.TextBox txtRazSoc 
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
            Left            =   1365
            TabIndex        =   5
            Top             =   300
            Width           =   2805
         End
         Begin VB.TextBox txtDomici 
            Enabled         =   0   'False
            Height          =   315
            Left            =   780
            TabIndex        =   47
            Top             =   645
            Width           =   4440
         End
         Begin VB.TextBox txtCiva 
            Enabled         =   0   'False
            Height          =   315
            Left            =   780
            TabIndex        =   46
            Top             =   1335
            Width           =   2265
         End
         Begin MSMask.MaskEdBox txtCuit 
            Height          =   315
            Left            =   3855
            TabIndex        =   51
            Top             =   1335
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   13
            Mask            =   "##-########-#"
            PromptChar      =   "_"
         End
         Begin VB.TextBox mRespo 
            Height          =   315
            Left            =   2205
            TabIndex        =   92
            Top             =   1365
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.TextBox txtcodCli 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   780
            TabIndex        =   0
            Top             =   300
            Width           =   570
         End
         Begin VB.Label lblSaldoFac 
            Caption         =   "Saldo Facturacion: $"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1440
            TabIndex        =   133
            Top             =   0
            Visible         =   0   'False
            Width           =   3855
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Ing. Brutos:"
            Height          =   195
            Left            =   2880
            TabIndex        =   116
            Top             =   1005
            Width           =   870
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Teléfono:"
            Height          =   195
            Left            =   4680
            TabIndex        =   106
            Top             =   960
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Nro Doc:"
            Height          =   195
            Left            =   90
            TabIndex        =   100
            Top             =   1040
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "C.U.I.T.:"
            Height          =   195
            Index           =   10
            Left            =   3135
            TabIndex        =   52
            Top             =   1395
            Width           =   660
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
            Height          =   195
            Left            =   90
            TabIndex        =   50
            Top             =   330
            Width           =   555
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio:"
            Height          =   195
            Left            =   90
            TabIndex        =   49
            Top             =   685
            Width           =   660
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   " I.V.A.:"
            Height          =   195
            Left            =   90
            TabIndex        =   48
            Top             =   1395
            Width           =   540
         End
      End
      Begin VB.Frame frameBuscar 
         Caption         =   "Buscar por..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1830
         Left            =   -74805
         TabIndex        =   35
         Top             =   420
         Width           =   10230
         Begin VB.TextBox txtBuscaNum 
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
            Left            =   6945
            TabIndex        =   24
            Top             =   980
            Width           =   1290
         End
         Begin VB.ComboBox cboTurnosB 
            Height          =   315
            Left            =   2505
            Style           =   2  'Dropdown List
            TabIndex        =   124
            Top             =   1320
            Width           =   1635
         End
         Begin VB.CommandButton cmdBuscaCli 
            Height          =   315
            Left            =   7830
            Picture         =   "frmFacturaCliente.frx":03D4
            Style           =   1  'Graphical
            TabIndex        =   103
            ToolTipText     =   "Buscar Cliente"
            Top             =   300
            Width           =   375
         End
         Begin VB.TextBox txtBuscarCliDescri 
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
            Left            =   3390
            MaxLength       =   50
            TabIndex        =   19
            Tag             =   "Descripción"
            Top             =   300
            Width           =   4395
         End
         Begin VB.TextBox txtBuscaCliente 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   2490
            MaxLength       =   40
            TabIndex        =   18
            Top             =   300
            Width           =   870
         End
         Begin VB.ComboBox cboFactura1 
            Height          =   315
            Left            =   2490
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   980
            Width           =   2400
         End
         Begin VB.CommandButton CmdBuscAprox 
            Caption         =   "Buscar"
            Height          =   390
            Left            =   8535
            MaskColor       =   &H000000FF&
            TabIndex        =   25
            ToolTipText     =   "Buscar "
            Top             =   915
            UseMaskColor    =   -1  'True
            Width           =   1335
         End
         Begin VB.TextBox txtBuscaSuc 
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
            Left            =   6360
            MaxLength       =   4
            TabIndex        =   23
            Top             =   980
            Width           =   555
         End
         Begin MSComCtl2.DTPicker FechaDesde 
            Height          =   315
            Left            =   2490
            TabIndex        =   20
            Top             =   655
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   54132737
            CurrentDate     =   41098
         End
         Begin MSComCtl2.DTPicker FechaHasta 
            Height          =   315
            Left            =   6360
            TabIndex        =   21
            Top             =   660
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   54132737
            CurrentDate     =   41098
         End
         Begin VB.Label lblFechaDesde 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   1395
            TabIndex        =   132
            Top             =   720
            Width           =   990
         End
         Begin VB.Label lblFechaHasta 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   5355
            TabIndex        =   131
            Top             =   720
            Width           =   960
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Nro Factura:"
            Height          =   195
            Left            =   5400
            TabIndex        =   130
            Top             =   1035
            Width           =   915
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Turno:"
            Height          =   195
            Left            =   1395
            TabIndex        =   125
            Top             =   1365
            Width           =   480
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   4
            Left            =   1395
            TabIndex        =   53
            Top             =   375
            Width           =   555
         End
         Begin VB.Label lbltipoFac 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Factura:"
            Height          =   195
            Left            =   1395
            TabIndex        =   44
            Top             =   1040
            Width           =   960
         End
      End
      Begin VB.Frame FrameFactura 
         Caption         =   "Factura..."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   5700
         TabIndex        =   33
         Top             =   345
         Width           =   4800
         Begin VB.TextBox txtNroSucursal 
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
            Left            =   840
            MaxLength       =   4
            TabIndex        =   2
            Top             =   690
            Width           =   555
         End
         Begin VB.ComboBox cboFactura 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   345
            Width           =   1890
         End
         Begin VB.TextBox txtNroFactura 
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
            Left            =   1420
            TabIndex        =   3
            Top             =   690
            Width           =   1290
         End
         Begin MSComCtl2.DTPicker FechaFactura 
            Height          =   315
            Left            =   840
            TabIndex        =   4
            Top             =   1080
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   54132737
            CurrentDate     =   41098
         End
         Begin VB.Label Ltipo_fac 
            AutoSize        =   -1  'True
            Caption         =   "A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   36
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   2970
            TabIndex        =   101
            Top             =   285
            Width           =   495
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   135
            TabIndex        =   41
            Top             =   375
            Width           =   360
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   135
            TabIndex        =   40
            Top             =   1065
            Width           =   495
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   135
            TabIndex        =   39
            Top             =   705
            Width           =   615
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   135
            TabIndex        =   38
            Top             =   1425
            Width           =   555
         End
         Begin VB.Label lblEstadoFactura 
            AutoSize        =   -1  'True
            Caption         =   "EST. FACTURA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   780
            TabIndex        =   37
            Top             =   1425
            Width           =   1170
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   4245
         Left            =   -74835
         TabIndex        =   26
         Top             =   2400
         Width           =   10260
         _ExtentX        =   18098
         _ExtentY        =   7488
         _Version        =   393216
         Cols            =   13
         FixedCols       =   0
         RowHeightMin    =   280
         BackColorSel    =   16761024
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         HighLight       =   2
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
      Begin VB.Frame Frame1 
         Height          =   615
         Left            =   105
         TabIndex        =   54
         Top             =   1965
         Width           =   10395
         Begin VB.ComboBox cboTurno 
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
            Left            =   4665
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   210
            Width           =   1635
         End
         Begin VB.ComboBox cboListaPrecio 
            Height          =   315
            Left            =   7860
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   210
            Width           =   2355
         End
         Begin VB.ComboBox cboVendedor 
            Height          =   315
            Left            =   825
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   210
            Width           =   2745
         End
         Begin VB.ComboBox cboCondicion 
            Height          =   315
            Left            =   8025
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   195
            Visible         =   0   'False
            Width           =   1860
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Turno:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4080
            TabIndex        =   123
            Top             =   255
            Width           =   465
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Lst Precio:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7035
            TabIndex        =   57
            Top             =   255
            Width           =   750
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Playero:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   60
            TabIndex        =   56
            Top             =   255
            Width           =   570
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Condición:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7125
            TabIndex        =   55
            Top             =   240
            Visible         =   0   'False
            Width           =   810
         End
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3810
         Left            =   105
         TabIndex        =   34
         Top             =   2505
         Width           =   10395
         Begin VB.TextBox txttasavial 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00008000&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   7200
            Locked          =   -1  'True
            TabIndex        =   127
            Top             =   3360
            Width           =   1410
         End
         Begin VB.TextBox txtSubtotalB 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0FF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   3420
            Locked          =   -1  'True
            TabIndex        =   122
            Top             =   2640
            Visible         =   0   'False
            Width           =   1650
         End
         Begin VB.TextBox txtImporteIvaB 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0FF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   5070
            Locked          =   -1  'True
            TabIndex        =   121
            Top             =   2640
            Visible         =   0   'False
            Width           =   1650
         End
         Begin VB.TextBox txtimpuestoB 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0FF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   1770
            Locked          =   -1  'True
            TabIndex        =   120
            Top             =   2640
            Visible         =   0   'False
            Width           =   1650
         End
         Begin VB.TextBox txtsubtotal1B 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0FF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   119
            Top             =   2640
            Visible         =   0   'False
            Width           =   1650
         End
         Begin VB.TextBox txtsubtotal1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FF0000&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   105
            Locked          =   -1  'True
            TabIndex        =   114
            Top             =   3360
            Width           =   1410
         End
         Begin VB.TextBox txtnoinsc 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FF0000&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   5745
            Locked          =   -1  'True
            TabIndex        =   112
            Top             =   3360
            Width           =   1410
         End
         Begin VB.TextBox txtimpuesto 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FF0000&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   108
            Top             =   3360
            Width           =   1410
         End
         Begin VB.TextBox txtImporteIva 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FF0000&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   4335
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   3360
            Width           =   1410
         End
         Begin VB.TextBox txtPorcentajeIva 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FF0000&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   -615
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   3600
            Visible         =   0   'False
            Width           =   1600
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   8715
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   3360
            Width           =   1530
         End
         Begin VB.TextBox txtSubtotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FF0000&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   2925
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   3360
            Width           =   1410
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H0000FFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   300
            TabIndex        =   27
            Top             =   510
            Visible         =   0   'False
            Width           =   1185
         End
         Begin MSFlexGridLib.MSFlexGrid grdGrilla 
            Height          =   2850
            Left            =   75
            TabIndex        =   10
            Top             =   165
            Width           =   10230
            _ExtentX        =   18045
            _ExtentY        =   5027
            _Version        =   393216
            Rows            =   3
            Cols            =   6
            FixedCols       =   0
            RowHeightMin    =   290
            BackColorSel    =   65535
            ForeColorSel    =   0
            GridColor       =   -2147483633
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            HighLight       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "TASA VIAL."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   7200
            TabIndex        =   128
            Top             =   3060
            Width           =   1400
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "IVA NO INSC."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   5745
            TabIndex        =   113
            Top             =   3060
            Width           =   1410
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "TOTAL"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   8715
            TabIndex        =   111
            Top             =   3060
            Width           =   1530
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "IVA INSC."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   4335
            TabIndex        =   110
            Top             =   3060
            Width           =   1410
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "SUB-TOTAL          "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2925
            TabIndex        =   109
            Top             =   3060
            Width           =   1410
         End
         Begin VB.Label lblConPago 
            AutoSize        =   -1  'True
            Caption         =   "Con Pago"
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
            Left            =   105
            TabIndex        =   58
            Top             =   3720
            Width           =   900
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "IMPUESTO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1515
            TabIndex        =   43
            Top             =   3060
            Width           =   1410
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "SUBT-TOTAL"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   42
            Top             =   3060
            Width           =   1410
         End
      End
      Begin VB.Label lblblockeado 
         Caption         =   "Label42"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4560
         TabIndex        =   142
         Top             =   0
         Visible         =   0   'False
         Width           =   6615
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones:"
         Height          =   195
         Left            =   120
         TabIndex        =   115
         Top             =   6360
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ordenado por :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   -74820
         TabIndex        =   32
         Top             =   570
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdFormaPago 
      Caption         =   "Forma Pago"
      Height          =   330
      Left            =   6510
      TabIndex        =   12
      Top             =   6825
      Width           =   1380
   End
   Begin VB.TextBox mDireccion 
      Height          =   285
      Left            =   48
      TabIndex        =   89
      Top             =   6828
      Visible         =   0   'False
      Width           =   1335
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   1620
      Top             =   6765
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Presione <F5> para actualizar Turno/Fecha"
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
      Left            =   2520
      TabIndex        =   126
      Top             =   6840
      Width           =   3690
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
      Left            =   195
      TabIndex        =   36
      Top             =   6990
      Width           =   660
   End
End
Attribute VB_Name = "frmFacturaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim i As Integer
Dim W As Integer
Dim J As Integer
Dim VBonificacion As Double
Dim VTotal As Double
Dim VEstadoFactura As Integer
Dim SImporte As String  'importe en letras para imprimir
Dim mFoco As Boolean
Dim mFormaPago As String
Public mQuienLlama As String
Public mQueFacturo As String
Public mDescrip As String
Private mRespuestaFiscal As Boolean
Dim mPrecio As Double
Dim mBuscador As Boolean
Dim mVerCta As Boolean
Dim mValorCta As Double
Dim mValorIvaIns As Double
'Dim mValIVA As Double
Dim mIVA_1 As Double
Dim mIVA_2 As Double
Dim SaldoCli As Double


Private Sub cboFactura_Click()
    If cboFactura.ListIndex = -1 Then Exit Sub
    If cboFactura.ItemData(cboFactura.ListIndex) = 1 Then
        Ltipo_fac.Caption = "A"
    ElseIf cboFactura.ItemData(cboFactura.ListIndex) = 1 Then
        Ltipo_fac.Caption = "B"
    End If
End Sub

'Private Sub cboCondicion_LostFocus()
'    If cboCondicion.ListIndex <> -1 Then
'        sql = "SELECT * FROM FORMA_PAGO"
'        sql = sql & " WHERE FPG_CODIGO=" & cboCondicion.ItemData(cboCondicion.ListIndex)
'        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'        If rec.EOF = False Then
'            If Not IsNull(rec!FPG_PORCEN) And CDbl(rec!FPG_PORCEN) <> 0 Then
'                If CDbl(rec!FPG_PORCEN) > 0 Then
'                    mFormaPago = (CDbl(rec!FPG_PORCEN) / 100) + 1
'                    lblConPago.Caption = "Sobre el Precio de Lista se Aplica un Incremento del " & Format(rec!FPG_PORCEN, "0.00") & " %"
'                Else
'                    mFormaPago = (CDbl(rec!FPG_PORCEN) / 100) + 1
'                    lblConPago.Caption = "Sobre el Precio de Lista se Aplica un Descuento del " & Format(CDbl(rec!FPG_PORCEN) * -1, "0.00") & " %"
'                End If
'            Else
'                mFormaPago = 0
'                lblConPago.Caption = ""
'            End If
'        Else
'            mFormaPago = 0
'            lblConPago.Caption = ""
'        End If
'        rec.Close
'    Else
'        mFormaPago = 0
'        lblConPago.Caption = ""
'    End If
'End Sub

Private Sub cboFormaPago_LostFocus()
    If Me.ActiveControl.Name = "grdPagos" Then
        Exit Sub
    End If
    If txtcodCli.Text = "1" Then
        If cboFormaPago.ItemData(cboFormaPago.ListIndex) = 2 Then
            MsgBox "No Puede Seleccionar Cta CTe para este Cliente", vbCritical, TIT_MSGBOX
            cboFormaPago.ListIndex = 0
            cboFormaPago.SetFocus
            Exit Sub
        End If
    End If
    fraTarjeta.Visible = False
    If Trim(cboFormaPago.Text) = "TARJETA DE CREDITO" Then
        cboPlan.Clear
        cboTarjeta.Clear
        cSQL = "SELECT TAR_CODIGO, TAR_DESCRI"
        cSQL = cSQL & " FROM TARJETA"
        cSQL = cSQL & " WHERE TTA_CODIGO=1" 'SOLO TARJETA DE CREDITO
        cSQL = cSQL & " ORDER BY TAR_DESCRI"
        rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
        If (rec.BOF And rec.EOF) = 0 Then
           Do While rec.EOF = False
              cboTarjeta.AddItem Trim(rec!TAR_DESCRI)
              cboTarjeta.ItemData(cboTarjeta.NewIndex) = rec!TAR_CODIGO
              rec.MoveNext
           Loop
           If cboTarjeta.ListCount > 0 Then cboTarjeta.ListIndex = 0
        End If
        rec.Close
        
        fraTarjeta.Top = 1485
        fraTarjeta.Left = 3330
        fraTarjeta.Visible = True
        cboTarjeta.SetFocus
        cboPlan.Enabled = True
        txtLote.Enabled = True
        txtCupon.Enabled = True
        txtTar_Autorizacion.Enabled = True
    End If
    
    If Trim(cboFormaPago.Text) = "TARJETA DE DEBITO" Then
        cboPlan.Clear
        cboTarjeta.Clear
        cSQL = "SELECT TAR_CODIGO, TAR_DESCRI"
        cSQL = cSQL & " FROM TARJETA"
        cSQL = cSQL & " WHERE TTA_CODIGO=2" 'SOLO TARJETA DE DEBITO
        cSQL = cSQL & " ORDER BY TAR_DESCRI"
        rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
        If (rec.BOF And rec.EOF) = 0 Then
           Do While rec.EOF = False
              cboTarjeta.AddItem Trim(rec!TAR_DESCRI)
              cboTarjeta.ItemData(cboTarjeta.NewIndex) = rec!TAR_CODIGO
              rec.MoveNext
           Loop
           If cboTarjeta.ListCount > 0 Then cboTarjeta.ListIndex = 0
        End If
        rec.Close
        
        fraTarjeta.Top = 1485
        fraTarjeta.Left = 3330
        fraTarjeta.Visible = True
        cboTarjeta.SetFocus
        cboPlan.Enabled = True
        txtLote.Enabled = True
        txtCupon.Enabled = True
        txtTar_Autorizacion.Enabled = True
    End If
    
    fracheque.Visible = False
    If Trim(cboFormaPago.Text) = "CHEQUE" Then
        'cboPlan.Clear
        'cboTarjeta.Clear
        'cSQL = "SELECT TAR_CODIGO, TAR_DESCRI"
        'cSQL = cSQL & " FROM TARJETA"
        'cSQL = cSQL & " WHERE TTA_CODIGO=1" 'SOLO TARJETA DE CREDITO
        'cSQL = cSQL & " ORDER BY TAR_DESCRI"
        'rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
        'If (rec.BOF And rec.EOF) = 0 Then
        '   Do While rec.EOF = False
        '      cboTarjeta.AddItem Trim(rec!TAR_DESCRI)
        '      cboTarjeta.ItemData(cboTarjeta.NewIndex) = rec!TAR_CODIGO
        '      rec.MoveNext
        '   Loop
        '   If cboTarjeta.ListCount > 0 Then cboTarjeta.ListIndex = 0
        'End If
        'rec.Close
        
        fracheque.Top = 1485
        fracheque.Left = 3330
        fracheque.Visible = True
        txtchebanco.SetFocus
        'cboPlan.Enabled = True
        'txtLote.Enabled = True
        'txtCupon.Enabled = True
        'txtTar_Autorizacion.Enabled = True
    End If
'    If Trim(UCase(Mid(cboFormaPago.Text, 1, 50))) = "DOLARES" Then
'        fraDolar.Top = 1980
'        fraDolar.Left = 3465
'        txtCotizacion.Text = Format(mCotiza, "0.00")
'        fraDolar.Visible = True
'        txtTotDolar.SetFocus
'    End If
'    If Trim(UCase(Mid(cboFormaPago.Text, 1, 50))) = "SE#A" Then
'        fraSenia.Visible = True
'        fraSenia.Top = 1880
'        fraSenia.Left = 1170
'        sql = "select v.suc_codigo, v.nrofac, v.tipo_fac, fecha, i.precio, i.descrip"
'        sql = sql & " from ventgral v, ventitem i"
'        sql = sql & " Where v.suc_codigo = i.suc_codigo"
'        sql = sql & " and v.tipo_fac = i.tipo_fac"
'        sql = sql & " and v.nrofac = i.nrofac"
'        sql = sql & " and codpieza = 'SENA'"
'        sql = sql & " and cliente = " & XN(mCodigo.Text)
'        sql = sql & " and SENIA_USADA = 'N'"
'        grdSenia.Rows = 1
'        If snp.State = 1 Then snp.Close
'        snp.Open sql, DBConn, adOpenStatic, adLockOptimistic
'        If snp.EOF = False Then
'            snp.MoveFirst
'            Do While Not snp.EOF
'
'                grdSenia.AddItem ("")
'                grdSenia.row = grdSenia.Rows - 1
'                grdSenia.TextMatrix(grdSenia.row, 0) = ChkNull(snp!suc_codigo)
'                grdSenia.TextMatrix(grdSenia.row, 1) = ChkNull(snp!TIPO_FAC)
'                grdSenia.TextMatrix(grdSenia.row, 2) = ChkNull(snp!NROFAC)
'                grdSenia.TextMatrix(grdSenia.row, 3) = ChkNull(snp!Fecha)
'                grdSenia.TextMatrix(grdSenia.row, 4) = ChkNull(snp!DESCRIP)
'                grdSenia.TextMatrix(grdSenia.row, 5) = Format(ChkNull(snp!precio), "0.00")
'
'                snp.MoveNext
'            Loop
'        End If
'        If grdSenia.Rows > 1 Then grdSenia.row = 1
'        grdSenia.SetFocus
'    End If
End Sub

Private Sub cboTarjeta_LostFocus()
    Dim mCodTar As String
    mCodTar = cboTarjeta.ItemData(cboTarjeta.ListIndex)
    cboPlan.Clear
    
    sql = "SELECT PLA_CODIGO, PLA_DESCRI"
    sql = sql & " FROM TARJETA_PLAN WHERE TAR_CODIGO = " & XN(mCodTar)
    sql = sql & " ORDER BY PLA_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboPlan.AddItem Trim(rec!PLA_DESCRI)
            cboPlan.ItemData(cboPlan.NewIndex) = rec!PLA_CODIGO
            rec.MoveNext
        Loop
    End If
    rec.Close
    If cboPlan.ListCount > 0 Then cboPlan.ListIndex = 0
End Sub

Private Sub cmdaceptarcheque_Click()
    fracheque.Visible = False
    'cboFormaPago.ListIndex = 0
    txtImportePago.SetFocus
End Sub

Private Sub cmdAceptarPagos_Click()
    If txtcodCli.Text = "1" Then
        If cboFormaPago.ItemData(cboFormaPago.ListIndex) = 2 Then
            MsgBox "No Puede Seleccionar Cta CTe para este Cliente", vbCritical, TIT_MSGBOX
            cboFormaPago.ListIndex = 0
            cboFormaPago.SetFocus
            Exit Sub
        End If
    End If
    fraPagos.Visible = False
    If cboVendedor.List(cboVendedor.ListIndex) = "" Then
        cboVendedor.SetFocus
    Else
        cmdGrabar.SetFocus
    End If
    
    If txtGrabar.Text = "S" Then
        'CBGrabar_Click
    Else
        'cboPara_Quien.SetFocus
    End If
    mValorCta = 0
    For i = 1 To grdPagos.Rows - 1
        If grdPagos.TextMatrix(i, 2) = "2" Then
            mValorCta = mValorCta + CDbl(Chk0(grdPagos.TextMatrix(i, 1)))
        End If
    Next
    If mValorCta > 0 Then
        'Call ImprimirPagare(CStr(mValorCta))
    End If
End Sub

Private Sub cmdAceptoTarjeta_Click()
    If cboPlan.ListIndex = -1 Then
        MsgBox "Falta Ingresar el Plan", vbExclamation, TIT_MSGBOX
        cboPlan.SetFocus
        Exit Sub
    End If
    If txtLote.Text = "" Then
        MsgBox "Falta Ingresar el Lote", vbExclamation, TIT_MSGBOX
        txtLote.SetFocus
        Exit Sub
    End If
    If txtCupon.Text = "" Then
        MsgBox "Falta Ingresar el Cupon", vbExclamation, TIT_MSGBOX
        txtCupon.SetFocus
        Exit Sub
    End If
    If txtTar_Autorizacion.Text = "" Then
        MsgBox "Falta Ingresar la Autorizacion", vbExclamation, TIT_MSGBOX
        txtTar_Autorizacion.SetFocus
        Exit Sub
    End If
        
    fraTarjeta.Visible = False
    'cboFormaPago.ListIndex = 0
    txtImportePago.SetFocus
End Sub

Private Sub cmdAltaPlan_Click()
    mOrigen = False
    ABMTarjetaPlan.vMode = 1
    ABMTarjetaPlan.Show vbModal
    sql = "SELECT PLA_CODIGO, PLA_DESCRI FROM TARJETA_PLAN WHERE TAR_CODIGO = " & XN(cboTarjeta.ItemData(cboTarjeta.ListIndex))
    sql = sql & " ORDER BY PLA_DESCRI"
    Call CargoComboBoxItemData(cboPlan, sql)
    cboPlan.ListIndex = 0
End Sub

Private Sub cmdAltaTarjeta_Click()
    mOrigen = False
    ABMTarjeta.vMode = 1
    ABMTarjeta.Show vbModal
    cSQL = "SELECT TAR_CODIGO, TAR_DESCRI FROM TARJETA ORDER BY TAR_DESCRI"
    Call CargoComboBoxItemData(cboTarjeta, cSQL)
    cboTarjeta.ListIndex = 0
End Sub

Private Sub cmdBorroFila_Click()
    If grdPagos.Rows <= 2 Then
        grdPagos.Rows = 1
    Else
        grdPagos.RemoveItem (grdPagos.row)
    End If
    Dim mTotalPagos As Double
    mTotalPagos = 0
    For i = 1 To grdPagos.Rows - 1
        mTotalPagos = CDbl(mTotalPagos) + CDbl(grdPagos.TextMatrix(i, 1))
    Next
    txtTotalPagos.Text = Format(CDbl(txtTotal.Text) - mTotalPagos, "0.00")
    cboFormaPago.SetFocus
End Sub

Private Sub cmdBuscaCli_Click()
    BuscarClientes "txtBuscaCliente", "CODIGO"
    txtBuscarCliDescri.SetFocus
End Sub

Private Sub cmdbuscaComp_Click()
    txtcodCli.Text = ""
    BuscarClientes "txtcodCli", "CODIGO"
    txtRazSoc.SetFocus
End Sub

Private Sub CmdBuscAprox_Click()
    GrdModulos.Rows = 1
    GrdModulos.HighLight = flexHighlightNever
    lblEstado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
    
    sql = "SELECT FC.*,"
    sql = sql & " C.CLI_RAZSOC,C.CLI_CODIGO,TC.TCO_ABREVIA"
    sql = sql & " FROM FACTURA_CLIENTE FC,CLIENTE C,"
    sql = sql & " TIPO_COMPROBANTE TC, FORMA_PAGO FP"
    sql = sql & " WHERE"
    sql = sql & " FC.TCO_CODIGO=TC.TCO_CODIGO"
    sql = sql & " AND FC.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & " AND FP.FPG_CODIGO = FC.FPG_CODIGO"
    If txtBuscaCliente.Text <> "" Then
        sql = sql & " AND FC.CLI_CODIGO=" & XN(txtBuscaCliente.Text)
    End If
    If FechaDesde.Value <> "" Then
        sql = sql & " AND FC.FCL_FECHA>=" & XDQ(FechaDesde.Value)
    End If
    If FechaHasta.Value <> "" Then
        sql = sql & " AND FC.FCL_FECHA<=" & XDQ(FechaHasta.Value)
    End If
    If cboFactura1.List(cboFactura1.ListIndex) <> "(Todas)" Then
        sql = sql & " AND FC.TCO_CODIGO=" & cboFactura1.ItemData(cboFactura1.ListIndex)
    End If
    
    If cboTurnosB.List(cboTurnosB.ListIndex) <> "" Then
        sql = sql & " AND FC.TUR_CODIGO=" & cboTurnosB.ItemData(cboTurnosB.ListIndex)
    End If
    If txtBuscaNum.Text <> "" Then
        sql = sql & " AND FC.FCL_SUCURSAL=" & txtBuscaSuc.Text
        sql = sql & " AND FC.FCL_NUMERO=" & txtBuscaNum.Text
    End If
    
    sql = sql & " ORDER BY FC.FCL_FECHA,FC.FCL_NUMERO"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            GrdModulos.AddItem rec!TCO_ABREVIA & Chr(9) & Format(rec!FCL_SUCURSAL, "0000") & "-" & Format(rec!FCL_NUMERO, "00000000") & Chr(9) & rec!FCL_FECHA _
                            & Chr(9) & rec!CLI_RAZSOC & Chr(9) & rec!EST_CODIGO _
                            & Chr(9) & rec!FCL_IVA & Chr(9) & rec!FCL_OBSERVACION _
                            & Chr(9) & rec!TCO_CODIGO & Chr(9) & rec!FPG_CODIGO _
                            & Chr(9) & rec!CLI_CODIGO & Chr(9) & Chk0(rec!FCL_TOTAL) _
                            & Chr(9) & Chk0(rec!FCL_IMPIVA) & Chr(9) & Chk0(rec!VEN_CODIGO) _
                            & Chr(9) & Chk0(rec!TUR_CODIGO)
            rec.MoveNext
        Loop
        GrdModulos.HighLight = flexHighlightAlways
        GrdModulos.SetFocus
        GrdModulos.Col = 0
        GrdModulos.row = 1
    Else
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        MsgBox "No se encontraron datos...", vbExclamation, TIT_MSGBOX
    End If
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    rec.Close
End Sub

Private Sub cmdCerrarPagos_Click()
    fraPagos.Visible = False
    cboFormaPago.ListIndex = 0
End Sub

Private Sub cmdCerrarTarjeta_Click()
    cboFormaPago.ListIndex = 0
    fraTarjeta.Visible = False
    cboFormaPago.SetFocus
End Sub
Private Function grabar_factura()
    sql = "SELECT * FROM FACTURA_CLIENTE"
    sql = sql & " WHERE TCO_CODIGO=" & cboFactura.ItemData(cboFactura.ListIndex)
    sql = sql & " AND FCL_NUMERO = " & XN(txtNroFactura.Text)
    sql = sql & " AND FCL_SUCURSAL=" & XN(txtNroSucursal.Text)
    If rec.State = 1 Then rec.Close
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic

    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando..."

    If rec.EOF = True Then
        'NUEVA FACTURA
        sql = "INSERT INTO FACTURA_CLIENTE"
        sql = sql & " (TCO_CODIGO,FCL_NUMERO,FCL_SUCURSAL,FCL_FECHA,"
        sql = sql & " FCL_IVA,FCL_IMPIVA,FPG_CODIGO,FCL_OBSERVACION,VEN_CODIGO,"
        sql = sql & " FCL_SUBTOTAL,FCL_TOTAL,FCL_SALDO,EST_CODIGO,"
        sql = sql & " FCL_NUMEROTXT,FCL_SUCURSALTXT,CLI_CODIGO,FCL_IMPINT,TUR_CODIGO,FCL_HORA,FCL_TASAVIAL,FCL_TOTALACT)"
        sql = sql & " VALUES ("
        sql = sql & cboFactura.ItemData(cboFactura.ListIndex) & ","
        sql = sql & XN(txtNroFactura.Text) & ","
        sql = sql & XN(txtNroSucursal.Text) & ","
        sql = sql & XDQ(FechaFactura.Value) & ","
        sql = sql & XN(txtPorcentajeIva.Text) & ","
        sql = sql & XN(txtImporteIvaB.Text) & "," 'uso este campo no visible para guardar la info de las FAC B
        sql = sql & cboFormaPago.ItemData(cboFormaPago.ListIndex) & ","
        sql = sql & XS(txtObservaciones) & ","
        sql = sql & cboVendedor.ItemData(cboVendedor.ListIndex) & ","
        sql = sql & XN(txtsubtotal1B.Text) & "," 'uso este campo no visible para guardar la info de las FAC B
        sql = sql & XN(txtTotal.Text) & ","
        sql = sql & XN(txtTotal.Text) & "," 'SALDO FACTURA
        sql = sql & "3," 'ESTADO DEFINITIVO
        sql = sql & XS(Format(txtNroFactura.Text, "00000000")) & ","
        sql = sql & XS(Format(txtNroSucursal.Text, "0000")) & ","
        sql = sql & XN(txtcodCli.Text) & "," 'CLIENTE
        sql = sql & XN(txtimpuestoB.Text) & "," 'uso este campo no visible para guardar la info de las FAC B
        sql = sql & cboTurno.ItemData(cboTurno.ListIndex) & ","
        sql = sql & XS(Format(Time, "hh:mm")) & ","
        sql = sql & XN(txttasavial.Text) & ","
        sql = sql & XN(txtTotal.Text) & ")"
        
        'Format(Valor, "mm/dd/yyyy") & "#"
        DBConn.Execute sql
           
        For i = 1 To grdGrilla.Rows - 1
            If grdGrilla.TextMatrix(i, 0) <> "" Then
                sql = "INSERT INTO DETALLE_FACTURA_CLIENTE"
                sql = sql & " (TCO_CODIGO, FCL_NUMERO, FCL_SUCURSAL, DFC_NROITEM,"
                sql = sql & " PTO_CODIGO, DFC_CONCEPTO, DFC_CANTIDAD, DFC_PRECIO, DFC_IVA,"
                sql = sql & " DFC_IMP, DFC_MONIMP, DFC_MONIVA,DFC_TASAVIAL,DFC_TOTALTVIAL)"
                ' guardar el imp, el monto del imp y el monto del iva
                
                sql = sql & " VALUES ("
                sql = sql & cboFactura.ItemData(cboFactura.ListIndex) & ","
                sql = sql & XN(txtNroFactura.Text) & ","
                sql = sql & XN(txtNroSucursal.Text) & ","
                sql = sql & i & "," 'PONER EL NRO ITEM
                sql = sql & XN(grdGrilla.TextMatrix(i, 0)) & ","
                sql = sql & XS(grdGrilla.TextMatrix(i, 1)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 2)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 3)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 6)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 7)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 8)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 9)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 11)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 12)) & ")"
                
                DBConn.Execute sql
                
                sql = "SELECT DST_STKFIS,DST_STKCON"
                sql = sql & " FROM STOCK"
                sql = sql & " WHERE STK_CODIGO = " & XN(Sucursal)
                sql = sql & " AND PTO_CODIGO = " & XN(grdGrilla.TextMatrix(i, 0))
                Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
                If Rec1.EOF = False Then
                    sql = "UPDATE STOCK SET"
                    sql = sql & " DST_STKFIS = DST_STKFIS - " & XN(grdGrilla.TextMatrix(i, 2))
                    sql = sql & " WHERE STK_CODIGO = " & XN(Sucursal)
                    sql = sql & " AND PTO_CODIGO = " & XN(grdGrilla.TextMatrix(i, 0))
                    DBConn.Execute sql
                End If
                Rec1.Close
            End If
        Next
        
        For i = 1 To grdPagos.Rows - 1
            sql = "insert into FACTURA_PAGOS (FCL_SUCURSAL,FCL_NUMERO,TCO_CODIGO,FPG_CODIGO,pag_importe,TAR_CODIGO,"
            sql = sql & " TAR_PLAN,TAR_CUPON, TAR_LOTE,TAR_AUTORIZACION,TOTDOLAR, COTIZACION, sen_sucursal, sen_tipo,"
            sql = sql & " sen_nro, PAG_SECUENCIA, FPG_SALDO,CHE_BANCO,CHE_NUMERO)"
            sql = sql & " values ("
            sql = sql & XN(txtNroSucursal.Text) & ", " & XN(txtNroFactura.Text) & ", " & XN(cboFactura.ItemData(cboFactura.ListIndex)) & ", "
            sql = sql & XN(grdPagos.TextMatrix(i, 2)) & ","
            sql = sql & XN(grdPagos.TextMatrix(i, 1))
            sql = sql & "," & XN(grdPagos.TextMatrix(i, 3))
            sql = sql & "," & XN(grdPagos.TextMatrix(i, 5))
            sql = sql & "," & XN(grdPagos.TextMatrix(i, 7))
            sql = sql & "," & XN(grdPagos.TextMatrix(i, 8))
            sql = sql & "," & XN(grdPagos.TextMatrix(i, 9))
            sql = sql & "," & XN(grdPagos.TextMatrix(i, 10))
            sql = sql & "," & XN(grdPagos.TextMatrix(i, 11))
            sql = sql & "," & XN(grdPagos.TextMatrix(i, 12))
            sql = sql & "," & XN(grdPagos.TextMatrix(i, 13))
            sql = sql & "," & XS(grdPagos.TextMatrix(i, 14)) & "," & i & ","
            If grdPagos.TextMatrix(i, 2) <> "2" Then
                sql = sql & "0"
            Else
                sql = sql & XN(grdPagos.TextMatrix(i, 1))
                mSaldo = mSaldo + CDbl(Chk0(grdPagos.TextMatrix(i, 1)))
            End If
            If grdPagos.TextMatrix(i, 2) = "5" Then
                sql = sql & "," & XS(grdPagos.TextMatrix(i, 4))
                sql = sql & "," & XS(grdPagos.TextMatrix(i, 6)) & ")"
            Else
                sql = sql & "," & XS("")
                sql = sql & "," & XS("") & ")"
            End If
            
            DBConn.Execute sql
        Next
        
'            sql = "UPDATE FACTURA_CLIENTE"
'            sql = sql & " SET FCL_SALDO=" & XN(CStr(mSaldo))
'            sql = sql & " WHERE"
'            sql = sql & " TCO_CODIGO=" & cboFactura.ItemData(cboFactura.ListIndex)
'            sql = sql & " AND FCL_NUMERO=" & XN(txtNroFactura.Text)
'            sql = sql & " AND FCL_SUCURSAL=" & XN(txtNroSucursal.Text)
'            DBConn.Execute sql
        
        'ACTUALIZO LA TABLA PARAMENTROS Y LE SUMO UNO A LA FACTURA QUE CORRESPONDE
'        Select Case cboFactura.ItemData(cboFactura.ListIndex)
'            Case 1
'                sql = "UPDATE PARAMETROS SET FACTURA_C=" & XN(txtNroFactura.Text)
'            Case 2
'                sql = "UPDATE PARAMETROS SET FACTURA_C=" & XN(txtNroFactura.Text)
'        End Select
'        DBConn.Execute sql
    End If
    rec.Close
End Function
Private Function seteo_iva()
    If txtImporteIva.Text = "" Then
        txtPorcentajeIva_LostFocus
    End If
    If mRespo.Text = "" Then
        sql = "SELECT I.IVA_LETRA"
        sql = sql & " FROM CLIENTE C, CONDICION_IVA I"
        sql = sql & " WHERE I.IVA_CODIGO = C.IVA_CODIGO"
        sql = sql & " AND C.CLI_CODIGO =" & XN(txtcodCli.Text)
        'If rec.State = 1 Then rec.Close
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            mRespo.Text = ChkNull(rec!IVA_LETRA)
        End If
        'If rec.State = 1 Then rec.Close
        rec.Close
    End If
End Function


Private Sub cmdcrerrarcheque_Click()
    cboFormaPago.ListIndex = 0
    fracheque.Visible = False
    cboFormaPago.SetFocus
End Sub

Private Sub cmdGrabar_Click()
    Dim VStockPendiente As String
    Dim mSaldo As Double
    mSaldo = 0
    If MsgBox("Confirma la impresion de la Factura?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        
        FechaFactura = Date
        
        seteo_iva
        
        If ValidarFactura = False Then Exit Sub
               
        If rec.State = 1 Then rec.Close
        
        'On Error GoTo HayErrorFactura
        DBConn.BeginTrans
        
        'Funcion que graba la factura en la BD
        grabar_factura
        
        DBConn.CommitTrans
             
        If VerificoSiGrabo = False Then
            If rec.State = 1 Then rec.Close
            Set frmFacturaCliente = Nothing
            Unload Me
        End If
        If rec.State = 1 Then rec.Close
        
        Do While VerificoSiGrabo = False
            grabar_factura
        Loop
       
        'IMPRIME COMPROBANTE FISCAL
        'If txtFiscal.Text = "F" And mImprime = "S" Then
        mRespuestaFiscal = True
                        
        If FISCAL = "TMT900FA" Then
            ImprimoFiscalEpsondll 2
        Else
            Imprimo_Fiscal
            errores_impresion
            ActualizoTotalesFiscales
        End If
        
        CmdNuevo_Click
    End If
    Exit Sub
    
HayErrorFactura:
    mMeDioError = True
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    If rec.State = 1 Then rec.Close
    If Rec1.State = 1 Then Rec1.Close
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub
Private Function errores_impresion()
    If mRespuestaFiscal = True Then
        'DBConn.CommitTrans
        mMeDioError = False
    Else
        mMeDioError = True
        Dim A As String
        Dim B As String
        
        A = pf.FiscalStatus
        B = pf.PrinterStatus
        
        MsgBox "Error al Imprimir", vbCritical, TIT_MSGBOX
        If A = "b220" Then
            'documento abierto y error
            MsgBox "Error en la Impresión Fiscal." & Chr(13) & "Apaguela y Préndala." & Chr(13) & "El Documento será ANULADO." & Chr(13) & Err.Description, 16, AppName
        
            sql = "UPDATE FACTURA_CLIENTE SET EST_CODIGO=2"
            sql = sql & " WHERE FCL_NUMERO=" & XN(txtNroFactura.Text)
            sql = sql & " AND TCO_CODIGO=" & XN(cboFactura.ItemData(cboFactura.ListIndex))
            sql = sql & " AND FCL_SUCURSAL = " & XN(txtNroSucursal.Text)
            DBConn.Execute sql
        
            'DBConn.CommitTrans
        ElseIf A = "8210" Then
            MsgBox "Error al Abrir el Ticket. Controle y vuelva a Intentar" & Chr(13) & Err.Description, 16, AppName
            
            If VerificoSiGrabo = True Then
                BorrarFactura
            End If
            'DBConn.RollbackTrans
            
        ElseIf A = "" Then
            MsgBox "Error de Comunicación con la Impresora Fiscal." & Chr(13) & "Controle las Conexiones." & Chr(13) & "Apáguela y Préndala" & Chr(13) & "El Documento será ANULADO." & Chr(13) & Err.Description, 16, AppName
            
            sql = "UPDATE FACTURA_CLIENTE SET EST_CODIGO=2"
            sql = sql & " WHERE FCL_NUMERO=" & XN(txtNroFactura.Text)
            sql = sql & " AND TCO_CODIGO=" & XN(cboFactura.ItemData(cboFactura.ListIndex))
            sql = sql & " AND FCL_SUCURSAL = " & XN(txtNroSucursal.Text)
            DBConn.Execute sql
        
            'DBConn.CommitTrans
        Else
            If VerificoSiGrabo = True Then
                BorrarFactura
            End If
            'DBConn.RollbackTrans
        End If
        'MsgBox "Error en la Impresión Fiscal" & Chr(13) & Err.Description, 16, APPNAME
        Screen.MousePointer = vbNormal
    End If
    
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
End Function
Private Function VerificoSiGrabo() As Boolean
    VerificoSiGrabo = True
    sql = "SELECT * FROM FACTURA_CLIENTE"
    sql = sql & " WHERE TCO_CODIGO=" & cboFactura.ItemData(cboFactura.ListIndex)
    sql = sql & " AND FCL_NUMERO = " & XN(txtNroFactura.Text)
    sql = sql & " AND FCL_SUCURSAL=" & XN(txtNroSucursal.Text)
    If rec.State = 1 Then rec.Close
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = True Then 'NO SE GRABO
        Screen.MousePointer = vbNormal
        lblEstado.Caption = ""
        'MsgBox "Error al Grabar la Factura...", vbCritical, TIT_MSGBOX
        If rec.State = 1 Then rec.Close
        VerificoSiGrabo = False
    Else 'SE GRABO
        VerificoSiGrabo = True
    End If
    If rec.State = 1 Then rec.Close
End Function
Private Function ActualizoTotalesFiscales()
    Dim vNetoFiscal As String
    Dim vIVAFiscal As String
    Dim vImpIntFiscal As String
    Dim vTotalFiscal As String
    
    If txttasavial.Text <> "0,00" Then 'es combustible
        'primero hago los calculos
        vNetoFiscal = CDbl(Chk0(txtNetoFiscal.Text)) - CDbl(txttasavial)
        vIVAFiscal = vNetoFiscal * mIVA_1 / 100
        vImpIntFiscal = CDbl(Chk0(txtIvaFiscal.Text))
        vTotalFiscal = CDbl(vNetoFiscal) + CDbl(vIVAFiscal) + CDbl(vImpIntFiscal) + CDbl(txttasavial)
        
        vNetoFiscal = Format(vNetoFiscal, "#,##0.00")
        vIVAFiscal = Format(vIVAFiscal, "#,##0.00")
        vImpIntFiscal = Format(vImpIntFiscal, "#,##0.00")
        vTotalFiscal = Format(vTotalFiscal, "#,##0.00")
        
      
        
        sql = "UPDATE FACTURA_CLIENTE SET"
        sql = sql & " FCL_SUBTOTAL = " & XN(vNetoFiscal)
        sql = sql & ",FCL_IMPIVA = " & XN(vIVAFiscal)
        sql = sql & ",FCL_IMPINT = " & XN(vImpIntFiscal)
        sql = sql & ",FCL_TOTAL = " & XN(vTotalFiscal)
        sql = sql & ",FCL_SALDO = " & XN(vTotalFiscal)
        sql = sql & " WHERE TCO_CODIGO=" & cboFactura.ItemData(cboFactura.ListIndex)
        sql = sql & " AND FCL_NUMERO = " & XN(txtNroFactura.Text)
        sql = sql & " AND FCL_SUCURSAL=" & XN(txtNroSucursal.Text)
        DBConn.Execute sql
    End If
End Function

Private Sub BorrarFactura()
    sql = "DELETE FROM FACTURA_PAGOS"
    sql = sql & " WHERE TCO_CODIGO=" & cboFactura.ItemData(cboFactura.ListIndex)
    sql = sql & " AND FCL_NUMERO = " & XN(txtNroFactura.Text)
    sql = sql & " AND FCL_SUCURSAL=" & XN(CInt(txtNroSucursal.Text))
    DBConn.Execute sql
    
    sql = "DELETE FROM DETALLE_FACTURA_CLIENTE"
    sql = sql & " WHERE TCO_CODIGO=" & cboFactura.ItemData(cboFactura.ListIndex)
    sql = sql & " AND FCL_NUMERO = " & XN(txtNroFactura.Text)
    sql = sql & " AND FCL_SUCURSAL=" & XN(CInt(txtNroSucursal.Text))
    DBConn.Execute sql
    
    sql = "DELETE FROM FACTURA_CLIENTE"
    sql = sql & " WHERE TCO_CODIGO=" & cboFactura.ItemData(cboFactura.ListIndex)
    sql = sql & " AND FCL_NUMERO = " & XN(txtNroFactura.Text)
    sql = sql & " AND FCL_SUCURSAL=" & XN(CInt(txtNroSucursal.Text))
    DBConn.Execute sql
End Sub
Private Function ValidarFactura() As Boolean
'    If txtNroFactura.Text = "" Then
'        MsgBox "Falta el Número de la Factura", vbExclamation, TIT_MSGBOX
'        ValidarFactura = False
'        Exit Function
'    End If
    If FechaFactura.Value = "" Then
        MsgBox "La Fecha de la Factura es requerida", vbExclamation, TIT_MSGBOX
        FechaFactura.SetFocus
        ValidarFactura = False
        Exit Function
    End If
    
    If cboVendedor.List(cboVendedor.ListIndex) = "" Then
        MsgBox "El Vendedor es Requerido", vbExclamation, TIT_MSGBOX
        cboVendedor.SetFocus
        ValidarFactura = False
        Exit Function
    End If
    
    If txtSubtotal.Text = "" Then
        MsgBox "El Sub Total de la Factura no puede ser Nulo", vbCritical, TIT_MSGBOX
        grdGrilla.Col = 0
        grdGrilla.row = 2
        grdGrilla.SetFocus
        ValidarFactura = False
        Exit Function
    End If
    If txtTotal.Text = "" Then
        MsgBox "El Total de la Factura no puede ser Nulo", vbCritical, TIT_MSGBOX
        grdGrilla.Col = 0
        grdGrilla.row = 2
        grdGrilla.SetFocus
        ValidarFactura = False
        Exit Function
    End If
    If mRespo.Text = "" Then
        MsgBox "Error", vbCritical, TIT_MSGBOX
        ValidarFactura = False
        Exit Function
    End If
    If grdPagos.Rows = 1 Then
        'If CDbl(Chk0(txtTotal.Text)) > 0 Then
            MsgBox "Debe indicar la Forma de Pago para poder grabar el movimiento !", vbInformation, TIT_MSGBOX
            fraPagos.Top = 930
            fraPagos.Left = 3345
            fraPagos.Visible = True
            
            Dim mTotalPagos As Double
            mTotalPagos = 0
            For i = 1 To grdPagos.Rows - 1
              mTotalPagos = mTotalPagos + CDbl(grdPagos.TextMatrix(i, 1))
            Next
            txtTotalPagos.Text = Format(CDbl(txtTotal.Text) - mTotalPagos, "0.00")
            
            txtGrabar.Text = "S"
            
            cboFormaPago.SetFocus
            ValidarFactura = False
            Exit Function
        'End If
    End If
    
    'If frmFactu1.txtFiscal.Text = "F" Then
    For i = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(i, 2) <> "" Then
            If CInt(grdGrilla.TextMatrix(i, 2)) < 0 Then
                MsgBox "UD. ESTA INTENTANDO EMITIR UN COMPROBANTE FISCAL CON CANTIDAD NEGATIVA." & Chr(13) & "ESTO PRODUCIRÁ UN ERROR EN EL CONTROLADOR FISCAL." & Chr(13) & "CORRIJA LA CANTIDAD O UTILICE LAS OPCIONES DE CARGA DE COMPROBANTES MANUALES !!!", vbCritical, TIT_MSGBOX
                ValidarFactura = False
                Exit For
            End If
        End If
    Next i
    If txtNroFactura.Text = "" Then
        MsgBox "Falta el Número de la Factura", vbExclamation, TIT_MSGBOX
        ValidarFactura = False
        Exit Function
    End If
    
    'End If
    ValidarFactura = True
End Function

Private Sub cmdImprimir_Click()
    'If MsgBox("¿Confirma Impresión Factura?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    'Set_Impresora
    'ImprimirFactura
    
    'BUSCA EL NUMERO DE FACTURA
    'If txtFiscal.Text = "F" Then
    '    If mImprime = "S" Then
'            Select Case cboFactura.ItemData(cboFactura.ListIndex)
'                Case 1 'FACTURAS A
'                    pf.Status ("A")
'                    txtNroFactura.Text = Val(pf.AnswerField_7) + 1
'                Case 2 'FACTURA B
'                    pf.Status ("A")
'                    txtNroFactura.Text = Val(pf.AnswerField_5) + 1
'                Case 3 'FACTURA C
'                Case 10000 'PARA TIKET
'                    pf.Status ("A")
'                    txtNroFactura.Text = Val(pf.AnswerField_4) + 1
'            End Select
            
        'Imprimo_Fiscal
        
'    If txtFiscal.Text = "F" And mImprime = "S" Then
'        mRespuestaFiscal = True
'        Imprimo_Fiscal
'    Else
'        mRespuestaFiscal = True
'    End If
'
'    If mRespuestaFiscal = True Then
'        DBConn.CommitTrans
'    Else
'
'        A = pf.FiscalStatus
'        B = pf.PrinterStatus
'
'        If A = "b220" Then
'            'documento abierto y error
'            MsgBox "Error en la Impresión Fiscal." & Chr(13) & "Apaguela y Préndala." & Chr(13) & "El Documento será ANULADO." & Chr(13) & Err.Description, 16, AppName
'
'            sql = "UPDATE VENTGRAL SET BAJA='S' WHERE NROFAC=" + XN(mNroFactu.Text) + " AND TIPO_FAC=" + XS(Ltipo_fac.Caption) + " AND SUC_CODIGO = " & XN(txtSucursal.Text)
'            DBConn.Execute sql
'
'            sql = "UPDATE VENTITEM SET BAJA='S' WHERE NROFAC = " + XN(mNroFactu.Text) + " AND TIPO_FAC=" + XS(Ltipo_fac.Caption) + " AND SUC_CODIGO = " & XN(txtSucursal.Text)
'            DBConn.Execute sql
'
'            DBConn.CommitTrans
'        End If
'
'        If A = "8210" Then
'            MsgBox "Error al Abrir el Ticket. Controle y vuelva a Intentar" & Chr(13) & Err.Description, 16, AppName
'            DBConn.RollbackTrans
'        End If
'
'        If A = "" Then
'            MsgBox "Error de Comunicación con la Impresora Fiscal." & Chr(13) & "Controle las Conexiones." & Chr(13) & "Apáguela y Préndala" & Chr(13) & "El Documento será ANULADO." & Chr(13) & Err.Description, 16, AppName
'
'            sql = "UPDATE VENTGRAL SET BAJA='S' WHERE NROFAC=" + XN(mNroFactu.Text) + " AND TIPO_FAC=" + XS(Ltipo_fac.Caption) + " AND SUC_CODIGO = " & XN(txtSucursal.Text)
'            DBConn.Execute sql
'
'            sql = "UPDATE VENTITEM SET BAJA='S' WHERE NROFAC = " + XN(mNroFactu.Text) + " AND TIPO_FAC=" + XS(Ltipo_fac.Caption) + " AND SUC_CODIGO = " & XN(txtSucursal.Text)
'            DBConn.Execute sql
'
'            DBConn.CommitTrans
'        End If
'
'
'        'MsgBox "Error en la Impresión Fiscal" & Chr(13) & Err.Description, 16, APPNAME
'
'        MP Normal
'    End If
'    '    Else
'         '   Select Case Ltipo_fac.Caption
'         '      Case "A"
'         '           cSQL$ = "SELECT NROFAC FROM PARAM"
'         '           If snp.State = 1 Then snp.Close
'         '           snp.Open cSQL$, DBConn, adOpenStatic, adLockOptimistic
'         '           mNroFactu.Text = snp("NROFAC")
'         '      Case "B"
'         '           cSQL$ = "SELECT NROFACB FROM PARAM"
'         '           If snp.State = 1 Then snp.Close
'         '           snp.Open cSQL$, DBConn, adOpenStatic, adLockOptimistic
'         '           mNroFactu.Text = snp("NROFACB")
'         '   End Select
'        'End If
'    'End If
End Sub

Private Sub Imprimo_Fiscal()

    Dim mContCanti As Integer
    Dim mContPrecio As Integer
    Dim mContInternos As Double
    Dim mVendedor As String
    Dim mIngBrutos As String
    Dim mItc As Double
    Dim mPneto As Double
    Dim mTasa As Double
    Dim mTVialDbl As Double
    
    Dim mTotaLl As Double
    Dim ITOTAL As String

    Dim iCanti As String
    Dim iPrecio As String
    Dim iImpInt As String
    Dim mIvaFE As String
    Dim mTasaVial As String
    
'    mContCanti = 1000
'    mContPrecio = 100
    mIVA_1 = BuscoIva
    mIVA_2 = BuscoIva_2
    
    'Lo cambio para el redondeo de combustibles
    mContCanti = 100
    mContPrecio = 1000
    mContInternos = 100000000
    'mContInternos = 1000
    
    mVendedor = "Playero: " & cboVendedor.List(cboVendedor.ListIndex)
    mVendedor = Mid(mVendedor, 1, 20)
    
    mIngBrutos = "Ing Brutos: " & IIf(txtIngBrutos.Text = "" Or txtIngBrutos.Text = "0", "NO POSEE", txtIngBrutos.Text)
    mIngBrutos = Mid(mIngBrutos, 1, 25)
    
    
    If InStr(1, mDireccion.Text, "Á") > 0 Or InStr(1, mDireccion.Text, "É") > 0 Or InStr(1, mDireccion.Text, "Í") > 0 Or InStr(1, mDireccion.Text, "Ó") > 0 Or InStr(1, mDireccion.Text, "Ú") > 0 Or InStr(1, mDireccion.Text, "Ñ") > 0 Or InStr(1, mDireccion.Text, "á") > 0 Or InStr(1, mDireccion.Text, "é") > 0 Or InStr(1, mDireccion.Text, "í") > 0 Or InStr(1, mDireccion.Text, "ó") > 0 Or InStr(1, mDireccion.Text, "ú") > 0 Or InStr(1, mDireccion.Text, "ñ") > 0 Or InStr(1, mDireccion.Text, "ü") > 0 Or InStr(1, mDireccion.Text, "Ü") > 0 Or InStr(1, mDireccion.Text, "º") > 0 Then
        mDireccion.Text = "DIRECCION"
    End If
    
    If InStr(1, mLocalidad.Text, "Á") > 0 Or InStr(1, mLocalidad.Text, "É") > 0 Or InStr(1, mLocalidad.Text, "Í") > 0 Or InStr(1, mLocalidad.Text, "Ó") > 0 Or InStr(1, mLocalidad.Text, "Ú") > 0 Or InStr(1, mLocalidad.Text, "Ñ") > 0 Or InStr(1, mLocalidad.Text, "á") > 0 Or InStr(1, mLocalidad.Text, "é") > 0 Or InStr(1, mLocalidad.Text, "í") > 0 Or InStr(1, mLocalidad.Text, "ó") > 0 Or InStr(1, mLocalidad.Text, "ú") > 0 Or InStr(1, mLocalidad.Text, "ñ") > 0 Or InStr(1, mLocalidad.Text, "ü") > 0 Or InStr(1, mLocalidad.Text, "Ü") > 0 Or InStr(1, mLocalidad.Text, "º") > 0 Then
        mLocalidad.Text = "LOCALIDAD"
    End If
    If InStr(1, mProvincia.Text, "Á") > 0 Or InStr(1, mProvincia.Text, "É") > 0 Or InStr(1, mProvincia.Text, "Í") > 0 Or InStr(1, mProvincia.Text, "Ó") > 0 Or InStr(1, mProvincia.Text, "Ú") > 0 Or InStr(1, mProvincia.Text, "Ñ") > 0 Or InStr(1, mProvincia.Text, "á") > 0 Or InStr(1, mProvincia.Text, "é") > 0 Or InStr(1, mProvincia.Text, "í") > 0 Or InStr(1, mProvincia.Text, "ó") > 0 Or InStr(1, mProvincia.Text, "ú") > 0 Or InStr(1, mProvincia.Text, "ñ") > 0 Or InStr(1, mProvincia.Text, "ü") > 0 Or InStr(1, mProvincia.Text, "Ü") > 0 Or InStr(1, mProvincia.Text, "º") > 0 Then
        mProvincia.Text = "PROVINCIA"
    End If
    
    If InStr(1, txtRazSoc.Text, "Á") > 0 Or InStr(1, txtRazSoc.Text, "É") > 0 Or InStr(1, txtRazSoc.Text, "Í") > 0 Or InStr(1, txtRazSoc.Text, "Ó") > 0 Or InStr(1, txtRazSoc.Text, "Ú") > 0 Or InStr(1, txtRazSoc.Text, "Ñ") > 0 Or InStr(1, txtRazSoc.Text, "á") > 0 Or InStr(1, txtRazSoc.Text, "é") > 0 Or InStr(1, txtRazSoc.Text, "í") > 0 Or InStr(1, txtRazSoc.Text, "ó") > 0 Or InStr(1, txtRazSoc.Text, "ú") > 0 Or InStr(1, txtRazSoc.Text, "ñ") > 0 Or InStr(1, txtRazSoc.Text, "ü") > 0 Or InStr(1, txtRazSoc.Text, "Ü") > 0 Or InStr(1, txtRazSoc.Text, "º") > 0 Then
        txtRazSoc.Text = SacoAcento(Trim(txtRazSoc.Text))
    End If
    
    If cboFactura.ItemData(cboFactura.ListIndex) = 1 Then 'factura A
        'mRespuestaFiscal = pf.OpenInvoice("T", "C", "A", "1", "P", "12", "I", mRespo.Text, txtRazSoc.Text, "", "CUIT", txtCuit.Text, "N", Trim(mDireccion.Text), Trim(mLocalidad.Text), Trim(mProvincia.Text), "", "", "C")
        
        'original
        'EMULADOR mRespuestaFiscal = pf.OpenInvoice("T", "C", "A", "1", "P", "12", "I", mRespo.Text, txtRazSoc.Text, "A", "CUIT", txtCuit.Text, "N", Trim(mIngBrutos), Trim(mVendedor), "X", "X", "B", "G")
        mRespuestaFiscal = pf.OpenInvoice("T", "C", "A", "1", "P", "12", "I", mRespo.Text, txtRazSoc.Text, "", "CUIT", txtCuit.Text, "N", Trim(mIngBrutos), Trim(mVendedor), "", Trim(txtNroFactura), "", "G")
        'talampaya
        'mRespuestaFiscal = pf.OpenInvoice("T", "C", "A", "1", "P", "12", "I", mRespo.Text, txtRazSoc.Text, "", "CUIT", txtCuit.Text, "N", "", Trim(mVendedor), "", "", "", "C")
        
        If mRespuestaFiscal = False Then Exit Sub
    End If
    If cboFactura.ItemData(cboFactura.ListIndex) = 2 Then 'factura B
        If txtCiva.Text = "CONSUMIDOR FINAL" Then
            'ABRO UN TIKET FACTURA B PERO CON TIPO DE DOCUMENTO DNI
            If txtRazSoc.Text = "" Then
                txtRazSoc.Text = "CLIENTE"
            End If
            If txtNRO_DOCUMENTO.Text = "" Then
                If txtCuit.Text = "" Then
                    txtNRO_DOCUMENTO.Text = "11111111"
                Else
                    txtNRO_DOCUMENTO.Text = txtCuit.Text
                End If
            End If
            'mRespuestaFiscal = pf.OpenInvoice("T", "C", "B", "1", "P", "12", "I", mRespo.Text, txtRazSoc.Text, "", "DNI", txtNRO_DOCUMENTO.Text, "N", Trim(mDireccion.Text), Trim(mLocalidad.Text), Trim(mProvincia.Text), "", "", "C")
            'EMULADOR mRespuestaFiscal = pf.OpenInvoice("T", "C", "B", "1", "P", "12", "I", mRespo.Text, txtRazSoc.Text, "A", "DNI", txtNRO_DOCUMENTO.Text, "N", "Z", Trim(mVendedor), "X", "X", "B", "C")
            mRespuestaFiscal = pf.OpenInvoice("T", "C", "B", "1", "P", "12", "I", mRespo.Text, txtRazSoc.Text, "", "DNI", txtNRO_DOCUMENTO.Text, "N", "", Trim(mVendedor), "", "", "", "C")
            If mRespuestaFiscal = False Then Exit Sub
        Else
            'MONOTRIBUTO - ABRO UN TIKET FACTURA B PERO CON TIPO DE DOCUMENTO CUIT
            'mRespuestaFiscal = pf.OpenInvoice("T", "C", "B", "1", "P", "12", "I", mRespo.Text, txtRazSoc.Text, "", "CUIT", txtCuit.Text, "N", Trim(mDireccion.Text), Trim(mLocalidad.Text), Trim(mProvincia.Text), "", "", "C")
            
            'mRespuestaFiscal = pf.OpenInvoice("T", "C", "B", "1", "P", "12", "I", mRespo.Text, txtRazSoc.Text, "", "CUIT", txtCuit.Text, "N", "", Trim(mVendedor), "", "", "", "C")
            'mRespuestaFiscal = pf.OpenInvoice("T", "C", "C", "1", "P", "12", "I", mRespo.Text, txtRazSoc.Text, "", "CUIT", txtCuit.Text, "N", "", Trim(mVendedor), "", "", "", "C")
            mRespuestaFiscal = pf.OpenInvoice("T", "C", "C", "1", "P", "12", "I", mRespo.Text, txtRazSoc.Text, "", "CUIT", txtCuit.Text, "N", Trim(mIngBrutos), Trim(mVendedor), "", Trim(txtNroFactura), "", "G")
            If mRespuestaFiscal = False Then Exit Sub
        End If
    End If
    
    For i = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(i, 0) <> "" Then
            'ACA HAY QUE CALCULAR EL PORCENTAJE DE INCIDENCIA DE LOS IMP INTERNOS EN EL LITRO DE COMB
            
            If cboFactura.ItemData(cboFactura.ListIndex) = 1 Then
                mItc = 0
                mTasa = 0
                If grdGrilla.TextMatrix(i, 0) <> 3 Then  'NAFTA / gnc y demas (el  imp es 0)
                    'NAFTA Y GNC
                    'primero calcular el precio neto del combustible luego el itc y tasa
                    mItc = CDbl(grdGrilla.TextMatrix(i, 2)) * CDbl(grdGrilla.TextMatrix(i, 7))
                    mPneto = CDbl(grdGrilla.TextMatrix(i, 2)) * (CDbl(grdGrilla.TextMatrix(i, 3)) - CDbl(grdGrilla.TextMatrix(i, 11))) - mItc
                    mPneto = mPneto / (1 + (mIVA_1 / 100))
                    mItc = mItc / mPneto
                    mItc = Format(mItc, "0.00000000")
                    
                    mPneto = mPneto / CDbl(grdGrilla.TextMatrix(i, 2))
                    mPneto = Format(mPneto, "0.000")
                Else
                    'GASOIL
                    mItc = CDbl(grdGrilla.TextMatrix(i, 2)) * CDbl(grdGrilla.TextMatrix(i, 7))
                    mPneto = CDbl(grdGrilla.TextMatrix(i, 2)) * (CDbl(grdGrilla.TextMatrix(i, 3)) - CDbl(grdGrilla.TextMatrix(i, 11))) - mItc ' ESTOY RESTANDO LA TASA VIAL COL 11
                    mPneto = mPneto / (1 + (mIVA_2 / 100)) '
                    
                    mTasa = mPneto * ((mIVA_2 - mIVA_1) / 100) ' RESTAR LOS DOS IVAS (40-21)
                    
                    mItc = (mItc + mTasa) / mPneto
                    mItc = Format(mItc, "0.00000000")
                    
                    mPneto = mPneto / CDbl(grdGrilla.TextMatrix(i, 2))
                    mPneto = Format(mPneto, "0.000")
                
                End If
            Else
                'facturas B
                mItc = 0
                mTasa = 0
                If grdGrilla.TextMatrix(i, 0) <> 3 Then  'NAFTA / gnc y demas (el  imp es 0)
                    'NAFTA Y GNC
                    'primero calcular el precio neto del combustible luego el itc y tasa
                    mItc = CDbl(grdGrilla.TextMatrix(i, 2)) * CDbl(grdGrilla.TextMatrix(i, 7))
                    'mPneto = CDbl(grdGrilla.TextMatrix(I, 2)) * CDbl(grdGrilla.TextMatrix(I, 3)) - mItc
                    'mPneto = mPneto / (1 + (mIVA_1 / 100))
                    mPneto = Format(CDbl(grdGrilla.TextMatrix(i, 2)) * (CDbl(grdGrilla.TextMatrix(i, 3)) - CDbl(grdGrilla.TextMatrix(i, 11))), "0.000") ' ESTOY RESTANDO LA TASA VIAL COL 11
                    If mPneto = 0 Then
                        mItc = mPneto
                        mItc = Format(mItc, "0.00000000")
                    
                        mPneto = mPneto
                        mPneto = Format(mPneto, "0.000")
                    Else
                        mItc = mItc / mPneto
                        mItc = Format(mItc, "0.00000000")
                    
                        mPneto = mPneto / CDbl(grdGrilla.TextMatrix(i, 2))
                        mPneto = Format(mPneto, "0.000")
                    End If
                    
                Else
                    'GASOIL
                    mItc = CDbl(grdGrilla.TextMatrix(i, 2)) * CDbl(grdGrilla.TextMatrix(i, 7))
                    'mPneto = CDbl(grdGrilla.TextMatrix(I, 2)) * CDbl(grdGrilla.TextMatrix(I, 3)) - mItc
                    'mPneto = mPneto / (1 + (mIVA_2 / 100)) '
                    mPneto = Format(CDbl(grdGrilla.TextMatrix(i, 2)) * (CDbl(grdGrilla.TextMatrix(i, 3)) - CDbl(grdGrilla.TextMatrix(i, 11))), "0.000") ' ESTOY RESTANDO LA TASA VIAL COL 11
                    
                    mTasa = mPneto * ((mIVA_2 - mIVA_1) / 100) ' RESTAR LOS DOS IVAS (40-21)
                                       
                    
                    mItc = (mItc + mTasa) / mPneto
                    mItc = Format(mItc, "0.00000000")
                    
                    mPneto = mPneto / CDbl(grdGrilla.TextMatrix(i, 2))
                    mPneto = Format(mPneto, "0.000")
                
                End If
                
                
                'mItc = 0
                
                
                grdGrilla.TextMatrix(i, 6) = 0
            End If
            iCanti = Str(Format(CDbl(grdGrilla.TextMatrix(i, 2)), "0.00") * mContCanti)
            iPrecio = Str(mPneto * mContPrecio)
            iImpInt = Str(Format(mItc * CDbl(grdGrilla.TextMatrix(i, 2)), "0.00000000") * mContInternos)
            'iiMpInt = Str(Format(CDbl(iiMpInt), "0.00") * mContInternos)
            mIvaFE = Str(mIVA_1 * 100)
            'mIvaFE = Str(CDbl(grdGrilla.TextMatrix(I, 6)) * 100)
            
            If cboFactura.ItemData(cboFactura.ListIndex) = 10000 Then  '"T" TIKET
                mRespuestaFiscal = pf.SendTicketItem(Trim(ChkNull(grdGrilla.TextMatrix(i, 1))), Trim(iCanti), Trim(iPrecio), Trim(mIvaFE), "M", "0", "0", Trim(iImpInt))
                If mRespuestaFiscal = False Then Exit Sub
            Else
                'mRespuestaFiscal = pf.SendInvoiceItem(Trim(ChkNull(grdGrilla.TextMatrix(i, 1))), Trim(iCanti), Trim(iPrecio), Trim(mIvaFE), "M", "0", "0", ChkNull(grdGrilla.TextMatrix(i, 0)) , "", "", "", Trim(iImpInt))
                mRespuestaFiscal = pf.SendInvoiceItem(Trim(ChkNull(grdGrilla.TextMatrix(i, 1))), Trim(iPrecio), Trim(iCanti), Trim(mIvaFE), "M", "0", "0", "", "", "", "", Trim(iImpInt))
                
                'TICKET
                'mRespuestaFiscal = pf.SendTicketItem(Trim(ChkNull(GRDGrilla.TextMatrix(i, 1))), Trim(iCanti), Trim(iPrecio), Trim(mIvaFE), "M", "0", "0")
                If mRespuestaFiscal = False Then Exit Sub
            End If
            'If txtTasaVial.Text <> "0,000" Then
            '    mTVialDbl = CDbl(GRDGrilla.TextMatrix(i, 11))
            '    mTasaVial = Str(mTVialDbl * mContPrecio)
            '    iCanti = Str(Format(CDbl(GRDGrilla.TextMatrix(i, 2)), "0.00") * mContCanti)
            '    'iPrecio = Str(mPneto * mContPrecio)
            '    If cboFactura.ItemData(cboFactura.ListIndex) = 10000 Then  '"T" TIKET
            '        'mRespuestaFiscal = pf.SendTicketItem(Trim(ChkNull(grdGrilla.TextMatrix(I, 1))), Trim(iCanti), Trim(iPrecio), Trim(mIvaFE), "M", "0", "0", Trim(iImpInt))
            '        mRespuestaFiscal = pf.SendTicketItem("Tasa Vial", Trim(iCanti), Trim(mTasaVial), "0", "M", "0", "0", "")
            '        If mRespuestaFiscal = False Then Exit Sub
            '    Else
            '        mRespuestaFiscal = pf.SendInvoiceItem("Tasa Vial", Trim(mTasaVial), Trim(iCanti), "", "M", "0", "0", "", "", "", "", "")
            '        If mRespuestaFiscal = False Then Exit Sub
            '    End If
            'End If
            
        End If
    Next
    
    'imprimo la tasa vial
'    If txtTasaVial.Text <> "0.00" Then
'        mTVialDbl = CDbl(txtTasaVial.Text)
'        mTasaVial = Str(mTVialDbl * mContPrecio)
'        iCanti = Str(Format(CDbl(grdGrilla.TextMatrix(I, 2)), "0.00") * mContCanti)
'        'iPrecio = Str(mPneto * mContPrecio)
'        If cboFactura.ItemData(cboFactura.ListIndex) = 10000 Then  '"T" TIKET
'            'mRespuestaFiscal = pf.SendTicketItem(Trim(ChkNull(grdGrilla.TextMatrix(I, 1))), Trim(iCanti), Trim(iPrecio), Trim(mIvaFE), "M", "0", "0", Trim(iImpInt))
'            mRespuestaFiscal = pf.SendTicketItem("Tasa Vial", "1", Trim(mTasaVial), "0", "M", "0", "0", "")
'            If mRespuestaFiscal = False Then Exit Sub
'        Else
'            mRespuestaFiscal = pf.SendInvoiceItem("Tasa Vial", Trim(mTasaVial), "1", "", "M", "0", "0", "", "", "", "", "")
'            If mRespuestaFiscal = False Then Exit Sub
'        End If
'    End If
    
'    sql = "SELECT VENTITEM.CODPIEZA, VENTITEM.DESCRIP, VENTITEM.CANTIDAD, VENTITEM.PRECIO, TOTAL"
'    sql = sql & " FROM VENTITEM, STOCK"
'    sql = sql & " WHERE NROFAC = " + XN(mNroFactu.Text) + " AND TIPO_FAC=" + XS(Ltipo_fac.Caption) + " AND SUC_CODIGO = " & XN(txtSucursal.Text)
'    sql = sql & " AND STOCK.CODPIEZA=VENTITEM.CODPIEZA"
'    If snp.State = 1 Then snp.Close
'    snp.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    Do While Not snp.EOF
'        iCanti = Str(Val(ChkNull(snp!cantidad)) * mContCanti)
'        iPrecio = Str((Val(ChkNull(snp!precio))) * mContPrecio)
'        iImpInt = Str((0 * Val(ChkNull(snp!precio))) * mContInternos)
'        mIvaFE = Str(mIVAi * 100)
'
'        If lblTipo_Ticket.Caption = "T" Then
'            mRespuestaFiscal = pf.SendTicketItem(Trim(ChkNull(snp!DESCRIP)), Trim(iCanti), Trim(iPrecio), Trim(mIvaFE), "M", "0", "0", Trim(iImpInt))
'            If mRespuestaFiscal = False Then Exit Sub
'        Else
'            mRespuestaFiscal = pf.SendInvoiceItem(Trim(ChkNull(snp!DESCRIP)), Trim(iCanti), Trim(iPrecio), Trim(mIvaFE), "M", "0", "0", ChkNull(snp!codpieza), "", "", "", Trim(iImpInt))
'            If mRespuestaFiscal = False Then Exit Sub
'        End If
'        snp.MoveNext
'    Loop
    
    'DESCUENTOS
'    If Val(txtImpDescGral.Caption) > 0 Then
'        mTotaLl = Val(txtImpDescGral.Caption) * 100
'        ITOTAL = Trim(Str(mTotaLl))
'        If lblTipo_Ticket.Caption = "T" Then
'            mRespuestaFiscal = pf.SendTicketPayment("DESCUENTO " + Trim(txtDescuentoGral.Text) + " %", Trim(ITOTAL), "D")
'            If mRespuestaFiscal = False Then Exit Sub
'        End If
'        If lblTipo_Ticket.Caption = "A" Then
'            mRespuestaFiscal = pf.SendInvoicePayment("DESCUENTO " + Trim(txtDescuentoGral.Text) + " %", Trim(ITOTAL), "D")
'            If mRespuestaFiscal = False Then Exit Sub
'        End If
'        If lblTipo_Ticket.Caption = "B" Then
'            mRespuestaFiscal = pf.SendInvoicePayment("DESCUENTO " + Trim(txtDescuentoGral.Text) + " %", Trim(ITOTAL), "D")
'            If mRespuestaFiscal = False Then Exit Sub
'        End If
'    End If
    
    'RECARGOS
'    If Val(lblRecargo.Caption) > 0 Then
'        mTotaLl = Val(lblRecargo.Caption) * 100
'        ITOTAL = Trim(Str(mTotaLl))
'        If lblTipo_Ticket.Caption = "T" Then
'            mRespuestaFiscal = pf.SendTicketPayment("RECARGO " + Trim(txtPorcRecargo.Text) + " %", Trim(ITOTAL), "R")
'            If mRespuestaFiscal = False Then Exit Sub
'        End If
'        If lblTipo_Ticket.Caption = "A" Then
'            mRespuestaFiscal = pf.SendInvoicePayment("RECARGO " + Trim(txtPorcRecargo.Text) + " %", Trim(ITOTAL), "R")
'            If mRespuestaFiscal = False Then Exit Sub
'        End If
'        If lblTipo_Ticket.Caption = "B" Then
'            mRespuestaFiscal = pf.SendInvoicePayment("RECARGO " + Trim(txtPorcRecargo.Text) + " %", Trim(ITOTAL), "R")
'            If mRespuestaFiscal = False Then Exit Sub
'        End If
'    End If
    
    'PAGOS
    If cboFactura.ItemData(cboFactura.ListIndex) = 10000 Then 'TIKET Then
        mRespuestaFiscal = pf.GetTicketSubtotal("P", "SUBTOTAL")
        If mRespuestaFiscal = False Then Exit Sub
    End If
    If cboFactura.ItemData(cboFactura.ListIndex) = 1 Then 'factura A Then
        mRespuestaFiscal = pf.GetInvoiceSubtotal("P", "SUBTOTAL")
        'ticket
        'mRespuestaFiscal = pf.GetTicketSubtotal("P", "SUBTOTAL")
        
        If mRespuestaFiscal = False Then Exit Sub
        
    End If
    If cboFactura.ItemData(cboFactura.ListIndex) = 2 Then 'factura B Then
        mRespuestaFiscal = pf.GetInvoiceSubtotal("P", "SUBTOTAL")
        If mRespuestaFiscal = False Then Exit Sub
    End If
    
    For i = 1 To grdPagos.Rows - 1
        mTotaLl = CDbl(grdPagos.TextMatrix(i, 1)) * 100
        ITOTAL = Str(mTotaLl)
        If cboFactura.ItemData(cboFactura.ListIndex) = 10000 Then 'TIKET Then
            mRespuestaFiscal = pf.SendTicketPayment(Mid(grdPagos.TextMatrix(i, 0), 1, 20), Trim(ITOTAL), "T")
            
            If mRespuestaFiscal = False Then Exit Sub
        End If
         If cboFactura.ItemData(cboFactura.ListIndex) = 1 Then 'factura A
            mRespuestaFiscal = pf.SendInvoicePayment(Mid(grdPagos.TextMatrix(i, 0), 1, 20), Trim(ITOTAL), "T")
            'ticket
            'mRespuestaFiscal = pf.SendTicketPayment(Mid(grdPagos.TextMatrix(i, 0), 1, 20), Trim(ITOTAL), "T")
            If mRespuestaFiscal = False Then Exit Sub
        End If
        If cboFactura.ItemData(cboFactura.ListIndex) = 2 Then 'factura B
            mRespuestaFiscal = pf.SendInvoicePayment(Mid(grdPagos.TextMatrix(i, 0), 1, 20), Trim(ITOTAL), "T")
            If mRespuestaFiscal = False Then Exit Sub
        End If
    Next
     
    ''**********aca poner lo de la leyenda final***********
    'LO HICE UNA VEZ Y YA QUEDA PUESTA
    'Dim leyenda1 As String
    'Dim leyenda2 As String
       
    'leyenda1 = "Res Nº54 14/08/96, Disp. Nº285 08/10/98"
    'leyenda1 = "Los Comb cumplen la Res Nº54 14/08/96"
    'leyenda2 = "y con la Disp Nº 285 del 08/10/98"
    'leyenda3 = "la Res de la ex Sec. de O y S Pub Nº54"
    'leyenda4 = "del 14/08/96 y la disp. de la subsec. "
    'leyenda5 = "de comb. Nº 285 del fecha 08/10/98"
    
    'leyenda = leyenda & " ESTABLECIDAS EN LA RES. DE LA EX SECRETARIA DE OBRAS SERV. PUBLICOS "
    'leyenda = leyenda & " N°54 DE FECHA 14/08/1996 Y CON LA DISP. DE LA SUBSECRETARIA DE COMB."
    'leyenda = leyenda & " N°285 DE FECHA 08/10/1998"
    
    'If mRespuestaFiscal Then mRespuestaFiscal = Me.pf.SetGetHeaderTrailer("S", "9", "")
    'If mRespuestaFiscal Then mRespuestaFiscal = Me.pf.SetGetHeaderTrailer("S", "10", "")
    'If mRespuestaFiscal Then mRespuestaFiscal = Me.pf.SetGetHeaderTrailer("S", "11", "")
    'If mRespuestaFiscal Then mRespuestaFiscal = Me.pf.SetGetHeaderTrailer("S", "12", "")
    'If mRespuestaFiscal Then mRespuestaFiscal = Me.pf.SetGetHeaderTrailer("S", "13", "")
    'If mRespuestaFiscal Then mRespuestaFiscal = Me.pf.SetGetHeaderTrailer("S", "13", "Los Comb cumplen la Res Nro 54 14/08/96")
    'If mRespuestaFiscal Then mRespuestaFiscal = Me.pf.SetGetHeaderTrailer("S", "14", "y con la Disp Nro 285 del 08/10/98")
    
    'If mRespuestaFiscal Then mRespuestaFiscal = Me.pf.SetGetHeaderTrailer("G", "11")
    '-----------------------------------------------------
    
    'CIERRO COMPROBANTE
    If cboFactura.ItemData(cboFactura.ListIndex) = 10000 Then 'TIKET Then
        pf.GetTicketSubtotal "P", "SUBTOTAL"
        txtTotalFiscal.Text = Round(CDbl(Chk0(pf.AnswerField_5)) / 100, 2)
        txtIvaFiscal.Text = Round(CDbl(Chk0(pf.AnswerField_9)) / 100, 2) ' es el impuesto interno
        txtNetoFiscal.Text = Round(CDbl(Chk0(pf.AnswerField_10)) / 100, 2)
        mRespuestaFiscal = pf.CloseTicket
        
        
        If mRespuestaFiscal = False Then Exit Sub
    End If
    If cboFactura.ItemData(cboFactura.ListIndex) = 1 Then 'factura A Then
        pf.GetInvoiceSubtotal "P", "SUBTOTAL"
        txtTotalFiscal.Text = Round(CDbl(Chk0(pf.AnswerField_5)) / 100, 2)
        txtIvaFiscal.Text = Round(CDbl(Chk0(pf.AnswerField_9)) / 100, 2) ' es el impuesto interno
        txtNetoFiscal.Text = Round(CDbl(Chk0(pf.AnswerField_10)) / 100, 2)
        mRespuestaFiscal = pf.CloseInvoice("T", "A", "TOTAL")
        
        
        'ACTUALIZO LA FACTURA CON EL IMPORTE EXACTO DEL TICKET
'        sql = "UPDATE FACTURA_CLIENTE SET"
'        sql = sql & " FCL_TOTAL=" & XN(txtTotalFiscal.Text)
'        sql = sql & " ,FCL_IMPIVA=" & XN(txtIvaFiscal.Text)
'        sql = sql & " ,FCL_SUBTOTAL=" & XN(txtNetoFiscal.Text)
'        sql = sql & " WHERE FCL_SUCURSAL=" & XN(txtNroSucursal.Text)
'        sql = sql & " AND FCL_NUMERO=" & XN(txtNroFactura.Text)
'        sql = sql & " AND TCO_CODIGO=1" ' FAC A
        
        
        'TICKET
        'mRespuestaFiscal = pf.CloseTicket
        
        If mRespuestaFiscal = False Then Exit Sub
        
    End If
    If cboFactura.ItemData(cboFactura.ListIndex) = 2 Then 'factura B
        'pf.GetInvoiceSubtotal "P", "SUBTOTAL"
        txtTotalFiscal.Text = Round(CDbl(Chk0(pf.AnswerField_5)) / 100, 2)
        txtIvaFiscal.Text = Round(CDbl(Chk0(pf.AnswerField_9)) / 100, 2) ' es el impuesto interno
        txtNetoFiscal.Text = Round(CDbl(Chk0(pf.AnswerField_10)) / 100, 2)
        mRespuestaFiscal = pf.CloseInvoice("T", "B", "TOTAL")
        
        'ACTUALIZO LA FACTURA CON EL IMPORTE EXACTO DEL TICKET
'        sql = "UPDATE FACTURA_CLIENTE SET"
'        sql = sql & " FCL_TOTAL=" & XN(txtTotalFiscal.Text)
'        sql = sql & " ,FCL_IMPIVA=" & XN(txtIvaFiscal.Text)
'        sql = sql & " ,FCL_SUBTOTAL=" & XN(txtNetoFiscal.Text)
'        sql = sql & " WHERE FCL_SUCURSAL=" & XN(txtNroSucursal.Text)
'        sql = sql & " AND FCL_NUMERO=" & XN(txtNroFactura.Text)
'        sql = sql & " AND TCO_CODIGO=2" ' FAC A
        
        If mRespuestaFiscal = False Then Exit Sub
    End If
    
    
End Sub


'Public Sub ImprimirFactura()
'    Dim Renglon As Double
'
'    Screen.MousePointer = vbHourglass
'    lblEstado.Caption = "Imprimiendo..."
'
'    ImprimirEncabezado
'
'    '---- IMPRESION DE LA FACTURA ------------------
'    Renglon = 2.5
'    Printer.FontSize = 6
'    For i = 1 To GRDGrilla.Rows - 1
'        If GRDGrilla.TextMatrix(i, 0) <> "" Then
'
'            Imprimir 0.5, Renglon, False, "(" & Trim(GRDGrilla.TextMatrix(i, 0)) & ") " & Trim(GRDGrilla.TextMatrix(i, 1))
'            Imprimir 6.8, Renglon, False, " x " & Trim(GRDGrilla.TextMatrix(i, 2)) & "     $" & CompletarConEspaciosIz(Trim(GRDGrilla.TextMatrix(i, 4)), 8)
'            'PARA LA SEGUNDA HOJA
'            Imprimir 10.5, Renglon, False, "(" & Trim(GRDGrilla.TextMatrix(i, 0)) & ") " & Trim(GRDGrilla.TextMatrix(i, 1))
'            Imprimir 16.8, Renglon, False, " x " & Trim(GRDGrilla.TextMatrix(i, 2)) & "     $" & CompletarConEspaciosIz(Trim(GRDGrilla.TextMatrix(i, 4)), 8)
'            Renglon = Renglon + 0.4 '0.8
'        End If
'    Next i
'
'    Printer.FontSize = 9
'    Renglon = 8
'    Printer.Line (0.4, Renglon)-(9, Renglon), , B
'    Imprimir 5.7, Renglon + 0.1, True, "TOTAL  " & Trim(TxtTotal.Text)
'    Printer.Line (0.4, Renglon + 0.6)-(9, Renglon + 0.6), , B
'    'PARA LA SEGUNDA HOJA
'    Printer.Line (10.4, Renglon)-(19, Renglon), , B
'    Imprimir 15.7, Renglon + 0.1, True, "TOTAL  " & Trim(TxtTotal.Text)
'    Printer.Line (10.4, Renglon + 0.6)-(19, Renglon + 0.6), , B
'
'    'PARA CAMBIOS
'    Printer.FontSize = 7
'    Imprimir 0.5, Renglon + 0.7, False, "- P/Cambios presentar esta Boleta"
'    'PARA LA SEGUNDA HOJA
'    Imprimir 10.5, Renglon + 0.7, False, "- P/Cambios presentar esta Boleta"
'    Printer.EndDoc
'    Screen.MousePointer = vbNormal
'    lblEstado.Caption = ""
'End Sub
'
'Public Sub ImprimirEncabezado()
' '-----------IMPRIME EL ENCABEZADO DE LA FACTURA-------------------
'    Set Rec1 = New ADODB.Recordset
'    sql = "SELECT P.RAZ_SOCIAL, S.SUC_DESCRI"
'    sql = sql & " FROM PARAMETROS P, SUCURSAL S"
'    sql = sql & " WHERE S.SUC_CODIGO=P.SUCURSAL"
'    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    If Rec1.EOF = False Then
'        Printer.FontSize = 12
'        Imprimir 0.5, 0.7, True, Trim(Rec1!RAZ_SOCIAL) & CompletarConEspaciosIz("X", 14)
'        Printer.FontSize = 8
'        Imprimir 0.5, 1.4, True, " Nº " & Trim(txtNroSucursal.Text) & "-" & Trim(txtNroFactura.Text) '& "   (Original)"
'        Imprimir 5, 1.4, True, Format(FechaFactura, "dd/mm/yyyy")
'        Printer.FontSize = 7
'        Imprimir 3.3, 1.4, False, "(Original)"
'
'        'DOCUMENTO NO VALIDO COMO FACTURA
'        Printer.FontSize = 7
'        Imprimir 6.8, 0.7, False, "   Movimiento  "
'        Imprimir 6.8, 1, False, "      Interno    "
'        Imprimir 6.8, 1.3, False, "(Doc. no valido"
'        Imprimir 6.8, 1.6, False, "como Factura)"
'
'        'PARA LA SEGUNDA HOJA
'        Printer.FontSize = 12
'        Imprimir 10.5, 0.7, True, Trim(Rec1!RAZ_SOCIAL) & CompletarConEspaciosIz("X", 10)
'        Printer.FontSize = 8
'        Imprimir 10.5, 1.4, True, " Nº " & Trim(txtNroSucursal.Text) & "-" & Trim(txtNroFactura.Text) '& "   (Duplicado)"
'        Imprimir 15, 1.4, True, Format(FechaFactura, "dd/mm/yyyy")
'        Printer.FontSize = 7
'        Imprimir 13.3, 1.4, False, "(Duplicado)"
'
'        'DOCUMENTO NO VALIDO COMO FACTURA
'        Printer.FontSize = 7
'        Imprimir 16.8, 0.7, False, "   Movimiento  "
'        Imprimir 16.8, 1, False, "      Interno    "
'        Imprimir 16.8, 1.3, False, "(Doc. no valido"
'        Imprimir 16.8, 1.6, False, "como Factura)"
'    End If
'    Rec1.Close
'
'    sql = "SELECT C.CLI_CODIGO, C.CLI_RAZSOC"
'    sql = sql & " FROM CLIENTE C"
'    sql = sql & " WHERE C.CLI_CODIGO=" & XN(txtcodCli.Text)
'    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    If Rec1.EOF = False Then
'        Printer.FontSize = 7
'        Imprimir 0.5, 1.9, False, "(" & Trim(Rec1!CLI_CODIGO) & ") " & Trim(Rec1!CLI_RAZSOC)
'        'PARA LA SEGUNDA HOJA
'        Imprimir 10.5, 1.9, False, "(" & Trim(Rec1!CLI_CODIGO) & ") " & Trim(Rec1!CLI_RAZSOC)
'    End If
'    Rec1.Close
'    Printer.Line (0.4, 2.3)-(9, 2.3), , B
'    Printer.Line (10.4, 2.3)-(19, 2.3), , B
'End Sub

Private Sub LIMPIOGRILLA()
    For i = 1 To grdGrilla.Rows - 1
        grdGrilla.TextMatrix(i, 0) = ""
        grdGrilla.TextMatrix(i, 1) = ""
        grdGrilla.TextMatrix(i, 2) = ""
        grdGrilla.TextMatrix(i, 3) = ""
        grdGrilla.TextMatrix(i, 4) = ""
        grdGrilla.TextMatrix(i, 5) = ""
        grdGrilla.TextMatrix(i, 6) = 0
        grdGrilla.TextMatrix(i, 7) = 0
        grdGrilla.TextMatrix(i, 8) = 0
        grdGrilla.TextMatrix(i, 9) = 0
        grdGrilla.TextMatrix(i, 10) = 0
    Next
End Sub

Private Sub cmdModificarCli_Click()
    FormLlamado = "frmFacturaCliente"
    If txtcodCli.Text <> "" Then
        ABMClientes.vMode = 2
        ABMClientes.vFieldID = "'" & txtcodCli.Text & "'"
        ABMClientes.txtId.Text = txtcodCli.Text
        ABMClientes.Show
    End If
End Sub

Private Sub cmdNC_Click()

If MsgBox("Confirma la impresion de la Nota de Credito?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then

''BUSCO EL NUMERO DE NOTA DE CREDITO EN EL FISCAL
    If FISCAL = "TMT900FA" Then
        ImprimoFiscalEpsondll 3
    Else
        Select Case cboFactura.ItemData(cboFactura.ListIndex)
            Case 1 'NOTA CREDITO A
                pf.Status ("A")
                txtNroFactura.Text = Val(pf.AnswerField_11) + 1
            Case 2 'NOTA CREDITO B
                pf.Status ("A")
                txtNroFactura.Text = Val(pf.AnswerField_12) + 1
        End Select
    
        Imprimo_NC_Fiscal
        'GrabarNC
    End If
    
'    If cboFactura.ItemData(cboFactura.ListIndex) = 1 Then 'NOTA CREDITO A
'            mRespuestaFiscal = pf.OpenInvoice("M", "C", "A", "1", "P", 12, "I", mRespo.Text, txtRazSoc.Text, "-", "CUIT", txtCuit.Text, "N", Trim(mVendedor), cpago, usua, cturno, "-", "C")
'
'            If mRespuestaFiscal = False Then Exit Sub
'    End If
'
'    If cboFactura.ItemData(cboFactura.ListIndex) = 2 Then 'NOTA CREDITO B
'          mRespuestaFiscal = pf.OpenInvoice("M", "C", "B", 1, "P", 12, "I", mRespo.Text, txtRazSoc.Text, "-", "DNI", txtNRO_DOCUMENTO.Text, "N", Trim(mVendedor), cpago, usua, cturno, "-", "C")
'    End If
'
'
'
'    For i = 1 To GRDGrilla.Rows - 1
'        If GRDGrilla.TextMatrix(i, 0) <> "" Then
'            iCanti = Str(CDbl(GRDGrilla.TextMatrix(i, 2)) * mContCanti)
'            iPrecio = Str(CDbl(GRDGrilla.TextMatrix(i, 3)) * mContPrecio)
'            iImpInt = Str((0 * CDbl(GRDGrilla.TextMatrix(i, 3))) * mContInternos)
'            'mIvaFE = Str(mIVAi * 100)
'            mIvaFE = Str(CDbl(GRDGrilla.TextMatrix(i, 6)) * 100)
'
'            If cboFactura.ItemData(cboFactura.ListIndex) = 10000 Then  '"T" TIKET
'                mRespuestaFiscal = pf.SendTicketItem(Trim(ChkNull(GRDGrilla.TextMatrix(i, 1))), Trim(iCanti), Trim(iPrecio), Trim(mIvaFE), "M", "0", "0", Trim(iImpInt))
'                If mRespuestaFiscal = False Then Exit Sub
'            Else
'                'mRespuestaFiscal = pf.SendInvoiceItem(Trim(ChkNull(grdGrilla.TextMatrix(i, 1))), Trim(iCanti), Trim(iPrecio), Trim(mIvaFE), "M", "0", "0", ChkNull(grdGrilla.TextMatrix(i, 0)) , "", "", "", Trim(iImpInt))
'                mRespuestaFiscal = pf.SendInvoiceItem(Trim(ChkNull(GRDGrilla.TextMatrix(i, 1))), Trim(iCanti), Trim(iPrecio), Trim(mIvaFE), "M", "0", "0", "", "", "", "", Trim(iImpInt))
'                If mRespuestaFiscal = False Then Exit Sub
'            End If
'        End If
'    Next
'
'
''PAGOS
'    If cboFactura.ItemData(cboFactura.ListIndex) = 10000 Then 'TIKET Then
'        mRespuestaFiscal = pf.GetTicketSubtotal("P", "SUBTOTAL")
'        If mRespuestaFiscal = False Then Exit Sub
'    End If
'    If cboFactura.ItemData(cboFactura.ListIndex) = 1 Then 'NOTA CREDITO A Then
'        mRespuestaFiscal = pf.GetInvoiceSubtotal("P", "SUBTOTAL")
'        If mRespuestaFiscal = False Then Exit Sub
'    End If
'    If cboFactura.ItemData(cboFactura.ListIndex) = 2 Then 'NOTA CREDITO B Then
'        mRespuestaFiscal = pf.GetInvoiceSubtotal("P", "SUBTOTAL")
'        If mRespuestaFiscal = False Then Exit Sub
'    End If
'''CIERRO COMPROBANTE
''
'    If cboFactura.ItemData(cboFactura.ListIndex) = 1 Then 'nota de credito A
'        'pf.GetInvoiceSubtotal "P", "SUBTOTAL"
'        txtTotalFiscal.Text = Round(CDbl(Chk0(pf.AnswerField_5)) / 100, 2)
'        txtIvaFiscal.Text = Round(CDbl(Chk0(pf.AnswerField_6)) / 100, 2)
'        txtNetoFiscal.Text = Round(CDbl(Chk0(pf.AnswerField_10)) / 100, 2)
'        mRespuestaFiscal = pf.CloseInvoice("M", "A", "TOTAL")
'        If mRespuestaFiscal = False Then Exit Sub
'    End If
'    If cboFactura.ItemData(cboFactura.ListIndex) = 2 Then 'nota de credito B
'        'pf.GetInvoiceSubtotal "P", "SUBTOTAL"
'        txtTotalFiscal.Text = Round(CDbl(Chk0(pf.AnswerField_5)) / 100, 2)
'        txtIvaFiscal.Text = Round(CDbl(Chk0(pf.AnswerField_6)) / 100, 2)
'        txtNetoFiscal.Text = Round(CDbl(Chk0(pf.AnswerField_10)) / 100, 2)
'        mRespuestaFiscal = pf.CloseInvoice("M", "B", "TOTAL")
'        If mRespuestaFiscal = False Then Exit Sub
'    End If
    CmdNuevo_Click
End If
End Sub
Private Function GrabarNC()
    sql = "SELECT * FROM NOTA_CREDITO_CLIENTE"
    sql = sql & " WHERE TCO_CODIGO=" & cboFactura.ItemData(cboFactura.ListIndex)
    sql = sql & " AND NCC_NUMERO = " & XN(txtNroFactura.Text)
    sql = sql & " AND NCC_SUCURSAL=" & XN(txtNroSucursal.Text)
    If rec.State = 1 Then rec.Close
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic

    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando..."

    If rec.EOF = True Then
        'NUEVA NOTA_CREDITO
        sql = "INSERT INTO NOTA_CREDITO_CLIENTE"
        sql = sql & " (TCO_CODIGO,NCC_NUMERO,NCC_SUCURSAL,NCC_FECHA,"
        sql = sql & " NCC_IVA,NCC_IMPIVA,FPG_CODIGO,NCC_OBSERVACION,VEN_CODIGO,"
        sql = sql & " NCC_SUBTOTAL,NCC_TOTAL,NCC_SALDO,EST_CODIGO,"
        sql = sql & " NCC_NUMEROTXT,NCC_SUCURSALTXT,CLI_CODIGO,NCC_IMPINT,TUR_CODIGO,NCC_HORA,NCC_TASAVIAL,NCC_TOTALACT)"
        sql = sql & " VALUES ("
        sql = sql & cboFactura.ItemData(cboFactura.ListIndex) & ","
        sql = sql & XN(txtNroFactura.Text) & ","
        sql = sql & XN(txtNroSucursal.Text) & ","
        sql = sql & XDQ(FechaFactura.Value) & ","
        sql = sql & XN(txtPorcentajeIva.Text) & ","
        sql = sql & XN(txtImporteIvaB.Text) & "," 'uso este campo no visible para guardar la info de las FAC B
        sql = sql & cboFormaPago.ItemData(cboFormaPago.ListIndex) & ","
        sql = sql & XS(txtObservaciones) & ","
        sql = sql & cboVendedor.ItemData(cboVendedor.ListIndex) & ","
        sql = sql & XN(txtsubtotal1B.Text) & "," 'uso este campo no visible para guardar la info de las FAC B
        sql = sql & XN(txtTotal.Text) & ","
        sql = sql & XN(txtTotal.Text) & "," 'SALDO NOTA_CREDITO
        sql = sql & "3," 'ESTADO DEFINITIVO
        sql = sql & XS(Format(txtNroFactura.Text, "00000000")) & ","
        sql = sql & XS(Format(txtNroSucursal.Text, "0000")) & ","
        sql = sql & XN(txtcodCli.Text) & "," 'CLIENTE
        sql = sql & XN(txtimpuestoB.Text) & "," 'uso este campo no visible para guardar la info de las FAC B
        sql = sql & cboTurno.ItemData(cboTurno.ListIndex) & ","
        sql = sql & XS(Format(Time, "hh:mm")) & ","
        sql = sql & XN(txttasavial.Text) & ","
        sql = sql & XN(txtTotal.Text) & ")"
        
        'Format(Valor, "mm/dd/yyyy") & "#"
        DBConn.Execute sql
           
        For i = 1 To grdGrilla.Rows - 1
            If grdGrilla.TextMatrix(i, 0) <> "" Then
                sql = "INSERT INTO DETALLE_NOTA_CREDITO_CLIENTE"
                sql = sql & " (TCO_CODIGO, NCC_NUMERO, NCC_SUCURSAL, DFC_NROITEM,"
                sql = sql & " PTO_CODIGO, DFC_CONCEPTO, DFC_CANTIDAD, DFC_PRECIO, DFC_IVA,"
                sql = sql & " DFC_IMP, DFC_MONIMP, DFC_MONIVA,DFC_TASAVIAL,DFC_TOTALTVIAL)"
                ' guardar el imp, el monto del imp y el monto del iva
                
                sql = sql & " VALUES ("
                sql = sql & cboFactura.ItemData(cboFactura.ListIndex) & ","
                sql = sql & XN(txtNroFactura.Text) & ","
                sql = sql & XN(txtNroSucursal.Text) & ","
                sql = sql & i & "," 'PONER EL NRO ITEM
                sql = sql & XN(grdGrilla.TextMatrix(i, 0)) & ","
                sql = sql & XS(grdGrilla.TextMatrix(i, 1)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 2)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 3)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 6)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 7)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 8)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 9)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 11)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 12)) & ")"
                
                DBConn.Execute sql
                
                sql = "SELECT DST_STKFIS,DST_STKCON"
                sql = sql & " FROM STOCK"
                sql = sql & " WHERE STK_CODIGO = " & XN(Sucursal)
                sql = sql & " AND PTO_CODIGO = " & XN(grdGrilla.TextMatrix(i, 0))
                Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
                If Rec1.EOF = False Then
                    sql = "UPDATE STOCK SET"
                    sql = sql & " DST_STKFIS = DST_STKFIS - " & XN(grdGrilla.TextMatrix(i, 2))
                    sql = sql & " WHERE STK_CODIGO = " & XN(Sucursal)
                    sql = sql & " AND PTO_CODIGO = " & XN(grdGrilla.TextMatrix(i, 0))
                    DBConn.Execute sql
                End If
                Rec1.Close
            End If
        Next
        
        For i = 1 To grdPagos.Rows - 1
            sql = "insert into FACTURA_PAGOS (NCC_SUCURSAL,NCC_NUMERO,TCO_CODIGO,FPG_CODIGO,pag_importe,TAR_CODIGO,"
            sql = sql & " TAR_PLAN,TAR_CUPON, TAR_LOTE,TAR_AUTORIZACION,TOTDOLAR, COTIZACION, sen_sucursal, sen_tipo,"
            sql = sql & " sen_nro, PAG_SECUENCIA, FPG_SALDO,CHE_BANCO,CHE_NUMERO)"
            sql = sql & " values ("
            sql = sql & XN(txtNroSucursal.Text) & ", " & XN(txtNroFactura.Text) & ", " & XN(cboFactura.ItemData(cboFactura.ListIndex)) & ", "
            sql = sql & XN(grdPagos.TextMatrix(i, 2)) & ","
            sql = sql & XN(grdPagos.TextMatrix(i, 1))
            sql = sql & "," & XN(grdPagos.TextMatrix(i, 3))
            sql = sql & "," & XN(grdPagos.TextMatrix(i, 5))
            sql = sql & "," & XN(grdPagos.TextMatrix(i, 7))
            sql = sql & "," & XN(grdPagos.TextMatrix(i, 8))
            sql = sql & "," & XN(grdPagos.TextMatrix(i, 9))
            sql = sql & "," & XN(grdPagos.TextMatrix(i, 10))
            sql = sql & "," & XN(grdPagos.TextMatrix(i, 11))
            sql = sql & "," & XN(grdPagos.TextMatrix(i, 12))
            sql = sql & "," & XN(grdPagos.TextMatrix(i, 13))
            sql = sql & "," & XS(grdPagos.TextMatrix(i, 14)) & "," & i & ","
            If grdPagos.TextMatrix(i, 2) <> "2" Then
                sql = sql & "0"
            Else
                sql = sql & XN(grdPagos.TextMatrix(i, 1))
                mSaldo = mSaldo + CDbl(Chk0(grdPagos.TextMatrix(i, 1)))
            End If
            If grdPagos.TextMatrix(i, 2) = "5" Then
                sql = sql & "," & XS(grdPagos.TextMatrix(i, 4))
                sql = sql & "," & XS(grdPagos.TextMatrix(i, 6)) & ")"
            Else
                sql = sql & "," & XS("")
                sql = sql & "," & XS("") & ")"
            End If
            
            DBConn.Execute sql
        Next
        
'            sql = "UPDATE NOTA_CREDITO_CLIENTE"
'            sql = sql & " SET NCC_SALDO=" & XN(CStr(mSaldo))
'            sql = sql & " WHERE"
'            sql = sql & " TCO_CODIGO=" & cboFactura.ItemData(cboFactura.ListIndex)
'            sql = sql & " AND NCC_NUMERO=" & XN(txtNroFactura.Text)
'            sql = sql & " AND NCC_SUCURSAL=" & XN(txtNroSucursal.Text)
'            DBConn.Execute sql
        
        'ACTUALIZO LA TABLA PARAMENTROS Y LE SUMO UNO A LA NOTA_CREDITO QUE CORRESPONDE
'        Select Case cboFactura.ItemData(cboFactura.ListIndex)
'            Case 1
'                sql = "UPDATE PARAMETROS SET NOTA_CREDITO_C=" & XN(txtNroFactura.Text)
'            Case 2
'                sql = "UPDATE PARAMETROS SET NOTA_CREDITO_C=" & XN(txtNroFactura.Text)
'        End Select
'        DBConn.Execute sql
    End If
    rec.Close
End Function
Private Sub CmdNuevo_Click()
   mIVA_1 = BuscoIva
   mBuscador = False
   mVerCta = True
   LIMPIOGRILLA
   mFoco = False
   cmdImprimir.Enabled = False
   lblConPago.Caption = ""
   txtNroFactura.Text = ""
   txtNroSucursal.Text = ""
   FechaFactura.Value = Date
   lblEstadoFactura.Caption = ""
   
   txtTotal.Text = "0,000"
   txtPorcentajeIva.Text = Format(mIVA_1, "0.0000")
   
   txtObservaciones.Text = ""
   'cboCondicion.ListIndex = 0
   lblEstado.Caption = ""
       
    txtsubtotal1.Text = "0,000"
    txtimpuesto.Text = "0,000"
    txtSubtotal.Text = "0,000"
    txtImporteIva.Text = "0,000"
    txttasavial.Text = "0,000"
    
    txtimpuestoB.Text = "0,000"
    txtsubtotal1B.Text = "0,000"
    txtSubtotalB.Text = "0,000"
    txtImporteIvaB.Text = "0,000"
   
   
    'CARGO ESTADO
    Call BuscoEstado(1, lblEstadoFactura) 'ESTADO PENDIENTE
   'BUSCO IVA
   'BuscoIva
    fraPagos.Visible = False
    fraTarjeta.Visible = False
    fracheque.Visible = False
    grdPagos.Rows = 1
    
    'txtPorcentajeIva.Text = "0,00"
    VEstadoFactura = 1
    '--------------
    'FrameFactura.Enabled = True
    FrameFactura.Enabled = False
    txtNroSucursal_LostFocus
    txtNroFactura_LostFocus
    
    tabDatos.Tab = 0
    FechaFactura.Value = Date
    cboFactura.ListIndex = 0
    'cboFactura.SetFocus
    txtcodCli.Text = ""
'   txtcodCli.Text = "1"
'   txtCodCli_LostFocus
    cboVendedor.ListIndex = -1
    
    cmdGrabar.Enabled = True
    cmdFormaPago.Enabled = True
    FrameCliente.Enabled = True
    txtcodCli.SetFocus
    
    cboTurno.Clear
    cboTurnosB.Clear
    LlenarComboTurnos
    
    txtBuscaNum.Text = ""
    lblSaldoFac.Visible = False
    cmdNC.Enabled = False
    lblblockeado.Visible = False
    
    cboCondicion.Clear
    cboFormaPago.Clear
    LlenarComboFormaPago
    
End Sub

Private Sub cmdNuevoCli_Click()
'    FormLlamado = "frmFacturaCliente"
'    If txtcodCli.Text <> "" Then
'        ABMClientes.vFieldID = "'" & txtcodCli.Text & "'"
'        ABMClientes.txtID.Text = txtcodCli.Text
'        ABMClientes.vMode = 2
'    Else
'        ABMClientes.vMode = 1
'    End If
'    ABMClientes.Show
    
    txtcodCli.Text = ""
    FormLlamado = "frmFacturaCliente"
    
    'PROGRAMAR ESTE TIPO DE LLAMADAS
    ABMClientes.vMode = 1
    ABMClientes.Show
End Sub

Private Sub CmdSalir_Click()
    If MsgBox("Seguro que desea Salir", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
    mMeDioError = True
    Set frmFacturaCliente = Nothing
    Unload Me
    End If
End Sub

Private Sub cmdFormaPago_Click()
    cboFormaPago.Enabled = True
    fraPagos.Top = 930
    fraPagos.Left = 3345
    fraPagos.Visible = True
    
    Dim mTotalPagos As Double
    mTotalPagos = 0
    For i = 1 To grdPagos.Rows - 1
        mTotalPagos = mTotalPagos + CDbl(grdPagos.TextMatrix(i, 1))
    Next
    txtTotalPagos.Text = Format(CDbl(txtTotal.Text) - mTotalPagos, "0.00")
    
    txtGrabar.Text = "N"
    cboFormaPago.ListIndex = 0
    cboFormaPago.SetFocus
    
'    If txtcodCli.Text = "1" Then
'        cboFormaPago.Enabled = False
'    End If
End Sub
Private Function removedata()

'TMP_LUBRICANTES, TMP_LUBRICANTES_STOCKFINAL, TMP_TASAVIAL
DBConn.Execute "DELETE * From CIERREZ"
DBConn.Execute "DELETE * From DETALLE_ENTRADA_DET_PRODUCTO"
DBConn.Execute "DELETE * From DETALLE_ENTRADA_PRODUCTO"
DBConn.Execute "DELETE * From DETALLE_FACTURA_CLIENTE"
DBConn.Execute "DELETE * From DETALLE_RECIBO_CLIENTE"
DBConn.Execute "DELETE * From ENTRADA_PRODUCTO"
DBConn.Execute "DELETE * From FACTURA_CLIENTE"
DBConn.Execute "DELETE * From FACTURA_PAGOS"
DBConn.Execute "DELETE * From FACTURAS_RECIBO_CLIENTE"
DBConn.Execute "DELETE * From RECIBO_CLIENTE"
DBConn.Execute "DELETE * From T_STOCK"
DBConn.Execute "DELETE * From TMP_CANTVEND"
DBConn.Execute "DELETE * From TMP_FACTURAS"
DBConn.Execute "DELETE * From TMP_INFORME"
DBConn.Execute "DELETE * From TMP_INFORME_RESUMEN"
DBConn.Execute "DELETE * From TMP_LIBRO_IVA_VENTAS"
DBConn.Execute "DELETE * From TMP_LUBRICANTES"
DBConn.Execute "DELETE * From TMP_LUBRICANTES_STOCKFINAL"
DBConn.Execute "DELETE * From TMP_TASAVIAL"


End Function

Private Sub Command2_Click()

End Sub

Private Sub Form_Activate()
'txtcodCli.SetFocus
'    mValorIvaIns = 0
'    If mQuienLlama = "frmComposturas" Then
'        txtcodCli.Text = frmComposturas.txtcodCli.Text
'        txtCodCli_LostFocus
'        mPrecio = 0
'        'BuscaCodigoProxItemData frmComposturas.cboVendedor.ItemData(frmComposturas.cboVendedor.ListIndex), cboVendedor
'        ' ACA HABRIA QUE PONER EL PLAYERO QUE ESTA DE TURNO O LOGUEADO EN LA MAQUINA
'        cboVendedor.ListIndex = 1
'
'        sql = "SELECT * FROM PRODUCTO WHERE PTO_CODIGO=2"
'        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'        If rec.EOF = False Then
'            grdGrilla.TextMatrix(1, 0) = rec!PTO_CODIGO
'            grdGrilla.TextMatrix(1, 1) = "COMP. " & Trim(mDescrip)
'            grdGrilla.TextMatrix(1, 2) = "1"
'            If Ltipo_fac.Caption = "B" Then
'                mPrecio = VALIDO_IMPORTE4(mQueFacturo)
'                txtImporteIva.Text = "0,00"
'            Else
'                mValorIvaIns = (1 + (mIVAi / 100))
'                mPrecio = VALIDO_IMPORTE4(CDbl(mQueFacturo) / mValorIvaIns)
'            End If
'            grdGrilla.TextMatrix(1, 3) = Format(mPrecio, "0.00")
'            grdGrilla.TextMatrix(1, 4) = Format(mPrecio, "0.00")
'            grdGrilla.TextMatrix(1, 5) = Trim(rec!PTO_CODIGO)
'            grdGrilla.TextMatrix(1, 6) = Trim(mIVAi)
'            txtSubtotal.Text = VALIDO_IMPORTE4(CStr(SumaTotal))
'            txtTotal.Text = VALIDO_IMPORTE4(CStr(SumaTotal))
'            txtPorcentajeIva_LostFocus
'            If CDbl(mQueFacturo) - (CDbl(txtSubtotal.Text) + CDbl(txtImporteIva.Text)) > 0 Then
'                txtTotal.Text = CDbl(txtSubtotal.Text) + CDbl(txtImporteIva.Text) + (CDbl(mQueFacturo) - (CDbl(txtSubtotal.Text) + CDbl(txtImporteIva.Text)))
'                txtTotal.Text = VALIDO_IMPORTE4(txtTotal.Text)
'            End If
'            Me.Refresh
'        End If
'        FrameCliente.Enabled = False
'        FrameFactura.Enabled = False
'        'Frame1.Enabled = False
'        cboListaPrecio.Enabled = False
'        Frame3.Enabled = False
'        CmdSalir.Enabled = True
'        CmdNuevo.Enabled = False
'        If cboVendedor.ListCount > 0 Then cboVendedor.ListIndex = 1
'        cmdFormaPago_Click
'
'    ElseIf mQuienLlama = "frmRevelados" Then
'        'txtcodCli.Text = frmRevelados.txtcodCli.Text
'        txtCodCli_LostFocus
'        mPrecio = 0
'        'BuscaCodigoProxItemData frmRevelados.cboVendedor.ItemData(frmComposturas.cboVendedor.ListIndex), cboVendedor
'        sql = "SELECT * FROM PRODUCTO WHERE PTO_CODIGO=3"
'        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'        If rec.EOF = False Then
'            grdGrilla.TextMatrix(1, 0) = rec!PTO_CODIGO
'            grdGrilla.TextMatrix(1, 1) = Trim(rec!PTO_DESCRI)
'            grdGrilla.TextMatrix(1, 2) = "1"
'            If Ltipo_fac.Caption = "B" Then
'                mPrecio = VALIDO_IMPORTE4(mQueFacturo)
'                txtImporteIva.Text = "0,00"
'            Else
'                mValorIvaIns = (1 + (mIVAi / 100))
'                mPrecio = VALIDO_IMPORTE4(CDbl(mQueFacturo) / mValorIvaIns)
'            End If
'            grdGrilla.TextMatrix(1, 3) = Format(mPrecio, "0.00")
'            grdGrilla.TextMatrix(1, 4) = Format(mPrecio, "0.00")
'            grdGrilla.TextMatrix(1, 5) = Trim(rec!PTO_CODIGO)
'            grdGrilla.TextMatrix(1, 6) = Trim(mIVAi)
'            txtSubtotal.Text = VALIDO_IMPORTE4(CStr(SumaTotal))
'            txtTotal.Text = VALIDO_IMPORTE4(CStr(SumaTotal))
'            txtPorcentajeIva_LostFocus
'            If CDbl(mQueFacturo) - (CDbl(txtSubtotal.Text) + CDbl(txtImporteIva.Text)) > 0 Then
'                txtTotal.Text = CDbl(txtSubtotal.Text) + CDbl(txtImporteIva.Text) + (CDbl(mQueFacturo) - (CDbl(txtSubtotal.Text) + CDbl(txtImporteIva.Text)))
'                txtTotal.Text = VALIDO_IMPORTE4(txtTotal.Text)
'            End If
'            Me.Refresh
'        End If
'        FrameCliente.Enabled = False
'        FrameFactura.Enabled = False
'        'Frame1.Enabled = False
'        cboListaPrecio.Enabled = False
'        Frame3.Enabled = False
'        CmdSalir.Enabled = True
'        CmdNuevo.Enabled = False
'        If cboVendedor.ListCount > 0 Then cboVendedor.ListIndex = 1
'        cmdFormaPago_Click
'    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 And ActiveControl.Name <> "grdGrilla" _
       And ActiveControl.Name <> "txtcodCli" And ActiveControl.Name <> "txtRazSoc" _
       And ActiveControl.Name <> "txtBuscaCliente" And ActiveControl.Name <> "txtBuscarCliDescri" Then
        tabDatos.Tab = 1
    End If
    If KeyCode = vbKeyF5 Then
        Dim vDesde(3) As Date
        Dim vHasta(3) As Date
        Dim i As Integer
        FechaFactura = Date
        sql = "SELECT * FROM TURNOS"
        sql = sql & " ORDER BY TUR_CODIGO"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        i = 0
        If rec.EOF = False Then
            Do While rec.EOF = False
                vDesde(i) = rec!TUR_DESDE
                vHasta(i) = rec!TUR_HASTA
                i = i + 1
                rec.MoveNext
            Loop
        End If
        rec.Close
        'POSICIONO EL TURNO DE ACUERDO A LA HORA ACTUAL
        If Time() >= vDesde(0) And Time() <= vHasta(0) Then
            Call BuscaCodigoProxItemData(1, cboTurno)
            Call BuscaCodigoProxItemData(1, cboTurnosB)
        Else
            If Time() >= vDesde(1) And Time() <= vHasta(1) Then
                Call BuscaCodigoProxItemData(2, cboTurno)
                Call BuscaCodigoProxItemData(2, cboTurnosB)
            Else
                Call BuscaCodigoProxItemData(3, cboTurno)
                Call BuscaCodigoProxItemData(3, cboTurnosB)
            End If
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Me.ActiveControl.Name <> "grdGrilla" And _
        Me.ActiveControl.Name <> "txtEdit" And _
        KeyAscii = vbKeyReturn Then
        MySendKeys Chr(9)
    End If
'    If KeyAscii = vbKeyEscape Then
'        cmdSalir_Click
'    End If
End Sub

Private Sub Form_Load()
    
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    Set Rec2 = New ADODB.Recordset
    
    mBuscador = False
    mVerCta = True
    
    'Me.Top = 0
    'Me.Left = 0
    Centrar_pantalla Me
    mIVA_1 = BuscoIva
    txtPorcentajeIva.Text = Format(mIVA_1, "0.0000")
    
    grdGrilla.FormatString = "^Código|<Descipción|^Cant.|>Precio|>Total|Codigo Producto|IVA|IMP|Monto IMP|Monto IVA|Neto B|Tasa Vial|Total Tasa Vial"
    grdGrilla.ColWidth(0) = 1200 'CODIGO
    grdGrilla.ColWidth(1) = 4700 'DESCRIPCION
    grdGrilla.ColWidth(2) = 1200 'CANTIDAD
    grdGrilla.ColWidth(3) = 1400 'PRECIO
    grdGrilla.ColWidth(4) = 1400 'TOTAL
    grdGrilla.ColWidth(5) = 0    'CODIGO PRODUCTO
    grdGrilla.ColWidth(6) = 0    'IVA
    grdGrilla.ColWidth(7) = 0    'IMPUESTO
    grdGrilla.ColWidth(8) = 0    'MONTO IMP
    grdGrilla.ColWidth(9) = 1400    'MONTO IVA
    grdGrilla.ColWidth(10) = 1400    'MONTO neto cuando la Factura es B
    grdGrilla.ColWidth(11) = 1000 'Tasa Vial
    grdGrilla.ColWidth(12) = 1000 'Total Tasa Vial
    grdGrilla.Rows = 30
    grdGrilla.Cols = 13
    'grdGrilla.HighLight = flexHighlightNever
    grdGrilla.BorderStyle = flexBorderNone
    grdGrilla.row = 0
    For i = 0 To grdGrilla.Cols - 1
        grdGrilla.Col = i
        grdGrilla.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        grdGrilla.CellBackColor = &H808080    'GRIS OSCURO
        grdGrilla.CellFontBold = True
    Next
    'Pongo en cero las columnas no visibles de impuestos
    For i = 1 To grdGrilla.Rows - 1
        grdGrilla.TextMatrix(i, 0) = ""
        grdGrilla.TextMatrix(i, 1) = ""
        grdGrilla.TextMatrix(i, 2) = ""
        grdGrilla.TextMatrix(i, 3) = ""
        grdGrilla.TextMatrix(i, 4) = ""
        grdGrilla.TextMatrix(i, 5) = ""
        grdGrilla.TextMatrix(i, 6) = 0
        grdGrilla.TextMatrix(i, 7) = 0
        grdGrilla.TextMatrix(i, 8) = 0
        grdGrilla.TextMatrix(i, 9) = 0
    Next
    
    'GRILLA (GrdModulos) PARA LA BUSQUEDA
    GrdModulos.FormatString = "^Tipo|^Número|^Fecha|Cliente|Cod_Estado|" _
                              & "PORCENTAJE IVA|OBSERVACIONES|" _
                              & "TIPO COMPROBANTE|CONDICION VENTA|CLI CODIGO|TOTAL|IMPIVA|VENDEDOR|TURNO"
                              
    GrdModulos.ColWidth(0) = 900  'TIPO FACTURA
    GrdModulos.ColWidth(1) = 1400 'NUMERO
    GrdModulos.ColWidth(2) = 1200 'FECHA
    GrdModulos.ColWidth(3) = 6000 'CLIENTE
    GrdModulos.ColWidth(4) = 0    'COD_ESTADO
    GrdModulos.ColWidth(5) = 0    'PORCENTAJE IVA
    GrdModulos.ColWidth(6) = 0    'OBSERVACIONES
    GrdModulos.ColWidth(7) = 0    'TIPO COMPROBANTE
    GrdModulos.ColWidth(8) = 0    'CONDICION VENTA
    GrdModulos.ColWidth(9) = 0    'CLI CODIGO
    GrdModulos.ColWidth(10) = 0    'TOTAL FACTURA
    GrdModulos.ColWidth(11) = 0    'IMPORTE IVA
    GrdModulos.ColWidth(12) = 0    'VENDEDOR
    GrdModulos.ColWidth(13) = 0    'TURNO
    GrdModulos.Cols = 14
    GrdModulos.Rows = 1
    GrdModulos.HighLight = flexHighlightNever
    GrdModulos.BorderStyle = flexBorderNone
    GrdModulos.row = 0
    For i = 0 To GrdModulos.Cols - 1
        GrdModulos.Col = i
        GrdModulos.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        GrdModulos.CellBackColor = &H808080    'GRIS OSCURO
        GrdModulos.CellFontBold = True
    Next
    '------------------------------------
    
    lblEstado.Caption = ""
    'CARGO COMBO CON LOS TIPOS DE FACTURA
    LlenarComboFactura
    'CARGO COMBO CON LAS CONDICIONES DE VENTA
    LlenarComboFormaPago
    'CARGO ESTADO
    Call BuscoEstado(1, lblEstadoFactura) 'ESTADO PENDIENTE
    VEstadoFactura = 1
    FechaFactura.Value = Date
    tabDatos.Tab = 0
    'BUSCO IVA
    'BuscoIva
    
    'CARGO COMBO VENDEDOR
    LlenarComboVendedor
    
    'COMBO DE TURNOS
    LlenarComboTurnos
    
    'CargoComboBox cboFormaPago, "FORMA_PAGO", "FPG_CODIGO", "FPG_DESCRI", "FPG_DESCRI"
    If cboFormaPago.ListCount > 0 Then cboFormaPago.ListIndex = 0
    
    
    
    CargoComboBox cboListaPrecio, "LISTA_PRECIO", "LIS_CODIGO", "LIS_DESCRI", "LIS_DESCRI"
    If cboListaPrecio.ListCount > 0 Then cboListaPrecio.ListIndex = 0
    
'    CargoComboBox cboTarjeta, "TARJETA", "TAR_CODIGO", "TAR_DESCRI", "TAR_DESCRI"
'    If cboTarjeta.ListCount > 0 Then cboTarjeta.ListIndex = 0
    
    'FrameFactura.Enabled = False
    txtNroSucursal_LostFocus
    txtNroFactura_LostFocus
'    txtcodCli.Text = "1"
'    txtCodCli_LostFocus
    lblConPago.Caption = ""
'    cboCondicion_LostFocus
    
    txtPorcentajeIva.Text = Format(CStr(mIVA_1), "0.0000")
    
    txtsubtotal1.Text = "0,000"
    txtimpuesto.Text = "0,000"
    txtSubtotal.Text = "0,000"
    txtImporteIva.Text = "0,000"
    txtnoinsc.Text = "0,000"
    txtTotal.Text = "0,000"
    txttasavial.Text = "0,000"

    
    cmdImprimir.Enabled = False
    
    grdPagos.FormatString = "^Forma Pago|^Importe|Cod.Forma Pago|Cod.Tarjeta|Desc.Tarjeta|Cod.Plan|Desc.Plan|Cupon|Lote|Autorizacion|Dolares|Cotizacion|SeniaSuc|SeniaTipo|SeniaNro"
    grdPagos.ColWidth(0) = 2000    'forma pago
    grdPagos.ColWidth(1) = 1000    'importe
    grdPagos.ColWidth(2) = 0       'cod forma pago
    grdPagos.ColWidth(3) = 0       'cod tarjeta
    grdPagos.ColWidth(4) = 2000    'desc tarjeta
    grdPagos.ColWidth(5) = 0       'cod plan
    grdPagos.ColWidth(6) = 1000    'desc plan
    grdPagos.ColWidth(7) = 1000    'cupon
    grdPagos.ColWidth(8) = 1000    'lote
    grdPagos.ColWidth(9) = 1000    'autorizacion
    grdPagos.ColWidth(10) = 1000   'dolares
    grdPagos.ColWidth(11) = 1000   'cotizacion
    grdPagos.ColWidth(12) = 1000   'seniasuc
    grdPagos.ColWidth(13) = 1000   'seniatipo
    grdPagos.ColWidth(14) = 1000   'senianro
    grdPagos.Rows = 1
    'grdPagos.HighLight = flexHighlightNever
    grdPagos.BorderStyle = flexBorderNone
    grdPagos.row = 0
    For i = 0 To grdPagos.Cols - 1
        grdPagos.Col = i
        grdPagos.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        grdPagos.CellBackColor = &H808080    'GRIS OSCURO
        grdPagos.CellFontBold = True
    Next
    fraPagos.Visible = False
    fraTarjeta.Visible = False
    fracheque.Visible = False
    mFoco = False
    mIVA_1 = 0
    mIVA_2 = 0
    
    'removedata
    
    'limpiar BASE DE DATOS
'    sql = "DELETE FROM CIERREZ"
'    sql = sql & " WHERE Z_FECHA<=" & XDQ("31/12/2013")
'    DBConn.Execute sql
'
'    sql = "DELETE FROM FACTURA_CLIENTE"
'    sql = sql & " WHERE FCL_FECHA<=" & XDQ("31/12/2013")
'    DBConn.Execute sql
'
'    sql = "DELETE FROM DETALLE_FACTURA_CLIENTE"
'    sql = sql & " WHERE TCO_CODIGO=1 AND FCL_NUMERO<=" & XN("31379")
'    DBConn.Execute sql
'
'    sql = "DELETE FROM DETALLE_FACTURA_CLIENTE"
'    sql = sql & " WHERE TCO_CODIGO=2 AND FCL_NUMERO<=" & XN("59778")
'    DBConn.Execute sql
'
'     sql = "DELETE FROM FACTURA_PAGOS"
'    sql = sql & " WHERE TCO_CODIGO=1 AND FCL_NUMERO<=" & XN("31379")
'    DBConn.Execute sql
'
'    sql = "DELETE FROM FACTURA_PAGOS"
'    sql = sql & " WHERE TCO_CODIGO=2 AND FCL_NUMERO<=" & XN("59778")
'    DBConn.Execute sql
'
'    sql = "DELETE FROM T_STOCK"
'    sql = sql & " WHERE T_FECHA<=" & XDQ("31/12/2013")
'    DBConn.Execute sql
'
''    CIERREZ Z_FECHA
''    DETALLE_FACTURA_CLIENTE FCL_NUMERO
''    FACTURA_CLIENTE FCL_FECHA
''    FACTURA_PAGOS FCL_NUMERO
''    T_STOCK T_FECHA

    
    
    
    
    
    'correr este script
   ' sql = "UPDATE FACTURA_CLIENTE SET FCL_IVA = 21"
   ' DBConn.Execute sql


'    sql = "SELECT * FROM FACTURA_CLIENTE"
'    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    If rec.EOF = False Then
'        Do While rec.EOF = False
'            If rec!TCO_CODIGO = 2 Then
'                sql = "UPDATE FACTURA_CLIENTE SET FCL_SUBTOTAL = " & XN(rec!FCL_TOTAL)
'                sql = sql & " WHERE FCL_NUMERO = " & rec!FCL_NUMERO
'
'                DBConn.Execute sql
'            End If
'
'        rec.MoveNext
'        Loop
'    End If
'    rec.Close


'' ACTUALIZO EL NUEVO CAMPO FCL_IMPINT CREADO PARA EL LIBRO DE IVA
'    sql = "SELECT FC.*,DF.DFC_CANTIDAD,DF.PTO_CODIGO,DF.DFC_IMP "
'    sql = sql & " FROM FACTURA_CLIENTE FC,DETALLE_FACTURA_CLIENTE DF"
'    sql = sql & " WHERE FC.TCO_CODIGO = DF.TCO_CODIGO"
'    sql = sql & " AND FC.FCL_NUMERO = DF.FCL_NUMERO"
'    sql = sql & " AND FC.FCL_SUCURSAL = DF.FCL_SUCURSAL"
'    'sql = sql & " AND P.PTO_CODIGO = DF.PTO_CODIGO"
'
'    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    If rec.EOF = False Then
'        Do While rec.EOF = False
'            If rec!PTO_CODIGO = 1 Or rec!PTO_CODIGO = 2 Or rec!PTO_CODIGO = 3 Or rec!PTO_CODIGO = 4 Then
'                Dim ImpInterno As String
'                Dim ImpIVA As String
'                Dim ImpSubtotal As String
'
'                ImpInterno = VALIDO_IMPORTE4(rec!DFC_CANTIDAD * rec!DFC_IMP)
'                ImpSubtotal = rec!FCL_TOTAL - CDbl(ImpInterno)
'                ImpSubtotal = VALIDO_IMPORTE4((ImpSubtotal) / 1.21)
'                ImpIVA = VALIDO_IMPORTE4(rec!FCL_TOTAL - CDbl(ImpInterno) - CDbl(ImpSubtotal))
'
'                ImpInterno = XN(ImpInterno)
'                ImpSubtotal = XN(ImpSubtotal)
'                ImpIVA = XN(ImpIVA)
'
'                sql = "UPDATE FACTURA_CLIENTE SET FCL_IMPINT = " & ImpInterno
'                sql = sql & " ,FCL_SUBTOTAL=" & ImpSubtotal
'                sql = sql & " ,FCL_IMPIVA=" & ImpIVA
'                sql = sql & " WHERE FCL_NUMERO = " & rec!FCL_NUMERO
'                sql = sql & " AND FCL_SUCURSAL = " & rec!FCL_SUCURSAL
'                sql = sql & " AND TCO_CODIGO = " & rec!TCO_CODIGO
'
'                    DBConn.Execute sql
'            Else
'            'ACTUALIZAR SOLO EL IVA
'
'                sql = "UPDATE FACTURA_CLIENTE SET FCL_IMPIVA = " & XN(rec!FCL_TOTAL * 21 / 100)
'
'                sql = sql & " WHERE FCL_NUMERO = " & rec!FCL_NUMERO
'                sql = sql & " AND FCL_SUCURSAL = " & rec!FCL_SUCURSAL
'                sql = sql & " AND TCO_CODIGO = " & rec!TCO_CODIGO
'
'                DBConn.Execute sql
'            End If
'
'        rec.MoveNext
'        Loop
'    End If
'    rec.Close

'' ACTUALIZO TASA VIAL EN FACTURA_CLIENTE TOMANDO DESDE DETALLE_FACTURA_CLIENTE
'    sql = "SELECT DF.TCO_CODIGO,DF.FCL_NUMERO,DF.FCL_SUCURSAL, SUM(DFC_TotalTVial) AS TASAVIAL "
'    sql = sql & " FROM FACTURA_CLIENTE FC,DETALLE_FACTURA_CLIENTE DF"
'    sql = sql & " WHERE FC.TCO_CODIGO = DF.TCO_CODIGO"
'    sql = sql & " AND FC.FCL_NUMERO = DF.FCL_NUMERO"
'    sql = sql & " AND FC.FCL_SUCURSAL = DF.FCL_SUCURSAL"
'    sql = sql & " AND FC.FCL_FECHA >= #07/09/2012#"
'    'sql = sql & " AND P.PTO_CODIGO = DF.PTO_CODIGO"
'    sql = sql & " GROUP BY DF.TCO_CODIGO,DF.FCL_NUMERO,DF.FCL_SUCURSAL"
'
'    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    If rec.EOF = False Then
'        Do While rec.EOF = False
'            'ACTUALIZAR SOLO EL IVA
'            If Not IsNull(rec!tasavial) Then
'
'                sql = "UPDATE FACTURA_CLIENTE SET FCL_TASAVIAL = " & XN(rec!tasavial)
'
'                sql = sql & " WHERE FCL_NUMERO = " & rec!FCL_NUMERO
'                sql = sql & " AND FCL_SUCURSAL = " & rec!FCL_SUCURSAL
'                sql = sql & " AND TCO_CODIGO = " & rec!TCO_CODIGO
'
'                DBConn.Execute sql
'            End If
'
'            rec.MoveNext
'        Loop
'    End If
'    rec.Close
'
' 'PONGO EN CERO LOS QUE TIENEN FCL_IMPINT NULL
'    sql = "UPDATE FACTURA_CLIENTE SET FCL_IMPINT = 0"
'    sql = sql & " WHERE FCL_IMPINT IS NULL "
'    DBConn.Execute sql
'
    
    'CORRER SCRIPT QUE ACTUALICE TURNOS DE ACUERDO A LA HORA REGISTRADA, SINO
    ' HAY HORA PONER A LA TARDE



''' ACTUALIZO NETO E iva EN FAC DE VARIOSS O ALMUERZOS DEL MES DE JULIO 2013 - FACTURAS A
'    Dim vSubtotal As Double
'    Dim vIVA As Double
'
'        sql = "SELECT DF.TCO_CODIGO,DF.FCL_NUMERO,DF.FCL_SUCURSAL,FC.FCL_FECHA,FC.FCL_TOTAL,FC.FCL_SUBTOTAL,FC.FCL_IMPIVA "
'    sql = sql & " FROM FACTURA_CLIENTE FC,DETALLE_FACTURA_CLIENTE DF"
'    sql = sql & " WHERE FC.TCO_CODIGO = DF.TCO_CODIGO"
'    sql = sql & " AND FC.FCL_NUMERO = DF.FCL_NUMERO"
'    sql = sql & " AND FC.FCL_SUCURSAL = DF.FCL_SUCURSAL"
'    sql = sql & " AND FC.FCL_NUMERO >= 24382"
'    sql = sql & " AND FC.TCO_CODIGO >= 1"
'    sql = sql & " AND DF.PTO_CODIGO = 69"
'
'
'    'sql = sql & " AND P.PTO_CODIGO = DF.PTO_CODIGO"
'    sql = sql & " ORDER BY DF.TCO_CODIGO,DF.FCL_NUMERO,DF.FCL_SUCURSAL"
'
'    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    If rec.EOF = False Then
'        Do While rec.EOF = False
'            'ACTUALIZAR SOLO EL IVA
'            MsgBox rec!TCO_CODIGO & " - " & rec!FCL_NUMERO & " - " & rec!FCL_FECHA
'            vSubtotal = rec!FCL_TOTAL / 1.21
'            vIVA = rec!FCL_TOTAL - vSubtotal
'
'            If Not IsNull(rec!tasavial) Then
'
'                sql = "UPDATE FACTURA_CLIENTE SET FCL_SUBTOTAL = " & vSubtotal
'                sql = sql & " ,FCL_IVA = " & vIVA
'                sql = sql & " WHERE FCL_NUMERO = " & rec!FCL_NUMERO
'                sql = sql & " AND FCL_SUCURSAL = " & rec!FCL_SUCURSAL
'                sql = sql & " AND TCO_CODIGO = " & rec!TCO_CODIGO
'
'                DBConn.Execute sql
'            End If
'
'            rec.MoveNext
'        Loop
'    End If
'    rec.Close
'
'    '' ACTUALIZO NETO E iva EN FAC DE VARIOSS O ALMUERZOS DEL MES DE JULIO 2013 - FACTURAS B
''    Dim vSubtotal As Double
''    Dim vIVA As Double
'
'    sql = "SELECT DF.TCO_CODIGO,DF.FCL_NUMERO,DF.FCL_SUCURSAL,FC.FCL_FECHA,FC.FCL_TOTAL,FC.FCL_SUBTOTAL,FC.FCL_IMPIVA "
'    sql = sql & " FROM FACTURA_CLIENTE FC,DETALLE_FACTURA_CLIENTE DF"
'    sql = sql & " WHERE FC.TCO_CODIGO = DF.TCO_CODIGO"
'    sql = sql & " AND FC.FCL_NUMERO = DF.FCL_NUMERO"
'    sql = sql & " AND FC.FCL_SUCURSAL = DF.FCL_SUCURSAL"
'    sql = sql & " AND FC.FCL_NUMERO >= 51824"
'    sql = sql & " AND FC.TCO_CODIGO >= 2"
'    sql = sql & " AND DF.PTO_CODIGO = 69"
'
'
'    'sql = sql & " AND P.PTO_CODIGO = DF.PTO_CODIGO"
'    sql = sql & " ORDER BY DF.TCO_CODIGO,DF.FCL_NUMERO,DF.FCL_SUCURSAL"
'
'    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    If rec.EOF = False Then
'        Do While rec.EOF = False
'            'ACTUALIZAR SOLO EL IVA
'            MsgBox rec!TCO_CODIGO & " - " & rec!FCL_NUMERO & " - " & rec!FCL_FECHA
'            vSubtotal = rec!FCL_TOTAL / 1.21
'            vIVA = rec!FCL_TOTAL - vSubtotal
'
'            If Not IsNull(rec!tasavial) Then
'
'                sql = "UPDATE FACTURA_CLIENTE SET FCL_SUBTOTAL = " & vSubtotal
'                sql = sql & " ,FCL_IVA = " & vIVA
'                sql = sql & " WHERE FCL_NUMERO = " & rec!FCL_NUMERO
'                sql = sql & " AND FCL_SUCURSAL = " & rec!FCL_SUCURSAL
'                sql = sql & " AND TCO_CODIGO = " & rec!TCO_CODIGO
'
'                DBConn.Execute sql
'            End If
'
'            rec.MoveNext
'        Loop
'    End If
'    rec.Close
'
'
End Sub

Private Sub LlenarComboFormaPago()
    sql = "SELECT FPG_DESCRI,FPG_CODIGO FROM FORMA_PAGO"
    sql = sql & " ORDER BY FPG_CODIGO"
    If rec.State = 1 Then rec.Close
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            If lblblockeado.Visible = True And rec!FPG_CODIGO = 2 Then 'si cliente esta bloqueado no cargar cta cte
                MsgBox "La opcion de Cuenta Corriente NO esta disponible para este Cliente", vbExclamation, TIT_MSGBOX
            Else
                cboCondicion.AddItem rec!FPG_DESCRI
                cboCondicion.ItemData(cboCondicion.NewIndex) = rec!FPG_CODIGO
                cboFormaPago.AddItem rec!FPG_DESCRI
                cboFormaPago.ItemData(cboFormaPago.NewIndex) = rec!FPG_CODIGO
            End If
            rec.MoveNext
            
        Loop
        cboCondicion.ListIndex = 0
        cboFormaPago.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub LlenarComboVendedor()
    sql = "SELECT VEN_NOMBRE,VEN_CODIGO FROM VENDEDOR WHERE VEN_ESTADO=" & XS("N")
    sql = sql & " ORDER BY VEN_NOMBRE"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboVendedor.AddItem ""
        Do While rec.EOF = False
            cboVendedor.AddItem rec!VEN_NOMBRE
            cboVendedor.ItemData(cboVendedor.NewIndex) = rec!VEN_CODIGO
            rec.MoveNext
        Loop
        If cboVendedor.ListCount > 0 Then cboVendedor.ListIndex = 0
    End If
    rec.Close
End Sub
Private Sub LlenarComboTurnos()
    Dim vDesde(3) As Date
    Dim vHasta(3) As Date
    Dim i As Integer
    sql = "SELECT * FROM TURNOS"
    sql = sql & " ORDER BY TUR_CODIGO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    i = 0
    If rec.EOF = False Then
        cboTurno.AddItem ""
        cboTurnosB.AddItem ""
        Do While rec.EOF = False
            cboTurno.AddItem rec!TUR_DESCRI
            cboTurno.ItemData(cboTurno.NewIndex) = rec!TUR_CODIGO
            
            cboTurnosB.AddItem rec!TUR_DESCRI
            cboTurnosB.ItemData(cboTurno.NewIndex) = rec!TUR_CODIGO
            vDesde(i) = rec!TUR_DESDE
            vHasta(i) = rec!TUR_HASTA
            i = i + 1
            rec.MoveNext
        Loop
    End If
    rec.Close
    'POSICIONO EL TURNO DE ACUERDO A LA HORA ACTUAL
    'For i = 0 To 2
        If Time() >= vDesde(0) And Time() <= vHasta(0) Then
            Call BuscaCodigoProxItemData(1, cboTurno)
            Call BuscaCodigoProxItemData(1, cboTurnosB)
        Else
            If Time() >= vDesde(1) And Time() <= vHasta(1) Then
                Call BuscaCodigoProxItemData(2, cboTurno)
                Call BuscaCodigoProxItemData(2, cboTurnosB)
            Else
                Call BuscaCodigoProxItemData(3, cboTurno)
                Call BuscaCodigoProxItemData(3, cboTurnosB)
            End If
        End If
    'Next i
    cboTurnosB.ListIndex = 0
End Sub
Private Function BuscoIva() As Double
    sql = "SELECT IVA FROM PARAMETROS"
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        BuscoIva = Chk0(Rec1!iva)
        'txtPorcentajeIva.Text = "0,00" 'IIf(IsNull(rec!IVA), "", Format(rec!IVA, "0.00"))
    End If
    Rec1.Close
End Function
Private Function BuscoIva_2() As Double
    sql = "SELECT IVA_2 FROM PARAMETROS"
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        BuscoIva_2 = Chk0(Rec1!IVA_2)
        'txtPorcentajeIva.Text = "0,00" 'IIf(IsNull(rec!IVA), "", Format(rec!IVA, "0.00"))
    End If
    Rec1.Close
End Function
Private Sub LlenarComboFactura()
    sql = "SELECT * FROM TIPO_COMPROBANTE"
    sql = sql & " WHERE TCO_DESCRI LIKE 'FACTURA%'"
    sql = sql & " ORDER BY TCO_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboFactura1.AddItem "(Todas)"
        Do While rec.EOF = False
            cboFactura.AddItem rec!TCO_DESCRI
            cboFactura.ItemData(cboFactura.NewIndex) = rec!TCO_CODIGO
            cboFactura1.AddItem rec!TCO_DESCRI
            cboFactura1.ItemData(cboFactura1.NewIndex) = rec!TCO_CODIGO
            rec.MoveNext
        Loop
        cboFactura.ListIndex = 0
        cboFactura1.ListIndex = 0
    End If
    rec.Close
End Sub

Private Function BuscoUltimaFactura(TipoFac As Integer) As String
    'ACA BUSCA EL NUMERO DE REMITO SIGUIENTE AL ULTIMO CARGADO
    sql = "SELECT (FACTURA_C) + 1 AS FAC_C"
    sql = sql & " FROM PARAMETROS"
    If rec.State = 1 Then rec.Close
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Select Case TipoFac
            Case 3
                BuscoUltimaFactura = IIf(IsNull(rec!FAC_C), 1, rec!FAC_C)
        End Select
    End If
    rec.Close
End Function

Private Sub Form_Unload(Cancel As Integer)
    FormLlamado = ""
End Sub

Private Sub grdGrilla_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Select Case grdGrilla.Col
        Case 0, 1
            LimpiarFilasDeGrilla grdGrilla, grdGrilla.RowSel
            grdGrilla.Col = 0
            txtSubtotal.Text = VALIDO_IMPORTE4(CStr(SumaTotal))
            txtTotal.Text = Format(CStr(SumaTotal), "0.00")
            txtPorcentajeIva_LostFocus
            Sumatorias
        End Select
    End If
    If KeyCode = vbKeyReturn Then
        Select Case grdGrilla.Col
            Case 1
                If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) = "" And grdGrilla.TextMatrix(grdGrilla.RowSel, 1) = "" Then
                    cmdFormaPago.SetFocus
                End If
        End Select
    End If
    If KeyCode = vbKeyF1 Then
        BuscarProducto grdGrilla, "CODIGO", , grdGrilla.RowSel
    End If
End Sub

Private Sub grdGrilla_KeyPress(KeyAscii As Integer)
    If (grdGrilla.Col = 0) Or (grdGrilla.Col = 1) Or (grdGrilla.Col = 2) Or (grdGrilla.Col = 4) Then  'Or (grdGrilla.Col = 3)
        If KeyAscii = vbKeyReturn Then
            If grdGrilla.Col = 3 Or grdGrilla.Col = 4 Then '2
                If grdGrilla.row < grdGrilla.Rows - 1 Then
                    grdGrilla.row = grdGrilla.row + 1
                    grdGrilla.Col = 0
                Else
                    MySendKeys Chr(9)
                End If
            Else
                grdGrilla.Col = grdGrilla.Col + 1
            End If
        Else
            If grdGrilla.Col = 2 Or grdGrilla.Col = 4 Then  'grdGrilla.Col = 0 Or Or grdGrilla.Col = 3
                If KeyAscii > 47 And KeyAscii < 58 Then
                    EDITAR grdGrilla, txtEdit, KeyAscii
                End If
            ElseIf grdGrilla.Col = 1 Or grdGrilla.Col = 0 Then
                EDITAR grdGrilla, txtEdit, KeyAscii
            End If
        End If
    Else
        If grdGrilla.Col = 3 Then
            If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) > 66 And grdGrilla.TextMatrix(grdGrilla.RowSel, 0) <= 71 Then
                If KeyAscii = vbKeyReturn Then
                    If grdGrilla.row < grdGrilla.Rows - 1 Then
                        grdGrilla.row = grdGrilla.row + 1
                        grdGrilla.Col = 0
                    Else
                        MySendKeys Chr(9)
                    End If
                Else
                    If KeyAscii > 47 And KeyAscii < 58 Then
                        EDITAR grdGrilla, txtEdit, KeyAscii
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub grdGrilla_LeaveCell()
    If txtEdit.Visible = False Then Exit Sub
    grdGrilla = txtEdit.Text
    txtEdit.Visible = False
End Sub

Private Sub grdGrilla_GotFocus()
    If grdGrilla.Rows > 1 Then
        If txtEdit.Visible = False And mFoco = False Then
            grdGrilla.Col = 0
            grdGrilla.row = 1
            Exit Sub
        End If
        grdGrilla = txtEdit.Text
        txtEdit.Visible = False
        mFoco = False
    End If
End Sub

Private Sub GrdModulos_dblClick()
    Dim mimpneto As Double
    If GrdModulos.Rows > 1 Then
        Set Rec1 = New ADODB.Recordset
        lblEstado.Caption = "Buscando..."
        Screen.MousePointer = vbHourglass
        'CABEZA FACTURA
        'tengo que limpiar
        CmdNuevo_Click
        cmdGrabar.Enabled = False
        mBuscador = True
        tabDatos.Tab = 0
        
        Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 7)), cboFactura)
        txtNroSucursal.Text = Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4)
        txtNroFactura.Text = Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 8)
        FechaFactura.Value = GrdModulos.TextMatrix(GrdModulos.RowSel, 2)
        Call BuscoEstado(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 4)), lblEstadoFactura)
        txtcodCli.Text = Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 9))
        mVerCta = False
        txtCodCli_LostFocus
        mVerCta = True
        
        If GrdModulos.TextMatrix(GrdModulos.RowSel, 6) <> "" Then
            txtObservaciones.Text = Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 6))
        End If
        'CONDICION VENTA
        Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 8)), cboCondicion)
        Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 12)), cboVendedor)
        Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 13)), cboTurno)
        
        '----BUSCO DETALLE DE LA FACTURA------------------
        sql = "SELECT P.PTO_CODIGO, DFC.DFC_CANTIDAD, DFC.DFC_PRECIO, P.PTO_DESCRI,P.PTO_CODBARRAS, DFC.DFC_CONCEPTO"
        sql = sql & " ,DFC.DFC_IVA, DFC.DFC_IMP, DFC.DFC_MONIMP, DFC.DFC_MONIVA,DFC.TCO_CODIGO,DFC_TASAVIAL, DFC_TOTALTVIAL"
        sql = sql & " FROM DETALLE_FACTURA_CLIENTE DFC, PRODUCTO  P"
        sql = sql & " WHERE DFC.FCL_NUMERO=" & XN(Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 8))
        sql = sql & " AND DFC.FCL_SUCURSAL=" & XN(Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4))
        sql = sql & " AND DFC.TCO_CODIGO=" & XN(GrdModulos.TextMatrix(GrdModulos.RowSel, 7))
        sql = sql & " AND DFC.PTO_CODIGO=P.PTO_CODIGO"
        sql = sql & " ORDER BY DFC.DFC_NROITEM"
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            i = 1
            Do While Rec1.EOF = False
                grdGrilla.TextMatrix(i, 0) = IIf(IsNull(Rec1!PTO_CODBARRAS), Rec1!PTO_CODIGO, Trim(Rec1!PTO_CODBARRAS))
                grdGrilla.TextMatrix(i, 1) = IIf(IsNull(Rec1!DFC_CONCEPTO), Trim(ChkNull(Rec1!PTO_DESCRI)), Trim(ChkNull(Rec1!DFC_CONCEPTO)))
                grdGrilla.TextMatrix(i, 2) = Chk0(Rec1!DFC_CANTIDAD)
                grdGrilla.TextMatrix(i, 3) = VALIDO_IMPORTE4(Chk0(Rec1!DFC_PRECIO))
                
                grdGrilla.TextMatrix(i, 6) = VALIDO_IMPORTE4(Chk0(Rec1!DFC_IVA))
                grdGrilla.TextMatrix(i, 7) = VALIDO_IMPORTE4(Chk0(Rec1!DFC_IMP))
                grdGrilla.TextMatrix(i, 8) = Format(Chk0(Rec1!DFC_MONIMP), "0.0000")
                grdGrilla.TextMatrix(i, 9) = Format(Chk0(Rec1!DFC_MONIVA), "0.0000")
                grdGrilla.TextMatrix(i, 11) = Format(Chk0(Rec1!DFC_TasaVial), "0.0000")
                grdGrilla.TextMatrix(i, 12) = Format(Chk0(Rec1!DFC_TOTALTVIAL), "0.0000")
                'aca tengo que ver si es boleta a y poner el neto sin iva
                If Rec1!TCO_CODIGO = 1 Then
                    mimpneto = CDbl(grdGrilla.TextMatrix(i, 2)) * CDbl(grdGrilla.TextMatrix(i, 3)) - (CDbl(grdGrilla.TextMatrix(i, 2)) * CDbl(grdGrilla.TextMatrix(i, 7)))
                    mimpneto = mimpneto / (1 + (Chk0(Rec1!DFC_IVA) / 100))
                
                    grdGrilla.TextMatrix(i, 4) = Format(mimpneto, "0.0000")
                    
                Else
                    grdGrilla.TextMatrix(i, 4) = Format(CDbl(grdGrilla.TextMatrix(i, 2)) * CDbl(grdGrilla.TextMatrix(i, 3)), "0.0000")
                End If
                'En esta columna auxiliar pongo el neto sin iva para q se guarde en la BD cuando es Fac B tmb
                grdGrilla.TextMatrix(i, 10) = Format(mimpneto, "0.0000")
                i = i + 1
                Rec1.MoveNext
            Loop
        End If
        Rec1.Close
        '--CARGO LOS TOTALES----
        '----BUSCO  LA FACTURA------------------
        sql = "SELECT TCO_CODIGO, FCL_SUCURSAL, FCL_NUMERO, FCL_SUBTOTAL, FCL_IMPIVA,FCL_IMPINT, FCL_TASAVIAL, FCL_TOTAL"
        sql = sql & " FROM FACTURA_CLIENTE "
        sql = sql & " WHERE FCL_NUMERO=" & XN(Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 8))
        sql = sql & " AND FCL_SUCURSAL=" & XN(Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4))
        sql = sql & " AND TCO_CODIGO=" & XN(GrdModulos.TextMatrix(GrdModulos.RowSel, 7))
        
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            If Rec1!TCO_CODIGO = 1 Then 'FAC A
                txtTotal.Text = Chk0(Rec1!FCL_TOTAL)
                txtTotal.Text = VALIDO_IMPORTE4(txtTotal.Text)
                
                txtsubtotal1.Text = Chk0(Rec1!FCL_SUBTOTAL)
                txtsubtotal1.Text = VALIDO_IMPORTE4(txtsubtotal1.Text)
                
                txtimpuesto.Text = Chk0(Rec1!FCL_IMPINT)
                txtimpuesto.Text = VALIDO_IMPORTE4(txtimpuesto.Text)
                
                txtSubtotal.Text = CDbl(txtsubtotal1.Text) + CDbl(txtimpuesto.Text)
                txtSubtotal.Text = VALIDO_IMPORTE4(txtSubtotal.Text)
                
                txtImporteIva.Text = Chk0(Rec1!FCL_IMPIVA)
                txtImporteIva.Text = VALIDO_IMPORTE4(txtImporteIva.Text)
                
                txttasavial.Text = Chk0(Rec1!FCL_TASAVIAL)
                txttasavial.Text = VALIDO_IMPORTE4(txttasavial.Text)
                'OCULTOS
                txtsubtotal1B.Text = Chk0(Rec1!FCL_SUBTOTAL)
                
                txtimpuestoB.Text = Chk0(Rec1!FCL_IMPINT)
                txtimpuestoB.Text = VALIDO_IMPORTE4(txtimpuestoB.Text)
                
                txtSubtotalB.Text = CDbl(txtsubtotal1B.Text) + CDbl(txtimpuestoB.Text)
                txtSubtotalB.Text = VALIDO_IMPORTE4(txtSubtotalB.Text)
                
                txtImporteIvaB.Text = Chk0(Rec1!FCL_IMPIVA)
                
            Else 'FAC B
                txtTotal.Text = Chk0(Rec1!FCL_TOTAL)
                txtTotal.Text = VALIDO_IMPORTE4(txtTotal.Text)
                                
                txtsubtotal1.Text = Chk0(Rec1!FCL_TOTAL)
                txtsubtotal1.Text = VALIDO_IMPORTE4(txtsubtotal1.Text)
                
                txtSubtotal.Text = Chk0(Rec1!FCL_TOTAL)
                txtSubtotal.Text = VALIDO_IMPORTE4(txtSubtotal.Text)
                                
                txttasavial.Text = Chk0(Rec1!FCL_TASAVIAL)
                txttasavial.Text = VALIDO_IMPORTE4(txttasavial.Text)
                
                'OCULTOS
                txtsubtotal1B.Text = Chk0(Rec1!FCL_SUBTOTAL)
                txtsubtotal1B.Text = VALIDO_IMPORTE4(txtsubtotal1B.Text)
                
                txtimpuestoB.Text = Chk0(Rec1!FCL_IMPINT)
                txtimpuestoB.Text = VALIDO_IMPORTE4(txtimpuestoB.Text)
                
                txtSubtotalB.Text = CDbl(txtsubtotal1B.Text) + CDbl(txtimpuestoB.Text)
                txtSubtotalB.Text = VALIDO_IMPORTE4(txtSubtotalB.Text)
                
                txtImporteIvaB.Text = Chk0(Rec1!FCL_IMPIVA)
                
                
            End If
        End If
        Rec1.Close
        
        
        
        'txtSubtotal.Text = Format(CStr(SumaSUBTotal), "0.00")
        'calculototales grdGrilla.TextMatrix(I, 7), grdGrilla.TextMatrix(I, 3)
        
        'txtTotal.Text = txtSubtotal.Text
        'txtTotal.Text = Format(GrdModulos.TextMatrix(GrdModulos.RowSel, 10), "0.00")
        'txtImporteIva.Text = Format(GrdModulos.TextMatrix(GrdModulos.RowSel, 11), "0.0000")
        'If GrdModulos.TextMatrix(GrdModulos.RowSel, 5) <> "" Then
        'txtPorcentajeIva = GrdModulos.TextMatrix(GrdModulos.RowSel, 5)
        'suma de totales con impuestos
        'Sumatorias
        'txtsubtotal1.Text = CDbl(txtsubtotal1.Text) - CDbl(txttasavial.Text)
        'txtsubtotal1.Text = VALIDO_IMPORTE4(txtsubtotal1.Text)
        '    txtPorcentajeIva_LostFocus
        'End If
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        '--------------
        'FrameFactura.Enabled = False
        FrameCliente.Enabled = False
        '--------------
        'tabDatos.Tab = 0
        'cboCondicion.SetFocus
        cmdGrabar.Enabled = False
        cmdFormaPago.Enabled = False
        cmdImprimir.Enabled = True
    '----------------------------------------------------------
        lblEstado.Caption = "Buscando..."
        Screen.MousePointer = vbHourglass
    
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        'tabDatos.Tab = 0
        
        cmdNC.Enabled = True
    End If
End Sub

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then GrdModulos_dblClick
End Sub

Private Sub tabDatos_Click(PreviousTab As Integer)
  If tabDatos.Tab = 1 Then
    'LimpiarBusqueda
    If Me.Visible = True Then txtBuscaCliente.SetFocus
    frameBuscar.Caption = "Buscar Factura por..."
  Else
    If VEstadoFactura = 1 Then
        cmdGrabar.Enabled = True
        cmdFormaPago.Enabled = True
    Else
        cmdGrabar.Enabled = False
        cmdFormaPago.Enabled = False
    End If
  End If
End Sub

Private Sub LimpiarBusqueda()
    txtBuscaCliente.Text = ""
    txtBuscarCliDescri.Text = ""
    FechaDesde.Value = ""
    FechaHasta.Value = ""
    cboFactura1.ListIndex = 0
    GrdModulos.HighLight = flexHighlightNever
    GrdModulos.Rows = 1
End Sub

Private Function BuscoCondicionIVA(IVACodigo As String) As String
    sql = "SELECT * FROM CONDICION_IVA"
    sql = sql & " WHERE IVA_CODIGO=" & XN(IVACodigo)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        BuscoCondicionIVA = rec!IVA_DESCRI
    Else
        BuscoCondicionIVA = ""
    End If
    rec.Close
End Function

Private Sub txtBuscaCliente_Change()
    If txtBuscaCliente.Text = "" Then
        txtBuscarCliDescri.Text = ""
    End If
End Sub

Private Sub txtBuscaCliente_GotFocus()
    SelecTexto txtBuscaCliente
End Sub

Private Sub txtBuscaCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarClientes "txtBuscaCliente", "CODIGO"
    End If
End Sub

Private Sub txtBuscaCliente_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtBuscaCliente_LostFocus()
    If txtBuscaCliente.Text <> "" Then
        Set rec = New ADODB.Recordset
        sql = "SELECT CLI_CODIGO, CLI_RAZSOC"
        sql = sql & " FROM CLIENTE"
        sql = sql & " WHERE "
        If txtBuscaCliente.Text <> "" Then
            sql = sql & " CLI_CODIGO=" & XN(txtBuscaCliente)
        Else
            sql = sql & " CLI_RAZSOC LIKE '" & Trim(txtBuscaCliente) & "%'"
        End If
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtBuscarCliDescri.Text = rec!CLI_RAZSOC
        Else
            MsgBox "El Cliente no existe", vbExclamation, TIT_MSGBOX
            txtBuscaCliente.SetFocus
        End If
        rec.Close
    End If
End Sub

Private Sub txtBuscaNum_GotFocus()
    seltxt
End Sub

Private Sub txtBuscaNum_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtBuscaNum_LostFocus()
    If txtNroFactura.Text <> "" Then
        txtNroFactura.Text = Format(txtNroFactura.Text, "00000000")
    End If

End Sub

Private Sub txtBuscarCliDescri_Change()
    If txtBuscarCliDescri.Text = "" Then
        txtBuscaCliente.Text = ""
    End If
End Sub

Private Sub txtBuscarCliDescri_GotFocus()
    SelecTexto txtBuscarCliDescri
End Sub

Private Sub txtBuscarCliDescri_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarClientes "txtBuscaCliente", "CODIGO"
    End If
End Sub

Private Sub txtBuscarCliDescri_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtBuscarCliDescri_LostFocus()
    If txtBuscaCliente.Text = "" And txtBuscarCliDescri.Text <> "" Then
        Set rec = New ADODB.Recordset
        sql = "SELECT CLI_CODIGO, CLI_RAZSOC"
        sql = sql & " FROM CLIENTE"
        sql = sql & " WHERE "
        If txtBuscaCliente.Text <> "" Then
            sql = sql & " CLI_CODIGO=" & XN(txtBuscaCliente)
        Else
            sql = sql & " CLI_RAZSOC LIKE '" & Trim(txtBuscarCliDescri) & "%'"
        End If
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If rec.RecordCount > 1 Then
                BuscarClientes "txtBuscaCliente", "CADENA", Trim(txtBuscarCliDescri.Text)
                If rec.State = 1 Then rec.Close
                txtBuscarCliDescri.SetFocus
            Else
                txtBuscaCliente.Text = rec!CLI_CODIGO
                txtBuscarCliDescri.Text = rec!CLI_RAZSOC
            End If
        Else
            MsgBox "No se encontro el Cliente", vbExclamation, TIT_MSGBOX
            txtBuscaCliente.SetFocus
        End If
        If rec.State = 1 Then rec.Close
    End If
End Sub

Private Sub txtBuscaSuc_GotFocus()
    seltxt
End Sub

Private Sub txtBuscaSuc_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtBuscaSuc_LostFocus()
    If txtBuscaSuc.Text = "" Then
        txtBuscaSuc.Text = Format(txtBuscaSuc, "0000")
    Else
        txtBuscaSuc.Text = Format(txtBuscaSuc.Text, "0000")
    End If

End Sub

Private Sub txtcodCli_Change()
    If txtcodCli.Text = "" Then
        txtRazSoc.Text = ""
        txtDomici.Text = ""
        txtCuit.Text = ""
        txtCiva.Text = ""
        txtNRO_DOCUMENTO.Text = ""
        txtTelefono.Text = ""
        txtIngBrutos.Text = ""
        mRespo.Text = ""
        'LIMPIOGRILLA
    End If
End Sub

Private Sub txtcodCli_GotFocus()
    SelecTexto txtcodCli
End Sub

Private Sub txtcodCli_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        txtcodCli.Text = ""
        BuscarClientes "txtcodCli", "CODIGO"
    End If
End Sub

Private Sub txtcodCli_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCodCli_LostFocus()
    If ActiveControl.Name = "cmdGrabar" Then Exit Sub
    If txtcodCli.Text <> "" Then
        sql = "SELECT C.CLI_CODIGO,C.CLI_RAZSOC,C.CLI_DOMICI,I.IVA_CODIGO,I.IVA_DESCRI,"
        sql = sql & "C.CLI_TELEFONO,C.CLI_CUIT,C.CLI_INGBRU, I.IVA_LETRA, C.CLI_NRODOC, C.CLI_CTACTE,C.CLI_BLOCKEADO"
        sql = sql & " FROM CLIENTE C, CONDICION_IVA I"
        sql = sql & " WHERE I.IVA_CODIGO = C.IVA_CODIGO"
        sql = sql & " AND C.CLI_CODIGO =" & XN(txtcodCli.Text)
        If rec.State = 1 Then rec.Close
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If mQuienLlama = "" Then
'                If mBuscador = False Then
'                    LIMPIOGRILLA
'                    txtsubtotal1.Text = "0.0000"
'                    txtSubtotal.Text = "0,000"
'                    txtimpuesto.Text = "0.0000"
'                    txtTotal.Text = "0,000"
'                    txtPorcentajeIva.Text = Format(mIVAi, "0.0000")
'                    txtImporteIva.Text = "0,000"
'                End If
            End If
   
            txtRazSoc.Text = Trim(ChkNull(rec!CLI_RAZSOC))
            txtDomici.Text = Trim(ChkNull(rec!CLI_DOMICI))
            txtCiva.Text = ChkNull(rec!IVA_DESCRI)
            txtCuit.Text = ChkNull(rec!CLI_CUIT)
            txtTelefono.Text = Trim(ChkNull(rec!CLI_TELEFONO))
            txtIngBrutos.Text = Trim(ChkNull(rec!CLI_INGBRU))
            mRespo.Text = ChkNull(rec!IVA_LETRA)
            QueFacturaUso (rec!IVA_CODIGO)
            txtNRO_DOCUMENTO.Text = Trim(ChkNull(rec!CLI_NRODOC))
            
            
            If txtcodCli.Text <> 1 Then
                If Chk0(rec!CLI_CTACTE) > 0 Then 'ojo cambio aca
                    lblSaldoFac.Visible = True
                    'lblSaldoFac = "Saldo Facturacion: $" & BuscarSaldoFactura(rec!CLI_CODIGO, rec!CLI_CTACTE)
                    'SaldoCli = BuscarSaldoFactura(rec!CLI_CODIGO, rec!CLI_CTACTE)
                Else
                    lblSaldoFac.Visible = False
                    
                End If
                If Chk0(rec!CLI_BLOCKEADO) = 1 Then
                    lblblockeado.Visible = True
                    lblblockeado = "CLIENTE CON CUENTA CORRIENTE BLOQUEADA"
                    
                Else
                    lblblockeado.Visible = False
                End If
            End If
            cboCondicion.Clear
            cboFormaPago.Clear
            LlenarComboFormaPago
            If mQuienLlama = "" Then
                If mVerCta = True Then
                    'Call BuscarPendienteClientes(txtcodCli.Text, True, True)
                End If
            End If
            If cmdGrabar.Enabled = True Then
                'BUSCO EL NUMERO DE FACTURA EN EL FISCAL
                If FISCAL = "TMT900FA" Then
                    txtNroFactura.Text = Epson_ConsultarNumeroComprobanteUltimo(cboFactura.ItemData(cboFactura.ListIndex)) + 1
                    txtNroFactura.Text = Format(txtNroFactura.Text, "00000000")
                Else
                    Select Case cboFactura.ItemData(cboFactura.ListIndex)
                        Case 1 'FACTURAS A
                            
                            
                            pf.Status ("A")
                            txtNroFactura.Text = Val(pf.AnswerField_7) + 1
    '                        sql = "SELECT FACTURA_C FROM PARAMETROS"
    '                        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    '
    '                        If Rec1.EOF = False Then
    '                                txtNroFactura.Text = Rec1!FACTURA_C + 1
    '                        End If
    '                        Rec1.Close
                            
                        Case 2 'FACTURA B
                            pf.Status ("A")
                            txtNroFactura.Text = Val(pf.AnswerField_5) + 1
    '                        sql = "SELECT FACTURA_B FROM PARAMETROS"
    '                        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    '
    '                        If Rec1.EOF = False Then
    '                                txtNroFactura.Text = Rec1!FACTURA_B + 1
    '                        End If
    '                        Rec1.Close
                        Case 3 'FACTURA C
                        Case 10000 'PARA TIKET
                            pf.Status ("A")
                            txtNroFactura.Text = Val(pf.AnswerField_4) + 1
                    End Select
               End If
            End If
        Else
            MsgBox "El Código no existe", vbInformation
            txtRazSoc.Text = ""
            txtcodCli.Text = ""
            txtcodCli.SetFocus
        End If
        If rec.State = 1 Then rec.Close
    End If
End Sub

Private Sub QueFacturaUso(iva As Integer)
    Select Case iva
        Case 1 'RESPONSABLE INSCRIPTO
            BuscaProx "FACTURA A", cboFactura
            Ltipo_fac.Caption = "A"
        Case Else ' EL RESTO DE LAS CONDICIONES USA FACTURA B
            BuscaProx "FACTURA B", cboFactura
            Ltipo_fac.Caption = "B"
    End Select
End Sub

Private Sub txtCupon_GotFocus()
    SelecTexto txtCupon
End Sub

Private Sub txtCupon_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(vbCr) Then KeyAscii = 0
    If grdGrilla.Col = 0 Then KeyAscii = CarNumeroEntero(KeyAscii)
    If grdGrilla.Col = 1 Then KeyAscii = CarTexto(KeyAscii)
    If grdGrilla.Col = 2 Then KeyAscii = CarNumeroDecimal(txtEdit, KeyAscii)
    If grdGrilla.Col = 3 Then KeyAscii = CarNumeroDecimal(txtEdit, KeyAscii)
    If grdGrilla.Col = 4 Then KeyAscii = CarNumeroDecimal(txtEdit, KeyAscii)
    CarTexto KeyAscii
End Sub

Private Sub TxtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim mimpneto As Double
    
    If KeyCode = vbKeyReturn Then
        Select Case grdGrilla.Col
            Case 0 'CODIGO
                mPrecio = 0
                If Trim(txtEdit) <> "" Then
                    Set rec = New ADODB.Recordset
                    sql = "SELECT P.PTO_CODIGO, P.PTO_DESCRI,P.PTO_IVA, P.PTO_PRECTO, P.PTO_TASAVIAL " ', D.LIS_IVA"
                    sql = sql & " FROM PRODUCTO P" ', DETALLE_LISTA_PRECIO D"
                    sql = sql & " WHERE "
                    'sql = sql & " P.PTO_CODIGO=D.PTO_CODIGO"
                    'If IsNumeric(txtEdit) Then
                        sql = sql & "  (P.PTO_CODIGO =" & XN(txtEdit) & " OR P.PTO_CODBARRAS=" & XS(txtEdit) & ")"
                    'Else
                        'sql = sql & " AND (P.PTO_CODBARRAS=" & XS(txtEdit) & ")"
                    'End If
                    sql = sql & " AND P.PTO_ESTADO=" & XS("N")
                    sql = sql & " AND P.LIS_CODIGO=" & cboListaPrecio.ItemData(cboListaPrecio.ListIndex)
                    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
                    If rec.EOF = False Then
                    
                        mIVA_1 = BuscoIva ' IVA 21%
                        mIVA_2 = BuscoIva_2 ' IVA 40% SE USA PARA EL GASOIL
                        
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 0) = Trim(txtEdit.Text)
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 1) = Trim(rec!PTO_DESCRI)
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = "1"
                        
                        
                        
                        If Ltipo_fac.Caption = "B" Then
                            mPrecio = VALIDO_IMPORTE4(Chk0(rec!PTO_PRECTO))
                        Else
                            grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = Format(Chk0(mIVA_1), "0.0000")
                            
                            mPrecio = VALIDO_IMPORTE4(Chk0(rec!PTO_PRECTO))
                            grdGrilla.TextMatrix(grdGrilla.RowSel, 10) = VALIDO_IMPORTE4(CStr(mPrecio))
                        End If
                        
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 3) = VALIDO_IMPORTE4(CStr(mPrecio))
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 4) = VALIDO_IMPORTE4(CStr(mPrecio))
                        
                        If Ltipo_fac.Caption = "B" Then
                            grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = Format(Chk0(mIVA_1), "0.0000")
                            mPrecio = VALIDO_IMPORTE4(Chk0(rec!PTO_PRECTO))
                            'pongo el neto en la columna auxiliar 10
                            grdGrilla.TextMatrix(grdGrilla.RowSel, 10) = VALIDO_IMPORTE4(CStr(mPrecio))
                        Else
                            'pongo el neto en la columna auxiliar 10
                            grdGrilla.TextMatrix(grdGrilla.RowSel, 10) = VALIDO_IMPORTE4(CStr(mPrecio))
                        End If
                        
                        
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 5) = Trim(rec!PTO_CODIGO)
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 7) = Format(Chk0(rec!PTO_IVA), "0.0000") ' ALICUOTA IMPUESTO INTERNO
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 11) = Format(Chk0(rec!PTO_TASAVIAL), "0.0000") ' ALICUOTA IMPUESTO INTERNO
                        'calculo tasa vial
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 12) = Format(CStr(CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 2)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 11))), "0.0000")
                        
                        'CAMBIA PARA FACTURAS A
                        If Ltipo_fac.Caption = "B" Then
                            txtTotal.Text = Format(CStr(SumaTotal), "0.00")
                            
                        Else
                            txtSubtotal.Text = Format(CStr(SumaSUBTotal), "0.00")
                        
                        End If
                        calculototales rec!PTO_IVA, (rec!PTO_PRECTO - Chk0(rec!PTO_TASAVIAL)) 'PTO_IVA ES EN REALIDAD EL IMPUESTO INTERNO
                                                
                        txtsubtotal1.Text = VALIDO_IMPORTE4(CStr(SumaSUBTotal))
                        txtTotal.Text = Format(CStr(SumaTotal), "0.00")
                        'txtPorcentajeIva_LostFocus
                        Sumatorias
                        mFoco = True
                        grdGrilla.Col = 0
                        grdGrilla.row = grdGrilla.RowSel
                    Else
                        MsgBox "El Producto NO Existe", vbCritical, TIT_MSGBOX
                        txtEdit.Text = ""
                    End If
                    rec.Close
                End If
                
            Case 1 'DESCRIPCION
                If Trim(txtEdit) <> "" Then
                    Set rec = New ADODB.Recordset
                    sql = "SELECT P.PTO_CODIGO, P.PTO_DESCRI, P.PTO_PRECTO " ', D.LIS_IVA"
                    sql = sql & " FROM PRODUCTO P" ' , DETALLE_LISTA_PRECIO D"
                    sql = sql & " WHERE "
                    'sql = sql & " P.PTO_CODIGO=D.PTO_CODIGO"
                    sql = sql & " P.PTO_DESCRI LIKE '%" & Trim(txtEdit) & "%'"
                    sql = sql & " AND P.PTO_ESTADO=" & XS("N")
                    sql = sql & " AND P.LIS_CODIGO=" & cboListaPrecio.ItemData(cboListaPrecio.ListIndex)
                    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
                    If rec.EOF = False Then
                        If rec.RecordCount > 1 Then
                            mFoco = True
                            BuscarProducto grdGrilla, "CADENA", txtEdit.Text, grdGrilla.RowSel
'                            If BuscoRepetetidos(CStr(grdGrilla.TextMatrix(grdGrilla.RowSel, 5)), grdGrilla.RowSel) = False Then
'                                grdGrilla.Col = 0
'                                grdGrilla_KeyDown vbKeyDelete, 0
'                            End If
                            grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = Format(Chk0(mIVA_1), "0.0000")
                            'calculo tasa vial
                            grdGrilla.TextMatrix(grdGrilla.RowSel, 12) = Format(CStr(CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 2)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 11))), "0.0000")
                            
                            
                            calculototales grdGrilla.TextMatrix(grdGrilla.RowSel, 7), (CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 3)) - CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 11)))
                            txtsubtotal1.Text = VALIDO_IMPORTE4(CStr(SumaSUBTotal))
                            txtSubtotal.Text = VALIDO_IMPORTE4(CStr(SumaSUBTotal))
                            txtTotal.Text = Format(CStr(SumaTotal), "0.00")
                            'txtPorcentajeIva_LostFocus
                            Sumatorias
                            grdGrilla.Col = 1
                            
                        Else
                            grdGrilla.TextMatrix(grdGrilla.RowSel, 0) = Trim(rec!PTO_CODIGO)
                            txtEdit.Text = Trim(rec!PTO_DESCRI)
                            grdGrilla.TextMatrix(grdGrilla.RowSel, 1) = Trim(rec!PTO_DESCRI)
                            grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = "1"
                            
                            grdGrilla.TextMatrix(grdGrilla.RowSel, 11) = Format(Chk0(rec!PTO_TASAVIAL), "0.0000") ' ALICUOTA IMPUESTO INTERNO
                            'calculo tasa vial
                            grdGrilla.TextMatrix(grdGrilla.RowSel, 12) = Format(CStr(CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 2)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 11))), "0.0000")
                            'ver que hago con esto depende de como lo maneje en la lista de precios
                            ' ahora estoy tomando el IVA de Parametros
                            
                            'mValIVA = Format(Chk0(rec!LIS_IVA), "0.0000")
                            
                            If Ltipo_fac.Caption = "B" Then
                                mPrecio = VALIDO_IMPORTE4(Chk0(rec!PTO_PRECTO))
                            Else
                                calculototales grdGrilla.TextMatrix(grdGrilla.RowSel, 7), (CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 3)) - CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 11)))
                                'mValorIvaIns = (1 + (mValIVA / 100))
                                mPrecio = VALIDO_IMPORTE4(Chk0(rec!PTO_PRECTO))
                            End If
                                                        
                            grdGrilla.TextMatrix(grdGrilla.RowSel, 3) = VALIDO_IMPORTE4(CStr(mPrecio))
                            'grdGrilla.TextMatrix(grdGrilla.RowSel, 4) = VALIDO_IMPORTE4(CStr(mPrecio))
                            grdGrilla.TextMatrix(grdGrilla.RowSel, 5) = Trim(rec!PTO_CODIGO)
                            grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = Format(Chk0(mIVA_1), "0.0000")
                        
'                            If BuscoRepetetidos(CStr(grdGrilla.TextMatrix(grdGrilla.RowSel, 5)), grdGrilla.RowSel) = False Then
'                                grdGrilla.Col = 0
'                                grdGrilla_KeyDown vbKeyDelete, 0
'                            End If
                            If Ltipo_fac.Caption = "B" Then
                                txtTotal.Text = Format(CStr(SumaTotal), "0.00")
                                txtSubtotal.Text = VALIDO_IMPORTE4(CStr(SumaSUBTotal))
                            Else
                                'txtSubtotal.Text = VALIDO_IMPORTE4(CStr(SumaSUBTotal))
                            End If
                            
                            txtTotal.Text = Format(CStr(SumaTotal), "0.00")
                            txtPorcentajeIva_LostFocus
                        End If
                    Else
                        MsgBox "El Producto NO Existe", vbCritical, TIT_MSGBOX
                        txtEdit.Text = ""
                    End If
                    rec.Close
                End If
                
            Case 2 'CANTIDAD
                If Trim(txtEdit) = "" Then grdGrilla.Text = "1"
                grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = txtEdit.Text
                'txtEdit.Text = Format(txtEdit.Text, "0.00")
'                If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) <> "" Then
'                    If Trim(txtEdit) <> "" Then
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 4) = VALIDO_IMPORTE4(CStr(CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 2)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 3))))
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 10) = VALIDO_IMPORTE4(CStr(CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 2)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 3))))
'                    End If
'                End If
                
                'calculo tasa vial
                grdGrilla.TextMatrix(grdGrilla.RowSel, 12) = Format(CStr(CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 2)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 11))), "0.0000")
                
                If Ltipo_fac.Caption = "B" Then
                    calculototales grdGrilla.TextMatrix(grdGrilla.RowSel, 7), (CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 3)) - CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 11)))
                    'Sumatorias
                    txtTotal.Text = Format(CStr(SumaTotal), "0.00")
                    txtsubtotal1.Text = VALIDO_IMPORTE4(CStr(SumaSUBTotal))
                    txtSubtotal = Format(CStr(SumaSUBTotal), "0.0000")
                Else
                    'txtSubtotal.Text = VALIDO_IMPORTE4(CStr(SumaSUBTotal))
'                    mimpneto = CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 2)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 3)) - (CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 2)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 7)))
'                    mimpneto = mimpneto / (1 + (mIVA_1 / 100))
'                    txtEdit.Text = Format(mimpneto, "0.0000")
'                    grdGrilla.TextMatrix(grdGrilla.RowSel, 4) = Format(txtEdit.Text, "0.0000")
'                    grdGrilla.TextMatrix(grdGrilla.RowSel, 10) = Format(txtEdit.Text, "0.0000")
'                    txtSubtotal.Text = Format(CStr(SumaSUBTotal), "0.00")
                    
                    calculototales grdGrilla.TextMatrix(grdGrilla.RowSel, 7), (CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 3)) - CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 11)))
                    'Sumatorias
                End If
                
                txtTotal.Text = Format(CStr(SumaTotal), "0.0000")
                'txtPorcentajeIva_LostFocus
                
            Case 3 'PRECIO
                
                If Trim(txtEdit) = "" Then grdGrilla.Text = "0,000"
                txtEdit.Text = Format(txtEdit.Text, "0.0000")
                grdGrilla.TextMatrix(grdGrilla.RowSel, 3) = Format(txtEdit.Text, "0.0000")
                If Ltipo_fac.Caption = "A" Then
                    
                    grdGrilla.TextMatrix(grdGrilla.RowSel, 3) = Format(txtEdit.Text, "0.0000")
                End If
                If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) <> "" Then
                    If Trim(txtEdit) <> "" Then
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 4) = Format(CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 2)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 3)), "0.0000")
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 10) = Format(CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 2)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 3)), "0.0000")
                        'calculo tasa vial
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 12) = Format(CStr(CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 2)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 11))), "0.0000")
                    End If
                End If
                'If Ltipo_fac.Caption = "A" Then
                    'mValorIvaIns = Format((1 + (mIVAi / 100)), "0.00")
                    'mPrecio = VALIDO_IMPORTE4(Chk0(grdGrilla.TextMatrix(grdGrilla.RowSel, 4)) / mValorIvaIns)
                    'grdGrilla.TextMatrix(grdGrilla.RowSel, 4) = VALIDO_IMPORTE4(CStr(mPrecio))
                'End If
                
                
                If Ltipo_fac.Caption = "B" Then
                    txtTotal.Text = Format(CStr(SumaTotal), "0.00")
                    txtsubtotal1.Text = VALIDO_IMPORTE4(CStr(SumaSUBTotal))
                    txtSubtotal.Text = VALIDO_IMPORTE4(CStr(SumaSUBTotal))
                    calculototales grdGrilla.TextMatrix(grdGrilla.RowSel, 7), (CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 3)) - CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 11)))
                Else
                    txtSubtotal.Text = Format(CStr(SumaSUBTotal), "0.00")
                    calculototales grdGrilla.TextMatrix(grdGrilla.RowSel, 7), (CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 3)) - CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 11)))
                End If
                
                txtTotal.Text = Format(CStr(SumaTotal), "0.00")
                txtPorcentajeIva_LostFocus
             
             Case 4 'importe
                If Trim(txtEdit) = "" Then grdGrilla.Text = "0,00"
                txtEdit.Text = Format(txtEdit.Text, "0.00")
                grdGrilla.TextMatrix(grdGrilla.RowSel, 4) = Format(txtEdit.Text, "0.0000")
                grdGrilla.TextMatrix(grdGrilla.RowSel, 10) = Format(txtEdit.Text, "0.0000")
                
                'If Ltipo_fac.Caption = "A" Then
                    
                '    grdGrilla.TextMatrix(grdGrilla.RowSel, 4) = Format(txtEdit.Text, "0.0000")
                '    grdGrilla.TextMatrix(grdGrilla.RowSel, 10) = Format(txtEdit.Text, "0.0000")
                'End If
                If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) <> "" Then
                    If Trim(txtEdit) <> "" Then
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = Format(CStr(CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 4)) / CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 3))), "0.0000")
                        'calculo tasa vial
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 12) = Format(CStr(CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 2)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 11))), "0.0000")
                    End If
                End If
                'If Ltipo_fac.Caption = "A" Then
                    'mValorIvaIns = Format((1 + (mIVAi / 100)), "0.00")
                    'mPrecio = VALIDO_IMPORTE4(Chk0(grdGrilla.TextMatrix(grdGrilla.RowSel, 4)) / mValorIvaIns)
                    'grdGrilla.TextMatrix(grdGrilla.RowSel, 4) = VALIDO_IMPORTE4(CStr(mPrecio))
                'End If
                
                
                If Ltipo_fac.Caption = "B" Then
                    calculototales grdGrilla.TextMatrix(grdGrilla.RowSel, 7), (CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 3)) - CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 11)))
                    'Sumatorias
                    txtTotal.Text = Format(CStr(SumaTotal), "0.00")
                    txtsubtotal1.Text = VALIDO_IMPORTE4(CStr(SumaSUBTotal))
                    txtSubtotal.Text = VALIDO_IMPORTE4(CStr(SumaSUBTotal))
                Else
                'importe neto es el precio de venta - el impuesto interno divido por el 1,21 de iva
                    mimpneto = CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 2)) * (CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 3)) - CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 11))) - (CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 2)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 7)))
                    mimpneto = mimpneto / (1 + (mIVA_1 / 100))
                    txtEdit.Text = Format(mimpneto, "0.0000")
                    grdGrilla.TextMatrix(grdGrilla.RowSel, 4) = Format(txtEdit.Text, "0.0000")
                    grdGrilla.TextMatrix(grdGrilla.RowSel, 10) = Format(txtEdit.Text, "0.0000")
                    txtSubtotal.Text = Format(CStr(SumaSUBTotal), "0.00")
                    calculototales grdGrilla.TextMatrix(grdGrilla.RowSel, 7), (CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 3)) - CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 11)))
                End If
                
                txtTotal.Text = Format(CStr(SumaTotal), "0.00")
                'txtPorcentajeIva_LostFocus
        End Select
        mFoco = True
        grdGrilla.SetFocus
    End If
    If lblSaldoFac.Visible = True Then
        If txtTotal > SaldoCli Then
            MsgBox "Ha superado el Saldo de Facturacion ", vbExclamation, TIT_MSGBOX
            'cmdGrabar.Enabled = False
            lblblockeado.Visible = True
            lblblockeado.Caption = "Limite de Facturacion en Cta Cte superado"
            cboCondicion.Clear
            cboFormaPago.Clear
            LlenarComboFormaPago
            txtEdit.Visible = False
            grdGrilla.SetFocus
        Else
            lblblockeado.Visible = False
            cmdGrabar.Enabled = True
        End If
    End If
    If KeyCode = vbKeyEscape Then
       txtEdit.Visible = False
       grdGrilla.SetFocus
    End If

End Sub
Private Sub calculototales(pImpue As Double, pPrecio As Double)
    'Dim Vimpuesto As Double
    Dim subtotal As Double
    Dim ImpIVA As Double
    Dim mTasa As Double
    Dim mIVA_1 As Double
    Dim mIVA_2 As Double
    
    'ARMAR LAS SUMATORIAS DE LOS IMPUESTOS Y VER BIEN EL TEMA
    'DE LOS ARTICULOS SIN IMPUESTOS
    
    subtotal = 0
    ImpIVA = 0
    mTasa = 0
    
    mIVA_1 = BuscoIva
    mIVA_2 = BuscoIva_2
    
    If pImpue <> 0 Then
'        pImpue = 1
        IMPUESTO = pImpue * grdGrilla.TextMatrix(grdGrilla.RowSel, 2) ' PASO 1
        grdGrilla.TextMatrix(grdGrilla.RowSel, 8) = pImpue * grdGrilla.TextMatrix(grdGrilla.RowSel, 2)
        ' SUMO TOTALES DE IMPUESTOS Y LOS GUARDO EN TXTIMPUESTO
        
        'NAFTA Y GNC
        If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) = 1 Or grdGrilla.TextMatrix(grdGrilla.RowSel, 0) = 2 Or grdGrilla.TextMatrix(grdGrilla.RowSel, 0) = 4 Then
            subtotal = grdGrilla.TextMatrix(grdGrilla.RowSel, 2) * pPrecio - grdGrilla.TextMatrix(grdGrilla.RowSel, 8)
            subtotal = subtotal / (1 + (mIVA_1 / 100)) 'paso 2
            If Ltipo_fac = "A" Then
                grdGrilla.TextMatrix(grdGrilla.RowSel, 4) = Format(subtotal, "0.0000")
            End If
            grdGrilla.TextMatrix(grdGrilla.RowSel, 10) = Format(subtotal, "0.0000")
            ImpIVA = subtotal * (mIVA_1 / 100)
            grdGrilla.TextMatrix(grdGrilla.RowSel, 9) = Format(ImpIVA, "0.0000")
        Else
        'GASOIL
            subtotal = grdGrilla.TextMatrix(grdGrilla.RowSel, 2) * pPrecio - grdGrilla.TextMatrix(grdGrilla.RowSel, 8)
            subtotal = subtotal / (1 + (mIVA_2 / 100)) 'paso 2
              If Ltipo_fac = "A" Then
                grdGrilla.TextMatrix(grdGrilla.RowSel, 4) = Format(subtotal, "0.0000")
            End If
            grdGrilla.TextMatrix(grdGrilla.RowSel, 10) = Format(subtotal, "0.0000")
            ImpIVA = subtotal * (mIVA_1 / 100)
            mTasa = subtotal * ((mIVA_2 - mIVA_1) / 100) '
            grdGrilla.TextMatrix(grdGrilla.RowSel, 8) = Format(IMPUESTO + mTasa, "0.0000")
            grdGrilla.TextMatrix(grdGrilla.RowSel, 9) = Format(ImpIVA, "0.0000")
        End If
        
    Else
                
        subtotal = grdGrilla.TextMatrix(grdGrilla.RowSel, 2) * pPrecio
        subtotal = subtotal / (1 + (mIVA_1 / 100))
        'impIVA = subtotal * (mIVA_1 / 100)
        If Ltipo_fac = "A" Then
            grdGrilla.TextMatrix(grdGrilla.RowSel, 4) = Format(subtotal, "0.0000")
        End If
        grdGrilla.TextMatrix(grdGrilla.RowSel, 10) = Format(subtotal, "0.0000")
        
        'txtimpuesto.Text = "0,00"
        grdGrilla.TextMatrix(grdGrilla.RowSel, 8) = 0
        ImpIVA = subtotal * (mIVA_1 / 100)
        grdGrilla.TextMatrix(grdGrilla.RowSel, 9) = Format(ImpIVA, "0.0000")
    End If
    
    txtnoinsc.Text = "0,00"
    Sumatorias
   
End Sub

Private Function Sumatorias()
    Dim Vimpuesto As Double
    Dim vSubtotal As Double
    Dim VsubtotalB As Double
    Dim Vimpiva As Double
    Dim vTasaVial As Double
    VTotal = 0
    Vimpiva = 0
    vTasaVial = 0
    For i = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(i, 0) <> "" Then
            
            vSubtotal = vSubtotal + CDbl(grdGrilla.TextMatrix(i, 4))
            
            VsubtotalB = VsubtotalB + CDbl(grdGrilla.TextMatrix(i, 10))
            
            
            
            Vimpuesto = Vimpuesto + CDbl(grdGrilla.TextMatrix(i, 8))
            Vimpiva = Vimpiva + CDbl(grdGrilla.TextMatrix(i, 9))
            'VTotal = VTotal + CDbl(grdGrilla.TextMatrix(I, 2)) * CDbl(grdGrilla.TextMatrix(I, 3))
            
            'sumo la tasa vial
            vTasaVial = vTasaVial + CDbl(grdGrilla.TextMatrix(i, 12))
        End If
    Next
    
    'ESTO ES LO QUE SE GUARDA EN LA bd PERO NO SE MUESTRA
    txtsubtotal1B.Text = Format(VsubtotalB, "0.0000")
    txtimpuestoB.Text = Format(Vimpuesto, "0.0000")
    txtSubtotalB.Text = Format(VsubtotalB + Vimpuesto, "0.0000")
    txtImporteIvaB.Text = Format(Vimpiva, "0.0000")
        
    
    txtsubtotal1.Text = Format(vSubtotal, "0.0000")
    If Ltipo_fac = "B" Then
        Vimpuesto = 0
        Vimpiva = 0
    End If
    txtimpuesto.Text = Format(Vimpuesto, "0.0000")
    txtSubtotal.Text = Format(vSubtotal + Vimpuesto, "0.0000")
    txtImporteIva.Text = Format(Vimpiva, "0.0000")
    txttasavial.Text = Format(vTasaVial, "0.0000")
    
End Function


Private Function BuscoRepetetidos(Codigo As String, Linea As Integer) As Boolean
    For i = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(i, 5) <> "" Then
            If Codigo = CStr(grdGrilla.TextMatrix(i, 5)) And (i <> Linea) Then
                MsgBox "El Producto ya fue elegido anteriormente", vbExclamation, TIT_MSGBOX
                BuscoRepetetidos = False
                Exit Function
            End If
        End If
    Next
    BuscoRepetetidos = True
End Function
Private Function SumaSUBTotal() As Double
    'If Ltipo_fac = "A" Then
        VTotal = 0
        For i = 1 To grdGrilla.Rows - 1
            If grdGrilla.TextMatrix(i, 4) <> "" Then
                VTotal = VTotal + CDbl(grdGrilla.TextMatrix(i, 4))
            End If
        Next
        SumaSUBTotal = VALIDO_IMPORTE4(CStr(VTotal))
    'End If
End Function


Private Function SumaTotal() As Double
   ' If Ltipo_fac = "B" Then
        VTotal = 0
        For i = 1 To grdGrilla.Rows - 1
            If grdGrilla.TextMatrix(i, 4) <> "" Then
                VTotal = VTotal + CDbl(grdGrilla.TextMatrix(i, 2)) * CDbl(grdGrilla.TextMatrix(i, 3))
            End If
        Next
        SumaTotal = VALIDO_IMPORTE4(CStr(VTotal))
    'End If
End Function

Private Function SumaBonificacion() As Double
    VTotal = 0
    For i = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(i, 4) <> "" Then
            VTotal = VTotal + CDbl(grdGrilla.TextMatrix(i, 4))
        End If
    Next
    SumaBonificacion = VALIDO_IMPORTE4(CStr(VTotal))
End Function

Private Sub txtImportePago_GotFocus()
    txtImportePago.Text = txtTotalPagos.Text
    SelecTexto txtImportePago
End Sub

Private Sub txtImportePago_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtImportePago, KeyAscii)
End Sub

Private Sub txtImportePago_LostFocus()
    If txtcodCli.Text = "1" Then
        If cboFormaPago.ItemData(cboFormaPago.ListIndex) = 2 Then
            'MsgBox "No Puede Seleccionar Cta CTe para este Cliente", vbCritical, TIT_MSGBOX
            'cboFormaPago.SetFocus
            Exit Sub
        End If
    End If
    
    If fraTarjeta.Visible = True Then Exit Sub
    If fracheque.Visible = True Then Exit Sub
    txtImportePago.Text = Format(txtImportePago.Text, "0.00")
    Dim mTotalPagos As Double
    mTotalPagos = 0
    
    
    
    For i = 1 To grdPagos.Rows - 1
        mTotalPagos = mTotalPagos + CDbl(grdPagos.TextMatrix(i, 1))
    Next
    If mTotalPagos + CDbl(Chk0(txtImportePago.Text)) > CDbl(txtTotal.Text) Then
        MsgBox "El Importe Ingresado Exede el Monto de la Compra!", vbInformation, TIT_MSGBOX
        txtImportePago.SetFocus
        Exit Sub
    Else
        If cboFormaPago.Text = "" Then
            MsgBox "Debe Indicar la Forma de Pago", vbCritical, TIT_MSGBOX
            cboFormaPago.SetFocus
            Exit Sub
        End If
        If CDbl(Chk0(txtImportePago.Text)) >= 0 Then
            grdPagos.AddItem ("")
            grdPagos.row = grdPagos.Rows - 1
            grdPagos.TextMatrix(grdPagos.row, 0) = Trim(Mid(cboFormaPago.Text, 1, 30))
            grdPagos.TextMatrix(grdPagos.row, 1) = txtImportePago.Text
            grdPagos.TextMatrix(grdPagos.row, 2) = cboFormaPago.ItemData(cboFormaPago.ListIndex)
            mFormaPago = cboFormaPago.ItemData(cboFormaPago.ListIndex)
            
            If Trim(UCase(Mid(cboFormaPago.Text, 1, 50))) = "TARJETA DE CREDITO" Then
                grdPagos.TextMatrix(grdPagos.row, 3) = cboTarjeta.ItemData(cboTarjeta.ListIndex)
                grdPagos.TextMatrix(grdPagos.row, 4) = cboTarjeta.List(cboTarjeta.ListIndex)
                grdPagos.TextMatrix(grdPagos.row, 5) = cboPlan.ItemData(cboPlan.ListIndex)
                grdPagos.TextMatrix(grdPagos.row, 6) = cboPlan.List(cboPlan.ListIndex)
                grdPagos.TextMatrix(grdPagos.row, 7) = txtCupon.Text
                grdPagos.TextMatrix(grdPagos.row, 8) = txtLote.Text
                grdPagos.TextMatrix(grdPagos.row, 9) = txtTar_Autorizacion.Text
            End If
            If Trim(UCase(Mid(cboFormaPago.Text, 1, 50))) = "TARJETA DE DEBITO" Then
                grdPagos.TextMatrix(grdPagos.row, 3) = cboTarjeta.ItemData(cboTarjeta.ListIndex)
                grdPagos.TextMatrix(grdPagos.row, 4) = cboTarjeta.List(cboTarjeta.ListIndex)
            End If
            If Trim(UCase(Mid(cboFormaPago.Text, 1, 50))) = "CHEQUE" Then
                grdPagos.TextMatrix(grdPagos.row, 2) = cboFormaPago.ItemData(cboFormaPago.ListIndex)
                grdPagos.TextMatrix(grdPagos.row, 4) = txtchebanco.Text
                grdPagos.TextMatrix(grdPagos.row, 6) = txtchenumero.Text
            End If
'            If Trim(UCase(Mid(cboFormaPago.Text, 1, 50))) = "DOLARES" Then
'                grdPagos.TextMatrix(grdPagos.row, 10) = txtTotDolar.Text
'                grdPagos.TextMatrix(grdPagos.row, 11) = txtCotizacion.Text
'            End If
        End If
    End If
    mTotalPagos = 0
    For i = 1 To grdPagos.Rows - 1
        mTotalPagos = CDbl(mTotalPagos) + CDbl(grdPagos.TextMatrix(i, 1))
    Next
    txtTotalPagos.Text = Format(CDbl(txtTotal.Text) - mTotalPagos, "0.00")
    txtImportePago.Text = Format(txtTotalPagos.Text, "0.00")
    If Val(txtTotalPagos.Text) = 0 Then
        cmdAceptarPagos.SetFocus
    Else
        cboFormaPago.ListIndex = 0
        cboFormaPago.SetFocus
    End If
    txtTar_Autorizacion.Text = ""
    txtLote.Text = ""
    txtCupon.Text = ""
    cboPlan.Clear
End Sub

Private Sub txtLote_GotFocus()
    SelecTexto txtLote
End Sub

Private Sub txtLote_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroFactura_GotFocus()
    SelecTexto txtNroFactura
End Sub

Private Sub txtNroFactura_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroFactura_LostFocus()
    If txtNroFactura.Text = "" Then
        'BUSCO EL NUMERO DE FACTURA QUE CORRESPONDE
        txtNroFactura.Text = Format(BuscoUltimaFactura(cboFactura.ItemData(cboFactura.ListIndex)), "00000000")
    Else
        txtNroFactura.Text = Format(txtNroFactura.Text, "00000000")
    End If
End Sub

Private Sub txtNroSucursal_GotFocus()
    SelecTexto txtNroSucursal
End Sub

Private Sub txtNroSucursal_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroSucursal_LostFocus()
    If txtNroSucursal.Text = "" Then
        txtNroSucursal.Text = Format(Sucursal, "0000")
    Else
        txtNroSucursal.Text = Format(txtNroSucursal.Text, "0000")
    End If
End Sub

Private Sub txtPorcentajeIva_GotFocus()
    SelecTexto txtPorcentajeIva
End Sub

Private Sub txtPorcentajeIva_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtPorcentajeIva, KeyAscii)
End Sub

Private Sub txtPorcentajeIva_LostFocus()
    If txtPorcentajeIva.Text <> "" And txtSubtotal.Text <> "" Then
        If ValidarPorcentaje(txtPorcentajeIva) = False Then
            txtPorcentajeIva.SetFocus
            Exit Sub
        End If
        If Ltipo_fac.Caption = "A" Then
'            Dim mCalculo As Double
'            mCalculo = 0
            txtImporteIva.Text = "0,00"
            'mValorIvaIns = (1 + (mValIVA / 100))
            'mCalculo = CDbl(txtSubtotal.Text * mValorIvaIns)
            'txtImporteIva.Text = (mCalculo) - (mCalculo / mValorIvaIns)
            For J = 1 To grdGrilla.Rows - 1
                If grdGrilla.TextMatrix(J, 0) <> "" Then
                   txtImporteIva.Text = CDbl(Chk0(txtImporteIva.Text)) + ((CDbl(Chk0(grdGrilla.TextMatrix(J, 4)) * CDbl(Chk0(grdGrilla.TextMatrix(J, 2)))) * CDbl(Chk0(grdGrilla.TextMatrix(J, 6)))) / 100)
                End If
            Next
            'ivains = valtot * mIVAi / 100
            'txtImporteIva.Text = (CDbl(txtSubtotal.Text) * CDbl(txtPorcentajeIva.Text)) / 100
            txtImporteIva.Text = VALIDO_IMPORTE4(txtImporteIva.Text)
        Else
            txtImporteIva.Text = "0,00"
        End If
'        txtImporteIva.Text = (CDbl(txtSubtotal.Text) * CDbl(txtPorcentajeIva.Text)) / 100
'        txtImporteIva.Text = VALIDO_IMPORTE4(txtImporteIva.Text)
        txtTotal.Text = CDbl(txtSubtotal.Text) + CDbl(txtImporteIva.Text)
        txtTotal.Text = Format(txtTotal.Text, "0.00")
    End If
End Sub

Private Sub txtRazSoc_Change()
    If txtRazSoc.Text = "" Then
        txtcodCli.Text = ""
        txtDomici.Text = ""
        txtCuit.Text = ""
        txtCiva.Text = ""
        txtTelefono.Text = ""
        txtIngBrutos.Text = ""
        mRespo.Text = ""
    'Else
    '    txtcodCli.Text = ""
    End If
End Sub
Private Sub txtRazSoc_GotFocus()
    SelecTexto txtRazSoc
End Sub

Private Sub txtRazSoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarClientes "txtcodCli", "CODIGO"
    End If
End Sub

Private Sub txtRazSoc_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtRazSoc_LostFocus()
    If txtcodCli.Text = "" And txtRazSoc.Text <> "" Then
        sql = "SELECT C.CLI_CODIGO,C.CLI_RAZSOC,C.CLI_DOMICI,I.IVA_DESCRI, C.CLI_CUIT,"
        sql = sql & " I.IVA_CODIGO, C.CLI_NRODOC, I.IVA_LETRA, C.CLI_TELEFONO, C.CLI_INGBRU"
        sql = sql & " FROM CLIENTE C, CONDICION_IVA I"
        sql = sql & " WHERE I.IVA_CODIGO = C.IVA_CODIGO"
        sql = sql & " AND CLI_RAZSOC LIKE '%" & XN(Trim(txtRazSoc.Text)) & "%'"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If rec.RecordCount > 1 Then
                BuscarClientes "txtcodCli", "CADENA", Trim(txtRazSoc.Text)
                If rec.State = 1 Then rec.Close
                txtRazSoc.SetFocus
            Else
                If mQuienLlama = "" Then
'                    If mBuscador = False Then
'                        LIMPIOGRILLA
'                        txtsubtotal1.Text = "0,000"
'                        txtSubtotal.Text = "0,000"
'                        txtTotal.Text = "0,000"
'                        txtPorcentajeIva.Text = Format(mIVAi, "0.0000")
'                        txtImporteIva.Text = "0,000"
'                    End If
                End If
                txtcodCli.Text = rec!CLI_CODIGO
                txtRazSoc.Text = rec!CLI_RAZSOC
                txtDomici.Text = ChkNull(rec!CLI_DOMICI) & ", " & Trim(ChkNull(LOC_DESCRI))
                txtCiva.Text = ChkNull(rec!IVA_DESCRI)
                txtCuit.Text = ChkNull(rec!CLI_CUIT)
                txtTelefono.Text = ChkNull(rec!CLI_TELEFONO)
                txtIngBrutos.Text = ChkNull(rec!CLI_INGBRU)
                mRespo.Text = ChkNull(rec!IVA_LETRA)
                QueFacturaUso (rec!IVA_CODIGO)
                txtNRO_DOCUMENTO.Text = Trim(ChkNull(rec!CLI_NRODOC))
                
                If mQuienLlama = "" Then
                    If mVerCta = True Then
                        'Call BuscarPendienteClientes(txtcodCli.Text, True, True)
                    End If
                End If
                If cmdGrabar.Enabled = True Then
                    'BUSCO EL NUMERO DE FACTURA EN EL FISCAL
                     Select Case cboFactura.ItemData(cboFactura.ListIndex)
                         Case 1 'FACTURAS A
                             pf.Status ("A")
                             txtNroFactura.Text = Val(pf.AnswerField_7) + 1
                         Case 2 'FACTURA B
                             pf.Status ("A")
                             txtNroFactura.Text = Val(pf.AnswerField_5) + 1
                         Case 3 'FACTURA C
                         Case 10000 'PARA TIKET
                             pf.Status ("A")
                             txtNroFactura.Text = Val(pf.AnswerField_4) + 1
                     End Select
                End If
            End If
        Else
            lblEstado.Caption = ""
            MsgBox "El Cliente no existe", vbExclamation, TIT_MSGBOX
            txtcodCli.Text = ""
            txtRazSoc.SetFocus
        End If
        If rec.State = 1 Then rec.Close
    End If
End Sub

Private Sub EstadoFactura(Estado As Integer)
        sql = "SELECT * FROM ESTADO_DOCUMENTO"
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            Do While Rec1.EOF = False
                If Rec1!EST_CODIGO = Estado Then
                    lblEstadoFactura.Caption = Rec1!EST_DESCRI
                End If
                Rec1.MoveNext
            Loop
        End If
        Rec1.Close
End Sub

Public Sub BuscarProducto(Txt As Control, mQuien As String, Optional mCadena As String, Optional mFila As Integer)
    Dim cSQL As String
    Dim hSQL As String
    Dim B As CBusqueda
    Dim i, posicion As Integer
    Dim cadena As String
    
    Set B = New CBusqueda
    With B
        'Set .Conn = DBConn
        cSQL = "SELECT P.PTO_DESCRI, P.PTO_CODIGO, P.PTO_PRECTO,P.PTO_IVA,P.PTO_TASAVIAL"
        cSQL = cSQL & " FROM PRODUCTO P" ', DETALLE_LISTA_PRECIO D"
        'cSQL = cSQL & "  " 'P.PTO_CODIGO=D.PTO_CODIGO"
        If mQuien = "CADENA" Then
            cSQL = cSQL & " WHERE  P.PTO_DESCRI LIKE '%" & Trim(mCadena) & "%'"
            cSQL = cSQL & " AND P.PTO_ESTADO=" & XS("N")
        End If
        'cSQL = cSQL & " AND D.LIS_CODIGO=" & cboListaPrecio.ItemData(cboListaPrecio.ListIndex)
        
        hSQL = "Descripción, Código, Precio, Iva, TasaVial"
        .sql = cSQL
        .Headers = hSQL
        .Field = "PTO_DESCRI"
        campo1 = .Field
        .Field = "PTO_CODIGO"
        campo2 = .Field
        .Field = "PTO_PRECTO"
        campo3 = .Field
        .Field = "PTO_IVA"
        campo4 = .Field
        .Field = "PTO_TASAVIAL"
        campo5 = .Field
        .OrderBy = "PTO_DESCRI"
        camponumerico = False
        .Titulo = "Busqueda de Productos :"
        .MaxRecords = 1
        .Show

        ' utilizar la coleccion de datos devueltos
        If .ResultFields.Count > 0 Then
            If mQuien = "CODIGO" Then
                grdGrilla.Col = 0
                txtEdit.Text = .ResultFields(2)
                TxtEdit_KeyDown 13, 0
                mFoco = True
                grdGrilla.Col = 0
                grdGrilla.row = mFila
            Else
                mPrecio = 0
                mFoco = True
                grdGrilla.TextMatrix(grdGrilla.RowSel, 0) = .ResultFields(2)
                txtEdit.Text = .ResultFields(1)
                grdGrilla.TextMatrix(grdGrilla.RowSel, 1) = .ResultFields(1)
                grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = "1"
                
                mIVA_1 = Format(Chk0(.ResultFields(4)), "0.0000")
                If Ltipo_fac.Caption = "B" Then
                    mPrecio = VALIDO_IMPORTE4(Chk0(.ResultFields(3)))
                Else
                    mValorIvaIns = (1 + (mIVA_1 / 100))
                    mPrecio = VALIDO_IMPORTE4(Chk0(.ResultFields(3)) / mValorIvaIns)
                End If
                                            
                grdGrilla.TextMatrix(grdGrilla.RowSel, 3) = VALIDO_IMPORTE4(CStr(mPrecio))
                grdGrilla.TextMatrix(grdGrilla.RowSel, 4) = VALIDO_IMPORTE4(CStr(mPrecio))
                grdGrilla.TextMatrix(grdGrilla.RowSel, 5) = .ResultFields(2)
                grdGrilla.TextMatrix(grdGrilla.RowSel, 7) = .ResultFields(4)
                grdGrilla.TextMatrix(grdGrilla.RowSel, 10) = VALIDO_IMPORTE4(CStr(mPrecio))
                grdGrilla.TextMatrix(grdGrilla.RowSel, 11) = .ResultFields(5)
                grdGrilla.Col = 1
                grdGrilla.row = mFila
            End If
        End If
    End With
    
    Set B = Nothing
End Sub

Public Sub BuscarClientes(Txt As String, mQuien As String, Optional mCadena As String)
    Dim cSQL As String
    Dim hSQL As String
    Dim B As CBusqueda
    Dim i, posicion As Integer
    Dim cadena As String
    
    Set B = New CBusqueda
    With B
        cSQL = "SELECT CLI_RAZSOC, CLI_DOMICI, CLI_CODIGO"
        cSQL = cSQL & " FROM CLIENTE C"
        If mQuien = "CADENA" Then
            cSQL = cSQL & " WHERE CLI_RAZSOC LIKE '%" & Trim(mCadena) & "%'"
        End If
        
        hSQL = "Nombre, Domicilio, Código"
        .sql = cSQL
        .Headers = hSQL
        .Field = "CLI_RAZSOC"
        campo1 = .Field
        .Field = "CLI_DOMICI"
        campo2 = .Field
        .Field = "CLI_CODIGO"
        campo3 = .Field
        .OrderBy = "CLI_RAZSOC"
        camponumerico = False
        .Titulo = "Busqueda de Clientes :"
        .MaxRecords = 1
        .Show

        ' utilizar la coleccion de datos devueltos
        If .ResultFields.Count > 0 Then
            If Txt = "txtcodCli" Then
                txtcodCli.Text = .ResultFields(3)
                txtCodCli_LostFocus
            Else
                txtBuscaCliente.Text = .ResultFields(3)
                txtBuscaCliente_LostFocus
            End If
        End If
    End With
    
    Set B = Nothing
End Sub

Private Sub txtTar_Autorizacion_GotFocus()
    SelecTexto txtTar_Autorizacion
End Sub

Private Sub txtTar_Autorizacion_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub ImprimirPagare(mCtaCteImp As String)
    Dim TotLetra As String
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    Rep.Formulas(2) = ""
    Rep.Formulas(3) = ""
    Rep.Formulas(4) = ""
    
    Rep.SelectionFormula = "{FORMA_PAGO.FPG_CODIGO}=2"
    Rep.Destination = crptToPrinter
    'Rep.Destination = crptToWindow
    'Rep.WindowState = crptMaximized
    'Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR
    'Rep.WindowTitle = "Listado de Composturas"
    TotLetra = LeeNro(CDbl(Format(mCtaCteImp, "0.00")), 80, 80, "$", "*", "*")
    
    Rep.Formulas(0) = "IMPORTE=" & XN(Chk0(mCtaCteImp))
    Rep.Formulas(1) = "PILAR='PILAR  " & Format(Date, "dddd, d") & " de " & Format(Date, "mmmm") & " de " & Format(Date, "yyyy") & "'"
    Rep.Formulas(2) = "LETRA='" & Mid(TotLetra, 1, 45) & "'"
    Rep.Formulas(3) = "LETRA1='" & Mid(TotLetra, 46, 100) & "'"
    If txtCuit.Text = "" Then
        Rep.Formulas(4) = "QUIEN='(" & txtNRO_DOCUMENTO.Text & ")   " & Trim(txtRazSoc.Text) & "'"
    Else
        Rep.Formulas(4) = "QUIEN='(" & Format(txtCuit.Text, "##-########-#") & ")   " & Trim(txtRazSoc.Text) & "'"
    End If
    
    
    Rep.ReportFileName = DRIVE & DirReport & "Pagare.rpt"
    Rep.Action = 1
End Sub

Private Function BuscarSaldoFactura(CodCli As String, limiteCtaCTe As Double) As Double
        'GrillaAplicar.Rows = 1
        Set Rec1 = New ADODB.Recordset
        Dim TotalDeuda As Double
        TotalDeuda = 0
        'BUSCA LAS FACTURAS
'        sql = "SELECT FCL_NUMERO AS NUMERO, FCL_SUCURSAL AS SUCURSAL, "
'        sql = sql & " FCL_FECHA AS FECHA, FCL_TOTAL AS TOTAL, FCL_SALDO AS SALDO"
'        sql = sql & " ,TCO_CODIGO AS TIPO, TCO_ABREVIA AS ABREVIA"
'        sql = sql & " FROM SALDO_FACTURAS_CLIENTE_V"
'        sql = sql & " WHERE "
'        sql = sql & " CLI_CODIGO=" & XN(CodCli)
'        sql = sql & " UNION ALL"
'
'        'BUSCA LAS NOTA DE DEBITO
'        sql = sql & " SELECT NDC_NUMERO AS NUMERO, NDC_SUCURSAL AS SUCURSAL, "
'        sql = sql & " NDC_FECHA AS FECHA, NDC_TOTAL AS TOTAL, NDC_SALDO AS SALDO"
'        sql = sql & " ,TCO_CODIGO AS TIPO, TCO_ABREVIA AS ABREVIA"
'        sql = sql & " FROM SALDO_NOTA_DEBITO_CLIENTE_V"
'        sql = sql & " WHERE "
'        sql = sql & " CLI_CODIGO=" & XN(CodCli)
'        sql = sql & " ORDER BY FECHA , NUMERO ASC"
           
        sql = "DELETE FROM CTA_CTE_CLIENTE"
        DBConn.Execute sql
       
        'TODAS LAS FACTURAS
        sql = "INSERT INTO CTA_CTE_CLIENTE (CLI_CODIGO,TCO_CODIGO,COM_NUMERO,COM_SUCURSAL,"
        sql = sql & " COM_FECHA,COM_IMPORTE,COM_IMP_DEBE,COM_IMP_HABER,CTA_CTE_DH,COM_NUMEROTXT)"
        sql = sql & " SELECT F.CLI_CODIGO,F.TCO_CODIGO,F.FCL_NUMERO,F.FCL_SUCURSAL,"
        sql = sql & " F.FCL_FECHA,F.FCL_TOTAL,F.FCL_TOTALACT,0 AS HABER,'D' AS DEBE,FCL_NUMEROTXT"
        sql = sql & " FROM FACTURA_CLIENTE F"
        sql = sql & " WHERE F.EST_CODIGO=3"
        sql = sql & " AND FPG_CODIGO=2"
        sql = sql & " AND F.CLI_CODIGO=" & XN(txtcodCli.Text)
        DBConn.Execute sql
    
        'TODAS LAS NOTAS DEBITOS CLIENTE
        sql = "INSERT INTO CTA_CTE_CLIENTE (CLI_CODIGO,TCO_CODIGO,COM_NUMERO,COM_SUCURSAL,"
        sql = sql & " COM_FECHA,COM_IMPORTE,COM_IMP_DEBE,COM_IMP_HABER,CTA_CTE_DH,"
        sql = sql & " COM_NUMEROTXT)"
        sql = sql & " SELECT DISTINCT N.CLI_CODIGO,N.TCO_CODIGO,N.NDC_NUMERO,N.NDC_SUCURSAL,"
        sql = sql & " N.NDC_FECHA,N.NDC_TOTAL,N.NDC_TOTAL,0 AS HABER,'D' AS DEBE,N.NDC_NUMEROTXT"
        sql = sql & " FROM NOTA_DEBITO_CLIENTE N"
        sql = sql & " WHERE N.EST_CODIGO=3"
        sql = sql & " AND N.CLI_CODIGO=" & XN(txtcodCli.Text)
        DBConn.Execute sql
        
        'TODAS LAS NOTAS CREDITO CLIENTE
        sql = "INSERT INTO CTA_CTE_CLIENTE (CLI_CODIGO,TCO_CODIGO,COM_NUMERO,COM_SUCURSAL,"
        sql = sql & " COM_FECHA,COM_IMPORTE,COM_IMP_DEBE,COM_IMP_HABER,CTA_CTE_DH,"
        sql = sql & " COM_NUMEROTXT)"
        sql = sql & " SELECT DISTINCT N.CLI_CODIGO,N.TCO_CODIGO,N.NCC_NUMERO,N.NCC_SUCURSAL,"
        sql = sql & " N.NCC_FECHA,N.NCC_TOTAL,0 AS DEBE,NCC_TOTAL,'C' AS CREDITO,N.NCC_NUMEROTXT"
        sql = sql & " FROM NOTA_CREDITO_CLIENTE N"
        sql = sql & " WHERE N.EST_CODIGO=3"
        sql = sql & " AND N.CLI_CODIGO=" & XN(txtcodCli.Text)
        DBConn.Execute sql
        
        'TODOS LOS RECIBOS
        sql = "INSERT INTO CTA_CTE_CLIENTE (CLI_CODIGO,TCO_CODIGO,COM_NUMERO,COM_SUCURSAL,"
        sql = sql & " COM_FECHA,COM_IMPORTE,COM_IMP_DEBE,COM_IMP_HABER,CTA_CTE_DH,"
        sql = sql & " COM_NUMEROTXT)"
        sql = sql & " SELECT DISTINCT R.CLI_CODIGO,R.TCO_CODIGO,R.REC_NUMERO,R.REC_SUCURSAL,"
        sql = sql & " R.REC_FECHA,R.REC_TOTAL,0 AS DEBE,REC_TOTAL,'C' AS CREDITO,R.REC_NUMEROTXT"
        sql = sql & " FROM RECIBO_CLIENTE R"
        sql = sql & " WHERE R.EST_CODIGO=3"
        sql = sql & " AND R.CLI_CODIGO=" & XN(txtcodCli.Text)
        DBConn.Execute sql
        
        sql = "SELECT SUM(COM_IMP_DEBE) AS DEBE,SUM(COM_IMP_HABER) AS HABER FROM CTA_CTE_CLIENTE"
        
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
        'Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            Do While Rec1.EOF = False
                'If Rec1!Saldo > 0 Then

                TotalDeuda = VALIDO_IMPORTE4(Chk0(Rec1!DEBE) - Chk0(Rec1!HABER))
                'End If
                Rec1.MoveNext
            Loop
'            GrillaAplicar.HighLight = flexHighlightAlways
            If limiteCtaCTe = 0 Then
                BuscarSaldoFactura = Format(TotalDeuda, "#,##0.0000")
            Else
                BuscarSaldoFactura = Format(limiteCtaCTe - TotalDeuda, "#,##0.0000")
            End If
            'txtSaldo.Text = Format(TotalDeuda, "0.00")
        Else
            BuscarSaldoFactura = limiteCtaCTe
        End If
        Rec1.Close
End Function

Private Sub Imprimo_NC_Fiscal()
    Dim mContCanti As Integer
    Dim mContPrecio As Integer
    Dim mContInternos As Double
    Dim mVendedor As String
    Dim mIngBrutos As String
    Dim mItc As Double
    Dim mPneto As Double
    Dim mTasa As Double
    Dim mTVialDbl As Double
    
    Dim mTotaLl As Double
    Dim ITOTAL As String

    Dim iCanti As String
    Dim iPrecio As String
    Dim iImpInt As String
    Dim mIvaFE As String
    Dim mTasaVial As String
    
'    mContCanti = 1000
'    mContPrecio = 100
    mIVA_1 = BuscoIva
    mIVA_2 = BuscoIva_2
    
    'Lo cambio para el redondeo de combustibles
    mContCanti = 100
    mContPrecio = 1000
    mContInternos = 100000000
    'mContInternos = 1000
    
    mVendedor = "Playero: " & cboVendedor.List(cboVendedor.ListIndex)
    mVendedor = Mid(mVendedor, 1, 20)
    
    mIngBrutos = "Ing Brutos: " & IIf(txtIngBrutos.Text = "" Or txtIngBrutos.Text = "0", "NO POSEE", txtIngBrutos.Text)
    mIngBrutos = Mid(mIngBrutos, 1, 25)
    
    
    If InStr(1, mDireccion.Text, "Á") > 0 Or InStr(1, mDireccion.Text, "É") > 0 Or InStr(1, mDireccion.Text, "Í") > 0 Or InStr(1, mDireccion.Text, "Ó") > 0 Or InStr(1, mDireccion.Text, "Ú") > 0 Or InStr(1, mDireccion.Text, "Ñ") > 0 Or InStr(1, mDireccion.Text, "á") > 0 Or InStr(1, mDireccion.Text, "é") > 0 Or InStr(1, mDireccion.Text, "í") > 0 Or InStr(1, mDireccion.Text, "ó") > 0 Or InStr(1, mDireccion.Text, "ú") > 0 Or InStr(1, mDireccion.Text, "ñ") > 0 Or InStr(1, mDireccion.Text, "ü") > 0 Or InStr(1, mDireccion.Text, "Ü") > 0 Or InStr(1, mDireccion.Text, "º") > 0 Then
        mDireccion.Text = "DIRECCION"
    End If
    
    If InStr(1, mLocalidad.Text, "Á") > 0 Or InStr(1, mLocalidad.Text, "É") > 0 Or InStr(1, mLocalidad.Text, "Í") > 0 Or InStr(1, mLocalidad.Text, "Ó") > 0 Or InStr(1, mLocalidad.Text, "Ú") > 0 Or InStr(1, mLocalidad.Text, "Ñ") > 0 Or InStr(1, mLocalidad.Text, "á") > 0 Or InStr(1, mLocalidad.Text, "é") > 0 Or InStr(1, mLocalidad.Text, "í") > 0 Or InStr(1, mLocalidad.Text, "ó") > 0 Or InStr(1, mLocalidad.Text, "ú") > 0 Or InStr(1, mLocalidad.Text, "ñ") > 0 Or InStr(1, mLocalidad.Text, "ü") > 0 Or InStr(1, mLocalidad.Text, "Ü") > 0 Or InStr(1, mLocalidad.Text, "º") > 0 Then
        mLocalidad.Text = "LOCALIDAD"
    End If
    If InStr(1, mProvincia.Text, "Á") > 0 Or InStr(1, mProvincia.Text, "É") > 0 Or InStr(1, mProvincia.Text, "Í") > 0 Or InStr(1, mProvincia.Text, "Ó") > 0 Or InStr(1, mProvincia.Text, "Ú") > 0 Or InStr(1, mProvincia.Text, "Ñ") > 0 Or InStr(1, mProvincia.Text, "á") > 0 Or InStr(1, mProvincia.Text, "é") > 0 Or InStr(1, mProvincia.Text, "í") > 0 Or InStr(1, mProvincia.Text, "ó") > 0 Or InStr(1, mProvincia.Text, "ú") > 0 Or InStr(1, mProvincia.Text, "ñ") > 0 Or InStr(1, mProvincia.Text, "ü") > 0 Or InStr(1, mProvincia.Text, "Ü") > 0 Or InStr(1, mProvincia.Text, "º") > 0 Then
        mProvincia.Text = "PROVINCIA"
    End If
    
    If InStr(1, txtRazSoc.Text, "Á") > 0 Or InStr(1, txtRazSoc.Text, "É") > 0 Or InStr(1, txtRazSoc.Text, "Í") > 0 Or InStr(1, txtRazSoc.Text, "Ó") > 0 Or InStr(1, txtRazSoc.Text, "Ú") > 0 Or InStr(1, txtRazSoc.Text, "Ñ") > 0 Or InStr(1, txtRazSoc.Text, "á") > 0 Or InStr(1, txtRazSoc.Text, "é") > 0 Or InStr(1, txtRazSoc.Text, "í") > 0 Or InStr(1, txtRazSoc.Text, "ó") > 0 Or InStr(1, txtRazSoc.Text, "ú") > 0 Or InStr(1, txtRazSoc.Text, "ñ") > 0 Or InStr(1, txtRazSoc.Text, "ü") > 0 Or InStr(1, txtRazSoc.Text, "Ü") > 0 Or InStr(1, txtRazSoc.Text, "º") > 0 Then
        txtRazSoc.Text = SacoAcento(Trim(txtRazSoc.Text))
    End If
    
    If cboFactura.ItemData(cboFactura.ListIndex) = 1 Then 'factura A
        'mRespuestaFiscal = pf.OpenInvoice("T", "C", "A", "1", "P", "12", "I", mRespo.Text, txtRazSoc.Text, "", "CUIT", txtCuit.Text, "N", Trim(mDireccion.Text), Trim(mLocalidad.Text), Trim(mProvincia.Text), "", "", "C")
        
        'original
        'EMULADOR mRespuestaFiscal = pf.OpenInvoice("T", "C", "A", "1", "P", "12", "I", mRespo.Text, txtRazSoc.Text, "A", "CUIT", txtCuit.Text, "N", Trim(mIngBrutos), Trim(mVendedor), "X", "X", "B", "G")
        mRespuestaFiscal = pf.OpenInvoice("M", "C", "A", "1", "P", "12", "I", mRespo.Text, txtRazSoc.Text, "", "CUIT", txtCuit.Text, "N", Trim(mIngBrutos), Trim(mVendedor), "", Trim(txtNroFactura), "", "G")
        'talampaya
        'mRespuestaFiscal = pf.OpenInvoice("T", "C", "A", "1", "P", "12", "I", mRespo.Text, txtRazSoc.Text, "", "CUIT", txtCuit.Text, "N", "", Trim(mVendedor), "", "", "", "C")
        
        If mRespuestaFiscal = False Then Exit Sub
    End If
    If cboFactura.ItemData(cboFactura.ListIndex) = 2 Then 'factura B
        If txtCiva.Text = "CONSUMIDOR FINAL" Then
            'ABRO UN TIKET FACTURA B PERO CON TIPO DE DOCUMENTO DNI
            If txtRazSoc.Text = "" Then
                txtRazSoc.Text = "CLIENTE"
            End If
            If txtNRO_DOCUMENTO.Text = "" Then
                If txtCuit.Text = "" Then
                    txtNRO_DOCUMENTO.Text = "11111111"
                Else
                    txtNRO_DOCUMENTO.Text = txtCuit.Text
                End If
            End If
            'mRespuestaFiscal = pf.OpenInvoice("T", "C", "B", "1", "P", "12", "I", mRespo.Text, txtRazSoc.Text, "", "DNI", txtNRO_DOCUMENTO.Text, "N", Trim(mDireccion.Text), Trim(mLocalidad.Text), Trim(mProvincia.Text), "", "", "C")
            'EMULADOR mRespuestaFiscal = pf.OpenInvoice("T", "C", "B", "1", "P", "12", "I", mRespo.Text, txtRazSoc.Text, "A", "DNI", txtNRO_DOCUMENTO.Text, "N", "Z", Trim(mVendedor), "X", "X", "B", "C")
            mRespuestaFiscal = pf.OpenInvoice("M", "C", "B", "1", "P", "12", "I", mRespo.Text, txtRazSoc.Text, "", "DNI", txtNRO_DOCUMENTO.Text, "N", "", Trim(mVendedor), "", "", "", "C")
            If mRespuestaFiscal = False Then Exit Sub
        Else
            'MONOTRIBUTO - ABRO UN TIKET FACTURA B PERO CON TIPO DE DOCUMENTO CUIT
            'mRespuestaFiscal = pf.OpenInvoice("T", "C", "B", "1", "P", "12", "I", mRespo.Text, txtRazSoc.Text, "", "CUIT", txtCuit.Text, "N", Trim(mDireccion.Text), Trim(mLocalidad.Text), Trim(mProvincia.Text), "", "", "C")
            
            'mRespuestaFiscal = pf.OpenInvoice("T", "C", "B", "1", "P", "12", "I", mRespo.Text, txtRazSoc.Text, "", "CUIT", txtCuit.Text, "N", "", Trim(mVendedor), "", "", "", "C")
            'mRespuestaFiscal = pf.OpenInvoice("T", "C", "C", "1", "P", "12", "I", mRespo.Text, txtRazSoc.Text, "", "CUIT", txtCuit.Text, "N", "", Trim(mVendedor), "", "", "", "C")
            mRespuestaFiscal = pf.OpenInvoice("M", "C", "C", "1", "P", "12", "I", mRespo.Text, txtRazSoc.Text, "", "CUIT", txtCuit.Text, "N", Trim(mIngBrutos), Trim(mVendedor), "", Trim(txtNroFactura), "", "G")
            If mRespuestaFiscal = False Then Exit Sub
        End If
    End If
    
    For i = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(i, 0) <> "" Then
            'ACA HAY QUE CALCULAR EL PORCENTAJE DE INCIDENCIA DE LOS IMP INTERNOS EN EL LITRO DE COMB
            
            If cboFactura.ItemData(cboFactura.ListIndex) = 1 Then
                mItc = 0
                mTasa = 0
                If grdGrilla.TextMatrix(i, 0) <> 3 Then  'NAFTA / gnc y demas (el  imp es 0)
                    'NAFTA Y GNC
                    'primero calcular el precio neto del combustible luego el itc y tasa
                    mItc = CDbl(grdGrilla.TextMatrix(i, 2)) * CDbl(grdGrilla.TextMatrix(i, 7))
                    mPneto = CDbl(grdGrilla.TextMatrix(i, 2)) * (CDbl(grdGrilla.TextMatrix(i, 3)) - CDbl(grdGrilla.TextMatrix(i, 11))) - mItc
                    mPneto = mPneto / (1 + (mIVA_1 / 100))
                    mItc = mItc / mPneto
                    mItc = Format(mItc, "0.00000000")
                    
                    mPneto = mPneto / CDbl(grdGrilla.TextMatrix(i, 2))
                    mPneto = Format(mPneto, "0.000")
                Else
                    'GASOIL
                    mItc = CDbl(grdGrilla.TextMatrix(i, 2)) * CDbl(grdGrilla.TextMatrix(i, 7))
                    mPneto = CDbl(grdGrilla.TextMatrix(i, 2)) * (CDbl(grdGrilla.TextMatrix(i, 3)) - CDbl(grdGrilla.TextMatrix(i, 11))) - mItc ' ESTOY RESTANDO LA TASA VIAL COL 11
                    mPneto = mPneto / (1 + (mIVA_2 / 100)) '
                    
                    mTasa = mPneto * ((mIVA_2 - mIVA_1) / 100) ' RESTAR LOS DOS IVAS (40-21)
                    
                    mItc = (mItc + mTasa) / mPneto
                    mItc = Format(mItc, "0.00000000")
                    
                    mPneto = mPneto / CDbl(grdGrilla.TextMatrix(i, 2))
                    mPneto = Format(mPneto, "0.000")
                
                End If
            Else
                'facturas B
                mItc = 0
                mTasa = 0
                If grdGrilla.TextMatrix(i, 0) <> 3 Then  'NAFTA / gnc y demas (el  imp es 0)
                    'NAFTA Y GNC
                    'primero calcular el precio neto del combustible luego el itc y tasa
                    mItc = CDbl(grdGrilla.TextMatrix(i, 2)) * CDbl(grdGrilla.TextMatrix(i, 7))
                    'mPneto = CDbl(grdGrilla.TextMatrix(I, 2)) * CDbl(grdGrilla.TextMatrix(I, 3)) - mItc
                    'mPneto = mPneto / (1 + (mIVA_1 / 100))
                    mPneto = Format(CDbl(grdGrilla.TextMatrix(i, 2)) * (CDbl(grdGrilla.TextMatrix(i, 3)) - CDbl(grdGrilla.TextMatrix(i, 11))), "0.000") ' ESTOY RESTANDO LA TASA VIAL COL 11
                    If mPneto = 0 Then
                        mItc = mPneto
                        mItc = Format(mItc, "0.00000000")
                    
                        mPneto = mPneto
                        mPneto = Format(mPneto, "0.000")
                    Else
                        mItc = mItc / mPneto
                        mItc = Format(mItc, "0.00000000")
                    
                        mPneto = mPneto / CDbl(grdGrilla.TextMatrix(i, 2))
                        mPneto = Format(mPneto, "0.000")
                    End If
                    
                Else
                    'GASOIL
                    mItc = CDbl(grdGrilla.TextMatrix(i, 2)) * CDbl(grdGrilla.TextMatrix(i, 7))
                    'mPneto = CDbl(grdGrilla.TextMatrix(I, 2)) * CDbl(grdGrilla.TextMatrix(I, 3)) - mItc
                    'mPneto = mPneto / (1 + (mIVA_2 / 100)) '
                    mPneto = Format(CDbl(grdGrilla.TextMatrix(i, 2)) * (CDbl(grdGrilla.TextMatrix(i, 3)) - CDbl(grdGrilla.TextMatrix(i, 11))), "0.000") ' ESTOY RESTANDO LA TASA VIAL COL 11
                    
                    mTasa = mPneto * ((mIVA_2 - mIVA_1) / 100) ' RESTAR LOS DOS IVAS (40-21)
                                       
                    
                    mItc = (mItc + mTasa) / mPneto
                    mItc = Format(mItc, "0.00000000")
                    
                    mPneto = mPneto / CDbl(grdGrilla.TextMatrix(i, 2))
                    mPneto = Format(mPneto, "0.000")
                
                End If
                
                
                'mItc = 0
                
                
                grdGrilla.TextMatrix(i, 6) = 0
            End If
            iCanti = Str(Format(CDbl(grdGrilla.TextMatrix(i, 2)), "0.00") * mContCanti)
            iPrecio = Str(mPneto * mContPrecio)
            iImpInt = Str(Format(mItc * CDbl(grdGrilla.TextMatrix(i, 2)), "0.00000000") * mContInternos)
            'iiMpInt = Str(Format(CDbl(iiMpInt), "0.00") * mContInternos)
            mIvaFE = Str(mIVA_1 * 100)
            'mIvaFE = Str(CDbl(grdGrilla.TextMatrix(I, 6)) * 100)
            
            If cboFactura.ItemData(cboFactura.ListIndex) = 10000 Then  '"T" TIKET
                mRespuestaFiscal = pf.SendTicketItem(Trim(ChkNull(grdGrilla.TextMatrix(i, 1))), Trim(iCanti), Trim(iPrecio), Trim(mIvaFE), "M", "0", "0", Trim(iImpInt))
                If mRespuestaFiscal = False Then Exit Sub
            Else
                'mRespuestaFiscal = pf.SendInvoiceItem(Trim(ChkNull(grdGrilla.TextMatrix(i, 1))), Trim(iCanti), Trim(iPrecio), Trim(mIvaFE), "M", "0", "0", ChkNull(grdGrilla.TextMatrix(i, 0)) , "", "", "", Trim(iImpInt))
                mRespuestaFiscal = pf.SendInvoiceItem(Trim(ChkNull(grdGrilla.TextMatrix(i, 1))), Trim(iPrecio), Trim(iCanti), Trim(mIvaFE), "M", "0", "0", "", "", "", "", Trim(iImpInt))
                
                'TICKET
                'mRespuestaFiscal = pf.SendTicketItem(Trim(ChkNull(GRDGrilla.TextMatrix(i, 1))), Trim(iCanti), Trim(iPrecio), Trim(mIvaFE), "M", "0", "0")
                If mRespuestaFiscal = False Then Exit Sub
            End If
            'If txttasavial.Text <> "0,000" Then
            '    mTVialDbl = CDbl(grdGrilla.TextMatrix(i, 11))
            '    mTasaVial = Str(mTVialDbl * mContPrecio)
            '    iCanti = Str(Format(CDbl(grdGrilla.TextMatrix(i, 2)), "0.00") * mContCanti)
            '    'iPrecio = Str(mPneto * mContPrecio)
            '    If cboFactura.ItemData(cboFactura.ListIndex) = 10000 Then  '"T" TIKET
            '        'mRespuestaFiscal = pf.SendTicketItem(Trim(ChkNull(grdGrilla.TextMatrix(I, 1))), Trim(iCanti), Trim(iPrecio), Trim(mIvaFE), "M", "0", "0", Trim(iImpInt))
            '        mRespuestaFiscal = pf.SendTicketItem("Tasa Vial", Trim(iCanti), Trim(mTasaVial), "0", "M", "0", "0", "")
            '        If mRespuestaFiscal = False Then Exit Sub
            '    Else
            '        mRespuestaFiscal = pf.SendInvoiceItem("Tasa Vial", Trim(mTasaVial), Trim(iCanti), "", "M", "0", "0", "", "", "", "", "")
            '        If mRespuestaFiscal = False Then Exit Sub
            '    End If
            'End If
            
        End If
    Next
    
   
    'PAGOS
    If cboFactura.ItemData(cboFactura.ListIndex) = 10000 Then 'TIKET Then
        mRespuestaFiscal = pf.GetTicketSubtotal("P", "SUBTOTAL")
        If mRespuestaFiscal = False Then Exit Sub
    End If
    If cboFactura.ItemData(cboFactura.ListIndex) = 1 Then 'factura A Then
        mRespuestaFiscal = pf.GetInvoiceSubtotal("P", "SUBTOTAL")
        'ticket
        'mRespuestaFiscal = pf.GetTicketSubtotal("P", "SUBTOTAL")
        
        If mRespuestaFiscal = False Then Exit Sub
        
    End If
    If cboFactura.ItemData(cboFactura.ListIndex) = 2 Then 'factura B Then
        mRespuestaFiscal = pf.GetInvoiceSubtotal("P", "SUBTOTAL")
        If mRespuestaFiscal = False Then Exit Sub
    End If
    
    For i = 1 To grdPagos.Rows - 1
        mTotaLl = CDbl(grdPagos.TextMatrix(i, 1)) * 100
        ITOTAL = Str(mTotaLl)
        If cboFactura.ItemData(cboFactura.ListIndex) = 10000 Then 'TIKET Then
            mRespuestaFiscal = pf.SendTicketPayment(Mid(grdPagos.TextMatrix(i, 0), 1, 20), Trim(ITOTAL), "T")
            
            If mRespuestaFiscal = False Then Exit Sub
        End If
         If cboFactura.ItemData(cboFactura.ListIndex) = 1 Then 'factura A
            mRespuestaFiscal = pf.SendInvoicePayment(Mid(grdPagos.TextMatrix(i, 0), 1, 20), Trim(ITOTAL), "T")
            'ticket
            'mRespuestaFiscal = pf.SendTicketPayment(Mid(grdPagos.TextMatrix(i, 0), 1, 20), Trim(ITOTAL), "T")
            If mRespuestaFiscal = False Then Exit Sub
        End If
        If cboFactura.ItemData(cboFactura.ListIndex) = 2 Then 'factura B
            mRespuestaFiscal = pf.SendInvoicePayment(Mid(grdPagos.TextMatrix(i, 0), 1, 20), Trim(ITOTAL), "T")
            If mRespuestaFiscal = False Then Exit Sub
        End If
    Next


'CIERRO COMPROBANTE


    If cboFactura.ItemData(cboFactura.ListIndex) = 1 Then 'nota de credito A
        'pf.GetInvoiceSubtotal "P", "SUBTOTAL"
        txtTotalFiscal.Text = Round(CDbl(Chk0(pf.AnswerField_5)) / 100, 2)
        txtIvaFiscal.Text = Round(CDbl(Chk0(pf.AnswerField_6)) / 100, 2)
        txtNetoFiscal.Text = Round(CDbl(Chk0(pf.AnswerField_10)) / 100, 2)
        mRespuestaFiscal = pf.CloseInvoice("M", "A", "TOTAL")
        If mRespuestaFiscal = False Then Exit Sub
    End If
    If cboFactura.ItemData(cboFactura.ListIndex) = 2 Then 'nota de credito B
        'pf.GetInvoiceSubtotal "P", "SUBTOTAL"
        txtTotalFiscal.Text = Round(CDbl(Chk0(pf.AnswerField_5)) / 100, 2)
        txtIvaFiscal.Text = Round(CDbl(Chk0(pf.AnswerField_6)) / 100, 2)
        txtNetoFiscal.Text = Round(CDbl(Chk0(pf.AnswerField_10)) / 100, 2)
        mRespuestaFiscal = pf.CloseInvoice("M", "B", "TOTAL")
        If mRespuestaFiscal = False Then Exit Sub
    End If

    CmdNuevo_Click
End Sub
Private Function ImprimoFiscalEpsondll(tipocbte As Integer)
    '"   1 - Tique.
    '"   2 - Tique factura A/B/C/M.
    '"   3 - Tique nota de crédito, tique nota crédito A/B/C/M.
    '"   4 - Tique nota de débito A/B/C/M.
    '"   21 - Documento no fiscal homologado genérico.
    '"   22 - Documento no fiscal homologado de uso interno.
 
    Dim mContCanti As Integer
    Dim mContPrecio As Integer
    Dim mContInternos As Double
    Dim mVendedor As String
    Dim mIngBrutos As String
    Dim mItc As Double
    Dim mPneto As Double
    Dim mTasa As Double
    Dim mTVialDbl As Double
    
    Dim mTotaLl As Double
    Dim ITOTAL As String

    Dim iCanti As String
    Dim iPrecio As String
    Dim iImpInt As String
    Dim mIvaFE As String
    Dim mTasaVial As String
    Dim mUnidadMedida As Integer
    
'    mContCanti = 1000
'    mContPrecio = 100
    mIVA_1 = BuscoIva
    mIVA_2 = BuscoIva_2
    
    'Lo cambio para el redondeo de combustibles
    'mContCanti = "1.000"
    'mContPrecio = "100.000"
    'mContInternos = "1000000.0000"
    'mContInternos = 1000
    
    mVendedor = "Playero: " & cboVendedor.List(cboVendedor.ListIndex)
    mVendedor = Mid(mVendedor, 1, 20)
    
    mIngBrutos = "Ing Brutos: " & IIf(txtIngBrutos.Text = "" Or txtIngBrutos.Text = "0", "NO POSEE", txtIngBrutos.Text)
    mIngBrutos = Mid(mIngBrutos, 1, 25)
    
    mDireccion.Text = IIf(txtDomici.Text = "", "NO INFORMADO", txtDomici.Text)
    If InStr(1, mDireccion.Text, "Á") > 0 Or InStr(1, mDireccion.Text, "É") > 0 Or InStr(1, mDireccion.Text, "Í") > 0 Or InStr(1, mDireccion.Text, "Ó") > 0 Or InStr(1, mDireccion.Text, "Ú") > 0 Or InStr(1, mDireccion.Text, "Ñ") > 0 Or InStr(1, mDireccion.Text, "á") > 0 Or InStr(1, mDireccion.Text, "é") > 0 Or InStr(1, mDireccion.Text, "í") > 0 Or InStr(1, mDireccion.Text, "ó") > 0 Or InStr(1, mDireccion.Text, "ú") > 0 Or InStr(1, mDireccion.Text, "ñ") > 0 Or InStr(1, mDireccion.Text, "ü") > 0 Or InStr(1, mDireccion.Text, "Ü") > 0 Or InStr(1, mDireccion.Text, "º") > 0 Then
        mDireccion.Text = "DIRECCION"
    End If
    
    If InStr(1, mLocalidad.Text, "Á") > 0 Or InStr(1, mLocalidad.Text, "É") > 0 Or InStr(1, mLocalidad.Text, "Í") > 0 Or InStr(1, mLocalidad.Text, "Ó") > 0 Or InStr(1, mLocalidad.Text, "Ú") > 0 Or InStr(1, mLocalidad.Text, "Ñ") > 0 Or InStr(1, mLocalidad.Text, "á") > 0 Or InStr(1, mLocalidad.Text, "é") > 0 Or InStr(1, mLocalidad.Text, "í") > 0 Or InStr(1, mLocalidad.Text, "ó") > 0 Or InStr(1, mLocalidad.Text, "ú") > 0 Or InStr(1, mLocalidad.Text, "ñ") > 0 Or InStr(1, mLocalidad.Text, "ü") > 0 Or InStr(1, mLocalidad.Text, "Ü") > 0 Or InStr(1, mLocalidad.Text, "º") > 0 Then
        mLocalidad.Text = "LOCALIDAD"
    End If
    If InStr(1, mProvincia.Text, "Á") > 0 Or InStr(1, mProvincia.Text, "É") > 0 Or InStr(1, mProvincia.Text, "Í") > 0 Or InStr(1, mProvincia.Text, "Ó") > 0 Or InStr(1, mProvincia.Text, "Ú") > 0 Or InStr(1, mProvincia.Text, "Ñ") > 0 Or InStr(1, mProvincia.Text, "á") > 0 Or InStr(1, mProvincia.Text, "é") > 0 Or InStr(1, mProvincia.Text, "í") > 0 Or InStr(1, mProvincia.Text, "ó") > 0 Or InStr(1, mProvincia.Text, "ú") > 0 Or InStr(1, mProvincia.Text, "ñ") > 0 Or InStr(1, mProvincia.Text, "ü") > 0 Or InStr(1, mProvincia.Text, "Ü") > 0 Or InStr(1, mProvincia.Text, "º") > 0 Then
        mProvincia.Text = "PROVINCIA"
    End If
    
    If InStr(1, txtRazSoc.Text, "Á") > 0 Or InStr(1, txtRazSoc.Text, "É") > 0 Or InStr(1, txtRazSoc.Text, "Í") > 0 Or InStr(1, txtRazSoc.Text, "Ó") > 0 Or InStr(1, txtRazSoc.Text, "Ú") > 0 Or InStr(1, txtRazSoc.Text, "Ñ") > 0 Or InStr(1, txtRazSoc.Text, "á") > 0 Or InStr(1, txtRazSoc.Text, "é") > 0 Or InStr(1, txtRazSoc.Text, "í") > 0 Or InStr(1, txtRazSoc.Text, "ó") > 0 Or InStr(1, txtRazSoc.Text, "ú") > 0 Or InStr(1, txtRazSoc.Text, "ñ") > 0 Or InStr(1, txtRazSoc.Text, "ü") > 0 Or InStr(1, txtRazSoc.Text, "Ü") > 0 Or InStr(1, txtRazSoc.Text, "º") > 0 Then
        txtRazSoc.Text = SacoAcento(Trim(txtRazSoc.Text))
    End If
    
    If txtNRO_DOCUMENTO.Text = "" Or txtNRO_DOCUMENTO.Text = "0" Then
        If txtCuit.Text = "" Then
            txtNRO_DOCUMENTO.Text = "11111111"
        Else
            txtNRO_DOCUMENTO.Text = txtCuit.Text
        End If
    End If



    Const ID_MODIFICADOR_AGREGAR_ITEM  As Long = 200
    Const ID_IMPUESTO_NINGUNO As Long = 1
    Const ID_CODIGO_INTERNO  As Long = 1
    Const AFIP_CODIGO_UNIDAD_MEDIDA_KILOGRAMO As Long = 1
     
    Dim msg
    Dim Error As Long
    Dim str_comprobante_numero As String
    Dim str_comprobante_tipo As String
   

    'connect
    Error = conectar_impresora()
    If Error <> ERROR_NINGUNO Then msg = MsgBox(Error, vbOKOnly, "Error: Conectar()")
    ' open
    'CONTINUAR ACA, VER EN DOCUMENTACION EL id_tipo_documento Y id_responsabilidad_iva
  
    If cboFactura.ItemData(cboFactura.ListIndex) = 1 Then 'factura A
        Error = CargarDatosCliente(txtRazSoc.Text, "", txtDomici.Text, "", "", 3, txtCuit.Text, 1)
        If Error <> ERROR_NINGUNO Then msg = MsgBox(Error, vbOKOnly, "Error: CargarDatosCliente()")
    End If
    If cboFactura.ItemData(cboFactura.ListIndex) = 2 Then 'factura B
        If txtCiva.Text = "EXENTO" Then
            Error = CargarDatosCliente(txtRazSoc.Text, "", mDireccion.Text, "", "", 3, txtCuit, 6)
            If Error <> ERROR_NINGUNO Then msg = MsgBox(Error, vbOKOnly, "Error: CargarDatosCliente()")
        ElseIf txtCiva.Text = "MONOTRIBUTO" Then
            Error = CargarDatosCliente(txtRazSoc.Text, "", mDireccion.Text, "", "", 3, txtCuit, 4)
            If Error <> ERROR_NINGUNO Then msg = MsgBox(Error, vbOKOnly, "Error: CargarDatosCliente()")
        ElseIf txtCiva.Text = "CONSUMIDOR FINAL" Then
            Error = CargarDatosCliente(txtRazSoc.Text, "", mDireccion.Text, "", "", 1, txtNRO_DOCUMENTO, 5)
            'Error = CargarDatosCliente("Nombre Comprador #1", "", "Domicilio Comparador #1", "", "", 1, "3478905", 5)
            If Error <> ERROR_NINGUNO Then msg = MsgBox(Error, vbOKOnly, "Error: CargarDatosCliente()")
        End If
        
    End If
    Error = CargarComprobanteAsociado("083-00001-00000027")
    'Dim cbte_relacionado As String
    'cbte_relacionado = "083-000001-" & txtNroFactura.Text
    
    'Error = CargarComprobanteAsociado(cbte_relacionado)
    If Error <> ERROR_NINGUNO Then msg = MsgBox(Error, vbOKOnly, "Error: CargarComprobanteAsociado()")
   
    Error = AbrirComprobante(tipocbte)
    If Error <> ERROR_NINGUNO Then msg = MsgBox(Error, vbOKOnly, "Error: AbrirComprobante()")
        
    ' consultar numero y tipo de comprobante actual
    str_comprobante_numero = String(60, vbNullChar)
    Error = ConsultarNumeroComprobanteActual(str_comprobante_numero, Len(str_comprobante_numero))
    If Error <> ERROR_NINGUNO Then msg = MsgBox(Error, vbOKOnly, "Error: ConsultarNumeroComprobanteActual()")
'    If Error = ERROR_NINGUNO Then
'        msg = MsgBox(str_comprobante_numero, vbOKOnly, "ConsultarNumeroComprobanteActual()")
'    Else
'        msg = MsgBox(Error, vbOKOnly, "Error: ConsultarNumeroComprobanteActual()")
'    End If
  
    str_comprobante_tipo = String(60, vbNullChar)
    Error = ConsultarTipoComprobanteActual(str_comprobante_tipo, Len(str_comprobante_tipo))
    If Error <> ERROR_NINGUNO Then msg = MsgBox(Error, vbOKOnly, "Error: ConsultarTipoComprobanteActual()")
'    If Error = ERROR_NINGUNO Then
'      msg = MsgBox(str_comprobante_tipo, vbOKOnly, "ConsultarTipoComprobanteActual()")
'    Else
'      msg = MsgBox(Error, vbOKOnly, "Error: ConsultarTipoComprobanteActual()")
'    End If
   
    
    ' item linea de descripcion extra
    'Error = CargarTextoExtra("Texto extra #1")
    'If Error <> 0 Then msg = MsgBox(Error, vbOKOnly, "Error: CargarTextoExtra() #1")
    
    'Error = CargarTextoExtra("Texto extra #2")
    'If Error <> 0 Then msg = MsgBox(Error, vbOKOnly, "Error: CargarTextoExtra() #2")
    
    For i = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(i, 0) <> "" Then
            'ACA HAY QUE CALCULAR EL PORCENTAJE DE INCIDENCIA DE LOS IMP INTERNOS EN EL LITRO DE COMB
            'defino unidad de medida 1,3,78,81,84,90
            
            If grdGrilla.TextMatrix(i, 0) = 1 Or grdGrilla.TextMatrix(i, 0) = 3 Or grdGrilla.TextMatrix(i, 0) = 78 _
                Or grdGrilla.TextMatrix(i, 0) = 81 Or grdGrilla.TextMatrix(i, 0) = 84 Or grdGrilla.TextMatrix(i, 0) = 90 Then
                mUnidadMedida = 5 'LITROS
            ElseIf grdGrilla.TextMatrix(i, 0) = 2 Or grdGrilla.TextMatrix(i, 0) = 4 Then
                mUnidadMedida = 4 'METROS CUBICOS
            Else
                mUnidadMedida = 7 'UNIDAD
            End If
            
            mItc = 0
            'primero calcular el precio neto del combustible luego el itc y tasa
            'mItc = CDbl(grdGrilla.TextMatrix(i, 2)) * CDbl(grdGrilla.TextMatrix(i, 7))
            mItc = CDbl(grdGrilla.TextMatrix(i, 7))
            mPneto = CDbl(grdGrilla.TextMatrix(i, 3)) - mItc
            mPneto = mPneto / (1 + (mIVA_1 / 100))
            grdGrilla.TextMatrix(i, 6) = 0
            
            iCanti = Format(grdGrilla.TextMatrix(i, 2), "00000.0000")
            iPrecio = Format(mPneto, "0000000.0000") 'Str(mPneto * mContPrecio)
            iImpInt = Format(mItc, "0000000.0000")
            
            iCanti = Replace(iCanti, ",", ".")
            iPrecio = Replace(iPrecio, ",", ".")
            iImpInt = Replace(iImpInt, ",", ".")
                        
            Error = ImprimirItem(ID_MODIFICADOR_AGREGAR_ITEM, Trim(ChkNull(grdGrilla.TextMatrix(i, 1))), Trim(iCanti), Trim(iPrecio), ID_TASA_IVA_21_00, ID_IMPUESTO_NINGUNO, Trim(iImpInt), ID_CODIGO_INTERNO, ChkNull(grdGrilla.TextMatrix(i, 0)), "", mUnidadMedida)
            If Error <> 0 Then msg = MsgBox(Error, vbOKOnly, "Error: ImprimirItem()")
            
            
        End If
    Next
    
    ' close
    Error = CerrarComprobante()
    If Error <> ERROR_NINGUNO Then msg = MsgBox(Error, vbOKOnly, "Error: CerrarComprobante()")
   
    '' cancelar
    ''error = Cancelar()
    ''msg = MsgBox(error, vbOKOnly, "Error: Cancelar()")
    
    
    ' close port
    Error = Desconectar()
    If Error <> 0 Then msg = MsgBox(Error, vbOKOnly, "Error: Desconectar()")
    
    Screen.MousePointer = vbNormal
    
End Function
