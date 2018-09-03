VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm Menu 
   BackColor       =   &H8000000C&
   Caption         =   "Sistema Gestión"
   ClientHeight    =   5310
   ClientLeft      =   105
   ClientTop       =   2085
   ClientWidth     =   9480
   Icon            =   "Menu.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar tbrPrincipal 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   10
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Planilla Playeros"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Facturacion"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Lista de Precios"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Clientes"
            Object.Tag             =   ""
            ImageIndex      =   18
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Cobranza"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Listado Facturas por Vendedor"
            Object.Tag             =   ""
            ImageIndex      =   16
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Organizar Ventanas"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   25
      Left            =   600
      Top             =   2640
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   9450
      TabIndex        =   2
      Top             =   420
      Width           =   9480
      Begin VB.CommandButton Command1 
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   350
         Left            =   240
         ScaleHeight     =   345
         ScaleWidth      =   11535
         TabIndex        =   3
         Top             =   120
         Width           =   11535
      End
   End
   Begin ComctlLib.StatusBar stbPrincipal 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   4995
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   556
      SimpleText      =   "Listo."
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   6
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   2
            Object.Width           =   6526
            MinWidth        =   6526
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   7673
            MinWidth        =   7673
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "NÚM"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Alignment       =   1
            Bevel           =   2
            Enabled         =   0   'False
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "MAYÚS"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "11:21"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1940
            MinWidth        =   1940
            TextSave        =   "16/08/2018"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
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
   Begin ComctlLib.ImageList ImageList1 
      Left            =   135
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   20
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":06A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":0EBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":16D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":1EF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":270C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":2F26
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":3100
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":341A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":3734
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":3A4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":3D68
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":4082
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":439C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":46B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":49D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":4CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":5004
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":531E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":5638
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuArc 
      Caption         =   "&Archivos"
      Begin VB.Menu mnuArchivoActualizaciones 
         Caption         =   "Actualizaciones"
         Begin VB.Menu mnuABMPais 
            Caption         =   "País"
         End
         Begin VB.Menu mnuABMProvincias 
            Caption         =   "Provincias"
         End
         Begin VB.Menu mnuABMLocalidades 
            Caption         =   "Localidades"
         End
         Begin VB.Menu mnuFacturacionTipoComprobante 
            Caption         =   "Tipo de Comprobante"
         End
         Begin VB.Menu mnuABMInscIVA 
            Caption         =   "Condición &IVA"
         End
         Begin VB.Menu mnuABMEstadoDocumento 
            Caption         =   "Estado de Documentos"
         End
         Begin VB.Menu mnuABMFormaPago 
            Caption         =   "Forma de Pago"
         End
         Begin VB.Menu mnuTarjetaCredito 
            Caption         =   "Tarjeta de Crédito"
         End
         Begin VB.Menu mnuTarjetaPlan 
            Caption         =   "Tarjeta Plan"
         End
         Begin VB.Menu mnuSucursal 
            Caption         =   "Sucursal"
         End
         Begin VB.Menu mnuBancos 
            Caption         =   "Bancos"
         End
      End
      Begin VB.Menu mnuClave 
         Caption         =   "Clave"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRaya11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConectar 
         Caption         =   "Conectar"
      End
      Begin VB.Menu mnuDesconectar 
         Caption         =   "Desconectar"
      End
      Begin VB.Menu mnuRaya4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuParametros 
         Caption         =   "&Parámetros"
      End
      Begin VB.Menu MNURAYA100 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUsuarios 
         Caption         =   "&Usuarios"
      End
      Begin VB.Menu mnuPermisos 
         Caption         =   "Permi&sos"
      End
      Begin VB.Menu mnuRaya3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCierreX 
         Caption         =   "Cierre X"
      End
      Begin VB.Menu mnuCierreZ 
         Caption         =   "Cierre Z"
      End
      Begin VB.Menu mnurayacierres 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArcSal 
         Caption         =   "Sali&r"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuraya29 
      Caption         =   "&Productos"
      Begin VB.Menu mnuABMLineas 
         Caption         =   "Líneas"
      End
      Begin VB.Menu mnuABMRubros 
         Caption         =   "Rubros"
      End
      Begin VB.Menu mnuMarcas 
         Caption         =   "Marcas"
      End
      Begin VB.Menu mnuRayaProd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStockABMProductos 
         Caption         =   "Productos"
      End
      Begin VB.Menu mnuRayaEstadoProd 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnulista 
         Caption         =   "Lista de Precios"
      End
      Begin VB.Menu mnuEstadoProductos 
         Caption         =   "Estado Productos"
         Visible         =   0   'False
      End
      Begin VB.Menu mnurayastock 
         Caption         =   "-"
      End
      Begin VB.Menu mnuentrada 
         Caption         =   "Entrada De Productos"
      End
      Begin VB.Menu mnuAjuste 
         Caption         =   "Ajuste de Stock"
      End
      Begin VB.Menu mnulistado 
         Caption         =   "Listado"
         Begin VB.Menu mnuEstaCantidadVendida 
            Caption         =   "Surtidor"
         End
      End
   End
   Begin VB.Menu mnuVentasFacturacion 
      Caption         =   "&Ventas"
      Begin VB.Menu mnuFacturaActualiza 
         Caption         =   "&Actualizaciones"
         Begin VB.Menu mnuABMClientes 
            Caption         =   "&Clientes"
         End
         Begin VB.Menu mnuABMVendedores 
            Caption         =   "&Vendedores"
         End
         Begin VB.Menu mnuTipoRevelado 
            Caption         =   "&Tipo Revelado"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuDestinos 
            Caption         =   "&Destinos"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuAparatos 
            Caption         =   "&Aparatos"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuRayaCompostura 
         Caption         =   "-"
      End
      Begin VB.Menu mnuComposturas 
         Caption         =   "Composturas"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRevelados 
         Caption         =   "Revelados"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFacturacionFacturacion 
         Caption         =   "&Facturación"
      End
      Begin VB.Menu mnuNC 
         Caption         =   "&Nota de Credito"
      End
      Begin VB.Menu mnuraya2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReciboClientes 
         Caption         =   "&Ingreso de Cobranza"
      End
      Begin VB.Menu mnuRayaRecibo 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCtaCteClientes 
         Caption         =   "Cuenta Corriente de Clientes"
      End
      Begin VB.Menu mnuRayaCtaCte 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConsultaAnulaciones 
         Caption         =   "C&onsulta - Anulaciones"
         Begin VB.Menu mnuConAnuFactura 
            Caption         =   "... de Factura"
         End
         Begin VB.Menu mnuAnulaRecibos 
            Caption         =   "... de Recibos"
         End
      End
      Begin VB.Menu mnuRayaListados 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLibroDebitoFiscal 
         Caption         =   "Libro Débito Fiscal"
      End
      Begin VB.Menu mnuRaya10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRendTasaVial 
         Caption         =   "Rendicion Tasa Vial"
      End
      Begin VB.Menu mnurayaRend 
         Caption         =   "-"
      End
      Begin VB.Menu Controles 
         Caption         =   "Controles"
         Begin VB.Menu Tarjetas 
            Caption         =   "Tarjetas"
         End
         Begin VB.Menu mnudepobanc 
            Caption         =   "Depositos bancarios"
         End
      End
      Begin VB.Menu mnurayaControles 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVentasListados 
         Caption         =   "&Listados"
         Begin VB.Menu mnuCantVend 
            Caption         =   "Cantidades Vendidas"
         End
         Begin VB.Menu mnuListadoVentaPorVendedor 
            Caption         =   "Ventas por Playero"
         End
         Begin VB.Menu mnurayalistVtas 
            Caption         =   "-"
         End
         Begin VB.Menu mnuInformeSecEn 
            Caption         =   "Informe para Secretaria de Energia"
         End
         Begin VB.Menu mnuPlanillaVtas 
            Caption         =   "Planilla de Ventas"
            Begin VB.Menu mnuCajaS 
               Caption         =   "Caja/Stock"
            End
            Begin VB.Menu mnuFac 
               Caption         =   "Facturas"
            End
         End
      End
   End
   Begin VB.Menu mnuCompras 
      Caption         =   "&Compras"
      Begin VB.Menu mnuComprasActualiza 
         Caption         =   "&Actualizaciones"
         Begin VB.Menu mnuABMTipoProveedores 
            Caption         =   "Tipo de Proveedores"
         End
         Begin VB.Menu mnuABMProveedores 
            Caption         =   "Proveedores"
         End
         Begin VB.Menu mnuABMTipoGastos 
            Caption         =   "Tipo de Gastos"
         End
      End
      Begin VB.Menu mnuProveedores 
         Caption         =   "&Proveedores"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuProveedoresFacturas 
         Caption         =   "&Facturas"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGastos 
         Caption         =   "&Gastos Generales"
      End
      Begin VB.Menu mnuRaya12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLibroCreditoFiscal 
         Caption         =   "Libro Crédito Fiscal"
      End
      Begin VB.Menu mnurayaListComp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuComprasListado 
         Caption         =   "&Listado de Proveedores"
      End
   End
   Begin VB.Menu mnuGestionStock 
      Caption         =   "&Gestión de Stock"
      Visible         =   0   'False
      Begin VB.Menu mnuStockAjuste 
         Caption         =   "Control de &Stock"
      End
      Begin VB.Menu mnuEntradaMercaderia 
         Caption         =   "Movimiento de Mercadería"
      End
      Begin VB.Menu mnuAjusteStock 
         Caption         =   "Ajuste de Stock"
      End
      Begin VB.Menu mnuRaya30 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListaPrecios 
         Caption         =   "Lista de &Precios"
      End
      Begin VB.Menu mnuListadoProductos 
         Caption         =   "Listado de Productos"
      End
   End
   Begin VB.Menu mnuVentana 
      Caption         =   "&Ventana"
      Visible         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu mnuMosHoriz 
         Caption         =   "Mosaico &horizontal"
      End
      Begin VB.Menu mnuMosVert 
         Caption         =   "Mosaico &vertical"
      End
      Begin VB.Menu mnuCascada 
         Caption         =   "&Cascada"
      End
      Begin VB.Menu mnuIconos 
         Caption         =   "Organizar &Iconos"
      End
   End
   Begin VB.Menu mnuFondos 
      Caption         =   "Fondos"
      Begin VB.Menu mnuabmbanco 
         Caption         =   "ABM de Banco"
      End
      Begin VB.Menu mnuCheques 
         Caption         =   "Carga de Cheques"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuChequesTerceros 
         Caption         =   "Carga de Cheques"
      End
      Begin VB.Menu mnurayaFondos 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListaCheques 
         Caption         =   "Listado de Cheques"
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "Mantenimiento"
      Begin VB.Menu mnubkpArchivos 
         Caption         =   "Backup de Archivos"
      End
      Begin VB.Menu mnurestArchivos 
         Caption         =   "Restaurar Archivos"
      End
      Begin VB.Menu mnuContenido 
         Caption         =   "&Contenido"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAAcerca 
         Caption         =   "&Acerca de..."
      End
   End
   Begin VB.Menu ContextBaseABM 
      Caption         =   "ContextBaseABM"
      Visible         =   0   'False
      Begin VB.Menu mnuContextABM 
         Caption         =   "Nuevo"
         Index           =   0
      End
      Begin VB.Menu mnuContextABM 
         Caption         =   "Editar"
         Index           =   1
      End
      Begin VB.Menu mnuContextABM 
         Caption         =   "Eliminar"
         Index           =   2
      End
      Begin VB.Menu mnuContextABM 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuContextABM 
         Caption         =   "Refrescar"
         Index           =   4
      End
      Begin VB.Menu mnuContextABM 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuContextABM 
         Caption         =   "Buscar"
         Index           =   6
      End
      Begin VB.Menu mnuContextABM 
         Caption         =   "Imprimir"
         Index           =   7
      End
      Begin VB.Menu mnuContextABM 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuContextABM 
         Caption         =   "Ver Datos"
         Index           =   9
      End
   End
   Begin VB.Menu ContextABMCta 
      Caption         =   "ContextABMCta"
      Visible         =   0   'False
      Begin VB.Menu mnuContextABMCta 
         Caption         =   "Nuevo"
         Index           =   0
      End
      Begin VB.Menu mnuContextABMCta 
         Caption         =   "Editar"
         Index           =   1
      End
      Begin VB.Menu mnuContextABMCta 
         Caption         =   "Eliminar"
         Index           =   2
      End
      Begin VB.Menu mnuContextABMCta 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuContextABMCta 
         Caption         =   "Refrescar"
         Index           =   4
      End
      Begin VB.Menu mnuContextABMCta 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuContextABMCta 
         Caption         =   "Ver Datos"
         Index           =   6
      End
   End
   Begin VB.Menu ContextABMPresu 
      Caption         =   "ContextABMPresu"
      Visible         =   0   'False
      Begin VB.Menu mnuContextABMPresu 
         Caption         =   "Nuevo"
         Index           =   0
      End
      Begin VB.Menu mnuContextABMPresu 
         Caption         =   "Editar"
         Index           =   1
      End
      Begin VB.Menu mnuContextABMPresu 
         Caption         =   "Eliminar"
         Index           =   2
      End
      Begin VB.Menu mnuContextABMPresu 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuContextABMPresu 
         Caption         =   "Refrescar"
         Index           =   4
      End
      Begin VB.Menu mnuContextABMPresu 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuContextABMPresu 
         Caption         =   "Ver Datos"
         Index           =   6
      End
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TituloPrincipal As String

Private Declare Function ShellAbout Lib "shell32.dll" Alias _
"ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, _
ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Dim Letrero As String


Private Sub Command1_Click()
    frmPlanillaStock.Show
End Sub

Private Sub MDIForm_Load()
    'If Dir("c:\windows\cpce.ini") = "" Then
    '    Menu.Picture = LoadPicture(App.Path & "\fotos\Demaría.bmp")
    'End If
    
    TituloPrincipal = TIT_MSGBOX '"Sistema de Gestión y Administración"
    Me.Caption = TituloPrincipal
    
    Me.Show
    FrmInicio.Show vbModal
    
    'Me.Caption = TituloPrincipal & " - (Usuario " & UCase(mNomUser) & " conectado a " & UCase(SERVIDOR) & " - " & UCase(BASEDATO) & ")"
    Me.Caption = TituloPrincipal & "    V. " & App.Major & "." & App.Minor & "." & App.Revision & "          - (Usuario " & UCase(mNomUser) & " conectado a " & UCase(SERVIDOR) & " - " & UCase(BASEDATO) & ")"
    Menu.mnuConectar.Enabled = False
        
    frmFacturaCliente.Show
    configuroLetrero Date
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Call mnuArcSal_Click
End Sub

Private Sub mnuAAcerca_Click()
    Call ShellAbout(Me.hWnd, "Sistema de Gestión Administrativo", "Copyright 2010, DANIEL", Me.Icon)
End Sub

Private Sub mnuabmbanco_Click()
    ABMBanco.Show
End Sub

Private Sub mnuABMClientes_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMClientes = New CListaBaseABM
    
    With vABMClientes
        .Caption = "Actualización de Clientes"
        .sql = "SELECT CLI_RAZSOC, CLI_CODIGO, CLI_DOMICI, CLI_TELEFONO, CLI_MAIL" & _
               " FROM CLIENTE"
        .HeaderSQL = "Razón Social, Código, Domicilio, Teléfono, e-mail"
        .FieldID = "CLI_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormClientes
        Set .FormDatos = ABMClientes
    End With
    
    Set auxDllActiva = vABMClientes
    
    vABMClientes.Show
End Sub

Private Sub mnuABMEstadoDocumento_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMEstadoDocumento = New CListaBaseABM
    
    With vABMEstadoDocumento
        .Caption = "Actualización de Estado Documento"
        .sql = "SELECT EST_DESCRI, EST_CODIGO FROM ESTADO_DOCUMENTO"
        .HeaderSQL = "Descripción, Código"
        .FieldID = "EST_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormEstadoDocumento
        Set .FormDatos = ABMEstadoDocumento
    End With
    
    Set auxDllActiva = vABMEstadoDocumento
    
    vABMEstadoDocumento.Show
End Sub

Private Sub mnuABMFormaPago_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMFormaPago = New CListaBaseABM
    
    With vABMFormaPago
        .Caption = "Actualización de Estado Documento"
        .sql = "SELECT FPG_DESCRI, FPG_CODIGO FROM FORMA_PAGO"
        .HeaderSQL = "Descripción, Código"
        .FieldID = "FPG_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormFormaPago
        Set .FormDatos = ABMFormaPago
    End With
    
    Set auxDllActiva = vABMFormaPago
    
    vABMFormaPago.Show
End Sub

Private Sub mnuABMInscIVA_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMCondicionIva = New CListaBaseABM
    
    With vABMCondicionIva
        .Caption = "Actualización de Condición de I.V.A."
        .sql = "SELECT IVA_DESCRI, IVA_CODIGO FROM CONDICION_IVA"
        .HeaderSQL = "Descripción, Código"
        .FieldID = "IVA_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormCondicionIva
        Set .FormDatos = ABMCondicionIva
    End With
    
    Set auxDllActiva = vABMCondicionIva
    
    vABMCondicionIva.Show
End Sub

Private Sub mnuABMLineas_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMLineas = New CListaBaseABM
    
    With vABMLineas
        .Caption = "Actualización de Líneas"
        .sql = "SELECT LNA_DESCRI, LNA_CODIGO FROM LINEAS"
        .HeaderSQL = "Descripción, Código"
        .FieldID = "LNA_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormLineas
        Set .FormDatos = ABMLineas
    End With
    
    Set auxDllActiva = vABMLineas
    
    vABMLineas.Show
End Sub

Private Sub mnuABMLocalidades_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMLocalidad = New CListaBaseABM
    
    With vABMLocalidad
        .Caption = "Actualización de Localidades"
        .sql = "SELECT L.LOC_DESCRI, L.LOC_CODIGO, P.PRO_DESCRI, P.PRO_CODIGO, PA.PAI_DESCRI, P.PAI_CODIGO" & _
               " FROM LOCALIDAD L, PROVINCIA P, PAIS PA" & _
               " WHERE P.PAI_CODIGO=PA.PAI_CODIGO" & _
               " AND L.PAI_CODIGO=PA.PAI_CODIGO" & _
               " AND L.PRO_CODIGO=P.PRO_CODIGO"
        .HeaderSQL = "Descripción, Código, Provincia, Código ,País, Código"
        .FieldID = "LOC_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormLocalidad
        Set .FormDatos = ABMLocalidad
    End With
    
    Set auxDllActiva = vABMLocalidad
    
    vABMLocalidad.Show
End Sub

Private Sub mnuABMPais_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMPais = New CListaBaseABM
    
    With vABMPais
        .Caption = "Actualización de País"
        .sql = "SELECT PAI_DESCRI, PAI_CODIGO FROM PAIS"
        .HeaderSQL = "Descripción, Código"
        .FieldID = "PAI_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormPais
        Set .FormDatos = ABMPais
    End With
    
    Set auxDllActiva = vABMPais
    
    vABMPais.Show
End Sub

Private Sub mnuABMProveedores_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMProveedor = New CListaBaseABM
    
    With vABMProveedor
        .Caption = "Actualización de Proveedores"
        .sql = "SELECT P.PROV_RAZSOC, P.PROV_CODIGO, T.TPR_DESCRI, T.TPR_CODIGO" & _
               " FROM PROVEEDOR P, TIPO_PROVEEDOR T" & _
               " WHERE P.TPR_CODIGO=T.TPR_CODIGO"
        .HeaderSQL = "Razón Social, Código, Tipo de Proveedor, Código"
        .FieldID = "PROV_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormProveedor
        Set .FormDatos = ABMProveedor
    End With
    
    Set auxDllActiva = vABMProveedor
    
    vABMProveedor.Show
End Sub

Private Sub mnuABMProvincias_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMProvincia = New CListaBaseABM
    
    With vABMProvincia
        .Caption = "Actualización de Provincias"
        .sql = "SELECT P.PRO_DESCRI, P.PRO_CODIGO, PA.PAI_DESCRI, P.PAI_CODIGO" & _
               " FROM PROVINCIA P, PAIS PA" & _
               " WHERE P.PAI_CODIGO=PA.PAI_CODIGO"
        .HeaderSQL = "Descripción, Código, País, Código"
        .FieldID = "PRO_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormProvincia
        Set .FormDatos = ABMProvincia
    End With
    
    Set auxDllActiva = vABMProvincia
    
    vABMProvincia.Show
End Sub

Private Sub mnuABMRubros_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMRubros = New CListaBaseABM
    
    With vABMRubros
        .Caption = "Actualización de Rubros"
        .sql = "SELECT R.RUB_DESCRI, R.RUB_CODIGO, L.LNA_DESCRI, L.LNA_CODIGO" & _
               " FROM RUBROS R, LINEAS L" & _
               " WHERE R.LNA_CODIGO=L.LNA_CODIGO"
        .HeaderSQL = "Descripción, Código, Línea, Código"
        .FieldID = "RUB_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormRubros
        Set .FormDatos = ABMRubros
    End With
    
    Set auxDllActiva = vABMRubros
    
    vABMRubros.Show
End Sub

Private Sub mnuABMTipoGastos_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMTipoGastos = New CListaBaseABM
    
    With vABMTipoGastos
        .Caption = "Actualización de Tipo de Gastos"
        .sql = "SELECT TGT_DESCRI, TGT_CODIGO FROM TIPO_GASTO"
        .HeaderSQL = "Descripción, Código"
        .FieldID = "TGT_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormTipoGastos
        Set .FormDatos = ABMTipoGatos
    End With
    
    Set auxDllActiva = vABMTipoGastos
    
    vABMTipoGastos.Show
End Sub

Private Sub mnuABMTipoProveedores_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMTipoProveedor = New CListaBaseABM
    
    With vABMTipoProveedor
        .Caption = "Actualización de Tipo de Proveedor"
        .sql = "SELECT TPR_DESCRI, TPR_CODIGO FROM TIPO_PROVEEDOR"
        .HeaderSQL = "Descripción, Código"
        .FieldID = "TPR_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormTipoProveedor
        Set .FormDatos = ABMTipoProveedor
    End With
    
    Set auxDllActiva = vABMTipoProveedor
    
    vABMTipoProveedor.Show
End Sub

Private Sub mnuABMVendedores_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMVendedor = New CListaBaseABM
    
    With vABMVendedor
        .Caption = "Actualización de Vendedores"
        .sql = "SELECT VEN_NOMBRE, VEN_CODIGO, VEN_DOMICI, VEN_TELEFONO, VEN_MAIL" & _
               " FROM VENDEDOR"
        .HeaderSQL = "Nombre, Código, Domicilio, Teléfono, e-mail"
        .FieldID = "VEN_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormVendedor
        Set .FormDatos = ABMVendedor
    End With
    
    Set auxDllActiva = vABMVendedor
    
    vABMVendedor.Show
End Sub

Private Sub mnuAjuste_Click()
    frmAjusteStock.Show
End Sub

Private Sub mnuAjusteStock_Click()
    frmAjusteStock.Show
End Sub

Private Sub mnuAnulaRecibos_Click()
    frmAnulaDocumentos.TipodeAnulacion = 4
    frmAnulaDocumentos.Show
End Sub

Private Sub mnuAparatos_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMAparato = New CListaBaseABM
    
    With vABMAparato
        .Caption = "Actualización de Aparatos"
        .sql = "SELECT APT_DESCRI, APT_CODIGO FROM APARATO"
        .HeaderSQL = "Descripción, Código"
        .FieldID = "APT_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormAparato
        Set .FormDatos = ABMAparato
    End With
    
    Set auxDllActiva = vABMAparato
    
    vABMAparato.Show
End Sub

Private Sub mnuArcSal_Click()
    On Error Resume Next
    'verifico si la conexión esta abierta antes de salir
    'If Me.mnuConexion.Enabled = False Then
    DBConn.CloseConnection
    Set DBConn = Nothing
    'End If
    Set Menu = Nothing
    End
End Sub

Private Sub mnuBancos_Click()
'    ABMBanco.Show
End Sub

Private Sub mnubkpArchivos_Click()
    With frmRestaurarBD
        .Caption = "Backup de Archivos"
        .optCopiarA.Value = True
        .Label1 = "Guardar Backup en: "
        .Show
    End With
End Sub

Private Sub mnuCajaS_Click()
    frmPlanillaStock.Show
End Sub

Private Sub mnuCantVend_Click()
    frmListadoCantidadesVendidas.Show
End Sub

Private Sub mnuCascada_Click()
    Me.Arrange 0
End Sub

Private Sub mnuCheques_Click()
'    FrmCargaCheques.Show
End Sub

Private Sub mnuChequesTerceros_Click()
    FrmCargaCheques.Show
End Sub

Private Sub mnuCierreX_Click()
    frmCierreX.Show
End Sub

Private Sub mnuCierreZ_Click()
    frmCierreZ.Show
End Sub

Private Sub mnuComposturas_Click()
    'frmComposturas.Show
End Sub

Private Sub mnuComprasListado_Click()
    frmListadoProvedores.Show
End Sub

Private Sub mnuConAnuFactura_Click()
    frmAnulaDocumentos.TipodeAnulacion = 3
    frmAnulaDocumentos.Show
End Sub

Private Sub mnuConectar_Click()
    FrmInicio.Show vbModal
    Me.Caption = TituloPrincipal & " - (Usuario " & UCase(mNomUser) & " conectado a " & UCase(SERVIDOR) & ")"
    Me.mnuConectar.Enabled = False
End Sub

Private Sub mnuContenido_Click()
    Call WinHelp(Me.hWnd, App.Path & "\help\AYUDA.HLP", HelpFinder, 0&)
End Sub

Public Sub mnuContextABM_Click(index As Integer)

Dim auxListView As ListView
Dim auxModo As Integer
    
    auxModo = 0
    Select Case index
        Case 0 'nuevo
            auxModo = 1
        Case 1 'editar
            auxModo = 2
        Case 2 'eliminar
            auxModo = 4
        Case 9 ' ver datos
            auxModo = 3
        'Case 7 ' imprimir
        '   auxModo = 7
    End Select
    
    If auxModo > 0 Then
        Set auxListView = auxDllActiva.FormBase.lstvLista
        auxDllActiva.FormDatos.SetWindow auxDllActiva.FormBase, auxDllActiva.sql, auxModo, auxListView, auxDllActiva.FieldID
        auxDllActiva.FormDatos.Show vbModal
    Else
        'si es una acción de edición de datos
        Select Case index
            Case 4 'refresh
                Screen.MousePointer = vbHourglass
                With auxDllActiva
                    Set auxListView = .FormBase.lstvLista
                    CargarListView .FormBase, auxListView, .sql, .FieldID, .HeaderSQL, .FormBase.ImgLstLista
                    .FormBase.sBarEstado.Panels(1).Text = auxListView.ListItems.Count & " Registro(s)"
                End With
                Screen.MousePointer = vbDefault

            Case 5 'refresh
                'auxDllActiva.FormBase.txtBusqueda.Text = ""
                'auxDllActiva.FormBase.fraFiltro.Visible = True
                'auxDllActiva.FormBase.txtBusqueda.SetFocus
                With auxDllActiva
                    If .Caption = "Actualización de Productos" Then
                        frmFiltroProducto.Show
                    Else
                        frmFiltro.Show
                    End If
                        
                End With

            Case 6 'Buscar
                    auxDllActiva.Find
                
            Case 7 'imprimir
                Select Case mQuienLlamo
                    Case "ABMProducto"
                        frmImprimeProducto.Show vbModal
                    Case Else
                        On Error GoTo ErrorReport
                        auxDllActiva.FormBase.rptListado.Action = 1
                        On Error GoTo 0
                End Select
        End Select
    End If
    Exit Sub
    
ErrorReport:
    
    Beep
    MsgBox "Error " & Err.Number & Chr(13) & Err.Description, vbCritical + vbOKOnly, App.Title
    
End Sub

Private Sub mnuCtaCteClientes_Click()
    frmCtaCteCliente.Show
End Sub

Private Sub mnudepobanc_Click()
    frmDepositosBanco.Show
End Sub

Private Sub mnuDesconectar_Click()
    If DBConn.State = adStateOpen Then
        DBConn.Close
        
        DeshabilitarMenu Me
        
        Me.mnuArc.Enabled = True
        Me.mnuConectar.Enabled = True
        Me.mnuArcSal.Enabled = True
        Me.mnuDesconectar.Enabled = False
        
        Me.Caption = TituloPrincipal & " - (No conectado)"
    End If
End Sub

Private Sub mnuEntradaProductos_Click()
    frmEntradaProductos.Show vbModal
End Sub

Private Sub mnuDestinos_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMDestinos = New CListaBaseABM
    
    With vABMDestinos
        .Caption = "Actualización de Destinos"
        .sql = "SELECT DES_DESCRI, DES_CODIGO FROM DESTINOS"
        .HeaderSQL = "Descripción, Código"
        .FieldID = "DES_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormDestinos
        Set .FormDatos = ABMDestinos
    End With
    
    Set auxDllActiva = vABMDestinos
    
    vABMDestinos.Show
End Sub

Private Sub mnuentrada_Click()
    frmEntradaProductos.Show
End Sub

Private Sub mnuEntradaMercaderia_Click()
    frmEntradaProductos.Show
End Sub

Private Sub mnuEstaCantidadVendida_Click()
    frmListadoSurtCantidadesVendidas.Show
End Sub

Private Sub mnuEstadoProductos_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMEstadoProducto = New CListaBaseABM
    
    With vABMEstadoProducto
        .Caption = "Actualización de Estado de Producto"
        .sql = "SELECT ESP_DESCRI, ESP_CODIGO, ESP_SIGNO FROM ESTADO_PRODUCTO"
        .HeaderSQL = "Descripción, Código, Signo"
        .FieldID = "ESP_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormEstadoProducto
        Set .FormDatos = ABMEstadoProducto
    End With
    
    Set auxDllActiva = vABMEstadoProducto
    
    vABMEstadoProducto.Show
End Sub

Private Sub mnuFac_Click()
    frmListadoPlanillaVentas.Show
End Sub

Private Sub mnuFacturacionFacturacion_Click()
    frmFacturaCliente.Show
    'frmFacturaClienteMarcos.Show
End Sub

Private Sub mnuFacturacionPorRemito_Click()
    frmFacturaCliente.Show vbModal
End Sub

Private Sub mnuFacturacionTipoComprobante_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMTipoCompronate = New CListaBaseABM
    
    With vABMTipoCompronate
        .Caption = "Actualización de Tipo de Comprobantes"
        .sql = "SELECT TCO_DESCRI, TCO_CODIGO, TCO_ABREVIA FROM TIPO_COMPROBANTE"
        .HeaderSQL = "Descripción, Código, Abrevia"
        .FieldID = "TCO_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormTipoComprobante
        Set .FormDatos = ABMTipoComprobante
    End With
    
    Set auxDllActiva = vABMTipoCompronate
    
    vABMTipoCompronate.Show
End Sub

Private Sub mnuFondosCargaIngresos_Click()

End Sub

Private Sub mnuGastosGeneralesRegistro_Click()
    
End Sub

Private Sub mnuGastos_Click()
    frmCargaGastosGenerales.Show
End Sub

Private Sub mnuIconos_Click()
    Me.Arrange 3
End Sub

Private Sub mnuInformeSecEn_Click()
    FrmInformeSecEnergia.Show
End Sub

Private Sub mnuLibroCreditoFiscal_Click()
    frmLibroCompras2.Show
End Sub

Private Sub mnuLibroDebitoFiscal_Click()
    'frmLibroIvaVentas.Show
    
    frmLibroVentas2.Show
End Sub

Private Sub mnuLista_Click()
    FrmListadePrecios.Show
End Sub

Private Sub mnuListaCheques_Click()
    'FrmListCheques.Show
End Sub

Private Sub mnuListadoProductos_Click()
    frmListadoProductos.Show
End Sub

Private Sub mnuListadoProveedore_Click()
    frmListadoProvedores.Show
End Sub

Private Sub mnuListadoVentaPorVendedor_Click()
    frmListadoCantVendidasVendedor.Show
End Sub

Private Sub mnuListaPrecios_Click()
    FrmListadePrecios.Show
End Sub

Private Sub mnuMarcas_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMMarcas = New CListaBaseABM
    
    With vABMMarcas
        .Caption = "Actualización de Marcas"
        .sql = "SELECT MAR_DESCRI, MAR_CODIGO FROM MARCAS"
        .HeaderSQL = "Descripción, Código"
        .FieldID = "MAR_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormMarcas
        Set .FormDatos = ABMMarcas
    End With
    
    Set auxDllActiva = vABMMarcas
    
    vABMMarcas.Show
End Sub

Private Sub mnuMosHoriz_Click()
    Me.Arrange 1
End Sub

Private Sub mnuMosVert_Click()
    Me.Arrange 2
End Sub

Private Sub mnuNC_Click()
    'frmNCCliente.Show
End Sub

Private Sub mnuParametros_Click()
    frmParametros.Show vbModal
End Sub

Private Sub mnuPermisos_Click()
    FrmPermisos.Show vbModal
End Sub

Private Sub mnuProveedoresFacturas_Click()
   'frmFacturaProveedores.Show
End Sub

Private Sub mnuReciboClientes_Click()
    frmReciboCliente.Show
End Sub

Private Sub mnuRendTasaVial_Click()
    frmRendicionTasaVial.Show
End Sub

Private Sub mnurestArchivos_Click()
    With frmRestaurarBD
        .Caption = "Restaurar Archivos"
        .optCopiarA.Value = True
        .Label1 = "Restaurar desde: "
        .Show
    End With
End Sub

Private Sub mnuRevelados_Click()
    'frmRevelados.Show
End Sub

Private Sub mnuStockABMProductos_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMProductos = New CListaBaseABM
    
    With vABMProductos
        .Caption = "Actualización de Productos"
        .sql = "SELECT P.PTO_DESCRI,P.PTO_PRECTO,P.PTO_IVA, P.PTO_CODIGO, R.RUB_DESCRI, L.LNA_DESCRI, M.MAR_DESCRI" & _
               " FROM PRODUCTO P, RUBROS R, LINEAS L, MARCAS M" & _
               " WHERE R.LNA_CODIGO=L.LNA_CODIGO" & _
               " AND P.LNA_CODIGO=L.LNA_CODIGO" & _
               " AND P.RUB_CODIGO=R.RUB_CODIGO" & _
               " AND P.MAR_CODIGO=M.MAR_CODIGO"
        .HeaderSQL = "Descripción, Precio, Impuesto, Código, Rubro, Línea, Marca"
        .FieldID = "PTO_CODIGO"
        Set .FormBase = vFormProductos
        Set .FormDatos = ABMProducto
        .Report = "C:"
    End With
    
    Set auxDllActiva = vABMProductos
    mQuienLlamo = "ABMProducto"
    
    vABMProductos.Show
End Sub

Private Sub mnuStockAjuste_Click()
    frmControlStock.Show
End Sub

Private Sub mnuSucursal_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMSucursal = New CListaBaseABM
    
    With vABMSucursal
        .Caption = "Actualización de la Sucursal"
        .sql = "SELECT SUC_DESCRI, SUC_CODIGO FROM SUCURSAL"
        .HeaderSQL = "Descripción, Código"
        .FieldID = "SUC_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormSucursal
        Set .FormDatos = ABMSucursal
    End With
    
    Set auxDllActiva = vABMSucursal
    
    vABMSucursal.Show
End Sub

Private Sub mnuTarjetaCredito_Click()
     Dim cSQL As String
    
    mOrigen = True
        
    Set vABMTarjeta = New CListaBaseABM
    
    With vABMTarjeta
        .Caption = "Actualización de la Tarjeta"
        .sql = "SELECT TAR_DESCRI, TAR_CODIGO, TAR_TELEFONO FROM TARJETA"
        .HeaderSQL = "Descripción, Teléfono"
        .FieldID = "TAR_CODIGO"
        '.Report = RPTPATH & "tarjeta_credito.rpt"
        Set .FormBase = vFormTarjeta
        Set .FormDatos = ABMTarjeta
    End With
    
    Set auxDllActiva = vABMTarjeta
    
    vABMTarjeta.Show
End Sub

Private Sub mnuTarjetaPlan_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMTarjetaPlan = New CListaBaseABM
    
    With vABMTarjetaPlan
        .Caption = "Actualización de Planes de Tarjetas"
        .sql = "SELECT P.PLA_DESCRI, P.PLA_CODIGO, T.TAR_DESCRI, T.TAR_CODIGO" & _
               " FROM TARJETA T, TARJETA_PLAN P" & _
               " WHERE P.TAR_CODIGO=T.TAR_CODIGO"
        .HeaderSQL = "Descripción, Código, Tarjeta, Código"
        .FieldID = "PLA_CODIGO"
        '.Report = RPTPATH & "tarjeta_plan.rpt"
        Set .FormBase = vFormTarjetaPlan
        Set .FormDatos = ABMTarjetaPlan
    End With
    
    Set auxDllActiva = vABMTarjetaPlan
    
    vABMTarjetaPlan.Show
End Sub

Private Sub mnuTipoRevelado_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMTipoRevelado = New CListaBaseABM
    
    With vABMTipoRevelado
        .Caption = "Actualización de Tipo Revelado"
        .sql = "SELECT TRE_DESCRI, TRE_CODIGO FROM TIPO_REVELADO"
        .HeaderSQL = "Descripción, Código"
        .FieldID = "TRE_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormTipoRevelado
        Set .FormDatos = ABMTipoRevelado
    End With
    
    Set auxDllActiva = vABMTipoRevelado
    
    vABMTipoRevelado.Show
End Sub

Private Sub mnuUsuarios_Click()
    FrmUsuarios.Show vbModal
End Sub

Private Sub Tarjetas_Click()
   frmTarjetasCredito.Show
End Sub

Private Sub tbrPrincipal_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.index
        Case 2: Call frmPlanillaStock.Show
        Case 2: Call mnuFacturacionFacturacion_Click
        Case 3: Call mnuLista_Click
        Case 4: Call mnuABMClientes_Click
        Case 6: Call mnuReciboClientes_Click
        Case 7: Call mnuFac_Click
        Case 9: Call mnuCascada_Click
    End Select
End Sub
Private Function configuroLetrero(Fecha As Date) As String
    Dim DIA As Integer
    Dim Doc As Integer
    'Dim nUltimaLista As Integer
    Dim nRazonSocial As String
    
    DIA = Weekday(Fecha, vbMonday)
    'Doc = 0
    'If User <> 99 Then
    '    Doc = XN(User)
    'End If
    
    'busco en parametros el nombre de la empresa
    sql = "SELECT RAZ_SOCIAL FROM PARAMETROS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        nRazonSocial = rec!RAZ_SOCIAL
    End If
    rec.Close
    
    'busco ultima lista de precios
'    sql = "SELECT MAX(LIS_CODIGO) AS ULTIMALISTA FROM LISTA_PRECIO"
'    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    If rec.EOF = False Then
'        nUltimaLista = rec!ULTIMALISTA
'    End If
'    rec.Close
    
    sql = "SELECT P.PTO_CODIGO, P.PTO_DESCRI, P.PTO_PRECTO "
    sql = sql & " FROM PRODUCTO P" ', DETALLE_LISTA_PRECIO D"
    sql = sql & " WHERE " 'P.PTO_CODIGO = D.PTO_CODIGO"
    sql = sql & " P.LIS_CODIGO = 1" ' & nUltimaLista
        sql = sql & " ORDER BY P.PTO_CODIGO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        
        Letrero = nRazonSocial & " // Precios del dia: " & WeekdayName(DIA, False) & " " & Day(Fecha) & " de " & MonthName(Month(Fecha), False) & " de " & Year(Fecha) & " : "
        Do While rec!PTO_CODIGO < 5
            'If rec!PTO_CODIGO < 4 Then  'MUESTRO NAFTA, GNC Y GASOIL EN EL LETRERO
             Letrero = Letrero & rec!PTO_DESCRI & "  $ " & Format(rec!PTO_PRECTO, "0.000") & " / "
            
            'End If
            'Letrero = Letrero & Format(rec!TUR_HORAD, "hh:mm") & " Hs - " & _
                          rec!CLI_RAZSOC & " , " & rec!TUR_MOTIVO & " // "
            
            rec.MoveNext
        Loop
    End If
    rec.Close
    Letrero = Letrero
        
End Function
Private Sub Timer1_Timer()
    
    Static Anterior As Boolean
    Static tamañoLetrero As Single
    Static X As Single
    If Not Anterior Then
        tamañoLetrero = Menu.Picture2.TextWidth(Letrero)
        Anterior = True
        X = Menu.Picture2.ScaleWidth
    End If
    Menu.Picture2.Cls
    Menu.Picture2.CurrentX = X
    Menu.Picture2.CurrentY = 100
'Para cambiar el tipo de letra
    Menu.Picture2.FontName = "Arial"
    Menu.Picture2.FontBold = True
    Menu.Picture2.Print Letrero
    X = X - 25
    If X < -tamañoLetrero Then X = Menu.Picture2.ScaleWidth
End Sub

