VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTarjetasCredito 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Tarjetas de Credito"
   ClientHeight    =   8112
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   10680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.4
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
   Picture         =   "frmTarjetasCredito.frx":0000
   ScaleHeight     =   8112
   ScaleWidth      =   10680
   Begin VB.CommandButton cmdConcilia 
      Caption         =   "&Conciliar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6150
      TabIndex        =   6
      Top             =   7583
      Width           =   1110
   End
   Begin VB.CommandButton cmdReporte 
      Caption         =   "&Reporte"
      Height          =   450
      Left            =   5040
      Picture         =   "frmTarjetasCredito.frx":0D82
      TabIndex        =   68
      Top             =   7575
      Width           =   1095
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   450
      Left            =   7290
      Picture         =   "frmTarjetasCredito.frx":108C
      TabIndex        =   2
      Top             =   7575
      Width           =   1095
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   9525
      Picture         =   "frmTarjetasCredito.frx":1396
      TabIndex        =   4
      Top             =   7575
      Width           =   1095
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&An&ular"
      Height          =   450
      Left            =   8400
      Picture         =   "frmTarjetasCredito.frx":16A0
      TabIndex        =   3
      Top             =   7575
      Width           =   1095
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   7500
      Left            =   15
      TabIndex        =   15
      Top             =   30
      Width           =   10620
      _ExtentX        =   18733
      _ExtentY        =   13229
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   529
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Datos"
      TabPicture(0)   =   "frmTarjetasCredito.frx":19AA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameGeneral"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FrameProducto"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtObservaciones"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraconcilia"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraTarjeta"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "frmTarjetasCredito.frx":19C6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GRDGrilla"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(2)=   "frameVer"
      Tab(1).ControlCount=   3
      Begin VB.Frame fraTarjeta 
         Height          =   3405
         Left            =   0
         TabIndex        =   71
         Top             =   2760
         Visible         =   0   'False
         Width           =   4935
         Begin VB.TextBox txttarjeta_importe 
            Height          =   315
            Left            =   1665
            TabIndex        =   74
            Top             =   1350
            Width           =   2505
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
            TabIndex        =   73
            ToolTipText     =   "Alta de Tarjeta"
            Top             =   990
            Width           =   480
         End
         Begin VB.CommandButton cmdCerrarTarjeta 
            Caption         =   "Cerrar"
            Height          =   375
            Left            =   3690
            TabIndex        =   81
            Top             =   2880
            Width           =   1095
         End
         Begin VB.TextBox txtTar_Autorizacion 
            Height          =   315
            Left            =   1665
            MaxLength       =   30
            TabIndex        =   77
            Top             =   2445
            Width           =   2505
         End
         Begin VB.ComboBox cbotarjeta_tarjeta 
            Height          =   315
            ItemData        =   "frmTarjetasCredito.frx":19E2
            Left            =   1665
            List            =   "frmTarjetasCredito.frx":19E4
            Style           =   2  'Dropdown List
            TabIndex        =   72
            Top             =   975
            Width           =   2505
         End
         Begin VB.TextBox txtCupon 
            Height          =   315
            Left            =   1665
            TabIndex        =   76
            Top             =   2085
            Width           =   2505
         End
         Begin VB.TextBox txtLote 
            Height          =   315
            Left            =   1665
            TabIndex        =   75
            Top             =   1725
            Width           =   2505
         End
         Begin VB.CommandButton cmdAceptoTarjeta 
            Caption         =   "Aceptar"
            Height          =   375
            Left            =   2220
            TabIndex        =   78
            Top             =   2880
            Width           =   1425
         End
         Begin MSComCtl2.DTPicker fechaTar 
            Height          =   315
            Left            =   1650
            TabIndex        =   70
            Top             =   600
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   550
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   106299393
            CurrentDate     =   41098
         End
         Begin VB.Label Label27 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fecha Tarjeta:"
            Height          =   315
            Left            =   405
            TabIndex        =   86
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label26 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Autorización:"
            Height          =   315
            Left            =   405
            TabIndex        =   85
            Top             =   2445
            Width           =   1215
         End
         Begin VB.Label Label25 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tarjeta:"
            Height          =   315
            Left            =   405
            TabIndex        =   84
            Top             =   975
            Width           =   1215
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Monto:"
            Height          =   315
            Left            =   405
            TabIndex        =   83
            Top             =   1365
            Width           =   1215
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cupón:"
            Height          =   315
            Left            =   405
            TabIndex        =   82
            Top             =   2085
            Width           =   1215
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Lote:"
            Height          =   315
            Left            =   405
            TabIndex        =   80
            Top             =   1725
            Width           =   1215
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackColor       =   &H00FF8080&
            Caption         =   "Datos Tarjeta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   30
            TabIndex        =   79
            Top             =   120
            Width           =   4845
         End
      End
      Begin VB.Frame fraconcilia 
         Caption         =   "Conciliacion"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   90
         TabIndex        =   43
         Top             =   315
         Visible         =   0   'False
         Width           =   10455
         Begin VB.TextBox txtImpuestos 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   3750
            MaxLength       =   40
            TabIndex        =   50
            Text            =   "0,00"
            Top             =   1147
            Width           =   1500
         End
         Begin VB.TextBox txtperIIBB 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   7140
            MaxLength       =   40
            TabIndex        =   53
            Text            =   "0,00"
            Top             =   1147
            Width           =   1500
         End
         Begin VB.TextBox txtdeduccionimp 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   7140
            MaxLength       =   40
            TabIndex        =   52
            Text            =   "0,00"
            Top             =   675
            Width           =   1500
         End
         Begin VB.TextBox txttotConc 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   8820
            MaxLength       =   40
            TabIndex        =   55
            Text            =   "0,00"
            Top             =   1567
            Width           =   1500
         End
         Begin VB.CommandButton cmdCoSalir 
            Caption         =   "&Salir"
            Height          =   450
            Left            =   9015
            Picture         =   "frmTarjetasCredito.frx":19E6
            TabIndex        =   64
            Top             =   720
            Width           =   1095
         End
         Begin VB.CommandButton cmdCoAceptar 
            Caption         =   "&Aceptar"
            Height          =   450
            Left            =   9000
            Picture         =   "frmTarjetasCredito.frx":1CF0
            TabIndex        =   62
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtNeto 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   750
            MaxLength       =   40
            TabIndex        =   45
            Text            =   "0,00"
            Top             =   1147
            Width           =   1500
         End
         Begin VB.TextBox txtSubTotal 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3750
            MaxLength       =   40
            TabIndex        =   49
            Text            =   "0,00"
            Top             =   675
            Width           =   1500
         End
         Begin VB.TextBox txtimp1IVA 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1470
            MaxLength       =   40
            TabIndex        =   48
            Text            =   "0,00"
            Top             =   1560
            Width           =   780
         End
         Begin VB.TextBox txtperIVA 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   3750
            MaxLength       =   40
            TabIndex        =   51
            Text            =   "0,00"
            Top             =   1567
            Width           =   1500
         End
         Begin VB.TextBox txtperGAN 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   7140
            MaxLength       =   40
            TabIndex        =   54
            Text            =   "0,00"
            Top             =   1567
            Width           =   1500
         End
         Begin MSComCtl2.DTPicker FechaComprobante 
            Height          =   315
            Left            =   750
            TabIndex        =   44
            Top             =   285
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   550
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   106299393
            CurrentDate     =   41098
         End
         Begin VB.TextBox txtIva 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   750
            MaxLength       =   40
            TabIndex        =   46
            Text            =   "0,00"
            Top             =   1560
            Width           =   660
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Exento:"
            Height          =   195
            Left            =   3135
            TabIndex        =   89
            Top             =   1200
            Width           =   540
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Percepcion IIBB:"
            Height          =   195
            Left            =   5925
            TabIndex        =   88
            Top             =   1200
            Width           =   1200
         End
         Begin VB.Label lbldeduccionimp 
            AutoSize        =   -1  'True
            Caption         =   "Deduccion Impositiva:"
            Height          =   195
            Left            =   5520
            TabIndex        =   87
            Top             =   735
            Width           =   1575
         End
         Begin VB.Label Label18 
            Caption         =   "Ventas:"
            Height          =   195
            Left            =   120
            TabIndex        =   66
            Top             =   735
            Width           =   630
         End
         Begin VB.Label lblVentas 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "lblventas"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   210
            Left            =   720
            TabIndex        =   65
            Top             =   720
            Width           =   1425
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Total Conciliado:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8880
            TabIndex        =   63
            Top             =   1320
            Width           =   1380
         End
         Begin VB.Label lbltarjeta 
            AutoSize        =   -1  'True
            Caption         =   "lbltarjeta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   270
            Left            =   2520
            TabIndex        =   61
            Top             =   240
            Width           =   1050
         End
         Begin VB.Label Label9 
            Caption         =   "Arancel:"
            Height          =   195
            Left            =   45
            TabIndex        =   60
            Top             =   1200
            Width           =   630
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   180
            TabIndex        =   59
            Top             =   315
            Width           =   495
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Percepcion Ganancias:"
            Height          =   195
            Left            =   5460
            TabIndex        =   58
            Top             =   1620
            Width           =   1665
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Sub Total:"
            Height          =   195
            Left            =   2940
            TabIndex        =   57
            Top             =   735
            Width           =   735
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Iva:"
            Height          =   195
            Left            =   405
            TabIndex        =   56
            Top             =   1590
            Width           =   270
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Percepcion IVA:"
            Height          =   195
            Left            =   2520
            TabIndex        =   47
            Top             =   1620
            Width           =   1155
         End
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   465
         Left            =   1275
         MaxLength       =   199
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   6945
         Width           =   9210
      End
      Begin VB.Frame frameVer 
         Caption         =   "Ver..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   -74910
         TabIndex        =   25
         Top             =   6720
         Visible         =   0   'False
         Width           =   10170
         Begin VB.OptionButton optSeleccion 
            Alignment       =   1  'Right Justify
            Caption         =   "... Listar Seleccionado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5280
            TabIndex        =   27
            Top             =   210
            Width           =   1935
         End
         Begin VB.OptionButton optTodos 
            Alignment       =   1  'Right Justify
            Caption         =   "... Listar Todos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1755
            TabIndex        =   26
            Top             =   210
            Value           =   -1  'True
            Width           =   1380
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Buscar Conciliaciones por..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1440
         Left            =   -74880
         TabIndex        =   16
         Top             =   345
         Width           =   10125
         Begin VB.ComboBox cbotarjeta_b 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   480
            Width           =   5085
         End
         Begin VB.TextBox txtOrden 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10320
            TabIndex        =   28
            Text            =   "A"
            Top             =   1080
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.ComboBox cbotipo_b 
            Height          =   315
            Left            =   7125
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   360
            Visible         =   0   'False
            Width           =   2325
         End
         Begin VB.CommandButton CmdBuscAprox 
            Caption         =   "Buscar Conciliaciones"
            Height          =   1020
            Left            =   8040
            MaskColor       =   &H8000000F&
            TabIndex        =   13
            ToolTipText     =   "Buscar Nota de Pedido"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker FechaDesde 
            Height          =   315
            Left            =   1965
            TabIndex        =   11
            Top             =   960
            Width           =   1455
            _ExtentX        =   2561
            _ExtentY        =   550
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   106299393
            CurrentDate     =   41098
         End
         Begin MSComCtl2.DTPicker FechaHasta 
            Height          =   315
            Left            =   5520
            TabIndex        =   12
            Top             =   960
            Width           =   1455
            _ExtentX        =   2561
            _ExtentY        =   550
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   106299393
            CurrentDate     =   41098
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Tarjeta:"
            Height          =   195
            Left            =   1200
            TabIndex        =   67
            Top             =   540
            Width           =   585
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   6660
            TabIndex        =   23
            Top             =   420
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Label lblFechaDesde 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   900
            TabIndex        =   18
            Top             =   945
            Width           =   990
         End
         Begin VB.Label lblFechaHasta 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   4365
            TabIndex        =   17
            Top             =   960
            Width           =   960
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3870
         Left            =   -74655
         TabIndex        =   19
         Top             =   2340
         Width           =   10455
         _ExtentX        =   18436
         _ExtentY        =   6837
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorSel    =   8388736
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid GRDGrilla 
         Height          =   5505
         Left            =   -74880
         TabIndex        =   14
         Top             =   1830
         Width           =   10185
         _ExtentX        =   17971
         _ExtentY        =   9694
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         RowHeightMin    =   280
         BackColorSel    =   16761024
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame FrameProducto 
         Caption         =   "Tarjetas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5085
         Left            =   90
         TabIndex        =   21
         Top             =   1845
         Width           =   10425
         Begin VB.TextBox txtTotalSel 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7440
            TabIndex        =   41
            Top             =   4680
            Width           =   1230
         End
         Begin VB.CommandButton CmdDeselec 
            Caption         =   "&Quitar Todos"
            Height          =   555
            Left            =   9360
            TabIndex        =   8
            Top             =   1320
            Width           =   990
         End
         Begin VB.CommandButton CmdSelec 
            Caption         =   "&Seleccionar Todos"
            Height          =   675
            Left            =   9360
            TabIndex        =   7
            Top             =   600
            Width           =   990
         End
         Begin VB.TextBox txttotal 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3240
            TabIndex        =   39
            Top             =   4680
            Width           =   1230
         End
         Begin MSFlexGridLib.MSFlexGrid GrdModulos 
            Height          =   4290
            Left            =   120
            TabIndex        =   5
            Top             =   345
            Width           =   9150
            _ExtentX        =   16150
            _ExtentY        =   7557
            _Version        =   393216
            Cols            =   5
            FixedCols       =   0
            RowHeightMin    =   280
            BackColorSel    =   16761024
            FocusRect       =   0
            SelectionMode   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.CommandButton cmdAgregar 
            Caption         =   "&Agregar Cupon"
            Height          =   555
            Left            =   9360
            TabIndex        =   69
            Top             =   2520
            Width           =   990
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "TOTAL SELECCIONADO: $"
            Height          =   195
            Left            =   5520
            TabIndex        =   42
            Top             =   4740
            Width           =   1890
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "TOTAL: $"
            Height          =   195
            Left            =   2520
            TabIndex        =   40
            Top             =   4740
            Width           =   675
         End
      End
      Begin VB.Frame FrameGeneral 
         Caption         =   "Filtros"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   90
         TabIndex        =   24
         Top             =   360
         Width           =   10425
         Begin VB.ComboBox cbotarjeta 
            Height          =   315
            Left            =   1275
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   240
            Width           =   4620
         End
         Begin VB.ComboBox cboTipo 
            Height          =   315
            Left            =   6435
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   600
            Visible         =   0   'False
            Width           =   3300
         End
         Begin VB.CommandButton cmdBuscarTarjetas 
            Caption         =   "Buscar Tarjetas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   930
            Left            =   8400
            MaskColor       =   &H000000FF&
            Picture         =   "frmTarjetasCredito.frx":1FFA
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Buscar Tarjetas"
            Top             =   360
            UseMaskColor    =   -1  'True
            Width           =   1605
         End
         Begin MSComCtl2.DTPicker Fecha 
            Height          =   315
            Left            =   13920
            TabIndex        =   0
            Top             =   360
            Width           =   1455
            _ExtentX        =   2561
            _ExtentY        =   550
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   106299393
            CurrentDate     =   41098
         End
         Begin MSComCtl2.DTPicker fdesdeT 
            Height          =   315
            Left            =   1275
            TabIndex        =   30
            Top             =   735
            Width           =   1455
            _ExtentX        =   2582
            _ExtentY        =   572
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   106299393
            CurrentDate     =   41098
         End
         Begin MSComCtl2.DTPicker fhastaT 
            Height          =   315
            Left            =   4500
            TabIndex        =   31
            Top             =   735
            Width           =   1455
            _ExtentX        =   2561
            _ExtentY        =   572
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   106299393
            CurrentDate     =   41098
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   6000
            TabIndex        =   37
            Top             =   660
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Tarjeta:"
            Height          =   195
            Left            =   600
            TabIndex        =   36
            Top             =   300
            Width           =   585
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   3465
            TabIndex        =   33
            Top             =   735
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   240
            TabIndex        =   32
            Top             =   720
            Width           =   990
         End
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones:"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   6960
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ordenado por :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   -74820
         TabIndex        =   20
         Top             =   570
         Width           =   1065
      End
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   0
      Top             =   7680
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowControls  =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lblestado 
      AutoSize        =   -1  'True
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   105
      TabIndex        =   22
      Top             =   6105
      Width           =   660
   End
End
Attribute VB_Name = "frmTarjetasCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim VnumeroListado As Long

Private Sub cboMovimiento_Click()
'    If cboMovimiento.ListIndex <> -1 Then
'        sql = "SELECT ESP_SIGNO "
'        sql = sql & " FROM ESTADO_PRODUCTO"
'        sql = sql & " WHERE ESP_CODIGO=" & cboMovimiento.ItemData(cboMovimiento.ListIndex)
'        Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
'        If Rec2.EOF = False Then
'            txtSigno.Text = ChkNull(Rec2!ESP_SIGNO)
'        End If
'        Rec2.Close
'    End If
End Sub

Private Sub cbotipo_b_LostFocus()
    'cbotarjeta_b.Clear
    cargocboTarjeta_b
End Sub

Private Sub cboTipo_LostFocus()
    cbotarjeta.Clear
    cargocboTarjeta
End Sub

Private Sub cmdAsignar_Click()
'    If TxtCODIGO.Text <> "" Then
'        GrdModulos.HighLight = flexHighlightAlways
'        If txtCantidad <> "" Then
'
'            If TxtCODIGO.Text = 1 Or TxtCODIGO.Text = 3 Then 'BUSCO STOCKS DE TANQUES
'                 For i = 1 To GrdModulos.Rows - 1
'                    If GrdModulos.TextMatrix(i, 4) = CLng(TxtCodInt.Text) Then
'                        If GrdModulos.TextMatrix(i, 4) = Right(Trim(IIf(optTanque1.Value = True, optTanque1.Caption, optTanque2.Caption)), 1) Then
'                            MsgBox "El combustible para ese tanque ya fue ingresado", vbExclamation, TIT_MSGBOX
'                            TxtCODIGO.SetFocus
'                            Exit Sub
'                        End If
'                    End If
'                Next
'
'
'                 GrdModulos.AddItem Trim(TxtCODIGO.Text) & Chr(9) & Trim(TxtDescri.Text) & " - " & IIf(optTanque1.Value = True, optTanque1.Caption, optTanque2.Caption) _
'                                & Chr(9) & Trim(txtCantidad.Text) & Chr(9) & "" & Chr(9) & Trim(TxtCodInt.Text) & Chr(9) & Right(Trim(IIf(optTanque1.Value = True, optTanque1.Caption, optTanque2.Caption)), 1)
'                'txtIngNuevo_Click
'                TxtCODIGO.Text = ""
'                TxtCODIGO.SetFocus
'                fraTanque.Visible = False
'            Else
'                If txtNumero.Text = "" Then
'                    For i = 1 To GrdModulos.Rows - 1
'                        If GrdModulos.TextMatrix(i, 0) = CLng(TxtCODIGO.Text) Then
'                            GrdModulos.TextMatrix(i, 2) = CDbl(GrdModulos.TextMatrix(i, 2)) + CDbl(txtCantidad.Text)
'                            TxtCODIGO.Text = ""
'                            TxtCODIGO.SetFocus
'                            Exit Sub
'                        End If
'                    Next
'                Else
'                    For i = 1 To GrdModulos.Rows - 1
'                        If GrdModulos.TextMatrix(i, 4) = CLng(TxtCodInt.Text) Then
'                            MsgBox "El producto ya fue ingresado", vbExclamation, TIT_MSGBOX
'                            TxtCODIGO.SetFocus
'                            Exit Sub
'                        End If
'                    Next
'                End If
'                GrdModulos.AddItem Trim(TxtCODIGO.Text) & Chr(9) & Trim(TxtDescri.Text) _
'                                & Chr(9) & Trim(txtCantidad.Text) & Chr(9) & "" & Chr(9) & Trim(TxtCodInt.Text)
'
'                'txtIngNuevo_Click
'                TxtCODIGO.Text = ""
'                TxtCODIGO.SetFocus
'            End If
'        Else
'            MsgBox "Debe Ingresar la cantidad", vbExclamation, TIT_MSGBOX
'            txtCantidad.SetFocus
'            Exit Sub
'        End If
'     Else
'        MsgBox "Debe seleccionar un Producto"
'    End If
End Sub

Private Sub cmdAceptoTarjeta_Click()
    If fechaTar.Value = "" Then
        MsgBox "Falta Ingresar la fecha del Cupon", vbExclamation, TIT_MSGBOX
        fechaTar.SetFocus
        Exit Sub
    End If
    'If cboPlan.ListIndex = -1 Then
    '    MsgBox "Falta Ingresar el Plan", vbExclamation, TIT_MSGBOX
    '    cboPlan.SetFocus
    '    Exit Sub
    'End If
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
        
    'agregar a grilla  CONTINUAR ACA, AGREGAR LA FECHA!!!!!
    
 '   GrdModulos.AddItem
        
        
    GrdModulos.AddItem Format(fechaTar.Value, "dd/mm/yyyy") & Chr(9) & cbotarjeta_tarjeta.Text & Chr(9) & _
                    VALIDO_IMPORTE(txttarjeta_importe) & Chr(9) & "" & Chr(9) & _
                    txtCupon.Text & Chr(9) & txtLote.Text & Chr(9) & _
                    txtTar_Autorizacion & Chr(9) & "NO"
                    'rec!TCO_CODIGO & Chr(9) & rec!FCL_NUMERO & Chr(9) & _
                    'rec!FCL_SUCURSAL & Chr(9) & rec!FPG_CODIGO & Chr(9) & _
                    'rec!PAG_SECUENCIA
                    'TOTAL = TOTAL + rec!PAG_IMPORTE
        
    fraTarjeta.Visible = False
    'cboFormaPago.ListIndex = 0
    cmdAgregar.SetFocus
    'txtImportePago.SetFocus
    calculototal
End Sub

Private Sub CmdAgregar_Click()
    fraTarjeta.Top = 1485
    fraTarjeta.Left = 3330
    fraTarjeta.Visible = True
    fechaTar.Enabled = True
    fechaTar.SetFocus
    cbotarjeta_tarjeta.Enabled = True
    txtLote.Enabled = True
    txtCupon.Enabled = True
    txtTar_Autorizacion.Enabled = True
    fraTarjeta.Visible = True
    
    
End Sub

Private Sub CmdBorrar_Click()
'    If txtNumero.Text <> "" Then
'        If GrdModulos.Rows <> 1 Then
'            If MsgBox("¿Seguro desea Anular el Movimineto de Producto Nro: " & XN(txtNumero.Text) & "? ", vbQuestion + vbYesNo + vbDefaultButton2, TIT_MSGBOX) = vbYes Then
'                lblestado.Caption = "Anulando..."
'                Screen.MousePointer = vbHourglass
'                On Error GoTo HayError1
'                DBConn.BeginTrans
'
'                'ANULO LA ENTRADA
'                sql = "UPDATE ENTRADA_PRODUCTO"
'                sql = sql & " SET EST_CODIGO=2"
'                sql = sql & " WHERE EPR_CODIGO=" & XN(txtNumero.Text)
'                DBConn.Execute sql
'
'                'ACTUALIZO EL DETALLE
'                For i = 1 To GrdModulos.Rows - 1
'                    sql = "UPDATE STOCK"
'                    sql = sql & " SET DST_STKFIS = DST_STKFIS "
'                    If Trim(txtSigno.Text) = "+" Then
'                        sql = sql & " - "
'                    Else
'                        sql = sql & " + "
'                    End If
'                    sql = sql & XN(GrdModulos.TextMatrix(i, 2))
'                    sql = sql & " WHERE STK_CODIGO = " & XN(cboStock.ItemData(cboStock.ListIndex))
'                    sql = sql & " AND PTO_CODIGO = " & XN(GrdModulos.TextMatrix(i, 4))
'                    DBConn.Execute sql
'                Next
'                DBConn.CommitTrans
'            End If
'            lblestado.Caption = ""
'            Screen.MousePointer = vbNormal
'            CmdNuevo_Click
'        End If
'    End If
'  Exit Sub
'HayError1:
'    lblestado.Caption = ""
'    Screen.MousePointer = vbNormal
'    DBConn.RollbackTrans
'    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub CmdBuscAprox_Click()
    BuscarConciliaciones
End Sub

Private Sub cmdBuscarTarjetas_Click()
    
    BuscarTarjetas
End Sub

Private Sub cmdCerrarTarjeta_Click()
    fraTarjeta.Visible = False
    
End Sub

Private Sub cmdCoAceptar_Click()
    If ValidarConciliacion = False Then Exit Sub
    If MsgBox("¿Confirma Conciliacion?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    On Error GoTo HayErrorCarga
    DBConn.BeginTrans
    'grabar conciliacion y conciliacion detalle
    grabarConciliacion
    DBConn.CommitTrans
    limpiarConciliacion
    BuscarTarjetas
    fraconcilia.Visible = False
    cmdConcilia.Enabled = True
    Exit Sub
    
HayErrorCarga:
    'lblestado.Caption = ""
    Screen.MousePointer = vbNormal
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub
Private Sub grabarConciliacion()
    Dim Numero As Integer
    'BUSCO ULTIMA CONCILIACION
    sql = "SELECT MAX(CON_NUMERO) AS NRO FROM CONCILIACION"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Numero = Chk0(rec!NRO) + 1
    End If
    rec.Close
    
    sql = "INSERT INTO CONCILIACION"
    sql = sql & "(CON_NUMERO,CON_TARJETA,CON_FECHA,CON_VENTAS,CON_NETO,CON_IVA,CON_IMP1IVA,CON_IMPUESTOS,"
    sql = sql & "CON_PERIIBB,CON_PERIVA,CON_PERGAN,CON_TOTAL,CON_DEDUIMP,EST_CODIGO)"
    sql = sql & " VALUES ("
    sql = sql & Numero & ","
    sql = sql & XS(lbltarjeta.Caption) & ","
    sql = sql & XDQ(FechaComprobante.Value) & ","
    sql = sql & XN(lblVentas) & ","
    sql = sql & XN(txtNeto) & ","
    sql = sql & XN(txtIva) & ","
    sql = sql & XN(txtimp1IVA) & ","
    sql = sql & XN(txtImpuestos) & "," 'exento
    sql = sql & XN(txtperIIBB) & ","
    sql = sql & XN(txtperIVA) & ","
    sql = sql & XN(txtperGAN) & ","
    sql = sql & XN(txttotConc) & ","
    sql = sql & XN(txtdeduccionimp) & ","
    sql = sql & 1 & ")"
    DBConn.Execute sql
    
    For i = 1 To GrdModulos.Rows - 1
        If GrdModulos.TextMatrix(i, 7) <> "NO" Then
            sql = "INSERT INTO DETALLE_CONCILIACION"
            sql = sql & "(CON_NUMERO,TAR_CODIGO,DCO_NROITEM,DCO_FECHA,DCO_MONTO,"
            sql = sql & "DCO_PLAN,DCO_CUPON,DCO_LOTE,DCO_AUTO)"
            sql = sql & " VALUES ("
            sql = sql & Numero & ","
            sql = sql & XS(GrdModulos.TextMatrix(i, 1)) & ","
            sql = sql & i & ","
            sql = sql & XDQ(GrdModulos.TextMatrix(i, 0)) & ","
            sql = sql & XN(GrdModulos.TextMatrix(i, 2)) & ","
            sql = sql & XS(GrdModulos.TextMatrix(i, 3)) & ","
            sql = sql & XS(GrdModulos.TextMatrix(i, 4)) & ","
            sql = sql & XS(GrdModulos.TextMatrix(i, 5)) & "," 'exento
            sql = sql & XS(GrdModulos.TextMatrix(i, 6)) & ")"
            DBConn.Execute sql
            
            
            'actualizar en factura pagos para indicar la conciliacion
            sql = "UPDATE FACTURA_PAGOS"
            sql = sql & " SET CON_NUMERO= " & Numero
            sql = sql & " WHERE TCO_CODIGO=" & XN(GrdModulos.TextMatrix(i, 8))
            sql = sql & " AND FCL_NUMERO=" & XN(GrdModulos.TextMatrix(i, 9))
            sql = sql & " AND FCL_SUCURSAL=" & XN(GrdModulos.TextMatrix(i, 10))
            sql = sql & " AND FPG_CODIGO=" & XN(GrdModulos.TextMatrix(i, 11))
            sql = sql & " AND PAG_SECUENCIA=" & XN(GrdModulos.TextMatrix(i, 12))
            DBConn.Execute sql
        End If
    Next

    

End Sub
Private Function SumaTotal()
        If txtSubTotal.Enabled = True Then
            txttotConc.Text = CDbl(lblVentas) - (CDbl(Chk0(txtImpuestos.Text)) + CDbl(txtSubTotal.Text) + Chk0(txtperIIBB.Text) + Chk0(txtperIVA.Text) + Chk0(txtperGAN.Text))
        Else ' tarjeta visa o visa-debito
            txttotConc.Text = CDbl(lblVentas) - (CDbl(Chk0(txtNeto.Text)) + Chk0(txtdeduccionimp.Text))
        End If
        txttotConc.Text = Valido_Importe2(txttotConc)
End Function

Private Function ValidarConciliacion() As Boolean

    If IsNull(FechaComprobante.Value) Then
        MsgBox "La Fecha de la conciliacion es requerida", vbExclamation, TIT_MSGBOX
        FechaComprobante.SetFocus
        ValidarConciliacion = False
        Exit Function
    End If
    If txtNeto.Text = "" Then
        MsgBox "El Arancel es requerido", vbExclamation, TIT_MSGBOX
        txtNeto.SetFocus
        ValidarConciliacion = False
        Exit Function
    End If
    If txtIva.Text = "" Then
        MsgBox "El Procentaje del I.V.A. es requerido", vbExclamation, TIT_MSGBOX
        txtIva.SetFocus
        ValidarConciliacion = False
        Exit Function
    End If
    If txttotal.Text = "" Then
        MsgBox "El Total es requerido", vbExclamation, TIT_MSGBOX
        txttotal.SetFocus
        ValidarConciliacion = False
        Exit Function
    End If

    ValidarConciliacion = True
End Function
Function sumaSeleccionados()
     Dim TOTAL As Double
     TOTAL = 0
     For i = 1 To GrdModulos.Rows - 1
        If GrdModulos.TextMatrix(i, 7) = "En Proceso" Then
            TOTAL = TOTAL + GrdModulos.TextMatrix(i, 2)
        End If
    Next
    txtTotalSel.Text = TOTAL
    txtTotalSel.Text = Valido_Importe2(txtTotalSel.Text)
End Function

Private Sub cmdConcilia_Click()
    If buscoaconciliar Then
        fraconcilia.Visible = True
        'lbltarjeta = tabButtons
        lbltarjeta = "TARJETA " & IIf(cbotarjeta.ListIndex > 0, cbotarjeta.Text, GrdModulos.TextMatrix(GrdModulos.RowSel, 1)) '& " - $" & txtTotalSel.Text
        lblVentas = Valido_Importe2(txtTotalSel.Text)
        'txtNeto.Text = txtTotalSel.Text
        cmdConcilia.Enabled = False
        If lbltarjeta = "TARJETA VISA" Or lbltarjeta = "TARJETA VISA DEBITO" Then
            habilitarConciliacion False
        Else
            habilitarConciliacion True
        End If
    Else
        MsgBox "No ha seleccionado ninguna tarjeta a conciliar", vbInformation, TIT_MSGBOX
    End If
End Sub
Private Function buscoaconciliar() As Boolean
    For i = 1 To GrdModulos.Rows - 1
        If GrdModulos.TextMatrix(i, 7) = "En Proceso" Then
            buscoaconciliar = True
            Exit Function
        End If
    Next
    buscoaconciliar = False
End Function

Private Sub cmdCoSalir_Click()
    limpiarConciliacion
    BuscarTarjetas
    fraconcilia.Visible = False
    cmdConcilia.Enabled = True
End Sub
Private Function limpiarConciliacion()
    FechaComprobante.Value = ""
    txtNeto.Text = "0,00"
    txtIva.Text = "0,00"
    txtimp1IVA.Text = "0,00"
    txtSubTotal = "0,00"
    txtImpuestos = "0,00"
    txtperIVA = "0,00"
    txtperIIBB = "0,00"
    txtperGAN = "0,00"
    txttotConc = "0,00"
    txtdeduccionimp = "0,00"
End Function
Private Function habilitarConciliacion(Estado As Boolean)
    'FechaComprobante.Enabled = estado
    'txtNeto.Enabled = estado
    txtIva.Enabled = Estado
    txtimp1IVA.Enabled = Estado
    txtSubTotal.Enabled = Estado
    txtImpuestos.Enabled = Estado
    txtperIVA.Enabled = Estado
    txtperIIBB.Enabled = Estado
    txtperGAN.Enabled = Estado
    txttotConc.Enabled = Estado
    txtdeduccionimp.Enabled = Not Estado
End Function
Private Sub CmdDeselec_Click()
    For i = 1 To GrdModulos.Rows - 1
        GrdModulos.TextMatrix(i, 7) = "NO"
        Call CambiaColorAFilaDeGrilla(GrdModulos, i, vbBlack, vbWhite)
    Next
    GrdModulos.SetFocus
    sumaSeleccionados
End Sub

Private Sub cmdGrabar_Click()
'    On Error GoTo HayError2
'
'    If ValidarEntrada = False Then Exit Sub
'
'        If MsgBox("¿Confirma Movomineto de Mercadería?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
'
'        Screen.MousePointer = vbHourglass
'        lblestado.Caption = "Guardando ..."
'        'DBConn.BeginTrans
'
'        sql = "SELECT EPR_FECHA FROM ENTRADA_PRODUCTO"
'        sql = sql & " WHERE EPR_CODIGO = " & XN(txtNumero.Text)
'        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'
'        If rec.EOF = True Then
'           'INSERTO EN LA TABLA ENTRADA_PRODUCTO
'           sql = "INSERT INTO ENTRADA_PRODUCTO(EPR_CODIGO,EPR_FECHA,VEN_CODIGO,"
'           sql = sql & " STK_CODIGO,ESP_CODIGO,"
'           sql = sql & " EST_CODIGO,EPR_OBSERVACIONES, EPR_HORA)"
'           sql = sql & " VALUES ("
'           sql = sql & XN(txtNumero) & ","
'           sql = sql & XDQ(Fecha.Value) & ","
'           sql = sql & XN(cboEmpleado.ItemData(cboEmpleado.ListIndex)) & ","
'           sql = sql & XN(cboStock.ItemData(cboStock.ListIndex)) & ","
'           sql = sql & XN(cboMovimiento.ItemData(cboMovimiento.ListIndex)) & ","
'           'sql = sql & XN(txtCodCliente.Text) & "," 'SI DEVUELVE PRODUCTOS
'           sql = sql & " 3," 'ESTADO DEFINITIVO
'           sql = sql & XS(txtObservaciones.Text) & ","
'           sql = sql & "#" & Format(Time, "hh:mm") & "#)"
'           DBConn.Execute sql
'
'           'INSERTO EN LA TABLA DETALLE_ENTRADA_PRODUCTO
'           'INSERTO EN LA TABLA DETALLE_ENTRADA_DET_PRODUCTO
'           For i = 1 To GrdModulos.Rows - 1
'               If GrdModulos.TextMatrix(i, 4) = 1 Or GrdModulos.TextMatrix(i, 4) = 3 Then
'                    sql = "INSERT INTO DETALLE_ENTRADA_DET_PRODUCTO(EPR_CODIGO,PTO_CODIGO,DPT_CODIGO,DPT_DETALLE,DEP_CANTIDAD)"
'                    sql = sql & " VALUES ("
'                    sql = sql & XN(txtNumero.Text) & ","
'                    sql = sql & XN(GrdModulos.TextMatrix(i, 4)) & ","
'                    sql = sql & XN(GrdModulos.TextMatrix(i, 5)) & ","
'                    sql = sql & XS(GrdModulos.TextMatrix(i, 1)) & ","
'                    sql = sql & XN(GrdModulos.TextMatrix(i, 2)) & " )"
'                    DBConn.Execute sql
'               Else
'                    sql = "INSERT INTO DETALLE_ENTRADA_PRODUCTO(EPR_CODIGO,PTO_CODIGO,DEP_CANTIDAD)"
'                    sql = sql & " VALUES ("
'                    sql = sql & XN(txtNumero.Text) & ","
'                    sql = sql & XN(GrdModulos.TextMatrix(i, 4)) & ","
'                    sql = sql & XN(GrdModulos.TextMatrix(i, 2)) & " )"
'                    DBConn.Execute sql
'               End If
'           Next
'
'            'ACTUALIZO DETALLE_STOCK
'            For i = 1 To GrdModulos.Rows - 1
'                If GrdModulos.TextMatrix(i, 4) = 1 Or GrdModulos.TextMatrix(i, 4) = 3 Then
'                    sql = "UPDATE PRODUCTO_DETALLE"
'                    sql = sql & " SET PDT_CANTIDAD= PDT_CANTIDAD " & Trim(txtSigno.Text) & XN(GrdModulos.TextMatrix(i, 2))
'                    sql = sql & " WHERE PDT_CODIGO=" & XN(GrdModulos.TextMatrix(i, 5))
'                    sql = sql & " AND PTO_CODIGO=" & XN(GrdModulos.TextMatrix(i, 4))
'                    DBConn.Execute sql
'                Else
'                    sql = "UPDATE STOCK"
'                    sql = sql & " SET DST_STKFIS = DST_STKFIS  " & Trim(txtSigno.Text) & XN(GrdModulos.TextMatrix(i, 2))
'                    sql = sql & " WHERE STK_CODIGO= " & XN(cboStock.ItemData(cboStock.ListIndex))
'                    sql = sql & " AND PTO_CODIGO =" & XN(GrdModulos.TextMatrix(i, 4))
'                    DBConn.Execute sql
'                End If
'            Next
'
'            'ACTUALIZO LA TABLA PARAMENTROS
'            sql = "UPDATE PARAMETROS SET RECEPCION_MERCADERIA=" & XN(txtNumero.Text)
'            DBConn.Execute sql
'        Else
'            MsgBox "La Recepción de Mercadería ya fue registrada", vbCritical, TIT_MSGBOX
'        End If
'        rec.Close
'        Screen.MousePointer = vbNormal
'        lblestado.Caption = ""
'        'DBConn.CommitTrans
'        CmdNuevo_Click
'    Exit Sub
'
'HayError2:
'         lblestado.Caption = ""
'         'DBConn.RollbackTrans
'         If rec.State = 1 Then rec.Close
'         Screen.MousePointer = vbNormal
'         MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Function ValidarEntrada()
'    If cboEmpleado.ListIndex = -1 Then
'        MsgBox "No ha ingresado el Encargado de Depósito", vbExclamation, TIT_MSGBOX
'        cboEmpleado.SetFocus
'        ValidarEntrada = False
'        Exit Function
'    End If
'    If Fecha.Value = "" Then
'        MsgBox "No ha ingresado la Fecha de Entrada de Productos", vbExclamation, TIT_MSGBOX
'        Fecha.SetFocus
'        ValidarEntrada = False
'        Exit Function
'    End If
'    If GrdModulos.Rows = 1 Then
'        MsgBox "Debe haber ingresar al menos un producto en la Grilla ", vbExclamation, TIT_MSGBOX
'        cmdAsignar.SetFocus
'        ValidarEntrada = False
'        Exit Function
'    End If
'    ValidarEntrada = True
End Function

Private Sub CmdNuevo_Click()
    txttotal.Text = ""
    txtObservaciones.Text = ""
    'cboTipo.ListIndex = 0
    cbotarjeta.ListIndex = 0
    fdesdeT.Value = ""
    fhastaT.Value = ""
    GrdModulos.HighLight = flexHighlightNever
    GrdModulos.Rows = 1
    tabDatos.Tab = 0
    fraconcilia.Visible = False
    txtTotalSel.Text = 0
    CmdSelec.Enabled = True
    CmdDeselec.Enabled = True
    cmdConcilia.Enabled = True
    cmdAgregar.Enabled = True
    cmdReporte.Enabled = False
    fraTarjeta.Visible = False
    limpiarConciliacion
End Sub

Private Sub cmdQuitar_Click()
'    If GrdModulos.Rows <> 1 Then
'        If MsgBox("¿Seguro desea Eliminar el Producto: " & Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 1)) & "? ", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
'            lblestado.Caption = "Borrando..."
'            Screen.MousePointer = vbHourglass
'            If GrdModulos.Rows = 2 Then
'                GrdModulos.HighLight = flexHighlightNever
'                GrdModulos.Rows = 1
'                TxtCODIGO.SetFocus
'            Else
'                GrdModulos.RemoveItem (GrdModulos.RowSel)
'                TxtCODIGO.SetFocus
'            End If
'            lblestado.Caption = ""
'            Screen.MousePointer = vbNormal
'        End If
'    End If
End Sub

Private Sub cmdReporte_Click()
'    If FechaDesde.value = "" Then
'        MsgBox "Falta Ingresar la Fecha Desde", vbExclamation, TIT_MSGBOX
'        FechaDesde.SetFocus
'        Exit Sub
'    End If
'    If FechaHasta.value = "" Then
'        MsgBox "Falta Ingresar la Fecha Hasta", vbExclamation, TIT_MSGBOX
'        FechaHasta.SetFocus
'        Exit Sub
'    End If
    
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    
'    Select Case cboDestino.ListIndex
'        Case 0
'            Rep.Destination = crptToWindow
'        Case 1
'            Rep.Destination = crptToPrinter
'        Case 2
'            Rep.Destination = crptToFile
'    End Select
    
    'SOLO FACTURAS DEFINITIVAS
    'Rep.SelectionFormula = " {TMP_CONCILIACION.EST_CODIGO}=3"
    
'    If cboVendedor.List(cboVendedor.ListIndex) <> "(Todos)" Then
'        If Rep.SelectionFormula = "" Then
'            Rep.SelectionFormula = " {TMP_CONCILIACION.VEN_CODIGO}=" & XN(cboVendedor.ItemData(cboVendedor.ListIndex))
'        Else
'            Rep.SelectionFormula = Rep.SelectionFormula & " AND {TMP_CONCILIACION.VEN_CODIGO}=" & XN(cboVendedor.ItemData(cboVendedor.ListIndex))
'        End If
'    End If
    If FechaDesde.Value <> "" Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = " {TMP_CONCILIACION.CON_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {TMP_CONCILIACION.CON_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
        End If
    End If
    If FechaHasta.Value <> "" Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = " {TMP_CONCILIACION.CON_FECHA}<= DATE( " & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {TMP_CONCILIACION.CON_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
        End If
    End If
   
    If FechaDesde.Value <> "" And FechaHasta.Value <> "" Then
        Rep.Formulas(0) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & FechaHasta.Value & "'"
    ElseIf FechaDesde.Value <> "" And FechaHasta.Value = "" Then
        Rep.Formulas(0) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & Date & "'"
    ElseIf FechaDesde.Value = "" And FechaHasta.Value <> "" Then
        Rep.Formulas(0) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & FechaHasta.Value & "'"
    ElseIf FechaDesde.Value = "" And FechaHasta.Value = "" Then
        Rep.Formulas(0) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & Date & "'"
    End If
    
    Rep.WindowState = crptNormal
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR
    
    Rep.WindowTitle = "Listado de Conciliaciones de Tarjetas"
    Rep.ReportFileName = DRIVE & DirReport & "rptconciliaciones.rpt"
    Rep.Action = 1
End Sub

Private Sub CmdSalir_Click()
    If MsgBox("Seguro que desea Salir", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        Set frmTarjetasCredito = Nothing
        Unload Me
    End If
End Sub

Private Sub cndBuscarCliente_Click()
'    frmBuscar.TipoBusqueda = 1
'    frmBuscar.TxtDescriB = ""
'    frmBuscar.Show vbModal
'    If frmBuscar.grdBuscar.Text <> "" Then
'        frmBuscar.grdBuscar.Col = 0
'        txtCodCliente.Text = frmBuscar.grdBuscar.Text
'        txtCodCliente_LostFocus
'        txtCliRazSoc.SetFocus
'    Else
'        txtCodCliente.SetFocus
'    End If
End Sub



Private Sub Command2_Click()

End Sub

Private Sub CmdSelec_Click()
    For i = 1 To GrdModulos.Rows - 1
        GrdModulos.TextMatrix(i, 7) = "En Proceso"
        Call CambiaColorAFilaDeGrilla(GrdModulos, i, vbRed, vbWhite)
    Next
    GrdModulos.SetFocus
    sumaSeleccionados
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 And ActiveControl.Name <> "txtcodigo" And ActiveControl.Name <> "txtdescri" Then
        tabDatos.Tab = 1
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then MySendKeys Chr(9)
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    Set Rec2 = New ADODB.Recordset
    lblestado.Caption = ""
    
    'Call Centrar_pantalla(Me)
    Me.Left = 0
    Me.Top = 0
    preparogrilla
    cmdReporte.Enabled = False
    'CARGO COMBO tipo
    'cargocboTipo
    'CARGO COMBO tarjeta
    cargocboTarjeta
    
    cargocboTarjeta_b
    
    cargocboTarjeta_Tarjeta
    
    tabDatos.Tab = 0
    'GrdModulos.HighLight = flexHighlightNever
    
End Sub
Private Sub preparogrilla()
    'GRILLA DONDE SE CRAGAN LOS PRODUCTOS
    GrdModulos.FormatString = "^Fecha|<Tarjeta|^Monto|Plan|Cupon|Lote|Autorizacion|Conciliacion|TCO_CODIGO|FCL_NUMERO|FCL_SUCURSAL|FPG_CODIGO|PAG_SECUENCIA"
    GrdModulos.ColWidth(0) = 1200 'Fecha
    GrdModulos.ColWidth(1) = 2000 'Tarjeta
    GrdModulos.ColWidth(2) = 1000 'Monto
    GrdModulos.ColWidth(3) = 0 'Plan
    GrdModulos.ColWidth(4) = 1000 'Cupon
    GrdModulos.ColWidth(5) = 1000 'Lote
    GrdModulos.ColWidth(6) = 1300 'Autoriza
    GrdModulos.ColWidth(7) = 1300 'Concilia SI/NO
    GrdModulos.ColWidth(8) = 0 'TCO_CODIGO
    GrdModulos.ColWidth(9) = 0 'FCL_NUMERO
    GrdModulos.ColWidth(10) = 0 'FCL_SUCURSAL
    GrdModulos.ColWidth(11) = 0 'FPG_CODIGO
    GrdModulos.ColWidth(12) = 0 'PAG_SECUENCIA
    GrdModulos.Cols = 13
    GrdModulos.Rows = 1
    GrdModulos.HighLight = flexHighlightWithFocus
    GrdModulos.BorderStyle = flexBorderNone
    GrdModulos.row = 0
    For i = 0 To GrdModulos.Cols - 1
        GrdModulos.Col = i
        GrdModulos.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        GrdModulos.CellBackColor = &H808080    'GRIS OSCURO
        GrdModulos.CellFontBold = True
    Next
    'X para cuando lo recupero de la tabla y tengo que modificarlo
    '"" para cuando no lo recupero de la base
    GrdModulos.Rows = 1
    'GRILLA PARA LA BUSQUEDA
    GRDGrilla.FormatString = "^Numero|^Fecha|<Tarjeta|Total"
    GRDGrilla.ColWidth(0) = 1200 'NUMERO
    GRDGrilla.ColWidth(1) = 1300 'FECHA
    GRDGrilla.ColWidth(2) = 5000 'TARJETA
    GRDGrilla.ColWidth(3) = 1300 'TOTAL
    GRDGrilla.Cols = 4
    GRDGrilla.Rows = 1
    GRDGrilla.HighLight = flexHighlightWithFocus
    GRDGrilla.BorderStyle = flexBorderNone
    GRDGrilla.row = 0
    For i = 0 To GRDGrilla.Cols - 1
        GRDGrilla.Col = i
        GRDGrilla.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        GRDGrilla.CellBackColor = &H808080    'GRIS OSCURO
        GRDGrilla.CellFontBold = True
    Next
    GRDGrilla.Rows = 1
End Sub
Private Function calculototal()
    Dim VTotal As Double
    VTotal = 0
    For i = 1 To GrdModulos.Rows - 1
        VTotal = VTotal + GrdModulos.TextMatrix(i, 2)
    Next
    txttotal.Text = VTotal
    txttotal.Text = VALIDO_IMPORTE(txttotal)
End Function
Private Sub cargocboTipo()
    sql = "SELECT TTA_CODIGO, TTA_DESCRI"
    sql = sql & " FROM TIPO_TARJETA ORDER BY TTA_CODIGO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboTipo.AddItem "(Todos)"
        cbotipo_b.AddItem "(Todos)"
        Do While rec.EOF = False
            cboTipo.AddItem rec!TTA_DESCRI
            cboTipo.ItemData(cboTipo.NewIndex) = rec!TTA_CODIGO
            cbotipo_b.AddItem rec!TTA_DESCRI
            cbotipo_b.ItemData(cboTipo.NewIndex) = rec!TTA_CODIGO
            rec.MoveNext
        Loop
        cboTipo.ListIndex = 0
        cbotipo_b.ListIndex = 0
    End If
    rec.Close
End Sub
Private Sub cargocboTipo_b()
    sql = "SELECT TTA_CODIGO, TTA_DESCRI"
    sql = sql & " FROM TIPO_TARJETA ORDER BY TTA_CODIGO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cbotipo_b.AddItem "(Todos)"
        Do While rec.EOF = False
            cbotipo_b.AddItem rec!TTA_DESCRI
            cbotipo_b.ItemData(cbotipo_b.NewIndex) = rec!TTA_CODIGO
            rec.MoveNext
        Loop
        cbotipo_b.ListIndex = 0
    End If
    rec.Close
End Sub


Private Sub grdGrilla_Click()
    If GRDGrilla.MouseRow = 0 Then
        GRDGrilla.Col = GRDGrilla.MouseCol
        GRDGrilla.ColSel = GRDGrilla.MouseCol
        
        If txtOrden.Text = "A" Then
            GRDGrilla.Sort = 2
            txtOrden.Text = "B"
        Else
            GRDGrilla.Sort = 1
            txtOrden.Text = "A"
        End If
    End If
End Sub

Private Sub GRDGrilla_DblClick()
' CARGAR EL LA CONCILIACION Y EL DETALLE
    tabDatos.Tab = 0
    fraconcilia.Visible = True
    Dim TOTAL As Double
    
    sql = "SELECT * FROM CONCILIACION "
    sql = sql & " WHERE CON_NUMERO=" & GRDGrilla.TextMatrix(GRDGrilla.RowSel, 0)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        lbltarjeta = rec!CON_TARJETA
        FechaComprobante = rec!CON_FECHA
        lblVentas.Caption = rec!CON_VENTAS
        lblVentas.Caption = Valido_Importe2(lblVentas.Caption)
        txtNeto.Text = rec!CON_NETO
        txtNeto.Text = Valido_Importe2(txtNeto.Text)
        txtIva.Text = rec!CON_IVA
        txtIva_LostFocus
        txtImpuestos.Text = rec!CON_IMPUESTOS
        txtImpuestos.Text = Valido_Importe2(txtNeto.Text)
        
        txtperIIBB.Text = rec!CON_PERIIBB
        txtperIIBB.Text = Valido_Importe2(txtperIIBB.Text)
        
        txtperIVA.Text = rec!CON_PERIVA
        txtperIVA.Text = Valido_Importe2(txtperIVA.Text)
        
        txtperGAN.Text = rec!CON_PERGAN
        txtperGAN.Text = Valido_Importe2(txtperGAN.Text)
        
        txttotConc.Text = rec!CON_TOTAL
        txttotConc.Text = Valido_Importe2(txttotConc.Text)
        
    End If
    rec.Close
    
    sql = "SELECT * FROM DETALLE_CONCILIACION "
    sql = sql & " WHERE CON_NUMERO=" & GRDGrilla.TextMatrix(GRDGrilla.RowSel, 0)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    GrdModulos.Rows = 1
    If rec.EOF = False Then
        Do While rec.EOF = False
            GrdModulos.AddItem rec!DCO_FECHA & Chr(9) & rec!TAR_CODIGO & Chr(9) & _
                    Format(rec!DCO_MONTO, "#,##0.00") & Chr(9) & rec!DCO_PLAN & Chr(9) & _
                    rec!DCO_CUPON & Chr(9) & rec!DCO_LOTE & Chr(9) & _
                    rec!DCO_AUTO & Chr(9) & "SI"
                    TOTAL = TOTAL + rec!DCO_MONTO
                    rec.MoveNext
        Loop
    End If
    rec.Close
    txttotal.Text = TOTAL
    txttotal.Text = Valido_Importe2(txttotal)
    'CARGO EL DETALLE DE LA CONCILIACION
    
'    If GRDGrilla.Rows > 1 Then
'        CmdNuevo_Click
'        txtNumero.Text = GRDGrilla.TextMatrix(GRDGrilla.RowSel, 0)
'        Fecha.Value = GRDGrilla.TextMatrix(GRDGrilla.RowSel, 1)
'        txtNumero_LostFocus
'        tabDatos.Tab = 0
'    End If
    CmdSelec.Enabled = False
    CmdDeselec.Enabled = False
    cmdConcilia.Enabled = False
    cmdAgregar.Enabled = False
    cmdReporte.Enabled = False
End Sub

Private Sub grdGrilla_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyReturn Then GRDGrilla_DblClick
End Sub


Private Sub GrdModulos_dblClick()
    If Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 7)) = "NO" Or _
       Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 7)) = "" Then 'NO IMPRIME
        Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbRed, vbWhite)
        GrdModulos.TextMatrix(GrdModulos.RowSel, 7) = "En Proceso"
    Else
        If Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 7)) = "En Proceso" Then
            Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbBlack, vbWhite)
            GrdModulos.TextMatrix(GrdModulos.RowSel, 7) = "NO"
        End If
    End If
    sumaSeleccionados
End Sub

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then GrdModulos_dblClick
End Sub

Private Sub tabDatos_Click(PreviousTab As Integer)
    If tabDatos.Tab = 1 Then
      'cmdGrabar.Enabled = False
      cmdBorrar.Enabled = False
      LimpiarBusqueda
      'If Me.Visible = True Then cboEmpleado1.SetFocus
    Else
      'cmdGrabar.Enabled = True
      cmdBorrar.Enabled = True
    End If
End Sub

Private Sub LimpiarBusqueda()
    'cboEmpleado1.ListIndex = 0
    'cboMovimiento1.ListIndex = 0
    FechaDesde.Value = ""
    FechaHasta.Value = ""
    frameVer.Enabled = False
    GRDGrilla.Rows = 1
End Sub

Private Sub txtCantidad_GotFocus()
    'SelecTexto txtCantidad
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
   ' KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCliRazSoc_Change()
   ' If txtCliRazSoc.Text = "" Then
   '     txtCodCliente.Text = ""
   ' End If
End Sub

Private Sub txtCliRazSoc_KeyPress(KeyAscii As Integer)
    'KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtCliRazSoc_LostFocus()
'    If txtCodCliente.Text = "" And txtCliRazSoc.Text <> "" Then
'        rec.Open BuscoCliente(txtCliRazSoc), DBConn, adOpenStatic, adLockOptimistic
'        If rec.EOF = False Then
'            If rec.RecordCount > 1 Then
'                frmBuscar.TipoBusqueda = 1
'                frmBuscar.TxtDescriB.Text = txtCliRazSoc.Text
'                frmBuscar.Show vbModal
'                If frmBuscar.grdBuscar.Text <> "" Then
'                    frmBuscar.grdBuscar.Col = 0
'                    txtCodCliente.Text = frmBuscar.grdBuscar.Text
'                    frmBuscar.grdBuscar.Col = 1
'                    txtCliRazSoc.Text = frmBuscar.grdBuscar.Text
'                    rec.Close
'                    txtCodCliente_LostFocus
'                    TxtCODIGO.SetFocus
'                Else
'                    txtCodCliente.SetFocus
'                End If
'            Else
'                txtCodCliente.Text = rec!CLI_CODIGO
'                txtCliRazSoc.Text = rec!CLI_RAZSOC
'                rec.Close
'            End If
'        Else
'            rec.Close
'            MsgBox "No se encontro el Cliente", vbExclamation, TIT_MSGBOX
'            txtCodCliente.SetFocus
'        End If
'    ElseIf txtCodCliente.Text = "" And txtCliRazSoc.Text = "" Then
'        MsgBox "Debe elegir un cliente", vbExclamation, TIT_MSGBOX
'        txtCodCliente.SetFocus
'    End If
End Sub

Private Sub txtCodCliente_Change()
    'If txtCodCliente.Text = "" Then
    '    txtCliRazSoc.Text = ""
    'End If
End Sub

Private Sub txtCodCliente_GotFocus()
    'SelecTexto txtCodCliente
End Sub

Private Sub txtCodCliente_KeyPress(KeyAscii As Integer)
    'KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCodCliente_LostFocus()
'    If txtCodCliente.Text <> "" Then
'        rec.Open BuscoCliente(txtCodCliente), DBConn, adOpenStatic, adLockOptimistic
'
'        If rec.EOF = False Then
'            txtCliRazSoc.Text = rec!CLI_RAZSOC
'        Else
'            MsgBox "El Cliente no existe", vbExclamation, TIT_MSGBOX
'            txtCodCliente.SetFocus
'        End If
'        rec.Close
'    End If
End Sub

Private Function BuscoCliente(Codigo As String) As String
'        sql = "SELECT C.CLI_CODIGO, C.CLI_RAZSOC"
'        sql = sql & " FROM CLIENTE C"
'        sql = sql & " WHERE"
'        If txtCodCliente.Text <> "" Then
'            sql = sql & " C.CLI_CODIGO=" & XN(Codigo)
'        Else
'            sql = sql & " C.CLI_RAZSOC LIKE '" & Trim(Codigo) & "%'"
'        End If
'        BuscoCliente = sql
End Function

Private Sub TxtCodigo_Change()
'    If TxtCODIGO.Text = "" Then
'        TxtCODIGO.Text = ""
'        TxtDescri.Text = ""
'        txtCantidad.Text = ""
'        TxtCodInt.Text = ""
'        cmdAsignar.Enabled = False
'    Else
'        cmdAsignar.Enabled = True
'    End If
End Sub

Private Sub TxtCodigo_GotFocus()
    'SelecTexto TxtCODIGO
End Sub

Private Sub txtcodigo_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyF1 Then
'        BuscarProducto "CODIGO"
'        TxtCODIGO.SetFocus
'    End If
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
   'KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub cargocboTarjeta_b()
    cbotarjeta_b.AddItem "Todos"
    cbotarjeta_b.ItemData(cbotarjeta_b.NewIndex) = 0 '19 Y 20
    
    cbotarjeta_b.AddItem "CABAL"
    cbotarjeta_b.ItemData(cbotarjeta_b.NewIndex) = 1 '19 Y 20
    
    cbotarjeta_b.AddItem "AMERICAN EXPRESS"
    cbotarjeta_b.ItemData(cbotarjeta_b.NewIndex) = 2 '18
    
    cbotarjeta_b.AddItem "MS-DEBIT"
    cbotarjeta_b.ItemData(cbotarjeta_b.NewIndex) = 3 '23
    
    cbotarjeta_b.AddItem "MAESTRO"
    cbotarjeta_b.ItemData(cbotarjeta_b.NewIndex) = 4 '12,25,22
    
    cbotarjeta_b.AddItem "MC CREDITO - BANCOR"
    cbotarjeta_b.ItemData(cbotarjeta_b.NewIndex) = 5 '6,21
    
    cbotarjeta_b.AddItem "NARANJA"
    cbotarjeta_b.ItemData(cbotarjeta_b.NewIndex) = 6 '1
    
    cbotarjeta_b.AddItem "VISA"
    cbotarjeta_b.ItemData(cbotarjeta_b.NewIndex) = 7 '5,17
    
    cbotarjeta_b.AddItem "DINNERS CLUB"
    cbotarjeta_b.ItemData(cbotarjeta_b.NewIndex) = 8 '24
    
    cbotarjeta_b.ListIndex = 0
End Sub

Private Sub cargocboTarjeta()
    cbotarjeta.AddItem "Todos"
    cbotarjeta.ItemData(cbotarjeta.NewIndex) = 0 '19 Y 20
    
    cbotarjeta.AddItem "CABAL"
    cbotarjeta.ItemData(cbotarjeta.NewIndex) = 1 '19 Y 20
    
    cbotarjeta.AddItem "AMERICAN EXPRESS"
    cbotarjeta.ItemData(cbotarjeta.NewIndex) = 2 '18
    
    cbotarjeta.AddItem "MS-DEBIT"
    cbotarjeta.ItemData(cbotarjeta.NewIndex) = 3 '23
    
    cbotarjeta.AddItem "MAESTRO"
    cbotarjeta.ItemData(cbotarjeta.NewIndex) = 4 '12,25,22
    
    cbotarjeta.AddItem "MC CREDITO - BANCOR"
    cbotarjeta.ItemData(cbotarjeta.NewIndex) = 5 '6,21
    
    cbotarjeta.AddItem "NARANJA"
    cbotarjeta.ItemData(cbotarjeta.NewIndex) = 6 '1
    
    cbotarjeta.AddItem "VISA"
    cbotarjeta.ItemData(cbotarjeta.NewIndex) = 7 '5,17
    
    cbotarjeta.AddItem "DINNERS CLUB"
    cbotarjeta.ItemData(cbotarjeta.NewIndex) = 8 '24
    
    cbotarjeta.ListIndex = 0

'
'
'    sql = "SELECT TAR_CODIGO, TAR_DESCRI "
'    sql = sql & " FROM TARJETA "
'    If cboTipo.ListIndex > 0 Then
'        sql = sql & " WHERE TTA_CODIGO = " & cboTipo.ItemData(cboTipo.ListIndex)
'    End If
'    'sql = sql & " ORDER BY S.STK_CODIGO"
'    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    If rec.EOF = False Then
'        cbotarjeta.AddItem "(Todos)"
'        Do While rec.EOF = False
'            cbotarjeta.AddItem rec!TAR_DESCRI
'            cbotarjeta.ItemData(cbotarjeta.NewIndex) = rec!TAR_CODIGO
'            rec.MoveNext
'        Loop
'        cbotarjeta.ListIndex = 0
'    End If
'    rec.Close
End Sub
Private Sub cargocboTarjeta_Tarjeta()
    cbotarjeta_tarjeta.AddItem "Todos"
    cbotarjeta_tarjeta.ItemData(cbotarjeta_tarjeta.NewIndex) = 0 '19 Y 20
    
    cbotarjeta_tarjeta.AddItem "CABAL"
    cbotarjeta_tarjeta.ItemData(cbotarjeta_tarjeta.NewIndex) = 1 '19 Y 20
    
    cbotarjeta_tarjeta.AddItem "AMERICAN EXPRESS"
    cbotarjeta_tarjeta.ItemData(cbotarjeta_tarjeta.NewIndex) = 2 '18
    
    cbotarjeta_tarjeta.AddItem "MS-DEBIT"
    cbotarjeta_tarjeta.ItemData(cbotarjeta_tarjeta.NewIndex) = 3 '23
    
    cbotarjeta_tarjeta.AddItem "MAESTRO"
    cbotarjeta_tarjeta.ItemData(cbotarjeta_tarjeta.NewIndex) = 4 '12,25,22
    
    cbotarjeta_tarjeta.AddItem "MC CREDITO - BANCOR"
    cbotarjeta_tarjeta.ItemData(cbotarjeta_tarjeta.NewIndex) = 5 '6,21
    
    cbotarjeta_tarjeta.AddItem "NARANJA"
    cbotarjeta_tarjeta.ItemData(cbotarjeta_tarjeta.NewIndex) = 6 '1
    
    cbotarjeta_tarjeta.AddItem "VISA"
    cbotarjeta_tarjeta.ItemData(cbotarjeta_tarjeta.NewIndex) = 7 '5,17
    
    cbotarjeta_tarjeta.AddItem "DINNERS CLUB"
    cbotarjeta_tarjeta.ItemData(cbotarjeta_tarjeta.NewIndex) = 8 '24
    
    cbotarjeta_tarjeta.ListIndex = 0

'
'
'    sql = "SELECT TAR_CODIGO, TAR_DESCRI "
'    sql = sql & " FROM TARJETA "
'    If cboTipo.ListIndex > 0 Then
'        sql = sql & " WHERE TTA_CODIGO = " & cboTipo.ItemData(cboTipo.ListIndex)
'    End If
'    'sql = sql & " ORDER BY S.STK_CODIGO"
'    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    If rec.EOF = False Then
'        cbotarjeta.AddItem "(Todos)"
'        Do While rec.EOF = False
'            cbotarjeta.AddItem rec!TAR_DESCRI
'            cbotarjeta.ItemData(cbotarjeta.NewIndex) = rec!TAR_CODIGO
'            rec.MoveNext
'        Loop
'        cbotarjeta.ListIndex = 0
'    End If
'    rec.Close
End Sub

Private Sub txtdescri_Change()
'    If TxtDescri.Text = "" Then
'        TxtCODIGO.Text = ""
'    End If
End Sub

Private Sub txtdescri_GotFocus()
    'SelecTexto TxtDescri
End Sub

Private Sub txtdescri_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyF1 Then
'        BuscarProducto "CODIGO"
'        TxtDescri.SetFocus
'    End If
End Sub

Private Sub txtdescri_KeyPress(KeyAscii As Integer)
    'KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtDescri_LostFocus()
'   If TxtCODIGO.Text = "" And TxtDescri.Text <> "" Then
'        Set Rec1 = New ADODB.Recordset
'        Screen.MousePointer = vbHourglass
'        sql = "SELECT PTO_CODIGO, PTO_DESCRI, PTO_CODBARRAS"
'        sql = sql & " FROM PRODUCTO"
'        sql = sql & " WHERE PTO_DESCRI LIKE '" & TxtDescri.Text & "%'"
'        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
'        If Rec1.EOF = False Then
'            If Rec1.RecordCount > 1 Then
'                'grdGrilla.SetFocus
'                BuscarProducto "CADENA", Trim(TxtDescri.Text)
'                TxtDescri.SetFocus
'            Else
'                TxtCODIGO.Text = Trim(ChkNull(Rec1!PTO_CODBARRAS))
'                TxtDescri.Text = Trim(Rec1!PTO_DESCRI)
'                TxtCodInt.Text = Trim(Rec1!PTO_CODIGO)
'            End If
'        Else
'                MsgBox "No se ha encontrado el Producto", vbExclamation, TIT_MSGBOX
'                TxtDescri.Text = ""
'        End If
'        Rec1.Close
'        Screen.MousePointer = vbNormal
'    ElseIf TxtCODIGO.Text = "" And TxtDescri.Text = "" Then
'        cmdAsignar.Enabled = False
'    End If
End Sub

Private Sub txtNumero_GotFocus()
    'SelecTexto txtNumero
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
    'KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNumero_LostFocus()
'    If txtNumero.Text <> "" Then
'        Set Rec1 = New ADODB.Recordset
'        sql = "SELECT * FROM ENTRADA_PRODUCTO"
'        sql = sql & " WHERE EPR_CODIGO=" & XN(txtNumero)
'        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
'        If Rec1.EOF = False Then
'            Fecha.Value = Rec1!EPR_FECHA
'            Call BuscaCodigoProxItemData(Rec1!VEN_CODIGO, cboEmpleado)
'            Call BuscaCodigoProxItemData(Rec1!STK_CODIGO, cboStock)
'            Call BuscaCodigoProxItemData(Rec1!ESP_CODIGO, cboMovimiento)
''            If Not IsNull(Rec1!CLI_CODIGO) Then
''                txtCodCliente.Text = Rec1!CLI_CODIGO
''                txtCodCliente_LostFocus
''            Else
''                txtCodCliente.Text = ""
''            End If
'            CargoGrilla (txtNumero)
'            Call BuscoEstado(CInt(Rec1!EST_CODIGO), lblEstadoRecepcion)
'            txtObservaciones.Text = ChkNull(Rec1!EPR_OBSERVACIONES)
'            If Rec1!EST_CODIGO = 2 Then
'               cmdBorrar.Enabled = False
'            Else
'               cmdBorrar.Enabled = True
'            End If
'            cmdGrabar.Enabled = False
'            FrameGeneral.Enabled = False
'            FrameProducto.Enabled = False
'        Else
'            MsgBox "El Movimiento no existe", vbExclamation, TIT_MSGBOX
'            CmdNuevo_Click
'            cboStock.SetFocus
'        End If
'        Rec1.Close
'    End If
End Sub

Private Sub CargoGrilla(Campo As Integer)
'    Dim Rec2 As ADODB.Recordset
'    Set Rec2 = New ADODB.Recordset
'    ' busco en DETALLE DE entrada de producto
'    Screen.MousePointer = vbHourglass
'    sql = "SELECT DISTINCT  P.PTO_DESCRI, P.PTO_CODBARRAS,"
'    sql = sql & " D.DEP_CANTIDAD, E.EPR_CODIGO, E.EPR_FECHA,P.PTO_CODIGO"
'    sql = sql & " FROM ENTRADA_PRODUCTO E, PRODUCTO P, DETALLE_ENTRADA_PRODUCTO D"
'    sql = sql & " WHERE P.PTO_CODIGO = D.PTO_CODIGO AND D.EPR_CODIGO = E.EPR_CODIGO"
'    sql = sql & " AND E.EPR_CODIGO = " & Campo & " ORDER BY E.EPR_CODIGO"
'
'    lblestado.Caption = "Buscando..."
'    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    If rec.EOF = False Then
'        GrdModulos.Rows = 1
'        GrdModulos.HighLight = flexHighlightAlways
'        Do While Not rec.EOF
'           GrdModulos.AddItem IIf(IsNull(rec!PTO_CODBARRAS), rec!PTO_CODIGO, rec!PTO_CODBARRAS) & Chr(9) & Trim(rec!PTO_DESCRI) _
'                              & Chr(9) & rec!DEP_CANTIDAD & Chr(9) & "X" & Chr(9) & rec!PTO_CODIGO
'
'           rec.MoveNext
'        Loop
'        rec.MoveFirst
'    Else
'        'busco en detalle de entrada de detaleproducto - ' PARA COMBUSTIBLES !
'        sql = "SELECT DISTINCT  D.DPT_DETALLE, D.PTO_CODIGO,D.DPT_CODIGO,"
'        sql = sql & " D.DEP_CANTIDAD, E.EPR_CODIGO, E.EPR_FECHA"
'        sql = sql & " FROM ENTRADA_PRODUCTO E, PRODUCTO P, DETALLE_ENTRADA_DET_PRODUCTO D"
'        sql = sql & " WHERE P.PTO_CODIGO = D.PTO_CODIGO AND D.EPR_CODIGO = E.EPR_CODIGO"
'        sql = sql & " AND E.EPR_CODIGO = " & Campo & " ORDER BY E.EPR_CODIGO, D.PTO_CODIGO,D.DPT_CODIGO"
'
'        lblestado.Caption = "Buscando..."
'        Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
'        If Rec2.EOF = False Then
'            GrdModulos.Rows = 1
'            GrdModulos.HighLight = flexHighlightAlways
'            Do While Not Rec2.EOF
'               GrdModulos.AddItem Rec2!PTO_CODIGO & Chr(9) & Trim(Rec2!DPT_DETALLE) _
'                                  & Chr(9) & Rec2!DEP_CANTIDAD & Chr(9) & "" & Chr(9) & Rec2!PTO_CODIGO & Chr(9) & Rec2!DPT_CODIGO
'
'               Rec2.MoveNext
'            Loop
'        End If
'
'    End If
'    If GrdModulos.Rows = 1 Then
'        lblestado.Caption = ""
'        MsgBox "No hay coincidencias en la busqueda.", vbOKOnly + vbCritical, TIT_MSGBOX
'        'Me.txtNumero.SetFocus
'    End If
'    rec.Close
'    Rec2.Close
'    Screen.MousePointer = vbNormal
'    lblestado.Caption = ""
    
End Sub

Private Sub txtdeduccionimp_GotFocus()
    seltxt
End Sub

Private Sub txtdeduccionimp_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtdeduccionimp, KeyAscii)
End Sub

Private Sub txtdeduccionimp_LostFocus()
    If txtdeduccionimp.Text <> "" Then
        txtdeduccionimp.Text = Valido_Importe2(txtdeduccionimp.Text)
        SumaTotal
    Else
        txtdeduccionimp.Text = "0,00"
    End If
End Sub

Private Sub txtImpuestos_GotFocus()
    SelecTexto txtImpuestos
End Sub

Private Sub txtImpuestos_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtImpuestos, KeyAscii)
End Sub

Private Sub txtImpuestos_LostFocus()
    If txtImpuestos.Text <> "" Then
        txtImpuestos.Text = Valido_Importe2(txtImpuestos.Text)
        SumaTotal
    Else
        txtImpuestos.Text = "0,00"
    End If
End Sub
Private Sub txtIva_GotFocus()
    SelecTexto txtIva
End Sub

Private Sub txtIva_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtIva, KeyAscii)
End Sub

Private Sub txtIva_LostFocus()
    If txtIva.Text <> "" Then
        If ValidarPorcentaje(txtIva) = False Then
            txtIva.SetFocus
            Exit Sub
        End If
        txtimp1IVA.Text = Valido_Importe2((CDbl(txtNeto.Text) * CDbl(txtIva.Text)) / 100)
        txtSubTotal.Text = CDbl(txtNeto.Text) + ((CDbl(txtNeto.Text) * CDbl(txtIva.Text)) / 100)
        txtSubTotal.Text = Valido_Importe2(txtSubTotal)
        'txttotal.Text = CDbl(txtSubTotal1.Text) + CDbl(txtSubTotal.Text) + CDbl(txtImpuestos.Text)
        'txttotal.Text = Valido_Importe2(txttotal)
    End If
End Sub

Private Sub txtNeto_GotFocus()
    SelecTexto txtNeto
End Sub

Private Sub txtNeto_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtNeto, KeyAscii)
End Sub

Private Sub txtNeto_LostFocus()
    If txtNeto.Text <> "" Then
        txtNeto.Text = Valido_Importe2(txtNeto)
        SumaTotal
    Else
        txtNeto.Text = "0,00"
    End If
End Sub


Private Sub txtObservaciones_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Public Sub BuscarTarjetas()
    GrdModulos.Rows = 1
    Dim TOTAL As Double
    sql = "SELECT F.FCL_FECHA,T.TAR_DESCRI,FP.PAG_IMPORTE,FP.TAR_PLAN,FP.TAR_CUPON,FP.TAR_LOTE,FP.TAR_AUTORIZACION"
    sql = sql & ",F.FCL_NUMERO,FP.TCO_CODIGO,FP.FCL_SUCURSAL,FP.FPG_CODIGO,FP.PAG_SECUENCIA"
    sql = sql & " FROM FACTURA_CLIENTE F, FACTURA_PAGOS FP,TARJETA T" ', TARJETA_PLAN TP"
    sql = sql & " WHERE F.FCL_NUMERO = FP.FCL_NUMERO"
    sql = sql & " AND F.FCL_SUCURSAL = FP.FCL_SUCURSAL"
    sql = sql & " AND FP.TAR_CODIGO = T.TAR_CODIGO"
    'sql = sql & " AND FP.TAR_PLAN = TP.PLA_CODIGO"
    'TENGO QUE BUSCAR LOS QUE NO TIENEN CON_NUMERO (COM_NUMERO IS NULL O VACIO)
    sql = sql & " AND ( ISNULL(FP.CON_NUMERO) OR FP.CON_NUMERO=0) "
    If fdesdeT.Value <> "" Then sql = sql & " AND F.FCL_FECHA>=" & XDQ(fdesdeT)
    If fhastaT.Value <> "" Then sql = sql & " AND F.FCL_FECHA<=" & XDQ(fhastaT)
    'If cboTipo.ListIndex > 0 Then sql = sql & " AND FP.FPG_CODIGO=" & cboTipo.ItemData(cboTipo.ListIndex) + 2 'FPG_CODIGO 3 CREDITO Y 4 DEBITO
    If cbotarjeta.ListIndex > 0 Then
        Select Case cbotarjeta.ItemData(cbotarjeta.ListIndex)
        Case 1
            sql = sql & " AND FP.TAR_CODIGO IN (19,20)"
        Case 2
            sql = sql & " AND FP.TAR_CODIGO IN (18)"
        Case 3
            sql = sql & " AND FP.TAR_CODIGO IN (23)"
        Case 4
            sql = sql & " AND FP.TAR_CODIGO IN (12,22,25)"
        Case 5
            sql = sql & " AND FP.TAR_CODIGO IN (6,21)"
        Case 6
            sql = sql & " AND FP.TAR_CODIGO IN (1)"
        Case 7
            sql = sql & " AND FP.TAR_CODIGO IN (5,17)"
        Case 8
            sql = sql & " AND FP.TAR_CODIGO IN (19,20)"
        End Select
        
    End If
    
    sql = sql & " ORDER BY F.FCL_FECHA"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        TOTAL = 0
        Do While rec.EOF = False
            GrdModulos.AddItem rec!FCL_FECHA & Chr(9) & rec!TAR_DESCRI & Chr(9) & _
                                Format(rec!PAG_IMPORTE, "#,##0.00") & Chr(9) & rec!TAR_PLAN & Chr(9) & _
                                rec!TAR_CUPON & Chr(9) & rec!TAR_LOTE & Chr(9) & _
                                rec!TAR_AUTORIZACION & Chr(9) & "NO" & Chr(9) & _
                                rec!TCO_CODIGO & Chr(9) & rec!FCL_NUMERO & Chr(9) & _
                                rec!FCL_SUCURSAL & Chr(9) & rec!FPG_CODIGO & Chr(9) & _
                                rec!PAG_SECUENCIA
                                TOTAL = TOTAL + rec!PAG_IMPORTE
            
            rec.MoveNext
        Loop
    Else
        MsgBox "No se encontraron tarjetas a conciliar", vbExclamation, TIT_MSGBOX
    End If
    txttotal = TOTAL
    txttotal = Valido_Importe2(txttotal)
    rec.Close
End Sub
Public Sub BuscarConciliaciones()
    'Dim TOTAL As Double
    sql = "SELECT DISTINCT C.*"
    sql = sql & " FROM CONCILIACION C, DETALLE_CONCILIACION FP"
    sql = sql & " WHERE C.CON_NUMERO=FP.CON_NUMERO "
    If FechaDesde.Value <> "" Then sql = sql & " AND CON_FECHA>=" & XDQ(FechaDesde)
    If FechaHasta.Value <> "" Then sql = sql & " AND CON_FECHA<=" & XDQ(FechaHasta)
    'If cboTipo.ListIndex > 0 Then sql = sql & " AND FP.FPG_CODIGO=" & cboTipo.ItemData(cboTipo.ListIndex) + 2 'FPG_CODIGO 3 CREDITO Y 4 DEBITO
    If cbotarjeta_b.ListIndex > 0 Then
        Select Case cbotarjeta_b.ItemData(cbotarjeta_b.ListIndex)
'        Case 1
'            sql = sql & " AND FP.TAR_CODIGO IN (19,20)"
'        Case 2
'            sql = sql & " AND FP.TAR_CODIGO IN (18)"
'        Case 3
'            sql = sql & " AND FP.TAR_CODIGO IN (23)"
'        Case 4
'            sql = sql & " AND FP.TAR_CODIGO IN (12,22,25)"
'        Case 5
'            sql = sql & " AND FP.TAR_CODIGO IN (6,21)"
'        Case 6
'            sql = sql & " AND FP.TAR_CODIGO IN (1)"
'        Case 7
'            sql = sql & " AND FP.TAR_CODIGO IN (5,17)"
'        Case 8
'            sql = sql & " AND FP.TAR_CODIGO IN (19,20)"
            
        Case 1
            sql = sql & " AND FP.TAR_CODIGO IN ('CABAL','CABAL DEBITO')"
        Case 2
            sql = sql & " AND FP.TAR_CODIGO IN ('AMERICAN EXPRESS')"
        Case 3
            sql = sql & " AND FP.TAR_CODIGO IN ('MC-DEBIT','MS-DEBIT')"
        Case 4
            sql = sql & " AND FP.TAR_CODIGO IN ('MAESTRO','MC-BANCOR DEBITO')"
        Case 5
            sql = sql & " AND FP.TAR_CODIGO IN ('MASTERCARD','MC-BANCOR')"
        Case 6
            sql = sql & " AND FP.TAR_CODIGO IN ('NARANJA')"
        Case 7
            sql = sql & " AND FP.TAR_CODIGO IN ('VISA','VISA DEBITO')"
        Case 8
            sql = sql & " AND FP.TAR_CODIGO IN ('CABAL','CABAL DEBITO')"
        End Select
        
    End If
    
    sql = sql & " ORDER BY CON_FECHA"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    GRDGrilla.Rows = 1
    DBConn.Execute "DELETE FROM TMP_CONCILIACION"
    
    If rec.EOF = False Then
        'TOTAL = 0
        Do While rec.EOF = False
            GRDGrilla.AddItem rec!CON_NUMERO & Chr(9) & rec!CON_FECHA & Chr(9) & rec!CON_TARJETA & Chr(9) & Valido_Importe2(rec!CON_TOTAL)
            
            sql = "INSERT INTO TMP_CONCILIACION"
            sql = sql & "(CON_NUMERO,CON_TARJETA,CON_FECHA,CON_VENTAS,CON_NETO,CON_IVA,CON_IMP1IVA,CON_IMPUESTOS,"
            sql = sql & "CON_PERIIBB,CON_PERIVA,CON_PERGAN,CON_TOTAL,CON_DEDUIMP)"
            sql = sql & " VALUES ("
            sql = sql & rec!CON_NUMERO & ","
            sql = sql & XS(rec!CON_TARJETA) & ","
            sql = sql & XDQ(rec!CON_FECHA) & ","
            sql = sql & XN(Chk0(rec!CON_VENTAS)) & ","
            sql = sql & XN(Chk0(rec!CON_NETO)) & ","
            sql = sql & XN(Chk0(rec!CON_IVA)) & ","
            sql = sql & XN(Chk0(rec!CON_IMP1IVA)) & ","
            sql = sql & XN(Chk0(rec!CON_IMPUESTOS)) & "," 'exento
            sql = sql & XN(Chk0(rec!CON_PERIIBB)) & ","
            sql = sql & XN(Chk0(rec!CON_PERIVA)) & ","
            sql = sql & XN(Chk0(rec!CON_PERGAN)) & ","
            sql = sql & XN(Chk0(rec!CON_TOTAL)) & ","
            sql = sql & XN(Chk0(rec!CON_DEDUIMP)) & ")"
            DBConn.Execute sql
            
            
            rec.MoveNext
        Loop
    End If
    'txttotal = TOTAL
    'txttotal = valido_importe2(txttotal)
    rec.Close
    If GRDGrilla.Rows > 1 Then
        cmdReporte.Enabled = True
    End If
End Sub

Private Sub txtperIVA_GotFocus()
    seltxt
End Sub

Private Sub txtperIVA_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtperIVA, KeyAscii)
End Sub

Private Sub txtperIVA_LostFocus()
    txtperIVA.Text = Valido_Importe2(txtperIVA.Text)
    SumaTotal
End Sub

Private Sub txtperIIBB_GotFocus()
    seltxt
End Sub

Private Sub txtperIIBB_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtperIIBB, KeyAscii)
End Sub

Private Sub txtperIIBB_LostFocus()
    txtperIIBB.Text = Valido_Importe2(txtperIIBB.Text)
    SumaTotal
End Sub

Private Sub txtperGAN_GotFocus()
    seltxt
End Sub

Private Sub txtperGAN_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtperGAN, KeyAscii)
End Sub

Private Sub txtperGAN_LostFocus()
    txtperGAN.Text = Valido_Importe2(txtperGAN.Text)
    SumaTotal
End Sub

Private Sub txtSubTotal_LostFocus()
    SumaTotal
End Sub

Private Sub txttarjeta_importe_GotFocus()
    seltxt
End Sub

Private Sub txttarjeta_importe_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txttarjeta_importe, KeyAscii)
End Sub

Private Sub txttarjeta_importe_LostFocus()
    txttarjeta_importe = Valido_Importe2(txttarjeta_importe)
End Sub
