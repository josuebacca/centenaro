VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEntradaProductos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimiento de Mercader�a"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9345
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
   Picture         =   "frmEntradaProductos.frx":0000
   ScaleHeight     =   6495
   ScaleWidth      =   9345
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   450
      Left            =   5955
      Picture         =   "frmEntradaProductos.frx":0D82
      TabIndex        =   12
      Top             =   6015
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   450
      Left            =   4830
      Picture         =   "frmEntradaProductos.frx":108C
      TabIndex        =   11
      Top             =   6015
      Width           =   1095
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   8205
      Picture         =   "frmEntradaProductos.frx":1396
      TabIndex        =   14
      Top             =   6015
      Width           =   1095
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&An&ular"
      Height          =   450
      Left            =   7080
      Picture         =   "frmEntradaProductos.frx":16A0
      TabIndex        =   13
      Top             =   6015
      Width           =   1095
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   5940
      Left            =   15
      TabIndex        =   22
      Top             =   30
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   10478
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   529
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
      TabPicture(0)   =   "frmEntradaProductos.frx":19AA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameGeneral"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FrameProducto"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtObservaciones"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "frmEntradaProductos.frx":19C6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GRDGrilla"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(2)=   "frameVer"
      Tab(1).ControlCount=   3
      Begin VB.TextBox txtObservaciones 
         Height          =   465
         Left            =   1275
         MaxLength       =   199
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   5385
         Width           =   7890
      End
      Begin VB.Frame frameVer 
         Caption         =   "Ver..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   -74910
         TabIndex        =   48
         Top             =   6480
         Width           =   9090
         Begin VB.OptionButton optSeleccion 
            Alignment       =   1  'Right Justify
            Caption         =   "... Listar Seleccionado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5280
            TabIndex        =   50
            Top             =   210
            Width           =   1935
         End
         Begin VB.OptionButton optTodos 
            Alignment       =   1  'Right Justify
            Caption         =   "... Listar Todos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1755
            TabIndex        =   49
            Top             =   210
            Value           =   -1  'True
            Width           =   1380
         End
      End
      Begin VB.Frame Frame4 
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
         Height          =   1440
         Left            =   -74880
         TabIndex        =   23
         Top             =   345
         Width           =   9045
         Begin VB.ComboBox cboMovimiento1 
            Height          =   315
            Left            =   1965
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   675
            Width           =   3645
         End
         Begin VB.TextBox txtOrden 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   105
            TabIndex        =   51
            Text            =   "A"
            Top             =   480
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.ComboBox cboEmpleado1 
            Height          =   315
            Left            =   1965
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   330
            Width           =   3645
         End
         Begin VB.CommandButton CmdBuscAprox 
            Caption         =   "Buscar"
            Height          =   420
            Left            =   6300
            MaskColor       =   &H8000000F&
            TabIndex        =   20
            ToolTipText     =   "Buscar Nota de Pedido"
            Top             =   885
            UseMaskColor    =   -1  'True
            Width           =   2175
         End
         Begin MSComCtl2.DTPicker FechaDesde 
            Height          =   315
            Left            =   1965
            TabIndex        =   18
            Top             =   1080
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
            Left            =   4560
            TabIndex        =   19
            Top             =   1080
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   51970049
            CurrentDate     =   41098
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Movimiento:"
            Height          =   195
            Left            =   900
            TabIndex        =   54
            Top             =   735
            Width           =   870
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor:"
            Height          =   195
            Left            =   900
            TabIndex        =   35
            Top             =   360
            Width           =   750
         End
         Begin VB.Label lblFechaDesde 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   900
            TabIndex        =   25
            Top             =   1065
            Width           =   990
         End
         Begin VB.Label lblFechaHasta 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   3525
            TabIndex        =   24
            Top             =   1080
            Width           =   960
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3870
         Left            =   -74655
         TabIndex        =   26
         Top             =   2340
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   6826
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
         Height          =   3960
         Left            =   -74895
         TabIndex        =   21
         Top             =   1830
         Width           =   9105
         _ExtentX        =   16060
         _ExtentY        =   6985
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame FrameProducto 
         Caption         =   "Producto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3765
         Left            =   90
         TabIndex        =   28
         Top             =   1605
         Width           =   9105
         Begin VB.Frame fraTanque 
            Height          =   830
            Left            =   3840
            TabIndex        =   56
            Top             =   30
            Visible         =   0   'False
            Width           =   1815
            Begin VB.OptionButton optTanque2 
               Caption         =   "Tanque 2"
               Height          =   255
               Left            =   240
               TabIndex        =   58
               Top             =   480
               Width           =   1215
            End
            Begin VB.OptionButton optTanque1 
               Caption         =   "Tanque 1"
               Height          =   195
               Left            =   240
               TabIndex        =   57
               Top             =   240
               Value           =   -1  'True
               Width           =   975
            End
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
            Left            =   1140
            TabIndex        =   7
            Top             =   480
            Width           =   4470
         End
         Begin VB.TextBox txtCantidad 
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
            Left            =   5640
            MaxLength       =   10
            TabIndex        =   8
            Top             =   480
            Width           =   885
         End
         Begin VB.CommandButton cmdBuscarProducto 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6585
            MaskColor       =   &H000000FF&
            Picture         =   "frmEntradaProductos.frx":19E2
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Buscar Producto"
            Top             =   480
            UseMaskColor    =   -1  'True
            Width           =   405
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
            Left            =   105
            TabIndex        =   6
            Top             =   480
            Width           =   1005
         End
         Begin VB.CommandButton cmdQuitar 
            Caption         =   "Quitar"
            Height          =   330
            Left            =   8055
            TabIndex        =   29
            ToolTipText     =   "Quitar Producto"
            Top             =   480
            Width           =   900
         End
         Begin MSFlexGridLib.MSFlexGrid GrdModulos 
            Height          =   2850
            Left            =   75
            TabIndex        =   15
            Top             =   825
            Width           =   8910
            _ExtentX        =   15716
            _ExtentY        =   5027
            _Version        =   393216
            Cols            =   5
            FixedCols       =   0
            RowHeightMin    =   280
            BackColorSel    =   16761024
            FocusRect       =   0
            SelectionMode   =   1
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
         Begin VB.TextBox txtCodInt 
            Height          =   345
            Left            =   6570
            TabIndex        =   55
            Top             =   195
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.CommandButton cmdAsignar 
            Caption         =   "A&gregar"
            Height          =   330
            Left            =   7110
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Asignar Producto"
            Top             =   480
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "C�digo"
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
            Left            =   180
            TabIndex        =   34
            Top             =   240
            Width           =   555
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad"
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
            Left            =   5685
            TabIndex        =   33
            ToolTipText     =   "Agregar Producto"
            Top             =   240
            Width           =   705
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Descripci�n"
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
            Left            =   1185
            TabIndex        =   32
            Top             =   240
            Width           =   915
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Devoluci�n Mercader�a"
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
         Left            =   45
         TabIndex        =   45
         Top             =   5565
         Visible         =   0   'False
         Width           =   1245
         Begin VB.CommandButton cndBuscarCliente 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2250
            MaskColor       =   &H000000FF&
            Picture         =   "frmEntradaProductos.frx":1CEC
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Buscar Cliente"
            Top             =   300
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtCliRazSoc 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2685
            MaxLength       =   50
            TabIndex        =   5
            Tag             =   "Descripci�n"
            Top             =   300
            Width           =   5850
         End
         Begin VB.TextBox txtCodCliente 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1245
            MaxLength       =   40
            TabIndex        =   4
            Top             =   300
            Width           =   960
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   645
            TabIndex        =   47
            Top             =   360
            Width           =   525
         End
      End
      Begin VB.Frame FrameGeneral 
         Caption         =   "Datos Generales"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   90
         TabIndex        =   36
         Top             =   360
         Width           =   9105
         Begin VB.TextBox txtSigno 
            Alignment       =   2  'Center
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
            Left            =   8145
            Locked          =   -1  'True
            TabIndex        =   53
            Top             =   630
            Width           =   345
         End
         Begin VB.ComboBox cboEmpleado 
            Height          =   315
            Left            =   5355
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   285
            Width           =   3135
         End
         Begin VB.TextBox txtNumero 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
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
            Left            =   795
            MaxLength       =   8
            TabIndex        =   37
            Top             =   240
            Width           =   1140
         End
         Begin VB.ComboBox cboStock 
            Height          =   315
            Left            =   795
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   630
            Width           =   3300
         End
         Begin VB.ComboBox cboMovimiento 
            Height          =   315
            Left            =   5355
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   630
            Width           =   2775
         End
         Begin MSComCtl2.DTPicker Fecha 
            Height          =   315
            Left            =   2640
            TabIndex        =   0
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
            AutoSize        =   -1  'True
            Caption         =   "Vendedor:"
            Height          =   195
            Left            =   4410
            TabIndex        =   44
            Top             =   345
            Width           =   750
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "N�mero:"
            Height          =   195
            Left            =   150
            TabIndex        =   43
            Top             =   300
            Width           =   615
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Stock:"
            Height          =   195
            Left            =   150
            TabIndex        =   42
            Top             =   690
            Width           =   450
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha:"
            Height          =   195
            Index           =   2
            Left            =   2115
            TabIndex        =   41
            Top             =   300
            Width           =   480
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Movimiento:"
            Height          =   195
            Left            =   4410
            TabIndex        =   40
            Top             =   690
            Width           =   870
         End
         Begin VB.Label lblEstadoRecepcion 
            AutoSize        =   -1  'True
            Caption         =   "ESTADO RECEPCION"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   795
            TabIndex        =   39
            Top             =   975
            Width           =   1845
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   150
            TabIndex        =   38
            Top             =   990
            Width           =   555
         End
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones:"
         Height          =   195
         Left            =   120
         TabIndex        =   52
         Top             =   5400
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
         TabIndex        =   27
         Top             =   570
         Width           =   1065
      End
   End
   Begin VB.Label lblestado 
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
      Left            =   105
      TabIndex        =   30
      Top             =   6105
      Width           =   660
   End
End
Attribute VB_Name = "frmEntradaProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim VnumeroListado As Long

Private Sub cboMovimiento_Click()
    If cboMovimiento.ListIndex <> -1 Then
        sql = "SELECT ESP_SIGNO "
        sql = sql & " FROM ESTADO_PRODUCTO"
        sql = sql & " WHERE ESP_CODIGO=" & cboMovimiento.ItemData(cboMovimiento.ListIndex)
        Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec2.EOF = False Then
            txtSigno.Text = ChkNull(Rec2!ESP_SIGNO)
        End If
        Rec2.Close
    End If
End Sub

Private Sub cmdAsignar_Click()
    If TxtCODIGO.Text <> "" Then
        GrdModulos.HighLight = flexHighlightAlways
        If txtCantidad <> "" Then
            
            If TxtCODIGO.Text = 1 Or TxtCODIGO.Text = 3 Then 'BUSCO STOCKS DE TANQUES
                 For i = 1 To GrdModulos.Rows - 1
                    If GrdModulos.TextMatrix(i, 4) = CLng(TxtCodInt.Text) Then
                        If GrdModulos.TextMatrix(i, 4) = Right(Trim(IIf(optTanque1.Value = True, optTanque1.Caption, optTanque2.Caption)), 1) Then
                            MsgBox "El combustible para ese tanque ya fue ingresado", vbExclamation, TIT_MSGBOX
                            TxtCODIGO.SetFocus
                            Exit Sub
                        End If
                    End If
                Next
                 
                 
                 GrdModulos.AddItem Trim(TxtCODIGO.Text) & Chr(9) & Trim(TxtDescri.Text) & " - " & IIf(optTanque1.Value = True, optTanque1.Caption, optTanque2.Caption) _
                                & Chr(9) & Trim(txtCantidad.Text) & Chr(9) & "" & Chr(9) & Trim(TxtCodInt.Text) & Chr(9) & Right(Trim(IIf(optTanque1.Value = True, optTanque1.Caption, optTanque2.Caption)), 1)
                'txtIngNuevo_Click
                TxtCODIGO.Text = ""
                TxtCODIGO.SetFocus
                fraTanque.Visible = False
            Else
                If txtNumero.Text = "" Then
                    For i = 1 To GrdModulos.Rows - 1
                        If GrdModulos.TextMatrix(i, 0) = CLng(TxtCODIGO.Text) Then
                            GrdModulos.TextMatrix(i, 2) = CDbl(GrdModulos.TextMatrix(i, 2)) + CDbl(txtCantidad.Text)
                            TxtCODIGO.Text = ""
                            TxtCODIGO.SetFocus
                            Exit Sub
                        End If
                    Next
                Else
                    For i = 1 To GrdModulos.Rows - 1
                        If GrdModulos.TextMatrix(i, 4) = CLng(TxtCodInt.Text) Then
                            MsgBox "El producto ya fue ingresado", vbExclamation, TIT_MSGBOX
                            TxtCODIGO.SetFocus
                            Exit Sub
                        End If
                    Next
                End If
                GrdModulos.AddItem Trim(TxtCODIGO.Text) & Chr(9) & Trim(TxtDescri.Text) _
                                & Chr(9) & Trim(txtCantidad.Text) & Chr(9) & "" & Chr(9) & Trim(TxtCodInt.Text)
                 
                'txtIngNuevo_Click
                TxtCODIGO.Text = ""
                TxtCODIGO.SetFocus
            End If
        Else
            MsgBox "Debe Ingresar la cantidad", vbExclamation, TIT_MSGBOX
            txtCantidad.SetFocus
            Exit Sub
        End If
     Else
        MsgBox "Debe seleccionar un Producto"
    End If
End Sub

Private Sub CmdBorrar_Click()
    If txtNumero.Text <> "" Then
        If GrdModulos.Rows <> 1 Then
            If MsgBox("�Seguro desea Anular el Movimineto de Producto Nro: " & XN(txtNumero.Text) & "? ", vbQuestion + vbYesNo + vbDefaultButton2, TIT_MSGBOX) = vbYes Then
                lblestado.Caption = "Anulando..."
                Screen.MousePointer = vbHourglass
                On Error GoTo HayError1
                DBConn.BeginTrans
                
                'ANULO LA ENTRADA
                sql = "UPDATE ENTRADA_PRODUCTO"
                sql = sql & " SET EST_CODIGO=2"
                sql = sql & " WHERE EPR_CODIGO=" & XN(txtNumero.Text)
                DBConn.Execute sql
                
                'ACTUALIZO EL DETALLE
                For i = 1 To GrdModulos.Rows - 1
                    sql = "UPDATE STOCK"
                    sql = sql & " SET DST_STKFIS = DST_STKFIS "
                    If Trim(txtSigno.Text) = "+" Then
                        sql = sql & " - "
                    Else
                        sql = sql & " + "
                    End If
                    sql = sql & XN(GrdModulos.TextMatrix(i, 2))
                    sql = sql & " WHERE STK_CODIGO = " & XN(cboStock.ItemData(cboStock.ListIndex))
                    sql = sql & " AND PTO_CODIGO = " & XN(GrdModulos.TextMatrix(i, 4))
                    DBConn.Execute sql
                Next
                DBConn.CommitTrans
            End If
            lblestado.Caption = ""
            Screen.MousePointer = vbNormal
            CmdNuevo_Click
        End If
    End If
  Exit Sub
HayError1:
    lblestado.Caption = ""
    Screen.MousePointer = vbNormal
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub CmdBuscAprox_Click()
    lblestado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
    
    Set rec = New ADODB.Recordset
    sql = "SELECT E.EPR_CODIGO, E.EPR_FECHA, V.VEN_NOMBRE"
    sql = sql & " FROM ENTRADA_PRODUCTO E, VENDEDOR V"
    sql = sql & " WHERE E.VEN_CODIGO = V.VEN_CODIGO"
    If cboEmpleado1.List(cboEmpleado1.ListIndex) <> "(Todos)" Then
        sql = sql & " AND V.VEN_CODIGO = " & XN(cboEmpleado1.ItemData(cboEmpleado1.ListIndex))
    End If
    If cboMovimiento1.List(cboMovimiento1.ListIndex) <> "(Todos)" Then
        sql = sql & " AND E.ESP_CODIGO=" & XN(cboMovimiento1.ItemData(cboMovimiento1.ListIndex))
    End If
    If FechaDesde.Value <> "" Then sql = sql & " AND E.EPR_FECHA>=" & XDQ(FechaDesde)
    If FechaHasta.Value <> "" Then sql = sql & " AND E.EPR_FECHA<=" & XDQ(FechaHasta)
    sql = sql & " ORDER BY E.EPR_CODIGO"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
      
    If rec.EOF = False Then
        GRDGrilla.Rows = 1
        Do While rec.EOF = False
            GRDGrilla.AddItem Format(rec!EPR_CODIGO, "00000000") & Chr(9) & rec!EPR_FECHA & Chr(9) & _
                              Trim(rec!VEN_NOMBRE)
            rec.MoveNext
        Loop
        GRDGrilla.Col = 0
        GRDGrilla.SetFocus
    Else
        lblestado.Caption = ""
        Screen.MousePointer = vbNormal
        GRDGrilla.Rows = 1
        MsgBox "No se encontraron datos...", vbExclamation, TIT_MSGBOX
    End If
    lblestado.Caption = ""
    Screen.MousePointer = vbNormal
    rec.Close
End Sub

Private Sub cmdBuscarProducto_Click()
    BuscarProducto "CODIGO"
    TxtCODIGO.SetFocus
End Sub

Private Sub cmdGrabar_Click()
    On Error GoTo HayError2
         
    If ValidarEntrada = False Then Exit Sub
           
        If MsgBox("�Confirma Movomineto de Mercader�a?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
        
        Screen.MousePointer = vbHourglass
        lblestado.Caption = "Guardando ..."
        'DBConn.BeginTrans
        
        sql = "SELECT EPR_FECHA FROM ENTRADA_PRODUCTO"
        sql = sql & " WHERE EPR_CODIGO = " & XN(txtNumero.Text)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
        If rec.EOF = True Then
           'INSERTO EN LA TABLA ENTRADA_PRODUCTO
           sql = "INSERT INTO ENTRADA_PRODUCTO(EPR_CODIGO,EPR_FECHA,VEN_CODIGO,"
           sql = sql & " STK_CODIGO,ESP_CODIGO,"
           sql = sql & " EST_CODIGO,EPR_OBSERVACIONES, EPR_HORA)"
           sql = sql & " VALUES ("
           sql = sql & XN(txtNumero) & ","
           sql = sql & XDQ(Fecha.Value) & ","
           sql = sql & XN(cboEmpleado.ItemData(cboEmpleado.ListIndex)) & ","
           sql = sql & XN(cboStock.ItemData(cboStock.ListIndex)) & ","
           sql = sql & XN(cboMovimiento.ItemData(cboMovimiento.ListIndex)) & ","
           'sql = sql & XN(txtCodCliente.Text) & "," 'SI DEVUELVE PRODUCTOS
           sql = sql & " 3," 'ESTADO DEFINITIVO
           sql = sql & XS(txtObservaciones.Text) & ","
           sql = sql & "#" & Format(Time, "hh:mm") & "#)"
           DBConn.Execute sql
           
           'INSERTO EN LA TABLA DETALLE_ENTRADA_PRODUCTO
           'INSERTO EN LA TABLA DETALLE_ENTRADA_DET_PRODUCTO
           For i = 1 To GrdModulos.Rows - 1
               If GrdModulos.TextMatrix(i, 4) = 1 Or GrdModulos.TextMatrix(i, 4) = 3 Then
                    sql = "INSERT INTO DETALLE_ENTRADA_DET_PRODUCTO(EPR_CODIGO,PTO_CODIGO,DPT_CODIGO,DPT_DETALLE,DEP_CANTIDAD)"
                    sql = sql & " VALUES ("
                    sql = sql & XN(txtNumero.Text) & ","
                    sql = sql & XN(GrdModulos.TextMatrix(i, 4)) & ","
                    sql = sql & XN(GrdModulos.TextMatrix(i, 5)) & ","
                    sql = sql & XS(GrdModulos.TextMatrix(i, 1)) & ","
                    sql = sql & XN(GrdModulos.TextMatrix(i, 2)) & " )"
                    DBConn.Execute sql
               Else
                    sql = "INSERT INTO DETALLE_ENTRADA_PRODUCTO(EPR_CODIGO,PTO_CODIGO,DEP_CANTIDAD)"
                    sql = sql & " VALUES ("
                    sql = sql & XN(txtNumero.Text) & ","
                    sql = sql & XN(GrdModulos.TextMatrix(i, 4)) & ","
                    sql = sql & XN(GrdModulos.TextMatrix(i, 2)) & " )"
                    DBConn.Execute sql
               End If
           Next
    
            'ACTUALIZO DETALLE_STOCK
            For i = 1 To GrdModulos.Rows - 1
                If GrdModulos.TextMatrix(i, 4) = 1 Or GrdModulos.TextMatrix(i, 4) = 3 Then
                    sql = "UPDATE PRODUCTO_DETALLE"
                    sql = sql & " SET PDT_CANTIDAD= PDT_CANTIDAD " & Trim(txtSigno.Text) & XN(GrdModulos.TextMatrix(i, 2))
                    sql = sql & " WHERE PDT_CODIGO=" & XN(GrdModulos.TextMatrix(i, 5))
                    sql = sql & " AND PTO_CODIGO=" & XN(GrdModulos.TextMatrix(i, 4))
                    DBConn.Execute sql
                Else
                    sql = "UPDATE STOCK"
                    sql = sql & " SET DST_STKFIS = DST_STKFIS  " & Trim(txtSigno.Text) & XN(GrdModulos.TextMatrix(i, 2))
                    sql = sql & " WHERE STK_CODIGO= " & XN(cboStock.ItemData(cboStock.ListIndex))
                    sql = sql & " AND PTO_CODIGO =" & XN(GrdModulos.TextMatrix(i, 4))
                    DBConn.Execute sql
                End If
            Next
            
            'ACTUALIZO LA TABLA PARAMENTROS
            sql = "UPDATE PARAMETROS SET RECEPCION_MERCADERIA=" & XN(txtNumero.Text)
            DBConn.Execute sql
        Else
            MsgBox "La Recepci�n de Mercader�a ya fue registrada", vbCritical, TIT_MSGBOX
        End If
        rec.Close
        Screen.MousePointer = vbNormal
        lblestado.Caption = ""
        'DBConn.CommitTrans
        CmdNuevo_Click
    Exit Sub
         
HayError2:
         lblestado.Caption = ""
         'DBConn.RollbackTrans
         If rec.State = 1 Then rec.Close
         Screen.MousePointer = vbNormal
         MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Function ValidarEntrada()
    If cboEmpleado.ListIndex = -1 Then
        MsgBox "No ha ingresado el Encargado de Dep�sito", vbExclamation, TIT_MSGBOX
        cboEmpleado.SetFocus
        ValidarEntrada = False
        Exit Function
    End If
    If Fecha.Value = "" Then
        MsgBox "No ha ingresado la Fecha de Entrada de Productos", vbExclamation, TIT_MSGBOX
        Fecha.SetFocus
        ValidarEntrada = False
        Exit Function
    End If
    If GrdModulos.Rows = 1 Then
        MsgBox "Debe haber ingresar al menos un producto en la Grilla ", vbExclamation, TIT_MSGBOX
        cmdAsignar.SetFocus
        ValidarEntrada = False
        Exit Function
    End If
    ValidarEntrada = True
End Function

Private Sub CmdNuevo_Click()
    txtNumero.Text = ""
    txtObservaciones.Text = ""
    cboEmpleado.ListIndex = 0
    cboMovimiento.ListIndex = 0
    Fecha.Value = Date
    TxtCODIGO.Text = ""
    txtCodCliente.Text = ""
    GrdModulos.Rows = 1
    GrdModulos.HighLight = flexHighlightNever
    Call BuscoEstado(1, lblEstadoRecepcion)
    tabDatos.Tab = 0
    BuscoNumeroRecepcion
    cmdBorrar.Enabled = False
    cmdGrabar.Enabled = True
    FrameGeneral.Enabled = True
    FrameProducto.Enabled = True
    cboStock.SetFocus
    fraTanque.Visible = False
End Sub

Private Sub cmdQuitar_Click()
    If GrdModulos.Rows <> 1 Then
        If MsgBox("�Seguro desea Eliminar el Producto: " & Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 1)) & "? ", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
            lblestado.Caption = "Borrando..."
            Screen.MousePointer = vbHourglass
            If GrdModulos.Rows = 2 Then
                GrdModulos.HighLight = flexHighlightNever
                GrdModulos.Rows = 1
                TxtCODIGO.SetFocus
            Else
                GrdModulos.RemoveItem (GrdModulos.RowSel)
                TxtCODIGO.SetFocus
            End If
            lblestado.Caption = ""
            Screen.MousePointer = vbNormal
        End If
    End If
End Sub

Private Sub CmdSalir_Click()
    If MsgBox("Seguro que desea Salir", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        Set frmEntradaProductos = Nothing
        Unload Me
    End If
End Sub

Private Sub cndBuscarCliente_Click()
    frmBuscar.TipoBusqueda = 1
    frmBuscar.TxtDescriB = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 0
        txtCodCliente.Text = frmBuscar.grdBuscar.Text
        txtCodCliente_LostFocus
        txtCliRazSoc.SetFocus
    Else
        txtCodCliente.SetFocus
    End If
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
    'CARGO COMBO EMPLEADO
    cargocboEmpl
    'CARGO COMBO STOCK
    CargocboStock
    'CARGO COMBO Movimiento
    CargoComboEstadoProducto
    tabDatos.Tab = 0
    cmdAsignar.Enabled = False
    cmdBorrar.Enabled = False
    GrdModulos.HighLight = flexHighlightNever
    'BUSCO NUMERO DE RECEPCION DE MERCADERIA
    BuscoNumeroRecepcion
    Call BuscoEstado(1, lblEstadoRecepcion)
    Fecha.Value = Date
        
    'llenar tablka stock
    
'    sql = "select * from producto"
'    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    If rec.EOF = False Then
'        Do While rec.EOF = False
'            cSQL = "INSERT INTO STOCK"
'            cSQL = cSQL & "  (STK_CODIGO, PTO_CODIGO, DST_STKFIS) "
'            cSQL = cSQL & "VALUES (" & XN("4") & " , " & XN(rec!pto_codigo) & " ,10)"
'            DBConn.Execute cSQL
'            rec.MoveNext
'        Loop
'    End If
'    rec.Close
    
    
    
End Sub

Private Sub BuscoNumeroRecepcion()
    sql = "SELECT (RECEPCION_MERCADERIA + 1) AS NUMERO_REP FROM PARAMETROS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        txtNumero.Text = Format(rec!NUMERO_REP, "00000000")
    End If
    rec.Close
End Sub

Private Sub CargoComboEstadoProducto()
    sql = "SELECT ESP_DESCRI,ESP_CODIGO "
    sql = sql & " FROM ESTADO_PRODUCTO"
    sql = sql & " ORDER BY ESP_CODIGO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboMovimiento1.AddItem "(Todos)"
        Do While rec.EOF = False
            cboMovimiento.AddItem rec!ESP_DESCRI
            cboMovimiento.ItemData(cboMovimiento.NewIndex) = rec!ESP_CODIGO
            cboMovimiento1.AddItem rec!ESP_DESCRI
            cboMovimiento1.ItemData(cboMovimiento1.NewIndex) = rec!ESP_CODIGO
            rec.MoveNext
        Loop
        cboMovimiento.ListIndex = 0
        cboMovimiento1.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub preparogrilla()
    'GRILLA DONDE SE CRAGAN LOS PRODUCTOS
    GrdModulos.FormatString = "^C�digo|<Producto|^Cantidad|marca|CODINT|dpt_codigo"
    GrdModulos.ColWidth(0) = 1200 'CODIGO PRODUCTO
    GrdModulos.ColWidth(1) = 6100 'PRODUCTO
    GrdModulos.ColWidth(2) = 1100 'CANTIDAD
    GrdModulos.ColWidth(3) = 0    'marca para saber cunado actualizo el stock
    GrdModulos.ColWidth(4) = 0    'CODINT
    GrdModulos.ColWidth(5) = 0    'dpt_codigo
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
    GRDGrilla.FormatString = "^Numero|^Fecha|<Vendedor"
    GRDGrilla.ColWidth(0) = 1200 'NUMERO
    GRDGrilla.ColWidth(1) = 1300 'FECHA
    GRDGrilla.ColWidth(2) = 5000 'EMPLEADO
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
End Sub

Private Sub cargocboEmpl()
    sql = "SELECT VEN_CODIGO, VEN_NOMBRE"
    sql = sql & " FROM VENDEDOR "
    sql = sql & " WHERE VEN_CODIGO = 11" 'JOSE INOCENZI
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboEmpleado1.AddItem "(Todos)"
        Do While rec.EOF = False
            cboEmpleado.AddItem rec!VEN_NOMBRE
            cboEmpleado.ItemData(cboEmpleado.NewIndex) = rec!VEN_CODIGO
            cboEmpleado1.AddItem rec!VEN_NOMBRE
            cboEmpleado1.ItemData(cboEmpleado1.NewIndex) = rec!VEN_CODIGO
            rec.MoveNext
        Loop
        cboEmpleado.ListIndex = 0
        cboEmpleado1.ListIndex = 0
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
    If GRDGrilla.Rows > 1 Then
        CmdNuevo_Click
        txtNumero.Text = GRDGrilla.TextMatrix(GRDGrilla.RowSel, 0)
        Fecha.Value = GRDGrilla.TextMatrix(GRDGrilla.RowSel, 1)
        txtNumero_LostFocus
        tabDatos.Tab = 0
    End If
End Sub

Private Sub grdGrilla_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyReturn Then GRDGrilla_DblClick
End Sub

Private Sub tabDatos_Click(PreviousTab As Integer)
    If tabDatos.Tab = 1 Then
      cmdGrabar.Enabled = False
      cmdBorrar.Enabled = False
      LimpiarBusqueda
      If Me.Visible = True Then cboEmpleado1.SetFocus
    Else
      cmdGrabar.Enabled = True
      cmdBorrar.Enabled = True
    End If
End Sub

Private Sub LimpiarBusqueda()
    cboEmpleado1.ListIndex = 0
    cboMovimiento1.ListIndex = 0
    FechaDesde.Value = ""
    FechaHasta.Value = ""
    frameVer.Enabled = False
    GRDGrilla.Rows = 1
End Sub

Private Sub txtCantidad_GotFocus()
    SelecTexto txtCantidad
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCliRazSoc_Change()
    If txtCliRazSoc.Text = "" Then
        txtCodCliente.Text = ""
    End If
End Sub

Private Sub txtCliRazSoc_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtCliRazSoc_LostFocus()
    If txtCodCliente.Text = "" And txtCliRazSoc.Text <> "" Then
        rec.Open BuscoCliente(txtCliRazSoc), DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If rec.RecordCount > 1 Then
                frmBuscar.TipoBusqueda = 1
                frmBuscar.TxtDescriB.Text = txtCliRazSoc.Text
                frmBuscar.Show vbModal
                If frmBuscar.grdBuscar.Text <> "" Then
                    frmBuscar.grdBuscar.Col = 0
                    txtCodCliente.Text = frmBuscar.grdBuscar.Text
                    frmBuscar.grdBuscar.Col = 1
                    txtCliRazSoc.Text = frmBuscar.grdBuscar.Text
                    rec.Close
                    txtCodCliente_LostFocus
                    TxtCODIGO.SetFocus
                Else
                    txtCodCliente.SetFocus
                End If
            Else
                txtCodCliente.Text = rec!CLI_CODIGO
                txtCliRazSoc.Text = rec!CLI_RAZSOC
                rec.Close
            End If
        Else
            rec.Close
            MsgBox "No se encontro el Cliente", vbExclamation, TIT_MSGBOX
            txtCodCliente.SetFocus
        End If
    ElseIf txtCodCliente.Text = "" And txtCliRazSoc.Text = "" Then
        MsgBox "Debe elegir un cliente", vbExclamation, TIT_MSGBOX
        txtCodCliente.SetFocus
    End If
End Sub

Private Sub txtCodCliente_Change()
    If txtCodCliente.Text = "" Then
        txtCliRazSoc.Text = ""
    End If
End Sub

Private Sub txtCodCliente_GotFocus()
    SelecTexto txtCodCliente
End Sub

Private Sub txtCodCliente_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCodCliente_LostFocus()
    If txtCodCliente.Text <> "" Then
        rec.Open BuscoCliente(txtCodCliente), DBConn, adOpenStatic, adLockOptimistic
        
        If rec.EOF = False Then
            txtCliRazSoc.Text = rec!CLI_RAZSOC
        Else
            MsgBox "El Cliente no existe", vbExclamation, TIT_MSGBOX
            txtCodCliente.SetFocus
        End If
        rec.Close
    End If
End Sub

Private Function BuscoCliente(Codigo As String) As String
        sql = "SELECT C.CLI_CODIGO, C.CLI_RAZSOC"
        sql = sql & " FROM CLIENTE C"
        sql = sql & " WHERE"
        If txtCodCliente.Text <> "" Then
            sql = sql & " C.CLI_CODIGO=" & XN(Codigo)
        Else
            sql = sql & " C.CLI_RAZSOC LIKE '" & Trim(Codigo) & "%'"
        End If
        BuscoCliente = sql
End Function

Private Sub TxtCodigo_Change()
    If TxtCODIGO.Text = "" Then
        TxtCODIGO.Text = ""
        TxtDescri.Text = ""
        txtCantidad.Text = ""
        TxtCodInt.Text = ""
        cmdAsignar.Enabled = False
    Else
        cmdAsignar.Enabled = True
    End If
End Sub

Private Sub TxtCodigo_GotFocus()
    SelecTexto TxtCODIGO
End Sub

Private Sub txtcodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarProducto "CODIGO"
        TxtCODIGO.SetFocus
    End If
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub TxtCodigo_LostFocus()
    If TxtCODIGO.Text <> "" Then
        Set rec = New ADODB.Recordset
        sql = " SELECT P.PTO_DESCRI, P.PTO_CODIGO"
        sql = sql & " FROM PRODUCTO P"
        sql = sql & " WHERE"
        If IsNumeric(TxtCODIGO.Text) Then
            sql = sql & " P.PTO_CODIGO =" & XN(TxtCODIGO.Text) & " OR P.PTO_CODBARRAS=" & XS(TxtCODIGO.Text)
        Else
            sql = sql & " P.PTO_CODBARRAS=" & XS(TxtCODIGO.Text)
        End If
        sql = sql & " ORDER BY P.PTO_CODIGO"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            TxtDescri.Text = Trim(rec!PTO_DESCRI)
            TxtCodInt.Text = rec!PTO_CODIGO
        Else
            MsgBox "El C�digo no existe", vbExclamation, TIT_MSGBOX
            TxtCODIGO.SetFocus
        End If
        rec.Close
        If TxtCODIGO.Text = 1 Or TxtCODIGO.Text = 3 Then 'BUSCO STOCKS DE TANQUES
            fraTanque.Visible = True
            If TxtCODIGO.Text = 3 Then 'GASOIL
                optTanque1.Caption = "Tanque 3"
                optTanque2.Caption = "Tanque 4"
            Else
                optTanque1.Caption = "Tanque 1"
                optTanque2.Caption = "Tanque 2"
            End If
        End If
        
    End If
End Sub

Private Sub CargocboStock()
    sql = "SELECT SUC_CODIGO, SUC_DESCRI "
    sql = sql & " FROM SUCURSAL R "
    sql = sql & " WHERE SUC_CODIGO = " & XN(Sucursal)
    'sql = sql & " ORDER BY S.STK_CODIGO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboStock.AddItem rec!SUC_DESCRI
            cboStock.ItemData(cboStock.NewIndex) = rec!SUC_CODIGO
            rec.MoveNext
        Loop
        cboStock.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub txtdescri_Change()
    If TxtDescri.Text = "" Then
        TxtCODIGO.Text = ""
    End If
End Sub

Private Sub txtdescri_GotFocus()
    SelecTexto TxtDescri
End Sub

Private Sub txtdescri_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarProducto "CODIGO"
        TxtDescri.SetFocus
    End If
End Sub

Private Sub txtdescri_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtDescri_LostFocus()
   If TxtCODIGO.Text = "" And TxtDescri.Text <> "" Then
        Set Rec1 = New ADODB.Recordset
        Screen.MousePointer = vbHourglass
        sql = "SELECT PTO_CODIGO, PTO_DESCRI, PTO_CODBARRAS"
        sql = sql & " FROM PRODUCTO"
        sql = sql & " WHERE PTO_DESCRI LIKE '" & TxtDescri.Text & "%'"
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            If Rec1.RecordCount > 1 Then
                'grdGrilla.SetFocus
                BuscarProducto "CADENA", Trim(TxtDescri.Text)
                TxtDescri.SetFocus
            Else
                TxtCODIGO.Text = Trim(ChkNull(Rec1!PTO_CODBARRAS))
                TxtDescri.Text = Trim(Rec1!PTO_DESCRI)
                TxtCodInt.Text = Trim(Rec1!PTO_CODIGO)
            End If
        Else
                MsgBox "No se ha encontrado el Producto", vbExclamation, TIT_MSGBOX
                TxtDescri.Text = ""
        End If
        Rec1.Close
        Screen.MousePointer = vbNormal
    ElseIf TxtCODIGO.Text = "" And TxtDescri.Text = "" Then
        cmdAsignar.Enabled = False
    End If
End Sub

Private Sub txtNumero_GotFocus()
    SelecTexto txtNumero
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNumero_LostFocus()
    If txtNumero.Text <> "" Then
        Set Rec1 = New ADODB.Recordset
        sql = "SELECT * FROM ENTRADA_PRODUCTO"
        sql = sql & " WHERE EPR_CODIGO=" & XN(txtNumero)
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            Fecha.Value = Rec1!EPR_FECHA
            Call BuscaCodigoProxItemData(Rec1!VEN_CODIGO, cboEmpleado)
            Call BuscaCodigoProxItemData(Rec1!STK_CODIGO, cboStock)
            Call BuscaCodigoProxItemData(Rec1!ESP_CODIGO, cboMovimiento)
'            If Not IsNull(Rec1!CLI_CODIGO) Then
'                txtCodCliente.Text = Rec1!CLI_CODIGO
'                txtCodCliente_LostFocus
'            Else
'                txtCodCliente.Text = ""
'            End If
            CargoGrilla (txtNumero)
            Call BuscoEstado(CInt(Rec1!EST_CODIGO), lblEstadoRecepcion)
            txtObservaciones.Text = ChkNull(Rec1!EPR_OBSERVACIONES)
            If Rec1!EST_CODIGO = 2 Then
               cmdBorrar.Enabled = False
            Else
               cmdBorrar.Enabled = True
            End If
            cmdGrabar.Enabled = False
            FrameGeneral.Enabled = False
            FrameProducto.Enabled = False
        Else
            MsgBox "El Movimiento no existe", vbExclamation, TIT_MSGBOX
            CmdNuevo_Click
            cboStock.SetFocus
        End If
        Rec1.Close
    End If
End Sub

Private Sub CargoGrilla(Campo As Integer)
    Dim Rec2 As ADODB.Recordset
    Set Rec2 = New ADODB.Recordset
    ' busco en DETALLE DE entrada de producto
    Screen.MousePointer = vbHourglass
    sql = "SELECT DISTINCT  P.PTO_DESCRI, P.PTO_CODBARRAS,"
    sql = sql & " D.DEP_CANTIDAD, E.EPR_CODIGO, E.EPR_FECHA,P.PTO_CODIGO"
    sql = sql & " FROM ENTRADA_PRODUCTO E, PRODUCTO P, DETALLE_ENTRADA_PRODUCTO D"
    sql = sql & " WHERE P.PTO_CODIGO = D.PTO_CODIGO AND D.EPR_CODIGO = E.EPR_CODIGO"
    sql = sql & " AND E.EPR_CODIGO = " & Campo & " ORDER BY E.EPR_CODIGO"
        
    lblestado.Caption = "Buscando..."
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        GrdModulos.Rows = 1
        GrdModulos.HighLight = flexHighlightAlways
        Do While Not rec.EOF
           GrdModulos.AddItem IIf(IsNull(rec!PTO_CODBARRAS), rec!PTO_CODIGO, rec!PTO_CODBARRAS) & Chr(9) & Trim(rec!PTO_DESCRI) _
                              & Chr(9) & rec!DEP_CANTIDAD & Chr(9) & "X" & Chr(9) & rec!PTO_CODIGO
    
           rec.MoveNext
        Loop
        rec.MoveFirst
    Else
        'busco en detalle de entrada de detaleproducto - ' PARA COMBUSTIBLES !
        sql = "SELECT DISTINCT  D.DPT_DETALLE, D.PTO_CODIGO,D.DPT_CODIGO,"
        sql = sql & " D.DEP_CANTIDAD, E.EPR_CODIGO, E.EPR_FECHA"
        sql = sql & " FROM ENTRADA_PRODUCTO E, PRODUCTO P, DETALLE_ENTRADA_DET_PRODUCTO D"
        sql = sql & " WHERE P.PTO_CODIGO = D.PTO_CODIGO AND D.EPR_CODIGO = E.EPR_CODIGO"
        sql = sql & " AND E.EPR_CODIGO = " & Campo & " ORDER BY E.EPR_CODIGO, D.PTO_CODIGO,D.DPT_CODIGO"
            
        lblestado.Caption = "Buscando..."
        Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec2.EOF = False Then
            GrdModulos.Rows = 1
            GrdModulos.HighLight = flexHighlightAlways
            Do While Not Rec2.EOF
               GrdModulos.AddItem Rec2!PTO_CODIGO & Chr(9) & Trim(Rec2!DPT_DETALLE) _
                                  & Chr(9) & Rec2!DEP_CANTIDAD & Chr(9) & "" & Chr(9) & Rec2!PTO_CODIGO & Chr(9) & Rec2!DPT_CODIGO
        
               Rec2.MoveNext
            Loop
        End If
        
    End If
    If GrdModulos.Rows = 1 Then
        lblestado.Caption = ""
        MsgBox "No hay coincidencias en la busqueda.", vbOKOnly + vbCritical, TIT_MSGBOX
        Me.txtNumero.SetFocus
    End If
    rec.Close
    Rec2.Close
    Screen.MousePointer = vbNormal
    lblestado.Caption = ""
    
End Sub

Private Sub txtObservaciones_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Public Sub BuscarProducto(mQuien As String, Optional mCadena As String)
    Dim cSQL As String
    Dim hSQL As String
    Dim B As CBusqueda
    Dim i, posicion As Integer
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
        
        hSQL = "Descripci�n, C�digo"
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
                TxtCODIGO.Text = .ResultFields(2)
                TxtCodigo_LostFocus
        End If
    End With
    Set B = Nothing
End Sub
