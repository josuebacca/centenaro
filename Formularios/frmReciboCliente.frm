VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmReciboCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recibo de Cliente"
   ClientHeight    =   6756
   ClientLeft      =   0
   ClientTop       =   756
   ClientWidth     =   12048
   ControlBox      =   0   'False
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
   ScaleHeight     =   6756
   ScaleWidth      =   12048
   Begin Crystal.CrystalReport Rep 
      Left            =   1965
      Top             =   6645
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   450
      Left            =   8070
      TabIndex        =   8
      Top             =   6270
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Cancelar"
      Height          =   450
      Left            =   9870
      TabIndex        =   9
      Top             =   6270
      Width           =   870
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   450
      Left            =   8970
      TabIndex        =   7
      Top             =   6270
      Width           =   870
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   10755
      TabIndex        =   10
      Top             =   6270
      Width           =   870
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   6195
      Left            =   15
      TabIndex        =   11
      Top             =   45
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   10922
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   512
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
      TabPicture(0)   =   "frmReciboCliente.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label10"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameRecibo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FrameRemito"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "tabValores"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "tabComprobantes"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtObservaciones"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "frmReciboCliente.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GrdModulos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "frameBuscar"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.TextBox txtObservaciones 
         BackColor       =   &H00C0FFFF&
         Height          =   465
         Left            =   1230
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   107
         Top             =   5640
         Width           =   10530
      End
      Begin TabDlg.SSTab tabComprobantes 
         Height          =   3840
         Left            =   105
         TabIndex        =   32
         Top             =   1695
         Width           =   5940
         _ExtentX        =   10478
         _ExtentY        =   6773
         _Version        =   393216
         Style           =   1
         Tabs            =   1
         TabsPerRow      =   2
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "C&omprobantes Pendientes"
         TabPicture(0)   =   "frmReciboCliente.frx":0038
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame5"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         Begin VB.Frame Frame5 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3480
            Left            =   75
            TabIndex        =   35
            Top             =   315
            Width           =   5790
            Begin VB.TextBox txtSaldoActual 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000A&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   375
               Left            =   960
               TabIndex        =   109
               Top             =   3000
               Visible         =   0   'False
               Width           =   1275
            End
            Begin VB.TextBox txtpagTar 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FF0000&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   420
               Left            =   120
               TabIndex        =   105
               Top             =   3015
               Visible         =   0   'False
               Width           =   810
            End
            Begin VB.TextBox txtSaldoEftTar 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   104
               Top             =   2640
               Visible         =   0   'False
               Width           =   795
            End
            Begin VB.TextBox txtSaldo 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   3855
               TabIndex        =   37
               Top             =   2625
               Width           =   1275
            End
            Begin VB.TextBox txtImporteApagar 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FF0000&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   420
               Left            =   3855
               TabIndex        =   36
               Top             =   3000
               Width           =   1290
            End
            Begin MSFlexGridLib.MSFlexGrid GrillaAplicar 
               Height          =   2415
               Left            =   105
               TabIndex        =   38
               Top             =   195
               Width           =   5445
               _ExtentX        =   9610
               _ExtentY        =   4276
               _Version        =   393216
               Cols            =   7
               FixedCols       =   0
               RowHeightMin    =   300
               BackColorSel    =   16761024
               AllowBigSelection=   -1  'True
               FocusRect       =   0
               HighLight       =   2
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
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Saldo Actual:"
               Height          =   195
               Left            =   960
               TabIndex        =   110
               Top             =   2760
               Visible         =   0   'False
               Width           =   945
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Saldo:"
               Height          =   195
               Left            =   3225
               TabIndex        =   40
               Top             =   2670
               Width           =   450
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Importe a Pagar:"
               Height          =   195
               Left            =   2445
               TabIndex        =   39
               Top             =   2985
               Width           =   1230
            End
         End
      End
      Begin TabDlg.SSTab tabValores 
         Height          =   3840
         Left            =   6090
         TabIndex        =   6
         Top             =   1695
         Width           =   5805
         _ExtentX        =   10224
         _ExtentY        =   6773
         _Version        =   393216
         Style           =   1
         Tabs            =   5
         Tab             =   4
         TabsPerRow      =   5
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "&Valores"
         TabPicture(0)   =   "frmReciboCliente.frx":0054
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame2"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Moneda"
         TabPicture(1)   =   "frmReciboCliente.frx":0070
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame4"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Valores a Cuenta"
         TabPicture(2)   =   "frmReciboCliente.frx":008C
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame6"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "T de Credito"
         TabPicture(3)   =   "frmReciboCliente.frx":00A8
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "fraTarjeta"
         Tab(3).Control(1)=   "fraPagos"
         Tab(3).ControlCount=   2
         TabCaption(4)   =   "Cheques"
         TabPicture(4)   =   "frmReciboCliente.frx":00C4
         Tab(4).ControlEnabled=   -1  'True
         Tab(4).Control(0)=   "Frame1"
         Tab(4).Control(0).Enabled=   0   'False
         Tab(4).ControlCount=   1
         Begin VB.Frame Frame1 
            Caption         =   "Cheques"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3435
            Left            =   120
            TabIndex        =   112
            Top             =   360
            Width           =   5535
            Begin VB.TextBox TxtCheImport 
               Height          =   330
               Left            =   3780
               TabIndex        =   123
               Top             =   315
               Width           =   900
            End
            Begin VB.TextBox txtTotalCheques 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
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
               Left            =   4305
               TabIndex        =   122
               Top             =   3045
               Width           =   1035
            End
            Begin VB.CommandButton cmdAgregarCheque 
               Caption         =   "Agregar"
               Height          =   345
               Left            =   4755
               TabIndex        =   81
               Top             =   1425
               Width           =   720
            End
            Begin VB.CommandButton cmdNuevoCheque 
               Height          =   315
               Left            =   2610
               MaskColor       =   &H000000FF&
               Picture         =   "frmReciboCliente.frx":00E0
               Style           =   1  'Graphical
               TabIndex        =   121
               ToolTipText     =   "Cargar Cheques"
               Top             =   330
               UseMaskColor    =   -1  'True
               Width           =   405
            End
            Begin VB.TextBox TxtCheNumero 
               Height          =   315
               Left            =   1110
               MaxLength       =   10
               TabIndex        =   75
               Top             =   330
               Width           =   1380
            End
            Begin VB.Frame frameBanco 
               Caption         =   "Banco"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   90
               TabIndex        =   114
               Top             =   660
               Width           =   4635
               Begin VB.CommandButton CmdBanco 
                  DisabledPicture =   "frmReciboCliente.frx":046A
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   4170
                  Picture         =   "frmReciboCliente.frx":0774
                  Style           =   1  'Graphical
                  TabIndex        =   80
                  Top             =   225
                  Width           =   375
               End
               Begin VB.TextBox TxtCodInt 
                  BackColor       =   &H80000018&
                  Height          =   300
                  Left            =   2670
                  TabIndex        =   116
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   420
               End
               Begin VB.TextBox TxtBanDescri 
                  BackColor       =   &H00C0C0C0&
                  Enabled         =   0   'False
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
                  Left            =   60
                  TabIndex        =   115
                  Top             =   615
                  Width           =   4500
               End
               Begin VB.TextBox TxtCODIGO 
                  Height          =   285
                  Left            =   3360
                  MaxLength       =   6
                  TabIndex        =   79
                  Top             =   255
                  Width           =   765
               End
               Begin VB.TextBox TxtLOCALIDAD 
                  Height          =   285
                  Left            =   1410
                  MaxLength       =   3
                  TabIndex        =   77
                  Top             =   240
                  Width           =   450
               End
               Begin VB.TextBox TxtBANCO 
                  Height          =   285
                  Left            =   525
                  MaxLength       =   3
                  TabIndex        =   76
                  Top             =   240
                  Width           =   450
               End
               Begin VB.TextBox TxtSUCURSAL 
                  Height          =   285
                  Left            =   2280
                  MaxLength       =   3
                  TabIndex        =   78
                  Top             =   255
                  Width           =   450
               End
               Begin VB.Label lbl 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Código:"
                  ForeColor       =   &H80000008&
                  Height          =   195
                  Index           =   0
                  Left            =   2790
                  TabIndex        =   120
                  Top             =   285
                  Width           =   540
               End
               Begin VB.Label lbl 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Suc:"
                  ForeColor       =   &H80000008&
                  Height          =   195
                  Index           =   5
                  Left            =   1935
                  TabIndex        =   119
                  Top             =   270
                  Width           =   330
               End
               Begin VB.Label lbl 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Bco:"
                  ForeColor       =   &H80000008&
                  Height          =   195
                  Index           =   10
                  Left            =   150
                  TabIndex        =   118
                  Top             =   270
                  Width           =   330
               End
               Begin VB.Label lbl 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Loc:"
                  ForeColor       =   &H80000008&
                  Height          =   195
                  Index           =   11
                  Left            =   1035
                  TabIndex        =   117
                  Top             =   270
                  Width           =   315
               End
            End
            Begin VB.CommandButton cmdAceptarCheques 
               Caption         =   "Aceptar"
               Height          =   360
               Left            =   585
               TabIndex        =   82
               Top             =   3015
               Width           =   960
            End
            Begin VB.CommandButton cmdCancelarCheques 
               Caption         =   "Cancelar"
               Height          =   360
               Left            =   1560
               TabIndex        =   83
               Top             =   3015
               Width           =   960
            End
            Begin VB.CommandButton cmdBuscaCheque 
               Height          =   315
               Left            =   3120
               MaskColor       =   &H000000FF&
               Picture         =   "frmReciboCliente.frx":08BE
               Style           =   1  'Graphical
               TabIndex        =   113
               ToolTipText     =   "Buscar Cheques"
               Top             =   330
               UseMaskColor    =   -1  'True
               Width           =   405
            End
            Begin MSFlexGridLib.MSFlexGrid GrillaCheques 
               Height          =   1170
               Left            =   75
               TabIndex        =   124
               Top             =   1815
               Width           =   5385
               _ExtentX        =   9504
               _ExtentY        =   2074
               _Version        =   393216
               Cols            =   9
               FixedCols       =   0
               BackColorSel    =   8388736
               AllowBigSelection=   -1  'True
               FocusRect       =   0
               HighLight       =   2
               SelectionMode   =   1
            End
            Begin MSComCtl2.DTPicker TxtCheFecVto 
               Height          =   315
               Left            =   3960
               TabIndex        =   125
               Top             =   480
               Visible         =   0   'False
               Width           =   1455
               _ExtentX        =   2561
               _ExtentY        =   550
               _Version        =   393216
               CheckBox        =   -1  'True
               DateIsNull      =   -1  'True
               Format          =   106299393
               CurrentDate     =   41098
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Total:"
               Height          =   195
               Left            =   3840
               TabIndex        =   127
               Top             =   3105
               Width           =   405
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Nro Cheque:"
               Height          =   195
               Index           =   7
               Left            =   90
               TabIndex        =   126
               Top             =   375
               Width           =   900
            End
         End
         Begin VB.Frame fraTarjeta 
            Height          =   2685
            Left            =   -69960
            TabIndex        =   87
            Top             =   660
            Visible         =   0   'False
            Width           =   4095
            Begin VB.CommandButton cmdAceptoTarjeta 
               Caption         =   "Aceptar"
               Height          =   375
               Left            =   1260
               TabIndex        =   96
               Top             =   2280
               Width           =   1425
            End
            Begin VB.TextBox txtLote 
               Height          =   315
               Left            =   1305
               TabIndex        =   92
               Top             =   1245
               Width           =   2505
            End
            Begin VB.TextBox txtCupon 
               Height          =   315
               Left            =   1305
               TabIndex        =   93
               Top             =   1605
               Width           =   2505
            End
            Begin VB.ComboBox cboPlan 
               Height          =   315
               ItemData        =   "frmReciboCliente.frx":0BC8
               Left            =   1305
               List            =   "frmReciboCliente.frx":0BCA
               Style           =   2  'Dropdown List
               TabIndex        =   91
               Top             =   885
               Width           =   2505
            End
            Begin VB.ComboBox cboTarjeta 
               Height          =   315
               ItemData        =   "frmReciboCliente.frx":0BCC
               Left            =   1305
               List            =   "frmReciboCliente.frx":0BCE
               Style           =   2  'Dropdown List
               TabIndex        =   90
               Top             =   495
               Width           =   2505
            End
            Begin VB.TextBox txtTar_Autorizacion 
               Height          =   315
               Left            =   1305
               MaxLength       =   30
               TabIndex        =   94
               Top             =   1965
               Width           =   2505
            End
            Begin VB.CommandButton cmdCerrarTarjeta 
               Caption         =   "Cerrar"
               Height          =   375
               Left            =   2730
               TabIndex        =   98
               Top             =   2280
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
               Left            =   3780
               TabIndex        =   89
               ToolTipText     =   "Alta de Tarjeta"
               Top             =   510
               Width           =   240
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
               Left            =   3780
               TabIndex        =   88
               ToolTipText     =   "Alta de Plan"
               Top             =   900
               Width           =   240
            End
            Begin VB.Label Label22 
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
               TabIndex        =   102
               Top             =   120
               Width           =   4005
            End
            Begin VB.Label Label21 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Lote:"
               Height          =   315
               Left            =   45
               TabIndex        =   101
               Top             =   1245
               Width           =   1215
            End
            Begin VB.Label Label20 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Cupón:"
               Height          =   315
               Left            =   45
               TabIndex        =   100
               Top             =   1605
               Width           =   1215
            End
            Begin VB.Label Label19 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Plan:"
               Height          =   315
               Left            =   45
               TabIndex        =   99
               Top             =   885
               Width           =   1215
            End
            Begin VB.Label Label17 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Tarjeta:"
               Height          =   315
               Left            =   45
               TabIndex        =   97
               Top             =   495
               Width           =   1215
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Autorización:"
               Height          =   315
               Left            =   45
               TabIndex        =   95
               Top             =   1965
               Width           =   1215
            End
         End
         Begin VB.Frame fraPagos 
            Height          =   3375
            Left            =   -74640
            TabIndex        =   65
            Top             =   375
            Width           =   5055
            Begin VB.TextBox txtSaldototal 
               Height          =   375
               Left            =   4800
               TabIndex        =   103
               Text            =   "0"
               Top             =   960
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.ComboBox cboFormaPago 
               Height          =   315
               ItemData        =   "frmReciboCliente.frx":0BD0
               Left            =   1470
               List            =   "frmReciboCliente.frx":0BD2
               Style           =   2  'Dropdown List
               TabIndex        =   69
               Top             =   840
               Width           =   3330
            End
            Begin VB.TextBox txtImportePago 
               Height          =   315
               Left            =   1470
               TabIndex        =   70
               Top             =   1215
               Width           =   1485
            End
            Begin VB.CommandButton cmdAceptarPagos 
               Caption         =   "Aceptar"
               Height          =   375
               Left            =   2160
               TabIndex        =   72
               Top             =   2895
               Width           =   1425
            End
            Begin VB.CommandButton cmdBorroFila 
               Caption         =   "Borrar Fila"
               Height          =   375
               Left            =   90
               TabIndex        =   84
               Top             =   2895
               Width           =   1095
            End
            Begin VB.Frame Frame3 
               Height          =   675
               Left            =   120
               TabIndex        =   67
               Top             =   110
               Width           =   4850
               Begin VB.TextBox txtTotalPagos 
                  Alignment       =   2  'Center
                  Enabled         =   0   'False
                  Height          =   375
                  Left            =   3120
                  TabIndex        =   68
                  Top             =   175
                  Width           =   1515
               End
               Begin VB.Label Label35 
                  Alignment       =   2  'Center
                  BackColor       =   &H000000FF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "T O T A L"
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
                  Left            =   90
                  TabIndex        =   73
                  Top             =   175
                  Width           =   3015
               End
            End
            Begin VB.TextBox txtGrabar 
               Height          =   285
               Left            =   3840
               TabIndex        =   66
               Top             =   1200
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.CommandButton cmdCerrarPagos 
               Caption         =   "Cancelar"
               Height          =   375
               Left            =   3630
               TabIndex        =   74
               Top             =   2895
               Width           =   1095
            End
            Begin MSFlexGridLib.MSFlexGrid grdPagos 
               Height          =   1365
               Left            =   120
               TabIndex        =   71
               Top             =   1545
               Width           =   4755
               _ExtentX        =   8382
               _ExtentY        =   2413
               _Version        =   393216
               Rows            =   1
               Cols            =   15
               FixedCols       =   0
               ForeColorSel    =   12632064
               ScrollTrack     =   -1  'True
               FocusRect       =   2
               HighLight       =   2
               SelectionMode   =   1
               FormatString    =   $"frmReciboCliente.frx":0BD4
            End
            Begin VB.Label Label36 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Forma Pago"
               Height          =   330
               Left            =   120
               TabIndex        =   86
               Top             =   840
               Width           =   1320
            End
            Begin VB.Label Label38 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Importe:"
               Height          =   330
               Left            =   120
               TabIndex        =   85
               Top             =   1215
               Width           =   1320
            End
         End
         Begin VB.Frame Frame4 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3480
            Left            =   -74880
            TabIndex        =   53
            Top             =   285
            Width           =   5535
            Begin VB.CommandButton cmdAgregarEfectivo 
               Caption         =   "Agregar"
               Height          =   345
               Left            =   2175
               TabIndex        =   59
               Top             =   705
               Width           =   885
            End
            Begin VB.TextBox txtEftImporte 
               Height          =   330
               Left            =   1125
               TabIndex        =   58
               Top             =   705
               Width           =   1005
            End
            Begin VB.ComboBox cboMoneda 
               Height          =   315
               Left            =   1125
               Style           =   2  'Dropdown List
               TabIndex        =   57
               Top             =   345
               Width           =   1950
            End
            Begin VB.TextBox txtTotalEfectivo 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FF0000&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   405
               Left            =   2745
               Locked          =   -1  'True
               TabIndex        =   56
               Top             =   2460
               Width           =   1290
            End
            Begin VB.CommandButton cmdAceptarMoneda 
               Caption         =   "Aceptar"
               Height          =   360
               Left            =   2115
               TabIndex        =   55
               Top             =   2925
               Width           =   960
            End
            Begin VB.CommandButton cmdCancelarMoneda 
               Caption         =   "Cancelar"
               Height          =   360
               Left            =   3090
               TabIndex        =   54
               Top             =   2925
               Width           =   960
            End
            Begin MSFlexGridLib.MSFlexGrid GrillaEfectivo 
               Height          =   1320
               Left            =   1110
               TabIndex        =   63
               Top             =   1110
               Width           =   2925
               _ExtentX        =   5144
               _ExtentY        =   2328
               _Version        =   393216
               Cols            =   3
               FixedCols       =   0
               RowHeightMin    =   300
               BackColorSel    =   16761024
               AllowBigSelection=   -1  'True
               FocusRect       =   0
               HighLight       =   2
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
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Importe:"
               Height          =   195
               Index           =   2
               Left            =   420
               TabIndex        =   62
               Top             =   765
               Width           =   630
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Moneda:"
               Height          =   195
               Left            =   420
               TabIndex        =   61
               Top             =   390
               Width           =   630
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "Total:"
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
               Left            =   2190
               TabIndex        =   60
               Top             =   2475
               Width           =   405
            End
         End
         Begin VB.Frame Frame6 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3480
            Left            =   -74925
            TabIndex        =   45
            Top             =   650
            Width           =   5415
            Begin VB.CommandButton cmdAgregarACta 
               Caption         =   "A&gregar"
               Height          =   420
               Left            =   3285
               TabIndex        =   49
               Top             =   2865
               Width           =   1065
            End
            Begin VB.TextBox txtSaldoACta 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1455
               Locked          =   -1  'True
               TabIndex        =   48
               Top             =   2535
               Width           =   1185
            End
            Begin VB.TextBox txtImporteACta 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1455
               TabIndex        =   47
               Top             =   2925
               Width           =   1185
            End
            Begin VB.CommandButton cmaAceptarACta 
               Caption         =   "A&ceptar"
               Height          =   420
               Left            =   4365
               TabIndex        =   46
               Top             =   2865
               Width           =   1065
            End
            Begin MSFlexGridLib.MSFlexGrid GrillaAFavor 
               Height          =   2175
               Left            =   30
               TabIndex        =   50
               Top             =   210
               Width           =   5355
               _ExtentX        =   9440
               _ExtentY        =   3831
               _Version        =   393216
               Cols            =   7
               FixedCols       =   0
               RowHeightMin    =   300
               BackColorSel    =   16761024
               AllowBigSelection=   -1  'True
               FocusRect       =   0
               HighLight       =   2
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
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               Caption         =   "Saldo:"
               Height          =   195
               Left            =   975
               TabIndex        =   52
               Top             =   2595
               Width           =   450
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               Caption         =   "Importe:"
               Height          =   195
               Left            =   795
               TabIndex        =   51
               Top             =   2970
               Width           =   630
            End
         End
         Begin VB.Frame Frame2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3480
            Left            =   -74900
            TabIndex        =   41
            Top             =   255
            Width           =   5550
            Begin VB.TextBox txtTotalValores 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FF0000&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   405
               Left            =   3945
               Locked          =   -1  'True
               TabIndex        =   42
               TabStop         =   0   'False
               Text            =   "0.00"
               Top             =   3000
               Width           =   1290
            End
            Begin MSFlexGridLib.MSFlexGrid grillaValores 
               Height          =   2430
               Left            =   90
               TabIndex        =   43
               Top             =   225
               Width           =   5430
               _ExtentX        =   9567
               _ExtentY        =   4276
               _Version        =   393216
               Cols            =   7
               FixedCols       =   0
               RowHeightMin    =   300
               BackColorSel    =   16761024
               FocusRect       =   0
               HighLight       =   0
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
            Begin VB.Label LblDineroaCta 
               AutoSize        =   -1  'True
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   285
               Left            =   165
               TabIndex        =   64
               Top             =   2670
               Width           =   75
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "Total:"
               Height          =   195
               Left            =   3420
               TabIndex        =   44
               Top             =   3000
               Width           =   420
            End
         End
      End
      Begin VB.Frame FrameRemito 
         Caption         =   "Cliente..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1305
         Left            =   4605
         TabIndex        =   18
         Top             =   360
         Width           =   7260
         Begin VB.TextBox txtDomici 
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
            Height          =   330
            Left            =   1350
            MaxLength       =   50
            TabIndex        =   33
            Top             =   840
            Width           =   5625
         End
         Begin VB.TextBox txtCliRazSoc 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2160
            MaxLength       =   50
            TabIndex        =   5
            Tag             =   "Descripción"
            Top             =   480
            Width           =   4815
         End
         Begin VB.TextBox txtCodCliente 
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
            Height          =   330
            Left            =   1350
            MaxLength       =   40
            TabIndex        =   4
            Top             =   480
            Width           =   780
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio:"
            Height          =   195
            Left            =   600
            TabIndex        =   34
            Top             =   870
            Width           =   675
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Recibimos de:"
            Height          =   195
            Left            =   285
            TabIndex        =   30
            Top             =   540
            Width           =   990
         End
      End
      Begin VB.Frame FrameRecibo 
         Caption         =   "Recibo..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1305
         Left            =   105
         TabIndex        =   24
         Top             =   360
         Width           =   4485
         Begin VB.TextBox txtNroSucursal 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1305
            MaxLength       =   4
            TabIndex        =   1
            Top             =   570
            Width           =   555
         End
         Begin VB.TextBox txtNroRecibo 
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
            Height          =   330
            Left            =   1890
            MaxLength       =   8
            TabIndex        =   2
            Top             =   570
            Width           =   1065
         End
         Begin VB.ComboBox cboRecibo 
            Height          =   315
            Left            =   1305
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   225
            Width           =   2400
         End
         Begin MSComCtl2.DTPicker FechaRecibo 
            Height          =   315
            Left            =   1320
            TabIndex        =   3
            Top             =   930
            Width           =   1215
            _ExtentX        =   2138
            _ExtentY        =   550
            _Version        =   393216
            DateIsNull      =   -1  'True
            Format          =   106299393
            CurrentDate     =   41176
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   720
            TabIndex        =   106
            Top             =   945
            Width           =   495
         End
         Begin VB.Label lblEstadoRecibo 
            AutoSize        =   -1  'True
            Caption         =   "EST. RECIBO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   3345
            TabIndex        =   28
            Top             =   1035
            Width           =   1005
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   2700
            TabIndex        =   27
            Top             =   1020
            Width           =   555
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   600
            TabIndex        =   26
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   855
            TabIndex        =   25
            Top             =   240
            Width           =   360
         End
      End
      Begin VB.Frame frameBuscar 
         Caption         =   "Buscar por..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1620
         Left            =   -74715
         TabIndex        =   19
         Top             =   480
         Width           =   11475
         Begin VB.TextBox txtCliente 
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
            Height          =   330
            Left            =   2805
            MaxLength       =   40
            TabIndex        =   12
            Top             =   300
            Width           =   750
         End
         Begin VB.TextBox txtDesCli 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3585
            MaxLength       =   50
            TabIndex        =   14
            Tag             =   "Descripción"
            Top             =   300
            Width           =   3990
         End
         Begin VB.CommandButton CmdBuscAprox 
            Caption         =   "Buscar"
            Height          =   450
            Left            =   8280
            MaskColor       =   &H80000006&
            Picture         =   "frmReciboCliente.frx":0BDA
            TabIndex        =   16
            ToolTipText     =   "Buscar "
            Top             =   915
            UseMaskColor    =   -1  'True
            Width           =   2085
         End
         Begin VB.ComboBox cboRecibo1 
            Height          =   315
            Left            =   2805
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   1095
            Width           =   2400
         End
         Begin MSComCtl2.DTPicker FechaHasta 
            Height          =   315
            Left            =   5415
            TabIndex        =   13
            Top             =   720
            Width           =   1455
            _ExtentX        =   2582
            _ExtentY        =   550
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   106299393
            CurrentDate     =   41098
         End
         Begin MSComCtl2.DTPicker FechaDesde 
            Height          =   315
            Left            =   2805
            TabIndex        =   111
            Top             =   720
            Width           =   1455
            _ExtentX        =   2561
            _ExtentY        =   550
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   106299393
            CurrentDate     =   41098
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
            Left            =   2145
            TabIndex        =   23
            Top             =   345
            Width           =   555
         End
         Begin VB.Label lblFechaDesde 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   1710
            TabIndex        =   22
            Top             =   810
            Width           =   990
         End
         Begin VB.Label lblFechaHasta 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   4380
            TabIndex        =   21
            Top             =   780
            Width           =   960
         End
         Begin VB.Label lbltipoFac 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Recibo:"
            Height          =   195
            Left            =   1815
            TabIndex        =   20
            Top             =   1125
            Width           =   885
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   3765
         Left            =   -74745
         TabIndex        =   17
         Top             =   2160
         Width           =   11520
         _ExtentX        =   20320
         _ExtentY        =   6646
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorSel    =   16761024
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         HighLight       =   2
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
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones:"
         Height          =   195
         Left            =   120
         TabIndex        =   108
         Top             =   5640
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
         TabIndex        =   29
         Top             =   570
         Width           =   1065
      End
   End
   Begin VB.Label lblEstado 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   135
      TabIndex        =   31
      Top             =   6705
      Width           =   660
   End
End
Attribute VB_Name = "frmReciboCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim TotFac As Double
Dim Estado As Integer
Dim mBorroTransfe As Boolean
Dim mImprimoRecibo As Boolean
 
Private Function SumaGrilla(Grilla As MSFlexGrid, COLUMNA As Integer) As String
    Dim Suma As Double
    Suma = 0
    For i = 1 To Grilla.Rows - 1
        Suma = Suma + CDbl(Chk0(Grilla.TextMatrix(i, COLUMNA)))
    Next
    SumaGrilla = VALIDO_IMPORTE(CStr(Suma))
End Function

Private Sub cboFormaPago_LostFocus()
    Dim mTotalPagos As Double
    Dim recargo As Double
    If Me.ActiveControl.Name = "grdPagos" Then
        Exit Sub
    End If
    If txtCodCliente.Text = "1" Then
        If cboFormaPago.ItemData(cboFormaPago.ListIndex) = 2 Then
            MsgBox "No Puede Seleccionar Cta CTe para este Cliente", vbCritical, TIT_MSGBOX
            cboFormaPago.ListIndex = 0
            cboFormaPago.SetFocus
            Exit Sub
        End If
    End If

    
    
    'fraTarjeta.Visible = False
     If Trim(cboFormaPago.Text) = "TARJETA DE CREDITO" Then
        cboPlan.Clear
        cbotarjeta.Clear
        cSQL = "SELECT TAR_CODIGO, TAR_DESCRI"
        cSQL = cSQL & " FROM TARJETA"
        cSQL = cSQL & " WHERE TTA_CODIGO=1" 'SOLO TARJETA DE CREDITO
        cSQL = cSQL & " ORDER BY TAR_DESCRI"
        rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
        If (rec.BOF And rec.EOF) = 0 Then
           Do While rec.EOF = False
              cbotarjeta.AddItem Trim(rec!TAR_DESCRI)
              cbotarjeta.ItemData(cbotarjeta.NewIndex) = rec!TAR_CODIGO
              rec.MoveNext
           Loop
           If cbotarjeta.ListCount > 0 Then cbotarjeta.ListIndex = 0
        End If
        rec.Close
                       
        'aplicar aumento de 10% O EL DEL CLIENTE SI TIENE
        sql = "SELECT CLI_PORC FROM CLIENTE WHERE CLI_CODIGO = " & XN(txtCodCliente)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If rec!CLI_PORC = 0 Or IsNull(rec!CLI_PORC) Then
                sql = "SELECT PORC FROM PARAMETROS"
                Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
                recargo = 0
                If Rec1.EOF = False Then
                    recargo = (CDbl(txtTotalPagos) * Chk0(Rec1!Porc)) / 100
                    txtTotalPagos.Text = CDbl(txtTotalPagos) + recargo
                    txtTotalPagos.Text = VALIDO_IMPORTE(txtTotalPagos.Text)
                    For i = 1 To grdPagos.Rows - 1
                        mTotalPagos = mTotalPagos + CDbl(grdPagos.TextMatrix(i, 1))
                    Next
                End If
                Rec1.Close
                If txtTotalValores.Text <> "0.00" Then
                    txtImporteApagar.Text = VALIDO_IMPORTE(CDbl(txtTotalValores.Text) + CDbl(txtTotalPagos.Text))
                Else
                    txtImporteApagar.Text = VALIDO_IMPORTE(CDbl(txtpagTar.Text) + recargo)
                End If
            Else
                txtTotalPagos.Text = CDbl(txtTotalPagos) + (CDbl(txtTotalPagos) * rec!CLI_PORC) / 100
                txtTotalPagos.Text = VALIDO_IMPORTE(txtTotalPagos.Text)
                For i = 1 To grdPagos.Rows - 1
                    mTotalPagos = mTotalPagos + CDbl(grdPagos.TextMatrix(i, 1))
                Next
                
                If txtTotalValores.Text <> "0.00" Then
                    txtImporteApagar.Text = VALIDO_IMPORTE(CDbl(txtTotalValores.Text) + CDbl(txtTotalPagos.Text))
                Else
                    txtImporteApagar.Text = VALIDO_IMPORTE(CDbl(txtpagTar.Text) + recargo)
                End If
            End If
        End If
        rec.Close
        
'        fraTarjeta.Top = 1000
'        fraTarjeta.Left = 1500
'        fraTarjeta.Visible = True
        'frmDatosTarjeta.cboTarjeta.SetFocus
        frmDatosTarjeta.cboPlan.Enabled = True
        frmDatosTarjeta.txtLote.Enabled = True
        frmDatosTarjeta.txtCupon.Enabled = True
        frmDatosTarjeta.txtTar_Autorizacion.Enabled = True
        frmDatosTarjeta.Show vbModal
        
    End If
    
    If Trim(cboFormaPago.Text) = "TARJETA DE DEBITO" Then
        cboPlan.Clear
        cbotarjeta.Clear
        cSQL = "SELECT TAR_CODIGO, TAR_DESCRI"
        cSQL = cSQL & " FROM TARJETA"
        cSQL = cSQL & " WHERE TTA_CODIGO=2" 'SOLO TARJETA DE DEBITO
        cSQL = cSQL & " ORDER BY TAR_DESCRI"
        rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
        If (rec.BOF And rec.EOF) = 0 Then
           Do While rec.EOF = False
              cbotarjeta.AddItem Trim(rec!TAR_DESCRI)
              cbotarjeta.ItemData(cbotarjeta.NewIndex) = rec!TAR_CODIGO
              rec.MoveNext
           Loop
           If cbotarjeta.ListCount > 0 Then cbotarjeta.ListIndex = 0
        End If
        rec.Close
        

        'frmDatosTarjeta.cboTarjeta.SetFocus
        frmDatosTarjeta.cboPlan.Enabled = False
        frmDatosTarjeta.txtLote.Enabled = False
        frmDatosTarjeta.txtCupon.Enabled = False
        frmDatosTarjeta.txtTar_Autorizacion.Enabled = False
        frmDatosTarjeta.Show vbModal


    End If

End Sub

Private Sub cboTarjeta_LostFocus()
    Dim mCodTar As String
    mCodTar = cbotarjeta.ItemData(cbotarjeta.ListIndex)
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

Private Sub cmdAceptarCheques_Click()
    Dim i, J As Integer
    'recorro grilla para saber si hay 6 cheques cargados (limite de impresion de cheques en recibo)
    For i = 1 To grillaValores.Rows - 1
        If grillaValores.TextMatrix(i, 0) = "CHE" Then
            J = J + 1
        End If
    Next
    If J < 6 Then
    
        If GrillaCheques.Rows > 1 Then
            'CARGO EN GRILLA VALORES
            For i = 1 To GrillaCheques.Rows - 1
                grillaValores.AddItem "CHE" & Chr(9) & GrillaCheques.TextMatrix(i, 6) & Chr(9) & _
                                      GrillaCheques.TextMatrix(i, 5) & Chr(9) & GrillaCheques.TextMatrix(i, 8) _
                                      & Chr(9) & GrillaCheques.TextMatrix(i, 4) & Chr(9) & GrillaCheques.TextMatrix(i, 7)
            Next
            txtTotalValores.Text = VALIDO_IMPORTE(CStr(SumaGrilla(grillaValores, 1)))
            grillaValores.HighLight = flexHighlightAlways
            GrillaCheques.Rows = 1
            txtTotalCheques.Text = ""
            tabValores.Tab = 0
        End If
    Else
        MsgBox "Ha superado el numero de cheques por recibo, cargue un nuevo recibo", vbInformation, TIT_MSGBOX
    End If
End Sub

Private Sub cmdAceptarPagos_Click()
    If grdPagos.Rows > 1 Then
        'CARGO EN GRILLA VALORES
        For i = 1 To grdPagos.Rows - 1
            grillaValores.AddItem "TAR" & Chr(9) & _
                                  grdPagos.TextMatrix(i, 1) & Chr(9) & _
                                  "" & Chr(9) & _
                                  grdPagos.TextMatrix(i, 4) & Chr(9) & _
                                  grdPagos.TextMatrix(i, 7) & Chr(9) & _
                                  grdPagos.TextMatrix(i, 2)
                                  
        Next
        txtTotalValores.Text = VALIDO_IMPORTE(CStr(SumaGrilla(grillaValores, 1)))
        grillaValores.HighLight = flexHighlightAlways
        GrillaEfectivo.Rows = 1
        txtTotalEfectivo.Text = ""
        tabValores.Tab = 0
    End If
    'cmdGrabar.Enabled = True
    'cmdGrabar.SetFocus
    Unload frmDatosTarjeta
End Sub

Private Sub cmdAceptoTarjeta_Click()
    If cboPlan.ListIndex = -1 Then
        MsgBox "Falta Ingresar el Plan", vbExclamation, TIT_MSGBOX
        cboPlan.SetFocus
        Exit Sub
    End If
    txtImportePago.SetFocus
    fraTarjeta.Visible = False
End Sub

Private Sub cmdAgregarCheque_Click()
    If GrillaCheques.Rows = 7 Then
        MsgBox "No se aceptan mas de 6 cheques por Recibo", vbExclamation, TIT_MSGBOX
        Exit Sub
    Else
    
        If TxtCheNumero.Text = "" Then
            MsgBox "Debe ingresar el número del cheque", vbExclamation, TIT_MSGBOX
            TxtCheNumero.SetFocus
            Exit Sub
        End If
        If TxtBANCO.Text = "" Then
            MsgBox "Debe ingresar el código del banco", vbExclamation, TIT_MSGBOX
            TxtBANCO.SetFocus
            Exit Sub
        End If
        If TxtLOCALIDAD.Text = "" Then
            MsgBox "Debe ingresar el código del banco", vbExclamation, TIT_MSGBOX
            TxtLOCALIDAD.SetFocus
            Exit Sub
        End If
        If TxtSUCURSAL.Text = "" Then
            MsgBox "Debe ingresar el código del banco", vbExclamation, TIT_MSGBOX
            TxtSUCURSAL.SetFocus
            Exit Sub
        End If
        If TxtCODIGO.Text = "" Then
            MsgBox "Debe ingresar el código del banco", vbExclamation, TIT_MSGBOX
            TxtCODIGO.SetFocus
            Exit Sub
        End If
        'VALIDO QUE EL CHEQUE NO SE HAYA CARGADO
        If GrillaCheques.Rows > 1 Then
            If ValidoIngCheques = False Then
                MsgBox "El Cheque ya fue ingresado", vbCritical, TIT_MSGBOX
                TxtCheNumero.Text = ""
                TxtCheNumero.SetFocus
                Exit Sub
            End If
        End If
        
        'CARGO GRILLA
        GrillaCheques.AddItem TxtBANCO.Text & Chr(9) & TxtLOCALIDAD.Text & Chr(9) & _
                              TxtSUCURSAL.Text & Chr(9) & TxtCODIGO.Text & Chr(9) & _
                              TxtCheNumero.Text & Chr(9) & TxtCheFecVto.Value & Chr(9) & _
                              VALIDO_IMPORTE(TxtCheImport.Text) & Chr(9) & TxtCodInt.Text & Chr(9) & TxtBanDescri.Text
        
        
        GrillaCheques.HighLight = flexHighlightAlways
        txtTotalCheques.Text = VALIDO_IMPORTE(CStr(SumaGrilla(GrillaCheques, 6)))
        LimpiarCheques
        cmdAgregarCheque.Enabled = False
        TxtCheNumero.SetFocus
    End If
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
    'txtTotalPagos.Text = Format(CDbl(txtTotal.Text) - mtotalpagos, "0.00")
    cboFormaPago.SetFocus
End Sub

Private Sub cmdBuscaCheque_Click()
    Dim codint As Integer
    frmBuscar.TipoBusqueda = 6
    frmBuscar.Show vbModal
    'TxtCheNumero.Text = frmBuscar.grdBuscar.Col
    TxtCheNumero.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 1)
    TxtBANCO.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 5)
    TxtLOCALIDAD.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 6)
    TxtSUCURSAL.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 7)
    TxtCODIGO.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 8)
    TxtCheImport.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 3)
    TxtCheFecVto.Value = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 2)
    TxtBanDescri.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 0)
    TxtCodInt.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 4)
End Sub

Private Sub cmdCancelarCheques_Click()
    GrillaCheques.Rows = 1
    txtTotalCheques.Text = ""
    tabValores.Tab = 0
End Sub

Private Sub cmdCerrarPagos_Click()
    grdPagos.Rows = 1
    txtTotalPagos.Text = ""
    cboFormaPago.ListIndex = 0
    txtImportePago.Text = ""
    
    tabValores.Tab = 1
End Sub

Private Sub cmdCerrarTarjeta_Click()
    cboFormaPago.ListIndex = 0
    fraTarjeta.Visible = False
    'cboFormaPago.SetFocus
End Sub

Private Sub cmdImprimir_Click()
    Dim J As Integer
    If MsgBox("¿Confirma Impresión del Recibo?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    'PONE A LA IMPRESORA  COMO PREDETERMINADA
    Dim X As Printer
    Dim mDriver As String
    mDriver = Impresora
    For Each X In Printers
        If UCase(X.DeviceName) = UCase(mDriver) Then
            ' La define como predeterminada del sistema.
            Set Printer = X
            Exit For
        End If
    Next
'-----------------------------------
    Set_Impresora
    'SeteoImpresora(256, 1, 7, -1, "Roman 10cpi", 10, False, 12220, 7950)
    J = 1
    For J = 1 To 2
        ImprimirFactura CDbl(J)
    Next J
    Printer.EndDoc
    Screen.MousePointer = vbNormal
    lblestado.Caption = ""

End Sub

Public Sub ImprimirFactura(Fila As Double)
    Dim Renglon As Double
    
    Screen.MousePointer = vbHourglass
    lblestado.Caption = "Imprimiendo..."
    
    'imprimir por duplicado
    If Fila = 2 Then Fila = 17
    ImprimirEncabezado Fila
    
    '---- IMPRESION DE LA FACTURA ------------------
    Renglon = 4
    'Printer.FontSize = 6
    
    'CAMBIAR LA GRILLA Y EL FORMATO DE LA SALIDA IMPRESA DEL RECIBO
    
    Imprimir 1, Fila + Renglon, False, "La cantidad de Pesos ($) " & UCase(EnLetras(txtTotalValores.Text))
    Imprimir 1, Fila + Renglon + 1, False, "En Concepto de pago a cuenta de "
    
    
    For i = 1 To GrillaAplicar.Rows - 1
        If GrillaAplicar.TextMatrix(i, 0) <> "" Then
            Imprimir 5, Fila + Renglon + 1.5, False, Trim(GrillaAplicar.TextMatrix(i, 0)) & " - " & Trim(GrillaAplicar.TextMatrix(i, 1)) & " - " & Trim(GrillaAplicar.TextMatrix(i, 2))
'            Imprimir 10, Fila + Renglon + 2.5, False, " - " & Trim(GrillaAplicar.TextMatrix(I, 1))
'            Imprimir 9.5, Fila + Renglon, False, Trim(GrillaAplicar.TextMatrix(I, 2))
'            Imprimir 11, Fila + Renglon, False, Trim(GrillaAplicar.TextMatrix(I, 3))
'            Imprimir 14, Fila + Renglon, False, Trim(GrillaAplicar.TextMatrix(I, 5))
'            Imprimir 16.2, Fila + Renglon, False, Trim(GrillaAplicar.TextMatrix(I, 6))
'            Imprimir 17.8, Fila + Renglon, False, Trim(GrillaAplicar.TextMatrix(I, 7))

            Renglon = Renglon + 0.8
        End If
    Next i
    
    'IMPRIMO PAGARE
    'HOJA 1
'    Imprimir 8, 2, False, "Sr/a: " & txtRazSoc.Text
    Imprimir 1, Fila + Renglon + 3, False, "Saldo actual: $ " & VALIDO_IMPORTE(txtSaldoActual.Text)
    Imprimir 17, Fila + Renglon + 3, False, "TOTAL: $ " & VALIDO_IMPORTE(txtTotalValores.Text)
'    Imprimir 25, 4, False, "Pilar (Cba.)"
    'Imprimir 1, Fila + Renglon + 1, False, "NOTAS:    "
''    Imprimir 1, Fila + Renglon + 1, False, "       Para cambios presentar esta Boleta"
''    Imprimir 1, Fila + Renglon + 1.5, False, "       Devoluciones o cambios dentro de los 7 dias de efectuada la compra"
'    Imprimir 1, Fila + Renglon + 1, False, "       Credito por Presupuesto: " & txtFacSuc & " - " & txtFacNro
    Imprimir 2, Fila + Renglon + 1.5, False, txtObservaciones.Text


End Sub

Public Sub ImprimirEncabezado(row As Double)
 '-----------IMPRIME EL ENCABEZADO DE RECIBO-------------------

    Dim año As String
    
    Imprimir 1, row - 1, False, "-CENTENARO Y CIA-"
    Imprimir 9.5, row, False, "RECIBO"
    Imprimir 14.5, row + 1, False, "Numero: " & txtNroSucursal.Text & " - " & txtNroRecibo.Text
    Imprimir 14.5, row + 1.5, False, "Fecha:  " & FechaRecibo
    
    Set Rec1 = New ADODB.Recordset
    sql = "SELECT C.CLI_RAZSOC" ',C.CLI_DOMICI, L.LOC_DESCRI"
    'sql = sql & ", P.PRO_DESCRI"
    sql = sql & " FROM CLIENTE C, LOCALIDAD L,"
    sql = sql & " PROVINCIA P"
    sql = sql & " WHERE "
    sql = sql & " CLI_CODIGO=" & XN(txtCodCliente.Text)
'    sql = sql & " AND C.LOC_CODIGO=L.LOC_CODIGO"
'    sql = sql & " AND C.PRO_CODIGO=L.PRO_CODIGO"
'    sql = sql & " AND C.PAI_CODIGO=L.PAI_CODIGO"
'    sql = sql & " AND L.PRO_CODIGO=P.PRO_CODIGO"
    

    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        'Hoja 1
        Imprimir 1, row + 2.5, False, "Recibi conforme de:   " & Trim(Rec1!CLI_RAZSOC)
        'Imprimir 1, row + 1.2, False, "Domicilio: " & Trim(IIf(IsNull(Rec1!CLI_DOMICI), "", Rec1!CLI_DOMICI)) & Trim(Rec1!LOC_DESCRI) & " - " & Trim(Rec1!PRO_DESCRI)
        
    End If
    Rec1.Close
'    'busco forma de pago
'    sql = "SELECT FP.FPG_DESCRI"
'    sql = sql & " FROM FACTURA_CLIENTE FC, FORMA_PAGO FP"
'    sql = sql & " WHERE FC.FPG_CODIGO=FP.FPG_CODIGO"
'    sql = sql & " AND FC.TCO_CODIGO = " & cboNotaCredito.ItemData(cboNotaCredito.ListIndex)
'    sql = sql & " AND FC.FCL_NUMERO = " & XN(txtNrorECIBO.Text)
'    sql = sql & " AND FC.FCL_SUCURSAL = " & XN(txtNroSucursal.Text)
'    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'
'    If rec.EOF = False Then
'        Imprimir 1, row + 2, False, "Forma de Pago:" & rec!FPG_DESCRI
'
'    End If
'
'    rec.Close
    'Hoja 1
    'Imprimir 14.5, row + 2, False, "Vendedor: " & cboVendedor.Text
'    Imprimir 1, row + 4, False, "CODIGO"
'    Imprimir 3, row + 4, False, "DETALLE"
'    Imprimir 9.5, row + 4, False, "TALLE"
'    Imprimir 11, row + 4, False, "COLOR"
'    Imprimir 14, row + 4, False, "CANTIDAD"
'    Imprimir 16.2, row + 4, False, "PRECIO"
'    Imprimir 17.8, row + 4, False, "TOTAL"
    
End Sub




Private Sub ReciboFacturas()
    Set Rec1 = New ADODB.Recordset
    'BUSCO FACTURAS_PROVEEDOR
    sql = "SELECT C.CLI_RAZSOC, C.CLI_DOMICI, L.LOC_DESCRI, PR.PRO_DESCRI, C.CLI_CUIT, C.CLI_INGBRU"
    sql = sql & ",CI.IVA_DESCRI, TC.TCO_ABREVIA, FR.FCL_SUCURSAL, FR.FCL_NUMERO, FR.FCL_FECHA ,F.FCL_TOTAL, FR.REC_IMPORTE, R.REC_TOTAL"
    sql = sql & " FROM CLIENTE C, RECIBO_CLIENTE R ,CONDICION_IVA CI ,LOCALIDAD L"
    sql = sql & " , PROVINCIA PR, TIPO_COMPROBANTE TC, FACTURAS_RECIBO_CLIENTE FR,"
    sql = sql & " FACTURA_CLIENTE F"
    sql = sql & " WHERE R.REC_NUMERO=" & XN(txtNroRecibo.Text)
    sql = sql & " AND R.REC_SUCURSAL=" & XN(txtNroSucursal.Text)
    sql = sql & " AND R.TCO_CODIGO=" & XN(cboRecibo.ItemData(cboRecibo.ListIndex))
    sql = sql & " AND R.REC_NUMERO=FR.REC_NUMERO"
    sql = sql & " AND R.REC_SUCURSAL=FR.REC_SUCURSAL"
    sql = sql & " AND R.TCO_CODIGO=FR.TCO_CODIGO"
    sql = sql & " AND R.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & " AND C.LOC_CODIGO=L.LOC_CODIGO"
    sql = sql & " AND C.PAI_CODIGO=L.PAI_CODIGO"
    sql = sql & " AND C.PRO_CODIGO=L.PRO_CODIGO"
    sql = sql & " AND L.PRO_CODIGO=PR.PRO_CODIGO"
    sql = sql & " AND L.PAI_CODIGO=PR.PAI_CODIGO"
    sql = sql & " AND C.IVA_CODIGO=CI.IVA_CODIGO"
    sql = sql & " AND FR.FCL_TCO_CODIGO=TC.TCO_CODIGO"
    sql = sql & " AND FR.FCL_TCO_CODIGO=F.TCO_CODIGO"
    sql = sql & " AND FR.FCL_SUCURSAL=F.FCL_SUCURSAL"
    sql = sql & " AND FR.FCL_NUMERO=F.FCL_NUMERO"

    'BUSCAR NOTA_DEBITO_PROVEEDOR
    sql = sql & " UNION ALL"
    sql = sql & " SELECT C.CLI_RAZSOC, C.CLI_DOMICI, L.LOC_DESCRI, PR.PRO_DESCRI, C.CLI_CUIT, C.CLI_INGBRU"
    sql = sql & ",CI.IVA_DESCRI, TC.TCO_ABREVIA, FR.FCL_SUCURSAL, FR.FCL_NUMERO, FR.FCL_FECHA ,N.NDC_TOTAL, FR.REC_IMPORTE, R.REC_TOTAL"
    sql = sql & " FROM CLIENTE C, RECIBO_CLIENTE R ,CONDICION_IVA CI ,LOCALIDAD L"
    sql = sql & " , PROVINCIA PR, TIPO_COMPROBANTE TC, FACTURAS_RECIBO_CLIENTE FR,"
    sql = sql & " NOTA_DEBITO_CLIENTE N"
    sql = sql & " WHERE R.REC_NUMERO=" & XN(txtNroRecibo.Text)
    sql = sql & " AND R.REC_SUCURSAL=" & XN(txtNroSucursal.Text)
    sql = sql & " AND R.TCO_CODIGO=" & XN(cboRecibo.ItemData(cboRecibo.ListIndex))
    sql = sql & " AND R.REC_NUMERO=FR.REC_NUMERO"
    sql = sql & " AND R.REC_SUCURSAL=FR.REC_SUCURSAL"
    sql = sql & " AND R.TCO_CODIGO=FR.TCO_CODIGO"
    sql = sql & " AND R.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & " AND C.LOC_CODIGO=L.LOC_CODIGO"
    sql = sql & " AND C.PAI_CODIGO=L.PAI_CODIGO"
    sql = sql & " AND C.PRO_CODIGO=L.PRO_CODIGO"
    sql = sql & " AND L.PRO_CODIGO=PR.PRO_CODIGO"
    sql = sql & " AND L.PAI_CODIGO=PR.PAI_CODIGO"
    sql = sql & " AND C.IVA_CODIGO=CI.IVA_CODIGO"
    sql = sql & " AND FR.FCL_TCO_CODIGO=TC.TCO_CODIGO"
    sql = sql & " AND FR.FCL_TCO_CODIGO=N.TCO_CODIGO"
    sql = sql & " AND FR.FCL_SUCURSAL=N.NDC_SUCURSAL"
    sql = sql & " AND FR.FCL_NUMERO=N.NDC_NUMERO"
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
            
            sql = "INSERT INTO TMP_RECIBO_CLIENTE ("
            sql = sql & "REC_NUMERO,REC_FECHA,CLI_RAZSOC,CLI_DOMICI,CLI_CUIT,CLI_INGBRU,"
            sql = sql & "LOC_DESCRI,PRO_DESCRI,IVA_DESCRI,TCO_ABREVIA,COM_FECHA,COM_NUMERO,COM_IMPORTE,"
            sql = sql & "REC_TOTAL,FAC_ABREVIA,FAC_NUMERO,FAC_FECHA,FAC_IMPORTE,FAC_TOTAL,REC_ITEM) VALUES ("
            sql = sql & XS(Format(txtNroSucursal.Text, "0000") & "-" & Format(txtNroRecibo.Text, "00000000")) & ","
            sql = sql & XDQ(FechaRecibo.Value) & ","
            sql = sql & XS(Rec1!CLI_RAZSOC) & ","
            sql = sql & XS(Rec1!CLI_DOMICI) & ","
            sql = sql & XS(Format(Rec1!CLI_CUIT, "##-########-#")) & ","
            sql = sql & XS(Format(Rec1!CLI_INGBRU, "###-#####-##")) & ","
            sql = sql & XS(Rec1!LOC_DESCRI) & ","
            sql = sql & XS(Rec1!PRO_DESCRI) & ","
            sql = sql & XS(Rec1!IVA_DESCRI) & ","
            sql = sql & "NULL,"
            sql = sql & "NULL,"
            sql = sql & "NULL,"
            sql = sql & "NULL,"
            sql = sql & XN(Rec1!REC_TOTAL) & ","
            sql = sql & XS(Rec1!TCO_ABREVIA) & ","
            sql = sql & XS(Format(Rec1!FCL_SUCURSAL, "0000") & "-" & Format(Rec1!FCL_NUMERO, "00000000")) & ","
            sql = sql & XS(Rec1!FCL_FECHA) & ","
            sql = sql & XN(Rec1!REC_IMPORTE) & ","
            sql = sql & XN(Rec1!FCL_TOTAL) & ","
            sql = sql & i & ")"
            DBConn.Execute sql
            
            i = i + 1
            Rec1.MoveNext
        Loop
    End If
    Rec1.Close
End Sub

Private Sub ReciboComprobante()
    Set Rec1 = New ADODB.Recordset
    sql = "SELECT C.CLI_RAZSOC, C.CLI_DOMICI, L.LOC_DESCRI, PR.PRO_DESCRI, C.CLI_CUIT, C.CLI_INGBRU"
    sql = sql & ",CI.IVA_DESCRI, TC.TCO_ABREVIA, DR.DRE_COMFECHA, DR.DRE_COMSUCURSAL ,DR.DRE_COMNUMERO, DR.DRE_COMIMP, R.REC_TOTAL"
    sql = sql & " FROM CLIENTE C, DETALLE_RECIBO_CLIENTE DR, RECIBO_CLIENTE R ,CONDICION_IVA CI"
    sql = sql & " ,LOCALIDAD L, PROVINCIA PR, TIPO_COMPROBANTE TC"
    sql = sql & " WHERE R.REC_NUMERO=" & XN(txtNroRecibo.Text)
    sql = sql & " AND R.REC_SUCURSAL=" & XN(txtNroSucursal.Text)
    sql = sql & " AND R.TCO_CODIGO=" & XN(cboRecibo.ItemData(cboRecibo.ListIndex))
    sql = sql & " AND R.REC_NUMERO=DR.REC_NUMERO"
    sql = sql & " AND R.REC_SUCURSAL=DR.REC_SUCURSAL"
    sql = sql & " AND R.TCO_CODIGO=DR.TCO_CODIGO"
    sql = sql & " AND R.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & " AND C.LOC_CODIGO=L.LOC_CODIGO"
    sql = sql & " AND C.PAI_CODIGO=L.PAI_CODIGO"
    sql = sql & " AND C.PRO_CODIGO=L.PRO_CODIGO"
    sql = sql & " AND L.PRO_CODIGO=PR.PRO_CODIGO"
    sql = sql & " AND L.PAI_CODIGO=PR.PAI_CODIGO"
    sql = sql & " AND C.IVA_CODIGO=CI.IVA_CODIGO"
    sql = sql & " AND DR.DRE_TCO_CODIGO=TC.TCO_CODIGO"
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
                        
            sql = "INSERT INTO TMP_RECIBO_CLIENTE ("
            sql = sql & "REC_NUMERO,REC_FECHA,CLI_RAZSOC,CLI_DOMICI,CLI_CUIT,CLI_INGBRU,"
            sql = sql & "LOC_DESCRI,PRO_DESCRI,IVA_DESCRI,TCO_ABREVIA,COM_FECHA,COM_NUMERO,COM_IMPORTE,"
            sql = sql & "REC_TOTAL,REC_ITEM) VALUES ("
            sql = sql & XS(Format(txtNroSucursal.Text, "0000") & "-" & Format(txtNroRecibo.Text, "00000000")) & ","
            sql = sql & XDQ(FechaRecibo.Value) & ","
            sql = sql & XS(Rec1!CLI_RAZSOC) & ","
            sql = sql & XS(Rec1!CLI_DOMICI) & ","
            sql = sql & XS(Format(Rec1!CLI_CUIT, "##-########-#")) & ","
            sql = sql & XS(Format(Rec1!CLI_INGBRU, "###-#####-##")) & ","
            sql = sql & XS(Rec1!LOC_DESCRI) & ","
            sql = sql & XS(Rec1!PRO_DESCRI) & ","
            sql = sql & XS(Rec1!IVA_DESCRI) & ","
            sql = sql & XS(Rec1!TCO_ABREVIA) & ","
            sql = sql & XDQ(Rec1!DRE_COMFECHA) & ","
            sql = sql & XS(Rec1!DRE_COMSUCURSAL & "-" & Format(Rec1!DRE_COMNUMERO, "00000000")) & ","
            sql = sql & XN(Rec1!DRE_COMIMP) & ","
            sql = sql & XN(Rec1!REC_TOTAL) & ","
            sql = sql & i & ")"
            DBConn.Execute sql
            
            i = i + 1
            Rec1.MoveNext
        Loop
    End If
    Rec1.Close
End Sub

Private Sub ReciboCheques()
    Set Rec1 = New ADODB.Recordset
    'PARA CHEQUES DE TERCEROS
    sql = "SELECT C.CLI_RAZSOC, C.CLI_DOMICI, L.LOC_DESCRI, PR.PRO_DESCRI, C.CLI_CUIT, C.CLI_INGBRU"
    sql = sql & ",CI.IVA_DESCRI, B.BAN_NOMCOR, CH.CHE_FECVTO ,DR.CHE_NUMERO, CH.CHE_IMPORT, R.REC_TOTAL"
    sql = sql & " FROM CLIENTE C, DETALLE_RECIBO_CLIENTE DR, RECIBO_CLIENTE R ,CONDICION_IVA CI"
    sql = sql & " ,LOCALIDAD L, PROVINCIA PR, CHEQUE CH, BANCO B"
    sql = sql & " WHERE R.REC_NUMERO=" & XN(txtNroRecibo.Text)
    sql = sql & " AND R.REC_SUCURSAL=" & XN(txtNroSucursal.Text)
    sql = sql & " AND R.TCO_CODIGO=" & XN(cboRecibo.ItemData(cboRecibo.ListIndex))
    sql = sql & " AND R.REC_NUMERO=DR.REC_NUMERO"
    sql = sql & " AND R.REC_SUCURSAL=DR.REC_SUCURSAL"
    sql = sql & " AND R.TCO_CODIGO=DR.TCO_CODIGO"
    sql = sql & " AND R.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & " AND C.LOC_CODIGO=L.LOC_CODIGO"
    sql = sql & " AND C.PAI_CODIGO=L.PAI_CODIGO"
    sql = sql & " AND C.PRO_CODIGO=L.PRO_CODIGO"
    sql = sql & " AND L.PRO_CODIGO=PR.PRO_CODIGO"
    sql = sql & " AND L.PAI_CODIGO=PR.PAI_CODIGO"
    sql = sql & " AND C.IVA_CODIGO=CI.IVA_CODIGO"
    sql = sql & " AND DR.BAN_CODINT=CH.BAN_CODINT"
    sql = sql & " AND DR.CHE_NUMERO=CH.CHE_NUMERO"
    sql = sql & " AND CH.BAN_CODINT=B.BAN_CODINT"
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
            
            sql = "INSERT INTO TMP_RECIBO_CLIENTE ("
            sql = sql & "REC_NUMERO,REC_FECHA,CLI_RAZSOC,CLI_DOMICI,CLI_CUIT,CLI_INGBRU,"
            sql = sql & "LOC_DESCRI,PRO_DESCRI,IVA_DESCRI,TCO_ABREVIA,COM_FECHA,COM_NUMERO,COM_IMPORTE,"
            sql = sql & "REC_TOTAL,REC_ITEM) VALUES ("
            sql = sql & XS(Format(txtNroSucursal.Text, "0000") & "-" & Format(txtNroRecibo.Text, "00000000")) & ","
            sql = sql & XDQ(FechaRecibo.Value) & ","
            sql = sql & XS(Rec1!CLI_RAZSOC) & ","
            sql = sql & XS(Rec1!CLI_DOMICI) & ","
            sql = sql & XS(Format(Rec1!CLI_CUIT, "##-########-#")) & ","
            sql = sql & XS(Format(Rec1!CLI_INGBRU, "###-#####-##")) & ","
            sql = sql & XS(Rec1!LOC_DESCRI) & ","
            sql = sql & XS(Rec1!PRO_DESCRI) & ","
            sql = sql & XS(Rec1!IVA_DESCRI) & ","
            sql = sql & XS(Rec1!BAN_NOMCOR) & ","
            sql = sql & XDQ(Rec1!CHE_FECVTO) & ","
            sql = sql & XS(Rec1!CHE_NUMERO) & ","
            sql = sql & XN(Rec1!che_import) & ","
            sql = sql & XN(Rec1!REC_TOTAL) & ","
            sql = sql & i & ")"
            DBConn.Execute sql
            
            i = i + 1
            Rec1.MoveNext
        Loop
    End If
    Rec1.Close
    'PARA CHEQUES PROPIOS
    sql = "SELECT C.CLI_RAZSOC, C.CLI_DOMICI, L.LOC_DESCRI, PR.PRO_DESCRI, C.CLI_CUIT, C.CLI_INGBRU"
    sql = sql & ",CI.IVA_DESCRI, B.BAN_NOMCOR, CH.CHEP_FECVTO ,DR.CHE_NUMERO, CH.CHEP_IMPORT, R.REC_TOTAL"
    sql = sql & " FROM CLIENTE C,DETALLE_RECIBO_CLIENTE DR, RECIBO_CLIENTE R ,CONDICION_IVA CI"
    sql = sql & " ,LOCALIDAD L, PROVINCIA PR, CHEQUE_PROPIO CH, BANCO B"
    sql = sql & " WHERE R.REC_NUMERO=" & XN(txtNroRecibo.Text)
    sql = sql & " AND R.REC_SUCURSAL=" & XN(txtNroSucursal.Text)
    sql = sql & " AND R.TCO_CODIGO=" & XN(cboRecibo.ItemData(cboRecibo.ListIndex))
    sql = sql & " AND R.REC_NUMERO=DR.REC_NUMERO"
    sql = sql & " AND R.REC_SUCURSAL=DR.REC_SUCURSAL"
    sql = sql & " AND R.TCO_CODIGO=DR.TCO_CODIGO"
    sql = sql & " AND R.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & " AND C.LOC_CODIGO=L.LOC_CODIGO"
    sql = sql & " AND C.PAI_CODIGO=L.PAI_CODIGO"
    sql = sql & " AND C.PRO_CODIGO=L.PRO_CODIGO"
    sql = sql & " AND L.PRO_CODIGO=PR.PRO_CODIGO"
    sql = sql & " AND L.PAI_CODIGO=PR.PAI_CODIGO"
    sql = sql & " AND C.IVA_CODIGO=CI.IVA_CODIGO"
    sql = sql & " AND DR.BAN_CODINT=CH.BAN_CODINT"
    sql = sql & " AND DR.CHE_NUMERO=CH.CHEP_NUMERO"
    sql = sql & " AND CH.BAN_CODINT=B.BAN_CODINT"
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
            
            sql = "INSERT INTO TMP_RECIBO_CLIENTE ("
            sql = sql & "REC_NUMERO,REC_FECHA,CLI_RAZSOC,CLI_DOMICI,CLI_CUIT,CLI_INGBRU,"
            sql = sql & "LOC_DESCRI,PRO_DESCRI,IVA_DESCRI,TCO_ABREVIA,COM_FECHA,COM_NUMERO,COM_IMPORTE,"
            sql = sql & "REC_TOTAL,REC_ITEM) VALUES ("
            sql = sql & XS(Format(txtNroSucursal.Text, "0000") & "-" & Format(txtNroRecibo.Text, "00000000")) & ","
            sql = sql & XDQ(FechaRecibo.Value) & ","
            sql = sql & XS(Rec1!CLI_RAZSOC) & ","
            sql = sql & XS(Rec1!CLI_DOMICI) & ","
            sql = sql & XS(Format(Rec1!CLI_CUIT, "##-########-#")) & ","
            sql = sql & XS(Format(Rec1!CLI_INGBRU, "###-#####-##")) & ","
            sql = sql & XS(Rec1!LOC_DESCRI) & ","
            sql = sql & XS(Rec1!PRO_DESCRI) & ","
            sql = sql & XS(Rec1!IVA_DESCRI) & ","
            sql = sql & XS(Rec1!BAN_NOMCOR) & ","
            sql = sql & XDQ(Rec1!CHEP_FECVTO) & ","
            sql = sql & XS(Rec1!CHE_NUMERO) & ","
            sql = sql & XN(Rec1!CHEP_IMPORT) & ","
            sql = sql & XN(Rec1!OPG_TOTAL) & ","
            sql = sql & i & ")"
            DBConn.Execute sql
            
            i = i + 1
            Rec1.MoveNext
        Loop
    End If
    Rec1.Close
End Sub

Private Sub ReciboMoneda()
    Set Rec1 = New ADODB.Recordset
    sql = "SELECT C.CLI_RAZSOC, C.CLI_DOMICI, L.LOC_DESCRI, PR.PRO_DESCRI, C.CLI_CUIT, C.CLI_INGBRU"
    sql = sql & ", M.MON_DESCRI, DR.DRE_MONIMP, R.REC_TOTAL, CI.IVA_DESCRI"
    sql = sql & " FROM CLIENTE C, DETALLE_RECIBO_CLIENTE DR, RECIBO_CLIENTE R"
    sql = sql & " ,LOCALIDAD L, PROVINCIA PR, MONEDA M, CONDICION_IVA CI"
    sql = sql & " WHERE R.REC_NUMERO=" & XN(txtNroRecibo.Text)
    sql = sql & " AND R.REC_SUCURSAL=" & XN(txtNroSucursal.Text)
    sql = sql & " AND R.TCO_CODIGO=" & XN(cboRecibo.ItemData(cboRecibo.ListIndex))
    sql = sql & " AND R.REC_NUMERO=DR.REC_NUMERO"
    sql = sql & " AND R.REC_SUCURSAL=DR.REC_SUCURSAL"
    sql = sql & " AND R.TCO_CODIGO=DR.TCO_CODIGO"
    sql = sql & " AND R.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & " AND DR.MON_CODIGO=M.MON_CODIGO"
    sql = sql & " AND C.LOC_CODIGO=L.LOC_CODIGO"
    sql = sql & " AND C.PAI_CODIGO=L.PAI_CODIGO"
    sql = sql & " AND C.PRO_CODIGO=L.PRO_CODIGO"
    sql = sql & " AND L.PRO_CODIGO=PR.PRO_CODIGO"
    sql = sql & " AND L.PAI_CODIGO=PR.PAI_CODIGO"
    sql = sql & " AND C.IVA_CODIGO=CI.IVA_CODIGO"
    
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
            
            sql = "INSERT INTO TMP_RECIBO_CLIENTE ("
            sql = sql & "REC_NUMERO,REC_FECHA,CLI_RAZSOC,CLI_DOMICI,CLI_CUIT,CLI_INGBRU,"
            sql = sql & "LOC_DESCRI,PRO_DESCRI,IVA_DESCRI,TCO_ABREVIA,COM_IMPORTE,"
            sql = sql & "REC_TOTAL,REC_ITEM) VALUES ("
            sql = sql & XS(Format(txtNroSucursal.Text, "0000") & "-" & Format(txtNroRecibo.Text, "00000000")) & ","
            sql = sql & XDQ(FechaRecibo.Value) & ","
            sql = sql & XS(Rec1!CLI_RAZSOC) & ","
            sql = sql & XS(Rec1!CLI_DOMICI) & ","
            sql = sql & XS(Format(Rec1!CLI_CUIT, "##-########-#")) & ","
            sql = sql & XS(Format(Rec1!CLI_INGBRU, "###-#####-##")) & ","
            sql = sql & XS(Rec1!LOC_DESCRI) & ","
            sql = sql & XS(Rec1!PRO_DESCRI) & ","
            sql = sql & XS(Rec1!IVA_DESCRI) & ","
            sql = sql & XS(Rec1!MON_DESCRI) & ","
            sql = sql & XN(Rec1!DRE_MONIMP) & ","
            sql = sql & XN(Rec1!REC_TOTAL) & ","
            sql = sql & i & ")"
            DBConn.Execute sql
            
            i = i + 1
            Rec1.MoveNext
        Loop
    End If
    Rec1.Close
End Sub

Private Sub cmaAceptarACta_Click()
    txtSaldoACta.Text = ""
    txtImporteACta.Text = ""
    tabValores.Tab = 0
End Sub

Private Sub cmdAceptarMoneda_Click()
    If GrillaEfectivo.Rows > 1 Then
        'CARGO EN GRILLA VALORES
        For i = 1 To GrillaEfectivo.Rows - 1
            grillaValores.AddItem "EFT" & Chr(9) & _
                                  GrillaEfectivo.TextMatrix(i, 1) & Chr(9) & _
                                  "" & Chr(9) & _
                                  GrillaEfectivo.TextMatrix(i, 0) & Chr(9) & _
                                  "" & Chr(9) & _
                                  GrillaEfectivo.TextMatrix(i, 2)
        Next
        txtTotalValores.Text = VALIDO_IMPORTE(CStr(SumaGrilla(grillaValores, 1)))
        grillaValores.HighLight = flexHighlightAlways
        GrillaEfectivo.Rows = 1
        txtTotalEfectivo.Text = ""
        tabValores.Tab = 0
    End If
    cmdGrabar.Enabled = True
    cmdGrabar.SetFocus
    
    If CDbl(txtImporteApagar.Text) > CDbl(txtTotalValores.Text) Then
        txtSaldoEftTar = CDbl(txtImporteApagar.Text) - CDbl(txtTotalValores.Text)
    End If
    txtSaldoActual.Text = CDbl(Chk0(txtSaldo.Text)) - CDbl(txtTotalValores)
    txtSaldoActual.Text = VALIDO_IMPORTE(txtSaldoActual.Text)
    'MsgBox EnLetras(txtTotalValores.Text)
End Sub

Private Sub cmdAgregarACta_Click()
    If GrillaAFavor.Rows > 1 Then
        If grillaValores.Rows > 1 Then
            For i = 1 To grillaValores.Rows - 1
                If GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 5) = grillaValores.TextMatrix(i, 5) _
                    And (GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 1)) = (grillaValores.TextMatrix(i, 4)) _
                    And CDate(GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 2)) = CDate(grillaValores.TextMatrix(i, 2)) _
                    And (GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 6)) = (grillaValores.TextMatrix(i, 6)) Then
                   MsgBox "El Valor ya fue ingresado", vbInformation, TIT_MSGBOX
                   txtSaldoACta.Text = ""
                   txtImporteACta.Text = ""
                   GrillaAFavor.SetFocus
                   Exit Sub
                End If
            Next
        End If
                
        'CARGO EN GRILLA VALORES
        grillaValores.AddItem "A-CTA" & Chr(9) & VALIDO_IMPORTE(txtImporteACta) & Chr(9) & GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 2) _
                                & Chr(9) & GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 0) & Chr(9) & GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 1) & Chr(9) & _
                                GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 5) & Chr(9) & GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 6)

        'ARREGLO EL SALDO DEL DINERO A CTA
        GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 4) = VALIDO_IMPORTE(CStr(CDbl(GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 4)) - CDbl(Chk0(txtImporteACta.Text))))
        
        txtTotalValores.Text = VALIDO_IMPORTE(CStr(SumaGrilla(grillaValores, 1)))
        grillaValores.HighLight = flexHighlightAlways

        txtSaldoACta.Text = ""
        txtImporteACta.Text = ""
        GrillaAFavor.SetFocus
    End If
End Sub

Private Function ValidoIngCheques() As Boolean
    For i = 1 To GrillaCheques.Rows - 1
        If TxtCodInt.Text = GrillaCheques.TextMatrix(i, 7) And _
           TxtCheNumero.Text = GrillaCheques.TextMatrix(i, 4) Then

           ValidoIngCheques = False
           Exit Function
        End If
    Next
    ValidoIngCheques = True
End Function

Private Sub LimpiarCheques()
    TxtBANCO.Text = ""
    TxtLOCALIDAD.Text = ""
    TxtSUCURSAL.Text = ""
    TxtCODIGO.Text = ""
    TxtCheNumero.Text = ""
    'TxtCheFecVto.Text = ""
    TxtCheImport.Text = ""
    TxtCodInt.Text = ""
    TxtBanDescri.Text = ""
    frameBanco.Enabled = False
    cmdAgregarCheque.Enabled = False
End Sub

Private Function BuscarTipoDocAbre(Codigo As String) As String
    Set Rec1 = New ADODB.Recordset
    sql = "SELECT TCO_ABREVIA"
    sql = sql & " FROM TIPO_COMPROBANTE"
    sql = sql & " WHERE TCO_CODIGO = " & XN(Codigo)
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        BuscarTipoDocAbre = Rec1!TCO_ABREVIA
    Else
        BuscarTipoDocAbre = ""
    End If
    Rec1.Close
End Function

Private Sub cmdAgregarEfectivo_Click()
    'VALIDO QUE EL CHEQUE NO SE HAYA CARGADO
    'If GrillaEfectivo.Rows > 1 Then
    '    If ValidoIngMoneda = False Then
    '        MsgBox "La Moneda ya fue ingresada", vbCritical, TIT_MSGBOX
    '        txtEftImporte.Text = ""
    '        cboMoneda.SetFocus
    '        Exit Sub
    '    End If
    'End If
    
    Dim TotalDineroaCta As String
    TotalDineroaCta = "0"
    
    If grillaValores.Rows > 1 Then
       For i = 1 To grillaValores.Rows - 1
          TotalDineroaCta = CDbl(TotalDineroaCta) + CDbl(grillaValores.TextMatrix(i, 1))
       Next i
       If txtEftImporte.Text = "" Then
          txtEftImporte.Text = Format(txtImporteApagar.Text - CDbl(TotalDineroaCta), "0.00")
       End If
    Else
       If txtEftImporte.Text = "" Then
          txtEftImporte.Text = txtImporteApagar.Text
       End If
    End If
    
    'CARGO GRILLA
    GrillaEfectivo.AddItem cboMoneda.Text & Chr(9) & txtEftImporte.Text & Chr(9) & cboMoneda.ItemData(cboMoneda.ListIndex)
                                                   
    GrillaEfectivo.HighLight = flexHighlightAlways
    txtTotalEfectivo.Text = VALIDO_IMPORTE(CStr(SumaGrilla(GrillaEfectivo, 1)))
    'txtEftImporte.Text = ""
    'cboMoneda.SetFocus
    cmdAceptarMoneda.SetFocus
End Sub

Private Function ValidoIngMoneda() As Boolean
    For i = 1 To GrillaEfectivo.Rows - 1
        If cboMoneda.ItemData(cboMoneda.ListIndex) = GrillaEfectivo.TextMatrix(i, 2) Then
           
           ValidoIngMoneda = False
           Exit Function
        End If
    Next
    ValidoIngMoneda = True
End Function

Private Function ValidoIngFactura(Combo As ComboBox, Grilla As MSFlexGrid, NROFAC As String, NroSuc As String) As Boolean
    For i = 1 To Grilla.Rows - 1
        If Combo.ItemData(Combo.ListIndex) = Grilla.TextMatrix(i, 4) And _
           NROFAC = Right(Grilla.TextMatrix(i, 1), 8) And _
           NroSuc = Left(Grilla.TextMatrix(i, 1), 4) Then
           
           ValidoIngFactura = False
           Exit Function
        End If
    Next
    ValidoIngFactura = True
End Function

Private Sub CmdBuscAprox_Click()
    GrdModulos.Rows = 1
    GrdModulos.HighLight = flexHighlightNever
    lblestado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
    
    Set Rec1 = New ADODB.Recordset
    sql = "SELECT RC.REC_NUMERO, RC.REC_SUCURSAL,"
    sql = sql & " RC.REC_FECHA, RC.TCO_CODIGO, TC.TCO_ABREVIA, C.CLI_RAZSOC, RC.REC_TOTAL "
    sql = sql & " FROM RECIBO_CLIENTE RC, CLIENTE C,  TIPO_COMPROBANTE TC"
    sql = sql & " WHERE RC.TCO_CODIGO=TC.TCO_CODIGO"
    sql = sql & "   AND RC.CLI_CODIGO=C.CLI_CODIGO"
    If txtCliente.Text <> "" Then sql = sql & " AND RC.CLI_CODIGO=" & XN(txtCliente.Text)
    If FechaDesde.Value <> "" Then sql = sql & " AND RC.REC_FECHA>=" & XDQ(FechaDesde.Value)
    If FechaHasta.Value <> "" Then sql = sql & " AND RC.REC_FECHA<=" & XDQ(FechaHasta.Value)
    If cboRecibo1.List(cboRecibo1.ListIndex) <> "(Todos)" Then sql = sql & " AND RC.TCO_CODIGO=" & XN(cboRecibo1.ItemData(cboRecibo1.ListIndex))
    sql = sql & " ORDER BY RC.REC_SUCURSAL, RC.REC_NUMERO"
    
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
            GrdModulos.AddItem Rec1!TCO_ABREVIA & Chr(9) & Format(Rec1!REC_SUCURSAL, "0000") & "-" & Format(Rec1!REC_NUMERO, "00000000") _
                               & Chr(9) & Rec1!REC_FECHA & Chr(9) & Rec1!CLI_RAZSOC _
                               & Chr(9) & Rec1!TCO_CODIGO & Chr(9) & VALIDO_IMPORTE(Rec1!REC_TOTAL)
            Rec1.MoveNext
        Loop
        GrdModulos.HighLight = flexHighlightAlways
        GrdModulos.SetFocus
        GrdModulos.Col = 0
    Else
        lblestado.Caption = ""
        Screen.MousePointer = vbNormal
        MsgBox "No se encontraron datos... ", vbExclamation, TIT_MSGBOX
        txtCliente.SetFocus
    End If
    Rec1.Close
    lblestado.Caption = ""
    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdBuscarCliente_Click()
    frmBuscar.TipoBusqueda = 1
    frmBuscar.TxtDescriB = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 0
        txtCodCliente.Text = frmBuscar.grdBuscar.Text
        txtCodCliente_LostFocus
    Else
        txtCodCliente.SetFocus
    End If
End Sub

Private Sub cmdCancelarMoneda_Click()
    GrillaEfectivo.Rows = 1
    txtTotalEfectivo.Text = ""
    tabValores.Tab = 0
End Sub

Private Sub cmdGrabar_Click()
    If ValidarRecibo = False Then Exit Sub
    If MsgBox("¿Confirma Recibo?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    
    On Error GoTo HayError
    DBConn.BeginTrans
    Screen.MousePointer = vbHourglass
    lblestado.Caption = "Guardando..."
    
    sql = "SELECT EST_CODIGO"
    sql = sql & " FROM RECIBO_CLIENTE"
    sql = sql & " WHERE TCO_CODIGO = " & XN(cboRecibo.ItemData(cboRecibo.ListIndex))
    sql = sql & "   AND REC_NUMERO = " & XN(txtNroRecibo.Text)
    sql = sql & "   AND REC_SUCURSAL = " & XN(txtNroSucursal.Text)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = True Then
        
        'CABEZA DEL RECIBO
        sql = "INSERT INTO RECIBO_CLIENTE ("
        sql = sql & " TCO_CODIGO, REC_NUMERO, REC_SUCURSAL, REC_FECHA,"
        sql = sql & " EST_CODIGO, CLI_CODIGO,"
        sql = sql & " REC_NUMEROTXT, REC_TOTAL,REC_OBSER,REC_SACTUAL)"
        sql = sql & " VALUES ("
        sql = sql & XN(cboRecibo.ItemData(cboRecibo.ListIndex)) & ", "
        sql = sql & XN(txtNroRecibo.Text) & ","
        sql = sql & XN(txtNroSucursal.Text) & ","
        sql = sql & XDQ(FechaRecibo.Value) & ","
        sql = sql & "3,"                          'ESTADO DEFINITIVO
        sql = sql & XN(txtCodCliente.Text) & ","
        sql = sql & XS(Format(txtNroRecibo.Text, "00000000")) & ","
        sql = sql & XN(txtpagTar.Text) & "," 'uso este campo oculto pq si el pago es con tarjeta no toma en cuenta el recargo
        sql = sql & XS(txtObservaciones.Text) & ","
        sql = sql & XN(txtSaldoActual) & ")"
        DBConn.Execute sql
        
        'DETALLE DEL RECIBO
        For i = 1 To grillaValores.Rows - 1
            sql = "INSERT INTO DETALLE_RECIBO_CLIENTE"
            sql = sql & " (TCO_CODIGO, REC_NUMERO, REC_SUCURSAL, REC_FECHA,"
            sql = sql & " DRE_NROITEM, BAN_CODINT, CHE_NUMERO, MON_CODIGO,"
            sql = sql & " DRE_MONIMP, DRE_TCO_CODIGO, DRE_COMFECHA, DRE_COMNUMERO,"
            sql = sql & " DRE_COMSUCURSAL, DRE_COMIMP,"
            sql = sql & " FPG_CODIGO,PAG_IMPORTE,TAR_CODIGO,TAR_PLAN,TAR_CUPON,TAR_LOTE,TAR_AUTORIZACION)"
            sql = sql & " VALUES ("
            sql = sql & XN(cboRecibo.ItemData(cboRecibo.ListIndex)) & ","
            sql = sql & XN(txtNroRecibo.Text) & ","
            sql = sql & XN(txtNroSucursal.Text) & ","
            sql = sql & XDQ(FechaRecibo.Value) & ","
            sql = sql & XN(CStr(i)) & ","
            If grillaValores.TextMatrix(i, 0) = "CHE" Then
                sql = sql & XN(grillaValores.TextMatrix(i, 5)) & ","
                sql = sql & XS(grillaValores.TextMatrix(i, 4)) & "," 'NUMERO DE CHEQUE
            Else
                sql = sql & "NULL,NULL,"
            End If
            If grillaValores.TextMatrix(i, 0) = "EFT" Then
                sql = sql & XN(grillaValores.TextMatrix(i, 5)) & "," 'MONEDA
                sql = sql & XN(grillaValores.TextMatrix(i, 1)) & "," 'IMPORTE
            Else
                sql = sql & "NULL,NULL,"
            End If
            
            If grillaValores.TextMatrix(i, 0) = "COMP" Or grillaValores.TextMatrix(i, 0) = "A-CTA" Then
                sql = sql & XN(grillaValores.TextMatrix(i, 5)) & ","
                sql = sql & XDQ(grillaValores.TextMatrix(i, 2)) & ","
                sql = sql & XN(Right(grillaValores.TextMatrix(i, 4), 8)) & "," 'NUMERO COMPROBANTE
                sql = sql & XN(Left(grillaValores.TextMatrix(i, 4), 4)) & ","  'NUMERO SUCURSAL
                sql = sql & XN(grillaValores.TextMatrix(i, 1)) & ")"
            Else
                sql = sql & "NULL,NULL,NULL,NULL,NULL"
            End If
            If grillaValores.TextMatrix(i, 0) = "TAR" Then
                If (i > 1) And (grillaValores.TextMatrix(i - 1, 0) <> "TAR") Then
                    sql = sql & "," & XN(grdPagos.TextMatrix(i - 1, 2)) 'F PAGO
                    sql = sql & "," & XN(grdPagos.TextMatrix(i - 1, 1))     '$ PAGO
                    sql = sql & "," & XN(grdPagos.TextMatrix(i - 1, 3)) 'TAR_CODIGO
                    sql = sql & "," & XN(grdPagos.TextMatrix(i - 1, 5)) 'TAR_PLAN
                    sql = sql & "," & XN(grdPagos.TextMatrix(i - 1, 7)) 'TAR_CUPON
                    sql = sql & "," & XN(grdPagos.TextMatrix(i - 1, 8)) 'TAR_LOTE
                    sql = sql & "," & XN(grdPagos.TextMatrix(i - 1, 9)) & ")" 'TAR_AUTORIZACION
                Else
                    sql = sql & "," & XN(grdPagos.TextMatrix(i, 2))   'F PAGO
                    sql = sql & "," & XN(grdPagos.TextMatrix(i, 1))       '$ PAGO
                    sql = sql & "," & XN(grdPagos.TextMatrix(i, 3)) 'TAR_CODIGO
                    sql = sql & "," & XN(grdPagos.TextMatrix(i, 5)) 'TAR_PLAN
                    sql = sql & "," & XN(grdPagos.TextMatrix(i, 7)) 'TAR_CUPON
                    sql = sql & "," & XN(grdPagos.TextMatrix(i, 8)) 'TAR_LOTE
                    sql = sql & "," & XN(grdPagos.TextMatrix(i, 9)) & ")" 'TAR_AUTORIZACION
                End If
            Else
                sql = sql & ",NULL,NULL,NULL,NULL,NULL,NULL,NULL)"
            End If
            
            DBConn.Execute sql
        Next
        
        'FACTURAS Y NOTA DE DEBITO CANCELADAS EN EL RECIBO
        For i = 1 To GrillaAplicar.Rows - 1
           If CDbl(txtpagTar.Text) > 0 Then
                sql = "INSERT INTO FACTURAS_RECIBO_CLIENTE"
                sql = sql & " (TCO_CODIGO, REC_NUMERO, REC_SUCURSAL, REC_FECHA,"
                sql = sql & " FCL_TCO_CODIGO, FCL_NUMERO, FCL_SUCURSAL,"
                sql = sql & " FCL_FECHA,REC_IMPORTE,REC_ABONA,REC_SALDO)"
                sql = sql & " VALUES ("
                sql = sql & XN(cboRecibo.ItemData(cboRecibo.ListIndex)) & ","
                sql = sql & XN(txtNroRecibo.Text) & ","
                sql = sql & XN(txtNroSucursal.Text) & ","
                sql = sql & XDQ(FechaRecibo) & ","
                sql = sql & XN(GrillaAplicar.TextMatrix(i, 6)) & ","           'TIPO FACTURA O NOTA DEBITO
                sql = sql & XN(Right(GrillaAplicar.TextMatrix(i, 1), 8)) & "," 'NUMERO FACTURA O NOTA DEBITO
                sql = sql & XN(Left(GrillaAplicar.TextMatrix(i, 1), 4)) & ","  'NUMERO SUCURSAL
                sql = sql & XDQ(GrillaAplicar.TextMatrix(i, 2)) & ","          'FECHA FACTURA O NOTA DEBITO
                
                'Comparo para ver si me queda saldo
                If CDbl(txtpagTar.Text) > VALIDO_IMPORTE(GrillaAplicar.TextMatrix(i, 5)) Then
                   'Importe TOTAL de la Factura
                   txtpagTar.Text = CDbl(txtpagTar.Text) - _
                                           VALIDO_IMPORTE(GrillaAplicar.TextMatrix(i, 5))
                   GrillaAplicar.TextMatrix(i, 4) = GrillaAplicar.TextMatrix(i, 5)
                   GrillaAplicar.TextMatrix(i, 5) = "0,00"
                   
                Else
                   'Importe del SALDO
                   GrillaAplicar.TextMatrix(i, 4) = txtpagTar.Text
                   GrillaAplicar.TextMatrix(i, 5) = Format(CDbl(GrillaAplicar.TextMatrix(i, 5)) - CDbl(GrillaAplicar.TextMatrix(i, 4)), "0.00")
                   txtpagTar.Text = "0,00"
                End If
                
                sql = sql & XN(GrillaAplicar.TextMatrix(i, 3)) & _
                     ", " & XN(GrillaAplicar.TextMatrix(i, 4)) & _
                     ", " & XN(GrillaAplicar.TextMatrix(i, 5)) & ")"
                DBConn.Execute sql
           End If
        Next
        
        'ACTUALIZO EL SALDO DE LAS FACTURAS ELEGIDAS
        For i = 1 To GrillaAplicar.Rows - 1
           If Trim(GrillaAplicar.TextMatrix(i, 4)) <> "0,00" Then
                sql = "UPDATE FACTURA_CLIENTE"
                sql = sql & " SET FCL_SALDO = " & XN(GrillaAplicar.TextMatrix(i, 5))
                sql = sql & " WHERE TCO_CODIGO=" & XN(GrillaAplicar.TextMatrix(i, 6))
                sql = sql & "   AND FCL_NUMERO=" & XN(Right(GrillaAplicar.TextMatrix(i, 1), 8))  'NUMERO FACTURA
                sql = sql & "   AND FCL_SUCURSAL=" & XN(Left(GrillaAplicar.TextMatrix(i, 1), 4)) 'NUMERO SUCURSAL
                DBConn.Execute sql
           End If
        Next

        'ACTUALIZO EL DINERO A CUENTA (RECIBO_CLIENTE_SALDO)
        For i = 1 To GrillaAFavor.Rows - 1
            If GrillaAFavor.TextMatrix(i, 5) <> "19" Then '19 ANTICIPO DE COBRO
                sql = "UPDATE RECIBO_CLIENTE_SALDO"
                sql = sql & " SET REC_SALDO = " & XN(GrillaAFavor.TextMatrix(i, 4))
                sql = sql & " WHERE TCO_CODIGO = " & XN(GrillaAFavor.TextMatrix(i, 5))
                sql = sql & "   AND REC_NUMERO = " & XN(Right(GrillaAFavor.TextMatrix(i, 1), 8)) 'NUMERO RECIBO
                sql = sql & "   AND REC_SUCURSAL = " & XN(Left(GrillaAFavor.TextMatrix(i, 1), 4)) 'NUMERO SUCURSAL
                DBConn.Execute sql
            Else
                sql = "UPDATE ANTICIPO_COBRO"
                sql = sql & " SET ANC_SALDO = " & XN(GrillaAFavor.TextMatrix(i, 4))
                sql = sql & " WHERE ANC_NUMERO = " & XN(Right(GrillaAFavor.TextMatrix(i, 1), 8)) 'NUMERO ANTICIPO
                sql = sql & "   AND ANC_SUCURSAL = " & XN(Left(GrillaAFavor.TextMatrix(i, 1), 4)) 'NUMERO SUCURSAL
                sql = sql & "   AND ANC_FECHA = " & XDQ(GrillaAFavor.TextMatrix(i, 2))
                DBConn.Execute sql
            End If
        Next

        'VERIFICO SI HAY DINERO A CUENTA
        If CDbl(txtpagTar.Text) > 0 Then
            sql = "INSERT INTO RECIBO_CLIENTE_SALDO"
            sql = sql & " (TCO_CODIGO, REC_NUMERO, REC_SUCURSAL, REC_FECHA,"
            sql = sql & " REC_TOTSALDO, REC_SALDO)"
            sql = sql & " VALUES ("
            sql = sql & XN(cboRecibo.ItemData(cboRecibo.ListIndex)) & ","
            sql = sql & XN(txtNroRecibo.Text) & ","
            sql = sql & XN(txtNroSucursal.Text) & ","
            sql = sql & XDQ(FechaRecibo.Value) & ","
            sql = sql & XN(CDbl(txtTotalValores.Text)) & ","
            sql = sql & XN(CDbl(txtpagTar.Text)) & ")"
            DBConn.Execute sql
        End If
                                                
        'ACTUALIZO LA TABLA PARAMENTROS Y LE SUMO UNO AL RECIBO QUE CORRESPONDA
        Select Case cboRecibo.ItemData(cboRecibo.ListIndex)
            Case 12
                sql = "UPDATE PARAMETROS SET RECIBO_C=" & XN(txtNroRecibo)
                DBConn.Execute sql
        End Select
        
        DBConn.CommitTrans
        mBorroTransfe = False
        
        
        'Aca puedo reutilizar esta parte
        'tengo que preguntarle a Marcos por el rptrecibo.rpt
        'sino veo que me conviene, tal vez hacer una impresion con formulario
        'preimpreso
        
        'If MsgBox("Desea Imprimir el Recibo?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        '    'MANDO RECIBO A IMPRESORA
        '    mImprimoRecibo = False
        '    cmdImprimir_Click
        'End If
    Else 'SI EXISTE
        MsgBox "El Recibo ya fue Registrado", vbCritical, TIT_MSGBOX
        'DBConn.CommitTrans
        mBorroTransfe = True
    End If
    Screen.MousePointer = vbNormal
    lblestado.Caption = ""
    rec.Close
    cmdImprimir_Click
    CmdNuevo_Click
    Exit Sub
    
HayError:
    lblestado.Caption = ""
    Screen.MousePointer = vbNormal
    If rec.State = 1 Then rec.Close
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Function ValidarRecibo() As Boolean
    
    If txtNroSucursal.Text = "" Or txtNroRecibo.Text = "" Then
        MsgBox "Debe ingresar el número de Recibo", vbCritical, TIT_MSGBOX
        txtNroSucursal.SetFocus
        ValidarRecibo = False
        Exit Function
    End If
    If IsNull(FechaRecibo.Value) Then
        MsgBox "Debe ingresar la fecha del Recibo", vbCritical, TIT_MSGBOX
        FechaRecibo.SetFocus
        ValidarRecibo = False
        Exit Function
    End If
    If txtCodCliente.Text = "" Then
        MsgBox "Debe ingresar un Cliente", vbCritical, TIT_MSGBOX
        txtCodCliente.SetFocus
        ValidarRecibo = False
        Exit Function
    End If
    If grillaValores.Rows = 1 Then
        MsgBox "Debe ingresar Valores Recibidos", vbCritical, TIT_MSGBOX
        ValidarRecibo = False
        Exit Function
    End If
    If GrillaAplicar.Rows = 1 Then
        MsgBox "No tiene Facturas pendientes", vbCritical, TIT_MSGBOX
        'cmdAgregarFactura.SetFocus
        ValidarRecibo = False
        Exit Function
    End If
    'If CDbl(txtSaldo.Text) > CDbl(txtTotalValores.Text) Then
    '    MsgBox "El Total de Facturas supera al Total de Valores Recibidos", vbCritical, TIT_MSGBOX
    '    ValidarRecibo = False
    '    Exit Function
    'End If
    If CDbl(txtImporteApagar.Text) < CDbl(txtTotalValores.Text) Then
        If MsgBox("El Total de Valores Recibidos supera al Total de Facturas," & Chr(13) & _
                "deja el importe (" & Format(CStr(CDbl(txtTotalValores.Text) - CDbl(txtImporteApagar.Text)), "#,##0.00") & _
                ") como dinero a cuenta?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then

            'cmdAgregarFactura.SetFocus
            ValidarRecibo = False
            Exit Function
        End If
    End If
    ValidarRecibo = True
End Function

Private Sub CmdNuevo_Click()
    Estado = 1
    If mBorroTransfe = True Then
       'VERIFICO SI HAY UNA TRASFERENCIA CARGADA
       'SI HAY LA BORRO DE LA TABLA DEBCRE_BANCARIOS
       For i = 1 To grillaValores.Rows - 1
           If grillaValores.TextMatrix(i, 0) = "COMP" Then
               If grillaValores.TextMatrix(i, 5) = 30 Then
                   DBConn.Execute "DELETE FROM DEBCRE_BANCARIOS WHERE DCB_NUMERO = " & XN(Right(Trim(grillaValores.TextMatrix(i, 4)), 8))
               End If
           End If
       Next
    End If
    cmdGrabar.Enabled = True
    FrameRecibo.Enabled = True
    FrameRemito.Enabled = True
    TxtCheNumero.Text = ""
    GrillaCheques.Rows = 1
    GrillaCheques.HighLight = flexHighlightNever
    txtEftImporte.Text = ""
    GrillaEfectivo.Rows = 1
    GrillaEfectivo.HighLight = flexHighlightNever
    GrillaAplicar.Rows = 1
    GrillaAplicar.HighLight = flexHighlightNever
    
    GrillaAFavor.Rows = 1
    GrillaAFavor.HighLight = flexHighlightNever
    LblDineroaCta.Caption = ""
    
    grillaValores.Rows = 1
    grillaValores.HighLight = flexHighlightNever
    
    txtCodCliente.Text = ""
    txtNroRecibo.Text = ""
    txtNroSucursal.Text = ""
    txtSaldo.Text = ""
    txtImporteApagar.Text = ""
    'FechaRendicion.Text = Date
    cboRecibo.ListIndex = 0
    'txtTotalCheques.Text = ""
    txtTotalEfectivo.Text = ""
    txtTotalValores.Text = "0.00"
    
    'txtTotalComprobante.Text = ""
    tabValores.Tab = 0
    tabComprobantes.Tab = 0
    
    'MANDO RECIBO A PANTALLA
    mImprimoRecibo = True
    
    'CARGO ESTADO
    Call BuscoEstado(1, lblEstadoRecibo) 'ESTADO PENDIENTE
    tabDatos.Tab = 0
    cboRecibo.SetFocus
    txtNroSucursal_LostFocus
    txtNroRecibo_LostFocus
    FechaRecibo_LostFocus
    txtCodCliente.SetFocus
    
    cmdImprimir.Visible = False
    grdPagos.Rows = 1
    txtTotalPagos.Text = ""
    cboFormaPago.ListIndex = 0
    txtImportePago.Text = ""
    txtSaldoEftTar.Text = ""
    txtpagTar.Text = ""
    txtSaldoActual.Text = ""
    FechaRecibo.Value = Date
    txtObservaciones.Text = ""
    
    
End Sub

'Private Sub cmdNuevoCheque_Click()
'    FrmCargaCheques.Show vbModal
'    'TxtCheNumero.SetFocus
'End Sub
'
'Private Sub cmdQuitarVal_Click()
'
'End Sub

Private Sub QuitoDineroACta()
    For i = 1 To GrillaAFavor.Rows - 1
        If GrillaAFavor.TextMatrix(i, 5) = grillaValores.TextMatrix(grillaValores.RowSel, 5) _
            And CLng(GrillaAFavor.TextMatrix(i, 1)) = CLng(grillaValores.TextMatrix(grillaValores.RowSel, 4)) _
            And CDate(GrillaAFavor.TextMatrix(i, 2)) = CDate(grillaValores.TextMatrix(grillaValores.RowSel, 2)) Then
            
            'ARREGLO EL SALDO DEL DINERO A CTA
            GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 4) = VALIDO_IMPORTE(CStr(CDbl(GrillaAFavor.TextMatrix(i, 4)) + CDbl(grillaValores.TextMatrix(grillaValores.RowSel, 1))))
           Exit For
        End If
    Next
End Sub

Private Sub cmdNuevoCheque_Click()
    FrmCargaCheques.Show vbModal
    TxtCheNumero.SetFocus
    If TxtCheNumero.Text <> "" Then
        TxtCheNumero_LostFocus
        cmdAgregarCheque_Click
    End If
End Sub
Private Sub CmdSalir_Click()
    'If MsgBox("Seguro que desea Salir", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        Set frmReciboCliente = Nothing
        Unload Me
    'End If
End Sub

Private Sub FechaRecibo_LostFocus()
    If IsNull(FechaRecibo.Value) Then
        FechaRecibo.Value = Date
    End If
End Sub

Private Sub Form_Activate()
    'txtCodCliente.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 And _
       ActiveControl.Name <> "txtCodCliente" And _
       ActiveControl.Name <> "txtCliRazSoc" And _
       ActiveControl.Name <> "txtCliente" And _
       ActiveControl.Name <> "txtDesCli" Then tabDatos.Tab = 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then MySendKeys Chr(9)
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Set Rec2 = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
   
    Me.Left = 0
    Me.Top = 0
    
    tabDatos.Tab = 0
    tabValores.Tab = 0
    tabComprobantes.Tab = 0
    Centrar_pantalla Me
    
    'CONFIGURO GRILLAS
    configurogrillas
    
    'CARGO COMBO CON LOS TIPOS DE RECIBO
    LlenarComboRecibo
    
    'CARGO COMBO CON LAS PROVINCIAS
    LLenarComboMoneda
    
    'CARGO COMBO CON COMPROBANTES PARA USO DE PAGO
    'Call CargoComboBox(cboComprobantes, "TIPO_COMPROBANTE", "TCO_CODIGO", "TCO_DESCRI")
    'cboComprobantes.ListIndex = 0

    'CARGO ESTADO
    Call BuscoEstado(1, lblEstadoRecibo) 'ESTADO PENDIENTE
    Estado = 1
    '------------------------
    'frameBanco.Enabled = False
    'cmdAgregarCheque.Enabled = False
    'cmdAgregarEfectivo.Enabled = False
    'FechaRendicion.Text = Date
    'Llenar Cbo Forma de Pago
    LLenarFPago
    
    txtNroRecibo.Enabled = True
    lblestado.Caption = ""
    mBorroTransfe = False
    
    'MANDO RECIBO A PANTALLA
    mImprimoRecibo = True
    
    txtNroSucursal_LostFocus
    txtNroRecibo_LostFocus
    FechaRecibo.Value = Date
    FechaRecibo_LostFocus
End Sub

Private Sub configurogrillas()
    'GRILLA CHEQUES
    GrillaCheques.FormatString = "^Bco|^Loc|^Suc|^Cod|^Nro Cheque|^Fec Vto|>Importe|COD INTERNO BANCO|DECRI BANCO"
    GrillaCheques.ColWidth(0) = 500   'BCO
    GrillaCheques.ColWidth(1) = 500   'LOC
    GrillaCheques.ColWidth(2) = 500   'SUC
    GrillaCheques.ColWidth(3) = 700   'COD
    GrillaCheques.ColWidth(4) = 1100  'NRO CHEQUE
    GrillaCheques.ColWidth(5) = 1000  'FEC VTO
    GrillaCheques.ColWidth(6) = 1000  'IMPORTE
    GrillaCheques.ColWidth(7) = 0     'COD INTERNO BANCO
    GrillaCheques.ColWidth(8) = 0     'DESCRI BANCO
    GrillaCheques.Rows = 1
    
    'GRILLA EFECTIVO
    GrillaEfectivo.FormatString = "Moneda|>Importe|Cód.Moneda"
    GrillaEfectivo.ColWidth(0) = 1900 'MONEDA
    GrillaEfectivo.ColWidth(1) = 1000 'IMPORTE
    GrillaEfectivo.ColWidth(2) = 0    'CODIGO MONEDA
    GrillaEfectivo.Rows = 1
    GrillaEfectivo.HighLight = flexHighlightNever
    GrillaEfectivo.BorderStyle = flexBorderNone
    GrillaEfectivo.row = 0
    For i = 0 To GrillaEfectivo.Cols - 1
        GrillaEfectivo.Col = i
        GrillaEfectivo.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        GrillaEfectivo.CellBackColor = &H808080    'GRIS OSCURO
        GrillaEfectivo.CellFontBold = True
    Next

    'GRILLA Aplicar A
    GrillaAplicar.FormatString = "^Comp.|^Número|^Fecha|>Total|>Abona|>Saldo|Cod.Comprob"
    GrillaAplicar.ColWidth(0) = 700  'COMPROBANTE
    GrillaAplicar.ColWidth(1) = 1250 'NUMERO
    GrillaAplicar.ColWidth(2) = 1000 'FECHA
    GrillaAplicar.ColWidth(3) = 900  'TOTAL
    GrillaAplicar.ColWidth(4) = 900  'ABONA
    GrillaAplicar.ColWidth(5) = 800  'SALDO
    GrillaAplicar.ColWidth(6) = 0    'CODIGO COMPROBANTE
    GrillaAplicar.Rows = 1
    GrillaAplicar.HighLight = flexHighlightNever
    GrillaAplicar.BorderStyle = flexBorderNone
    GrillaAplicar.row = 0
    For i = 0 To GrillaAplicar.Cols - 1
        GrillaAplicar.Col = i
        GrillaAplicar.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        GrillaAplicar.CellBackColor = &H808080    'GRIS OSCURO
        GrillaAplicar.CellFontBold = True
    Next
    
    'GRILLA BUSQUEDA
    GrdModulos.FormatString = "^Tipo|^Nro Recibo|^Fecha Recibo|Cliente|Tipo Recibo|>Monto"
    GrdModulos.ColWidth(0) = 1000 'TIPO RECIBO
    GrdModulos.ColWidth(1) = 1600 'NRO RECIBO
    GrdModulos.ColWidth(2) = 1600 'FECHA RECIBO
    GrdModulos.ColWidth(3) = 5000 'CLIENTE
    GrdModulos.ColWidth(4) = 0    'TIPO RECIBO (TCO_CODIGO)
    GrdModulos.ColWidth(5) = 1600 'MONTO
    GrdModulos.Cols = 6
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
    
    'GRILLA VALORES
    grillaValores.FormatString = "^Tipo|>Importe||Descripción|Número||"
    grillaValores.ColWidth(0) = 700  'TIPO DE VALOR (CHE,EFT...)
    grillaValores.ColWidth(1) = 1000 'IMPORTE
    grillaValores.ColWidth(2) = 0    'FECHA
    grillaValores.ColWidth(3) = 2000 'DESCRIPCIÓN
    grillaValores.ColWidth(4) = 1350 'NÚMERO
    grillaValores.ColWidth(5) = 0    'CÓDIGO
    grillaValores.ColWidth(6) = 0    'REPRESENTADA
    grillaValores.Rows = 1
    grillaValores.HighLight = flexHighlightNever
    grillaValores.BorderStyle = flexBorderNone
    grillaValores.row = 0
    For i = 0 To grillaValores.Cols - 1
        grillaValores.Col = i
        grillaValores.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        grillaValores.CellBackColor = &H808080    'GRIS OSCURO
        grillaValores.CellFontBold = True
    Next
    
    'GRILLA A FAVOR
    GrillaAFavor.FormatString = "^Comp.|^Número|^Fecha|>Total|>Saldo|codigo comprobante|REPRESENTADA"
    GrillaAFavor.ColWidth(0) = 850  'COMPROBANTE
    GrillaAFavor.ColWidth(1) = 1300 'NUMERO
    GrillaAFavor.ColWidth(2) = 1000 'FECHA
    GrillaAFavor.ColWidth(3) = 1000 'TOTAL
    GrillaAFavor.ColWidth(4) = 1000 'SALDO
    GrillaAFavor.ColWidth(5) = 0    'CODIGO COMPROBANTE
    GrillaAFavor.ColWidth(6) = 0    'REPRESENTADA
    GrillaAFavor.Rows = 1
    GrillaAFavor.HighLight = flexHighlightNever
    GrillaAFavor.BorderStyle = flexBorderNone
    GrillaAFavor.row = 0
    For i = 0 To GrillaAFavor.Cols - 1
        GrillaAFavor.Col = i
        GrillaAFavor.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        GrillaAFavor.CellBackColor = &H808080    'GRIS OSCURO
        GrillaAFavor.CellFontBold = True
    Next
    
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
    'fraPagos.Visible = False
    'fraTarjeta.Visible = False
End Sub

Private Sub LlenarComboRecibo()
    sql = "SELECT * FROM TIPO_COMPROBANTE"
    sql = sql & " WHERE TCO_DESCRI LIKE 'RECIBO C%'"
    sql = sql & " ORDER BY TCO_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboRecibo1.AddItem "(Todos)"
        Do While rec.EOF = False
            cboRecibo.AddItem rec!TCO_DESCRI
            cboRecibo.ItemData(cboRecibo.NewIndex) = rec!TCO_CODIGO
            cboRecibo1.AddItem rec!TCO_DESCRI
            cboRecibo1.ItemData(cboRecibo1.NewIndex) = rec!TCO_CODIGO
            rec.MoveNext
        Loop
        cboRecibo.ListIndex = 0
        cboRecibo1.ListIndex = 0
    End If
    rec.Close
End Sub
Private Sub LLenarFPago()
        'cargofromadePago
    'CargoComboBox cboFormaPago, "FORMA_PAGO", "FPG_CODIGO", "FPG_DESCRI", "FPG_DESCRI"
    'If cboFormaPago.ListCount > 0 Then cboFormaPago.ListIndex = 0
    '" & txtCliRazSoc.Text & "%'"
    sql = "SELECT * FROM FORMA_PAGO WHERE FPG_DESCRI LIKE 'TARJETA%' ORDER BY FPG_DESCRI"
    cboFormaPago.Clear
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboFormaPago.AddItem rec!FPG_DESCRI
            cboFormaPago.ItemData(cboFormaPago.NewIndex) = rec!FPG_CODIGO
            rec.MoveNext
        Loop
        cboFormaPago.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub LLenarComboMoneda()
    sql = "SELECT * FROM MONEDA ORDER BY MON_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboMoneda.AddItem rec!MON_DESCRI
            cboMoneda.ItemData(cboMoneda.NewIndex) = rec!MON_CODIGO
            rec.MoveNext
        Loop
        cboMoneda.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub GrdModulos_dblClick()
     If GrdModulos.Rows > 1 Then
        mBorroTransfe = False
        CmdNuevo_Click
        Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 4)), cboRecibo)
        txtNroRecibo.Text = Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 8)
        txtNroSucursal.Text = Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4)
        FechaRecibo.Value = GrdModulos.TextMatrix(GrdModulos.RowSel, 2)
        tabDatos.Tab = 0
        Call BuscarRecibo(GrdModulos.TextMatrix(GrdModulos.RowSel, 4), Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 8), Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4))
        cmdImprimir.Visible = True
     End If
End Sub

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then GrdModulos_dblClick
End Sub

Private Sub GrillaAFavor_DblClick()
    If GrillaAFavor.Rows > 1 Then
        txtSaldoACta.Text = VALIDO_IMPORTE(GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 4))
        txtImporteACta.Text = VALIDO_IMPORTE(GrillaAFavor.TextMatrix(GrillaAFavor.RowSel, 4))
        txtImporteACta.SetFocus
    End If
End Sub

Private Sub GrillaAFavor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If GrillaAFavor.Rows > 1 Then
           GrillaAFavor_DblClick
        End If
    End If
End Sub

Private Sub GrillaEfectivo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If GrillaEfectivo.Rows > 2 Then
           GrillaEfectivo.RemoveItem GrillaEfectivo.RowSel
        Else
           GrillaEfectivo.Rows = 1
           GrillaEfectivo.HighLight = flexHighlightNever
           cboMoneda.SetFocus
        End If
        txtTotalEfectivo.Text = SumaGrilla(GrillaEfectivo, 1)
        'txtTotalValores.Text = VALIDO_IMPORTE(CStr(CDbl(SumaGrilla(GrillaCheques, 6)) + CDbl(SumaGrilla(GrillaEfectivo, 1))))
    End If
End Sub

Private Sub grillaValores_DblClick()
    If grillaValores.Rows > 1 Then
        If MsgBox("¿Seguro que desea eliminar?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
            If grillaValores.Rows > 2 Then
                If grillaValores.TextMatrix(grillaValores.RowSel, 0) = "A-CTA" Then
                    QuitoDineroACta
                ElseIf grillaValores.TextMatrix(grillaValores.RowSel, 0) = "COMP" Then
                    'VEO SI ES UNA TRASFERENCIA
                    'SI ES LA BORRO DE LA TABLA DEBCRE_BANCARIOS
                    If grillaValores.TextMatrix(grillaValores.RowSel, 5) = 30 Then
                        DBConn.Execute "DELETE FROM DEBCRE_BANCARIOS WHERE DCB_NUMERO = " & XN(Right(Trim(grillaValores.TextMatrix(grillaValores.RowSel, 4)), 8))
                    End If
                End If
                grillaValores.RemoveItem grillaValores.RowSel
                txtTotalValores.Text = SumaGrilla(grillaValores, 1)
            Else
                If grillaValores.TextMatrix(grillaValores.RowSel, 0) = "A-CTA" Then
                    QuitoDineroACta
                ElseIf grillaValores.TextMatrix(grillaValores.RowSel, 0) = "COMP" Then
                    'VEO SI ES UNA TRASFERENCIA
                    'SI ES LA BORRO DE LA TABLA DEBCRE_BANCARIOS
                    If grillaValores.TextMatrix(grillaValores.RowSel, 5) = 30 Then
                        DBConn.Execute "DELETE FROM DEBCRE_BANCARIOS WHERE DCB_NUMERO = " & XN(Right(Trim(grillaValores.TextMatrix(grillaValores.RowSel, 4)), 8))
                    End If
                End If
                grillaValores.Rows = 1
                txtTotalValores.Text = ""
                grillaValores.HighLight = flexHighlightNever
            End If
        End If
    End If
End Sub

Private Sub tabComprobantes_Click(PreviousTab As Integer)
    If tabComprobantes.Tab = 1 Then
        GrillaAplicar.SetFocus
    End If
    If tabComprobantes.Tab = 0 Then
        'If Me.tabComprobantes.Visible = True Then cmdAgregarFactura.SetFocus
        If GrillaAplicar.Rows > 1 Then
          ' txtTotalAplicar.Text = VALIDO_IMPORTE(SumaGrilla(GrillaAplicar, 1))
        End If
    End If
End Sub

Private Sub tabDatos_Click(PreviousTab As Integer)
  If tabDatos.Tab = 1 Then
    'LimpiarBusqueda
    cmdGrabar.Enabled = False
    If Me.Visible = True Then txtCliente.SetFocus
  End If
End Sub

Private Sub LimpiarBusqueda()
    txtCliente.Text = ""
    txtDesCli.Text = ""
    FechaDesde.Value = ""
    FechaHasta.Value = ""
    GrdModulos.Rows = 1
    GrdModulos.HighLight = flexHighlightNever
   ' cboBuscaRep.ListIndex = -1
    cboRecibo1.ListIndex = 0
   ' chkRepresentada.Value = Unchecked
End Sub

Private Sub tabValores_Click(PreviousTab As Integer)
    If tabValores.Tab = 1 Then
       BuscaProx "PESOS", cboMoneda
       txtEftImporte.SetFocus
    ElseIf tabValores.Tab = 2 Then
            If GrillaAFavor.Rows > 1 Then
               GrillaAFavor.Col = 0
               GrillaAFavor.row = 1
               GrillaAFavor.SetFocus
            End If
    ElseIf tabValores.Tab = 3 Then
        txtTotalPagos.Text = txtImporteApagar.Text
    'ElseIf tabValores.Tab = 4 Then
        cboFormaPago.SetFocus
        
        If txtSaldoEftTar.Text <> "" Then
            txtTotalPagos.Text = CDbl(txtTotalPagos.Text) - (CDbl(txtImporteApagar.Text) - CDbl(txtSaldoEftTar.Text))
            txtTotalPagos.Text = VALIDO_IMPORTE(txtTotalPagos.Text)
        End If
    End If
End Sub

Private Sub TxtBANCO_GotFocus()
    SelecTexto TxtBANCO
End Sub

Private Sub TxtBANCO_LostFocus()
    If Len(TxtBANCO.Text) < 3 Then TxtBANCO.Text = CompletarConCeros(TxtBANCO.Text, 3)
End Sub

Private Sub TxtCheNumero_Change()
    If TxtCheNumero.Text = "" Then
        LimpiarCheques
    Else
        frameBanco.Enabled = True
        cmdAgregarCheque.Enabled = True
    End If
End Sub

Private Sub TxtCheNumero_GotFocus()
    SelecTexto TxtCheNumero
End Sub

Private Sub TxtCheNumero_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtCheNumero_LostFocus()
    If TxtCheNumero.Text <> "" Then
        If Len(TxtCheNumero.Text) < 8 Then TxtCheNumero.Text = CompletarConCeros(TxtCheNumero.Text, 8)
    'sql = "SELECT * FROM CHEQUE WHERE "
        sql = "SELECT DISTINCT CE.CHE_NUMERO, CH.CHE_IMPORT, CH.CHE_FECVTO, CE.BAN_CODINT, B.BAN_BANCO, B.BAN_LOCALIDAD,"
        sql = sql & " B.BAN_SUCURSAL, B.BAN_CODIGO, B.BAN_NOMCOR,CE.CES_DESCRI,B.BAN_DESCRI"
        sql = sql & " FROM CHEQUE_ESTADOS CE, CHEQUE CH, BANCO B,ESTADO_CHEQUE E"
        sql = sql & " Where "
        sql = sql & " CE.CHE_NUMERO = CH.CHE_NUMERO And "
        sql = sql & " CE.BAN_CODINT = CH.BAN_CODINT And "
        sql = sql & " CH.BAN_CODINT=B.BAN_CODINT  "
        'sql = sql & " CE.ECH_CODIGO= E.ECH_CODIGO AND" '
        'sql = sql & " E.ECH_CODIGO=7" ' 7-entregado
        sql = sql & " AND CH.CHE_NUMERO LIKE '%" & Trim(TxtCheNumero) & "%'"  'CODIGO (1) ES CHEQUE EN CARTERA
        
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            TxtCheNumero.Text = rec!CHE_NUMERO
            TxtBANCO.Text = rec!BAN_BANCO
            TxtLOCALIDAD.Text = rec!BAN_LOCALIDAD
            TxtSUCURSAL.Text = rec!BAN_SUCURSAL
            TxtCODIGO.Text = rec!BAN_CODIGO
            TxtCheImport.Text = rec!che_import
            TxtCheFecVto.Value = rec!CHE_FECVTO
            TxtBanDescri.Text = rec!BAN_DESCRI
            TxtCodInt.Text = rec!BAN_CODINT
        End If
        rec.Close
    End If
    
End Sub

Private Sub txtCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarClientes "txtCliente", "CODIGO"
    End If
End Sub

Private Sub txtCliRazSoc_Change()
    If txtCliRazSoc.Text = "" Then
        txtCodCliente.Text = ""
        txtDomici.Text = ""
    End If
End Sub

Private Sub txtCliRazSoc_GotFocus()
    SelecTexto txtCliRazSoc
End Sub

Private Sub txtCliRazSoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarClientes "txtCodCliente", "CODIGO"
    End If
End Sub

Private Sub txtCliRazSoc_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtCliRazSoc_LostFocus()
    If txtCodCliente.Text = "" And txtCliRazSoc.Text <> "" Then
        sql = "SELECT C.CLI_CODIGO,C.CLI_RAZSOC,C.CLI_DOMICI"
        sql = sql & "  FROM CLIENTE C "
        sql = sql & " WHERE C.CLI_RAZSOC LIKE '" & txtCliRazSoc.Text & "%'"
        If rec.State = 1 Then rec.Close
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If rec.RecordCount > 1 Then
                BuscarClientes "txtCodCliente", "CADENA", Trim(txtCliRazSoc.Text)
                If rec.State = 1 Then rec.Close
                txtCliRazSoc.SetFocus
            Else
                txtCodCliente.Text = rec!CLI_CODIGO
                txtCliRazSoc.Text = rec!CLI_RAZSOC
                txtDomici.Text = ChkNull(rec!CLI_DOMICI)
                If Estado = 1 Then
                    If BuscarFactura(txtCodCliente) = False Then
                        MsgBox "El Cliente NO tiene facturas pendientes de pago. Verifique!", vbExclamation, TIT_MSGBOX
                        txtCodCliente.Text = ""
                        txtCodCliente.SetFocus
                        FrameRecibo.Enabled = True
                    Else
                        'Call BuscarSaldosAFavor(txtCodCliente)
                        FrameRecibo.Enabled = False
                        txtImporteApagar.SetFocus
                        'GrillaAplicar.SetFocus
                    End If
                End If
            End If
        Else
            MsgBox "No se encontro el Cliente", vbExclamation, TIT_MSGBOX
            txtCodCliente.SetFocus
        End If
        If rec.State = 1 Then rec.Close
    End If
End Sub

Private Sub txtCliente_Change()
    If txtCliente.Text = "" Then
        txtDesCli.Text = ""
    End If
End Sub

Private Sub txtCliente_GotFocus()
    SelecTexto txtCliente
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCliente_LostFocus()
    If txtCliente.Text <> "" Then
        Set rec = New ADODB.Recordset
        sql = "SELECT CLI_RAZSOC FROM CLIENTE"
        sql = sql & " WHERE CLI_CODIGO=" & XN(txtCliente)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtDesCli.Text = rec!CLI_RAZSOC
        Else
            MsgBox "El Cliente no existe", vbExclamation, TIT_MSGBOX
            txtDesCli.Text = ""
            txtCliente.SetFocus
        End If
        If rec.State = 1 Then rec.Close
    End If
End Sub

Private Sub cmdBuscarCli_Click()
    frmBuscar.TipoBusqueda = 1
    frmBuscar.TxtDescriB = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 0
        txtCliente.Text = frmBuscar.grdBuscar.Text
        txtCliente.SetFocus
        txtCliente_LostFocus
    Else
        txtCliente.SetFocus
    End If
End Sub

Private Sub txtCodCliente_Change()
    If txtCodCliente.Text = "" Then
        txtCliRazSoc.Text = ""
        txtDomici.Text = ""
    End If
End Sub

Private Sub txtCodCliente_GotFocus()
    SelecTexto txtCodCliente
End Sub

Private Sub txtCodCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarClientes "txtCodCliente", "CODIGO"
    End If
End Sub

Private Sub txtCodCliente_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCodCliente_LostFocus()
    If txtCodCliente.Text <> "" Then
        Set rec = New ADODB.Recordset
        sql = "SELECT C.CLI_RAZSOC,C.CLI_DOMICI,C.CLI_CTACTE"
        sql = sql & " FROM CLIENTE C"
        sql = sql & " WHERE"
        sql = sql & " C.CLI_CODIGO=" & XN(txtCodCliente)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtCliRazSoc.Text = rec!CLI_RAZSOC
            txtDomici.Text = ChkNull(rec!CLI_DOMICI)
            If Estado = 1 And txtCodCliente.Text <> 1 Then
                If BuscarFactura(txtCodCliente) = False Then
                    If MsgBox("El Cliente NO tiene facturas pendientes de pago. Desea ingresar un saldo a favor?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then
                        txtCodCliente.Text = ""
                        txtCodCliente.SetFocus
                        FrameRecibo.Enabled = True
                    Else
                        GrillaAplicar.AddItem "A FAVOR" & Chr(9) & "0001-00000001" _
                                    & Chr(9) & Date & Chr(9) & "0,00" _
                                    & Chr(9) & "0,00" & Chr(9) & "0,00" & Chr(9) & 15
                        'TotalDeuda = CDbl(TotalDeuda) + VALIDO_IMPORTE(Rec1!Saldo)
                    End If
                Else
                    Call BuscarSaldosAFavor(txtCodCliente)
                    FrameRecibo.Enabled = False
                    GrillaAplicar.SetFocus
                    'MsgBox "El Limite de Cta Cte es: " & rec!CLI_CTACTE, vbInformation, TIT_MSGBOX
                End If
            End If
        Else
            MsgBox "El Cliente no existe", vbExclamation, TIT_MSGBOX
            FrameRecibo.Enabled = True
            txtCliRazSoc.Text = ""
            txtCodCliente.SetFocus
        End If
        If rec.State = 1 Then rec.Close
    End If
End Sub

Private Sub TxtCodigo_GotFocus()
    SelecTexto TxtCODIGO
End Sub

Private Sub txtCupon_GotFocus()
    SelecTexto txtCupon
End Sub

Private Sub txtCupon_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
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
        BuscarClientes "txtCliente", "CODIGO"
    End If
End Sub

Private Sub txtDesCli_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtDesCli_LostFocus()
    If txtCliente.Text = "" And txtDesCli.Text <> "" Then
        sql = "SELECT C.CLI_CODIGO,C.CLI_RAZSOC,C.CLI_DOMICI"
        sql = sql & " FROM CLIENTE C"
        sql = sql & " WHERE"
        sql = sql & " C.CLI_RAZSOC LIKE '" & txtDesCli.Text & "%'"
        If rec.State = 1 Then rec.Close
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If rec.RecordCount > 1 Then
                BuscarClientes "txtCliente", "CADENA", Trim(txtDesCli.Text)
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

Private Sub txtEftImporte_Change()
    'If txtEftImporte.Text = "" Then
    '    cmdAgregarEfectivo.Enabled = False
    'Else
    '    cmdAgregarEfectivo.Enabled = True
    'End If
End Sub

Private Sub txtEftImporte_GotFocus()
    
    
    SelecTexto txtEftImporte
End Sub

Private Sub txtEftImporte_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtEftImporte, KeyAscii)
End Sub

Private Sub txtEftImporte_LostFocus()
    If txtEftImporte.Text <> "" Then
        txtEftImporte.Text = VALIDO_IMPORTE(txtEftImporte.Text)
        cmdAgregarEfectivo.Enabled = True
        'cmdAgregarEfectivo.SetFocus
    End If
End Sub

Private Sub txtImporteACta_Change()
    If txtSaldoACta.Text <> "" And txtImporteACta.Text <> "" Then
        cmdAgregarACta.Enabled = True
    Else
        cmdAgregarACta.Enabled = False
    End If
End Sub

Private Sub txtImporteACta_GotFocus()
    SelecTexto txtImporteACta
End Sub

Private Sub txtImporteACta_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtImporteACta, KeyAscii)
End Sub

Private Sub txtImporteACta_LostFocus()
    If txtSaldoACta.Text <> "" Then
        If txtImporteACta.Text = "" Then
            txtImporteACta.Text = txtSaldoACta.Text
        ElseIf CDbl(txtImporteACta.Text) > CDbl(txtSaldoACta.Text) Then
            MsgBox "Importe mayor al Saldo. Verifique!", vbCritical, TIT_MSGBOX
            txtImporteACta.Text = txtSaldoACta.Text
            txtImporteACta.SetFocus
        End If
        txtImporteACta.Text = VALIDO_IMPORTE(txtImporteACta)
    End If
End Sub

Private Sub txtImporteApagar_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtImporteApagar, KeyAscii)
End Sub

Private Sub txtImporteApagar_LostFocus()
  If Me.ActiveControl.Name <> "CmdSalir" Then
    txtImporteApagar.Text = VALIDO_IMPORTE(txtImporteApagar.Text)
    txtpagTar.Text = VALIDO_IMPORTE(txtImporteApagar.Text)
    
    txtEftImporte = txtImporteApagar.Text
    
    If GrillaAFavor.Rows > 1 Then
       tabValores.Tab = 2
       GrillaAFavor.Col = 0
       GrillaAFavor.row = 1
    Else
       tabValores.Tab = 1
       BuscaProx "PESOS", cboMoneda
       txtEftImporte.SetFocus
    End If
    
    'ingreso saldo a favor
    If GrillaAplicar.TextMatrix(1, 0) = "A FAVOR" Then
        GrillaAplicar.TextMatrix(1, 3) = txtImporteApagar.Text
    End If
  End If
End Sub

Private Function BuscarFactura(CodCli As String) As Boolean
        GrillaAplicar.Rows = 1
        Set Rec1 = New ADODB.Recordset
        Dim TotalDeuda As Double
        TotalDeuda = 0
        'BUSCA LAS FACTURAS
        sql = "SELECT FCL_NUMERO AS NUMERO, FCL_SUCURSAL AS SUCURSAL, "
        sql = sql & " FCL_FECHA AS FECHA, FCL_TOTAL AS TOTAL, FCL_SALDO AS SALDO"
        sql = sql & " ,TCO_CODIGO AS TIPO, TCO_ABREVIA AS ABREVIA"
        sql = sql & " FROM SALDO_FACTURAS_CLIENTE_V"
        sql = sql & " WHERE "
        sql = sql & " CLI_CODIGO=" & XN(CodCli)
        sql = sql & " UNION ALL"
        
        'BUSCA LAS NOTA DE DEBITO
        sql = sql & " SELECT NDC_NUMERO AS NUMERO, NDC_SUCURSAL AS SUCURSAL, "
        sql = sql & " NDC_FECHA AS FECHA, NDC_TOTAL AS TOTAL, NDC_SALDO AS SALDO"
        sql = sql & " ,TCO_CODIGO AS TIPO, TCO_ABREVIA AS ABREVIA"
        sql = sql & " FROM SALDO_NOTA_DEBITO_CLIENTE_V"
        sql = sql & " WHERE "
        sql = sql & " CLI_CODIGO=" & XN(CodCli)
        sql = sql & " ORDER BY FECHA , NUMERO ASC"
        
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            Do While Rec1.EOF = False
                If Rec1!Saldo > 0 Then
                    GrillaAplicar.AddItem Rec1!ABREVIA & Chr(9) & Format(Rec1!Sucursal, "0000") & "-" & Format(Rec1!Numero, "00000000") _
                                    & Chr(9) & Rec1!Fecha & Chr(9) & VALIDO_IMPORTE(Rec1!TOTAL) _
                                    & Chr(9) & "0,00" & Chr(9) & VALIDO_IMPORTE(Rec1!Saldo) & Chr(9) & Rec1!Tipo
                    TotalDeuda = CDbl(TotalDeuda) + VALIDO_IMPORTE(Rec1!Saldo)
                End If
                Rec1.MoveNext
            Loop
            GrillaAplicar.HighLight = flexHighlightAlways
            BuscarFactura = True
            txtSaldo.Text = Format(TotalDeuda, "0.00")
        Else
            BuscarFactura = False
        End If
        Rec1.Close
End Function

Private Sub BuscarSaldosAFavor(CodCli As String)
        GrillaAFavor.Rows = 1
        Set Rec1 = New ADODB.Recordset
        sql = "SELECT RS.TCO_CODIGO, RS.REC_NUMERO, RS.REC_SUCURSAL, RS.REC_FECHA,"
        sql = sql & " RS.REC_TOTSALDO,RS.REC_SALDO, T.TCO_ABREVIA"
        sql = sql & " FROM RECIBO_CLIENTE_SALDO RS, RECIBO_CLIENTE R,TIPO_COMPROBANTE T"
        sql = sql & " WHERE RS.TCO_CODIGO = T.TCO_CODIGO"
        sql = sql & "   AND RS.TCO_CODIGO = R.TCO_CODIGO"
        sql = sql & "   AND RS.REC_NUMERO = R.REC_NUMERO"
        sql = sql & "   AND RS.REC_SUCURSAL = R.REC_SUCURSAL"
        sql = sql & "   AND RS.REC_FECHA = R.REC_FECHA"
        sql = sql & "   AND RS.REC_SALDO > 0"
        sql = sql & "   AND R.CLI_CODIGO = " & XN(CodCli)
        sql = sql & " ORDER BY RS.TCO_CODIGO,RS.REC_SUCURSAL,RS.REC_NUMERO, RS.REC_FECHA"
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            GrillaAFavor.HighLight = flexHighlightAlways
            Do While Rec1.EOF = False
               If Rec1!REC_SALDO > 0 Then
                  GrillaAFavor.AddItem Rec1!TCO_ABREVIA & Chr(9) & Format(Rec1!REC_SUCURSAL, "0000") & "-" & Format(Rec1!REC_NUMERO, "00000000") & Chr(9) & Rec1!REC_FECHA & Chr(9) & VALIDO_IMPORTE(Rec1!REC_TOTSALDO) & Chr(9) & VALIDO_IMPORTE(Rec1!REC_SALDO) & Chr(9) & Rec1!TCO_CODIGO
               End If
               Rec1.MoveNext
            Loop
        End If
        Rec1.Close
        
        If GrillaAFavor.Rows > 1 Then
           LblDineroaCta.Caption = "El Cliente tiene Dinero a Cuenta"
        Else
           LblDineroaCta.Caption = ""
        End If
                        
        'BUSCO ANTICIPOS DE COBRO
        'sql = "SELECT A.ANC_NUMERO, A.ANC_FECHA, A.ANC_SUCURSAL,"
        'sql = sql & " A.ANC_MONTO,A.ANC_SALDO"
        'sql = sql & " FROM ANTICIPO_COBRO A, CLIENTE C"
        'sql = sql & " WHERE"
        'sql = sql & " A.CLI_CODIGO=C.CLI_CODIGO"
        'sql = sql & " AND A.ANC_SALDO > 0"
        'sql = sql & " AND A.CLI_CODIGO=" & XN(CodCli)
        'sql = sql & " ORDER BY A.ANC_FECHA,A.ANC_NUMERO"
        'Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        'If Rec1.EOF = False Then
        '    GrillaAFavor.HighLight = flexHighlightAlways
        '    Do While Rec1.EOF = False
        '        If Rec1!ANC_SALDO > 0 Then
        '            GrillaAFavor.AddItem "ANT-COB" & Chr(9) & Format(Rec1!ANC_SUCURSAL, "0000") & "-" & Format(Rec1!ANC_NUMERO, "00000000") _
        '                            & Chr(9) & Rec1!ANC_FECHA & Chr(9) & VALIDO_IMPORTE(Rec1!ANC_MONTO) _
        '                            & Chr(9) & VALIDO_IMPORTE(Rec1!ANC_SALDO) & Chr(9) & "19" 'TIPO DE COMPROBANTE NRO 19
        '        End If
        '        Rec1.MoveNext
        '    Loop
        'End If
        'Rec1.Close
End Sub

Private Function BuscoComprobanteEnRecibo() As Boolean
'    Set Rec2 = New ADODB.Recordset
'
'    sql = "SELECT DR.REC_NUMERO"
'    sql = sql & " FROM DETALLE_RECIBO_CLIENTE DR, RECIBO_CLIENTE RC"
'    sql = sql & " WHERE"
'    sql = sql & " DR.DRE_TCO_CODIGO=" & XN(cboComprobantes.ItemData(cboComprobantes.ListIndex))
'    sql = sql & " AND DR.DRE_COMNUMERO=" & XN(txtNroComprobantes)
'    sql = sql & " AND DR.DRE_COMSUCURSAL=" & XN(txtNroCompSuc)
'    sql = sql & " AND RC.CLI_CODIGO=" & XN(txtCodCliente)
'    sql = sql & " AND DR.REC_NUMERO=RC.REC_NUMERO"
'    sql = sql & " AND DR.REC_SUCURSAL=RC.REC_SUCURSAL"
'    sql = sql & " AND DR.TCO_CODIGO=RC.TCO_CODIGO"
'    sql = sql & " AND RC.EST_CODIGO=3"
'    Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
'
'    If Rec2.EOF = False Then
'        BuscoComprobanteEnRecibo = False
'    Else
'        BuscoComprobanteEnRecibo = True
'    End If
'    Rec2.Close
    
End Function

Private Sub txtImportePago_GotFocus()
    txtImportePago.Text = txtTotalPagos.Text 'txtImporteApagar
    SelecTexto txtImportePago
End Sub

Private Sub txtImportePago_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtImportePago, KeyAscii)
End Sub

Private Sub txtImportePago_LostFocus()
    If txtCodCliente.Text = "1" Then
        If cboFormaPago.ItemData(cboFormaPago.ListIndex) = 2 Then
            'MsgBox "No Puede Seleccionar Cta CTe para este Cliente", vbCritical, TIT_MSGBOX
            'cboFormaPago.SetFocus
            Exit Sub
        End If
    End If
    If frmDatosTarjeta.cboPlan.ListIndex = -1 Then
        'frmDatosTarjeta.cboFormaPago.SetFocus
        Exit Sub
    
    End If
    
    'If fraTarjeta.Visible = True Then Exit Sub
    txtImportePago.Text = Format(txtImportePago.Text, "0.00")
    Dim mTotalPagos As Double
    mTotalPagos = 0
    
    
    
    For i = 1 To grdPagos.Rows - 1
        mTotalPagos = mTotalPagos + CDbl(grdPagos.TextMatrix(i, 1))
    Next
'    If mtotalpagos + CDbl(Chk0(txtImportePago.Text)) > CDbl(txtTotalPagos.Text) Then
'        MsgBox "El Importe Ingresado Exede el Monto de la Compra!", vbInformation, TIT_MSGBOX
'        'txtImportePago.SetFocus
'        Exit Sub
'    Else
        If cboFormaPago.Text = "" Then
            MsgBox "Debe Indicar la Forma de Pago", vbCritical, TIT_MSGBOX
            cboFormaPago.SetFocus
            Exit Sub
        End If
        If CDbl(Chk0(txtImportePago.Text)) > 0 Then
            grdPagos.AddItem ("")
            grdPagos.row = grdPagos.Rows - 1
            grdPagos.TextMatrix(grdPagos.row, 0) = Trim(Mid(cboFormaPago.Text, 1, 30))
            grdPagos.TextMatrix(grdPagos.row, 1) = txtImportePago.Text
            grdPagos.TextMatrix(grdPagos.row, 2) = cboFormaPago.ItemData(cboFormaPago.ListIndex)
            'mFormaPago = cboFormaPago.ItemData(cboFormaPago.ListIndex)
            
            If Trim(UCase(Mid(cboFormaPago.Text, 1, 50))) = "TARJETA DE CREDITO" Then
                grdPagos.TextMatrix(grdPagos.row, 3) = frmDatosTarjeta.cbotarjeta.ItemData(frmDatosTarjeta.cbotarjeta.ListIndex)
                grdPagos.TextMatrix(grdPagos.row, 4) = frmDatosTarjeta.cbotarjeta.List(frmDatosTarjeta.cbotarjeta.ListIndex)
                grdPagos.TextMatrix(grdPagos.row, 5) = frmDatosTarjeta.cboPlan.ItemData(frmDatosTarjeta.cboPlan.ListIndex)
                grdPagos.TextMatrix(grdPagos.row, 6) = frmDatosTarjeta.cboPlan.List(frmDatosTarjeta.cboPlan.ListIndex)
                grdPagos.TextMatrix(grdPagos.row, 7) = frmDatosTarjeta.txtCupon.Text
                grdPagos.TextMatrix(grdPagos.row, 8) = frmDatosTarjeta.txtLote.Text
                grdPagos.TextMatrix(grdPagos.row, 9) = frmDatosTarjeta.txtTar_Autorizacion.Text
            End If
            If Trim(UCase(Mid(cboFormaPago.Text, 1, 50))) = "TARJETA DE DEBITO" Then
                grdPagos.TextMatrix(grdPagos.row, 3) = frmDatosTarjeta.cbotarjeta.ItemData(frmDatosTarjeta.cbotarjeta.ListIndex)
                grdPagos.TextMatrix(grdPagos.row, 4) = frmDatosTarjeta.cbotarjeta.List(frmDatosTarjeta.cbotarjeta.ListIndex) & " DEBITO"
            End If
'            If Trim(UCase(Mid(cboFormaPago.Text, 1, 50))) = "DOLARES" Then
'                grdPagos.TextMatrix(grdPagos.row, 10) = txtTotDolar.Text
'                grdPagos.TextMatrix(grdPagos.row, 11) = txtCotizacion.Text
'            End If
        End If
'    End If
    mTotalPagos = 0
    For i = 1 To grdPagos.Rows - 1
        mTotalPagos = CDbl(mTotalPagos) + CDbl(grdPagos.TextMatrix(i, 1))
    Next
    If txtsaldototal.Text < 0 Then
        txtTotalPagos.Text = Format(CDbl(txtsaldototal.Text) * (-1) - mTotalPagos, "0.00")
    Else
        txtTotalPagos.Text = Format(CDbl(txtImporteApagar.Text) - mTotalPagos, "0.00")
    End If
    txtImportePago.Text = Format(txtTotalPagos.Text, "0.00")
    If Val(txtTotalPagos.Text) = 0 Then
        'cmdAceptarPagos.SetFocus
    Else
        cboFormaPago.ListIndex = 0
        'cboFormaPago.SetFocus
    End If
    txtTar_Autorizacion.Text = ""
    txtLote.Text = ""
    txtCupon.Text = ""
    cboPlan.Clear
    
    If txtImporteApagar.Text > txtTotalValores.Text Then
        txtSaldoEftTar = CDbl(txtImporteApagar.Text) - CDbl(txtTotalValores.Text)
    End If
End Sub

Private Sub TxtLOCALIDAD_GotFocus()
    SelecTexto TxtLOCALIDAD
End Sub

Private Sub Txtlocalidad_KeyPress(KeyAscii As Integer)
     KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtLOCALIDAD_LostFocus()
    If Len(TxtLOCALIDAD.Text) < 3 Then TxtLOCALIDAD.Text = CompletarConCeros(TxtLOCALIDAD.Text, 3)
End Sub

Private Sub txtLote_GotFocus()
    SelecTexto txtLote
End Sub

Private Sub txtLote_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub
Private Sub txtNroRecibo_Change()
    If txtNroRecibo.Text = "" Then
        FechaRecibo.Value = Date
    End If
End Sub

Private Sub txtNroRecibo_GotFocus()
    SelecTexto txtNroRecibo
End Sub

Private Sub txtNroRecibo_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroRecibo_LostFocus()
    If txtNroRecibo.Text = "" Then
        'BUSCO EL NUMERO DE RECIBO QUE CORRESPONDE
        txtNroRecibo.Text = Format(BuscoUltimoRecibo(cboRecibo.ItemData(cboRecibo.ListIndex)), "00000000")
    Else
        If txtNroSucursal.Text = "" Then
            txtNroSucursal.Text = Sucursal
        End If
        txtNroRecibo.Text = Format(txtNroRecibo.Text, "00000000")
        Call BuscarRecibo(CStr(cboRecibo.ItemData(cboRecibo.ListIndex)), _
                          txtNroRecibo, txtNroSucursal)
    End If
End Sub

Private Function BuscoUltimoRecibo(TipoRec As Integer) As String
    'ACA BUSCA EL NUMERO DE REMITO SIGUIENTE AL ULTIMO CARGADO
    sql = "SELECT (RECIBO_C) + 1 AS REC_C"
    sql = sql & " FROM PARAMETROS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Select Case TipoRec
            Case 12
                BuscoUltimoRecibo = IIf(IsNull(rec!REC_C), 1, rec!REC_C)
        End Select
    End If
    rec.Close
End Function

Private Sub BuscarRecibo(TipoRec As String, NroRec As String, NroSuc As String)
    Set Rec2 = New ADODB.Recordset
    
    sql = "SELECT * "
    sql = sql & "  FROM RECIBO_CLIENTE"
    sql = sql & " WHERE TCO_CODIGO = " & XN(TipoRec)
    sql = sql & "   AND REC_NUMERO = " & XN(NroRec)
    sql = sql & "   AND REC_SUCURSAL = " & XN(NroSuc)
    Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec2.EOF = False Then
        If Rec2.RecordCount > 2 Then
            Rec2.Close
            tabDatos.Tab = 1
            Exit Sub
        End If
        'CABEZA DEL RECIDO
        FechaRecibo.Value = Rec2!REC_FECHA
        'FechaRendicion.Text = Rec2!REC_FECHA_RENDICION
        'CARGO ESTADO
        Call BuscoEstado(CInt(Rec2!EST_CODIGO), lblEstadoRecibo)
        Estado = CInt(Rec2!EST_CODIGO)
        txtCodCliente.Text = Rec2!CLI_CODIGO
        txtCodCliente_LostFocus
        
        txtObservaciones.Text = ChkNull(Rec2!REC_OBSER)
        txtSaldoActual.Text = Chk0(Rec2!REC_SACTUAL)
        
        'DETALLE_DEL RECIBO CHEQUES
        Set rec = New ADODB.Recordset
        sql = "SELECT *"
        sql = sql & " FROM DETALLE_RECIBO_CLIENTE"
        sql = sql & " WHERE TCO_CODIGO =" & XN(TipoRec)
        sql = sql & "   AND REC_NUMERO =" & XN(NroRec)
        sql = sql & "   AND REC_SUCURSAL =" & XN(NroSuc)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
        If rec.EOF = False Then
            Do While rec.EOF = False
                If Not IsNull(rec!BAN_CODINT) Then 'BANCO
                    Call BuscarCheque(rec!BAN_CODINT, rec!CHE_NUMERO)
                ElseIf Not IsNull(rec!MON_CODIGO) Then 'MONEDA
                    grillaValores.AddItem "EFT" & Chr(9) & VALIDO_IMPORTE(rec!DRE_MONIMP) _
                                    & Chr(9) & "" & Chr(9) & BuscarMoneda(rec!MON_CODIGO) _
                                    & Chr(9) & "" & Chr(9) & rec!MON_CODIGO
                              
                ElseIf Not IsNull(rec!DRE_TCO_CODIGO) Then 'COMPROBANTE
                    Dim QueEs As String
                    If rec!DRE_TCO_CODIGO >= 10 And rec!DRE_TCO_CODIGO <= 13 Then
                        QueEs = "A-CTA"
                    ElseIf (rec!DRE_TCO_CODIGO = 19) Then
                        QueEs = "A-CTA"
                    Else
                        QueEs = "COMP"
                    End If
                    grillaValores.AddItem QueEs & Chr(9) & VALIDO_IMPORTE(rec!DRE_COMIMP) _
                                    & Chr(9) & rec!DRE_COMFECHA & Chr(9) & BuscarTipoDocAbre(rec!DRE_TCO_CODIGO) _
                                    & Chr(9) & Format(ChkNull(rec!DRE_COMSUCURSAL), "0000") & "-" & Format(rec!DRE_COMNUMERO, "00000000") _
                                    & Chr(9) & rec!DRE_TCO_CODIGO
                Else
                    If Not IsNull(rec!FPG_CODIGO) Then
                        grillaValores.AddItem "TAR" & Chr(9) & VALIDO_IMPORTE(rec!PAG_IMPORTE) _
                                    & Chr(9) & "" & Chr(9) & BuscarTarjeta(rec!TAR_CODIGO) _
                                    & Chr(9) & rec!TAR_CUPON & Chr(9) & ""
                    End If
                End If
                rec.MoveNext
            Loop
            
            grillaValores.HighLight = flexHighlightAlways
            txtTotalValores.Text = SumaGrilla(grillaValores, 1)
            
        End If
        rec.Close
                   
        'DETALLE_DEL RECIBO FACTURA
        sql = "SELECT * "
        sql = sql & " FROM FACTURAS_RECIBO_CLIENTE"
        sql = sql & " WHERE TCO_CODIGO=" & XN(TipoRec)
        sql = sql & "   AND REC_NUMERO=" & XN(NroRec)
        sql = sql & "   AND REC_SUCURSAL=" & XN(NroSuc)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
        If rec.EOF = False Then
            Do While rec.EOF = False
                GrillaAplicar.AddItem BuscarTipoDocAbre(rec!FCL_TCO_CODIGO) & Chr(9) & _
                            Format(rec!FCL_SUCURSAL, "0000") & "-" & Format(rec!FCL_NUMERO, "00000000") & Chr(9) & rec!FCL_FECHA _
                             & Chr(9) & VALIDO_IMPORTE(rec!REC_IMPORTE) & Chr(9) & VALIDO_IMPORTE(rec!REC_ABONA) & Chr(9) & VALIDO_IMPORTE(rec!REC_SALDO) & Chr(9) & rec!FCL_TCO_CODIGO
                            
                rec.MoveNext
            Loop
            GrillaAplicar.HighLight = flexHighlightAlways
            txtImporteApagar.Text = SumaGrilla(GrillaAplicar, 4)
        End If
        FrameRecibo.Enabled = False
        FrameRemito.Enabled = False
        rec.Close
        cmdNuevo.SetFocus
        cmdGrabar.Enabled = False
        mBorroTransfe = False
    End If
    Rec2.Close
End Sub

Private Function BuscarCheque(Codigo As String, NroChe As String) As String
    
    Set Rec1 = New ADODB.Recordset
    sql = "SELECT B.BAN_DESCRI,C.CHE_IMPORT,C.CHE_FECVTO"
    sql = sql & " FROM BANCO B, CHEQUE C"
    sql = sql & " WHERE C.BAN_CODINT=" & XN(Codigo)
    sql = sql & " AND C.CHE_NUMERO=" & XS(NroChe)
    sql = sql & " AND C.BAN_CODINT=B.BAN_CODINT"
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        grillaValores.AddItem "CHE" & Chr(9) & VALIDO_IMPORTE(Rec1!che_import) & Chr(9) & Rec1!CHE_FECVTO _
                           & Chr(9) & Rec1!BAN_DESCRI & Chr(9) & NroChe & Chr(9) & Codigo
    End If
    Rec1.Close
End Function

Private Function BuscarMoneda(Codigo As String) As String
    
    Set Rec1 = New ADODB.Recordset
    sql = "SELECT MON_DESCRI"
    sql = sql & " FROM MONEDA"
    sql = sql & " WHERE MON_CODIGO=" & XN(Codigo)
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        BuscarMoneda = Rec1!MON_DESCRI
    Else
        BuscarMoneda = ""
    End If
    Rec1.Close
End Function
Private Function BuscarTarjeta(Codigo As String) As String
    
    Set Rec1 = New ADODB.Recordset
    sql = "SELECT TAR_DESCRI"
    sql = sql & " FROM TARJETA"
    sql = sql & " WHERE TAR_CODIGO=" & XN(Codigo)
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        BuscarTarjeta = Rec1!TAR_DESCRI
    Else
        BuscarTarjeta = ""
    End If
    Rec1.Close
End Function

Private Sub txtNroSucursal_GotFocus()
    SelecTexto txtNroSucursal
End Sub

Private Sub txtNroSucursal_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroSucursal_LostFocus()
    If txtNroSucursal.Text = "" Then
        txtNroSucursal.Text = Sucursal
    Else
        txtNroSucursal.Text = Format(txtNroSucursal.Text, "0000")
    End If
End Sub

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
            If Txt = "txtCodCliente" Then
                txtCodCliente.Text = .ResultFields(2)
                txtCodCliente_LostFocus
            Else
                txtCliente.Text = .ResultFields(2)
                txtCliente_LostFocus
            End If
        End If
    End With
    
    Set B = Nothing
End Sub

Private Sub txtSucursal_GotFocus()
    SelecTexto TxtSUCURSAL
End Sub

Private Sub TxtSucursal_KeyPress(KeyAscii As Integer)
     KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtTar_Autorizacion_GotFocus()
    SelecTexto txtTar_Autorizacion
End Sub

Private Sub txtTar_Autorizacion_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub


Private Sub txtTotalPagos_GotFocus()
    SelecTexto txtTotalPagos
End Sub

Private Sub txtTotalPagos_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtTotalPagos, KeyAscii)
End Sub
