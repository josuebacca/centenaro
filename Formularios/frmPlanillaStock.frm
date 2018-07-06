VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPlanillaStock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planilla de Ventas"
   ClientHeight    =   7716
   ClientLeft      =   156
   ClientTop       =   540
   ClientWidth     =   10752
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7716
   ScaleWidth      =   10752
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   8298
      TabIndex        =   151
      Top             =   7320
      Width           =   1155
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      Height          =   345
      Left            =   4800
      TabIndex        =   146
      Top             =   7320
      Width           =   1155
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   345
      Left            =   7132
      TabIndex        =   150
      Top             =   7320
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   9465
      TabIndex        =   153
      Top             =   7320
      Width           =   1155
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   5966
      TabIndex        =   148
      Top             =   7320
      Width           =   1155
   End
   Begin TabDlg.SSTab TabDatos 
      Height          =   7095
      Left            =   120
      TabIndex        =   75
      Top             =   120
      Width           =   10575
      _ExtentX        =   18648
      _ExtentY        =   12510
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Resumen"
      TabPicture(0)   =   "frmPlanillaStock.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraResumen"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame10"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame11"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdGetDatos"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Stock"
      TabPicture(1)   =   "frmPlanillaStock.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(2)=   "Frame5"
      Tab(1).Control(3)=   "Frame8"
      Tab(1).Control(4)=   "Frame6"
      Tab(1).Control(5)=   "Frame7"
      Tab(1).Control(6)=   "Frame2"
      Tab(1).Control(7)=   "Frame4"
      Tab(1).Control(8)=   "Frame13"
      Tab(1).Control(9)=   "Frame9"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Facturas"
      TabPicture(2)   =   "frmPlanillaStock.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label80"
      Tab(2).Control(1)=   "grdFacturas"
      Tab(2).Control(2)=   "txtTotFac(0)"
      Tab(2).Control(3)=   "txtTotFac(1)"
      Tab(2).Control(4)=   "txtTotFac(2)"
      Tab(2).Control(5)=   "txtTotFac(3)"
      Tab(2).Control(6)=   "txtTotFac(4)"
      Tab(2).Control(7)=   "txtTotFac(5)"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Lubricantes"
      TabPicture(3)   =   "frmPlanillaStock.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label79"
      Tab(3).Control(1)=   "Label81"
      Tab(3).Control(2)=   "grdLubricantes"
      Tab(3).Control(3)=   "txtTotLub"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Buscar Anteriores"
      TabPicture(4)   =   "frmPlanillaStock.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "grdModulos"
      Tab(4).Control(1)=   "Frame12"
      Tab(4).ControlCount=   2
      Begin VB.Frame Frame12 
         Caption         =   "Listar por..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -74880
         TabIndex        =   199
         Top             =   360
         Width           =   10335
         Begin VB.CommandButton cmdListar 
            Height          =   660
            Left            =   9000
            Picture         =   "frmPlanillaStock.frx":008C
            Style           =   1  'Graphical
            TabIndex        =   203
            Top             =   360
            Width           =   615
         End
         Begin VB.ComboBox cboPlaBus 
            Height          =   315
            Left            =   2280
            Style           =   2  'Dropdown List
            TabIndex        =   200
            Top             =   405
            Width           =   5235
         End
         Begin MSComCtl2.DTPicker FechaDesde 
            Height          =   315
            Left            =   2280
            TabIndex        =   201
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
         Begin MSComCtl2.DTPicker FechaHasta 
            Height          =   315
            Left            =   6050
            TabIndex        =   202
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
         Begin VB.Label Label76 
            AutoSize        =   -1  'True
            Caption         =   "Fecha      Desde:"
            Height          =   195
            Left            =   1020
            TabIndex        =   206
            Top             =   825
            Width           =   1215
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Hasta:"
            Height          =   195
            Index           =   2
            Left            =   5520
            TabIndex        =   205
            Top             =   780
            Width           =   480
         End
         Begin VB.Label Label75 
            AutoSize        =   -1  'True
            Caption         =   "Playero:"
            Height          =   195
            Left            =   1665
            TabIndex        =   204
            Top             =   450
            Width           =   570
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "GASOIL - Manguera 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -68100
         TabIndex        =   184
         Top             =   2400
         Width           =   3435
         Begin VB.TextBox txtNSLAnt 
            Height          =   315
            Index           =   13
            Left            =   2160
            TabIndex        =   192
            Top             =   870
            Width           =   1215
         End
         Begin VB.TextBox txtNSLTot 
            BackColor       =   &H80000002&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   315
            Index           =   13
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   191
            Top             =   1230
            Width           =   1215
         End
         Begin VB.TextBox txtNSLPesos 
            BackColor       =   &H80000002&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   315
            Index           =   13
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   190
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtNSLActual 
            Height          =   315
            Index           =   13
            Left            =   2160
            TabIndex        =   189
            Text            =   " "
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtNSLAnt 
            Height          =   315
            Index           =   12
            Left            =   960
            TabIndex        =   188
            Top             =   870
            Width           =   1215
         End
         Begin VB.TextBox txtNSLTot 
            BackColor       =   &H80000002&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   315
            Index           =   12
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   187
            Top             =   1230
            Width           =   1215
         End
         Begin VB.TextBox txtNSLPesos 
            BackColor       =   &H80000002&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   315
            Index           =   12
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   186
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtNSLActual 
            Height          =   315
            Index           =   12
            Left            =   960
            TabIndex        =   185
            Text            =   " "
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label86 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2º Numerac"
            Height          =   255
            Left            =   2160
            TabIndex        =   198
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label85 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Litros"
            Height          =   315
            Left            =   960
            TabIndex        =   197
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label Label84 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Anterior"
            Height          =   315
            Left            =   120
            TabIndex        =   196
            Top             =   870
            Width           =   855
         End
         Begin VB.Label Label83 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Actual"
            Height          =   315
            Left            =   120
            TabIndex        =   195
            Top             =   510
            Width           =   855
         End
         Begin VB.Label Label82 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PESOS"
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
            Left            =   120
            TabIndex        =   194
            Top             =   1530
            Width           =   855
         End
         Begin VB.Label Label78 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "VENTA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   193
            Top             =   1200
            Width           =   690
         End
      End
      Begin VB.TextBox txtTotFac 
         BackColor       =   &H80000002&
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   5
         Left            =   -65715
         TabIndex        =   183
         Top             =   6720
         Width           =   1050
      End
      Begin VB.TextBox txtTotFac 
         BackColor       =   &H80000002&
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   4
         Left            =   -66756
         TabIndex        =   182
         Top             =   6720
         Width           =   1050
      End
      Begin VB.TextBox txtTotFac 
         BackColor       =   &H80000002&
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   3
         Left            =   -67797
         TabIndex        =   181
         Top             =   6720
         Width           =   1050
      End
      Begin VB.TextBox txtTotFac 
         BackColor       =   &H80000002&
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   2
         Left            =   -68838
         TabIndex        =   180
         Top             =   6720
         Width           =   1050
      End
      Begin VB.TextBox txtTotFac 
         BackColor       =   &H80000002&
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   1
         Left            =   -69879
         TabIndex        =   179
         Top             =   6720
         Width           =   1050
      End
      Begin VB.CommandButton cmdGetDatos 
         Caption         =   "&Obtener datos del Turno"
         Height          =   975
         Left            =   4800
         Picture         =   "frmPlanillaStock.frx":0416
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtTotLub 
         BackColor       =   &H80000002&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   -65760
         TabIndex        =   175
         Top             =   6720
         Width           =   1215
      End
      Begin VB.TextBox txtTotFac 
         BackColor       =   &H80000002&
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   0
         Left            =   -70920
         TabIndex        =   174
         Top             =   6720
         Width           =   1050
      End
      Begin VB.Frame Frame13 
         Caption         =   "Puente de Medicion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -74880
         TabIndex        =   168
         Top             =   6360
         Width           =   10215
         Begin VB.TextBox txtMedicion 
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
            Index           =   1
            Left            =   6120
            TabIndex        =   74
            Text            =   "0,00"
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox txtMedicion 
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
            Index           =   0
            Left            =   2280
            TabIndex        =   73
            Text            =   "0,00"
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label77 
            AutoSize        =   -1  'True
            Caption         =   "Mecanica"
            Height          =   195
            Left            =   5280
            TabIndex        =   170
            Top             =   300
            Width           =   705
         End
         Begin VB.Label Label74 
            AutoSize        =   -1  'True
            Caption         =   "Digital"
            Height          =   195
            Left            =   1680
            TabIndex        =   169
            Top             =   300
            Width           =   435
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "NAFTA SUPER - Manguera 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -69675
         TabIndex        =   156
         Top             =   480
         Width           =   4400
         Begin VB.TextBox txtNSLActual 
            Height          =   315
            Index           =   3
            Left            =   2880
            TabIndex        =   37
            Text            =   " "
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtNSLPesos 
            BackColor       =   &H80000002&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   315
            Index           =   3
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   1560
            Width           =   1335
         End
         Begin VB.TextBox txtNSLTot 
            BackColor       =   &H80000002&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   315
            Index           =   3
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   1230
            Width           =   1335
         End
         Begin VB.TextBox txtNSLAnt 
            Height          =   315
            Index           =   3
            Left            =   2880
            TabIndex        =   38
            Top             =   870
            Width           =   1335
         End
         Begin VB.TextBox txtNSLActual 
            Height          =   315
            Index           =   2
            Left            =   1560
            TabIndex        =   33
            Text            =   " "
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtNSLPesos 
            BackColor       =   &H80000002&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   315
            Index           =   2
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   1560
            Width           =   1335
         End
         Begin VB.TextBox txtNSLTot 
            BackColor       =   &H80000002&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   315
            Index           =   2
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   1230
            Width           =   1335
         End
         Begin VB.TextBox txtNSLAnt 
            Height          =   315
            Index           =   2
            Left            =   1560
            TabIndex        =   34
            Top             =   870
            Width           =   1335
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PESOS"
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
            Left            =   120
            TabIndex        =   162
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "TOTAL VENTA"
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
            Left            =   120
            TabIndex        =   161
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Litros"
            Height          =   315
            Left            =   1560
            TabIndex        =   160
            Top             =   240
            Width           =   1305
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Anterior"
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   159
            Top             =   870
            Width           =   1455
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Actual"
            Height          =   315
            Left            =   120
            TabIndex        =   158
            Top             =   510
            Width           =   1455
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2º Numerac"
            Height          =   255
            Left            =   2880
            TabIndex        =   157
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "GASOIL - Manguera 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -71475
         TabIndex        =   147
         Top             =   2400
         Width           =   3435
         Begin VB.TextBox txtNSLActual 
            Height          =   315
            Index           =   7
            Left            =   2160
            TabIndex        =   53
            Text            =   " "
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtNSLPesos 
            BackColor       =   &H80000002&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   315
            Index           =   7
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   56
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtNSLTot 
            BackColor       =   &H80000002&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   315
            Index           =   7
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   55
            Top             =   1230
            Width           =   1215
         End
         Begin VB.TextBox txtNSLAnt 
            Height          =   315
            Index           =   7
            Left            =   2160
            TabIndex        =   54
            Top             =   870
            Width           =   1215
         End
         Begin VB.TextBox txtNSLActual 
            Height          =   315
            Index           =   6
            Left            =   960
            TabIndex        =   49
            Text            =   " "
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtNSLPesos 
            BackColor       =   &H80000002&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   315
            Index           =   6
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   52
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtNSLTot 
            BackColor       =   &H80000002&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   315
            Index           =   6
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   51
            Top             =   1230
            Width           =   1215
         End
         Begin VB.TextBox txtNSLAnt 
            Height          =   315
            Index           =   6
            Left            =   960
            TabIndex        =   50
            Top             =   870
            Width           =   1215
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "VENTA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   166
            Top             =   1200
            Width           =   690
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PESOS"
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
            Left            =   120
            TabIndex        =   165
            Top             =   1530
            Width           =   855
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Actual"
            Height          =   315
            Left            =   120
            TabIndex        =   155
            Top             =   510
            Width           =   855
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Anterior"
            Height          =   315
            Left            =   120
            TabIndex        =   154
            Top             =   870
            Width           =   855
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Litros"
            Height          =   315
            Left            =   960
            TabIndex        =   152
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2º Numerac"
            Height          =   255
            Left            =   2160
            TabIndex        =   149
            Top             =   240
            Width           =   1140
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "GNC - Surtidor 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   -69720
         TabIndex        =   138
         Top             =   4320
         Width           =   2535
         Begin VB.TextBox txtNSLAnt 
            Height          =   315
            Index           =   10
            Left            =   1200
            TabIndex        =   66
            Top             =   870
            Width           =   1260
         End
         Begin VB.TextBox txtNSLTot 
            BackColor       =   &H80000002&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   315
            Index           =   10
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   67
            Top             =   1230
            Width           =   1260
         End
         Begin VB.TextBox txtNSLPesos 
            BackColor       =   &H80000002&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   315
            Index           =   10
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   68
            Top             =   1560
            Width           =   1260
         End
         Begin VB.TextBox txtNSLActual 
            Height          =   315
            Index           =   10
            Left            =   1200
            TabIndex        =   65
            Text            =   " "
            Top             =   480
            Width           =   1260
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Actual"
            Height          =   315
            Left            =   120
            TabIndex        =   139
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Pesos                "
            Height          =   255
            Left            =   120
            TabIndex        =   144
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Total Venta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   143
            Top             =   1200
            Width           =   1065
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Anterior"
            Height          =   315
            Left            =   120
            TabIndex        =   142
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Manguera 1"
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
            Left            =   120
            TabIndex        =   141
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Metros"
            Height          =   255
            Left            =   1200
            TabIndex        =   140
            Top             =   240
            Width           =   1260
         End
      End
      Begin VB.Frame Frame6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   -72240
         TabIndex        =   131
         Top             =   4320
         Width           =   2535
         Begin VB.TextBox txtNSLAnt 
            Height          =   315
            Index           =   9
            Left            =   1200
            TabIndex        =   62
            Top             =   870
            Width           =   1260
         End
         Begin VB.TextBox txtNSLTot 
            BackColor       =   &H80000002&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   315
            Index           =   9
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   63
            Top             =   1230
            Width           =   1260
         End
         Begin VB.TextBox txtNSLPesos 
            BackColor       =   &H80000002&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   315
            Index           =   9
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   64
            Top             =   1560
            Width           =   1260
         End
         Begin VB.TextBox txtNSLActual 
            Height          =   315
            Index           =   9
            Left            =   1200
            TabIndex        =   61
            Text            =   " "
            Top             =   480
            Width           =   1260
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Metros"
            Height          =   255
            Left            =   1200
            TabIndex        =   137
            Top             =   240
            Width           =   1260
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Manguera 2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   136
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Pesos                    "
            Height          =   255
            Left            =   120
            TabIndex        =   135
            Top             =   1560
            Width           =   1155
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Total Venta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   134
            Top             =   1200
            Width           =   1065
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Anterior"
            Height          =   315
            Left            =   120
            TabIndex        =   133
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Actual"
            Height          =   315
            Left            =   120
            TabIndex        =   132
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.Frame Frame11 
         Height          =   3615
         Left            =   1380
         TabIndex        =   125
         Top             =   360
         Width           =   3375
         Begin VB.ComboBox cboTurno 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   480
            Width           =   1395
         End
         Begin VB.TextBox txtId 
            Height          =   375
            Left            =   1440
            TabIndex        =   145
            Top             =   360
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtObservaciones 
            BackColor       =   &H00C0FFFF&
            Height          =   1005
            Left            =   150
            MaxLength       =   60
            TabIndex        =   5
            Top             =   2400
            Width           =   3090
         End
         Begin VB.ComboBox cboPla2 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1800
            Width           =   3195
         End
         Begin VB.ComboBox cboPla1 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1080
            Width           =   3195
         End
         Begin VB.TextBox txtTurno 
            Height          =   350
            Left            =   2160
            TabIndex        =   6
            Top             =   720
            Visible         =   0   'False
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker Fecha 
            Height          =   315
            Left            =   120
            TabIndex        =   0
            Top             =   480
            Width           =   1455
            _ExtentX        =   2561
            _ExtentY        =   550
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   106299393
            CurrentDate     =   41098
         End
         Begin VB.Label Label73 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones:"
            Height          =   195
            Left            =   120
            TabIndex        =   130
            Top             =   2160
            Width           =   1125
         End
         Begin VB.Label Label72 
            AutoSize        =   -1  'True
            Caption         =   "Encargado del Turno Nº 2:"
            Height          =   195
            Left            =   120
            TabIndex        =   129
            Top             =   1560
            Width           =   1905
         End
         Begin VB.Label Label71 
            AutoSize        =   -1  'True
            Caption         =   "Encargado del Turno Nº 1:"
            Height          =   195
            Left            =   120
            TabIndex        =   128
            Top             =   840
            Width           =   1905
         End
         Begin VB.Label Label70 
            AutoSize        =   -1  'True
            Caption         =   "Turno:"
            Height          =   195
            Left            =   1920
            TabIndex        =   127
            Top             =   240
            Width           =   465
         End
         Begin VB.Label frmPlanillaStock 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   120
            TabIndex        =   126
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   6120
         TabIndex        =   110
         Top             =   360
         Width           =   3375
         Begin VB.TextBox txtTRend 
            BackColor       =   &H80000001&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000004&
            Height          =   350
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   24
            Text            =   "0,00"
            Top             =   6120
            Width           =   1335
         End
         Begin VB.TextBox txtDiff 
            Height          =   350
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   23
            Text            =   "0,00"
            Top             =   5630
            Width           =   1335
         End
         Begin VB.TextBox txtVarios 
            Height          =   350
            Left            =   1800
            TabIndex        =   22
            Text            =   "0,00 "
            Top             =   5142
            Width           =   1335
         End
         Begin VB.TextBox txtDolares 
            Height          =   350
            Left            =   1800
            TabIndex        =   17
            Text            =   "0,00"
            Top             =   2702
            Width           =   1335
         End
         Begin VB.TextBox txtCheq 
            Height          =   350
            Left            =   1800
            TabIndex        =   21
            Text            =   "0,00"
            Top             =   4654
            Width           =   1335
         End
         Begin VB.TextBox txtDiario 
            Height          =   350
            Left            =   1800
            TabIndex        =   20
            Text            =   "0,00"
            Top             =   4166
            Width           =   1335
         End
         Begin VB.TextBox txtTar 
            Height          =   350
            Left            =   1800
            TabIndex        =   19
            Text            =   "0,00"
            Top             =   3678
            Width           =   1335
         End
         Begin VB.TextBox txtFacP 
            Height          =   350
            Left            =   1800
            TabIndex        =   16
            Text            =   "0,00"
            Top             =   2214
            Width           =   1335
         End
         Begin VB.TextBox txtVale 
            Height          =   350
            Left            =   1800
            TabIndex        =   15
            Text            =   "0,00"
            Top             =   1726
            Width           =   1335
         End
         Begin VB.TextBox txtefe 
            Height          =   350
            Left            =   1800
            TabIndex        =   14
            Text            =   "0,00"
            Top             =   1238
            Width           =   1335
         End
         Begin VB.TextBox txtRet 
            Height          =   350
            Left            =   1800
            TabIndex        =   13
            Text            =   "0,00"
            Top             =   750
            Width           =   1335
         End
         Begin VB.TextBox txtCtaC 
            Height          =   350
            Left            =   1800
            TabIndex        =   18
            Text            =   "0,00"
            Top             =   3190
            Width           =   1350
         End
         Begin VB.Label Label69 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "TOTAL RENDIDO"
            Height          =   345
            Left            =   120
            TabIndex        =   124
            Top             =   6120
            Width           =   1605
         End
         Begin VB.Label Label68 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Diferencia de Caja"
            Height          =   345
            Left            =   120
            TabIndex        =   123
            Top             =   5620
            Width           =   1605
         End
         Begin VB.Label Label67 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Varios"
            Height          =   345
            Left            =   120
            TabIndex        =   122
            Top             =   5130
            Width           =   1605
         End
         Begin VB.Label Label66 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Monedas"
            Height          =   315
            Left            =   120
            TabIndex        =   121
            Top             =   2710
            Width           =   1605
         End
         Begin VB.Label Label65 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cheques"
            Height          =   345
            Left            =   120
            TabIndex        =   120
            Top             =   4640
            Width           =   1605
         End
         Begin VB.Label Label64 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Diario"
            Height          =   345
            Left            =   120
            TabIndex        =   119
            Top             =   4150
            Width           =   1605
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tarjetas"
            Height          =   345
            Left            =   120
            TabIndex        =   118
            Top             =   3660
            Width           =   1605
         End
         Begin VB.Label Label62 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Facturas Pagadas"
            Height          =   345
            Left            =   120
            TabIndex        =   117
            Top             =   2220
            Width           =   1605
         End
         Begin VB.Label Label61 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vales autorizados"
            Height          =   345
            Left            =   120
            TabIndex        =   116
            Top             =   1730
            Width           =   1605
         End
         Begin VB.Label Label60 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Efectivo"
            Height          =   345
            Left            =   120
            TabIndex        =   115
            Top             =   1240
            Width           =   1605
         End
         Begin VB.Label Label59 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Retiro de efectivo"
            Height          =   345
            Left            =   120
            TabIndex        =   114
            Top             =   750
            Width           =   1605
         End
         Begin VB.Label Label58 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000C&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Resumen de CAJA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   350
            Left            =   120
            TabIndex        =   113
            Top             =   300
            Width           =   1635
         End
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000C&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Pesos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   345
            Left            =   1800
            TabIndex        =   112
            Top             =   300
            Width           =   1350
         End
         Begin VB.Label Label56 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cuentas Corrientes"
            Height          =   345
            Left            =   120
            TabIndex        =   111
            Top             =   3170
            Width           =   1605
         End
      End
      Begin VB.Frame fraResumen 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   1380
         TabIndex        =   102
         Top             =   3960
         Width           =   3375
         Begin VB.TextBox txtNaftaEco 
            Height          =   350
            Left            =   1800
            TabIndex        =   8
            Text            =   "0,00"
            Top             =   1850
            Width           =   1335
         End
         Begin VB.TextBox txtTotal1 
            BackColor       =   &H80000002&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   350
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   12
            Text            =   "0,00"
            Top             =   2520
            Width           =   1335
         End
         Begin VB.TextBox txtGasOil1 
            Height          =   350
            Left            =   1800
            TabIndex        =   9
            Text            =   "0,00"
            Top             =   1110
            Width           =   1335
         End
         Begin VB.TextBox txtGNC1 
            Height          =   350
            Left            =   1800
            TabIndex        =   10
            Text            =   "0,00"
            Top             =   1470
            Width           =   1335
         End
         Begin VB.TextBox txtLub1 
            Height          =   350
            Left            =   1800
            TabIndex        =   11
            Text            =   "0,00"
            Top             =   2190
            Width           =   1335
         End
         Begin VB.TextBox txtNSuper1 
            Height          =   350
            Left            =   1800
            TabIndex        =   7
            Text            =   "0,00"
            Top             =   750
            Width           =   1335
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "GNC M3"
            Height          =   315
            Left            =   120
            TabIndex        =   167
            Top             =   1850
            Width           =   1680
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "TOTAL GENERAL"
            Height          =   345
            Left            =   120
            TabIndex        =   109
            Top             =   2520
            Width           =   1680
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000C&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Pesos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   345
            Left            =   1800
            TabIndex        =   108
            Top             =   300
            Width           =   1350
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000C&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Resumen de Venta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   345
            Left            =   120
            TabIndex        =   107
            Top             =   300
            Width           =   1680
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nafta Super"
            Height          =   345
            Left            =   120
            TabIndex        =   106
            Top             =   750
            Width           =   1680
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Gas Oil"
            Height          =   345
            Left            =   120
            TabIndex        =   105
            Top             =   1110
            Width           =   1680
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "GNC"
            Height          =   345
            Left            =   120
            TabIndex        =   104
            Top             =   1470
            Width           =   1680
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Lubricantes"
            Height          =   345
            Left            =   120
            TabIndex        =   103
            Top             =   2190
            Width           =   1680
         End
      End
      Begin VB.Frame Frame8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   -67200
         TabIndex        =   95
         Top             =   4320
         Width           =   2535
         Begin VB.TextBox txtNSLAnt 
            Height          =   315
            Index           =   11
            Left            =   1200
            TabIndex        =   70
            Top             =   870
            Width           =   1260
         End
         Begin VB.TextBox txtNSLTot 
            BackColor       =   &H80000002&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   315
            Index           =   11
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   71
            Top             =   1230
            Width           =   1260
         End
         Begin VB.TextBox txtNSLPesos 
            BackColor       =   &H80000002&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   315
            Index           =   11
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   72
            Top             =   1560
            Width           =   1260
         End
         Begin VB.TextBox txtNSLActual 
            Height          =   315
            Index           =   11
            Left            =   1200
            TabIndex        =   69
            Text            =   " "
            Top             =   480
            Width           =   1260
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Actual"
            Height          =   315
            Left            =   120
            TabIndex        =   101
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Anterior"
            Height          =   315
            Left            =   120
            TabIndex        =   100
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Total Venta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   99
            Top             =   1200
            Width           =   1065
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Pesos              "
            Height          =   255
            Left            =   120
            TabIndex        =   98
            Top             =   1560
            Width           =   1005
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Manguera 2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   97
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Metros"
            Height          =   255
            Left            =   1200
            TabIndex        =   96
            Top             =   240
            Width           =   1260
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "GNC - Surtidor 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   -74880
         TabIndex        =   88
         Top             =   4320
         Width           =   2655
         Begin VB.TextBox txtNSLAnt 
            Height          =   315
            Index           =   8
            Left            =   1200
            TabIndex        =   58
            Top             =   870
            Width           =   1260
         End
         Begin VB.TextBox txtNSLTot 
            BackColor       =   &H80000002&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   315
            Index           =   8
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   59
            Top             =   1230
            Width           =   1260
         End
         Begin VB.TextBox txtNSLPesos 
            BackColor       =   &H80000002&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   315
            Index           =   8
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   60
            Top             =   1560
            Width           =   1260
         End
         Begin VB.TextBox txtNSLActual 
            Height          =   315
            Index           =   8
            Left            =   1200
            TabIndex        =   57
            Text            =   " "
            Top             =   480
            Width           =   1260
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Actual"
            Height          =   315
            Left            =   120
            TabIndex        =   93
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Metros"
            Height          =   255
            Left            =   1200
            TabIndex        =   89
            Top             =   240
            Width           =   1260
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Manguera 1"
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
            Left            =   120
            TabIndex        =   94
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Anterior"
            Height          =   315
            Left            =   120
            TabIndex        =   92
            Top             =   870
            Width           =   1095
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Total Venta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   91
            Top             =   1230
            Width           =   1065
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Pesos                  "
            Height          =   255
            Left            =   120
            TabIndex        =   90
            Top             =   1590
            Width           =   1065
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "GASOIL - Manguera 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -74880
         TabIndex        =   83
         Top             =   2400
         Width           =   3435
         Begin VB.TextBox txtNSLActual 
            Height          =   315
            Index           =   5
            Left            =   2160
            TabIndex        =   45
            Text            =   " "
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtNSLPesos 
            BackColor       =   &H80000002&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   315
            Index           =   5
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   48
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtNSLTot 
            BackColor       =   &H80000002&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   315
            Index           =   5
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   47
            Top             =   1230
            Width           =   1215
         End
         Begin VB.TextBox txtNSLAnt 
            Height          =   315
            Index           =   5
            Left            =   2160
            TabIndex        =   46
            Top             =   870
            Width           =   1215
         End
         Begin VB.TextBox txtNSLActual 
            Height          =   315
            Index           =   4
            Left            =   960
            TabIndex        =   41
            Text            =   " "
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtNSLPesos 
            BackColor       =   &H80000002&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   315
            Index           =   4
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   44
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtNSLTot 
            BackColor       =   &H80000002&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   315
            Index           =   4
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   43
            Top             =   1230
            Width           =   1215
         End
         Begin VB.TextBox txtNSLAnt 
            Height          =   315
            Index           =   4
            Left            =   960
            TabIndex        =   42
            Top             =   870
            Width           =   1215
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PESOS"
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
            Left            =   120
            TabIndex        =   164
            Top             =   1560
            Width           =   855
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "VENTA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   163
            Top             =   1230
            Width           =   690
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2º Numerac"
            Height          =   255
            Left            =   2160
            TabIndex        =   84
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Litros"
            Height          =   315
            Left            =   960
            TabIndex        =   87
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Anterior"
            Height          =   315
            Left            =   120
            TabIndex        =   86
            Top             =   870
            Width           =   855
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Actual"
            Height          =   315
            Left            =   120
            TabIndex        =   85
            Top             =   510
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "NAFTA SUPER - Manguera 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -74040
         TabIndex        =   76
         Top             =   480
         Width           =   4400
         Begin VB.TextBox txtNSLPesos 
            BackColor       =   &H80000002&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   315
            Index           =   1
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   1560
            Width           =   1335
         End
         Begin VB.TextBox txtNSLTot 
            BackColor       =   &H80000002&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   315
            Index           =   1
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   1230
            Width           =   1335
         End
         Begin VB.TextBox txtNSLAnt 
            Height          =   315
            Index           =   1
            Left            =   2880
            TabIndex        =   30
            Top             =   870
            Width           =   1335
         End
         Begin VB.TextBox txtNSLActual 
            Height          =   315
            Index           =   1
            Left            =   2880
            TabIndex        =   29
            Text            =   " "
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtNSLTot 
            BackColor       =   &H80000002&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   315
            Index           =   0
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   1230
            Width           =   1335
         End
         Begin VB.TextBox txtNSLPesos 
            BackColor       =   &H80000002&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   315
            Index           =   0
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   1560
            Width           =   1335
         End
         Begin VB.TextBox txtNSLActual 
            Height          =   315
            Index           =   0
            Left            =   1560
            TabIndex        =   25
            Text            =   " "
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtNSLAnt 
            Height          =   315
            Index           =   0
            Left            =   1560
            TabIndex        =   26
            Top             =   870
            Width           =   1335
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2º Numerac"
            Height          =   255
            Left            =   2880
            TabIndex        =   77
            Top             =   240
            Width           =   1380
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Actual"
            Height          =   315
            Left            =   120
            TabIndex        =   82
            Top             =   510
            Width           =   1455
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Anterior"
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   81
            Top             =   870
            Width           =   1455
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "TOTAL VENTA"
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
            Left            =   120
            TabIndex        =   80
            Top             =   1230
            Width           =   1455
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PESOS"
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
            Left            =   120
            TabIndex        =   79
            Top             =   1550
            Width           =   1455
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Litros"
            Height          =   315
            Left            =   1560
            TabIndex        =   78
            Top             =   240
            Width           =   1275
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grdModulos 
         Height          =   5370
         Left            =   -74880
         TabIndex        =   171
         Top             =   1680
         Width           =   10335
         _ExtentX        =   18225
         _ExtentY        =   9483
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         BackColorSel    =   8388736
         FocusRect       =   0
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid grdFacturas 
         Height          =   6210
         Left            =   -74880
         TabIndex        =   172
         Top             =   480
         Width           =   10335
         _ExtentX        =   18225
         _ExtentY        =   10964
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         BackColorSel    =   8388736
         FocusRect       =   0
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid grdLubricantes 
         Height          =   6210
         Left            =   -74880
         TabIndex        =   173
         Top             =   480
         Width           =   10335
         _ExtentX        =   18225
         _ExtentY        =   10964
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         BackColorSel    =   8388736
         FocusRect       =   0
         SelectionMode   =   1
      End
      Begin VB.Label Label81 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Presione <F5> para actualizar Lubricantes"
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
         Left            =   -74880
         TabIndex        =   178
         Top             =   6840
         Width           =   3600
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<F5> para actualizar Facturas"
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
         Left            =   -74880
         TabIndex        =   177
         Top             =   6840
         Width           =   2565
      End
      Begin VB.Label Label79 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Venta"
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
         Left            =   -66840
         TabIndex        =   176
         Top             =   6720
         Width           =   1065
      End
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   120
      Top             =   7800
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowControls  =   -1  'True
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmPlanillaStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vEstado As String
Dim vUltTurno As Integer
Private Function buscaUltimoTurno()
    sql = "SELECT MAX(T_ID) AS TURNO FROM T_STOCK"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        vUltTurno = IIf(IsNull(rec!turno), 0, rec!turno)
    End If
    rec.Close
End Function
    
Private Function Validar() As Boolean
'    If txtNroFactura.Text = "" Then
'        MsgBox "Falta el Número de la Factura", vbExclamation, TIT_MSGBOX
'        ValidarFactura = False
'        Exit Function
'    End If
    If Fecha.Value = "" Then
        MsgBox "La Fecha es requerida", vbExclamation, TIT_MSGBOX
        Fecha.SetFocus
        Validar = False
        Exit Function
    End If
    If cboTurno.ListIndex = -1 Then
        MsgBox "El Turno es requerido", vbCritical, TIT_MSGBOX
        cboTurno.SetFocus
        Validar = False
        Exit Function
    End If
    
    If cboPla1.List(cboPla1.ListIndex) = "" Then
        MsgBox "El Encargado del Turno Nº 1 es Requerido", vbExclamation, TIT_MSGBOX
        cboPla1.SetFocus
        Validar = False
        Exit Function
    End If
    If txtefe.Text = "0,00" Then
        MsgBox "Debe ingresar el efectivo a rendir", vbCritical, TIT_MSGBOX
        txtefe.SetFocus
        Validar = False
        Exit Function
    End If
'    If cboPla2.List(cboPla2.ListIndex) = "" Then
'        MsgBox "El Encargado del Turno Nº 2 es Requerido", vbExclamation, TIT_MSGBOX
'        cboPla2.SetFocus
'        Validar = False
'        Exit Function
'    End If
    
'    If txtTotal.Text = "" Then
'        MsgBox "El Total de la Factura no puede ser Nulo", vbCritical, TIT_MSGBOX
'        grdGrilla.Col = 0
'        grdGrilla.row = 2
'        grdGrilla.SetFocus
'        ValidarFactura = False
'        Exit Function
'    End If
    
    Validar = True
End Function

Private Function BuscoTurnoAnterior()
    Dim turnoAnt As Integer
    Dim i As Integer
    sql = "SELECT MAX(T_ID) AS TURNO FROM T_STOCK"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        turnoAnt = IIf(IsNull(rec!turno), 0, rec!turno)
    End If
    rec.Close
    If turnoAnt <> 0 Then
        sql = "SELECT T_NSACTUAL,T_NSACTUAL2, T_NEACTUAL, T_NEACTUAL2"
        sql = sql & ",T_G1ACTUAL,T_G1ACTUAL2,T_G2ACTUAL,T_G2ACTUAL2"
        sql = sql & ",T_GNC1ACTUAL,T_GNC2ACTUAL,T_GNC3ACTUAL,T_GNC4ACTUAL "
        sql = sql & ",T_G3ACTUAL,T_G3ACTUAL2,T_G3ACTUAL,T_G3ACTUAL2"
        sql = sql & " FROM T_STOCK "
        sql = sql & " WHERE T_ID=" & turnoAnt
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            For i = 0 To 13
                If Not IsNull(rec.Fields(i)) Then
                    txtNSLAnt(i).Text = IIf(IsNull(rec.Fields(i)), "0,00", VALIDO_IMPORTE(rec.Fields(i)))
                    txtNSLActual(i).Text = IIf(IsNull(rec.Fields(i)), "0,00", VALIDO_IMPORTE(rec.Fields(i)))
                Else
                    txtNSLAnt(i).Text = "0,00"
                    txtNSLActual(i).Text = "0,00"
                End If
            Next i
        End If
        rec.Close
    End If
End Function

Private Sub cboTurno_LostFocus()
    If cboTurno.ItemData(cboTurno.ListIndex) = 3 Then 'TURNONOCHE
        Fecha.Value = Date - 1
        cboPla1.SetFocus
    End If
End Sub

Private Sub cmdAceptar_Click()
 Dim vSComb(3) As Double '0: nafta super T 1 - 1: nafta super t2 - 2: gasoil t 4 - Gasoil T3
' Dim vSneco As Double
' Dim vSgasoil1 As Double
' Dim vSgasoil2 As Double
 Dim vtid As Integer
 Dim i As Integer
 'Resume Error
    Dim cSQL As String
    
    'If Validar(vMode) = True Then
     If Validar = False Then Exit Sub
        
        'On Error GoTo ErrorTran
        
        Screen.MousePointer = vbHourglass
    
        'DBConn.BeginTrans
        
        If vEstado = "NUEVO" Then
            'NUEVO
            sql = "SELECT MAX(T_ID) AS MAXID FROM T_STOCK"
            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If rec.EOF = False Then
                vtid = IIf(IsNull(rec!MAXID), 1, rec!MAXID + 1)
            
            End If
            rec.Close
            txtId.Text = vtid
        
            'TENGO QUE HACER EL MODIFICAR!!!!! :S
            'Select Case vMode
            '    Case 1 'nuevo
            
            
            
            
            
            ' ACTUALIZO STOCKS DE TANQUES
            ' nafta manguera 1 tanque 2 txtNSLTot(0)
            ' nafta manguera 2 tanque 1 txtNSLTot(2)
            ' gasoil manguera 1 tanque 4 txtNSLTot(4)
            ' gasoil manguera 2 tanque 3 txtNSLTot(6)
            ' gasoil manguera 3 tanque 4 txtNSLTot(12)
            
            'ACTUALIZO TANQUES DE NAFTA
            sql = "SELECT * FROM PRODUCTO_DETALLE WHERE PTO_CODIGO = 1 ORDER BY PDT_CODIGO DESC"
            If Rec1.State = 1 Then
                Rec1.Close
            End If
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec1.EOF = False Then
                i = 0
                Do While Rec1.EOF = False
                
                    sql = "UPDATE PRODUCTO_DETALLE"
                    sql = sql & " SET PDT_CANTIDAD = " & XN(Rec1!PDT_CANTIDAD - CDbl(txtNSLTot(i).Text))
                    sql = sql & " WHERE PTO_CODIGO = " & Rec1!PTO_CODIGO
                    If i = 0 Then 'manguera 1 descuenta del tanque 2
                        sql = sql & " AND PDT_CODIGO = 2"
                    Else 'manguera 2 descuenta del tanque 1
                        sql = sql & " AND PDT_CODIGO = 1"
                    End If
                    DBConn.Execute sql
                    i = i + 2
                    Rec1.MoveNext
                    
                Loop
            End If
            Rec1.Close
            
            'ACTUALIZO TANQUES DE GASOIL
            sql = "SELECT * FROM PRODUCTO_DETALLE WHERE PTO_CODIGO = 3 ORDER BY PDT_CODIGO DESC"
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec1.EOF = False Then
                i = 4
                Do While Rec1.EOF = False
                    sql = "UPDATE PRODUCTO_DETALLE"
                    sql = sql & " SET PDT_CANTIDAD = " & XN(Rec1!PDT_CANTIDAD - CDbl(txtNSLTot(i).Text))
                    sql = sql & " WHERE PTO_CODIGO = " & Rec1!PTO_CODIGO
                    sql = sql & " AND PDT_CODIGO = " & Rec1!PDT_CODIGO
                    DBConn.Execute sql
                    i = i + 2
                    Rec1.MoveNext
                Loop
                Rec1.MoveFirst ' SE MUEVE AL TANQUE 4 DE GASOIL
                sql = "UPDATE PRODUCTO_DETALLE"
                sql = sql & " SET PDT_CANTIDAD = " & XN(Rec1!PDT_CANTIDAD - CDbl(txtNSLTot(12).Text))
                sql = sql & " WHERE PTO_CODIGO = " & Rec1!PTO_CODIGO
                sql = sql & " AND PDT_CODIGO = " & Rec1!PDT_CODIGO
                DBConn.Execute sql
            End If
            Rec1.Close
            
            
            
            
            'Busco el stock de Nafta Super Tanque 1, Nafta Super Tanque 2 y Gasoil T 1 y T 2
            sql = " SELECT P.PTO_CODIGO,PD.PDT_CODIGO, PD.PDT_CANTIDAD"
            sql = sql & " FROM PRODUCTO P, PRODUCTO_DETALLE PD"
            sql = sql & " WHERE"
            sql = sql & " P.PTO_CODIGO=PD.PTO_CODIGO"
            'sql = sql & " AND P.PTO_CODIGO =" & XN(txtcodigo.Text)
            sql = sql & " ORDER BY PD.PDT_CODIGO"
            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If rec.EOF = False Then
                i = 0
                Do While rec.EOF = False
                    vSComb(i) = Chk0(rec!PDT_CANTIDAD)
                    i = i + 1
                    rec.MoveNext
                    
                Loop
            End If
            rec.Close
            
            cSQL = "INSERT INTO T_STOCK "
            cSQL = cSQL & " (T_ID, T_FECHA, T_TURNO, VEN_CODIGO1,"
            cSQL = cSQL & " VEN_CODIGO2, T_OBSER, T_NSUPER1, T_NSUPERECO, T_GASOIL1,"
            cSQL = cSQL & " T_GNC1, T_LUB1, T_TOTAL1, T_RET,"
            cSQL = cSQL & " T_EFE, T_VALE, T_FACP, T_CTAC,"
            cSQL = cSQL & " T_TAR, T_DIARIO, T_CHEQ, T_DOLARES,"
            cSQL = cSQL & " T_VARIOS, T_DIFF,"
            'TAB STOCK
            'NAFTA SUPER
            cSQL = cSQL & " T_NSACTUAL, T_NSANTER, T_NSTOTAL,T_NSPESOS,"
            cSQL = cSQL & " T_NSACTUAL2, T_NSANTER2, T_NSTOTAL2,T_NSPESOS2,"
            'NAFTA ECOLOGICA
            cSQL = cSQL & " T_NEACTUAL, T_NEANTER, T_NETOTAL,T_NEPESOS,"
            cSQL = cSQL & " T_NEACTUAL2, T_NEANTER2, T_NETOTAL2,T_NEPESOS2,"
            'GASOIL 1
            cSQL = cSQL & " T_G1ACTUAL, T_G1ANTER, T_G1TOTAL,T_G1PESOS,"
            cSQL = cSQL & " T_G1ACTUAL2, T_G1ANTER2, T_G1TOTAL2,T_G1PESOS2,"
            'GASOIL 2
            cSQL = cSQL & " T_G2ACTUAL, T_G2ANTER, T_G2TOTAL,T_G2PESOS,"
            cSQL = cSQL & " T_G2ACTUAL2, T_G2ANTER2, T_G2TOTAL2,T_G2PESOS2,"
            'GNC1
            cSQL = cSQL & " T_GNC1ACTUAL, T_GNC1ANTER, T_GNC1TOTAL,T_GNC1PESOS,"
            'GNC2
            cSQL = cSQL & " T_GNC2ACTUAL, T_GNC2ANTER, T_GNC2TOTAL,T_GNC2PESOS,"
            'GNC3
            cSQL = cSQL & " T_GNC3ACTUAL, T_GNC3ANTER, T_GNC3TOTAL,T_GNC3PESOS,"
            'GNC4
            cSQL = cSQL & " T_GNC4ACTUAL, T_GNC4ANTER, T_GNC4TOTAL,T_GNC4PESOS,"
'
            'GASOIL 3
            cSQL = cSQL & " T_G3ACTUAL, T_G3ANTER, T_G3TOTAL,T_G3PESOS,"
            cSQL = cSQL & " T_G3ACTUAL2, T_G3ANTER2, T_G3TOTAL2,T_G3PESOS2,"
            
            cSQL = cSQL & " T_TREND,T_MEDDIG,T_MEDMEC,"
            
            'TOTALES - STOCK
            cSQL = cSQL & " T_SNSUPER,T_SNECO,T_SGASOIL1,T_SGASOIL2)"
                        
            cSQL = cSQL & " VALUES "
            'RENG_1
            cSQL = cSQL & " (" & XN(txtId.Text) & ", " & XDQ(Fecha) & ", "
            cSQL = cSQL & XS(cboTurno.Text) & ", "
            cSQL = cSQL & cboPla1.ItemData(cboPla1.ListIndex) & ", "
            'RENG_2
            If cboPla2.ListIndex = -1 Then
                cSQL = cSQL & 0 & ", "
            Else
                cSQL = cSQL & cboPla2.ItemData(cboPla2.ListIndex) & ", "
            End If
            cSQL = cSQL & XS(txtObservaciones.Text) & ", "
            cSQL = cSQL & XN(txtNSuper1.Text) & ", "
            cSQL = cSQL & XN(txtNaftaEco.Text) & ", "
            cSQL = cSQL & XN(txtGasOil1.Text) & ", "
            'RENG_3
            cSQL = cSQL & XN(txtGNC1.Text) & ", "
            cSQL = cSQL & XN(txtLub1.Text) & ", "
            cSQL = cSQL & XN(txtTotal1.Text) & ", "
            cSQL = cSQL & XN(txtRet.Text) & ", "
            'RENG_4
            cSQL = cSQL & XN(txtefe.Text) & ", "
            cSQL = cSQL & XN(txtVale.Text) & ", "
            cSQL = cSQL & XN(txtFacP.Text) & ", "
            cSQL = cSQL & XN(txtCtaC.Text) & ", "
            'RENG_5
            cSQL = cSQL & XN(txtTar.Text) & ", "
            cSQL = cSQL & XN(txtDiario.Text) & ", "
            cSQL = cSQL & XN(txtCheq.Text) & ", "
            cSQL = cSQL & XN(txtDolares.Text) & ", "
            'RENG_6
            cSQL = cSQL & XN(txtVarios.Text) & ", "
            cSQL = cSQL & XN(txtDiff.Text) & ", "
            For i = 0 To 13
                cSQL = cSQL & XN(txtNSLActual(i).Text) & ", "
                cSQL = cSQL & XN(txtNSLAnt(i).Text) & ", "
                cSQL = cSQL & XN(txtNSLTot(i).Text) & ", "
                cSQL = cSQL & XN(txtNSLPesos(i).Text) & ", " 'ojo aca cuando es el ultimo
            Next i
            cSQL = cSQL & XN(txtTRend.Text) & ", "
            cSQL = cSQL & XN(txtMedicion(0).Text) & ", "
            cSQL = cSQL & XN(txtMedicion(1).Text) & ", "
            
            'TOTALES - STOCK
            cSQL = cSQL & Replace(vSComb(0), ",", ".") & ", " 'STOCK NAFTA SUPER M1
            cSQL = cSQL & Replace(vSComb(1), ",", ".") & ", " 'STOCK NAFTA SUPER M2
            cSQL = cSQL & Replace(vSComb(3), ",", ".") & ", " 'STOCK GASOIL 1- Tanque 4
            cSQL = cSQL & Replace(vSComb(2), ",", ".") & ") " 'STOCK GASOIL 2- Tanque 3 'FALTA GASOIL 3, PERO SON 2 TANQUES ;)
            
            
            
            
            'ACTUALIZO STOCK FINAL LUBRICANTES PARA USARLO COMO STOCK INICIAL DEL TURNO SIGUIENTE
                          
             'DBConn.Execute "DELETE FROM TMP_LUBRICANTES_STOCKFINAL"
             
             For i = 1 To grdLubricantes.Rows - 1
                sql = "INSERT INTO TMP_LUBRICANTES_STOCKFINAL (PTO_CODIGO,PTO_DESCRI,PTO_PRECTO,INICIAL,ENTRO, TOTAL,"
                sql = sql & " PESOS,VENTA,FINAL,FECHA,TURNO, T_ID) "
                sql = sql & " VALUES ("
                sql = sql & XN(grdLubricantes.TextMatrix(i, 8))
                sql = sql & "," & XS(grdLubricantes.TextMatrix(i, 0))
                sql = sql & "," & XN(grdLubricantes.TextMatrix(i, 1))
                sql = sql & "," & XN(grdLubricantes.TextMatrix(i, 2))
                sql = sql & "," & XN(grdLubricantes.TextMatrix(i, 3))
                sql = sql & "," & XN(grdLubricantes.TextMatrix(i, 4))
                sql = sql & "," & XN(grdLubricantes.TextMatrix(i, 5))
                sql = sql & "," & XN(grdLubricantes.TextMatrix(i, 6))
                sql = sql & "," & XN(grdLubricantes.TextMatrix(i, 7))
                sql = sql & "," & XDQ(Fecha)
                sql = sql & "," & XS(cboTurno.Text)
                sql = sql & "," & XN(txtId.Text) & ")"
                DBConn.Execute sql
            Next i
                       
            
            
            'ESTE METODO ACTUALIZA LOS PARAMETROS DE TURNOS
            'ActualizoTurnos
        Else
        'MODIFICO - no se puede modificar por ahora
'            cSQL = "UPDATE T_STOCK SET"
'            cSQL = cSQL & " T_ID="
'            cSQL = cSQL & " T_FECHA=" & XDQ(Fecha)
'            cSQL = cSQL & ",T_TURNO=" & XS(txtTurno.Text)
'            cSQL = cSQL & ",VEN_CODIGO1=" & cboPla1.ItemData(cboPla1.ListIndex)
'            cSQL = cSQL & ",VEN_CODIGO2=" & cboPla2.ItemData(cboPla2.ListIndex)
'            cSQL = cSQL & ",T_OBSER=" & XS(txtObservaciones.Text)
'            cSQL = cSQL & ",T_NSUPER1=" & XN(txtNSuper1.Text)
'            cSQL = cSQL & ",T_GASOIL1=" & XN(txtGasOil1.Text)
'            cSQL = cSQL & ",T_GNC1=" & XN(txtGNC1.Text)
'            cSQL = cSQL & ",T_LUB1=" & XN(txtLub1.Text)
'            cSQL = cSQL & ",T_TOTAL1=" & XN(txtTotal1.Text)
'            cSQL = cSQL & ",T_RET=" & XN(txtRet.Text)
'            cSQL = cSQL & ",T_EFE=" & XN(txtefe.Text)
'            cSQL = cSQL & ",T_VALE=" & XN(txtVale.Text)
'            cSQL = cSQL & ",T_FACP=" & XN(txtFacP.Text)
'            cSQL = cSQL & ",T_CTAC=" & XN(txtCtaC.Text)
'            cSQL = cSQL & ",T_TAR=" & XN(txtTar.Text)
'            cSQL = cSQL & ",T_DIARIO=" & XN(txtDiario.Text)
'            cSQL = cSQL & ",T_CHEQ=" & XN(txtCheq.Text)
'            cSQL = cSQL & ",T_DOLARES=" & XN(txtDolares.Text)
'            cSQL = cSQL & ",T_VARIOS=" & XN(txtVarios.Text)
'            cSQL = cSQL & ",T_DIFF=" & XN(txtDiff.Text)
'            cSQL = cSQL & ",T_TREND=" & XN(txtTRend.Text)
'            cSQL = cSQL & " WHERE T_ID= " & XN(txtId.Text)
'
        End If
        
        
        
        DBConn.Execute cSQL
        
        'DBConn.CommitTrans
        'On Error GoTo 0
        
        'actualizo la lista base
        'ActualizarListaBase vMode
        Screen.MousePointer = vbDefault
        If MsgBox("Desea imprimir la Planilla?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
            cmdImprimir_Click
        End If
        CmdNuevo_Click
        'Unload Me
'    End If
    Exit Sub
    
'ErrorTran:
'
'    DBConn.RollbackTrans
'    Screen.MousePointer = vbDefault
'    'Resume Error
'    'manejo el error
'    'ManejoDeErrores DBConn.ErrorNative
'    MsgBox Err.Description, vbCritical
    
End Sub
Private Function SumarCaja()
    Dim totRend As Double
    totRend = CDbl(txtRet) + CDbl(txtefe) + _
                    CDbl(txtVale) + CDbl(txtFacP) + _
                    CDbl(txtCtaC) + CDbl(txtTar) + _
                    CDbl(txtDiario) + CDbl(txtCheq) + _
                    CDbl(txtDolares) + CDbl(txtVarios)
                    
    txtDiff.Text = CDbl(txtTRend.Text) - totRend
    txtDiff.Text = Format(txtDiff.Text, "#,##0.00")
    

End Function
Private Function ListarFacturas(pVend1 As Integer, pVend2 As Integer, pFecha As Date, pturno As Integer)
    'grdFacturas.FormatString = "^Tipo Comp|Numero|Playero|M3 GNC " & 1 & "|M3 GNC " & 2 & "|Lts Nafta|" _
                              & "Lts Gasoil|Aceites|Bar|Total Fac"
    Dim VTotal(6) As Double
    Dim vDesde As String
    Dim vHasta As String

    grdFacturas.Rows = 1
    grdFacturas.HighLight = flexHighlightNever
    'lblEstado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
    
    sql = "SELECT * FROM TURNOS WHERE TUR_CODIGO = " & pturno 'NOCHE
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    '***********************************************************************
    '***********************BUSCO FACTURAS DE COMBUSTIBLES******************
    '***********************************************************************
    sql = "SELECT DISTINCT TC.TCO_ABREVIA,FC.FCL_SUCURSAL,FC.FCL_FECHA,FC.FCL_NUMERO,V.VEN_NOMBRE"
    sql = sql & " ,DFC.DFC_CANTIDAD,DFC.PTO_CODIGO,FC.FCL_TOTAL"
    sql = sql & " FROM FACTURA_CLIENTE FC,DETALLE_FACTURA_CLIENTE DFC,"
    sql = sql & " TIPO_COMPROBANTE TC, VENDEDOR V"
    sql = sql & " WHERE"
    sql = sql & " FC.FCL_SUCURSAL = DFC.FCL_SUCURSAL"
    sql = sql & " AND FC.FCL_NUMERO = DFC.FCL_NUMERO"
    sql = sql & " AND FC.TCO_CODIGO = DFC.TCO_CODIGO"
    sql = sql & " AND FC.TCO_CODIGO=TC.TCO_CODIGO"
    sql = sql & " AND FC.VEN_CODIGO = V.VEN_CODIGO"
    sql = sql & " AND FC.EST_CODIGO <> 2"
    sql = sql & " AND DFC.PTO_CODIGO IN (1,2,3,4,78,81,84)" 'COMBUSTIBLES
    
'    If pVend1 <> 0 And pVend2 <> 0 Then
'        sql = sql & " AND FC.VEN_CODIGO IN (" & pVend1 & "," & pVend2 & ")"
'    Else
'        If pVend1 <> 0 Then
'            sql = sql & " AND FC.VEN_CODIGO = " & pVend1
'        Else
'            sql = sql & " AND FC.VEN_CODIGO = " & pVend2
'        End If
'    End If
    If pturno = 3 Then ' TURNO NOCHE
        sql = sql & " AND ((FC.FCL_FECHA= " & XDQ(pFecha) & ""
        sql = sql & " AND FC.FCL_HORA  >=#" & Rec1!TUR_DESDE & "#)" 'vDesde & "#)" '
        sql = sql & " OR (FC.FCL_FECHA= " & XDQ(pFecha + 1) & ""
        sql = sql & " AND FC.FCL_HORA  <=#" & Rec1!TUR_HASTA & "#))" 'vHasta & "#))" '
        'sql = sql & " AND FC.TUR_CODIGO=" & pTurno
    Else
        sql = sql & " AND FC.FCL_FECHA=" & XDQ(pFecha)
        sql = sql & " AND FC.FCL_HORA  >=#" & Rec1!TUR_DESDE & "# " 'vDesde & "#)" '
        sql = sql & " AND FC.FCL_HORA  <=#" & Rec1!TUR_HASTA & "# " 'vHasta & "#))" '
        
        'sql = sql & " AND FC.TUR_CODIGO=" & pTurno
    End If

'    Rec1.Close
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    DBConn.Execute "DELETE FROM TMP_FACTURAS"
    
    If rec.EOF = False Then
        Do While rec.EOF = False
            cSQL = "INSERT INTO TMP_FACTURAS"
            cSQL = cSQL & " (TCO_CODIGO,FCL_SUCURSAL,FCL_NUMERO,FCL_FECHA,"
            cSQL = cSQL & " M3GNC_1,M3GNC_2,LNAFTA,LNAFTAE,GASOIL,"
            cSQL = cSQL & " TOTALF,VEN_CODIGO)" 'USO TMP AUX1 PARA EL PLAYERO
            cSQL = cSQL & " VALUES (" & XS(rec!TCO_ABREVIA) & ", "
            cSQL = cSQL & XS(Format(rec!FCL_SUCURSAL, "0000")) & ", "
            cSQL = cSQL & XS(Format(rec!FCL_NUMERO, "00000000")) & ", "
            cSQL = cSQL & XDQ(rec!FCL_FECHA) & ", "
            cSQL = cSQL & IIf(rec!PTO_CODIGO = 2, Replace(Chk0(rec!DFC_CANTIDAD), ",", "."), "0.00") & ", "
            cSQL = cSQL & IIf(rec!PTO_CODIGO = 4, Replace(Chk0(rec!DFC_CANTIDAD), ",", "."), "0.00") & ", "
            cSQL = cSQL & IIf((rec!PTO_CODIGO = 1) Or (rec!PTO_CODIGO = 84), Replace(Chk0(rec!DFC_CANTIDAD), ",", "."), "0.00") & ", "
            cSQL = cSQL & IIf(rec!PTO_CODIGO = 78, Replace(Chk0(rec!DFC_CANTIDAD), ",", "."), "0.00") & ", "
            cSQL = cSQL & IIf((rec!PTO_CODIGO = 3) Or (rec!PTO_CODIGO = 81), Replace(Chk0(rec!DFC_CANTIDAD), ",", "."), "0.00") & ", "
            cSQL = cSQL & XN(Chk0(rec!FCL_TOTAL)) & ", "
            cSQL = cSQL & XS(Chk0(rec!VEN_NOMBRE)) & ") "
            DBConn.Execute cSQL
            rec.MoveNext
        Loop
    End If
    rec.Close
    '***********************************************************************
    '***********************BUSCO FACTURAS DE LUBRICANTES Y DEMAS***********
    '***********************************************************************
    sql = "SELECT DISTINCT TC.TCO_ABREVIA,FC.FCL_SUCURSAL,FC.FCL_FECHA,FC.FCL_NUMERO,V.VEN_NOMBRE"
    sql = sql & " ,FC.FCL_TOTAL"
    sql = sql & " FROM FACTURA_CLIENTE FC,DETALLE_FACTURA_CLIENTE DFC,"
    sql = sql & " TIPO_COMPROBANTE TC, VENDEDOR V"
    sql = sql & " WHERE"
    sql = sql & " FC.FCL_SUCURSAL = DFC.FCL_SUCURSAL"
    sql = sql & " AND FC.FCL_NUMERO = DFC.FCL_NUMERO"
    sql = sql & " AND FC.TCO_CODIGO = DFC.TCO_CODIGO"
    sql = sql & " AND FC.TCO_CODIGO=TC.TCO_CODIGO"
    sql = sql & " AND FC.VEN_CODIGO = V.VEN_CODIGO"
    sql = sql & " AND FC.EST_CODIGO <> 2"
    sql = sql & " AND DFC.PTO_CODIGO NOT IN (1,2,3,4,78,81,84)" 'LUBRICANTES Y DEMAS
    If pturno = 3 Then ' TURNO NOCHE
        sql = sql & " AND ((FC.FCL_FECHA= " & XDQ(pFecha) & ""
        sql = sql & " AND FC.FCL_HORA  >=#" & Rec1!TUR_DESDE & "#)" 'vDesde & "#)" '
        sql = sql & " OR (FC.FCL_FECHA= " & XDQ(pFecha + 1) & ""
        sql = sql & " AND FC.FCL_HORA  <=#" & Rec1!TUR_HASTA & "#))" 'vHasta & "#))" '
        'sql = sql & " AND FC.TUR_CODIGO=" & pTurno
    Else
        sql = sql & " AND FC.FCL_FECHA=" & XDQ(pFecha)
        sql = sql & " AND FC.FCL_HORA  >=#" & Rec1!TUR_DESDE & "# " 'vDesde & "#)" '
        sql = sql & " AND FC.FCL_HORA  <=#" & Rec1!TUR_HASTA & "# " 'vHasta & "#))" '
        
        'sql = sql & " AND FC.TUR_CODIGO=" & pTurno
    End If
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = False Then
        Do While rec.EOF = False
            cSQL = "INSERT INTO TMP_FACTURAS"
            cSQL = cSQL & " (TCO_CODIGO,FCL_SUCURSAL,FCL_NUMERO,FCL_FECHA,"
            cSQL = cSQL & " M3GNC_1,M3GNC_2,LNAFTA,LNAFTAE,GASOIL,"
            cSQL = cSQL & " TOTALF,VEN_CODIGO)"
            cSQL = cSQL & " VALUES (" & XS(rec!TCO_ABREVIA) & ", "
            cSQL = cSQL & XS(Format(rec!FCL_SUCURSAL, "0000")) & ", "
            cSQL = cSQL & XS(Format(rec!FCL_NUMERO, "00000000")) & ", "
            cSQL = cSQL & XDQ(rec!FCL_FECHA) & ", "
            cSQL = cSQL & "0.00" & ", "
            cSQL = cSQL & "0.00" & ", "
            cSQL = cSQL & "0.00" & ", "
            cSQL = cSQL & "0.00" & ", "
            cSQL = cSQL & "0.00" & ", "
            cSQL = cSQL & XN(Chk0(rec!FCL_TOTAL)) & ", "
            cSQL = cSQL & XS(Chk0(rec!VEN_NOMBRE)) & ") "
            DBConn.Execute cSQL
            rec.MoveNext
        Loop
    End If
    rec.Close
    Rec1.Close
    
    'CONSULTO Y MUESTRO LA TABLA TEMPORAL
    sql = "SELECT * FROM TMP_FACTURAS "
    sql = sql & "ORDER BY TCO_CODIGO,FCL_NUMERO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    i = 0
    If rec.EOF = False Then
        Do While rec.EOF = False
                    
            grdFacturas.AddItem rec!TCO_CODIGO & Chr(9) & Format(rec!FCL_SUCURSAL, "0000") & "-" & Format(rec!FCL_NUMERO, "00000000") _
                            & Chr(9) & rec!VEN_CODIGO _
                            & Chr(9) & rec!M3GNC_1 _
                            & Chr(9) & rec!M3GNC_2 _
                            & Chr(9) & rec!LNAFTA _
                            & Chr(9) & rec!LNAFTAE _
                            & Chr(9) & rec!GASOIL _
                            & Chr(9) & VALIDO_IMPORTE(Chk0(rec!TOTALF))
            VTotal(0) = VTotal(0) + CDbl(VALIDO_IMPORTE(Chk0(rec!M3GNC_1)))
            VTotal(1) = VTotal(1) + CDbl(VALIDO_IMPORTE(Chk0(rec!M3GNC_2)))
            VTotal(2) = VTotal(2) + CDbl(VALIDO_IMPORTE(Chk0(rec!LNAFTA)))
            VTotal(3) = VTotal(3) + CDbl(VALIDO_IMPORTE(Chk0(rec!LNAFTAE)))
            VTotal(4) = VTotal(4) + CDbl(VALIDO_IMPORTE(Chk0(rec!GASOIL)))
            VTotal(5) = VTotal(5) + CDbl(VALIDO_IMPORTE(Chk0(rec!TOTALF)))
            i = i + 1
            rec.MoveNext
            
        Loop
        grdFacturas.HighLight = flexHighlightAlways
        grdFacturas.SetFocus
        grdFacturas.Col = 0
        grdFacturas.row = 1

    End If
    rec.Close
    
    'lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    For i = 0 To 5
        txtTotFac(i).Text = VTotal(i)
        txtTotFac(i).Text = VALIDO_IMPORTE(txtTotFac(i).Text)
    Next
End Function
Private Function ListarLubricantes(pVend1 As Integer, pVend2 As Integer, pFecha As Date, pturno As Integer, Estado As String)
    'grdFacturas.FormatString = "^Tipo Comp|Numero|Playero|M3 GNC " & 1 & "|M3 GNC " & 2 & "|Lts Nafta|" _
                              & "Lts Gasoil|Aceites|Bar|Total Fac"
    Dim VTotal As Double
    'grdLubricantes.Rows = 1
    
    grdLubricantes.HighLight = flexHighlightNever
    'lblEstado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
    ' primero cargo todos los productos
    buscaUltimoTurno
        
    If Estado = "NUEVO" Then
        
        Lubricantes
        
        sql = "SELECT * FROM TURNOS WHERE TUR_CODIGO = " & pturno
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
        'segundo, actualizo los valores vendidos
        sql = "SELECT P.PTO_DESCRI,P.PTO_PRECTO,SUM(DFC.DFC_CANTIDAD) AS CANTIDAD"
        sql = sql & " ,P.LNA_CODIGO,DFC.PTO_CODIGO"
        sql = sql & " FROM FACTURA_CLIENTE FC,DETALLE_FACTURA_CLIENTE DFC,"
        sql = sql & " PRODUCTO P,VENDEDOR V"
        sql = sql & " WHERE"
        sql = sql & " FC.FCL_SUCURSAL = DFC.FCL_SUCURSAL"
        sql = sql & " AND FC.FCL_NUMERO = DFC.FCL_NUMERO"
        sql = sql & " AND FC.TCO_CODIGO = DFC.TCO_CODIGO"
        sql = sql & " AND FC.VEN_CODIGO = V.VEN_CODIGO"
        sql = sql & " AND P.PTO_CODIGO = DFC.PTO_CODIGO"
        sql = sql & " AND P.LNA_CODIGO <> 1 AND P.RUB_CODIGO NOT IN(12,13) "
    '    If pVend1 <> 0 And pVend2 <> 0 Then
    '        sql = sql & " AND FC.VEN_CODIGO IN (" & pVend1 & "," & pVend2 & ")"
    '    Else
    '        If pVend1 <> 0 Then
    '            sql = sql & " AND FC.VEN_CODIGO = " & pVend1
    '        Else
    '            sql = sql & " AND FC.VEN_CODIGO = " & pVend2
    '        End If
    '    End If
        If pturno = 3 Then ' TURNO NOCHE
            If Date = pFecha Then
                pFecha = pFecha - 1 'VUELVO UN DIA ARTRAS CUANDO CIERRAN TURNO NOCHE CON FECHA DEL DIAS SIGUIENTE
            End If
            sql = sql & " AND ((FC.FCL_FECHA= " & XDQ(pFecha) & ""
            sql = sql & " AND FC.FCL_HORA  >=#" & Rec1!TUR_DESDE & "#)" 'vDesde & "#)" '
            sql = sql & " OR (FC.FCL_FECHA= " & XDQ(pFecha + 1) & ""
            sql = sql & " AND FC.FCL_HORA  <=#" & Rec1!TUR_HASTA & "#))" 'vHasta & "#))" '
            'sql = sql & " AND FC.TUR_CODIGO=" & pTurno
        Else
            sql = sql & " AND FC.FCL_FECHA=" & XDQ(pFecha)
            sql = sql & " AND FC.FCL_HORA  >=#" & Rec1!TUR_DESDE & "# " 'vDesde & "#)" '
            sql = sql & " AND FC.FCL_HORA  <=#" & Rec1!TUR_HASTA & "# " 'vHasta & "#))" '
            
            'sql = sql & " AND FC.TUR_CODIGO=" & pTurno
        End If
    
        Rec1.Close
        
        sql = sql & " GROUP BY P.PTO_DESCRI,P.PTO_PRECTO,P.LNA_CODIGO,DFC.PTO_CODIGO"
        sql = sql & " ORDER BY P.PTO_DESCRI"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Do While rec.EOF = False
                'Actualizo grilla Lubricantes
                For i = 1 To grdLubricantes.Rows - 1
                    If grdLubricantes.TextMatrix(i, 8) = rec!PTO_CODIGO Then
                        'grdLubricantes.TextMatrix(i, 2) = StockFinal(rec!PTO_CODIGO) '(CDbl(grdLubricantes.TextMatrix(i, 2)) + rec!CANTIDAD) ' - CDbl(grdLubricantes.TextMatrix(i, 3)) '(StockActual+Venta) - Entro
                        grdLubricantes.TextMatrix(i, 4) = StockFinal(rec!PTO_CODIGO) + CDbl(grdLubricantes.TextMatrix(i, 3)) 'TOTAL = INICAL + ENTRO (3)
                        grdLubricantes.TextMatrix(i, 5) = rec!cantidad 'VENDIDO
                        grdLubricantes.TextMatrix(i, 6) = CDbl(grdLubricantes.TextMatrix(i, 4)) - rec!cantidad 'FINAL
                        grdLubricantes.TextMatrix(i, 7) = VALIDO_IMPORTE(rec!PTO_PRECTO * rec!cantidad) 'IMPORTE
                        Exit For
                    End If
                Next i
                VTotal = VTotal + CDbl(VALIDO_IMPORTE(rec!PTO_PRECTO * rec!cantidad))
                rec.MoveNext
            Loop
            grdLubricantes.HighLight = flexHighlightAlways
            grdLubricantes.SetFocus
            grdLubricantes.Col = 0
            grdLubricantes.row = 1
        Else
    '        If vEstado = "NUEVO" Then
    '            Screen.MousePointer = vbNormal
    '            MsgBox "No se encontraron Lubricantes...", vbExclamation, TIT_MSGBOX
    '        End If
        End If
        Screen.MousePointer = vbNormal
        rec.Close
        txtTotLub.Text = VTotal
        txtTotLub.Text = VALIDO_IMPORTE(txtTotLub.Text)
        txtLub1.Text = VALIDO_IMPORTE(txtTotLub.Text)
    Else
    
        sql = "SELECT PTO_DESCRI,PTO_PRECTO,INICIAL, ENTRO, TOTAL, VENTA, FINAL,PESOS,PTO_CODIGO"
        sql = sql & " FROM TMP_LUBRICANTES_STOCKFINAL"
        sql = sql & " WHERE T_ID=" & XN(txtId.Text)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            i = 1
            Do While rec.EOF = False
                'Actualizo grilla Lubricantes
                    grdLubricantes.Rows = i + 1
                    grdLubricantes.TextMatrix(i, 0) = rec!PTO_DESCRI
                    grdLubricantes.TextMatrix(i, 1) = Chk0(PTO_PRECTO)
                    grdLubricantes.TextMatrix(i, 2) = Chk0(rec!INICIAL)
                    grdLubricantes.TextMatrix(i, 3) = Chk0(rec!ENTRO)
                    grdLubricantes.TextMatrix(i, 4) = Chk0(rec!TOTAL)
                    grdLubricantes.TextMatrix(i, 5) = Chk0(rec!VENTA)
                    grdLubricantes.TextMatrix(i, 6) = Chk0(rec!FINAL)
                    grdLubricantes.TextMatrix(i, 7) = Chk0(rec!PESOS)
                    grdLubricantes.TextMatrix(i, 9) = Chk0(rec!PTO_CODIGO)
                    i = i + 1
                VTotal = VTotal + CDbl(VALIDO_IMPORTE(Chk0(rec!PESOS)))
                rec.MoveNext
            Loop
            grdLubricantes.HighLight = flexHighlightAlways
            grdLubricantes.SetFocus
            grdLubricantes.Col = 0
            'grdLubricantes.row = 1

        End If
        Screen.MousePointer = vbNormal
        rec.Close
        txtTotLub.Text = VTotal
        txtTotLub.Text = VALIDO_IMPORTE(txtTotLub.Text)
        txtLub1.Text = VALIDO_IMPORTE(txtTotLub.Text)
    
    
    End If
End Function
Private Function StockFinal(Codigo As Integer) As Double
    'Busco el stock final del turno anterior de los Lubricantes
    sql = "SELECT FINAL FROM TMP_LUBRICANTES_STOCKFINAL WHERE T_ID=" & vUltTurno & " AND PTO_CODIGO=" & Codigo
    Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec2.EOF = False Then
        StockFinal = Rec2!FINAL
    Else
        StockFinal = 0
    End If
    Rec2.Close
    
End Function

Private Function HayEntrada(Fecha As Date, turno As Integer) As Boolean
    Dim Rec2 As New ADODB.Recordset
    sql = "SELECT * FROM TURNOS WHERE TUR_CODIGO = " & turno
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    sql = "SELECT DE.DEP_CANTIDAD FROM ENTRADA_PRODUCTO E, DETALLE_ENTRADA_PRODUCTO DE"
    sql = sql & " WHERE DE.EPR_CODIGO=E.EPR_CODIGO"
    sql = sql & " AND DE.PTO_CODIGO NOT IN (1,2,3,4,78,81)" 'NO ES PARA COMBUSTIBLES
    
     If turno = 3 Then ' TURNO NOCHE
        If Date = Fecha Then
            Fecha = Fecha - 1 'VUELVO UN DIA ARTRAS CUANDO CIERRAN TURNO NOCHE CON FECHA DEL DIAS SIGUIENTE
        End If
        sql = sql & " AND ((E.EPR_FECHA= " & XDQ(Fecha) & ""
        sql = sql & " AND E.EPR_HORA  >=#" & Rec1!TUR_DESDE & "#)" 'vDesde & "#)" '
        sql = sql & " OR (E.EPR_FECHA= " & XDQ(Fecha + 1) & ""
        sql = sql & " AND E.EPR_HORA  <=#" & Rec1!TUR_HASTA & "#))" 'vHasta & "#))" '
        'sql = sql & " AND FC.TUR_CODIGO=" & pTurno
    Else
        sql = sql & " AND E.EPR_FECHA=" & XDQ(Fecha)
        sql = sql & " AND E.EPR_HORA  >=#" & Rec1!TUR_DESDE & "# " 'vDesde & "#)" '
        sql = sql & " AND E.EPR_HORA  <=#" & Rec1!TUR_HASTA & "# " 'vHasta & "#))" '
        
        'sql = sql & " AND FC.TUR_CODIGO=" & pTurno
    End If
    
    Rec1.Close
    Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec2.EOF = False Then
        HayEntrada = True
    Else
        HayEntrada = False
    End If
    Rec2.Close
    
End Function
Private Function BuscoEntPto(pto As Integer, Fecha As Date, turno As Integer) As Double
    Dim Rec2 As New ADODB.Recordset
    sql = "SELECT * FROM TURNOS WHERE TUR_CODIGO = " & turno
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    sql = "SELECT DE.DEP_CANTIDAD FROM ENTRADA_PRODUCTO E, DETALLE_ENTRADA_PRODUCTO DE"
    sql = sql & " WHERE DE.EPR_CODIGO=E.EPR_CODIGO"
    sql = sql & " AND DE.PTO_CODIGO = " & pto
    sql = sql & " AND DE.PTO_CODIGO NOT IN (1,2,3,4,78,81)" 'NO ES PARA COMBUSTIBLES
    
     If turno = 3 Then ' TURNO NOCHE
        If Date = Fecha Then
            Fecha = Fecha - 1 'VUELVO UN DIA ARTRAS CUANDO CIERRAN TURNO NOCHE CON FECHA DEL DIAS SIGUIENTE
        End If
        sql = sql & " AND ((E.EPR_FECHA= " & XDQ(Fecha) & ""
        sql = sql & " AND E.EPR_HORA  >=#" & Rec1!TUR_DESDE & "#)" 'vDesde & "#)" '
        sql = sql & " OR (E.EPR_FECHA= " & XDQ(Fecha + 1) & ""
        sql = sql & " AND E.EPR_HORA  <=#" & Rec1!TUR_HASTA & "#))" 'vHasta & "#))" '
        'sql = sql & " AND FC.TUR_CODIGO=" & pTurno
    Else
        sql = sql & " AND E.EPR_FECHA=" & XDQ(Fecha)
        sql = sql & " AND E.EPR_HORA  >=#" & Rec1!TUR_DESDE & "# " 'vDesde & "#)" '
        sql = sql & " AND E.EPR_HORA  <=#" & Rec1!TUR_HASTA & "# " 'vHasta & "#))" '
        
        'sql = sql & " AND FC.TUR_CODIGO=" & pTurno
    End If
    
    Rec1.Close
    Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec2.EOF = False Then
        BuscoEntPto = Rec2!DEP_CANTIDAD
    Else
        BuscoEntPto = 0
    End If
    Rec2.Close
    
End Function
Private Function LlenatTMPFacturas()
    Dim i As Integer
    
    DBConn.Execute "DELETE FROM TMP_FACTURAS"
    For i = 1 To grdFacturas.Rows - 1
        cSQL = "INSERT INTO TMP_FACTURAS"
        cSQL = cSQL & " (TCO_CODIGO,FCL_SUCURSAL,FCL_NUMERO,FCL_FECHA,"
        cSQL = cSQL & " VEN_CODIGO,M3GNC_1,M3GNC_2,LNAFTA,LNAFTAE,GASOIL,"
        cSQL = cSQL & " TOTALF,TURNO,TMPFAC_AUX1,TMPFAC_AUX2,TMPFAC_AUX3)"
        cSQL = cSQL & " VALUES (" & XS(grdFacturas.TextMatrix(i, 0)) & ", "
        cSQL = cSQL & XS(Left(grdFacturas.TextMatrix(i, 1), 4)) & ", "
        cSQL = cSQL & XS(Right(grdFacturas.TextMatrix(i, 1), 8)) & ", "
        cSQL = cSQL & XDQ(Fecha) & ", "
        cSQL = cSQL & XS(grdFacturas.TextMatrix(i, 2)) & ", "
        cSQL = cSQL & XN(grdFacturas.TextMatrix(i, 3)) & ", "
        cSQL = cSQL & XN(grdFacturas.TextMatrix(i, 4)) & ", "
        cSQL = cSQL & XN(grdFacturas.TextMatrix(i, 5)) & ", "
        cSQL = cSQL & XN(grdFacturas.TextMatrix(i, 6)) & ", "
        cSQL = cSQL & XN(grdFacturas.TextMatrix(i, 7)) & ", "
        cSQL = cSQL & XN(grdFacturas.TextMatrix(i, 8)) & ", "
        cSQL = cSQL & XS(cboTurno.Text) & ", "
        cSQL = cSQL & XS("Aux_1") & ", "
        cSQL = cSQL & XS("Aux_2") & ", "
        cSQL = cSQL & XS("Aux_2") & ") "
        DBConn.Execute cSQL
    Next i

End Function
Private Function Lubricantes()
    Dim hayMov As Boolean
    
    grdLubricantes.Rows = 1
    'hago esta funcion para que no entre siempre a buscar movimientos de mercaderias
    hayMov = HayEntrada(Fecha.Value, cboTurno.ItemData(cboTurno.ListIndex))
    
    sql = "SELECT P.PTO_DESCRI,P.PTO_PRECTO,P.PTO_CODIGO,S.DST_STKFIS"
    sql = sql & " FROM PRODUCTO P, STOCK S"
    sql = sql & " WHERE"
    sql = sql & " P.PTO_CODIGO = S.PTO_CODIGO"
    sql = sql & " AND P.LNA_CODIGO <> 1 AND P.RUB_CODIGO NOT IN(12,13)"
    sql = sql & " ORDER BY P.PTO_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            grdLubricantes.AddItem rec!PTO_DESCRI & Chr(9) & VALIDO_IMPORTE(rec!PTO_PRECTO) _
                            & Chr(9) & StockFinal(rec!PTO_CODIGO) _
                            & Chr(9) & IIf(hayMov = True, BuscoEntPto(rec!PTO_CODIGO, Fecha.Value, cboTurno.ItemData(cboTurno.ListIndex)), 0) _
                            & Chr(9) & StockFinal(rec!PTO_CODIGO) _
                            & Chr(9) & "0" _
                            & Chr(9) & StockFinal(rec!PTO_CODIGO) _
                            & Chr(9) & "0" _
                            & Chr(9) & rec!PTO_CODIGO _
                            & Chr(9) & StockFinal(rec!PTO_CODIGO)
            rec.MoveNext
        
        
        Loop
        grdLubricantes.HighLight = flexHighlightAlways
        grdLubricantes.SetFocus
        grdLubricantes.Col = 0
        grdLubricantes.row = 1
    End If
    Screen.MousePointer = vbNormal
    rec.Close
    txtTotLub.Text = "0,00"
End Function
Private Function LlenatTMPLubricantes()
    Dim i As Integer
    DBConn.Execute "DELETE FROM TMP_LUBRICANTES"
    For i = 1 To grdLubricantes.Rows - 1
        cSQL = "INSERT INTO TMP_LUBRICANTES"
        cSQL = cSQL & " (PTO_CODIGO,PTO_DESCRI,PTO_PRECTO,INICIAL,ENTRO,"
        cSQL = cSQL & " TOTAL,VENTA,FINAL,TURNO,FECHA,TMPLUB_AUX1,TMPLUB_AUX2,TMPLUB_AUX3)"
        cSQL = cSQL & " VALUES (" & XN(grdLubricantes.TextMatrix(i, 8)) & ", "
        cSQL = cSQL & XS(grdLubricantes.TextMatrix(i, 0)) & ", "
        cSQL = cSQL & XN(grdLubricantes.TextMatrix(i, 1)) & ", "
        cSQL = cSQL & XN(grdLubricantes.TextMatrix(i, 2)) & ", "
        cSQL = cSQL & XN(grdLubricantes.TextMatrix(i, 3)) & ", "
        cSQL = cSQL & XN(grdLubricantes.TextMatrix(i, 4)) & ", "
        cSQL = cSQL & XN(grdLubricantes.TextMatrix(i, 5)) & ", "
        cSQL = cSQL & XN(grdLubricantes.TextMatrix(i, 6)) & ", "
        cSQL = cSQL & XS(cboTurno.Text) & ", "
        cSQL = cSQL & XDQ(Fecha) & ", "
        cSQL = cSQL & XS(cboTurno.Text) & ", "
        cSQL = cSQL & XS("Aux_2") & ", "
        cSQL = cSQL & XS("Aux_2") & ") "
        DBConn.Execute cSQL
    Next i
End Function
Private Function SumarResumen()
    txtTotal1.Text = CDbl(txtNSuper1) + CDbl(txtGasOil1) + _
                    CDbl(txtGNC1) + CDbl(txtLub1)
    txtTotal1.Text = Format(txtTotal1.Text, "#,##0.00")
    txtTRend.Text = Format(txtTotal1.Text, "#,##0.00")
    If txtDiff.Text = "0,00" Then
        txtDiff.Text = Format(txtTotal1.Text, "#,##0.00")
    End If
End Function

Private Sub CmdCancelar_Click()
    If MsgBox("Seguro que desea Salir", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
    'Set frmPlanillaStock = Nothing
    Unload Me
    End If
End Sub

Private Sub cmdEliminar_Click()
    If MsgBox("Confirma eliminar la Planilla?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        cSQL = "DELETE FROM T_STOCK WHERE T_ID  = " & XN(txtId.Text)
        DBConn.Execute cSQL
        CmdNuevo_Click
    End If
End Sub

Private Sub cmdGetDatos_Click()
  'Ver que hago cuando hay un solo playero y desde donde ejecutar esa funcion
    Dim vPla1 As Integer
    Dim vPla2 As Integer
    
       
    'Limpio stock
    Dim i As Integer
    If vEstado = "NUEVO" Then
        For i = 0 To 13
            txtNSLActual(i).Text = "0,00"
            txtNSLAnt(i).Text = "0,00"
            txtNSLTot(i).Text = "0,00"
            txtNSLPesos(i).Text = "0,00"
        Next
        txtMedicion(0).Text = "0,00"
        txtMedicion(1).Text = "0,00"
    End If
    
    If cboTurno.ListIndex <> -1 Then
        If cboPla1.ListIndex = -1 Then vPla1 = 0 Else vPla1 = cboPla1.ItemData(cboPla1.ListIndex)
        If cboPla2.ListIndex = -1 Then vPla2 = 0 Else vPla2 = cboPla2.ItemData(cboPla2.ListIndex)
        ListarFacturas vPla1, vPla2, _
                       Fecha, cboTurno.ItemData(cboTurno.ListIndex)
        
        
        ListarLubricantes vPla1, vPla2, _
                          Date, cboTurno.ItemData(cboTurno.ListIndex), vEstado
        'buscar stocks del turno anterior
        If vEstado = "NUEVO" Then
            BuscoTurnoAnterior
        End If
    End If
    cmdGetDatos.Enabled = False
End Sub

Private Sub cmdImprimir_Click()

    'Imprimir Planilla Stock 1
    ListarPlanillaStock1
    'Imprimir Facturas
    ListarPlanillaFacturas
    'Imprimir Lubricantes
    ListarPlanillaLubricantes
    
End Sub
Private Function ListarPlanillaStock1()
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
    If txtId.Text <> "" Then
        Rep.SelectionFormula = " {T_STOCK.T_ID}=" & txtId.Text
    End If

    
    Rep.WindowState = crptMaximized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR
    
    Rep.WindowTitle = "Planilla Diaria - Stock"
    Rep.ReportFileName = DRIVE & DirReport & "PlanillaStock_1.rpt"
    Rep.Action = 1
End Function
Private Function ListarPlanillaFacturas()
    LlenatTMPFacturas
    
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
   
'    If txtId.Text <> "" Then
'        Rep.SelectionFormula = " {T_STOCK.T_ID}=" & txtId.Text
'    End If
   
    Rep.WindowState = crptMaximized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR
    
    Rep.WindowTitle = "Planilla Diaria - Facturas"
    Rep.ReportFileName = DRIVE & DirReport & "PlanillaStock_Fac.rpt"
    Rep.Action = 1
End Function
Private Function ListarPlanillaLubricantes()
    LlenatTMPLubricantes
    
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""

    Rep.WindowState = crptMaximized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR
    
    Rep.WindowTitle = "Planilla Diaria - Lubricantes"
    Rep.ReportFileName = DRIVE & DirReport & "PlanillaStock_Lub.rpt"
    Rep.Action = 1
End Function
Private Sub cmdListar_Click()
    sql = "SELECT DISTINCT T.T_TURNO,T.T_FECHA,V1.VEN_NOMBRE AS PLA1,T.VEN_CODIGO2,T.T_TOTAL1,T.T_ID"
    sql = sql & " FROM T_STOCK T, VENDEDOR V1"
    sql = sql & " WHERE T.VEN_CODIGO1 = V1.VEN_CODIGO"
    
    If FechaDesde.Value <> "" Then
        sql = sql & " AND T.T_FECHA>=" & XDQ(FechaDesde.Value)
    End If
    If FechaHasta.Value <> "" Then
        sql = sql & " AND T.T_FECHA<=" & XDQ(FechaHasta.Value)
    End If
    sql = sql & " ORDER BY T.T_FECHA,T.T_TURNO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    GrdModulos.Rows = 1
    If rec.EOF = False Then
        Do While rec.EOF = False
            
            GrdModulos.AddItem rec!T_TURNO & Chr(9) & rec!T_FECHA & Chr(9) & _
                               rec!PLA1 & Chr(9) & IIf(IsNull(rec!VEN_CODIGO2), "", BuscoVend(rec!VEN_CODIGO2)) & Chr(9) & _
                               VALIDO_IMPORTE(rec!T_TOTAL1) & Chr(9) & rec!T_ID
                               
            rec.MoveNext
        Loop
    End If
    rec.Close
    
End Sub
Private Function BuscoVend(Codigo As Integer) As String
    sql = "SELECT VEN_NOMBRE FROM VENDEDOR WHERE VEN_CODIGO=" & Codigo
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        BuscoVend = Rec1!VEN_NOMBRE
    Else
        BuscoVend = ""
    End If
    Rec1.Close
    
End Function

Private Sub CmdNuevo_Click()
    vEstado = "NUEVO"
    cmdImprimir.Enabled = False
    cmdAceptar.Enabled = True
    txtNSuper1.Text = "0,00"
    txtNaftaEco.Text = "0,00"
    txtGasOil1.Text = "0,00"
    txtGNC1.Text = "0,00"
    txtLub1.Text = "0,00"
    txtTotal1.Text = "0,00"
    
    txtRet.Text = "0,00"
    txtefe.Text = "0,00"
    txtVale.Text = "0,00"
    txtFacP.Text = "0,00"
    txtCtaC.Text = "0,00"
    txtTar.Text = "0,00"
    txtDiario.Text = "0,00"
    txtCheq.Text = "0,00"
    txtDolares.Text = "0,00"
    txtVarios.Text = "0,00"
    txtDiff.Text = "0,00"
    txtTRend.Text = "0,00"
    
    txtObservaciones.Text = ""
    cboPla1.ListIndex = -1
    cboPla2.ListIndex = -1
    Fecha = Date
    txtTurno.Text = ""
    txtId.Text = ""
    
    tabDatos.Tab = 0
    GrdModulos.Rows = 1
    
    Dim i As Integer
    For i = 0 To 13
        txtNSLActual(i).Text = "0,00"
        txtNSLAnt(i).Text = "0,00"
        txtNSLTot(i).Text = "0,00"
        txtNSLPesos(i).Text = "0,00"
    Next
    txtMedicion(0).Text = "0,00"
    txtMedicion(1).Text = "0,00"
    grdFacturas.Rows = 1
    grdLubricantes.Rows = 1
    
    cmdGetDatos.Enabled = True
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        Dim vPla1 As Integer
        Dim vPla2 As Integer
        If cboTurno.ListIndex <> -1 Then
            If cboPla1.ListIndex = -1 Then vPla1 = 0 Else vPla1 = cboPla1.ItemData(cboPla1.ListIndex)
            If cboPla2.ListIndex = -1 Then vPla2 = 0 Else vPla2 = cboPla2.ItemData(cboPla2.ListIndex)
            ListarFacturas vPla1, vPla2, _
                           Fecha, cboTurno.ItemData(cboTurno.ListIndex)
            ListarLubricantes vPla1, vPla2, _
                              Fecha, cboTurno.ItemData(cboTurno.ListIndex), vEstado
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
    
    If KeyAscii = 27 Then
        CmdCancelar_Click
    End If
End Sub

Private Sub Form_Load()
    Centrar_pantalla Me
    'Me.Top = 0
    Fecha.Value = Date
    CargoComboBox cboPla1, "VENDEDOR", "VEN_CODIGO", "VEN_NOMBRE", "VEN_NOMBRE"
    CargoComboBox cboPla2, "VENDEDOR", "VEN_CODIGO", "VEN_NOMBRE", "VEN_NOMBRE"
    CargoComboBox cboPlaBus, "VENDEDOR", "VEN_CODIGO", "VEN_NOMBRE", "VEN_NOMBRE"
    LlenarComboTurnos
    preparogrillas
    
    For i = 0 To 13
        txtNSLActual(i).Text = "0,00"
        txtNSLAnt(i).Text = "0,00"
        txtNSLTot(i).Text = "0,00"
        txtNSLPesos(i).Text = "0,00"
    Next
    vEstado = "NUEVO"
    
    'deshabilito la edicion de anterior
    
    For i = 0 To 13
        txtNSLAnt(i).Locked = True
    Next i
    If mNomUser = "A" Then
        fraResumen.Enabled = True
    End If
    
'    sql = "UPDATE T_STOCK"
'    sql = sql & " SET"
'    sql = sql & " T_G3ACTUAL=0"
'    sql = sql & ", T_G3ANTER=0"
'    sql = sql & ", T_G3TOTAL=0"
'    sql = sql & ", T_G3PESOS=0"
'    sql = sql & ", T_G3ACTUAL2=0"
'    sql = sql & ", T_G3ANTER2=0"
'    sql = sql & ", T_G3TOTAL2=0"
'    sql = sql & ", T_G3PESOS2=0"
'
'    DBConn.Execute sql
        
    
    
End Sub
Private Sub LlenarComboTurnos()
    sql = "SELECT * FROM TURNOS"
    sql = sql & " ORDER BY TUR_CODIGO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboTurno.AddItem ""
        Do While rec.EOF = False
            cboTurno.AddItem rec!TUR_DESCRI
            cboTurno.ItemData(cboTurno.NewIndex) = rec!TUR_CODIGO
            rec.MoveNext
        Loop
    End If
    rec.Close

    Dim vDesde(3) As Date
    Dim vHasta(3) As Date
    Dim i As Integer
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
        
    Else
        If Time() >= vDesde(1) And Time() <= vHasta(1) Then
            Call BuscaCodigoProxItemData(2, cboTurno)
            
        Else
            Call BuscaCodigoProxItemData(3, cboTurno)
            
        End If
    End If

End Sub
Private Function preparogrillas()
    'Grilla grdFacturas
    grdFacturas.FormatString = "^Tipo Comp|Numero|Playero|M3 GNC " & 1 & "|M3 GNC " & 2 & "|Lts NS - M1|" _
                              & "Lts NS - M2|Lts Gasoil|Total Fac"
    grdFacturas.ColWidth(0) = 1000 'TURNO
    grdFacturas.ColWidth(1) = 1200 'Numero
    grdFacturas.ColWidth(2) = 1800 'PLAYERO
    grdFacturas.ColWidth(3) = 1000 'M3 GNC 1
    grdFacturas.ColWidth(4) = 1000 'Metros GNC a 2
    grdFacturas.ColWidth(5) = 1000 'Litros Nafta
    grdFacturas.ColWidth(6) = 1000 'Litros Nafta Ecologica
    grdFacturas.ColWidth(7) = 1000 'Litros Gasoil
    grdFacturas.ColWidth(8) = 1000 'Importe Tot Fact.
    grdFacturas.Cols = 9
    grdFacturas.Rows = 1
    grdFacturas.HighLight = flexHighlightNever
    grdFacturas.BorderStyle = flexBorderNone
    grdFacturas.row = 0
    For i = 0 To grdFacturas.Cols - 1
        grdFacturas.Col = i
        grdFacturas.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        grdFacturas.CellBackColor = &H808080    'GRIS OSCURO
        grdFacturas.CellFontBold = True
    Next
    tabDatos.Tab = 0
    
    'Grilla grdLubricantes
    'grdLubricantes.FormatString = "^Producto|^Precio|^Inicial|^Total|^Venta|^Final|PTO_CODIGO|STK ACTUAL "
    grdLubricantes.FormatString = "Producto|^Precio|^Inicial|^Entro|^Total|^Venta|^Final|^Importe|PTO_CODIGO|STK ACTUAL "
    grdLubricantes.ColWidth(0) = 3200 'PRODUCTO
    grdLubricantes.ColWidth(1) = 1000 'PRECIO
    grdLubricantes.ColWidth(2) = 1000 'INICIAL
    grdLubricantes.ColWidth(3) = 1000 'ENTRO
    grdLubricantes.ColWidth(4) = 1000 'TOTAL
    grdLubricantes.ColWidth(5) = 1000 'VENTA
    grdLubricantes.ColWidth(6) = 1000 'FINAL
    grdLubricantes.ColWidth(7) = 1000 'IMPORTE
    grdLubricantes.ColWidth(8) = 0 'PTO_CODIGO
    grdLubricantes.ColWidth(9) = 0 'STK ACTUAL
    grdLubricantes.Cols = 10
    grdLubricantes.Rows = 1
    grdLubricantes.HighLight = flexHighlightNever
    grdLubricantes.BorderStyle = flexBorderNone
    grdLubricantes.row = 0
    For i = 0 To grdLubricantes.Cols - 1
        grdLubricantes.Col = i
        grdLubricantes.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        grdLubricantes.CellBackColor = &H808080    'GRIS OSCURO
        grdLubricantes.CellFontBold = True
    Next
    tabDatos.Tab = 0
    
    'Grilla Busqcar Anteriores
    GrdModulos.FormatString = "^Turno|Fecha|Encargado 1|Encargado 2|Total|id|TurnoID"
    GrdModulos.ColWidth(0) = 1200 'TURNO
    GrdModulos.ColWidth(1) = 1200 'FECHA
    GrdModulos.ColWidth(2) = 3000 'ENCARGADO 1
    GrdModulos.ColWidth(3) = 3000 'ENCARGADO 2
    GrdModulos.ColWidth(4) = 1500    'TOTAL
    GrdModulos.ColWidth(5) = 0    'ID
    GrdModulos.ColWidth(6) = 0    'turno ID
    GrdModulos.Cols = 7
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
    tabDatos.Tab = 0
End Function

Private Sub GrdModulos_dblClick()
    vEstado = "CONSULTA"
    If mNomUser = "A" Then
        cmdEliminar.Enabled = True
    End If
    cmdImprimir.Enabled = True
    cmdAceptar.Enabled = False
    sql = "SELECT *"
    sql = sql & " FROM T_STOCK "
    sql = sql & " WHERE T_ID = " & GrdModulos.TextMatrix(GrdModulos.RowSel, 5)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        txtNSuper1.Text = VALIDO_IMPORTE(Chk0(rec!T_NSUPER1))
        txtGasOil1.Text = VALIDO_IMPORTE(Chk0(rec!T_GASOIL1))
        txtGNC1.Text = VALIDO_IMPORTE(Chk0(rec!T_GNC1))
        txtLub1.Text = VALIDO_IMPORTE(Chk0(rec!T_LUB1))
        txtTotal1.Text = VALIDO_IMPORTE(Chk0(rec!T_TOTAL1))
        
        txtRet.Text = VALIDO_IMPORTE(Chk0(rec!T_RET))
        txtefe.Text = VALIDO_IMPORTE(Chk0(rec!T_EFE))
        txtVale.Text = VALIDO_IMPORTE(Chk0(rec!T_VALE))
        txtFacP.Text = VALIDO_IMPORTE(Chk0(rec!T_FACP))
        txtCtaC.Text = VALIDO_IMPORTE(Chk0(rec!T_CTAC))
        txtTar.Text = VALIDO_IMPORTE(Chk0(rec!T_TAR))
        txtDiario.Text = VALIDO_IMPORTE(Chk0(rec!T_DIARIO))
        txtCheq.Text = VALIDO_IMPORTE(Chk0(rec!T_CHEQ))
        txtDolares.Text = VALIDO_IMPORTE(Chk0(rec!T_DOLARES))
        txtVarios.Text = VALIDO_IMPORTE(Chk0(rec!T_VARIOS))
        txtDiff.Text = VALIDO_IMPORTE(Chk0(rec!T_DIFF))
        txtTRend.Text = VALIDO_IMPORTE(Chk0(rec!T_TREND))
        
        txtObservaciones.Text = ChkNull(rec!T_OBSER)
        Call BuscaCodigoProxItemData(Chk0(rec!VEN_CODIGO1), cboPla1)
        If rec!VEN_CODIGO2 <> 0 Then
            Call BuscaCodigoProxItemData(Chk0(rec!VEN_CODIGO2), cboPla2)
        Else
            cboPla2.ListIndex = -1
        End If
        'cboPla1.ListIndex = -1
        'cboPla2.ListIndex = -1
        Fecha = ChkNull(rec!T_FECHA)
        'txtTurno.Text = ChkNull(rec!T_TURNO)
        cboTurno.Text = ChkNull(rec!T_TURNO)
        txtId.Text = Chk0(rec!T_ID)
        
        'mostrar campos de tab stock
        Dim i, J, z As Integer
        i = 0
        J = 24
        For i = 0 To 13
            txtNSLActual(i).Text = IIf(IsNull(rec.Fields(J)), "", rec.Fields(J))
            txtNSLAnt(i).Text = IIf(IsNull(rec.Fields(J + 1)), "", rec.Fields(J + 1))
            txtNSLTot(i).Text = IIf(IsNull(rec.Fields(J + 2)), "", rec.Fields(J + 2))
            txtNSLPesos(i).Text = IIf(IsNull(rec.Fields(J + 3)), "", rec.Fields(J + 3))
            txtNSLActual(i).Text = VALIDO_IMPORTE(txtNSLActual(i).Text)
            txtNSLAnt(i).Text = VALIDO_IMPORTE(txtNSLAnt(i).Text)
            txtNSLTot(i).Text = VALIDO_IMPORTE(txtNSLTot(i).Text)
            txtNSLPesos(i).Text = VALIDO_IMPORTE(txtNSLPesos(i).Text)
            
            J = J + 4
        Next i
        
    End If
    rec.Close
    cmdGetDatos_Click
    'ListarFacturas grdModulos.TextMatrix(grdModulos.RowSel, 2), grdModulos.TextMatrix(grdModulos.RowSel, 3), _
                       grdModulos.TextMatrix(grdModulos.RowSel, 1), cboTurno.ItemData(cboTurno.ListIndex)
    'ListarLubricantes grdModulos.TextMatrix(grdModulos.RowSel, 2), grdModulos.TextMatrix(grdModulos.RowSel, 3), _
                       grdModulos.TextMatrix(grdModulos.RowSel, 1), cboTurno.ItemData(cboTurno.ListIndex)
    tabDatos.Tab = 0
    
    
End Sub

Private Sub grdModulos_GotFocus()
    GrdModulos.Col = 0
    GrdModulos.ColSel = 1
    GrdModulos.HighLight = flexHighlightAlways
End Sub



Private Sub txtCheq_GotFocus()
    SelecTexto txtCheq
End Sub

Private Sub txtCheq_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtCheq, KeyAscii)
End Sub

Private Sub txtCheq_LostFocus()
    If txtCheq = "" Then
        txtCheq.Text = "0,00"
    End If
    txtCheq.Text = VALIDO_IMPORTE(txtCheq)
    SumarCaja
End Sub

Private Sub txtCtaC_GotFocus()
    SelecTexto txtCtaC
End Sub

Private Sub txtCtaC_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtCtaC, KeyAscii)
End Sub

Private Sub txtCtaC_LostFocus()
    If txtCtaC = "" Then
        txtCtaC.Text = "0,00"
    End If
    txtCtaC.Text = VALIDO_IMPORTE(txtCtaC)
    SumarCaja
End Sub

Private Sub txtDiario_GotFocus()
    SelecTexto txtDiario
End Sub

Private Sub txtDiario_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtDiario, KeyAscii)
End Sub

Private Sub txtDiario_LostFocus()
    If txtDiario = "" Then
        txtDiario.Text = "0,00"
    End If
    txtDiario.Text = VALIDO_IMPORTE(txtDiario)
    SumarCaja
End Sub

Private Sub txtDiff_GotFocus()
    SelecTexto txtDiff
End Sub

Private Sub txtDiff_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtDiff, KeyAscii)
End Sub

Private Sub txtDiff_LostFocus()
    If txtDiff = "" Then
        txtDiff.Text = "0,00"
    End If
    txtDiff.Text = VALIDO_IMPORTE(txtDiff)
    SumarCaja
End Sub

Private Sub txtDolares_GotFocus()
    SelecTexto txtDolares
End Sub

Private Sub txtDolares_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtDolares, KeyAscii)
End Sub

Private Sub txtDolares_LostFocus()
    If txtDolares = "" Then
        txtDolares.Text = "0,00"
    End If
    txtDolares.Text = VALIDO_IMPORTE(txtDolares)
    SumarCaja
End Sub

Private Sub txtefe_GotFocus()
    SelecTexto txtefe
End Sub

Private Sub txtefe_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtefe, KeyAscii)
End Sub

Private Sub txtefe_LostFocus()
    If txtefe = "" Then
        txtefe.Text = "0,00"
    End If
    txtefe.Text = VALIDO_IMPORTE(txtefe)
    SumarCaja
End Sub

Private Sub txtFacP_GotFocus()
    SelecTexto txtFacP
End Sub

Private Sub txtFacP_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtFacP, KeyAscii)
End Sub

Private Sub txtFacP_LostFocus()
    If txtFacP = "" Then
        txtFacP.Text = "0,00"
    End If
    txtFacP.Text = VALIDO_IMPORTE(txtFacP)
    SumarCaja
End Sub

Private Sub txtGasOil1_GotFocus()
    SelecTexto txtGasOil1
End Sub

Private Sub txtGasOil1_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtGasOil1, KeyAscii)
End Sub

Private Sub txtGasOil1_LostFocus()
    If txtGasOil1 = "" Then
        txtGasOil1.Text = "0,00"
    End If
    txtGasOil1.Text = VALIDO_IMPORTE(txtGasOil1)
    SumarResumen
End Sub

Private Sub txtGNC1_GotFocus()
    SelecTexto txtGNC1
End Sub

Private Sub txtGNC1_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtGNC1, KeyAscii)
End Sub

Private Sub txtGNC1_LostFocus()
    If txtGNC1 = "" Then
        txtGNC1.Text = "0,00"
    End If
    txtGNC1.Text = VALIDO_IMPORTE(txtGNC1)
    SumarResumen
End Sub



Private Sub txtLub1_GotFocus()
    SelecTexto txtLub1
End Sub

Private Sub txtLub1_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtLub1, KeyAscii)
End Sub

Private Sub txtLub1_LostFocus()
    If txtLub1 = "" Then
        txtLub1.Text = "0,00"
    End If
    txtLub1.Text = VALIDO_IMPORTE(txtLub1)
    SumarResumen
End Sub

'Private Sub txtNSLActual_GotFocus()
'    seltxt
'End Sub
'
'Private Sub txtNSLActual_KeyPress(KeyAscii As Integer)
'    KeyAscii = CarNumeroDecimal(txtNSLActual, KeyAscii)
'End Sub

'Private Sub txtNSLAnt_GotFocus()
'    seltxt
'End Sub
'
'Private Sub txtNSLAnt_KeyPress(KeyAscii As Integer)
'    KeyAscii = CarNumeroDecimal(txtNSLAnt, KeyAscii)
'End Sub

'Private Sub txtNSLAnt_LostFocus()
'    If txtNSLAnt.Text <> "" Then
'        'Calculo el total de la venta
'        txtNSLTot.Text = CDbl(txtNSLActual.Text - txtNSLAnt.Text)
'        txtNSLTot.Text = VALIDO_IMPORTE(txtNSLTot.Text)
'        'calculo el monto de la venta
'        'Busco precio nafta super
'        txtNSLPesos.Text = CDbl(txtNSLTot.Text) * BuscoPrecio(1)
'        txtNSLPesos.Text = VALIDO_IMPORTE(txtNSLPesos.Text)
'    End If
'End Sub
Private Function BuscoPrecio(Codigo As Integer) As Double
    Dim producto As Integer
    '1: nafta super
    '2: gnc
    '3: gasoil
    '4: gnc
    '78: nafta ecologica
    Select Case Codigo
    Case Is = 0, 1
        producto = 1
    Case Is = 2, 3
        producto = 78
    Case Is = 4, 5, 6, 7, 12, 13
        producto = 3
    Case Is = 8, 9, 10, 11
        producto = 2
    End Select
    
    
    'codigo tiene el indicedel producto
    BuscoPrecio = 0
    sql = "SELECT PTO_PRECTO FROM PRODUCTO WHERE PTO_CODIGO =" & producto
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        BuscoPrecio = rec!PTO_PRECTO
    End If
    rec.Close
End Function

Private Sub txtNSNActual_GotFocus()
    seltxt
End Sub

Private Sub txtNSNAnt_GotFocus()
    seltxt
End Sub

Private Sub txtMedicion_GotFocus(index As Integer)
    seltxt
End Sub

Private Sub txtMedicion_KeyPress(index As Integer, KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtMedicion(index), KeyAscii)
End Sub

Private Sub txtNaftaEco_GotFocus()
    SelecTexto txtNaftaEco
End Sub

Private Sub txtNSLActual_GotFocus(index As Integer)
    seltxt
End Sub

Private Sub txtNSLActual_KeyPress(index As Integer, KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtNSLActual(index), KeyAscii)
End Sub

Private Sub txtNSLActual_LostFocus(index As Integer)
    If txtNSLActual(index).Text <> "" And txtNSLAnt(index).Text <> "" Then
        
            If CDbl(txtNSLActual(index).Text) < CDbl(txtNSLAnt(index).Text) Then
                If index < 10 Then
                    txtNSLTot(index).Text = CDbl(1000000 + CDbl(txtNSLActual(index).Text) - txtNSLAnt(index).Text)
                Else
                    txtNSLTot(index).Text = CDbl(100000 + CDbl(txtNSLActual(index).Text) - txtNSLAnt(index).Text)
                End If
                txtNSLTot(index).Text = VALIDO_IMPORTE(txtNSLTot(index).Text)
            Else
                txtNSLTot(index).Text = CDbl(txtNSLActual(index).Text - txtNSLAnt(index).Text)
                txtNSLTot(index).Text = VALIDO_IMPORTE(txtNSLTot(index).Text)
            End If
        
        'Calculo el total de la venta
        
        'calculo el monto de la venta
        'Busco precio nafta super
        txtNSLPesos(index).Text = CDbl(txtNSLTot(index).Text) * BuscoPrecio(index)
        txtNSLPesos(index).Text = VALIDO_IMPORTE(txtNSLPesos(index).Text)
        Select Case index
        Case 0, 2
            txtNSuper1.Text = VALIDO_IMPORTE(CDbl(txtNSLPesos(0).Text) + CDbl(txtNSLPesos(2).Text))
        'Case 2 'm3 gnc
        
       Case 4, 6, 12
            txtGasOil1.Text = VALIDO_IMPORTE(CDbl(Chk0(txtNSLPesos(4).Text)) + CDbl(Chk0(txtNSLPesos(6).Text)) + CDbl(Chk0(txtNSLPesos(12).Text)))
        Case 8, 9, 10, 11
            txtGNC1.Text = VALIDO_IMPORTE(CDbl(Chk0(txtNSLPesos(8).Text)) + CDbl(Chk0(txtNSLPesos(9).Text)) + _
                                             CDbl(Chk0(txtNSLPesos(10).Text)) + CDbl(Chk0(txtNSLPesos(11).Text)))
            'm3 gnc
            txtNaftaEco.Text = VALIDO_IMPORTE(CDbl(Chk0(txtNSLTot(8).Text)) + CDbl(Chk0(txtNSLTot(9).Text)) + _
                                             CDbl(Chk0(txtNSLTot(10).Text)) + CDbl(Chk0(txtNSLTot(11).Text)))
        End Select
        txtNSuper1_LostFocus
    End If
    txtNSLActual(index).Text = VALIDO_IMPORTE(txtNSLActual(index).Text)
End Sub

Private Sub txtNSLAnt_GotFocus(index As Integer)
    SelecTexto txtNSLAnt(index)
End Sub

Private Sub txtNSLAnt_KeyPress(index As Integer, KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtNSLAnt(index).Text, KeyAscii)
End Sub

Private Sub txtNSLAnt_LostFocus(index As Integer)
    If txtNSLActual(index).Text <> "" And txtNSLAnt(index).Text <> "" Then
        'Calculo el total de la venta
        If CDbl(txtNSLActual(index).Text) < CDbl(txtNSLAnt(index).Text) Then
            txtNSLTot(index).Text = CDbl(100000 + CDbl(txtNSLActual(index).Text) - txtNSLAnt(index).Text)
            txtNSLTot(index).Text = VALIDO_IMPORTE(txtNSLTot(index).Text)
        Else
            txtNSLTot(index).Text = CDbl(txtNSLActual(index).Text - txtNSLAnt(index).Text)
            txtNSLTot(index).Text = VALIDO_IMPORTE(txtNSLTot(index).Text)
        End If
        'calculo el monto de la venta
        'Busco precio nafta super
        txtNSLPesos(index).Text = CDbl(txtNSLTot(index).Text) * BuscoPrecio(index)
        txtNSLPesos(index).Text = VALIDO_IMPORTE(txtNSLPesos(index).Text)
        Select Case index
        Case 0, 2
            txtNSuper1.Text = VALIDO_IMPORTE(CDbl(txtNSLPesos(0).Text) + CDbl(txtNSLPesos(2).Text))
        'Case 2
        
        Case 4, 6, 12
            txtGasOil1.Text = VALIDO_IMPORTE(CDbl(Chk0(txtNSLPesos(4).Text)) + CDbl(Chk0(txtNSLPesos(6).Text)))
        Case 8, 9, 10, 11
            txtGNC1.Text = VALIDO_IMPORTE(CDbl(Chk0(txtNSLPesos(8).Text)) + CDbl(Chk0(txtNSLPesos(9).Text)) + _
                                             CDbl(Chk0(txtNSLPesos(10).Text)) + CDbl(Chk0(txtNSLPesos(11).Text)))
            'm3 gnc
            txtNaftaEco.Text = VALIDO_IMPORTE(CDbl(Chk0(txtNSLTot(8).Text)) + CDbl(Chk0(txtNSLTot(9).Text)) + _
                                             CDbl(Chk0(txtNSLTot(10).Text)) + CDbl(Chk0(txtNSLTot(11).Text)))
        End Select
        txtNSuper1_LostFocus
    End If
    
End Sub

Private Sub txtNSuper1_GotFocus()
    SelecTexto txtNSuper1
End Sub

Private Sub txtNSuper1_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtNSuper1, KeyAscii)
End Sub

Private Sub txtNSuper1_LostFocus()
    If txtNSuper1 = "" Then
        txtNSuper1.Text = "0,00"
    End If
    txtNSuper1.Text = VALIDO_IMPORTE(txtNSuper1)
    SumarResumen
End Sub

Private Sub txtObservaciones_GotFocus()
    SelecTexto txtObservaciones
End Sub

Private Sub txtRet_GotFocus()
    SelecTexto txtRet
End Sub

Private Sub txtRet_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtRet, KeyAscii)
End Sub

Private Sub txtRet_LostFocus()
    If txtRet = "" Then
        txtRet.Text = "0,00"
    End If
    txtRet.Text = VALIDO_IMPORTE(txtRet)
    SumarCaja
End Sub

Private Sub txtTar_GotFocus()
    SelecTexto txtTar
End Sub

Private Sub txtTar_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtTar, KeyAscii)
End Sub

Private Sub txtTar_LostFocus()
    If txtTar = "" Then
        txtTar.Text = "0,00"
    End If
    txtTar.Text = VALIDO_IMPORTE(txtTar)
    SumarCaja
End Sub

Private Sub txtTurno_GotFocus()
    SelecTexto txtTurno
End Sub

Private Sub txtVale_GotFocus()
    SelecTexto txtVale
End Sub

Private Sub txtVale_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtVale, KeyAscii)
End Sub

Private Sub txtVale_LostFocus()
    If txtVale = "" Then
        txtVale.Text = "0,00"
    End If
    txtVale.Text = VALIDO_IMPORTE(txtVale)
    SumarCaja
End Sub

Private Sub txtVarios_GotFocus()
    SelecTexto txtVarios
End Sub

Private Sub txtVarios_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtVarios, KeyAscii)
End Sub

Private Sub txtVarios_LostFocus()
    If txtVarios = "" Then
        txtVarios.Text = "0,00"
    End If
    txtVarios.Text = VALIDO_IMPORTE(txtVarios)
    SumarCaja
End Sub
Private Function ActualizoTurnos()
    Dim vDesde(3) As Date
    Dim vHasta(3) As Date
    Dim i As Integer
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
        'actualizo turno mañana
        ActualizoHora 1
    Else
        If Time() >= vDesde(1) And Time() <= vHasta(1) Then
            'actualizo turno tarde
            ActualizoHora 2
        Else
            'actualizo turno noche
            ActualizoHora 3
        End If
    End If
End Function
Private Function ActualizoHora(pturno As Integer)
    Dim pturnosig As Integer
    
    Select Case pturno
    Case 1
      pturnosig = 2
    Case 2
      pturnosig = 3
    Case 3
      pturnosig = 1
    End Select
    
    
    'ACTUALIZO HORA HASTA TURNO
    sql = "UPDATE TURNOS SET TUR_HASTA = " & XS(Format(Time(), "hh:mm")) & " WHERE TUR_CODIGO = " & pturno
    DBConn.Execute sql
    
    'ACTUALIZO HORA DESDE TURNO SIGUIENTE
    sql = "UPDATE TURNOS SET TUR_DESDE = " & XS(Format(Time(), "hh:mm")) & " WHERE TUR_CODIGO = " & pturnosig
    DBConn.Execute sql
End Function
