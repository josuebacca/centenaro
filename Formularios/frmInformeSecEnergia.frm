VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmInformeSecEnergia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GNC-Informe para Secretaria de Energia"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
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
   ScaleHeight     =   8355
   ScaleWidth      =   7920
   Begin VB.CommandButton cmdtxt 
      Caption         =   "&Generar txt"
      Enabled         =   0   'False
      Height          =   345
      Left            =   3120
      TabIndex        =   55
      Top             =   7980
      Width           =   1215
   End
   Begin VB.CommandButton cmdexportar 
      Caption         =   "Exportar"
      Height          =   345
      Left            =   4350
      TabIndex        =   54
      Top             =   7980
      Width           =   855
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   345
      Left            =   6105
      Picture         =   "frmInformeSecEnergia.frx":0000
      TabIndex        =   26
      ToolTipText     =   "Imprimir lista de Precios"
      Top             =   7980
      Width           =   870
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   120
      TabIndex        =   31
      Top             =   -75
      Width           =   7750
      Begin VB.CommandButton cmdListar 
         Height          =   300
         Left            =   7440
         Picture         =   "frmInformeSecEnergia.frx":27A2
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   255
      End
      Begin VB.ComboBox cboAnio 
         Height          =   315
         Left            =   6600
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   120
         Width           =   780
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         Left            =   5040
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   120
         Width           =   1140
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "de:"
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
         Left            =   6240
         TabIndex        =   53
         Top             =   180
         Width           =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "* GNC-Informe para Secretaria de Energia:"
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
         Left            =   120
         TabIndex        =   33
         Top             =   172
         Width           =   3870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Mes:"
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
         Left            =   4560
         TabIndex        =   32
         Top             =   180
         Width           =   390
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   345
      Left            =   6990
      Picture         =   "frmInformeSecEnergia.frx":2B2C
      TabIndex        =   27
      Top             =   7980
      Width           =   870
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   345
      Left            =   5220
      Picture         =   "frmInformeSecEnergia.frx":2E36
      TabIndex        =   25
      ToolTipText     =   "Imprimir lista de Precios"
      Top             =   7980
      Width           =   870
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
      Left            =   1635
      TabIndex        =   24
      Text            =   "A"
      Top             =   7950
      Visible         =   0   'False
      Width           =   390
   End
   Begin MSFlexGridLib.MSFlexGrid GrdGNC 
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   2143
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
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   2160
      Top             =   7920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSFlexGridLib.MSFlexGrid grdNafta 
      Height          =   1215
      Left            =   120
      TabIndex        =   23
      Top             =   2760
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   2143
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
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid grdGasoil 
      Height          =   1215
      Left            =   120
      TabIndex        =   15
      Top             =   4800
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   2143
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
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid grdTOTAL 
      Height          =   1215
      Left            =   120
      TabIndex        =   22
      Top             =   6720
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   2143
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
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   34
      Top             =   1860
      Width           =   7770
      Begin VB.TextBox txtGNCNum 
         Height          =   315
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   180
         Width           =   1095
      End
      Begin VB.TextBox txtGNCPPPsI 
         BackColor       =   &H00FFFFC0&
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
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   180
         Width           =   1095
      End
      Begin VB.TextBox txtGNCPPPcI 
         BackColor       =   &H00FFFFC0&
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
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   180
         Width           =   1095
      End
      Begin VB.TextBox txtGNCDen 
         Height          =   315
         Left            =   5085
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "PPP c/I ="
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
         Left            =   3000
         TabIndex        =   39
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "PPP s/I ="
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
         Left            =   240
         TabIndex        =   38
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6240
         TabIndex        =   37
         Top             =   195
         Width           =   195
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4920
         TabIndex        =   36
         Top             =   180
         Width           =   135
      End
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   120
      TabIndex        =   35
      Top             =   3880
      Width           =   7770
      Begin VB.TextBox txtNaPPPcIden 
         Height          =   315
         Left            =   5805
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtNaftaPPPcI 
         BackColor       =   &H00FFFFC0&
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
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtNaPPPcInum 
         Height          =   315
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtNaPPPsIDEN 
         Height          =   315
         Left            =   1965
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtNAFTAPPPsI 
         BackColor       =   &H00FFFFC0&
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
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtNaPPPsINUM 
         Height          =   315
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5640
         TabIndex        =   45
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6650
         TabIndex        =   44
         Top             =   255
         Width           =   195
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "PPP c/I ="
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
         Left            =   3960
         TabIndex        =   43
         Top             =   300
         Width           =   795
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   42
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2820
         TabIndex        =   41
         Top             =   255
         Width           =   195
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "PPP s/I ="
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
         Left            =   120
         TabIndex        =   40
         Top             =   300
         Width           =   795
      End
   End
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   120
      TabIndex        =   46
      Top             =   5920
      Width           =   7770
      Begin VB.TextBox txtGaPPPsINUM 
         Height          =   315
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtGASOILPPPsI 
         BackColor       =   &H00FFFFC0&
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
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtGaPPPsIDEN 
         Height          =   315
         Left            =   1965
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtGaPPPcInum 
         Height          =   315
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtGasoilPPPcI 
         BackColor       =   &H00FFFFC0&
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
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtGaPPPcIden 
         Height          =   315
         Left            =   5805
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "PPP s/I ="
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
         Left            =   120
         TabIndex        =   52
         Top             =   300
         Width           =   795
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2820
         TabIndex        =   51
         Top             =   255
         Width           =   195
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   50
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "PPP c/I ="
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
         Left            =   3960
         TabIndex        =   49
         Top             =   300
         Width           =   795
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6650
         TabIndex        =   48
         Top             =   255
         Width           =   195
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5640
         TabIndex        =   47
         Top             =   240
         Width           =   135
      End
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "GASOIL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   4560
      Width           =   7755
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NAFTA SUPER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   2520
      Width           =   7755
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "GNC"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   480
      Width           =   7755
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
      Left            =   105
      TabIndex        =   8
      Top             =   7950
      Width           =   660
   End
End
Attribute VB_Name = "FrmInformeSecEnergia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CodigoProducto As String
Dim J As Integer
Dim i As Integer
Dim vMes As Date
Private Function lleno_tmp()
    sql = "DELETE FROM TMP_INFORME"
    DBConn.Execute sql
    
    sql = "DELETE FROM TMP_INFORME_RESUMEN"
    DBConn.Execute sql
    
    'GNC
    For i = 1 To GrdGNC.Rows - 2
        sql = "INSERT INTO TMP_INFORME (PTO_CODIGO,PTO_DESCRI,FCL_FECHA,PTO_NETO,PTO_PVENTA,PTO_CVEND,PTO_TOTPESOS)"
        sql = sql & " VALUES("
        sql = sql & XN("1") & ","
        sql = sql & XS("GNC") & ","
        sql = sql & XS(cboMes.Text & "/" & cboAnio.Text) & ","
        sql = sql & XN(GrdGNC.TextMatrix(i, 1)) & ","
        sql = sql & XN(GrdGNC.TextMatrix(i, 2)) & ","
        sql = sql & XN(GrdGNC.TextMatrix(i, 3)) & ","
        sql = sql & XN(GrdGNC.TextMatrix(i, 4)) & ")"
        DBConn.Execute sql
    Next
    
    'nafta
    For i = 1 To grdNafta.Rows - 2
        sql = "INSERT INTO TMP_INFORME (PTO_CODIGO,PTO_DESCRI,FCL_FECHA,PTO_NETO,PTO_PVENTA,PTO_CVEND,PTO_TOTPESOS)"
        sql = sql & " VALUES("
        sql = sql & XN("2") & ","
        sql = sql & XS("NAFTA") & ","
        sql = sql & XS(cboMes.Text & "/" & cboAnio.Text) & ","
        sql = sql & XN(grdNafta.TextMatrix(i, 1)) & ","
        sql = sql & XN(grdNafta.TextMatrix(i, 2)) & ","
        sql = sql & XN(grdNafta.TextMatrix(i, 3)) & ","
        sql = sql & XN(grdNafta.TextMatrix(i, 4)) & ")"
        DBConn.Execute sql
    Next
    
    'gasoil
    For i = 1 To grdGasoil.Rows - 2
        sql = "INSERT INTO TMP_INFORME (PTO_CODIGO,PTO_DESCRI,FCL_FECHA,PTO_NETO,PTO_PVENTA,PTO_CVEND,PTO_TOTPESOS)"
        sql = sql & " VALUES("
        sql = sql & XN("3") & ","
        sql = sql & XS("GASOIL") & ","
        sql = sql & XS(cboMes.Text & "/" & cboAnio.Text) & ","
        sql = sql & XN(grdGasoil.TextMatrix(i, 1)) & ","
        sql = sql & XN(grdGasoil.TextMatrix(i, 2)) & ","
        sql = sql & XN(grdGasoil.TextMatrix(i, 3)) & ","
        sql = sql & XN(grdGasoil.TextMatrix(i, 4)) & ")"
        DBConn.Execute sql
    Next
    
    
    'RESUMEN
    For i = 1 To grdTOTAL.Rows - 1
        sql = "INSERT INTO TMP_INFORME_RESUMEN "
        sql = sql & "(PTO_CODIGO,PPPsi,PPPci,VOLUMEN,PFINAL)"
        sql = sql & " VALUES("
        sql = sql & i & ","
        sql = sql & XN(grdTOTAL.TextMatrix(i, 2)) & ","
        sql = sql & XN(grdTOTAL.TextMatrix(i, 3)) & ","
        sql = sql & XN(grdTOTAL.TextMatrix(i, 4)) & ","
        sql = sql & XN(grdTOTAL.TextMatrix(i, 5)) & ")"

        DBConn.Execute sql
    Next
End Function

Private Sub cmdexportar_Click()
    lleno_tmp
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
    Rep.Destination = crptToWindow
    Rep.Formulas(0) = "FECHA='" & cboMes.Text & " - " & cboAnio.Text & "'"
        
    Rep.WindowState = crptNormal
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR
    
    Rep.WindowTitle = "Informe Secretaria de Energia"
    Rep.ReportFileName = DRIVE & DirReport & "informe_secenergia.rpt"
    
    
    Rep.Action = 1
    
End Sub

Private Sub cmdImprimir_Click()
    'programar la impresion del informe!!!!
    If MsgBox("¿Confirma Impresión del Informe?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    'PONE A LA IMPRESORA  COMO PREDETERMINADA
    Dim X As Printer
    Dim mDriver As String
    mDriver = Impresora
    For Each X In Printers
        If X.DeviceName = mDriver Then
            ' La define como predeterminada del sistema.
            Set Printer = X
            Exit For
        End If
    Next
'-----------------------------------
    Set_Impresora
    ImprimirInforme
    
End Sub
Public Sub ImprimirInforme()
    Dim Renglon As Double
    Dim canttxt As Integer
    Dim W As Integer
    Screen.MousePointer = vbHourglass
    lblestado.Caption = "Imprimiendo..."
            
    For W = 1 To 1 'SE IMPRIME POR DUPLICADO
        ' encabezado
        Imprimir 5, 0.5, True, "Marisa Centenaro y CIA"
        Imprimir 2, 2.3, True, "* GNC - Informe para Secretaría de Energía"
        Imprimir 13, 2.3, True, "Mes: " & cboMes.Text & " de " & cboAnio.Text
                '---- IMPRESION DEL INFORME ------------------
        Renglon = 4
        '---------Encabezado de grilla---------------
        Imprimir 2, Renglon - 1, True, "GNC"
        Imprimir 2.5, Renglon - 0.5, True, "Neto"
        Imprimir 7, Renglon - 0.5, True, "Precio de Venta"
        Imprimir 10.5, Renglon - 0.5, True, "Metros Vendidos"
        Imprimir 14, Renglon - 0.5, True, "Total en Pesos"
        '-------------------------------------------
        For i = 1 To GrdGNC.Rows - 1
          Imprimir 2.5, Renglon, False, GrdGNC.TextMatrix(i, 1)
          Imprimir 7, Renglon, False, GrdGNC.TextMatrix(i, 2)
          Imprimir 10.5, Renglon, False, GrdGNC.TextMatrix(i, 3)
          Imprimir 14, Renglon, False, GrdGNC.TextMatrix(i, 4)
          Renglon = Renglon + 0.5
        Next
        Imprimir 2.5, Renglon + 1, True, "PPP s/I = "
        Imprimir 4.5, Renglon + 1, False, txtGNCPPPsI.Text
        Imprimir 7.5, Renglon + 1, True, "PPP c/I = "
        Imprimir 9.5, Renglon + 0.35, False, txtGNCNum.Text
        Imprimir 9.5, Renglon + 0.75, False, "________"
        Imprimir 9.5, Renglon + 1.5, False, txtGNCDen.Text
        Imprimir 11.5, Renglon + 1, False, " = " & txtGNCPPPcI.Text
                
        Renglon = Renglon + 4
        
        '---------Encabezado de grilla---------------
        Imprimir 2, Renglon - 1, True, "NAFTA"
        Imprimir 3, Renglon - 0.5, True, "Precio de Venta"
        Imprimir 7, Renglon - 0.5, True, "Litros Vendidos"
        Imprimir 11, Renglon - 0.5, True, "Total en Pesos"
        '-------------------------------------------
        For i = 1 To grdNafta.Rows - 1
          
          Imprimir 3, Renglon, False, grdNafta.TextMatrix(i, 2)
          Imprimir 7, Renglon, False, grdNafta.TextMatrix(i, 3)
          Imprimir 11, Renglon, False, grdNafta.TextMatrix(i, 4)
          Renglon = Renglon + 0.5
        Next
        Imprimir 2.5, Renglon + 1, True, "PPP s/I = "
        Imprimir 4.5, Renglon + 0.35, False, txtNaPPPsINUM.Text
        Imprimir 4.5, Renglon + 0.75, False, "________"
        Imprimir 4.5, Renglon + 1.5, False, txtNaPPPsIDEN.Text
        Imprimir 6.5, Renglon + 1, True, " = " & txtNAFTAPPPsI.Text
        
        Imprimir 9.5, Renglon + 1, True, "PPP c/I = "
        Imprimir 11.5, Renglon + 0.35, False, txtNaPPPcInum.Text
        Imprimir 11.5, Renglon + 0.75, False, "________"
        Imprimir 11.5, Renglon + 1.5, False, txtNaPPPcIden.Text
        Imprimir 13.5, Renglon + 1, True, " = " & txtNaftaPPPcI.Text
        
        Renglon = Renglon + 4
        '---------Encabezado de grilla---------------
        Imprimir 2, Renglon - 1, True, "GASOIL"
        Imprimir 3, Renglon - 0.5, True, "Precio de Venta"
        Imprimir 7, Renglon - 0.5, True, "Litros Vendidos"
        Imprimir 11, Renglon - 0.5, True, "Total en Pesos"
        '--------------------------------------------
        For i = 1 To grdGasoil.Rows - 1
          
          Imprimir 3, Renglon, False, grdGasoil.TextMatrix(i, 2)
          Imprimir 7, Renglon, False, grdGasoil.TextMatrix(i, 3)
          Imprimir 11, Renglon, False, grdGasoil.TextMatrix(i, 4)
          Renglon = Renglon + 0.5
        Next
        Imprimir 2.5, Renglon + 1, True, "PPP s/I = "
        Imprimir 4.5, Renglon + 0.35, False, txtGaPPPsINUM.Text
        Imprimir 4.5, Renglon + 0.75, False, "________"
        Imprimir 4.5, Renglon + 1.5, False, txtGaPPPsIDEN.Text
        Imprimir 6.5, Renglon + 1, True, " = " & txtGASOILPPPsI.Text
        
        Imprimir 9.5, Renglon + 1, True, "PPP c/I = "
        Imprimir 11.5, Renglon + 0.35, False, txtGaPPPcInum.Text
        Imprimir 11.5, Renglon + 0.75, False, "________"
        Imprimir 11.5, Renglon + 1.5, False, txtGaPPPcIden.Text
        Imprimir 13.5, Renglon + 1, True, " = " & txtGasoilPPPcI.Text
        
        Renglon = Renglon + 4
        '---------Encabezado de grilla---------------
        Imprimir 6, Renglon - 0.5, True, "PPP s/I"
        Imprimir 9, Renglon - 0.5, True, "PPP c/I"
        Imprimir 12, Renglon - 0.5, True, "Volumen"
        Imprimir 14.5, Renglon - 0.5, True, "Precio Final"
        '--------------------------------------------
              
        For i = 1 To grdTOTAL.Rows - 1
          Imprimir 0.5, Renglon, False, grdTOTAL.TextMatrix(i, 0)
          Imprimir 3.5, Renglon, False, grdTOTAL.TextMatrix(i, 1)
          Imprimir 6, Renglon, False, grdTOTAL.TextMatrix(i, 2)
          Imprimir 9, Renglon, False, grdTOTAL.TextMatrix(i, 3)
          Imprimir 12, Renglon, False, grdTOTAL.TextMatrix(i, 4)
          Imprimir 14.5, Renglon, False, grdTOTAL.TextMatrix(i, 5)
          Renglon = Renglon + 0.5
        Next
        
        Printer.EndDoc
    Next W
    Screen.MousePointer = vbNormal
    lblestado.Caption = ""
End Sub

Private Sub cmdListar_Click()
    BuscarGNC
    BuscarNAFTA
    BuscarGASOIL
    BuscarTOTAL
    cmdtxt.Enabled = True
End Sub
Private Function BuscarGNC()
    ' calcular y mostrar los precios del GNC en el mes
    Dim Fecha As Date
    Dim Primer As Date
    Dim Ultimo As Date
    Dim vPrecioNeto As Double
    Dim vPrecioVta As Double
    Dim vMetros As Double
    Dim vTotPesos As Double
    Dim vAcuTotal As Double
    Dim VNETO As Double
    Dim vNetosum As Double
    Dim vNetoavg As Double
    Dim vMetTotal As Double
    Dim vAcuNetoTotal As Double
    
    Fecha = "10/" & cboMes.Text & "/" & cboAnio
     
    'Usamos la funcion DAteSerial para obtener el primero y el ultimo dia
    Primer = DateSerial(Year(Fecha), Month(Fecha) + 0, 1)
    Ultimo = DateSerial(Year(Fecha), Month(Fecha) + 1, 0)
      
      
    sql = "SELECT DF.DFC_PRECIO,DF.DFC_IMP,DF.DFC_TasaVial,SUM(DF.DFC_CANTIDAD) AS METROS"
    sql = sql & " FROM FACTURA_CLIENTE F, DETALLE_FACTURA_CLIENTE DF"
    sql = sql & " WHERE F.TCO_CODIGO = DF.TCO_CODIGO "
    sql = sql & " AND F.FCL_NUMERO = DF.FCL_NUMERO "
    sql = sql & " AND F.EST_CODIGO<>2"
    sql = sql & " AND F.FCL_SUCURSAL = DF.FCL_SUCURSAL "
    sql = sql & " AND (DF.PTO_CODIGO = 2 OR DF.PTO_CODIGO = 4)"
    sql = sql & " AND F.FCL_FECHA BETWEEN " & XDQ(Primer) & " AND " & XDQ(Ultimo)
    sql = sql & " GROUP BY DF.DFC_PRECIO,DF.DFC_IMP,DF.DFC_TasaVial"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    GrdGNC.Rows = 1
    If rec.EOF = False Then
        Do While rec.EOF = False
            VNETO = NetoGNC(rec!DFC_IMP, Format(rec!DFC_PRECIO, "0.00"), rec!DFC_TasaVial)
            GrdGNC.AddItem rec!DFC_IMP & Chr(9) & Format(VNETO, "0.00") & Chr(9) & _
                           Format(rec!DFC_PRECIO, "0.00") & Chr(9) & Format(rec!METROS, "0.00") _
                           & Chr(9) & Format(rec!DFC_PRECIO * rec!METROS, "0.00")
                        vMetTotal = vMetTotal + Format(rec!METROS, "0.00")
                        vAcuNetoTotal = vAcuNetoTotal + Format(VNETO * rec!METROS, "0.00")
                        vAcuTotal = vAcuTotal + Format(rec!DFC_PRECIO * rec!METROS, "0.00")
                        vNetosum = vNetosum + Format(VNETO, "0.00")
            
            rec.MoveNext
        Loop
        
        'vNetoavg = vNetosum / (GrdGNC.Rows - 1)
        GrdGNC.AddItem "" & Chr(9) & vNetosum & " : " & GrdGNC.Rows - 1 & " =  " & Format(vNetoavg, "0.00") & Chr(9) & _
                        "" & Chr(9) & vMetTotal & Chr(9) & vAcuTotal
        
        'Calculo el PPP S/I y el PPP C/I
        
        txtGNCPPPsI.Text = Format(vAcuNetoTotal / vMetTotal, "0.00")
        
        'txtGNCPPPsI.Text = Format(vNetoavg, "0.00")
        
        txtGNCNum.Text = Format(vAcuTotal, "0.00")
        txtGNCDen.Text = Format(vMetTotal, "0.00")
        txtGNCPPPcI.Text = Format(txtGNCNum.Text / txtGNCDen, "0.00")
        
        
    End If
    rec.Close
End Function

Private Function BuscarNAFTA()
    ' calcular y mostrar los precios del NAFTA en el mes
    Dim Fecha As Date
    Dim Primer As Date
    Dim Ultimo As Date
    Dim vPrecioNeto As Double
    Dim vPrecioVta As Double
    Dim vMetros As Double
    Dim vTotPesos As Double
    Dim vAcuTotal As Double
    Dim VNETO As Double
    Dim vPreciosum As Double
    Dim vPrecioavg As Double
    Dim vLitTotal As Double
    Dim vAcuImp As Double
    Dim vImpAvg As Double
    Dim vNetoNafta As Double
    Dim vAcuNetoTotal As Double
    
    Fecha = "10/" & cboMes.Text & "/" & cboAnio
     
    'Usamos la funcion DAteSerial para obtener el primero y el ultimo dia
    Primer = DateSerial(Year(Fecha), Month(Fecha) + 0, 1)
    Ultimo = DateSerial(Year(Fecha), Month(Fecha) + 1, 0)
      
      
    sql = "SELECT DF.DFC_PRECIO,DF.DFC_TasaVial,DF.DFC_IMP,SUM(DF.DFC_CANTIDAD) AS LITROS"
    sql = sql & " FROM FACTURA_CLIENTE F, DETALLE_FACTURA_CLIENTE DF, PRODUCTO P"
    sql = sql & " WHERE F.TCO_CODIGO = DF.TCO_CODIGO "
    sql = sql & " AND F.FCL_NUMERO = DF.FCL_NUMERO "
    sql = sql & " AND F.EST_CODIGO<>2"
    sql = sql & " AND F.FCL_SUCURSAL = DF.FCL_SUCURSAL "
    sql = sql & " AND DF.PTO_CODIGO = P.PTO_CODIGO "
    'sql = sql & " AND (DF.PTO_CODIGO = 1 or DF.PTO_CODIGO = 78 or DF.PTO_CODIGO = 84)" ' nafta codigo 1
    sql = sql & " AND P.RUB_CODIGO = 1 "
    sql = sql & " AND F.FCL_FECHA BETWEEN " & XDQ(Primer) & " AND " & XDQ(Ultimo)
    sql = sql & " GROUP BY DF.DFC_PRECIO,DF.DFC_IMP,DF.DFC_TasaVial"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    grdNafta.Rows = 1
    If rec.EOF = False Then
        Do While rec.EOF = False
            VNETO = NetoGNC(Chk0(rec!DFC_IMP), Format(Chk0(rec!DFC_PRECIO), "0.00"), Chk0(rec!DFC_TasaVial))
            grdNafta.AddItem rec!DFC_IMP & Chr(9) & Format(VNETO, "0.00") & Chr(9) & _
                           Format(rec!DFC_PRECIO, "0.00") & Chr(9) & Format(rec!LITROS, "0.00") _
                           & Chr(9) & Format(rec!DFC_PRECIO * rec!LITROS, "0.00")
                        
                        vLitTotal = vLitTotal + Format(Chk0(rec!LITROS), "0.00")
                        vAcuTotal = vAcuTotal + Format(Chk0(rec!DFC_PRECIO) * Chk0(rec!LITROS), "0.00")
                        vAcuNetoTotal = vAcuNetoTotal + Format(VNETO * Chk0(rec!LITROS), "0.00")
                        'vPreciosum = vPreciosum + Format(rec!DFC_PRECIO - rec!DFC_TasaVial, "0.00")
                        vPreciosum = vPreciosum + VNETO
                        ' esto lo hago por las dudas cambie el impuesto en un mes
                        vAcuImp = vAcuImp + Format(Chk0(rec!DFC_IMP), "0.00")
                        
            rec.MoveNext
        Loop
        
        vImpAvg = vAcuImp / (grdNafta.Rows - 1)
'        vPrecioavg = vPreciosum / (grdNafta.Rows - 1)
        vPrecioavg = Format(vAcuNetoTotal / vLitTotal, "0.00")
                
        grdNafta.AddItem "" & Chr(9) & "" & Chr(9) & _
                        "Total $ " & vPreciosum & " : " & grdNafta.Rows - 1 & " =  " & Format(vPrecioavg, "0.00") & Chr(9) & vLitTotal & Chr(9) & vAcuTotal
        
                 
        txtNaPPPsINUM.Text = Format(vPrecioavg, "0.00")
        
        'calcular los imp y el iva
        'vNetoNafta = NetoGNC(vImpAvg, vPrecioavg, 0)
        txtNaPPPsIDEN.Text = "1,000" 'Format(vPrecioavg - vNetoNafta, "0.00")
        txtNAFTAPPPsI.Text = Format(txtNaPPPsINUM.Text / txtNaPPPsIDEN, "0.00")
        
        
        
        txtNaPPPcInum.Text = Format(vAcuTotal, "0.00")
        txtNaPPPcIden.Text = Format(vLitTotal, "0.00")
        txtNaftaPPPcI.Text = Format(txtNaPPPcInum.Text / txtNaPPPcIden.Text, "0.00")
    
    End If
    rec.Close
End Function
Private Function BuscarGASOIL()
    ' calcular y mostrar los precios del GASOIL en el mes
    Dim Fecha As Date
    Dim Primer As Date
    Dim Ultimo As Date
    Dim vPrecioNeto As Double
    Dim vPrecioVta As Double
    Dim vMetros As Double
    Dim vTotPesos As Double
    Dim vAcuTotal As Double
    Dim VNETO As Double
    Dim vPreciosum As Double
    Dim vPrecioavg As Double
    Dim vLitTotal As Double
    Dim vAcuImp As Double
    Dim vImpAvg As Double
    Dim vNetoGasoil As Double
    Dim vAcuNetoTotal As Double
    
    Fecha = "10/" & cboMes.Text & "/" & cboAnio
     
    'Usamos la funcion DAteSerial para obtener el primero y el ultimo dia
    Primer = DateSerial(Year(Fecha), Month(Fecha) + 0, 1)
    Ultimo = DateSerial(Year(Fecha), Month(Fecha) + 1, 0)
      
      
    sql = "SELECT DF.DFC_PRECIO,DF.DFC_IMP,DF.DFC_TasaVial,SUM(DF.DFC_CANTIDAD) AS LITROS"
    sql = sql & " FROM FACTURA_CLIENTE F, DETALLE_FACTURA_CLIENTE DF"
    sql = sql & " WHERE F.TCO_CODIGO = DF.TCO_CODIGO "
    sql = sql & " AND F.FCL_NUMERO = DF.FCL_NUMERO "
    sql = sql & " AND F.EST_CODIGO<>2"
    sql = sql & " AND F.FCL_SUCURSAL = DF.FCL_SUCURSAL "
    sql = sql & " AND (DF.PTO_CODIGO = 3 OR DF.PTO_CODIGO = 81)" ' GASOIL codigo 3
    sql = sql & " AND F.FCL_FECHA BETWEEN " & XDQ(Primer) & " AND " & XDQ(Ultimo)
    sql = sql & " GROUP BY DF.DFC_PRECIO,DF.DFC_IMP,DF.DFC_TasaVial"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    grdGasoil.Rows = 1
    If rec.EOF = False Then
        Do While rec.EOF = False
            VNETO = NetoGNC(Chk0(rec!DFC_IMP), Format(rec!DFC_PRECIO, "0.00"), rec!DFC_TasaVial)
            grdGasoil.AddItem rec!DFC_IMP & Chr(9) & Format(VNETO, "0.00") & Chr(9) & _
                           Format(rec!DFC_PRECIO, "0.00") & Chr(9) & Format(rec!LITROS, "0.00") _
                           & Chr(9) & Format(rec!DFC_PRECIO * rec!LITROS, "0.00")
                        vLitTotal = vLitTotal + Format(rec!LITROS, "0.00")
                        vAcuTotal = vAcuTotal + Format(rec!DFC_PRECIO * rec!LITROS, "0.00")
                        vAcuNetoTotal = vAcuNetoTotal + Format(VNETO * rec!LITROS, "0.00")
                        
                        vPreciosum = vPreciosum + Format(rec!DFC_PRECIO)
                        ' esto lo hago por las dudas cambie el impuesto en un mes
                        vAcuImp = vAcuImp + Format(Chk0(rec!DFC_IMP), "0.00")
            rec.MoveNext
        Loop
        vImpAvg = vAcuImp / (grdGasoil.Rows - 1)
        vPrecioavg = Format(vAcuNetoTotal / vLitTotal, "0.00")
        
        grdGasoil.AddItem "" & Chr(9) & "" & Chr(9) & _
                        "Total $ " & vPreciosum & " : " & grdGasoil.Rows - 1 & " =  " & Format(vPrecioavg, "0.00") & Chr(9) & vLitTotal & Chr(9) & vAcuTotal
        
    
    
        txtGaPPPsINUM.Text = Format(vPrecioavg, "0.00")
        'calcular los imp y el iva
        'vNetoGasoil = NetoGASOIL(vImpAvg, vPrecioavg)
        txtGaPPPsIDEN.Text = "1,000" 'Format(vPrecioavg - vNetoGasoil, "0.00")
        txtGASOILPPPsI.Text = Format(txtGaPPPsINUM.Text / txtGaPPPsIDEN, "0.00")
        
        
        txtGaPPPcInum.Text = Format(vAcuTotal, "0.00")
        txtGaPPPcIden.Text = Format(vLitTotal, "0.00")
        txtGasoilPPPcI.Text = Format(txtGaPPPcInum.Text / txtGaPPPcIden.Text, "0.00")
    
    End If
    rec.Close
End Function
Private Function BuscarTOTAL()
    
    'CARGO TOTALES DE GNC
    If GrdGNC.Rows > 1 Or grdGasoil.Rows > 1 Or grdNafta.Rows > 1 Then
        grdTOTAL.Rows = 1
        grdTOTAL.AddItem "" & Chr(9) & "GNC" & Chr(9) & _
                        txtGNCPPPsI.Text & Chr(9) & txtGNCPPPcI.Text & Chr(9) & _
                        txtGNCDen.Text & Chr(9) & GrdGNC.TextMatrix(GrdGNC.Rows - 2, 2)
                         
        'CARGO TOTALES DE NAFTA
        
        grdTOTAL.AddItem "" & Chr(9) & "NAFTA" & Chr(9) & _
                        txtNAFTAPPPsI.Text & Chr(9) & txtNaftaPPPcI.Text & Chr(9) & _
                        txtNaPPPcIden.Text & Chr(9) & grdNafta.TextMatrix(grdNafta.Rows - 2, 2)
        
        
        'CARGO TOTALES DE GASOIL
        
        grdTOTAL.AddItem "" & Chr(9) & "GASOIL" & Chr(9) & _
                         txtGASOILPPPsI.Text & Chr(9) & txtGasoilPPPcI.Text & Chr(9) & _
                         txtGaPPPcIden.Text & Chr(9) & grdGasoil.TextMatrix(grdGasoil.Rows - 2, 2)
        
    End If
End Function

Private Function NetoGNC(pImpue As Double, pPrecio As Double, pTasaVial As Double) As Double
    
    Dim subtotal As Double
    Dim mIVA_1 As Double
    Dim mIVA_2 As Double
    Dim IMPUESTO As Double
    
    subtotal = 0
    
    mIVA_1 = BuscoIva
    'mIVA_2 = BuscoIva_2
    
    If pImpue <> 0 Then
        IMPUESTO = pImpue * 1 ' PASO 1
        subtotal = 1 * pPrecio - IMPUESTO - pTasaVial
        subtotal = subtotal / (1 + (mIVA_1 / 100))
    End If
    NetoGNC = subtotal 'Neto sin impuestos
End Function
Private Function NetoGASOIL(pImpue As Double, pPrecio As Double) As Double
    
    Dim subtotal As Double
    Dim mIVA_1 As Double
    Dim mIVA_2 As Double
    Dim IMPUESTO As Double
    
    subtotal = 0
    
    'mIVA_1 = BuscoIva
    mIVA_2 = BuscoIva_2
    
    If pImpue <> 0 Then
        IMPUESTO = pImpue * 1 ' PASO 1
        subtotal = 1 * pPrecio - IMPUESTO
        subtotal = subtotal / (1 + (mIVA_2 / 100))
    End If
    NetoGASOIL = subtotal
End Function
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
Private Sub ObtenerPrimerUltimoDia(Fecha As Date)



     'MsgBox " Primer día : " & Primer & vbNewLine & _
            " Último día : " & Ultimo, vbInformation

End Sub

Private Sub CmdNuevo_Click()
    GrdGNC.Rows = 4
    For i = 1 To GrdGNC.Rows - 1
        For J = 0 To GrdGNC.Cols - 1
            GrdGNC.TextMatrix(i, J) = ""
        Next J
    Next i
    txtGNCPPPsI.Text = ""
    txtGNCNum.Text = ""
    txtGNCDen.Text = ""
    txtGNCPPPcI.Text = ""
    
    grdNafta.Rows = 4
    For i = 1 To grdNafta.Rows - 1
        For J = 0 To grdNafta.Cols - 1
            grdNafta.TextMatrix(i, J) = ""
        Next J
    Next i
    txtNaPPPsINUM.Text = ""
    txtNaPPPsIDEN.Text = ""
    txtNAFTAPPPsI.Text = ""
    txtNaPPPcInum.Text = ""
    txtNaPPPcIden.Text = ""
    txtNaftaPPPcI.Text = ""
    
    grdGasoil.Rows = 4
    For i = 1 To grdGasoil.Rows - 1
        For J = 0 To grdGasoil.Cols - 1
            grdGasoil.TextMatrix(i, J) = ""
        Next J
    Next i
    txtGaPPPsINUM.Text = ""
    txtGaPPPsIDEN.Text = ""
    txtGASOILPPPsI.Text = ""
    txtGaPPPcInum.Text = ""
    txtGaPPPcIden.Text = ""
    txtGasoilPPPcI.Text = ""
    
    grdTOTAL.Rows = 4
    For i = 1 To grdTOTAL.Rows - 1
        For J = 0 To grdTOTAL.Cols - 1
            grdTOTAL.TextMatrix(i, J) = ""
        Next J
    Next i
    cmdtxt.Enabled = False
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdtxt_Click()
    CrearArchvioAFIP
End Sub
Private Function CrearArchvioAFIP()
    Dim cadena(2) As String
    Dim TipoAgente As String
    Dim TipoCombN As String
    Dim CantLitrosN As String
    Dim PrecioVentaN As String
    Dim TipoCombG As String
    Dim CantLitrosG As String
    Dim PrecioVentaG As String
    Dim MesAnio As String
    
    MesAnio = cboMes.Text & "_" & cboAnio.Text
    '************* NAFTA SUPER *****************
    'Tipo de Agente
    TipoAgente = "A6"
    'Tipo de Combustible
    TipoCombN = 2 'NAFTA SUPER
    'Cant Litros
    CantLitrosN = Round(txtNaPPPcIden.Text, 0)
    CantLitrosN = String(10 - Len(CantLitrosN), "0") & CantLitrosN
    'PrecioVenta
    PrecioVentaN = txtNAFTAPPPsI.Text
    'PrecioVentaN = Replace(PrecioVentaN, ".", "")
    PrecioVentaN = String(18 - Len(PrecioVentaN), "0") & PrecioVentaN
    '*******************************************
    '************* GASOIL *****************
    'Tipo de Combustible
    TipoCombG = 5 'GASOIL GRADO 2
    'Cant Litros
    CantLitrosG = Round(txtGaPPPcIden.Text, 0)
    CantLitrosG = String(10 - Len(CantLitrosG), "0") & CantLitrosG
    'PrecioVenta
    PrecioVentaG = txtGASOILPPPsI.Text
    'PrecioVentaG = Replace(PrecioVentaG, ".", "")
    PrecioVentaG = String(18 - Len(PrecioVentaG), "0") & PrecioVentaG
    '*******************************************
    
    
    'ARMO UNA LINEA DEL ARCHIVO COMPROBANTES
    cadena(0) = TipoAgente & ";;;;;;;;;;;;;;;;;" & TipoCombN & ";" & CantLitrosN & ";" & PrecioVentaN & ";"
    cadena(1) = TipoAgente & ";;;;;;;;;;;;;;;;;" & TipoCombG & ";" & CantLitrosG & ";" & PrecioVentaG & ";"
        
    If EstadoDeArchivo(DirAFIP & "Agentes_de_Combustibles_Liquidos_" & MesAnio & ".txt") Then
        Kill (DirAFIP & "Agentes_de_Combustibles_Liquidos_" & MesAnio & ".txt")
    End If
        
    'GENERO LOS ARCHIVOS
    For i = 0 To 1
        Open DirAFIP & "Agentes_de_Combustibles_Liquidos_" & MesAnio & ".txt" For Append As #1
        Print #1, cadena(i)
        Close #1
    Next
    
    
    MsgBox "Se genero correctamente el archivo " & DirAFIP & "Agentes_de_Combustibles_Liquidos_" & MesAnio & ".txt", vbInformation, TIT_MSGBOX
   
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then MySendKeys Chr(9)
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Call Centrar_pantalla(Me)
    'Me.Left = 0
    Me.Top = 0
    SeteoInicial
    CboMes_Año

End Sub
Private Function CboMes_Año()
    For i = 1 To 12
        cboMes.AddItem MonthName(i)
        cboMes.ItemData(cboMes.NewIndex) = i
    Next i
    cboMes.ListIndex = Month(Date) - 1
    For i = 1980 To 2099
        cboAnio.AddItem i
    Next i
    cboAnio.Text = Year(Date)
End Function
Private Sub SeteoInicial()
    'CONFIGURO GRILLA GNC
    lblestado.Caption = ""
    
    GrdGNC.FormatString = "^Código Interno|Neto|Precio de Venta|Metros Vendidos|Total en Pesos"
    GrdGNC.ColWidth(0) = 0    'CODIGO INTERNO PRODUCTO
    GrdGNC.ColWidth(1) = 1875 'NETO
    GrdGNC.ColWidth(2) = 1875 'PRECIO DE VENTA
    GrdGNC.ColWidth(3) = 1875 'METROS VENDIDOS
    GrdGNC.ColWidth(4) = 1875 'TOTAL EN PESOS
    GrdGNC.Rows = 4
    GrdGNC.Cols = 5
    GrdGNC.BorderStyle = flexBorderNone
    GrdGNC.row = 0
    For i = 0 To 4
        GrdGNC.Col = i
        GrdGNC.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        GrdGNC.CellBackColor = &H808080    'GRIS OSCURO
        GrdGNC.CellFontBold = True
    Next
    GrdGNC.HighLight = flexHighlightWithFocus
    
    'CONFIGURO GRILLA NAFTA
    grdNafta.FormatString = "^Código Interno|Neto|Precio de Venta|Litros Vendidos|Total en Pesos"
    grdNafta.ColWidth(0) = 0    'CODIGO INTERNO PRODUCTO
    grdNafta.ColWidth(1) = 0 'NETO
    grdNafta.ColWidth(2) = 2450 'PRECIO DE VENTA
    grdNafta.ColWidth(3) = 2450 'METROS VENDIDOS
    grdNafta.ColWidth(4) = 2450 'TOTAL EN PESOS
    grdNafta.Rows = 4
    grdNafta.Cols = 5
    grdNafta.BorderStyle = flexBorderNone
    grdNafta.row = 0
    For i = 0 To 4
        grdNafta.Col = i
        grdNafta.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        grdNafta.CellBackColor = &H808080    'GRIS OSCURO
        grdNafta.CellFontBold = True
    Next
    grdNafta.HighLight = flexHighlightWithFocus
    
    
    'CONFIGURO GRILLA GASOIL
    grdGasoil.FormatString = "^Código Interno|Neto|Precio de Venta|Litros Vendidos|Total en Pesos"
    grdGasoil.ColWidth(0) = 0    'CODIGO INTERNO PRODUCTO
    grdGasoil.ColWidth(1) = 0 'NETO
    grdGasoil.ColWidth(2) = 2450 'PRECIO DE VENTA
    grdGasoil.ColWidth(3) = 2450 'METROS VENDIDOS
    grdGasoil.ColWidth(4) = 2450 'TOTAL EN PESOS
    grdGasoil.Rows = 4
    grdGasoil.Cols = 5
    grdGasoil.BorderStyle = flexBorderNone
    grdGasoil.row = 0
    For i = 0 To 4
        grdGasoil.Col = i
        grdGasoil.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        grdGasoil.CellBackColor = &H808080    'GRIS OSCURO
        grdGasoil.CellFontBold = True
    Next
    grdGasoil.HighLight = flexHighlightWithFocus
    
    
    'CONFIGURO GRILLA TOTAL
    grdTOTAL.FormatString = "^Código Interno| |PPP s/I|PPP c/I|Volumen|Precio Final"
    grdTOTAL.ColWidth(0) = 0    'CODIGO INTERNO PRODUCTO
    grdTOTAL.ColWidth(1) = 1450 'PPP s/I
    grdTOTAL.ColWidth(2) = 1550 'PPP s/I
    grdTOTAL.ColWidth(3) = 1550 'PPP c/I
    grdTOTAL.ColWidth(4) = 1550 'Volumen
    grdTOTAL.ColWidth(5) = 1550 'Precio Final
    grdTOTAL.Rows = 4
    grdTOTAL.BorderStyle = flexBorderNone
    grdTOTAL.row = 0
    For i = 0 To 5
        grdTOTAL.Col = i
        grdTOTAL.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        grdTOTAL.CellBackColor = &H808080    'GRIS OSCURO
        grdTOTAL.CellFontBold = True
    Next
    grdTOTAL.HighLight = flexHighlightWithFocus
End Sub
Public Function EstadoDeArchivo(ByVal Archivo As String) As Boolean
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If (fso.FileExists(Archivo)) Then
        EstadoDeArchivo = True
    Else
        EstadoDeArchivo = False
    End If
End Function

