VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmListadoCantVendidasVendedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Cantidades Vendidas por Playero"
   ClientHeight    =   2520
   ClientLeft      =   1515
   ClientTop       =   1740
   ClientWidth     =   5145
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmListadoCantVendidasVendedor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2520
   ScaleWidth      =   5145
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   2520
      TabIndex        =   6
      Top             =   2085
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   3825
      TabIndex        =   7
      Top             =   2085
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      Caption         =   "Listar por..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   45
      TabIndex        =   25
      Top             =   30
      Width           =   5055
      Begin VB.ComboBox cboDesde 
         BackColor       =   &H8000000E&
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
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1140
         Width           =   1450
      End
      Begin VB.ComboBox cbohasta 
         BackColor       =   &H8000000E&
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
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1140
         Width           =   1450
      End
      Begin VB.ComboBox cboVendedor 
         Height          =   315
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   405
         Width           =   3470
      End
      Begin MSComCtl2.DTPicker FechaDesde 
         Height          =   315
         Left            =   1350
         TabIndex        =   1
         Top             =   772
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   56688641
         CurrentDate     =   41098
      End
      Begin MSComCtl2.DTPicker FechaHasta 
         Height          =   315
         Left            =   3360
         TabIndex        =   2
         Top             =   772
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   56688641
         CurrentDate     =   41098
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Index           =   1
         Left            =   2880
         TabIndex        =   30
         Top             =   1200
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Hora        Desde:"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   1200
         Width           =   1215
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
         Left            =   705
         TabIndex        =   28
         Top             =   450
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Index           =   0
         Left            =   2880
         TabIndex        =   27
         Top             =   780
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha      Desde:"
         Height          =   195
         Left            =   105
         TabIndex        =   26
         Top             =   825
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2250
      Left            =   6735
      TabIndex        =   15
      Top             =   210
      Visible         =   0   'False
      Width           =   6915
      Begin VB.TextBox txtEmpresaCuit 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   900
         TabIndex        =   23
         Top             =   660
         Width           =   2235
      End
      Begin VB.TextBox txtEmp_Id 
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
         Left            =   4365
         TabIndex        =   22
         Top             =   1065
         Visible         =   0   'False
         Width           =   675
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
         Height          =   75
         Left            =   90
         TabIndex        =   21
         Top             =   1560
         Width           =   6795
      End
      Begin VB.TextBox txtTipoLibro 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4365
         TabIndex        =   20
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtEmpresa 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   900
         TabIndex        =   19
         Top             =   375
         Width           =   3075
      End
      Begin VB.TextBox txtMes_LibroI 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   900
         MaxLength       =   2
         TabIndex        =   8
         Top             =   945
         Width           =   795
      End
      Begin VB.TextBox txtAnio_LibroI 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1725
         MaxLength       =   4
         TabIndex        =   9
         Top             =   945
         Width           =   1380
      End
      Begin VB.TextBox txtLibro_IdI 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4380
         TabIndex        =   16
         Top             =   360
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label Label6 
         Caption         =   "C.U.I.T."
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
         Left            =   120
         TabIndex        =   24
         Top             =   660
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Empresa"
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
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Periodo"
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
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   1005
         Width           =   540
      End
   End
   Begin VB.Frame fraImpresion 
      Caption         =   "Destino"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   45
      TabIndex        =   10
      Top             =   1680
      Width           =   2175
      Begin VB.PictureBox picSalida 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   135
         Picture         =   "frmListadoCantVendidasVendedor.frx":0442
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   14
         Top             =   315
         Width           =   240
      End
      Begin VB.PictureBox picSalida 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   135
         Picture         =   "frmListadoCantVendidasVendedor.frx":0544
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   13
         Top             =   315
         Width           =   240
      End
      Begin VB.PictureBox picSalida 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   135
         Picture         =   "frmListadoCantVendidasVendedor.frx":0646
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   12
         Top             =   315
         Width           =   240
      End
      Begin VB.ComboBox cboDestino 
         Height          =   315
         ItemData        =   "frmListadoCantVendidasVendedor.frx":0748
         Left            =   450
         List            =   "frmListadoCantVendidasVendedor.frx":0755
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   270
         Width           =   1635
      End
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   2340
      Top             =   1725
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowControls  =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Modo 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6690
      TabIndex        =   11
      Top             =   2985
      Visible         =   0   'False
      Width           =   315
   End
End
Attribute VB_Name = "frmListadoCantVendidasVendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboDestino_Click()
    picSalida(0).Visible = False
    picSalida(1).Visible = False
    picSalida(2).Visible = False
    picSalida(cboDestino.ListIndex).Visible = True
End Sub

Private Sub cboListar_Click()
    If frmListadoCantVendidasVendedor.Visible = True Then
        If cboListar.ListIndex = 0 Then
            cboAgrupar.Enabled = True
            cboAgrupar.ListIndex = 0
        Else
            cboAgrupar.Enabled = False
            cboAgrupar.ListIndex = -1
        End If
    End If
End Sub

Private Sub cmdAceptar_Click()
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
    
    Select Case cboDestino.ListIndex
        Case 0
            Rep.Destination = crptToWindow
        Case 1
            Rep.Destination = crptToPrinter
        Case 2
            Rep.Destination = crptToFile
    End Select
    
    'SOLO FACTURAS DEFINITIVAS
    Rep.SelectionFormula = " {FACTURA_CLIENTE.EST_CODIGO}=3"
    
    If cboVendedor.List(cboVendedor.ListIndex) <> "(Todos)" Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = " {FACTURA_CLIENTE.VEN_CODIGO}=" & XN(cboVendedor.ItemData(cboVendedor.ListIndex))
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {FACTURA_CLIENTE.VEN_CODIGO}=" & XN(cboVendedor.ItemData(cboVendedor.ListIndex))
        End If
    End If
    If FechaDesde.Value <> "" Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = " {FACTURA_CLIENTE.FCL_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {FACTURA_CLIENTE.FCL_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
        End If
    End If
    If FechaHasta.Value <> "" Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = " {FACTURA_CLIENTE.FCL_FECHA}<= DATE( " & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {FACTURA_CLIENTE.FCL_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
        End If
    End If
    
'    If cboDesde.ListIndex <> -1 Then
'        If Rep.SelectionFormula = "" Then
'            Rep.SelectionFormula = " {FACTURA_CLIENTE.FCL_HORA}>=" & Hour(cboDesde.Text)
'        Else
'            Rep.SelectionFormula = Rep.SelectionFormula & " AND {FACTURA_CLIENTE.FCL_HORA}>=" & Hour(cboDesde.Text)
'        End If
'    End If
'    If cbohasta.ListIndex <> -1 Then
'        If Rep.SelectionFormula = "" Then
'            Rep.SelectionFormula = " {FACTURA_CLIENTE.FCL_HORA}<=" & Hour(cbohasta.Text)
'        Else
'            Rep.SelectionFormula = Rep.SelectionFormula & " AND {FACTURA_CLIENTE.FCL_HORA}<=" & Hour(cbohasta.Text)
'        End If
'    End If
    
    
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
    
    Rep.WindowTitle = "Listado de Cantidades Vendidas"
    Rep.ReportFileName = DRIVE & DirReport & "cantidades_vendidas_vendedor.rpt"
    Rep.Action = 1
End Sub

Private Sub CmdCancelar_Click()
    Set frmListadoCantVendidasVendedor = Nothing
    'mQuienLlamo = "ABMProducto"
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub LlenarComboHoras()
    Dim cItems As Integer
    'rec.Open "SELECT HS_DESDE,HS_HASTA FROM PARAMETROS", DBConn, adOpenStatic, adLockOptimistic
    'If rec.EOF = False Then
        hDesde = Hour("00:00")
        hHasta = Hour("23:59")
    'End If
    'rec.Close
    cItems = (hHasta - hDesde) * 2 + 1
    i = 0
    For J = hDesde To hHasta
        cboDesde.AddItem Format(J, "00") & ":00"
        cboDesde.ItemData(cboDesde.NewIndex) = i
        cboDesde.AddItem Format(J, "00") & ":30"
        cboDesde.ItemData(cboDesde.NewIndex) = i + 0.5
        
        cbohasta.AddItem Format(J, "00") & ":00"
        cbohasta.ItemData(cbohasta.NewIndex) = i
        cbohasta.AddItem Format(J, "00") & ":30"
        cbohasta.ItemData(cbohasta.NewIndex) = i + 0.5
        
        i = i + 1
    Next
    cboDesde.ListIndex = -1
    cbohasta.ListIndex = -1
    
End Sub

Private Sub Form_Load()
    Dim rec As ADODB.Recordset
    Set rec = New ADODB.Recordset
    Me.Top = 0
    Me.Left = 0
    cboVendedor.AddItem "(Todos)"
    CargoComboBox cboVendedor, "VENDEDOR", "VEN_CODIGO", "VEN_NOMBRE", "VEN_NOMBRE"
    If cboVendedor.ListCount > 0 Then cboVendedor.ListIndex = 0
    'CARGO COMBO LINEA
    cboDestino.ListIndex = 0
    FechaDesde.Value = Date
    FechaHasta.Value = Date
    LlenarComboHoras
End Sub

