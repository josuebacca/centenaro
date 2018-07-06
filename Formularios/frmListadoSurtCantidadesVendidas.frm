VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmListadoSurtCantidadesVendidas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Cantidades Vendidas por Surtidor"
   ClientHeight    =   2670
   ClientLeft      =   1515
   ClientTop       =   1740
   ClientWidth     =   5625
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
   Icon            =   "frmListadoSurtCantidadesVendidas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2670
   ScaleWidth      =   5625
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   3000
      TabIndex        =   5
      Top             =   2205
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   4305
      TabIndex        =   6
      Top             =   2205
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
      Height          =   1815
      Left            =   45
      TabIndex        =   24
      Top             =   30
      Width           =   5535
      Begin VB.ComboBox cboTurno 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   840
         Width           =   1515
      End
      Begin VB.OptionButton optCantVend 
         Caption         =   "Cantidades Vendidas"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.OptionButton optSurtidor 
         Caption         =   "Combustibles Vendidos por Surtidor"
         Height          =   375
         Left            =   2280
         TabIndex        =   4
         Top             =   1320
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker FechaDesde 
         Height          =   315
         Left            =   1200
         TabIndex        =   0
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   52035585
         CurrentDate     =   41098
      End
      Begin MSComCtl2.DTPicker FechaHasta 
         Height          =   315
         Left            =   3840
         TabIndex        =   1
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   52035585
         CurrentDate     =   41098
      End
      Begin VB.Label Label70 
         AutoSize        =   -1  'True
         Caption         =   "Turno:"
         Height          =   195
         Left            =   720
         TabIndex        =   28
         Top             =   840
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Index           =   0
         Left            =   2895
         TabIndex        =   26
         Top             =   435
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   255
         TabIndex        =   25
         Top             =   450
         Width           =   990
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
      Left            =   8880
      TabIndex        =   14
      Top             =   120
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   7
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
         TabIndex        =   8
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
         TabIndex        =   15
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
         TabIndex        =   23
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
         TabIndex        =   17
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
         TabIndex        =   16
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
      TabIndex        =   9
      Top             =   1920
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
         Picture         =   "frmListadoSurtCantidadesVendidas.frx":0442
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
         Index           =   1
         Left            =   135
         Picture         =   "frmListadoSurtCantidadesVendidas.frx":0544
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   12
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
         Picture         =   "frmListadoSurtCantidadesVendidas.frx":0646
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   11
         Top             =   315
         Width           =   240
      End
      Begin VB.ComboBox cboDestino 
         Height          =   315
         ItemData        =   "frmListadoSurtCantidadesVendidas.frx":0748
         Left            =   450
         List            =   "frmListadoSurtCantidadesVendidas.frx":0755
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   270
         Width           =   1635
      End
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   2340
      Top             =   2085
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
      TabIndex        =   10
      Top             =   2985
      Visible         =   0   'False
      Width           =   315
   End
End
Attribute VB_Name = "frmListadoSurtCantidadesVendidas"
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
    If frmListadoSurtCantidadesVendidas.Visible = True Then
        If cboListar.ListIndex = 0 Then
            cboAgrupar.Enabled = True
            cboAgrupar.ListIndex = 0
        Else
            cboAgrupar.Enabled = False
            cboAgrupar.ListIndex = -1
        End If
    End If
End Sub

Private Sub chkSurtidor_Click()
End Sub

Private Function cantVend()
    DBConn.Execute "DELETE FROM TMP_CANTVEND"
    
    sql = "SELECT * FROM TURNOS WHERE TUR_CODIGO = 3" 'NOCHE
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic


    sql = "SELECT DF.PTO_CODIGO,P.PTO_DESCRI,SUM(DF.DFC_CANTIDAD)AS CANTI,SUM(DF.DFC_CANTIDAD*DF.DFC_PRECIO) AS MONTO"
    sql = sql & " FROM FACTURA_CLIENTE FC, DETALLE_FACTURA_CLIENTE DF, PRODUCTO P"
    sql = sql & " WHERE FC.TCO_CODIGO = DF.TCO_CODIGO"
    sql = sql & " AND FC.FCL_SUCURSAL = DF.FCL_SUCURSAL"
    sql = sql & " AND FC.FCL_NUMERO = DF.FCL_NUMERO"
    sql = sql & " AND FC.EST_CODIGO = 3"
    sql = sql & " AND DF.PTO_CODIGO = P.PTO_CODIGO"
    If cboTurno.ItemData(cboTurno.ListIndex) = 3 Then
        If FechaDesde.Value = FechaHasta.Value Then
            sql = sql & " AND ((FC.FCL_FECHA= " & XDQ(FechaDesde.Value) & ""
            sql = sql & " AND FC.FCL_HORA  >=#" & Rec1!TUR_DESDE & "#)" 'vDesde & "#)" '
            sql = sql & " OR (FC.FCL_FECHA = " & XDQ(DateValue(FechaHasta.Value) + 1) & ""
            sql = sql & " AND FC.FCL_HORA  <=#" & Rec1!TUR_HASTA & "#))" 'vHasta & "#))" '
        End If
        sql = sql & " AND FC.TUR_CODIGO = 3 "
    Else
        If FechaDesde.Value <> "" Then sql = sql & " AND FC.FCL_FECHA >=" & XDQ(FechaDesde.Value)
        If FechaHasta.Value <> "" Then sql = sql & " AND FC.FCL_FECHA <=" & XDQ(FechaHasta.Value)
        If cboTurno.ListIndex > 0 Then sql = sql & " AND FC.TUR_CODIGO =" & cboTurno.ItemData(cboTurno.ListIndex)
    End If
    sql = sql & " GROUP BY DF.PTO_CODIGO,P.PTO_DESCRI"
    Rec1.Close
    
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
            sql = "INSERT INTO TMP_CANTVEND (PTO_CODIGO,PTO_DESCRI,PTO_CANTIDAD,PTO_IMPORTE)"
            sql = sql & " VALUES ("
            sql = sql & XN(Rec1!PTO_CODIGO) & ","
            sql = sql & XS(Rec1!PTO_DESCRI) & ","
            sql = sql & XN(Rec1!CANTI) & ","
            sql = sql & XN(Rec1!MONTO) & ")"
            DBConn.Execute sql
            Rec1.MoveNext
        Loop
    
    End If
    Rec1.Close

End Function

Private Sub cmdAceptar_Click()
    If FechaDesde.Value = "" Then
        MsgBox "Falta Ingresar la Fecha Desde", vbExclamation, TIT_MSGBOX
        FechaDesde.SetFocus
        Exit Sub
    End If
    If FechaHasta.Value = "" Then
        MsgBox "Falta Ingresar la Fecha Hasta", vbExclamation, TIT_MSGBOX
        FechaHasta.SetFocus
        Exit Sub
    End If
    
    
    cantVend 'FUNCION QUE LLENA LA TABLA TEMPORAL
    
    
    
    
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
    
    If optSurtidor.Value = True Then 'VEO SI ESTA SELECCIONADO EL REPORTE POR SURTIDOR
        If cboTurno.ListIndex <> 0 Then
        
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " {T_STOCK.T_TURNO}='" & cboTurno.Text & "'"
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {T_STOCK.T_TURNO}='" & cboTurno.Text & "'"
            End If
        End If
        
        If FechaDesde.Value <> "" Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " {T_STOCK.T_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {T_STOCK.T_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
            End If
        End If
        If FechaHasta.Value <> "" Then
            If FechaDesde.Value = FechaHasta.Value Then ' MISMO DIA
                If cboTurno.ItemData(cboTurno.ListIndex) = 3 Then ' SI ES TURNO NOCHE // ESTO LO HAGO PARA BUSCAR LAS DEL TURNO NOCHE DEL DIA SIGUIENTE
                    If Rep.SelectionFormula = "" Then
                        Rep.SelectionFormula = " {T_STOCK.T_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) + 1 & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
                    Else
                        Rep.SelectionFormula = Rep.SelectionFormula & " AND {T_STOCK.T_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) + 1 & ")"
                        
                    End If
                Else
                    If Rep.SelectionFormula = "" Then
                        Rep.SelectionFormula = " {T_STOCK.T_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
                    Else
                        Rep.SelectionFormula = Rep.SelectionFormula & " AND {T_STOCK.T_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
                    End If
                End If
            Else
                If Rep.SelectionFormula = "" Then
                    Rep.SelectionFormula = " {T_STOCK.T_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
                Else
                    Rep.SelectionFormula = Rep.SelectionFormula & " AND {T_STOCK.T_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
                End If
            End If
        End If
    Else
        cantVend
        'SOLO FACTURAS DEFINITIVAS
        
'        Rep.SelectionFormula = " {FACTURA_CLIENTE.EST_CODIGO}=3"
'
'        If cboTurno.ListIndex <> 0 Then
'            If Rep.SelectionFormula = "" Then
'                Rep.SelectionFormula = " {FACTURA_CLIENTE.TUR_CODIGO}=" & cboTurno.ItemData(cboTurno.ListIndex) & ""
'            Else
'                Rep.SelectionFormula = Rep.SelectionFormula & " AND {FACTURA_CLIENTE.TUR_CODIGO}=" & cboTurno.ItemData(cboTurno.ListIndex) & ""
'            End If
'        End If
'
'        If FechaDesde.value <> "" Then
'            If Rep.SelectionFormula = "" Then
'                Rep.SelectionFormula = " {FACTURA_CLIENTE.FCL_FECHA}>= DATE (" & Mid(FechaDesde.value, 7, 4) & "," & Mid(FechaDesde.value, 4, 2) & "," & Mid(FechaDesde.value, 1, 2) & ")"
'            Else
'                Rep.SelectionFormula = Rep.SelectionFormula & " AND {FACTURA_CLIENTE.FCL_FECHA}>= DATE (" & Mid(FechaDesde.value, 7, 4) & "," & Mid(FechaDesde.value, 4, 2) & "," & Mid(FechaDesde.value, 1, 2) & ")"
'            End If
'        End If
'        If FechaHasta.value <> "" Then
'            If FechaDesde.value = FechaHasta.value Then ' MISMO DIA
'                If cboTurno.ItemData(cboTurno.ListIndex) = 3 Then ' SI ES TURNO NOCHE // ESTO LO HAGO PARA BUSCAR LAS DEL TURNO NOCHE DEL DIA SIGUIENTE
'                    If Rep.SelectionFormula = "" Then
'                        Rep.SelectionFormula = " {FACTURA_CLIENTE.FCL_FECHA}<= DATE (" & Mid(FechaHasta.value, 7, 4) + 1 & "," & Mid(FechaHasta.value, 4, 2) & "," & Mid(FechaHasta.value, 1, 2) & ")"
'                    Else
'                        Rep.SelectionFormula = Rep.SelectionFormula & " AND {FACTURA_CLIENTE.FCL_FECHA}<= DATE (" & Mid(FechaHasta.value, 7, 4) & "," & Mid(FechaHasta.value, 4, 2) & "," & Mid(FechaHasta.value, 1, 2) + 1 & ")"
'                        'Rep.SelectionFormula = Rep.SelectionFormula & " AND {FACTURA_CLIENTE.FCL_HORA}<= TIME(22:15:00)"
'                    End If
'                Else
'                    If Rep.SelectionFormula = "" Then
'                        Rep.SelectionFormula = " {FACTURA_CLIENTE.FCL_FECHA}<= DATE (" & Mid(FechaHasta.value, 7, 4) & "," & Mid(FechaHasta.value, 4, 2) & "," & Mid(FechaHasta.value, 1, 2) & ")"
'                    Else
'                        Rep.SelectionFormula = Rep.SelectionFormula & " AND {FACTURA_CLIENTE.FCL_FECHA}<= DATE (" & Mid(FechaHasta.value, 7, 4) & "," & Mid(FechaHasta.value, 4, 2) & "," & Mid(FechaHasta.value, 1, 2) & ")"
'                    End If
'                End If
'            Else
'                If Rep.SelectionFormula = "" Then
'                    Rep.SelectionFormula = " {FACTURA_CLIENTE.FCL_FECHA}<= DATE (" & Mid(FechaHasta.value, 7, 4) & "," & Mid(FechaHasta.value, 4, 2) & "," & Mid(FechaHasta.value, 1, 2) & ")"
'                Else
'                    Rep.SelectionFormula = Rep.SelectionFormula & " AND {FACTURA_CLIENTE.FCL_FECHA}<= DATE (" & Mid(FechaHasta.value, 7, 4) & "," & Mid(FechaHasta.value, 4, 2) & "," & Mid(FechaHasta.value, 1, 2) & ")"
'                End If
'            End If
'        End If
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
    If cboTurno.ListIndex <> 0 Then
        If FechaDesde.Value = FechaHasta.Value Then ' MISMO DIA
            If cboTurno.ItemData(cboTurno.ListIndex) = 3 Then
                Rep.Formulas(1) = "Turno='" & "Turno: " & cboTurno.Text & " - Comienza el " & FechaDesde.Value & " y termina el " & DateValue(FechaHasta.Value) + 1 & "'"
            Else
                Rep.Formulas(1) = "Turno='" & "Turno: " & cboTurno.Text & "'"
            End If
        Else
            Rep.Formulas(1) = "Turno='" & "Turno: " & cboTurno.Text & "'"
        End If
        
    End If
    
    Rep.WindowState = crptNormal
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR
    If optSurtidor.Value = True Then 'VEO SI ESTA SELECCIONADO EL REPORTE POR SURTIDOR
        Rep.WindowTitle = "Listado de Combustibles Vendidos por Surtidor"
        Rep.ReportFileName = DRIVE & DirReport & "cant_vendidas_surtidor.rpt"
    Else
        Rep.WindowTitle = "Listado de Cantidades Vendidas"
        Rep.ReportFileName = DRIVE & DirReport & "cantidades_vendidas.rpt"
    End If
    
    Rep.Action = 1
End Sub

Private Sub CmdCancelar_Click()
    Set frmListadoSurtCantidadesVendidas = Nothing
    'mQuienLlamo = "ABMProducto"
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    Dim rec As ADODB.Recordset
    Set rec = New ADODB.Recordset
    Me.Top = 0
    Me.Left = 0
    'CARGO COMBO LINEA
    cboDestino.ListIndex = 0
    FechaDesde.Value = Date
    FechaHasta.Value = Date
    LlenarComboTurnos
    
End Sub

Private Sub LlenarComboTurnos()
    sql = "SELECT * FROM TURNOS"
    sql = sql & " ORDER BY TUR_CODIGO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboTurno.AddItem "<TODOS>"
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

