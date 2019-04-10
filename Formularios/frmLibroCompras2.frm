VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLibroCompras2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Libro IVA Compras"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport rep 
      Index           =   0
      Left            =   120
      Top             =   5280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame optObserva 
      Caption         =   "Tipo de Listado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      TabIndex        =   30
      Top             =   1560
      Width           =   5715
      Begin VB.CheckBox chkTodos 
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   32
         Top             =   0
         Width           =   975
      End
      Begin VB.CheckBox chkreporte 
         Caption         =   "NO - Libro Observaciones"
         Height          =   255
         Index           =   10
         Left            =   2880
         TabIndex        =   12
         Top             =   1920
         Width           =   2175
      End
      Begin VB.CheckBox chkreporte 
         Caption         =   "Informe Favor por Concepto"
         Height          =   255
         Index           =   9
         Left            =   2880
         TabIndex        =   11
         Top             =   1440
         Width           =   2535
      End
      Begin VB.CheckBox chkreporte 
         Caption         =   "Informe Proveedores por Concepto"
         Height          =   255
         Index           =   8
         Left            =   2880
         TabIndex        =   10
         Top             =   1080
         Width           =   2775
      End
      Begin VB.CheckBox chkreporte 
         Caption         =   "NO - Informe por Concepto"
         Height          =   255
         Index           =   7
         Left            =   2880
         TabIndex        =   9
         Top             =   720
         Width           =   2655
      End
      Begin VB.CheckBox chkreporte 
         Caption         =   "NO - Libro IVA"
         Height          =   255
         Index           =   6
         Left            =   2880
         TabIndex        =   8
         Top             =   360
         Width           =   2175
      End
      Begin VB.CheckBox chkreporte 
         Caption         =   "Libro Observaciones"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   7
         Top             =   1920
         Width           =   2175
      End
      Begin VB.CheckBox chkreporte 
         Caption         =   "Libro Combustibles"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   6
         Top             =   1608
         Width           =   2415
      End
      Begin VB.CheckBox chkreporte 
         Caption         =   "Informe Favor por Proveedor"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   5
         Top             =   1296
         Width           =   2895
      End
      Begin VB.CheckBox chkreporte 
         Caption         =   "Informe por Proveedor"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   984
         Width           =   1935
      End
      Begin VB.CheckBox chkreporte 
         Caption         =   "Informe por Concepto"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   672
         Width           =   2415
      End
      Begin VB.CheckBox chkreporte 
         Caption         =   "Libro IVA"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   4770
      Picture         =   "frmLibroCompras2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4905
      Width           =   840
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   5715
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   225
         Left            =   75
         TabIndex        =   23
         Top             =   1185
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   397
         _Version        =   327682
         Appearance      =   1
      End
      Begin MSComCtl2.DTPicker FechaDesde 
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   54198273
         CurrentDate     =   41098
      End
      Begin MSComCtl2.DTPicker FechaHasta 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   54198273
         CurrentDate     =   41098
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   180
         TabIndex        =   28
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   225
         TabIndex        =   27
         Top             =   660
         Width           =   960
      End
      Begin VB.Label lblPeriodo1 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2850
         TabIndex        =   26
         Top             =   255
         Width           =   1785
      End
      Begin VB.Label lblPeriodo2 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2850
         TabIndex        =   25
         Top             =   630
         Width           =   1785
      End
      Begin VB.Label lblPor 
         AutoSize        =   -1  'True
         Caption         =   "100 %"
         Height          =   195
         Left            =   4950
         TabIndex        =   24
         Top             =   1215
         Width           =   435
      End
   End
   Begin VB.Frame FrameImpresora 
      Caption         =   "impresora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   19
      Top             =   3990
      Width           =   5730
      Begin VB.OptionButton optImpresora 
         Caption         =   "Impresora"
         Height          =   195
         Left            =   2370
         TabIndex        =   15
         Top             =   315
         Width           =   1050
      End
      Begin VB.OptionButton optPantalla 
         Caption         =   "Pantalla"
         Height          =   195
         Left            =   945
         TabIndex        =   13
         Top             =   315
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.CommandButton CBImpresora 
         Caption         =   "Configurar Impresora"
         Height          =   375
         Left            =   3810
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   1665
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   150
         TabIndex        =   21
         Top             =   315
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmLibroCompras2.frx":030A
      Height          =   735
      Left            =   3915
      Picture         =   "frmLibroCompras2.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4905
      Width           =   840
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   825
      Top             =   5055
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtHoja 
      Height          =   315
      Left            =   2280
      TabIndex        =   14
      ToolTipText     =   "Ingrese el ultimo numero de hoja impreso"
      Top             =   5340
      Width           =   495
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   735
      Left            =   3060
      Picture         =   "frmLibroCompras2.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4905
      Width           =   840
   End
   Begin Crystal.CrystalReport rep 
      Index           =   1
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport rep 
      Index           =   2
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport rep 
      Index           =   3
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport rep 
      Index           =   4
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport rep 
      Index           =   5
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport rep 
      Index           =   6
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport rep 
      Index           =   7
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport rep 
      Index           =   8
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport rep 
      Index           =   9
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport rep 
      Index           =   10
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Nro. Hoja:"
      Height          =   195
      Left            =   1560
      TabIndex        =   31
      Top             =   5400
      Width           =   720
   End
   Begin VB.Label lblEstado 
      AutoSize        =   -1  'True
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   105
      TabIndex        =   29
      Top             =   4920
      Width           =   750
   End
End
Attribute VB_Name = "frmLibroCompras2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Registro As Long
Dim Tamanio As Long
Dim TotIva As Double

Private Sub CBImpresora_Click()
    CDImpresora.PrinterDefault = True
    CDImpresora.ShowPrinter
    FrameImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
End Sub
Private Sub ArmarListado(Numero)
     Registro = 0
     Tamanio = 0
     TotIva = 0
     Dim Tabla As String
     If IsNull(FechaDesde.Value) Then
        MsgBox "Debe ingresar el periodo", vbExclamation, TIT_MSGBOX
        FechaDesde.SetFocus
        Exit Sub
     End If
     If IsNull(FechaHasta.Value) Then
        MsgBox "Debe ingresar el periodo", vbExclamation, TIT_MSGBOX
        FechaHasta.SetFocus
        Exit Sub
     End If
     
     On Error GoTo CLAVO
     Screen.MousePointer = vbHourglass
     DBConn.BeginTrans
     lblestado.Caption = "Buscando Datos..."
     
    If Numero = 0 Then
        'BORRO LA TABLA TMP_LIBRO_IVA_COMPRAS
        Tabla = "TMP_LIBRO_IVA_COMPRAS"
    End If
    If Numero = 4 Then
        'BORRO LA TABLA TMP_LIBRO_IVA_COMPRAS COMBUSTIBLES
        Tabla = "TMP_LIBRO_IVA_COMPRAS_COMB"
    End If
    If Numero = 6 Then
        'BORRO LA TABLA TMP_LIBRO_IVA_COMPRAS
        Tabla = "TMP_LIBRO_NO_IVA_COMPRAS"
    End If
    
    sql = "DELETE FROM " & Tabla
    DBConn.Execute sql
    
    'BUSCO COMPROBANTES DENTRO DE LOS GASTOS GENERALES -----
    sql = "SELECT GG.GGR_NROSUCTXT,GG.GGR_NROCOMPTXT,GG.GGR_FECHACOMP,GG.GGR_IVA,GG.GGR_IVA1,GG.GGR_NETO,GG.GGR_TOTAL,"
    sql = sql & " GG.GGR_IVA1,GG.GGR_NETO1,GG.GGR_IMPUESTOS,GGR_PERIIBB,GGR_PERIVA,GGR_PERGAN,"
    sql = sql & " P.PROV_CUIT,P.PROV_INGBRU,GGR_IMP1IVA,GGR_IMP2IVA,"
    sql = sql & " P.PROV_RAZSOC,TC.TCO_ABREVIA,GG.GGR_NAFTA,GG.GGR_GASOIL,TG.TGT_DESCRI"
    sql = sql & " FROM GASTOS_GENERALES GG,"
    sql = sql & " TIPO_COMPROBANTE TC, PROVEEDOR P, TIPO_GASTO TG"
    sql = sql & " WHERE GG.TCO_CODIGO=TC.TCO_CODIGO"
    sql = sql & " AND GG.TPR_CODIGO=P.TPR_CODIGO"
    sql = sql & " AND GG.PROV_CODIGO=P.PROV_CODIGO"
    sql = sql & " AND GG.TGT_CODIGO=TG.TGT_CODIGO"
    'sql = sql & " AND GG.GGR_FAVOR=0"
    'sql = sql & " AND GG.TGT_CODIGO<>13" 'RETENCION
    If Numero = 0 Then
        sql = sql & " AND GG.GGR_LIBROIVA = " & XS("S")
        If FechaDesde <> "" Then sql = sql & " AND GG.GGR_PERIODO>=" & XDQ(FechaDesde)
        If FechaHasta <> "" Then sql = sql & " AND GG.GGR_PERIODO<=" & XDQ(FechaHasta)
    Else
        If Numero = 6 Then
            sql = sql & " AND GG.GGR_FAVOR<>1" '2-no IVA (no incluye ningun favor)
            sql = sql & " AND GG.GGR_LIBROIVA = " & XS("N")
            If FechaDesde <> "" Then sql = sql & " AND (GG.GGR_FECHACOMP >=" & XDQ(FechaDesde) & " OR GG.GGR_PERIODO >=" & XDQ(FechaDesde) & " )"
            If FechaHasta <> "" Then sql = sql & " AND (GG.GGR_FECHACOMP<=" & XDQ(FechaHasta) & " OR GG.GGR_PERIODO <=" & XDQ(FechaHasta) & " )"
        End If
    End If
    If Numero = 4 Then
        sql = sql & " AND GG.TGT_CODIGO=1" ' COMBUSTIBLES
        If FechaDesde <> "" Then sql = sql & " AND GG.GGR_PERIODO>=" & XDQ(FechaDesde)
        If FechaHasta <> "" Then sql = sql & " AND GG.GGR_PERIODO<=" & XDQ(FechaHasta)
    End If

    sql = sql & " ORDER BY GG.GGR_FECHACOMP"

    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Registro = 0
        Tamanio = rec.RecordCount
        Do While rec.EOF = False
            sql = "INSERT INTO " & Tabla & " (FECHA,COMPROBANTE,NUMERO,"
            sql = sql & "PROVEEDOR,CUIT,INGBRUTOS,SUBTOTAL,IVA,TOTIVA,"
            sql = sql & "SUBTOTAL1,TOTOTROIVA,IMPUESTOS,TOTAL,RETIIBB,RETIVA,RETGAN,IVA1,NRO_HOJA,NAFTA,GASOIL,CONCEPTO)"
            sql = sql & "VALUES ("
            sql = sql & XDQ(rec!GGR_FECHACOMP) & ","
            sql = sql & XS(rec!TCO_ABREVIA) & ","
            sql = sql & XS(rec!GGR_NROSUCTXT & "-" & rec!GGR_NROCOMPTXT) & ","
            sql = sql & XS(rec!PROV_RAZSOC) & ","
            sql = sql & XS(Format(rec!PROV_CUIT, "##-########-#")) & ","
            sql = sql & "NULL" & ","
            'sql = sql & XS(Format(rec!CLI_INGBRU, "###-#####-##")) & ","
            sql = sql & XN(rec!GGR_NETO) & ","
            sql = sql & XN(rec!GGR_IVA) & ","
                'TotIva = (CDbl(rec!GGR_NETO) * CDbl(rec!GGR_IVA)) / 100
            sql = sql & XN(Chk0(rec!GGR_IMP1IVA)) & ","
            sql = sql & XN(Chk0(rec!GGR_NETO1)) & ","  'OTRO NETO
                'TotIva = (CDbl(Chk0(rec!GGR_NETO1)) * CDbl(Chk0(rec!GGR_IVA1))) / 100

            sql = sql & XN(Chk0(rec!GGR_IMP2IVA)) & ","  'OTRO IVA
            If rec!GGR_IVA = 0 Then
                'PONGO EL TOTAL CUANDO EL GASTO CON IVA = 0
                ' EJ MONOTRIBUTISTA
                sql = sql & XN(Chk0(rec!GGR_TOTAL)) & ","  'IMPUESTOS
            Else
                sql = sql & XN(Chk0(rec!GGR_IMPUESTOS)) & ","  'IMPUESTOS
            End If
            sql = sql & XN(Chk0(rec!GGR_TOTAL)) & ","
            sql = sql & XN(Chk0(rec!GGR_PERIIBB)) & ","
            sql = sql & XN(Chk0(rec!GGR_PERIVA)) & ","
            sql = sql & XN(Chk0(rec!GGR_PERGAN)) & ","
            sql = sql & XN(Chk0(rec!GGR_IVA1)) & ","
            sql = sql & XN(Chk0(txtHoja.Text)) & ","
            sql = sql & XN(Chk0(rec!GGR_NAFTA)) & ","
            sql = sql & XN(Chk0(rec!GGR_GASOIL)) & ","
            sql = sql & XS(ChkNull(rec!TGT_DESCRI)) & ")"
            DBConn.Execute sql
            rec.MoveNext

            Registro = Registro + 1
            ProgressBar1.Value = Format((Registro * 100) / Tamanio, "0.0")
            lblPor.Caption = Format((Registro * 100) / Tamanio, "0.0") & " %"
        Loop
    End If
    rec.Close

    lblestado.Caption = ""
    DBConn.CommitTrans
    Screen.MousePointer = vbNormal
    Exit Sub

CLAVO:
 Screen.MousePointer = vbNormal
 lblestado.Caption = ""
 DBConn.RollbackTrans
 If rec.State = 1 Then rec.Close
 MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub chkLibro_Click()

End Sub

Private Sub chkResumen_Click()

End Sub

Private Function generar_reporte(Numero As Integer)
    Select Case Numero
        Case 0 'Libro IVA
            ArmarListado Numero
            ListarLibroIVA Numero
        Case 1 'Informe por Concepto
            InformeConcepto Numero
            ListarResumenIVA Numero
        Case 2 'Informe por Proveedor
            InformeProveedor
            ListarProveedor Numero
        Case 3 'Informe Favor por Proveedor
            InformeFavor
            ListarProveedor Numero
        Case 4 'Libro Combustibles
            ArmarListado Numero
            ListarLibroIVA_Combustibles Numero
        Case 5 'Libro Observaciones
            ArmarListObs Numero
            ListarObs Numero
        Case 6 'NO - Libro IVA
            ArmarListado Numero
            ListarLibroIVA Numero
        Case 7 'NO - Informe por Concepto
            InformeConcepto Numero
            ListarResumenIVA Numero
        Case 8 'Informe Proveedores por Concepto
            InformeFavorProvConcepto Numero
            ListarConcepto Numero
        Case 9 'Informe Favor por Concepto
            InformeFavorProvConcepto Numero
            ListarConcepto Numero
        Case 10 'NO - Libro Observaciones
            ArmarListObs Numero
            ListarObs Numero
    End Select

End Function

Private Sub chkTodos_Click()
    Dim i As Integer
    If chkTodos.Value = Checked Then
        For i = 0 To 10
            chkreporte(i).Value = Checked
        Next
    Else
        For i = 0 To 10
            chkreporte(i).Value = Unchecked
        Next
    End If
End Sub

Private Sub cmdAceptar_Click()
    Dim i As Integer
    For i = 0 To 10
        If chkreporte(i).Value = 1 Then
            'llamar a funcion generar_reporte
            generar_reporte i
        End If
    Next
    
End Sub
Private Function ListarObs(Numero As Integer)
    lblestado.Caption = "Buscando Listado..."
    Rep(Numero).WindowState = crptMaximized
    Rep(Numero).WindowBorderStyle = crptNoBorder
    Rep(Numero).Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=CENTENARO"
    Rep(Numero).Formulas(0) = ""
    Rep(Numero).Formulas(1) = ""
    Rep(Numero).Formulas(2) = ""
    Rep(Numero).Formulas(3) = ""
        
    sql = "SELECT CUIT,ING_BRUTOS,RAZ_SOCIAL,DIRECCION,TELEFONO FROM PARAMETROS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Rep(Numero).Formulas(0) = "EMPRESA='" & Trim(rec!RAZ_SOCIAL) & "'"
        Rep(Numero).Formulas(1) = "CUIT='" & Trim(rec!DIRECCION) & " - " & Trim(rec!TELEFONO) & " - " & Format(rec!cuit, "##-########-#") & "'"
        Rep(Numero).Formulas(2) = "INGBRUTOS='Ing. Brutos:  " & Format(rec!ING_BRUTOS, "###-#####-##") & "'"
    End If
    rec.Close
    
    If FechaDesde.Value <> "" And Not IsNull(FechaHasta.Value) Then
        Rep(Numero).Formulas(3) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & FechaHasta.Value & "'"
    ElseIf FechaDesde.Value <> "" And IsNull(FechaHasta.Value) Then
        Rep(Numero).Formulas(3) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & Date & "'"
    ElseIf IsNull(FechaDesde.Value) And Not IsNull(FechaHasta.Value) Then
        Rep(Numero).Formulas(3) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & FechaHasta.Value & "'"
    ElseIf IsNull(FechaDesde.Value) And IsNull(FechaHasta.Value) Then
        Rep(Numero).Formulas(3) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & Date & "'"
    End If
    
    If Numero = 5 Then
        Rep(Numero).WindowTitle = "Libro Compras con Observaciones - IVA"
        Rep(Numero).ReportFileName = DRIVE & DirReport & "rptlibroivacompras_obs.rpt"
    Else
        If Numero = 10 Then
            Rep(Numero).WindowTitle = "Libro Compras con Observaciones - NO IVA"
            Rep(Numero).ReportFileName = DRIVE & DirReport & "rptlibro_NO_ivacompras_obs.rpt"
        End If
    End If
       
    
    
    If optPantalla.Value = True Then
        Rep(Numero).Destination = crptToWindow
    ElseIf optImpresora.Value = True Then
        Rep(Numero).Destination = crptToPrinter
    End If
        
     Rep(Numero).Action = 1
     
     
     lblestado.Caption = ""
     Rep(Numero).Formulas(0) = ""
     Rep(Numero).Formulas(1) = ""
     Rep(Numero).Formulas(3) = ""
     
End Function
Private Sub ArmarListObs(Numero As Integer)
     Registro = 0
     Tamanio = 0
     TotIva = 0
     Dim Tabla As String
     If IsNull(FechaDesde.Value) Then
        MsgBox "Debe ingresar el periodo", vbExclamation, TIT_MSGBOX
        FechaDesde.SetFocus
        Exit Sub
     End If
     If IsNull(FechaHasta.Value) Then
        MsgBox "Debe ingresar el periodo", vbExclamation, TIT_MSGBOX
        FechaHasta.SetFocus
        Exit Sub
     End If
     
     On Error GoTo CLAVO
     Screen.MousePointer = vbHourglass
     DBConn.BeginTrans
     lblestado.Caption = "Buscando Datos..."
     
     If Numero = 5 Then
        Tabla = "TMP_LIBRO_IVA_COMPRAS_OBS"
     End If
     If Numero = 10 Then
        Tabla = "TMP_LIBRO_NO_IVA_COMPRAS_OBS"
     End If
    'BORRO LA TABLA TMP_LIBRO_IVA_COMPRAS
    sql = "DELETE FROM " & Tabla
    DBConn.Execute sql
    
    'BUSCO COMPROBANTES DENTRO DE LOS GASTOS GENERALES -----
    sql = "SELECT GG.GGR_NROSUCTXT,GG.GGR_NROCOMPTXT,GG.GGR_FECHACOMP,GG.GGR_TOTAL,"
    sql = sql & " P.PROV_CUIT,P.PROV_RAZSOC,P.PROV_FANTASIA,TC.TCO_ABREVIA,GG.GGR_OBSER"
    sql = sql & " FROM GASTOS_GENERALES GG,"
    sql = sql & " TIPO_COMPROBANTE TC, PROVEEDOR P"
    sql = sql & " WHERE GG.TCO_CODIGO=TC.TCO_CODIGO"
    sql = sql & " AND GG.TPR_CODIGO=P.TPR_CODIGO"
    sql = sql & " AND GG.PROV_CODIGO=P.PROV_CODIGO"
    sql = sql & " AND GG.GGR_FAVOR=0"
    'sql = sql & " AND GG.TGT_CODIGO<>13" 'RETENCION
    If Numero = 5 Then
        sql = sql & " AND GG.GGR_LIBROIVA = " & XS("S")
        If FechaDesde <> "" Then sql = sql & " AND GG.GGR_PERIODO>=" & XDQ(FechaDesde)
        If FechaHasta <> "" Then sql = sql & " AND GG.GGR_PERIODO<=" & XDQ(FechaHasta)
    Else
    'Libro de Observaciones NO IVA - numero = 10
        sql = sql & " AND GG.GGR_LIBROIVA = " & XS("N")
        If FechaDesde <> "" Then sql = sql & " AND (GG.GGR_FECHACOMP >=" & XDQ(FechaDesde) & " OR GG.GGR_PERIODO >=" & XDQ(FechaDesde) & " )"
        If FechaHasta <> "" Then sql = sql & " AND (GG.GGR_FECHACOMP<=" & XDQ(FechaHasta) & " OR GG.GGR_PERIODO <=" & XDQ(FechaHasta) & " )"
    End If

    sql = sql & " ORDER BY GG.GGR_FECHACOMP"

    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Registro = 0
        Tamanio = rec.RecordCount
        Do While rec.EOF = False
            sql = "INSERT INTO " & Tabla & " (FECHA,COMPROBANTE,NUMERO,"
            sql = sql & "PROVEEDOR,FANTASIA,CUIT,TOTAL,OBSERVA)"
            sql = sql & "VALUES ("
            sql = sql & XDQ(rec!GGR_FECHACOMP) & ","
            sql = sql & XS(rec!TCO_ABREVIA) & ","
            sql = sql & XS(rec!GGR_NROSUCTXT & "-" & rec!GGR_NROCOMPTXT) & ","
            sql = sql & XS(rec!PROV_RAZSOC) & ","
            sql = sql & XS(ChkNull(rec!PROV_FANTASIA)) & ","
            sql = sql & XS(Format(rec!PROV_CUIT, "##-########-#")) & ","
            sql = sql & XN(Chk0(rec!GGR_TOTAL)) & ","
            sql = sql & XS(ChkNull(rec!GGR_OBSER)) & ")"
            DBConn.Execute sql
            rec.MoveNext

            Registro = Registro + 1
            ProgressBar1.Value = Format((Registro * 100) / Tamanio, "0.0")
            lblPor.Caption = Format((Registro * 100) / Tamanio, "0.0") & " %"
        Loop
    End If
    rec.Close
    
lblestado.Caption = ""
DBConn.CommitTrans
Screen.MousePointer = vbNormal
Exit Sub

CLAVO:
 Screen.MousePointer = vbNormal
 lblestado.Caption = ""
 DBConn.RollbackTrans
 If rec.State = 1 Then rec.Close
 MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub ListarConcepto(Numero As Integer)
    lblestado.Caption = "Buscando Listado..."
    Rep(Numero).WindowState = crptMaximized
    Rep(Numero).WindowBorderStyle = crptNoBorder
    Rep(Numero).Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=CENTENARO"
    Rep(Numero).Formulas(0) = ""
    Rep(Numero).Formulas(1) = ""
    Rep(Numero).Formulas(2) = ""
    'rep(Numero).Formulas(3) = ""
        
    sql = "SELECT CUIT,ING_BRUTOS,RAZ_SOCIAL,DIRECCION,TELEFONO FROM PARAMETROS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Rep(Numero).Formulas(0) = "EMPRESA='" & Trim(rec!RAZ_SOCIAL) & "'"
        Rep(Numero).Formulas(1) = "CUIT='" & Trim(rec!DIRECCION) & " - " & Trim(rec!TELEFONO) & " - " & Format(rec!cuit, "##-########-#") & "'"
        Rep(Numero).Formulas(2) = "INGBRUTOS='Ing. Brutos:  " & Format(rec!ING_BRUTOS, "###-#####-##") & "'"
    End If
    rec.Close
    
    If FechaDesde.Value <> "" And Not IsNull(FechaHasta.Value) Then
        Rep(Numero).Formulas(3) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & FechaHasta.Value & "'"
    ElseIf FechaDesde.Value <> "" And IsNull(FechaHasta.Value) Then
        Rep(Numero).Formulas(3) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & Date & "'"
    ElseIf IsNull(FechaDesde.Value) And Not IsNull(FechaHasta.Value) Then
        Rep(Numero).Formulas(3) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & FechaHasta.Value & "'"
    ElseIf IsNull(FechaDesde.Value) And IsNull(FechaHasta.Value) Then
        Rep(Numero).Formulas(3) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & Date & "'"
    End If
    
    If Numero = 8 Then
        Rep(Numero).WindowTitle = "Informe de Proveedores por Concepto"
        Rep(Numero).ReportFileName = DRIVE & DirReport & "rptcompras_conceptos_prov.rpt"
    End If
    If Numero = 9 Then
        Rep(Numero).WindowTitle = "Informe a favor por Concepto"
        Rep(Numero).ReportFileName = DRIVE & DirReport & "rptcompras_FAVOR_conceptos_prov.rpt"
    End If
     
    
    If optPantalla.Value = True Then
        Rep(Numero).Destination = crptToWindow
    ElseIf optImpresora.Value = True Then
        Rep(Numero).Destination = crptToPrinter
    End If
     Rep(Numero).Action = 1
     
     lblestado.Caption = ""
     Rep(Numero).Formulas(0) = ""
     Rep(Numero).Formulas(1) = ""
     Rep(Numero).Formulas(2) = ""
     'rep(Numero).Formulas(3) = ""
End Sub
Private Sub InformeFavorProvConcepto(Numero As Integer)
    Dim Tabla As String
On Error GoTo CLAVO
     Screen.MousePointer = vbHourglass
     DBConn.BeginTrans
     lblestado.Caption = "Buscando Datos..."

        If Numero = 8 Then
            Tabla = "TMP_LIBRO_IVA_COMPRAS_CONCEPTO"
        End If
        If Numero = 9 Then
            Tabla = "TMP_LIBRO_IVA_FAVOR_COMPRAS_CONCEPTO"
        End If
        'BORRO LA TABLA TMP_LIBRO_IVA_COMPRAS
        sql = "DELETE FROM " & Tabla
        DBConn.Execute sql
        
        'BUSCO COMPROBANTES DENTRO DE LOS GASTOS GENERALES -----
        sql = "SELECT TG.TGT_DESCRI,GG.GGR_IVA,GG.GGR_IVA1,SUM(GG.GGR_NETO) AS NETO,SUM(GG.GGR_TOTAL) AS TOTAL,"
        sql = sql & " SUM(GG.GGR_NETO1) AS NETO1,SUM(GG.GGR_IMPUESTOS) AS IMPUESTO,SUM(GGR_PERIIBB) AS PERIIBB,SUM(GGR_PERIVA) AS PERIVA,SUM(GGR_PERGAN) AS PERGAN"
        sql = sql & " FROM GASTOS_GENERALES GG, TIPO_GASTO TG"
        sql = sql & " WHERE "
        sql = sql & " TG.TGT_CODIGO=GG.TGT_CODIGO"
        If Numero = 8 Then
            sql = sql & " AND GG.GGR_FAVOR=0" ' NORMAL (NO A FAVOR)
        Else
            If Numero = 9 Then
                sql = sql & " AND GG.GGR_FAVOR=1" ' A FAVOR
            End If
        End If
        sql = sql & " AND GG.GGR_LIBROIVA = " & XS("S")
        If FechaDesde <> "" Then sql = sql & " AND GG.GGR_PERIODO>=" & XDQ(FechaDesde)
        If FechaHasta <> "" Then sql = sql & " AND GG.GGR_PERIODO<=" & XDQ(FechaHasta)
        
        sql = sql & " GROUP BY TG.TGT_DESCRI,GG.GGR_IVA,GG.GGR_IVA1"
        sql = sql & " ORDER BY TG.TGT_DESCRI"

        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Registro = 0
            Tamanio = rec.RecordCount
            Do While rec.EOF = False
                sql = "INSERT INTO " & Tabla & " (CONCEPTO,"
                sql = sql & "SUBTOTAL,IVA,TOTIVA,"
                sql = sql & "SUBTOTAL1,TOTOTROIVA,IMPUESTOS,TOTAL,RETIIBB,RETIVA,RETGAN,IVA1)"
                sql = sql & "VALUES ("
                sql = sql & XS(rec!TGT_DESCRI) & ","
                sql = sql & XN(rec!NETO) & ","
                sql = sql & XN(rec!GGR_IVA) & ","
                    TotIva = (CDbl(rec!NETO) * CDbl(rec!GGR_IVA)) / 100
                sql = sql & XN(CStr(TotIva)) & ","
                sql = sql & XN(Chk0(rec!NETO1)) & "," 'OTRO NETO
                    TotIva = (CDbl(Chk0(rec!NETO1)) * CDbl(Chk0(rec!GGR_IVA1))) / 100
                
                sql = sql & XN(CStr(TotIva)) & "," 'OTRO IVA
                If rec!GGR_IVA = 0 Then
                    'PONGO EL TOTAL CUANDO EL GASTO CON IVA = 0
                    ' EJ MONOTRIBUTISTA
                    sql = sql & XN(Chk0(rec!TOTAL)) & "," 'IMPUESTOS
                Else
                    sql = sql & XN(Chk0(rec!IMPUESTO)) & "," 'IMPUESTOS
                End If
                sql = sql & XN(Chk0(rec!TOTAL)) & ","
                sql = sql & XN(Chk0(rec!PERIIBB)) & ","
                sql = sql & XN(Chk0(rec!PERIVA)) & ","
                sql = sql & XN(Chk0(rec!PERGAN)) & ","
                sql = sql & XN(Chk0(rec!GGR_IVA1)) & ")"
                DBConn.Execute sql
                rec.MoveNext

                Registro = Registro + 1
                ProgressBar1.Value = Format((Registro * 100) / Tamanio, "0.0")
                lblPor.Caption = Format((Registro * 100) / Tamanio, "0.0") & " %"
            Loop
        End If
        rec.Close
        
        
        'BUSCO COMPROBANTES DENTRO DE LOS GASTOS GENERALES  -----
        sql = "SELECT TG.TGT_DESCRI,GG.GGR_IVA,GG.GGR_IVA1,SUM(GG.GGR_NETO) AS NETO,SUM(GG.GGR_TOTAL) AS TOTAL,"
        sql = sql & " SUM(GG.GGR_NETO1) AS NETO1,SUM(GG.GGR_IMPUESTOS) AS IMPUESTO,SUM(GGR_PERIIBB) AS PERIIBB,SUM(GGR_PERIVA) AS PERIVA,SUM(GGR_PERGAN) AS PERGAN"
        sql = sql & " FROM GASTOS_GENERALES GG, TIPO_GASTO TG"
        sql = sql & " WHERE "
        sql = sql & " TG.TGT_CODIGO=GG.TGT_CODIGO"
        If Numero = 8 Then
            sql = sql & " AND GG.GGR_FAVOR=0" ' NORMAL (NO A FAVOR)
        Else
            If Numero = 9 Then
                sql = sql & " AND GG.GGR_FAVOR=1" ' NORMAL (NO A FAVOR)
            End If
        End If
        sql = sql & " AND GG.GGR_LIBROIVA = " & XS("N")
        If FechaDesde <> "" Then sql = sql & " AND (GG.GGR_FECHACOMP >=" & XDQ(FechaDesde) & " OR GG.GGR_PERIODO >=" & XDQ(FechaDesde) & " )"
        If FechaHasta <> "" Then sql = sql & " AND (GG.GGR_FECHACOMP<=" & XDQ(FechaHasta) & " OR GG.GGR_PERIODO <=" & XDQ(FechaHasta) & " )"
        
        sql = sql & " GROUP BY TG.TGT_DESCRI,GG.GGR_IVA,GG.GGR_IVA1"
        sql = sql & " ORDER BY TG.TGT_DESCRI"

        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Registro = 0
            Tamanio = rec.RecordCount
            Do While rec.EOF = False
                sql = "INSERT INTO " & Tabla & " (CONCEPTO,"
                sql = sql & "SUBTOTAL,IVA,TOTIVA,"
                sql = sql & "SUBTOTAL1,TOTOTROIVA,IMPUESTOS,TOTAL,RETIIBB,RETIVA,RETGAN,IVA1)"
                sql = sql & "VALUES ("
                sql = sql & XS(rec!TGT_DESCRI) & ","
                sql = sql & XN(rec!NETO) & ","
                sql = sql & XN(rec!GGR_IVA) & ","
                    TotIva = (CDbl(rec!NETO) * CDbl(rec!GGR_IVA)) / 100
                sql = sql & XN(CStr(TotIva)) & ","
                sql = sql & XN(Chk0(rec!NETO1)) & "," 'OTRO NETO
                    TotIva = (CDbl(Chk0(rec!NETO1)) * CDbl(Chk0(rec!GGR_IVA1))) / 100
                
                sql = sql & XN(CStr(TotIva)) & "," 'OTRO IVA
                If rec!GGR_IVA = 0 Then
                    'PONGO EL TOTAL CUANDO EL GASTO CON IVA = 0
                    ' EJ MONOTRIBUTISTA
                    sql = sql & XN(Chk0(rec!TOTAL)) & "," 'IMPUESTOS
                Else
                    sql = sql & XN(Chk0(rec!IMPUESTO)) & "," 'IMPUESTOS
                End If
                sql = sql & XN(Chk0(rec!TOTAL)) & ","
                sql = sql & XN(Chk0(rec!PERIIBB)) & ","
                sql = sql & XN(Chk0(rec!PERIVA)) & ","
                sql = sql & XN(Chk0(rec!PERGAN)) & ","
                sql = sql & XN(Chk0(rec!GGR_IVA1)) & ")"
                DBConn.Execute sql
                rec.MoveNext

                Registro = Registro + 1
                ProgressBar1.Value = Format((Registro * 100) / Tamanio, "0.0")
                lblPor.Caption = Format((Registro * 100) / Tamanio, "0.0") & " %"
            Loop
        End If
        rec.Close

    lblestado.Caption = ""
    DBConn.CommitTrans
    Screen.MousePointer = vbNormal
    Exit Sub

CLAVO:
 Screen.MousePointer = vbNormal
 lblestado.Caption = ""
 DBConn.RollbackTrans
 If rec.State = 1 Then rec.Close
 MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub
Private Sub InformeProvConcepto()

End Sub
Private Sub InformeFavor()
On Error GoTo CLAVO
     Screen.MousePointer = vbHourglass
     DBConn.BeginTrans
     lblestado.Caption = "Buscando Datos..."


        'BORRO LA TABLA TMP_LIBRO_IVA_COMPRAS
        sql = "DELETE FROM TMP_LIBRO_IVA_COMPRAS_PROVEEDOR_FAVOR"
        DBConn.Execute sql
        
       'BUSCO COMPROBANTES DENTRO DE LOS GASTOS GENERALES -----
        sql = "SELECT P.PROV_RAZSOC,GG.GGR_IVA,GG.GGR_IVA1,SUM(GG.GGR_NETO) AS NETO,SUM(GG.GGR_TOTAL) AS TOTAL,"
        sql = sql & " SUM(GG.GGR_NETO1) AS NETO1,SUM(GG.GGR_IMPUESTOS) AS IMPUESTO,SUM(GGR_PERIIBB) AS PERIIBB,SUM(GGR_PERIVA) AS PERIVA,SUM(GGR_PERGAN) AS PERGAN"
        sql = sql & " FROM GASTOS_GENERALES GG, PROVEEDOR P, TIPO_PROVEEDOR TP"
        sql = sql & " WHERE "
        sql = sql & " GG.PROV_CODIGO=P.PROV_CODIGO"
        sql = sql & " AND GG.TPR_CODIGO=TP.TPR_CODIGO"
        sql = sql & " AND P.TPR_CODIGO=TP.TPR_CODIGO"
        sql = sql & " AND GG.GGR_FAVOR=1" 'FAVOR
        sql = sql & " AND GG.GGR_LIBROIVA = " & XS("S") 'FAVOR(TODOS LOS FAVOR TIENEN IVA)
        If FechaDesde <> "" Then sql = sql & " AND GG.GGR_PERIODO>=" & XDQ(FechaDesde)
        If FechaHasta <> "" Then sql = sql & " AND GG.GGR_PERIODO<=" & XDQ(FechaHasta)
                
        sql = sql & " GROUP BY P.PROV_RAZSOC,GG.GGR_IVA,GG.GGR_IVA1"
        sql = sql & " ORDER BY P.PROV_RAZSOC"

        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Registro = 0
            Tamanio = rec.RecordCount
            Do While rec.EOF = False
                sql = "INSERT INTO TMP_LIBRO_IVA_COMPRAS_PROVEEDOR_FAVOR (PROVEEDOR,"
                sql = sql & "SUBTOTAL,IVA,TOTIVA,"
                sql = sql & "SUBTOTAL1,TOTOTROIVA,IMPUESTOS,TOTAL,RETIIBB,RETIVA,RETGAN,IVA1)"
                sql = sql & "VALUES ("
                sql = sql & XS(rec!PROV_RAZSOC) & ","
                sql = sql & XN(rec!NETO) & ","
                sql = sql & XN(rec!GGR_IVA) & ","
                    TotIva = (CDbl(rec!NETO) * CDbl(rec!GGR_IVA)) / 100
                sql = sql & XN(CStr(TotIva)) & ","
                sql = sql & XN(Chk0(rec!NETO1)) & "," 'OTRO NETO
                    TotIva = (CDbl(Chk0(rec!NETO1)) * CDbl(Chk0(rec!GGR_IVA1))) / 100
                
                sql = sql & XN(CStr(TotIva)) & "," 'OTRO IVA
                If rec!GGR_IVA = 0 Then
                    'PONGO EL TOTAL CUANDO EL GASTO CON IVA = 0
                    ' EJ MONOTRIBUTISTA
                    sql = sql & XN(Chk0(rec!TOTAL)) & "," 'IMPUESTOS
                Else
                    sql = sql & XN(Chk0(rec!IMPUESTO)) & "," 'IMPUESTOS
                End If
                sql = sql & XN(Chk0(rec!TOTAL)) & ","
                sql = sql & XN(Chk0(rec!PERIIBB)) & ","
                sql = sql & XN(Chk0(rec!PERIVA)) & ","
                sql = sql & XN(Chk0(rec!PERGAN)) & ","
                sql = sql & XN(Chk0(rec!GGR_IVA1)) & ")"
                DBConn.Execute sql
                rec.MoveNext

                Registro = Registro + 1
                ProgressBar1.Value = Format((Registro * 100) / Tamanio, "0.0")
                lblPor.Caption = Format((Registro * 100) / Tamanio, "0.0") & " %"
            Loop
        End If
        rec.Close


    lblestado.Caption = ""
    DBConn.CommitTrans
    Screen.MousePointer = vbNormal
    Exit Sub

CLAVO:
 Screen.MousePointer = vbNormal
 lblestado.Caption = ""
 DBConn.RollbackTrans
 If rec.State = 1 Then rec.Close
 MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub
Private Sub InformeProveedor()
On Error GoTo CLAVO
     Screen.MousePointer = vbHourglass
     DBConn.BeginTrans
     lblestado.Caption = "Buscando Datos..."


            'BORRO LA TABLA TMP_LIBRO_IVA_COMPRAS
        sql = "DELETE FROM TMP_LIBRO_IVA_COMPRAS_PROVEEDOR"
        DBConn.Execute sql
        
        'BUSCO COMPROBANTES DENTRO DE LOS GASTOS GENERALES -----
        'COMPROBANTES FISCALES
        sql = "SELECT P.PROV_RAZSOC,P.PROV_FANTASIA,GG.GGR_IVA,GG.GGR_IVA1,SUM(GG.GGR_NETO) AS NETO,SUM(GG.GGR_TOTAL) AS TOTAL,"
        sql = sql & " SUM(GG.GGR_NETO1) AS NETO1,SUM(GG.GGR_IMPUESTOS) AS IMPUESTO,SUM(GGR_PERIIBB) AS PERIIBB,SUM(GGR_PERIVA) AS PERIVA,SUM(GGR_PERGAN) AS PERGAN"
        sql = sql & " FROM GASTOS_GENERALES GG, PROVEEDOR P, TIPO_PROVEEDOR TP"
        sql = sql & " WHERE "
        sql = sql & " GG.PROV_CODIGO=P.PROV_CODIGO"
        sql = sql & " AND GG.TPR_CODIGO=TP.TPR_CODIGO"
        sql = sql & " AND P.TPR_CODIGO=TP.TPR_CODIGO"
        sql = sql & " AND GG.GGR_FAVOR=0" ' NORMAL (NO A FAVOR)
        sql = sql & " AND GG.GGR_LIBROIVA = " & XS("S")
        If FechaDesde <> "" Then sql = sql & " AND GG.GGR_PERIODO>=" & XDQ(FechaDesde)
        If FechaHasta <> "" Then sql = sql & " AND GG.GGR_PERIODO<=" & XDQ(FechaHasta)
        sql = sql & " GROUP BY P.PROV_RAZSOC,P.PROV_FANTASIA,GG.GGR_IVA,GG.GGR_IVA1"
        sql = sql & " ORDER BY P.PROV_RAZSOC"
        'COMPROBANTES NO FISCALES
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Registro = 0
            Tamanio = rec.RecordCount
            Do While rec.EOF = False
                sql = "INSERT INTO TMP_LIBRO_IVA_COMPRAS_PROVEEDOR (PROVEEDOR,"
                sql = sql & "SUBTOTAL,IVA,TOTIVA,"
                sql = sql & "SUBTOTAL1,TOTOTROIVA,IMPUESTOS,TOTAL,RETIIBB,RETIVA,RETGAN,IVA1,FANTASIA)"
                sql = sql & "VALUES ("
                sql = sql & XS(rec!PROV_RAZSOC) & ","
                sql = sql & XN(rec!NETO) & ","
                sql = sql & XN(rec!GGR_IVA) & ","
                    TotIva = (CDbl(rec!NETO) * CDbl(rec!GGR_IVA)) / 100
                sql = sql & XN(CStr(TotIva)) & ","
                sql = sql & XN(Chk0(rec!NETO1)) & "," 'OTRO NETO
                    TotIva = (CDbl(Chk0(rec!NETO1)) * CDbl(Chk0(rec!GGR_IVA1))) / 100
                
                sql = sql & XN(CStr(TotIva)) & "," 'OTRO IVA
                If rec!GGR_IVA = 0 Then
                    'PONGO EL TOTAL CUANDO EL GASTO CON IVA = 0
                    ' EJ MONOTRIBUTISTA
                    sql = sql & XN(Chk0(rec!TOTAL)) & "," 'IMPUESTOS
                Else
                    sql = sql & XN(Chk0(rec!IMPUESTO)) & "," 'IMPUESTOS
                End If
                sql = sql & XN(Chk0(rec!TOTAL)) & ","
                sql = sql & XN(Chk0(rec!PERIIBB)) & ","
                sql = sql & XN(Chk0(rec!PERIVA)) & ","
                sql = sql & XN(Chk0(rec!PERGAN)) & ","
                sql = sql & XN(Chk0(rec!GGR_IVA1)) & ","
                sql = sql & XS(ChkNull(rec!PROV_FANTASIA)) & ")"
                DBConn.Execute sql
                rec.MoveNext

                Registro = Registro + 1
                ProgressBar1.Value = Format((Registro * 100) / Tamanio, "0.0")
                lblPor.Caption = Format((Registro * 100) / Tamanio, "0.0") & " %"
            Loop
        End If
        rec.Close
        
        sql = " SELECT P.PROV_RAZSOC,P.PROV_FANTASIA,GG.GGR_IVA,GG.GGR_IVA1,SUM(GG.GGR_NETO) AS NETO,SUM(GG.GGR_TOTAL) AS TOTAL,"
        sql = sql & " SUM(GG.GGR_NETO1) AS NETO1,SUM(GG.GGR_IMPUESTOS) AS IMPUESTO,SUM(GGR_PERIIBB) AS PERIIBB,SUM(GGR_PERIVA) AS PERIVA,SUM(GGR_PERGAN) AS PERGAN"
        sql = sql & " FROM GASTOS_GENERALES GG, PROVEEDOR P, TIPO_PROVEEDOR TP"
        sql = sql & " WHERE "
        sql = sql & " GG.PROV_CODIGO=P.PROV_CODIGO"
        sql = sql & " AND GG.TPR_CODIGO=TP.TPR_CODIGO"
        sql = sql & " AND P.TPR_CODIGO=TP.TPR_CODIGO"
        sql = sql & " AND GG.GGR_FAVOR=0" ' NORMAL
        sql = sql & " AND GG.GGR_LIBROIVA=" & XS("N")
        If FechaDesde <> "" Then sql = sql & " AND (GG.GGR_FECHACOMP >=" & XDQ(FechaDesde) & " OR GG.GGR_PERIODO >=" & XDQ(FechaDesde) & " )"
        If FechaHasta <> "" Then sql = sql & " AND (GG.GGR_FECHACOMP<=" & XDQ(FechaHasta) & " OR GG.GGR_PERIODO <=" & XDQ(FechaHasta) & " )"
        sql = sql & " GROUP BY P.PROV_RAZSOC,P.PROV_FANTASIA,GG.GGR_IVA,GG.GGR_IVA1"
        sql = sql & " ORDER BY P.PROV_RAZSOC"
            
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Registro = 0
            Tamanio = rec.RecordCount
            Do While rec.EOF = False
                sql = "INSERT INTO TMP_LIBRO_IVA_COMPRAS_PROVEEDOR (PROVEEDOR,"
                sql = sql & "SUBTOTAL,IVA,TOTIVA,"
                sql = sql & "SUBTOTAL1,TOTOTROIVA,IMPUESTOS,TOTAL,RETIIBB,RETIVA,RETGAN,IVA1,FANTASIA)"
                sql = sql & "VALUES ("
                sql = sql & XS(rec!PROV_RAZSOC) & ","
                sql = sql & XN(rec!NETO) & ","
                sql = sql & XN(rec!GGR_IVA) & ","
                    TotIva = (CDbl(rec!NETO) * CDbl(rec!GGR_IVA)) / 100
                sql = sql & XN(CStr(TotIva)) & ","
                sql = sql & XN(Chk0(rec!NETO1)) & "," 'OTRO NETO
                    TotIva = (CDbl(Chk0(rec!NETO1)) * CDbl(Chk0(rec!GGR_IVA1))) / 100
                
                sql = sql & XN(CStr(TotIva)) & "," 'OTRO IVA
                If rec!GGR_IVA = 0 Then
                    'PONGO EL TOTAL CUANDO EL GASTO CON IVA = 0
                    ' EJ MONOTRIBUTISTA
                    sql = sql & XN(Chk0(rec!TOTAL)) & "," 'IMPUESTOS
                Else
                    sql = sql & XN(Chk0(rec!IMPUESTO)) & "," 'IMPUESTOS
                End If
                sql = sql & XN(Chk0(rec!TOTAL)) & ","
                sql = sql & XN(Chk0(rec!PERIIBB)) & ","
                sql = sql & XN(Chk0(rec!PERIVA)) & ","
                sql = sql & XN(Chk0(rec!PERGAN)) & ","
                sql = sql & XN(Chk0(rec!GGR_IVA1)) & ","
                sql = sql & XS(ChkNull(rec!PROV_FANTASIA)) & ")"
                DBConn.Execute sql
                rec.MoveNext

                Registro = Registro + 1
                ProgressBar1.Value = Format((Registro * 100) / Tamanio, "0.0")
                lblPor.Caption = Format((Registro * 100) / Tamanio, "0.0") & " %"
            Loop
        End If
        rec.Close


    lblestado.Caption = ""
    DBConn.CommitTrans
    Screen.MousePointer = vbNormal
    Exit Sub

CLAVO:
 Screen.MousePointer = vbNormal
 lblestado.Caption = ""
 DBConn.RollbackTrans
 If rec.State = 1 Then rec.Close
 MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub
Private Sub InformeConcepto(Numero As Integer)
     Dim Tabla As String
     On Error GoTo CLAVO
     Screen.MousePointer = vbHourglass
     DBConn.BeginTrans
     lblestado.Caption = "Buscando Datos..."
        
        If Numero = 1 Then
            Tabla = "TMP_LIBRO_IVA_COMPRAS_CONCEPTO"
        End If
        If Numero = 7 Then
            Tabla = "TMP_LIBRO_NO_IVA_COMPRAS_CONCEPTO"
        End If

        'BORRO LA TABLA TMP_LIBRO_IVA_COMPRAS
        sql = "DELETE FROM " & Tabla
        DBConn.Execute sql
        
        'BUSCO COMPROBANTES DENTRO DE LOS GASTOS GENERALES -----
        sql = "SELECT TG.TGT_DESCRI,GG.GGR_IVA,GG.GGR_IVA1,SUM(GG.GGR_NETO) AS NETO,SUM(GG.GGR_TOTAL) AS TOTAL,"
        sql = sql & " SUM(GG.GGR_NETO1) AS NETO1,SUM(GG.GGR_IMPUESTOS) AS IMPUESTO,SUM(GGR_PERIIBB) AS PERIIBB,SUM(GGR_PERIVA) AS PERIVA,SUM(GGR_PERGAN) AS PERGAN"
        sql = sql & " FROM GASTOS_GENERALES GG, TIPO_GASTO TG"
        sql = sql & " WHERE "
        sql = sql & " TG.TGT_CODIGO=GG.TGT_CODIGO"
        'sql = sql & " AND GG.TGT_CODIGO<>13" 'RETENCION
        
        If Numero = 1 Then
            sql = sql & " AND GG.GGR_LIBROIVA = " & XS("S")
            If FechaDesde <> "" Then sql = sql & " AND GG.GGR_PERIODO>=" & XDQ(FechaDesde)
            If FechaHasta <> "" Then sql = sql & " AND GG.GGR_PERIODO<=" & XDQ(FechaHasta)
        Else
            If Numero = 7 Then
                sql = sql & " AND GG.GGR_FAVOR<>1" '2-no IVA (no incluye ningun favor)
                sql = sql & " AND GG.GGR_LIBROIVA = " & XS("N")
                If FechaDesde <> "" Then sql = sql & " AND (GG.GGR_FECHACOMP >=" & XDQ(FechaDesde) & " OR GG.GGR_PERIODO >=" & XDQ(FechaDesde) & " )"
                If FechaHasta <> "" Then sql = sql & " AND (GG.GGR_FECHACOMP<=" & XDQ(FechaHasta) & " OR GG.GGR_PERIODO <=" & XDQ(FechaHasta) & " )"
            End If
        End If
        
        sql = sql & " GROUP BY TG.TGT_DESCRI,GG.GGR_IVA,GG.GGR_IVA1"
        sql = sql & " ORDER BY TG.TGT_DESCRI"

        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Registro = 0
            Tamanio = rec.RecordCount
            Do While rec.EOF = False
                sql = "INSERT INTO  " & Tabla & "  (CONCEPTO,"
                sql = sql & "SUBTOTAL,IVA,TOTIVA,"
                sql = sql & "SUBTOTAL1,TOTOTROIVA,IMPUESTOS,TOTAL,RETIIBB,RETIVA,RETGAN,IVA1)"
                sql = sql & "VALUES ("
                sql = sql & XS(rec!TGT_DESCRI) & ","
                sql = sql & XN(rec!NETO) & ","
                sql = sql & XN(rec!GGR_IVA) & ","
                    TotIva = (CDbl(rec!NETO) * CDbl(rec!GGR_IVA)) / 100
                sql = sql & XN(CStr(TotIva)) & ","
                sql = sql & XN(Chk0(rec!NETO1)) & "," 'OTRO NETO
                    TotIva = (CDbl(Chk0(rec!NETO1)) * CDbl(Chk0(rec!GGR_IVA1))) / 100
                
                sql = sql & XN(CStr(TotIva)) & "," 'OTRO IVA
                If rec!GGR_IVA = 0 Then
                    'PONGO EL TOTAL CUANDO EL GASTO CON IVA = 0
                    ' EJ MONOTRIBUTISTA
                    sql = sql & XN(Chk0(rec!TOTAL)) & "," 'IMPUESTOS
                Else
                    sql = sql & XN(Chk0(rec!IMPUESTO)) & "," 'IMPUESTOS
                End If
                sql = sql & XN(Chk0(rec!TOTAL)) & ","
                sql = sql & XN(Chk0(rec!PERIIBB)) & ","
                sql = sql & XN(Chk0(rec!PERIVA)) & ","
                sql = sql & XN(Chk0(rec!PERGAN)) & ","
                sql = sql & XN(Chk0(rec!GGR_IVA1)) & ")"
                DBConn.Execute sql
                rec.MoveNext

                Registro = Registro + 1
                ProgressBar1.Value = Format((Registro * 100) / Tamanio, "0.0")
                lblPor.Caption = Format((Registro * 100) / Tamanio, "0.0") & " %"
            Loop
        End If
        rec.Close
     
    lblestado.Caption = ""
    DBConn.CommitTrans
    Screen.MousePointer = vbNormal
    Exit Sub

CLAVO:
 Screen.MousePointer = vbNormal
 lblestado.Caption = ""
 DBConn.RollbackTrans
 If rec.State = 1 Then rec.Close
 MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub


Private Sub ListarLibroIVA(Numero As Integer)
    lblestado.Caption = "Buscando Listado..."
    Rep(Numero).WindowState = crptMaximized
    Rep(Numero).WindowBorderStyle = crptNoBorder
    Rep(Numero).Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=CENTENARO"
    Rep(Numero).Formulas(0) = ""
    Rep(Numero).Formulas(1) = ""
    Rep(Numero).Formulas(2) = ""
    Rep(Numero).Formulas(3) = ""
        
    sql = "SELECT CUIT,ING_BRUTOS,RAZ_SOCIAL,DIRECCION,TELEFONO FROM PARAMETROS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Rep(Numero).Formulas(0) = "EMPRESA='" & Trim(rec!RAZ_SOCIAL) & "'"
        Rep(Numero).Formulas(1) = "CUIT='" & Trim(rec!DIRECCION) & " - " & Trim(rec!TELEFONO) & " - " & Format(rec!cuit, "##-########-#") & "'"
        Rep(Numero).Formulas(2) = "INGBRUTOS='Ing. Brutos:  " & Format(rec!ING_BRUTOS, "###-#####-##") & "'"
    End If
    rec.Close
    
    If FechaDesde.Value <> "" And Not IsNull(FechaHasta.Value) Then
        Rep(Numero).Formulas(3) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & FechaHasta.Value & "'"
    ElseIf FechaDesde.Value <> "" And IsNull(FechaHasta.Value) Then
        Rep(Numero).Formulas(3) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & Date & "'"
    ElseIf IsNull(FechaDesde.Value) And Not IsNull(FechaHasta.Value) Then
        Rep(Numero).Formulas(3) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & FechaHasta.Value & "'"
    ElseIf IsNull(FechaDesde.Value) And IsNull(FechaHasta.Value) Then
        Rep(Numero).Formulas(3) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & Date & "'"
    End If
    
    If Numero = 0 Then
        Rep(Numero).WindowTitle = "Libro I.V.A. Compras"
        Rep(Numero).ReportFileName = DRIVE & DirReport & "rptlibroivacompras.rpt"
    End If
    If Numero = 4 Then
        Rep(Numero).WindowTitle = "Libro I.V.A. Compras - COMBUSTIBLES"
        'rep(Numero).Formulas(0) = ""
        'rep(Numero).Formulas(1) = ""
        'rep(Numero).Formulas(2) = ""
        'rep(Numero).Formulas(3) = ""
        Rep(Numero).ReportFileName = DRIVE & DirReport & "rptlibroivacompras_combustibles.rpt"
    End If
    If Numero = 6 Then
        Rep(Numero).WindowTitle = "Libro I.V.A. Compras - NO INCLUIDAS"
        Rep(Numero).ReportFileName = DRIVE & DirReport & "rptlibro_NO_ivacompras.rpt"
    End If
    
    If optPantalla.Value = True Then
        Rep(Numero).Destination = crptToWindow
    ElseIf optImpresora.Value = True Then
        Rep(Numero).Destination = crptToPrinter
    End If
     Rep(Numero).Action = 1
     
     lblestado.Caption = ""
     Rep(Numero).Formulas(0) = ""
     Rep(Numero).Formulas(1) = ""
     Rep(Numero).Formulas(2) = ""
     Rep(Numero).Formulas(3) = ""
End Sub
Private Sub ListarLibroIVA_Combustibles(Numero As Integer)
    lblestado.Caption = "Buscando Listado..."
    Rep(Numero).WindowState = crptMaximized
    Rep(Numero).WindowBorderStyle = crptNoBorder
    Rep(Numero).Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=CENTENARO"
    Rep(Numero).Formulas(0) = ""
    Rep(Numero).Formulas(1) = ""
    Rep(Numero).Formulas(2) = ""
    Rep(Numero).Formulas(3) = ""
        
    sql = "SELECT CUIT,ING_BRUTOS,RAZ_SOCIAL,DIRECCION,TELEFONO FROM PARAMETROS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Rep(Numero).Formulas(0) = "EMPRESA='" & Trim(rec!RAZ_SOCIAL) & "'"
        Rep(Numero).Formulas(1) = "CUIT='" & Trim(rec!DIRECCION) & " - " & Trim(rec!TELEFONO) & " - " & Format(rec!cuit, "##-########-#") & "'"
        Rep(Numero).Formulas(2) = "INGBRUTOS='" & "Ing. Brutos:  " & Format(rec!ING_BRUTOS, "###-#####-##") & "'"
    End If
    rec.Close
    
    If FechaDesde.Value <> "" And Not IsNull(FechaHasta.Value) Then
        Rep(Numero).Formulas(3) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & FechaHasta.Value & "'"
    ElseIf FechaDesde.Value <> "" And IsNull(FechaHasta.Value) Then
        Rep(Numero).Formulas(3) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & Date & "'"
    ElseIf IsNull(FechaDesde.Value) And Not IsNull(FechaHasta.Value) Then
        Rep(Numero).Formulas(3) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & FechaHasta.Value & "'"
    ElseIf IsNull(FechaDesde.Value) And IsNull(FechaHasta.Value) Then
        Rep(Numero).Formulas(3) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & Date & "'"
    End If
    
    If Numero = 4 Then
        Rep(Numero).WindowTitle = "Libro I.V.A. Compras - COMBUSTIBLES"
        Rep(Numero).ReportFileName = DRIVE & DirReport & "rptlibroivacompras_combustibles.rpt"
    End If
    
    If optPantalla.Value = True Then
        Rep(Numero).Destination = crptToWindow
    ElseIf optImpresora.Value = True Then
        Rep(Numero).Destination = crptToPrinter
    End If
     Rep(Numero).Action = 1
     
     lblestado.Caption = ""
     Rep(Numero).Formulas(0) = ""
     Rep(Numero).Formulas(1) = ""
     Rep(Numero).Formulas(2) = ""
     Rep(Numero).Formulas(3) = ""
     
End Sub
Private Sub ListarProveedor(Numero As Integer)
    lblestado.Caption = "Buscando Listado..."
    Rep(Numero).WindowState = crptMaximized
    Rep(Numero).WindowBorderStyle = crptNoBorder
    Rep(Numero).Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=CENTENARO"
    Rep(Numero).Formulas(0) = ""
    Rep(Numero).Formulas(1) = ""
    Rep(Numero).Formulas(2) = ""
    Rep(Numero).Formulas(3) = ""
    Rep(Numero).Formulas(4) = ""
        
    sql = "SELECT CUIT,ING_BRUTOS,RAZ_SOCIAL,DIRECCION,TELEFONO FROM PARAMETROS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Rep(Numero).Formulas(0) = "EMPRESA='" & Trim(rec!RAZ_SOCIAL) & "'"
        Rep(Numero).Formulas(1) = "CUIT='" & Trim(rec!DIRECCION) & " - " & Trim(rec!TELEFONO) & " - " & Format(rec!cuit, "##-########-#") & "'"
        Rep(Numero).Formulas(2) = "INGBRUTOS='Ing. Brutos:  " & Format(rec!ING_BRUTOS, "###-#####-##") & "'"
    End If
    rec.Close
    
    If FechaDesde.Value <> "" And Not IsNull(FechaHasta.Value) Then
        Rep(Numero).Formulas(3) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & FechaHasta.Value & "'"
    ElseIf FechaDesde.Value <> "" And IsNull(FechaHasta.Value) Then
        Rep(Numero).Formulas(3) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & Date & "'"
    ElseIf IsNull(FechaDesde.Value) And Not IsNull(FechaHasta.Value) Then
        Rep(Numero).Formulas(3) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & FechaHasta.Value & "'"
    ElseIf IsNull(FechaDesde.Value) And IsNull(FechaHasta.Value) Then
        Rep(Numero).Formulas(3) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & Date & "'"
    End If
    
    If Numero = 2 Then
        Rep(Numero).WindowTitle = "Informe de Compras por Proveedor"
        Rep(Numero).Formulas(4) = "TIPO='" & "FAVOR" & "'"
        Rep(Numero).ReportFileName = DRIVE & DirReport & "rptlibroivacompras_proveedor.rpt"
    Else
        If Numero = 3 Then
            Rep(Numero).WindowTitle = "Informe de Favor por Proveedor"
            Rep(Numero).Formulas(4) = "TIPO='" & "NORMAL" & "'"
            Rep(Numero).ReportFileName = DRIVE & DirReport & "rptlibroivacompras_FAVOR_proveedor.rpt"
        End If
    End If
    'If chkResumen Then
     
    'Else
    '    rep(Numero).WindowTitle = "NO - Informe por Concepto I.V.A. Compras"
    'End If
    
    
    
    If optPantalla.Value = True Then
        Rep(Numero).Destination = crptToWindow
    ElseIf optImpresora.Value = True Then
        Rep(Numero).Destination = crptToPrinter
    End If
     Rep(Numero).Action = 1
     
     lblestado.Caption = ""
     Rep(Numero).Formulas(0) = ""
     Rep(Numero).Formulas(1) = ""
     Rep(Numero).Formulas(3) = ""

End Sub
Private Sub ListarResumenIVA(Numero As Integer)
    lblestado.Caption = "Buscando Listado..."
    Rep(Numero).WindowState = crptMaximized
    Rep(Numero).WindowBorderStyle = crptNoBorder
    Rep(Numero).Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=CENTENARO"
    Rep(Numero).Formulas(0) = ""
    Rep(Numero).Formulas(1) = ""
    Rep(Numero).Formulas(2) = ""
    'rep(Numero).Formulas(3) = ""
        
    sql = "SELECT CUIT,ING_BRUTOS,RAZ_SOCIAL,DIRECCION,TELEFONO FROM PARAMETROS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Rep(Numero).Formulas(0) = "EMPRESA='" & Trim(rec!RAZ_SOCIAL) & "'"
        Rep(Numero).Formulas(1) = "CUIT='" & Trim(rec!DIRECCION) & " - " & Trim(rec!TELEFONO) & " - " & Format(rec!cuit, "##-########-#") & "'"
        Rep(Numero).Formulas(2) = "INGBRUTOS='Ing. Brutos:  " & Format(rec!ING_BRUTOS, "###-#####-##") & "'"
    End If
    rec.Close
    
    If FechaDesde.Value <> "" And Not IsNull(FechaHasta.Value) Then
        Rep(Numero).Formulas(3) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & FechaHasta.Value & "'"
    ElseIf FechaDesde.Value <> "" And IsNull(FechaHasta.Value) Then
        Rep(Numero).Formulas(3) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & Date & "'"
    ElseIf IsNull(FechaDesde.Value) And Not IsNull(FechaHasta.Value) Then
        Rep(Numero).Formulas(3) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & FechaHasta.Value & "'"
    ElseIf IsNull(FechaDesde.Value) And IsNull(FechaHasta.Value) Then
        Rep(Numero).Formulas(3) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & Date & "'"
    End If
    
    If Numero = 1 Then
        Rep(Numero).WindowTitle = "Informe por Concepto I.V.A. Compras"
        Rep(Numero).ReportFileName = DRIVE & DirReport & "rptlibroivacompras_conceptos.rpt"
    Else
        If Numero = 7 Then
            Rep(Numero).WindowTitle = "NO - Informe por Concepto I.V.A. Compras"
            Rep(Numero).ReportFileName = DRIVE & DirReport & "rptlibro_NO_ivacompras_conceptos.rpt"
        End If
    End If
    
    
    
    If optPantalla.Value = True Then
        Rep(Numero).Destination = crptToWindow
    ElseIf optImpresora.Value = True Then
        Rep(Numero).Destination = crptToPrinter
    End If
     Rep(Numero).Action = 1
     
     lblestado.Caption = ""
     Rep(Numero).Formulas(0) = ""
     Rep(Numero).Formulas(1) = ""
     Rep(Numero).Formulas(2) = ""
     'rep(Numero).Formulas(3) = ""
     
End Sub
Private Sub CmdNuevo_Click()
    Dim i As Integer
    FechaDesde.Value = ""
    lblPeriodo1.Caption = ""
    FechaHasta.Value = ""
    lblPeriodo2.Caption = ""
    txtHoja.Text = ""
    FechaDesde.SetFocus
    For i = 0 To 10
        chkreporte(i).Value = Unchecked
    Next
    chkTodos.Value = Unchecked
End Sub

Private Sub CmdSalir_Click()
    Set frmLibroCompras2 = Nothing
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    lblestado.Caption = ""
    lblPor.Caption = "100 %"
    Call Centrar_pantalla(Me)
    Set rec = New ADODB.Recordset
    FrameImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
End Sub

Private Sub FechaDesde_LostFocus()
    If Trim(FechaDesde.Value) <> "" Then
        FechaHasta.Value = FechaDesde.Value
        lblPeriodo1.Caption = UCase(Format(FechaDesde.Value, "mmmm/yyyy"))
    Else
        lblPeriodo1.Caption = ""
    End If
End Sub

Private Sub FechaHasta_LostFocus()
    If Trim(FechaHasta.Value) <> "" Then
        lblPeriodo2.Caption = UCase(Format(FechaHasta.Value, "mmmm/yyyy"))
    Else
        lblPeriodo2.Caption = ""
    End If
End Sub

Private Sub Option2_Click()

End Sub

Private Sub OptProveedor_Click()

End Sub

Private Sub txtHoja_GotFocus()
    seltxt
End Sub

Private Sub txtHoja_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub
