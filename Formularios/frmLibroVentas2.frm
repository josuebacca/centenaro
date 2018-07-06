VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLibroVentas2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Libro IVA Ventas"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optCierreZ 
      Caption         =   "Cierres Z"
      Height          =   195
      Left            =   360
      TabIndex        =   32
      Top             =   2280
      Width           =   1050
   End
   Begin VB.OptionButton optlibro 
      Caption         =   "Libro IVA"
      Height          =   195
      Left            =   360
      TabIndex        =   31
      Top             =   120
      Value           =   -1  'True
      Width           =   1245
   End
   Begin VB.Frame fracierrez 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   120
      TabIndex        =   22
      Top             =   2280
      Width           =   5355
      Begin VB.TextBox txtZdesde 
         Height          =   315
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   23
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtZHasta 
         Height          =   315
         Left            =   3960
         MaxLength       =   6
         TabIndex        =   24
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cierre Z Desde:"
         Height          =   195
         Left            =   480
         TabIndex        =   26
         Top             =   540
         Width           =   1110
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Cierre Z Hasta:"
         Height          =   195
         Left            =   2805
         TabIndex        =   25
         Top             =   540
         Width           =   1065
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
      TabIndex        =   17
      Top             =   3600
      Width           =   5490
      Begin VB.OptionButton optImpresora 
         Caption         =   "Impresora"
         Height          =   195
         Left            =   2370
         TabIndex        =   7
         Top             =   315
         Width           =   1050
      End
      Begin VB.OptionButton optPantalla 
         Caption         =   "Pantalla"
         Height          =   195
         Left            =   945
         TabIndex        =   6
         Top             =   315
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.CommandButton CBImpresora 
         Caption         =   "Configurar Impresora"
         Height          =   375
         Left            =   3570
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   1665
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   150
         TabIndex        =   19
         Top             =   315
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   735
      Left            =   2925
      Picture         =   "frmLibroVentas2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4410
      Width           =   840
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   4635
      Picture         =   "frmLibroVentas2.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4410
      Width           =   840
   End
   Begin VB.Frame fralibro 
      Height          =   2160
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   5355
      Begin VB.TextBox txtFBDesde 
         Height          =   315
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   4
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtFBHasta 
         Height          =   315
         Left            =   4320
         MaxLength       =   6
         TabIndex        =   5
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtFADesde 
         Height          =   315
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   2
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtFAHasta 
         Height          =   315
         Left            =   4320
         MaxLength       =   6
         TabIndex        =   3
         Top             =   1200
         Width           =   735
      End
      Begin MSComCtl2.DTPicker FechaDesde 
         Height          =   315
         Left            =   1680
         TabIndex        =   0
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   51904513
         CurrentDate     =   41098
      End
      Begin MSComCtl2.DTPicker FechaHasta 
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   51904513
         CurrentDate     =   41098
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "FB Desde:"
         Height          =   195
         Left            =   840
         TabIndex        =   30
         Top             =   1620
         Width           =   750
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "FB Hasta:"
         Height          =   195
         Left            =   3525
         TabIndex        =   29
         Top             =   1620
         Width           =   705
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "FA Desde:"
         Height          =   195
         Left            =   840
         TabIndex        =   28
         Top             =   1260
         Width           =   750
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "FA Hasta:"
         Height          =   195
         Left            =   3525
         TabIndex        =   27
         Top             =   1260
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   600
         TabIndex        =   15
         Top             =   405
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   645
         TabIndex        =   14
         Top             =   750
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
         Left            =   3225
         TabIndex        =   13
         Top             =   390
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
         Left            =   3225
         TabIndex        =   12
         Top             =   720
         Width           =   1785
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmLibroVentas2.frx":0614
      Height          =   735
      Left            =   3780
      Picture         =   "frmLibroVentas2.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4410
      Width           =   840
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   1590
      Top             =   4590
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   2130
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   225
      Left            =   120
      TabIndex        =   20
      Top             =   3360
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   397
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label lblPor 
      AutoSize        =   -1  'True
      Caption         =   "100 %"
      Height          =   195
      Left            =   4995
      TabIndex        =   21
      Top             =   3390
      Width           =   435
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
      Left            =   90
      TabIndex        =   16
      Top             =   4575
      Width           =   750
   End
End
Attribute VB_Name = "frmLibroVentas2"
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

Private Function actualizar_cierrez(zeta As Long)
    Dim ultima_fa As Long
    Dim ultima_fb As Long
    Dim ultima_fa_anterior As Long
    Dim ultima_fb_anterior As Long
    ultima_fa_anterior = BUSCO_UFACZANT(zeta, 1)
    ultima_fb_anterior = BUSCO_UFACZANT(zeta, 2)
    ultima_fa = BUSCO_UFACZANT(zeta + 1, 1)
    ultima_fb = BUSCO_UFACZANT(zeta + 1, 2)
    ultima_fa = ultima_fa - 1
    ultima_fb = ultima_fb - 1
    'MsgBox ultima_fa_anterior, ultima_fb_anterior
    'MsgBox ultima_fa, ultima_fb
    
    sql = "UPDATE FACTURA_CLIENTE"
    sql = sql & " SET FCL_CIERREZ = " & zeta
    sql = sql & " WHERE FCL_NUMERO BETWEEN " & ultima_fa_anterior & " AND " & ultima_fa
    sql = sql & " AND TCO_CODIGO = 1"
    DBConn.Execute sql
    
    sql = "UPDATE FACTURA_CLIENTE"
    sql = sql & " SET FCL_CIERREZ = " & zeta
    sql = sql & " WHERE FCL_NUMERO BETWEEN " & ultima_fb_anterior & " AND " & ultima_fb
    sql = sql & " AND TCO_CODIGO = 2"
    DBConn.Execute sql
    
    
End Function

Private Sub cmdAceptar_Click()
     Registro = 0
     Tamanio = 0
     TotIva = 0
     
'     If FechaDesde.value = "" Then
'        MsgBox "Debe ingresar el periodo", vbExclamation, TIT_MSGBOX
'        FechaDesde.SetFocus
'        Exit Sub
'     End If
'     If FechaHasta.value = "" Then
'        MsgBox "Debe ingresar el periodo", vbExclamation, TIT_MSGBOX
'        FechaHasta.SetFocus
'        Exit Sub
'     End If
     
     'On Error GoTo CLAVO
     Screen.MousePointer = vbHourglass
     'DBConn.BeginTrans
     lblestado.Caption = "Buscando Datos..."
     
    If optlibro.Value = True Then
        'BORRO LA TABLA TEMPORAL DE IVA VENTAS
        sql = "DELETE FROM TMP_LIBRO_IVA_VENTAS"
        DBConn.Execute sql
        
        'FACTURAS
        BUSCO_FACTURAS
        'NOTAS DE CREDITO
        'BUSCO_NOTA_CREDITO
        'NOTAS DE DEBITO
        'BUSCO_NOTA_DEBITO
        'RETENCIONES
        'BUSCO_RETENCIONES
            
        lblestado.Caption = ""
        'DBConn.CommitTrans
        'cargo el reporte
        ListarLibroIVA
    Else
         
        sql = "DELETE FROM TMP_CIERREZ"
        DBConn.Execute sql
        'PROBANDO LIBRO IVA POR CIERRE Z
        BUSCO_FACTURAS_CIERREZ
        'ListarLibroIVA
        
        BUSCO_CIERRESZ
        lblestado.Caption = ""
        'DBConn.CommitTrans
        ListarCierreZ
    End If
    Screen.MousePointer = vbNormal
    
    Exit Sub

'CLAVO:
' Screen.MousePointer = vbNormal
' lblEstado.Caption = ""
' DBConn.RollbackTrans
' If rec.State = 1 Then rec.Close
' MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub
Private Function CalculoTasaVial(ftipo As Integer, fdesde As Long, fhasta As Long) As Double
    sql = "SELECT SUM(FCL_TASAVIAL) AS TASAVIAL, TCO_CODIGO"
    sql = sql & " FROM FACTURA_CLIENTE"
    sql = sql & " WHERE TCO_CODIGO = " & ftipo
    sql = sql & " AND  FCL_NUMERO >=" & fdesde
    sql = sql & " AND  FCL_NUMERO <=" & fhasta
    sql = sql & " AND  EST_CODIGO = 3"
    sql = sql & " GROUP BY TCO_CODIGO"
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        CalculoTasaVial = Chk0(Rec1!TASAVIAL)
    Else
        CalculoTasaVial = 0
    End If
    Rec1.Close

End Function
Private Function CalculoImpuesto(ftipo As Integer, fdesde As Long, fhasta As Long) As Double
    sql = "SELECT SUM(FCL_IMPINT) AS IMPINT, TCO_CODIGO"
    sql = sql & " FROM FACTURA_CLIENTE"
    sql = sql & " WHERE TCO_CODIGO = " & ftipo
    sql = sql & " AND  FCL_SUCURSAL = 5"
    sql = sql & " AND  FCL_NUMERO >=" & fdesde
    sql = sql & " AND  FCL_NUMERO <=" & fhasta
    sql = sql & " AND  EST_CODIGO = 3"
    sql = sql & " GROUP BY TCO_CODIGO"
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        CalculoImpuesto = Chk0(Rec1!IMPINT)
    Else
        CalculoImpuesto = 0
    End If
    Rec1.Close

End Function
Private Function BUSCO_UFACZANT(z_nro As Long, ftipo As Integer) As Long
    Dim zeta As Long
    Dim EXIST As Integer '1 SI 0 NO
    zeta = z_nro - 1
    EXIST = 0
    Do While EXIST = 0
        sql = "SELECT * FROM CIERREZ WHERE Z_NUMERO = " & zeta
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            EXIST = 1
        Else
            zeta = zeta - 1
        End If
        Rec1.Close
    Loop
    
    sql = "SELECT Z_ULTIMA_FACTURA,Z_ULTIMO_TICKET "
    sql = sql & " FROM CIERREZ "
    sql = sql & " WHERE Z_NUMERO=" & zeta
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        If ftipo = 1 Then
            BUSCO_UFACZANT = Chk0(Rec1!Z_ULTIMA_FACTURA) + 1
        Else 'ftipo= 2
            BUSCO_UFACZANT = Chk0(Rec1!Z_ULTIMO_TICKET) + 1
        End If
    Else
        BUSCO_UFACZANT = 0
    End If
    Rec1.Close
    sql = "UPDATE FACTURA_CLIENTE "
    sql = sql & " SET FCL_CIERREZ=" & zeta
    sql = sql & " WHERE FCL_NUMERO BETWEEN "
    
End Function
Private Function BUSCO_CIERRESZ()
    Dim A_TazaVial As Double
    Dim B_TazaVial As Double
    Dim TazaVial As String
    Dim A_Imp As Double
    Dim B_Imp As Double
    Dim Impuestos As String
    
    TotIva = 0
    'BUSCO FACTURAS POR REMITO ---------------------------------
    sql = "SELECT * "
    sql = sql & "FROM CIERREZ WHERE 1=1 "
    If FechaDesde <> "" Then sql = sql & " AND Z_FECHA>=" & XDQ(FechaDesde)
    If FechaHasta <> "" Then sql = sql & " AND Z_FECHA<=" & XDQ(FechaHasta)
    If txtZdesde.Text <> "" Then sql = sql & " AND Z_NUMERO >=" & txtZdesde
    If txtZHasta.Text <> "" Then sql = sql & " AND Z_NUMERO <=" & txtZHasta
    sql = sql & " ORDER BY Z_NUMERO, Z_FECHA"
    
'    If txtZdesde.Text <> "" Or txtZHasta.Text <> "" Then
'        'BUSCAR Y GENERAR CIERREZ PERDIDOS
'
'    End If
    

    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Tamanio = rec.RecordCount
        Do While rec.EOF = False
            actualizar_cierrez rec!Z_NUMERO
            'A_TazaVial = CalculoTasaVial(1, BUSCO_UFACZANT(rec!Z_NUMERO, 1), rec!Z_ULTIMA_FACTURA)
            'B_TazaVial = CalculoTasaVial(2, BUSCO_UFACZANT(rec!Z_NUMERO, 2), rec!Z_ULTIMO_TICKET)
            'TazaVial = A_TazaVial + B_TazaVial
            TazaVial = 0
            A_Imp = CalculoImpuesto(1, BUSCO_UFACZANT(rec!Z_NUMERO, 1), rec!Z_ULTIMA_FACTURA)
            B_Imp = CalculoImpuesto(2, BUSCO_UFACZANT(rec!Z_NUMERO, 2), rec!Z_ULTIMO_TICKET)
            Impuestos = A_Imp + B_Imp
            
            sql = "INSERT INTO TMP_CIERREZ (Z_SECUENCIA,Z_NUMERO,Z_FECHA,Z_NETO,Z_IVA,Z_IMPUESTOS,Z_TASAVIAL,Z_TOTAL)"
            sql = sql & "VALUES ("
            sql = sql & XN(rec!Z_SECUENCIA) & ","
            sql = sql & XN(rec!Z_NUMERO) & ","
            sql = sql & XDQ(rec!Z_FECHA) & ","
            sql = sql & XN(rec!Z_TOTAL - (rec!Z_IVA + CDbl(Impuestos))) & "," 'neto
            sql = sql & XN(rec!Z_IVA) & "," 'iva
            sql = sql & XN(Impuestos) & "," 'impuestos
            sql = sql & XN(TazaVial) & "," ' tasa_vial
            sql = sql & XN(rec!Z_TOTAL) & ")"
            DBConn.Execute sql
            rec.MoveNext
            
            Registro = Registro + 1
'            ProgressBar1.Value = Format((Registro * 100) / Tamanio, "0.0")
 '           lblPor.Caption = Format((Registro * 100) / Tamanio, "0.0") & " %"
        Loop
    End If
    rec.Close
End Function
'Private Function GenerarCierreZPerdido()
'
'End Function
'Private Function BuscarCierreZPerdido()
'    Dim Rec2 As ADODB.Recordset
'    Set Rec2 = New ADODB.Recordset
'    sql = "SELECT * "
'    sql = sql & "FROM CIERREZ "
'    sql = sql & " WHERE Z_NUMERO >=" & txtZdesde
'    sql = sql & " AND Z_NUMERO <=" & txtZHasta
'
'    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    If Rec1!EOF = False Then
'        Do While Rec1.EOF = False
'            sql = "SELECT * "
'            sql = sql & "FROM CIERREZ "
'            sql = sql & " WHERE Z_NUMERO >=" & txtZdesde
'            sql = sql & " AND Z_NUMERO <=" & txtZHasta
'            sql = sql & " AND Z_NUMERO =" & Rec1!Z_NUMERO
'            Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
'            If Rec2!EOF = False Then
'                GenerarCierreZPerdido
'            End If
'            Rec2.Close
'            Rec1.MoveNext
'        Loop
'    End If
'    Rec1.Close
'
'End Function
Private Sub ListarLibroIVA()
    lblestado.Caption = "Buscando Listado..."
    Rep.WindowState = crptMaximized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SIHDG"
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    Rep.Formulas(2) = ""
        
    sql = "SELECT CUIT,ING_BRUTOS,RAZ_SOCIAL FROM PARAMETROS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Rep.Formulas(0) = "EMPRESA='     Empresa:  " & Trim(rec!RAZ_SOCIAL) & "'"
        Rep.Formulas(1) = "CUIT='       C.U.I.T.:  " & Format(rec!cuit, "##-########-#") & "'"
        Rep.Formulas(2) = "INGBRUTOS='Ing. Brutos:  " & Format(rec!ING_BRUTOS, "###-#####-##") & "'"
    End If
    rec.Close
    
    Rep.WindowTitle = "Libro I.V.A. Ventas"
    Rep.ReportFileName = DRIVE & DirReport & "rptlibroivaventas.rpt"
    
    If optPantalla.Value = True Then
        Rep.Destination = crptToWindow
    ElseIf optImpresora.Value = True Then
        Rep.Destination = crptToPrinter
    End If
     Rep.Action = 1
     
     lblestado.Caption = ""
     Rep.Formulas(0) = ""
     Rep.Formulas(1) = ""
     Rep.Formulas(2) = ""
     Rep.Connect = ""
     Rep.Reset
     
End Sub
Private Sub ListarCierreZ()
    lblestado.Caption = "Buscando Listado..."
    Rep.WindowState = crptMaximized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SIHDG"
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    Rep.Formulas(2) = ""
        
    sql = "SELECT CUIT,ING_BRUTOS,RAZ_SOCIAL FROM PARAMETROS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Rep.Formulas(0) = "EMPRESA='     Empresa:  " & Trim(rec!RAZ_SOCIAL) & "'"
        Rep.Formulas(1) = "CUIT='       C.U.I.T.:  " & Format(rec!cuit, "##-########-#") & "'"
        Rep.Formulas(2) = "INGBRUTOS='Ing. Brutos:  " & Format(rec!ING_BRUTOS, "###-#####-##") & "'"
    End If
    rec.Close
    
    Rep.WindowTitle = "Cierres Z"
    Rep.ReportFileName = DRIVE & DirReport & "rptcierresz.rpt"
    
    If optPantalla.Value = True Then
        Rep.Destination = crptToWindow
    ElseIf optImpresora.Value = True Then
        Rep.Destination = crptToPrinter
    End If
     Rep.Action = 1
     
     lblestado.Caption = ""
     Rep.Formulas(0) = ""
     Rep.Formulas(1) = ""
     Rep.Formulas(2) = ""
     Rep.Reset
End Sub

Private Sub CmdNuevo_Click()
    FechaDesde.Value = ""
    lblPeriodo1.Caption = ""
    FechaHasta.Value = ""
    lblPeriodo2.Caption = ""
    optlibro.Value = True
'    FechaDesde.SetFocus
    txtFADesde.Text = ""
    txtFAHasta.Text = ""
    txtFBDesde.Text = ""
    txtFBHasta.Text = ""
    txtZdesde.Text = ""
    txtZHasta.Text = ""
End Sub

Private Sub CmdSalir_Click()
    Set frmLibroVentas2 = Nothing
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

Private Sub BUSCO_FACTURAS_CIERREZ()
    'Dim SecDesde As Long
    'Dim SecHasta As Long
    Dim sql1 As String
    Dim FacADesde As Long
    Dim FacAHasta As Long
    Dim FacBDesde As Long
    Dim FacBHasta As Long
    TotIva = 0
    'BUSCO FACTURAS POR REMITO fac A---------------------------------
    sql1 = "SELECT FC.FCL_NUMERO, FC.FCL_SUCURSAL, FC.FCL_FECHA, FC.FCL_IVA,"
    sql1 = sql1 & " FC.FCL_SUBTOTAL, FC.FCL_TOTAL,"
    sql1 = sql1 & " FC.EST_CODIGO, C.CLI_CUIT, C.CLI_INGBRU,FCL_TASAVIAL,"
    sql1 = sql1 & " C.CLI_RAZSOC, TC.TCO_ABREVIA,FC.FCL_IMPINT"
    sql1 = sql1 & " FROM FACTURA_CLIENTE FC, CLIENTE C,"
    sql1 = sql1 & " TIPO_COMPROBANTE TC"
    sql1 = sql1 & " WHERE"
    sql1 = sql1 & " FC.TCO_CODIGO=TC.TCO_CODIGO"
    sql1 = sql1 & " AND FC.CLI_CODIGO=C.CLI_CODIGO"
    'busco todas las facs
    'sql1 = sql1 & " AND FC.EST_CODIGO = 3" 'ESTADO DEFINITIVO Y ANULADO
    'If FechaDesde <> "" Then sql1 = sql1 & " AND FC.FCL_FECHA>=" & XDQ(FechaDesde)
    'If FechaHasta <> "" Then sql1 = sql1 & " AND FC.FCL_FECHA<=" & XDQ(FechaHasta)
'    If txtZdesde.Text <> "" Then
'        FacADesde = BUSCO_UFACZANT(txtZdesde.Text, 1)
'        sql1 = sql1 & " AND FC.TCO_CODIGO=1 AND FC.FCL_NUMERO >= " & FacADesde
'    End If
'    If txtZHasta.Text <> "" Then
'        FacADesde = BUSCO_UFACZANT(txtZHasta.Text, 1)
'        sql1 = sql1 & " AND FC.TCO_CODIGO=1 AND FC.FCL_NUMERO <= " & FacAHasta
'    End If
    'If txtFADesde.Text <> "" Then
    '    sql1 = sql1 & " AND FC.TCO_CODIGO=1 AND FC.FCL_NUMERO >= " & XN(txtFADesde.Text)
    'End If
    'If txtFAHasta.Text <> "" Then
    '    sql1 = sql1 & " AND FC.TCO_CODIGO=1 AND FC.FCL_NUMERO <= " & XN(txtFAHasta.Text)
    'End If
    
    If txtZdesde.Text <> "" Then
        sql1 = sql1 & " AND FC.TCO_CODIGO=1 AND FC.FCL_CIERREZ >= " & XN(txtZdesde.Text)
    End If
    If txtZHasta.Text <> "" Then
        sql1 = sql1 & " AND FC.TCO_CODIGO=1 AND FC.FCL_CIERREZ <= " & XN(txtZHasta.Text)
    End If
    
    sql1 = sql1 & " UNION "
    sql1 = sql1 & " SELECT FC.FCL_NUMERO, FC.FCL_SUCURSAL, FC.FCL_FECHA, FC.FCL_IVA,"
    sql1 = sql1 & " FC.FCL_SUBTOTAL, FC.FCL_TOTAL,"
    sql1 = sql1 & " FC.EST_CODIGO, C.CLI_CUIT, C.CLI_INGBRU,FCL_TASAVIAL,"
    sql1 = sql1 & " C.CLI_RAZSOC, TC.TCO_ABREVIA,FC.FCL_IMPINT"
    sql1 = sql1 & " FROM FACTURA_CLIENTE FC, CLIENTE C,"
    sql1 = sql1 & " TIPO_COMPROBANTE TC"
    sql1 = sql1 & " WHERE"
    sql1 = sql1 & " FC.TCO_CODIGO=TC.TCO_CODIGO"
    sql1 = sql1 & " AND FC.CLI_CODIGO=C.CLI_CODIGO"
    'sql1 = sql1 & " AND FC.EST_CODIGO = 3" 'ESTADO DEFINITIVO Y ANULADO
    'If FechaDesde <> "" Then sql1 = sql1 & " AND FC.FCL_FECHA>=" & XDQ(FechaDesde)
    'If FechaHasta <> "" Then sql1 = sql1 & " AND FC.FCL_FECHA<=" & XDQ(FechaHasta)
'    If txtZdesde.Text <> "" Then
'        FacBDesde = BUSCO_UFACZANT(txtZdesde.Text, 2)
'        sql1 = sql1 & " AND FC.TCO_CODIGO=2 AND FC.FCL_NUMERO >= " & FacBDesde
'    End If
'    If txtZHasta.Text <> "" Then
'        FacBDesde = BUSCO_UFACZANT(txtZHasta.Text, 2)
'        sql1 = sql1 & " AND FC.TCO_CODIGO=2 AND FC.FCL_NUMERO <= " & FacBHasta
'    End If
    'If txtFBDesde.Text <> "" Then
    '    sql1 = sql1 & " AND FC.TCO_CODIGO=2 AND FC.FCL_NUMERO >= " & XN(txtFBDesde.Text)
    'End If
    'If txtFBHasta.Text <> "" Then
    '    sql1 = sql1 & " AND FC.TCO_CODIGO=2 AND FC.FCL_NUMERO <= " & XN(txtFBHasta.Text)
    'End If
    If txtZdesde.Text <> "" Then
        sql1 = sql1 & " AND FC.TCO_CODIGO=2 AND FC.FCL_CIERREZ >= " & XN(txtZdesde.Text)
    End If
    If txtZHasta.Text <> "" Then
        sql1 = sql1 & " AND FC.TCO_CODIGO=2 AND FC.FCL_CIERREZ <= " & XN(txtZHasta.Text)
    End If
    sql1 = sql1 & " ORDER BY FC.FCL_NUMERO, FC.FCL_FECHA"
    
    rec.Open sql1, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Tamanio = rec.RecordCount
        Do While rec.EOF = False
            sql1 = "INSERT INTO TMP_LIBRO_IVA_VENTAS (FECHA,COMPROBANTE,NUMERO,"
            sql1 = sql1 & "CLIENTE,CUIT,INGBRUTOS,TASAVIAL,SUBTOTAL,IVA,TOTIVA,TOTAL,RETENCION,IMPINT)"
            sql1 = sql1 & "VALUES ("
            sql1 = sql1 & XDQ(rec!FCL_FECHA) & ","
            sql1 = sql1 & XS(rec!TCO_ABREVIA) & ","
            sql1 = sql1 & XS(Format(rec!FCL_SUCURSAL, "0000") & "-" & Format(rec!FCL_NUMERO, "00000000")) & ","
            If rec!EST_CODIGO = 2 Then ' anuladas
                sql1 = sql1 & XS(rec!CLI_RAZSOC & " - ERROR") & ","
                sql1 = sql1 & XS(Format(ChkNull(rec!CLI_CUIT), "##-########-#")) & ","
                sql1 = sql1 & "NULL" & ","
                sql1 = sql1 & "0" & ","
                sql1 = sql1 & "0" & ","
                sql1 = sql1 & "0" & ","
                sql1 = sql1 & "0" & ","
                sql1 = sql1 & "0" & ","
                sql1 = sql1 & "0" & "," 'RETENCIONES
                sql1 = sql1 & "0" & ")" 'Impuesto Interno
            Else
                If rec!EST_CODIGO = 5 Then ' anuladas
                    sql1 = sql1 & XS("ANULADA") & ","
                Else
                    sql1 = sql1 & XS(rec!CLI_RAZSOC) & ","
                End If
                sql1 = sql1 & XS(Format(ChkNull(rec!CLI_CUIT), "##-########-#")) & ","
                sql1 = sql1 & "NULL" & ","
                sql1 = sql1 & XN(Chk0(rec!FCL_TASAVIAL)) & ","
                sql1 = sql1 & XN(Chk0(rec!FCL_SUBTOTAL)) & ","
                sql1 = sql1 & XN(rec!FCL_IVA) & ","
                TotIva = (CDbl(Chk0(rec!FCL_SUBTOTAL)) * CDbl(rec!FCL_IVA)) / 100
                sql1 = sql1 & XN(CStr(TotIva)) & ","
                sql1 = sql1 & XN(Chk0(rec!FCL_TOTAL)) & ","
                sql1 = sql1 & "0" & "," 'RETENCIONES
                sql1 = sql1 & XN(rec!FCL_IMPINT) & ")" 'Impuesto Interno
            End If
            DBConn.Execute sql1
            rec.MoveNext
            
            Registro = Registro + 1
            ProgressBar1.Value = Format((Registro * 100) / Tamanio, "0.0")
            lblPor.Caption = Format((Registro * 100) / Tamanio, "0.0") & " %"
        Loop
    End If
    rec.Close
End Sub
Private Sub BUSCO_FACTURAS()
    'Dim SecDesde As Long
    'Dim SecHasta As Long
    Dim sql1 As String
    Dim FacADesde As Long
    Dim FacAHasta As Long
    Dim FacBDesde As Long
    Dim FacBHasta As Long
    TotIva = 0
    'BUSCO FACTURAS POR REMITO fac A---------------------------------
    sql1 = "SELECT FC.FCL_NUMERO, FC.FCL_SUCURSAL, FC.FCL_FECHA, FC.FCL_IVA,"
    sql1 = sql1 & " FC.FCL_SUBTOTAL, FC.FCL_TOTAL,"
    sql1 = sql1 & " FC.EST_CODIGO, C.CLI_CUIT, C.CLI_INGBRU,FCL_TASAVIAL,"
    sql1 = sql1 & " C.CLI_RAZSOC, TC.TCO_ABREVIA,FC.FCL_IMPINT,FCL_CIERREZ"
    sql1 = sql1 & " FROM FACTURA_CLIENTE FC, CLIENTE C,"
    sql1 = sql1 & " TIPO_COMPROBANTE TC"
    sql1 = sql1 & " WHERE"
    sql1 = sql1 & " FC.TCO_CODIGO=TC.TCO_CODIGO"
    sql1 = sql1 & " AND FC.CLI_CODIGO=C.CLI_CODIGO"
    'busco todas las facs
    'sql1 = sql1 & " AND FC.EST_CODIGO = 3" 'ESTADO DEFINITIVO Y ANULADO
    If FechaDesde <> "" Then sql1 = sql1 & " AND FC.FCL_FECHA>=" & XDQ(FechaDesde)
    If FechaHasta <> "" Then sql1 = sql1 & " AND FC.FCL_FECHA<=" & XDQ(FechaHasta)
'    If txtZdesde.Text <> "" Then
'        FacADesde = BUSCO_UFACZANT(txtZdesde.Text, 1)
'        sql1 = sql1 & " AND FC.TCO_CODIGO=1 AND FC.FCL_NUMERO >= " & FacADesde
'    End If
'    If txtZHasta.Text <> "" Then
'        FacADesde = BUSCO_UFACZANT(txtZHasta.Text, 1)
'        sql1 = sql1 & " AND FC.TCO_CODIGO=1 AND FC.FCL_NUMERO <= " & FacAHasta
'    End If
    If txtFADesde.Text <> "" Then
        sql1 = sql1 & " AND FC.TCO_CODIGO=1 AND FC.FCL_NUMERO >= " & XN(txtFADesde.Text)
    End If
    If txtFAHasta.Text <> "" Then
        sql1 = sql1 & " AND FC.TCO_CODIGO=1 AND FC.FCL_NUMERO <= " & XN(txtFAHasta.Text)
    End If
    
    sql1 = sql1 & " UNION "
    sql1 = sql1 & " SELECT FC.FCL_NUMERO, FC.FCL_SUCURSAL, FC.FCL_FECHA, FC.FCL_IVA,"
    sql1 = sql1 & " FC.FCL_SUBTOTAL, FC.FCL_TOTAL,"
    sql1 = sql1 & " FC.EST_CODIGO, C.CLI_CUIT, C.CLI_INGBRU,FCL_TASAVIAL,"
    sql1 = sql1 & " C.CLI_RAZSOC, TC.TCO_ABREVIA,FC.FCL_IMPINT, FCL_CIERREZ"
    sql1 = sql1 & " FROM FACTURA_CLIENTE FC, CLIENTE C,"
    sql1 = sql1 & " TIPO_COMPROBANTE TC"
    sql1 = sql1 & " WHERE"
    sql1 = sql1 & " FC.TCO_CODIGO=TC.TCO_CODIGO"
    sql1 = sql1 & " AND FC.CLI_CODIGO=C.CLI_CODIGO"
    'sql1 = sql1 & " AND FC.EST_CODIGO = 3" 'ESTADO DEFINITIVO Y ANULADO
    If FechaDesde <> "" Then sql1 = sql1 & " AND FC.FCL_FECHA>=" & XDQ(FechaDesde)
    If FechaHasta <> "" Then sql1 = sql1 & " AND FC.FCL_FECHA<=" & XDQ(FechaHasta)
'    If txtZdesde.Text <> "" Then
'        FacBDesde = BUSCO_UFACZANT(txtZdesde.Text, 2)
'        sql1 = sql1 & " AND FC.TCO_CODIGO=2 AND FC.FCL_NUMERO >= " & FacBDesde
'    End If
'    If txtZHasta.Text <> "" Then
'        FacBDesde = BUSCO_UFACZANT(txtZHasta.Text, 2)
'        sql1 = sql1 & " AND FC.TCO_CODIGO=2 AND FC.FCL_NUMERO <= " & FacBHasta
'    End If
    If txtFBDesde.Text <> "" Then
        sql1 = sql1 & " AND FC.TCO_CODIGO=2 AND FC.FCL_NUMERO >= " & XN(txtFBDesde.Text)
    End If
    If txtFBHasta.Text <> "" Then
        sql1 = sql1 & " AND FC.TCO_CODIGO=2 AND FC.FCL_NUMERO <= " & XN(txtFBHasta.Text)
    End If
    sql1 = sql1 & " ORDER BY FC.FCL_NUMERO, FC.FCL_FECHA"
    
    rec.Open sql1, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Tamanio = rec.RecordCount
        Do While rec.EOF = False
            sql1 = "INSERT INTO TMP_LIBRO_IVA_VENTAS (FECHA,COMPROBANTE,NUMERO,"
            sql1 = sql1 & "CLIENTE,CUIT,INGBRUTOS,TASAVIAL,SUBTOTAL,IVA,TOTIVA,TOTAL,RETENCION,IMPINT,CIERREZ)"
            sql1 = sql1 & "VALUES ("
            sql1 = sql1 & XDQ(rec!FCL_FECHA) & ","
            sql1 = sql1 & XS(rec!TCO_ABREVIA) & ","
            sql1 = sql1 & XS(Format(rec!FCL_SUCURSAL, "0000") & "-" & Format(rec!FCL_NUMERO, "00000000")) & ","
            If rec!EST_CODIGO = 2 Then ' anuladas
                sql1 = sql1 & XS(rec!CLI_RAZSOC & " - ERROR") & ","
                sql1 = sql1 & XS(Format(ChkNull(rec!CLI_CUIT), "##-########-#")) & ","
                sql1 = sql1 & "NULL" & ","
                sql1 = sql1 & "0" & ","
                sql1 = sql1 & "0" & ","
                sql1 = sql1 & "0" & ","
                sql1 = sql1 & "0" & ","
                sql1 = sql1 & "0" & ","
                sql1 = sql1 & "0" & "," 'RETENCIONES
                sql1 = sql1 & "0" & "," 'Impuesto Interno
                sql1 = sql1 & XN(Chk0(rec!FCL_CIERREZ)) & ")" 'ZETA
            Else
                If rec!EST_CODIGO = 5 Then ' anuladas
                    sql1 = sql1 & XS("ANULADA") & ","
                Else
                    sql1 = sql1 & XS(rec!CLI_RAZSOC) & ","
                End If
                sql1 = sql1 & XS(Format(ChkNull(rec!CLI_CUIT), "##-########-#")) & ","
                sql1 = sql1 & "NULL" & ","
                sql1 = sql1 & XN(Chk0(rec!FCL_TASAVIAL)) & ","
                sql1 = sql1 & XN(Chk0(rec!FCL_SUBTOTAL)) & ","
                sql1 = sql1 & XN(rec!FCL_IVA) & ","
                TotIva = (CDbl(Chk0(rec!FCL_SUBTOTAL)) * CDbl(rec!FCL_IVA)) / 100
                sql1 = sql1 & XN(CStr(TotIva)) & ","
                sql1 = sql1 & XN(Chk0(rec!FCL_TOTAL)) & ","
                sql1 = sql1 & "0" & "," 'RETENCIONES
                sql1 = sql1 & XN(rec!FCL_IMPINT) & "," 'Impuesto Interno
                sql1 = sql1 & XN(Chk0(rec!FCL_CIERREZ)) & ")" 'ZETA
            End If
            DBConn.Execute sql1
            rec.MoveNext
            
            Registro = Registro + 1
            ProgressBar1.Value = Format((Registro * 100) / Tamanio, "0.0")
            lblPor.Caption = Format((Registro * 100) / Tamanio, "0.0") & " %"
        Loop
    End If
    rec.Close
End Sub


Private Sub BUSCO_NOTA_CREDITO()
    TotIva = 0
    'BUSCO NOTA DE CREDITO------------------------------------
     sql = "SELECT NC.NCC_NUMERO, NC.NCC_SUCURSAL, NC.NCC_FECHA, NC.NCC_IVA,"
     sql = sql & " NC.NCC_SUBTOTAL, NC.NCC_TOTAL,"
     sql = sql & " NC.EST_CODIGO,C.CLI_CUIT,C.CLI_INGBRU,"
     sql = sql & " C.CLI_RAZSOC,TC.TCO_ABREVIA"
     sql = sql & " FROM NOTA_CREDITO_CLIENTE NC,"
     sql = sql & " TIPO_COMPROBANTE TC, CLIENTE C"
     sql = sql & " WHERE"
     sql = sql & " NC.TCO_CODIGO=TC.TCO_CODIGO"
     sql = sql & " AND NC.EST_CODIGO = 3"
     sql = sql & " AND NC.CLI_CODIGO=C.CLI_CODIGO"
     If FechaDesde <> "" Then sql = sql & " AND NC.NCC_FECHA>=" & XDQ(FechaDesde)
     If FechaHasta <> "" Then sql = sql & " AND NC.NCC_FECHA<=" & XDQ(FechaHasta)
 
     sql = sql & " ORDER BY NC.NCC_FECHA"
     
     rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Registro = 0
        Tamanio = rec.RecordCount
        Do While rec.EOF = False
            sql = "INSERT INTO TMP_LIBRO_IVA_VENTAS (FECHA,COMPROBANTE,NUMERO,"
            sql = sql & "CLIENTE,CUIT,INGBRUTOS,SUBTOTAL,IVA,TOTIVA,TOTAL,RETENCION)"
            sql = sql & "VALUES ("
            sql = sql & XDQ(rec!NCC_FECHA) & ","
            sql = sql & XS(rec!TCO_ABREVIA) & ","
            sql = sql & XS(Format(rec!NCC_SUCURSAL, "0000") & "-" & Format(rec!NCC_NUMERO, "00000000")) & ","
            sql = sql & XS(rec!CLI_RAZSOC) & ","
            sql = sql & XS(Format(ChkNull(rec!CLI_CUIT), "##-########-#")) & ","
            sql = sql & "NULL" & ","
            'sql = sql & XS(Format(rec!CLI_INGBRU, "###-#####-##")) & ","
            If rec!EST_CODIGO = 2 Then
                sql = sql & "0" & ","
                sql = sql & XN(rec!NCC_IVA) & ","
                sql = sql & "0" & ","
                sql = sql & "0" & ","
                sql = sql & "0" & ")" 'RETENCION
            Else
                sql = sql & XN(CStr((-1) * CDbl(rec!NCC_SUBTOTAL))) & ","
                sql = sql & XN(rec!NCC_IVA) & ","
                TotIva = (CDbl(rec!NCC_SUBTOTAL) * CDbl(rec!NCC_IVA)) / 100
                sql = sql & XN(CStr((-1) * CDbl(TotIva))) & ","
                sql = sql & XN(CStr((-1) * CDbl(rec!NCC_TOTAL))) & ","
                sql = sql & "0" & ")" 'RETENCION
            End If
            DBConn.Execute sql
            rec.MoveNext
            
            Registro = Registro + 1
            ProgressBar1.Value = Format((Registro * 100) / Tamanio, "0.0")
            lblPor.Caption = Format((Registro * 100) / Tamanio, "0.0") & " %"
        Loop
    End If
    rec.Close
End Sub

Private Sub BUSCO_NOTA_DEBITO()
    TotIva = 0
    'BUSCO NOTA DE DEBITO SERVICIOS, CONCEPTO Y CHEQUES DEVUELTOS-----
    sql = "SELECT ND.NDC_NUMERO, ND.NDC_SUCURSAL, ND.NDC_FECHA, ND.NDC_IVA,"
    sql = sql & " ND.NDC_SUBTOTAL, ND.NDC_TOTAL,"
    sql = sql & " ND.EST_CODIGO, C.CLI_CUIT, C.CLI_INGBRU,"
    sql = sql & " C.CLI_RAZSOC,TC.TCO_ABREVIA"
    sql = sql & " FROM NOTA_DEBITO_CLIENTE ND,"
    sql = sql & " TIPO_COMPROBANTE TC , CLIENTE C"
    sql = sql & " WHERE ND.TCO_CODIGO=TC.TCO_CODIGO"
    sql = sql & " AND ND.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & " AND ND.EST_CODIGO = 3"
    If FechaDesde <> "" Then sql = sql & " AND ND.NDC_FECHA>=" & XDQ(FechaDesde)
    If FechaHasta <> "" Then sql = sql & " AND ND.NDC_FECHA<=" & XDQ(FechaHasta)
    

    sql = sql & " ORDER BY ND.NDC_FECHA"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Registro = 0
        Tamanio = rec.RecordCount
        Do While rec.EOF = False
            sql = "INSERT INTO TMP_LIBRO_IVA_VENTAS (FECHA,COMPROBANTE,NUMERO,"
            sql = sql & "CLIENTE,CUIT,INGBRUTOS,SUBTOTAL,IVA,TOTIVA,TOTAL,RETENCION)"
            sql = sql & "VALUES ("
            sql = sql & XDQ(rec!NDC_FECHA) & ","
            sql = sql & XS(rec!TCO_ABREVIA) & ","
            sql = sql & XS(Format(rec!NDC_SUCURSAL, "0000") & "-" & Format(rec!NDC_NUMERO, "00000000")) & ","
            sql = sql & XS(rec!CLI_RAZSOC) & ","
            sql = sql & XS(Format(ChkNull(rec!CLI_CUIT), "##-########-#")) & ","
            sql = sql & "NULL" & ","
            'sql = sql & XS(Format(rec!CLI_INGBRU, "###-#####-##")) & ","
            If rec!EST_CODIGO = 2 Then
                sql = sql & "0" & ","
                sql = sql & XN(rec!NDC_IVA) & ","
                sql = sql & "0" & ","
                sql = sql & "0" & ","
                sql = sql & "0" & ")" 'RETENCION
            Else
                sql = sql & XN(rec!NDC_SUBTOTAL) & ","
                sql = sql & XN(rec!NDC_IVA) & ","
                TotIva = (CDbl(rec!NDC_SUBTOTAL) * CDbl(rec!NDC_IVA)) / 100
                sql = sql & XN(CStr(TotIva)) & ","
                sql = sql & XN(rec!NDC_TOTAL) & ","
                sql = sql & "0" & ")" 'RETENCION
            End If
            DBConn.Execute sql
            rec.MoveNext
            
            Registro = Registro + 1
            ProgressBar1.Value = Format((Registro * 100) / Tamanio, "0.0")
            lblPor.Caption = Format((Registro * 100) / Tamanio, "0.0") & " %"
        Loop
    End If
    rec.Close
End Sub

Private Sub BUSCO_RETENCIONES()
    TotIva = 0
    'BUSCO NOTA DE CREDITO------------------------------------
     sql = "SELECT DR.DRE_COMNUMERO, DR.DRE_COMSUCURSAL, DR.DRE_COMFECHA,"
     sql = sql & " DR.DRE_COMIMP, R.EST_CODIGO, C.CLI_CUIT,C.CLI_INGBRU,"
     sql = sql & " C.CLI_RAZSOC,TC.TCO_ABREVIA"
     sql = sql & " FROM RECIBO_CLIENTE R,DETALLE_RECIBO_CLIENTE DR"
     sql = sql & ",TIPO_COMPROBANTE TC , CLIENTE C"
     sql = sql & " WHERE"
     sql = sql & " R.TCO_CODIGO=DR.TCO_CODIGO"
     sql = sql & " AND R.REC_NUMERO=DR.REC_NUMERO"
     sql = sql & " AND R.REC_SUCURSAL=DR.REC_SUCURSAL"
     sql = sql & " AND DR.DRE_TCO_CODIGO=TC.TCO_CODIGO"
     sql = sql & " AND DR.DRE_TCO_CODIGO IN (14,15,16)" 'LAS TRES RETENCIONES
     sql = sql & " AND R.CLI_CODIGO=C.CLI_CODIGO"
     sql = sql & " AND R.EST_CODIGO = 3"
     If FechaDesde <> "" Then sql = sql & " AND R.REC_FECHA>=" & XDQ(FechaDesde)
     If FechaHasta <> "" Then sql = sql & " AND R.REC_FECHA<=" & XDQ(FechaHasta)
         
     sql = sql & " ORDER BY R.REC_FECHA"
     
     rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Registro = 0
        Tamanio = rec.RecordCount
        Do While rec.EOF = False
            sql = "INSERT INTO TMP_LIBRO_IVA_VENTAS (FECHA,COMPROBANTE,NUMERO,"
            sql = sql & "CLIENTE,CUIT,INGBRUTOS,SUBTOTAL,IVA,TOTIVA,TOTAL,RETENCION)"
            sql = sql & "VALUES ("
            sql = sql & XDQ(rec!DRE_COMFECHA) & ","
            sql = sql & XS(rec!TCO_ABREVIA) & ","
            sql = sql & XS(Format(rec!DRE_COMSUCURSAL, "0000") & "-" & Format(rec!DRE_COMNUMERO, "00000000")) & ","
            sql = sql & XS(rec!CLI_RAZSOC) & ","
            sql = sql & XS(Format(ChkNull(rec!CLI_CUIT), "##-########-#")) & ","
            sql = sql & "NULL" & ","
            'sql = sql & XS(Format(rec!CLI_INGBRU, "###-#####-##")) & ","
            If rec!EST_CODIGO = 2 Then
                sql = sql & "0" & ","
                sql = sql & "0" & ","
                sql = sql & "0" & ","
                sql = sql & "0" & ","
                sql = sql & "0" & ")" 'RETENCION
            Else
                sql = sql & "0" & ","
                sql = sql & "0" & ","
                sql = sql & "0" & ","
                sql = sql & XN(CStr((-1) * CDbl(rec!DRE_COMIMP))) & "," 'TOTAL
                sql = sql & XN(CStr((-1) * CDbl(rec!DRE_COMIMP))) & ")" 'RETENCION
            End If
            DBConn.Execute sql
            rec.MoveNext
            
            Registro = Registro + 1
            ProgressBar1.Value = Format((Registro * 100) / Tamanio, "0.0")
            lblPor.Caption = Format((Registro * 100) / Tamanio, "0.0") & " %"
        Loop
    End If
    rec.Close
End Sub

Private Sub optCierreZ_Click()
    fracierrez.Enabled = True
    fralibro.Enabled = False
    txtZdesde.SetFocus
End Sub

Private Sub optlibro_Click()
    fralibro.Enabled = True
    fracierrez.Enabled = False
    txtFADesde.SetFocus
End Sub

Private Sub txtFADesde_GotFocus()
    seltxt
End Sub

Private Sub txtFADesde_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtFAHasta_GotFocus()
    seltxt
End Sub

Private Sub txtFAHasta_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtFBHasta_GotFocus()
    seltxt
End Sub

Private Sub txtFBHasta_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtZdesde_GotFocus()
    seltxt
End Sub

Private Sub txtZdesde_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtZHasta_GotFocus()
    seltxt
End Sub

Private Sub txtZHasta_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub xtFBDesde_Change()

End Sub

Private Sub TxtFBDesde_GotFocus()
    seltxt
End Sub

Private Sub TxtFBDesde_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub
