VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRendicionTasaVial 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rendicion Tasa Vial"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   15
      Top             =   2040
      Width           =   5610
      Begin VB.OptionButton optImpresora 
         Caption         =   "Impresora"
         Height          =   195
         Left            =   2370
         TabIndex        =   3
         Top             =   315
         Width           =   1050
      End
      Begin VB.OptionButton optPantalla 
         Caption         =   "Pantalla"
         Height          =   195
         Left            =   945
         TabIndex        =   2
         Top             =   315
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.CommandButton CBImpresora 
         Caption         =   "Configurar Impresora"
         Height          =   375
         Left            =   3810
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   1665
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   150
         TabIndex        =   17
         Top             =   315
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   735
      Left            =   3045
      Picture         =   "frmRendicionTasaVial.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2850
      Width           =   840
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   4755
      Picture         =   "frmRendicionTasaVial.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2850
      Width           =   840
   End
   Begin VB.Frame Frame2 
      Height          =   2040
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   5595
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   225
         Left            =   210
         TabIndex        =   8
         Top             =   1395
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   397
         _Version        =   327682
         Appearance      =   1
      End
      Begin MSComCtl2.DTPicker FechaDesde 
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   56295425
         CurrentDate     =   41098
      End
      Begin MSComCtl2.DTPicker FechaHasta 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   56295425
         CurrentDate     =   41098
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   360
         TabIndex        =   13
         Top             =   525
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   405
         TabIndex        =   12
         Top             =   870
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
         Left            =   2985
         TabIndex        =   11
         Top             =   510
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
         Left            =   2985
         TabIndex        =   10
         Top             =   840
         Width           =   1785
      End
      Begin VB.Label lblPor 
         AutoSize        =   -1  'True
         Caption         =   "100 %"
         Height          =   195
         Left            =   5085
         TabIndex        =   9
         Top             =   1425
         Width           =   435
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmRendicionTasaVial.frx":0614
      Height          =   735
      Left            =   3900
      Picture         =   "frmRendicionTasaVial.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2850
      Width           =   840
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   1590
      Top             =   3030
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   2130
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      TabIndex        =   14
      Top             =   3015
      Width           =   750
   End
End
Attribute VB_Name = "frmRendicionTasaVial"
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

Private Sub cmdAceptar_Click()
     Registro = 0
     Tamanio = 0
     TotIva = 0
     
     If FechaDesde.Value = "" Then
        MsgBox "Debe ingresar el periodo", vbExclamation, TIT_MSGBOX
        FechaDesde.SetFocus
        Exit Sub
     End If
     If FechaHasta.Value = "" Then
        MsgBox "Debe ingresar el periodo", vbExclamation, TIT_MSGBOX
        FechaHasta.SetFocus
        Exit Sub
     End If
     
     On Error GoTo CLAVO
     Screen.MousePointer = vbHourglass
'     DBConn.BeginTrans
     lblEstado.Caption = "Buscando Datos..."
     
    'BORRO LA TABLA TEMPORAL DE IVA VENTAS
    sql = "DELETE FROM TMP_TASAVIAL"
    DBConn.Execute sql
    
    'FACTURAS
    'DBConn.CommitTrans
    BUSCO_FACTURAS

        
    lblEstado.Caption = ""
    
    'cargo el reporte
    
    ListarRendicion
    
        
    Screen.MousePointer = vbNormal
    
    Exit Sub

CLAVO:
 Screen.MousePointer = vbNormal
 lblEstado.Caption = ""
 'DBConn.RollbackTrans
 If rec.State = 1 Then rec.Close
 MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub ListarRendicion()
    lblEstado.Caption = "Buscando Listado..."
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
    
     If FechaDesde.Value <> "" And FechaHasta.Value <> "" Then
        Rep.Formulas(3) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & FechaHasta.Value & "'"
    ElseIf FechaDesde.Value <> "" And FechaHasta.Value = "" Then
        Rep.Formulas(3) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & Date & "'"
    ElseIf FechaDesde.Value = "" And FechaHasta.Value <> "" Then
        Rep.Formulas(3) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & FechaHasta.Value & "'"
    ElseIf FechaDesde.Value = "" And FechaHasta.Value = "" Then
        Rep.Formulas(3) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & Date & "'"
    End If
    
    Rep.WindowTitle = "Rendicion Tasa Vial"
    Rep.ReportFileName = DRIVE & DirReport & "rptRendicionTasaVial.rpt"
    
    If optPantalla.Value = True Then
        Rep.Destination = crptToWindow
    ElseIf optImpresora.Value = True Then
        Rep.Destination = crptToPrinter
    End If
     Rep.Action = 1
     
     lblEstado.Caption = ""
     Rep.Formulas(0) = ""
     Rep.Formulas(1) = ""
     Rep.Formulas(2) = ""
End Sub

Private Sub cmdNuevo_Click()
    FechaDesde.Value = ""
    lblPeriodo1.Caption = ""
    FechaHasta.Value = ""
    lblPeriodo2.Caption = ""
    FechaDesde.SetFocus
End Sub

Private Sub CmdSalir_Click()
    Set frmRendicionTasaVial = Nothing
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    lblEstado.Caption = ""
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

Private Sub BUSCO_FACTURAS()
    TotIva = 0
    'BUSCO FACTURAS POR REMITO ---------------------------------
    sql = "SELECT P.PTO_CODIGO, P.PTO_DESCRI,"
    sql = sql & " DFC.DFC_TasaVial,SUM(DFC.DFC_CANTIDAD) AS CANTI,SUM(DFC.DFC_TotalTVial) AS TotalTVial"
    sql = sql & " FROM FACTURA_CLIENTE FC, DETALLE_FACTURA_CLIENTE DFC, PRODUCTO P"
    sql = sql & " WHERE"
    sql = sql & " FC.TCO_CODIGO=DFC.TCO_CODIGO"
    sql = sql & " AND FC.FCL_NUMERO=DFC.FCL_NUMERO"
    sql = sql & " AND FC.FCL_SUCURSAL=DFC.FCL_SUCURSAL"
    sql = sql & " AND FC.EST_CODIGO = 3 " 'ANULADO
    sql = sql & " AND DFC.PTO_CODIGO = P.PTO_CODIGO"
    sql = sql & " AND P.LNA_CODIGO =1 " 'combustibles
    
    If FechaDesde <> "" Then sql = sql & " AND FC.FCL_FECHA>=" & XDQ(FechaDesde)
    If FechaHasta <> "" Then sql = sql & " AND FC.FCL_FECHA<=" & XDQ(FechaHasta)
    
    sql = sql & " GROUP BY P.PTO_CODIGO, P.PTO_DESCRI, DFC.DFC_TasaVial"
    sql = sql & " ORDER BY P.PTO_CODIGO"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Tamanio = rec.RecordCount
        Do While rec.EOF = False
            sql = "INSERT INTO TMP_TASAVIAL(PTO_CODIGO,PTO_DESCRI,"
            sql = sql & "DFC_TasaVial,DFC_CANTIDAD,DFC_TotalTVial)"
            sql = sql & "VALUES ("
            sql = sql & XN(rec!PTO_CODIGO) & ","
            sql = sql & XS(rec!PTO_DESCRI) & ","
            sql = sql & XN(rec!DFC_TasaVial) & ","
            sql = sql & XN(rec!CANTI) & ","
            sql = sql & XN(rec!TotalTVial) & ")"
            DBConn.Execute sql
            rec.MoveNext
            
            Registro = Registro + 1
            ProgressBar1.Value = Format((Registro * 100) / Tamanio, "0.0")
            lblPor.Caption = Format((Registro * 100) / Tamanio, "0.0") & " %"
        Loop
    End If
    rec.Close
End Sub

