VERSION 5.00
Object = "{AFD24A52-2823-4FBD-B75D-C282C11E1D98}#1.0#0"; "IFEpson.ocx"
Begin VB.Form frmCierreZ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cierre Z"
   ClientHeight    =   2760
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   675
      Left            =   3210
      TabIndex        =   2
      Top             =   1905
      Width           =   1095
   End
   Begin EPSON_Impresora_Fiscal.PrinterFiscal pf 
      Left            =   1800
      Top             =   1305
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Realizar Cierre"
      Height          =   675
      Left            =   420
      TabIndex        =   0
      Top             =   1905
      Width           =   2745
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Este Proceso Realiza el Cierre Fiscal Z y Guarda los Valores obtenidos de la Impresora Fiscal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   330
      TabIndex        =   1
      Top             =   345
      Width           =   4005
   End
End
Attribute VB_Name = "frmCierreZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCerrar_Click()
    Dim mSec As String
    Dim mZZ_Numero As String
    Dim mZZ_Fecha As String
    Dim mZZ_Tickets As String
    Dim mZZ_Facturas As String
    Dim mZZ_UltDoc As String
    Dim mZZ_TotalVta As String
    Dim mZZ_TotalIva As String
    Dim mZZ_UltFac As String
    Dim mZZ_sucursal As String
    
    Dim mZZ_TotalImpuesto As String
    
    If MsgBox("¿Esta Seguro que desea realizar el Cierre Z ?", 36, "Cierre") = 7 Then
        Exit Sub
    End If
    If FISCAL = "TMT900FA" Then
        Error = conectar_impresora()
        If Error = 0 Then
            Error = ImprimirCierreZ()
            MsgBox "El Cierre Z se ha realizado Exitosamente!", vbInformation, TIT_MSGBOX
            Screen.MousePointer = vbNormal
            Error = Desconectar
        End If
    Else
        pf.CloseJournal "Z"
        
        '*numero de cierre Z
        mZZ_Numero = Val(pf.AnswerField_3)
        '*fecha
        mZZ_Fecha = Date
        '*Cantidad de Comprobantes Fiscales Ticket - B - C
        mZZ_Tickets = CDbl(pf.AnswerField_7)
        '*Cantidad de Comprobantes A
        mZZ_Facturas = CDbl(pf.AnswerField_8)
        '*Nro del Ultimo Comprobantes Fiscales Ticket - B - C
        mZZ_UltDoc = CDbl(pf.AnswerField_9)
        '*Monto Total facturado
        mZZ_TotalVta = CDbl(pf.AnswerField_10) / 100
        '*Monto Total de Iva Cobrado
        mZZ_TotalIva = CDbl(pf.AnswerField_11) / 100
        
        '*Monto Total de IMPUESTOS
        'mZZ_TotalImpuesto = CDbl(pf.AnswerField_12) / 100
        
        '*Nro del Ultimo Comprobantes A
        mZZ_UltFac = CDbl(pf.AnswerField_13)
        '*nro de pto de vta
        pf.Status ("C")
        mZZ_sucursal = CDbl(pf.AnswerField_4)
        
        sql = "SELECT max(Z_SECUENCIA) AS maxi from CIERREZ"
        If rec.State = 1 Then rec.Close
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            mSec = Chk0(rec!maxi) + 1
        Else
            mSec = 1
        End If
        rec.Close
    
         DBConn.BeginTrans
         On Error GoTo ErrorTran
         sql = "insert into CIERREZ (Z_SECUENCIA,Z_SUCURSAL,Z_NUMERO,Z_FECHA,"
         sql = sql & "Z_CANT_TICKETS,Z_CANT_FACTURAS,Z_TOTAL,Z_IVA,"
         sql = sql & "Z_ULTIMO_TICKET,Z_ULTIMA_FACTURA) VALUES ("
         sql = sql & XN(mSec) & ","
         sql = sql & XN(mZZ_sucursal) & ","
         sql = sql & XN(mZZ_Numero) & ","
         sql = sql & XDQ(mZZ_Fecha) & ","
         sql = sql & XN(mZZ_Tickets) & ","
         sql = sql & XN(mZZ_Facturas) & ","
         sql = sql & XN(mZZ_TotalVta) & ","
         sql = sql & XN(mZZ_TotalIva) & ","
         sql = sql & XN(mZZ_UltDoc) & ","
         sql = sql & XN(mZZ_UltFac) & ")"
         DBConn.Execute sql
        DBConn.CommitTrans
         
         MsgBox "El Cierre Z se ha realizado Exitosamente!", vbInformation, TIT_MSGBOX
         Screen.MousePointer = vbNormal
         Unload Me
         Exit Sub
    End If
    Unload Me
    Exit Sub
ErrorTran:
    MsgBox "Error en la Transacción" & Chr(13) & Err.Description, 16, AppName
    Screen.MousePointer = vbNormal
    DBConn.RollbackTrans
End Sub

Private Sub CmdSalir_Click()
    Unload Me
    Set frmCierreZ = Nothing
End Sub
