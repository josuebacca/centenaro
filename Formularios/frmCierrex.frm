VERSION 5.00
Object = "{AFD24A52-2823-4FBD-B75D-C282C11E1D98}#1.0#0"; "IFEpson.ocx"
Begin VB.Form frmCierreX 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cierre X"
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
   Begin VB.ComboBox cboTurno 
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1140
      Width           =   2115
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   675
      Left            =   3210
      TabIndex        =   3
      Top             =   1860
      Width           =   1095
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Realizar Cierre"
      Height          =   675
      Left            =   420
      TabIndex        =   1
      Top             =   1860
      Width           =   2745
   End
   Begin EPSON_Impresora_Fiscal.PrinterFiscal pf 
      Left            =   120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Seleccione el Turno"
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
      Left            =   480
      TabIndex        =   4
      Top             =   1200
      Width           =   1635
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Este Proceso Realiza el Cierre Fiscal X"
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
      Left            =   375
      TabIndex        =   2
      Top             =   465
      Width           =   4050
   End
End
Attribute VB_Name = "frmCierreX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCerrar_Click()
    If MsgBox("¿Esta Seguro que desea realizar el Cierre X ?", 36, "Cierre") = 7 Then
        Exit Sub
    End If
    
    If cboTurno.ListIndex = 0 Then
        MsgBox "Debe seleccionar el Turno a cerrar!", vbExclamation, TIT_MSGBOX
        cboTurno.SetFocus
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    ActualizoHora cboTurno.ItemData(cboTurno.ListIndex)
    Error = 0
    If FISCAL = "TMT900FA" Then
          Error = conectar_impresora()
          If Error = 0 Then
            Error = ImprimirCierreX()
            ' close port
            Error = Desconectar()
          End If
    Else
        pf.CloseJournal "X", "P"
    End If
    If Error = 0 Then
        MsgBox "El Cierre X se ha realizado Exitosamente!", vbInformation, TIT_MSGBOX
    Else
        MsgBox "El Cierre X NO pudo generarse", vbExclamation, TIT_MSGBOX
    End If
    Screen.MousePointer = vbNormal
    'Actualizo hora de turno
    'ActualizoTurnos
    
    Unload Me
    Exit Sub

ErrorTran:
    MsgBox "Error en la Transacción" & Chr(13) & Err.Description, 16, AppName
    Screen.MousePointer = vbNormal
    DBConn.RollbackTrans
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
Private Sub CmdSalir_Click()
    Unload Me
    Set frmCierreZ = Nothing
End Sub

Private Sub Form_Load()
    LlenarComboTurnos

End Sub
Private Sub LlenarComboTurnos()
    sql = "SELECT * FROM TURNOS"
    sql = sql & " ORDER BY TUR_CODIGO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboTurno.AddItem ""
        Do While rec.EOF = False
            cboTurno.AddItem rec!TUR_CODIGO & " - " & rec!TUR_DESCRI
            cboTurno.ItemData(cboTurno.NewIndex) = rec!TUR_CODIGO
            rec.MoveNext
        Loop
    End If
    rec.Close
    cboTurno.ListIndex = 0

'    Dim vDesde(3) As Date
'    Dim vHasta(3) As Date
'    Dim i As Integer
'    sql = "SELECT * FROM TURNOS"
'    sql = sql & " ORDER BY TUR_CODIGO"
'    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    i = 0
'    If rec.EOF = False Then
'        Do While rec.EOF = False
'            vDesde(i) = rec!TUR_DESDE
'            vHasta(i) = rec!TUR_HASTA
'            i = i + 1
'            rec.MoveNext
'        Loop
'    End If
'    rec.Close
'    'POSICIONO EL TURNO DE ACUERDO A LA HORA ACTUAL
'    If Time() >= vDesde(0) And Time() <= vHasta(0) Then
'        Call BuscaCodigoProxItemData(1, cboTurno)
'
'    Else
'        If Time() >= vDesde(1) And Time() <= vHasta(1) Then
'            Call BuscaCodigoProxItemData(2, cboTurno)
'
'        Else
'            Call BuscaCodigoProxItemData(3, cboTurno)
'
'        End If
'    End If
End Sub
