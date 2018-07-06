VERSION 5.00
Begin VB.Form frmDatosTarjeta 
   BorderStyle     =   0  'None
   Caption         =   "Datos Tarjeta"
   ClientHeight    =   2745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraTarjeta 
      Height          =   2685
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   4095
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
         TabIndex        =   3
         ToolTipText     =   "Alta de Plan"
         Top             =   900
         Width           =   240
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
         TabIndex        =   1
         ToolTipText     =   "Alta de Tarjeta"
         Top             =   510
         Width           =   240
      End
      Begin VB.CommandButton cmdCerrarTarjeta 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   2730
         TabIndex        =   8
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txtTar_Autorizacion 
         Height          =   315
         Left            =   1305
         MaxLength       =   30
         TabIndex        =   6
         Top             =   1965
         Width           =   2505
      End
      Begin VB.ComboBox cboTarjeta 
         Height          =   315
         ItemData        =   "frmDatosTarjeta.frx":0000
         Left            =   1305
         List            =   "frmDatosTarjeta.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   495
         Width           =   2505
      End
      Begin VB.ComboBox cboPlan 
         Height          =   315
         ItemData        =   "frmDatosTarjeta.frx":0004
         Left            =   1305
         List            =   "frmDatosTarjeta.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   885
         Width           =   2505
      End
      Begin VB.TextBox txtCupon 
         Height          =   315
         Left            =   1305
         TabIndex        =   5
         Top             =   1605
         Width           =   2505
      End
      Begin VB.TextBox txtLote 
         Height          =   315
         Left            =   1305
         TabIndex        =   4
         Top             =   1245
         Width           =   2505
      End
      Begin VB.CommandButton cmdAceptoTarjeta 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   1260
         TabIndex        =   7
         Top             =   2280
         Width           =   1425
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Autorización:"
         Height          =   315
         Left            =   45
         TabIndex        =   15
         Top             =   1965
         Width           =   1215
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tarjeta:"
         Height          =   315
         Left            =   45
         TabIndex        =   14
         Top             =   495
         Width           =   1215
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Plan:"
         Height          =   315
         Left            =   45
         TabIndex        =   13
         Top             =   885
         Width           =   1215
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cupón:"
         Height          =   315
         Left            =   45
         TabIndex        =   12
         Top             =   1605
         Width           =   1215
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lote:"
         Height          =   315
         Left            =   45
         TabIndex        =   11
         Top             =   1245
         Width           =   1215
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "Datos Tarjeta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   30
         TabIndex        =   10
         Top             =   120
         Width           =   4005
      End
   End
End
Attribute VB_Name = "frmDatosTarjeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboTarjeta_LostFocus()
    Dim mCodTar As String
    mCodTar = cboTarjeta.ItemData(cboTarjeta.ListIndex)
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


Private Sub cmdAceptoTarjeta_Click()
    If cboPlan.ListIndex = -1 Then
        MsgBox "Falta Ingresar el Plan", vbExclamation, TIT_MSGBOX
        cboPlan.SetFocus
        Exit Sub
    End If
    'txtImportePago.SetFocus
    'fraTarjeta.Visible = False
    Me.Visible = False
End Sub

Private Sub cmdCerrarTarjeta_Click()
    'cboFormaPago.ListIndex = 0
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then MySendKeys Chr(9)
    If KeyAscii = vbKeyEscape Then cmdCerrarTarjeta_Click
End Sub

Private Sub Form_Load()
    Centrar_pantalla Me
    cboPlan.Clear
    cboTarjeta.Clear
    cSQL = "SELECT TAR_CODIGO, TAR_DESCRI"
    cSQL = cSQL & " FROM TARJETA"
    cSQL = cSQL & " WHERE TTA_CODIGO=1" 'SOLO TARJETA DE CREDITO
    cSQL = cSQL & " ORDER BY TAR_DESCRI"
    rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If (rec.BOF And rec.EOF) = 0 Then
       Do While rec.EOF = False
          cboTarjeta.AddItem Trim(rec!TAR_DESCRI)
          cboTarjeta.ItemData(cboTarjeta.NewIndex) = rec!TAR_CODIGO
          rec.MoveNext
       Loop
       If cboTarjeta.ListCount > 0 Then cboTarjeta.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub txtCupon_GotFocus()
    SelecTexto txtCupon
End Sub

Private Sub txtCupon_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtLote_GotFocus()
    SelecTexto txtLote
End Sub

Private Sub txtLote_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtTar_Autorizacion_GotFocus()
    SelecTexto txtTar_Autorizacion
End Sub
