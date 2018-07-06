Attribute VB_Name = "EpsonFiscalInterfacedll"
Option Explicit
Public Error As Long
Const ERROR_NINGUNO As Long = 0
Const EpsonFiscaldll As String = "D:\Sistema\EpsonFiscalInterface.dll"
Public Const ID_TASA_IVA_21_00 As Long = 5



' /*=============================================================================
' **  VISUAL BASIC 6.0                                    Epson Latin America  **
' **=============================================================================
' **  Rutinas de ejemplo en Visual Basic 6.0 para impresora fiscal Epson       **
' **                                                                           **
' **  Para la inmortalidad del creador de estas rutinas, daremos a conocer su  **
' **  nombre:                                                                  **
' **          Autor:  Rubén Pantaleón Miranda, rpm+                            **
' **          Fecha:  16/Junio/2017  - Bs.As. - Argentina.-                    **
' **  Actualizacion:  24/Agosto/2017 - Bs.As. - Argentina.-                    **
' **                                                                           **
' **===========================================================================*/
Public Declare Function EnviarComando Lib "D:\Sistema\EpsonFiscalInterface.dll" (ByVal comando As String) As Long
Public Declare Function ObtenerRespuestaExtendida Lib "D:\Sistema\EpsonFiscalInterface.dll" (ByVal numero_campo As Long, ByVal buffer_salida As Long, ByVal largo_buffer_salida As Long, ByVal largo_final_buffer_salida As Long) As Long

Public Declare Function Cancelar Lib "D:\Sistema\EpsonFiscalInterface.dll" () As Long

Public Declare Function ConsultarVersionDll Lib "D:\Sistema\EpsonFiscalInterface.dll" (ByVal Descripcion As String, ByVal descripcion_largo_maximo As Long, ByVal mayor As Long, ByVal menor As Long) As Long
Public Declare Function ConsultarVersionEquipo Lib "D:\Sistema\EpsonFiscalInterface.dll" (ByVal Descripcion As String, ByVal descripcion_largo_maximo As Long, ByVal mayor As Long, ByVal menor As Long) As Long
Public Declare Function ConsultarFechaHora Lib "D:\Sistema\EpsonFiscalInterface.dll" (ByVal respuesta As String, ByVal descripcion_largo_maximo As Long) As Long
Public Declare Function ConsultarDescripcionDeError Lib "D:\Sistema\EpsonFiscalInterface.dll" (ByVal numero_de_errr As Long, ByVal respuesta_descripcion As String, ByVal respuesta_descripcion_largo_maximo As Long) As Long

Public Declare Function ConsultarEstado Lib "D:\Sistema\EpsonFiscalInterface.dll" (ByVal id_consulta As Long, ByVal respuesta As Long) As Long

Public Declare Function ConsultarNumeroPuntoDeVenta Lib "D:\Sistema\EpsonFiscalInterface.dll" (ByVal respuesta As String, ByVal respuesta_largo_maximo As Long) As Long
Public Declare Function ConsultarNumeroComprobanteUltimo Lib "D:\Sistema\EpsonFiscalInterface.dll" (ByVal tipo_de_comprobante As String, ByVal respuesta As String, ByVal respuesta_largo_maximo As Long) As Long
Public Declare Function ConsultarNumeroComprobanteActual Lib "D:\Sistema\EpsonFiscalInterface.dll" (ByVal respuesta As String, ByVal respuesta_largo_maximo As Long) As Long
Public Declare Function ConsultarTipoComprobanteActual Lib "D:\Sistema\EpsonFiscalInterface.dll" (ByVal respuesta As String, ByVal respuesta_largo_maximo As Long) As Long

Public Declare Function CargarDatosCliente Lib "D:\Sistema\EpsonFiscalInterface.dll" (ByVal nombre_o_razon_social1 As String, ByVal nombre_o_razon_social2 As String, ByVal domicilio1 As String, ByVal domicilio2 As String, ByVal domicilio3 As String, ByVal id_tipo_documento As Long, ByVal numero_documento As String, ByVal id_responsabilidad_iva As Long) As Long
Public Declare Function CargarComprobanteAsociado Lib "D:\Sistema\EpsonFiscalInterface.dll" (ByVal Descripcion As String) As Long
Public Declare Function AbrirComprobante Lib "D:\Sistema\EpsonFiscalInterface.dll" (ByVal id_tipo_documento As Long) As Long
Public Declare Function CargarTextoExtra Lib "D:\Sistema\EpsonFiscalInterface.dll" (ByVal Descripcion As String) As Long
Public Declare Function ImprimirItem Lib "D:\Sistema\EpsonFiscalInterface.dll" (ByVal id_modificador As Long, ByVal Descripcion As String, ByVal cantidad As String, ByVal precio As String, ByVal id_tasa_iva As Long, ByVal ii_id As Long, ByVal ii_valor As String, ByVal id_codigo As Long, ByVal Codigo As String, ByVal codigo_unidad_matrix As String, ByVal codigo_unidad_medida As Long) As Long
Public Declare Function ImprimirTextoLibre Lib "D:\Sistema\EpsonFiscalInterface.dll" (ByVal Descripcion As String) As Long
Public Declare Function CerrarComprobante Lib "D:\Sistema\EpsonFiscalInterface.dll" () As Long
Public Declare Function CargarLogo Lib "D:\Sistema\EpsonFiscalInterface.dll" (ByVal nombre_de_archivo As String) As Long
Public Declare Function EliminarLogo Lib "D:\Sistema\EpsonFiscalInterface.dll" () As Long

Public Declare Function ConfigurarVelocidad Lib "D:\Sistema\EpsonFiscalInterface.dll" (ByVal velocidad As Long) As Long
Public Declare Function ConfigurarPuerto Lib "D:\Sistema\EpsonFiscalInterface.dll" (ByVal puerto As String) As Long
Public Declare Function Conectar Lib "D:\Sistema\EpsonFiscalInterface.dll" () As Long
Public Declare Function ImprimirCierreX Lib "D:\Sistema\EpsonFiscalInterface.dll" () As Long
Public Declare Function ImprimirCierreZ Lib "D:\Sistema\EpsonFiscalInterface.dll" () As Long
Public Declare Function Desconectar Lib "D:\Sistema\EpsonFiscalInterface.dll" () As Long
Const BaudRate As Long = 9600
Const Port As Long = 1
Public str_tipo_comprobante As String
Public Function conectar_impresora() As Long

  Dim Error As Long
  
  ' connect
  ConfigurarVelocidad (BaudRate)
  ConfigurarPuerto (Port)
  Error = Conectar()

  ' retornar valor
  conectar_impresora = Error

End Function

Public Function Epson_ConsultarNumeroComprobanteUltimo(tipo_cbte As Integer) As Long

    ' NOTA:    1º Este ejemplo es util solo para *** TM-T900FA ***, vesion Ceres 1.00.-
    '-------
    '          2º La siguiente version de DLL "D:\Downloads\EpsonFiscalInterface.02.03.02\examples\visual_basic_6\Release\EpsonFiscalInterface.dll" contendrá
    '             una función que resuelve esta consulta para todas la impresoras.
    '
    '          3º El nombre de la funcion es: ConsultarNumeroComprobanteUltimo()
    
                  
                  
    ' constante
    Const str_numero_comprobante_largo_maximo As Long = 60
                  
    ' definicion
    Dim msg
    Dim index
    Dim Error As Long
    Dim str_cmd As String
    
    Dim str_numero_comprobante As String     ' respuesta que buscamos
    Dim str_numero_comprobante_largo_real As Long
    
    ' inicializacion
    str_cmd = ""
    
    Select Case tipo_cbte
        Case 1 'fac-a
            str_tipo_comprobante = "081"
        Case 2 'fac-b
            str_tipo_comprobante = "082"
        'Case "NOTACREDITO-A"
        'Case "NOTACREDITO-A"
        
    End Select
        
    str_numero_comprobante = ""
    str_numero_comprobante_largo_real = 0
    
    ' connect
    Error = conectar_impresora()
    
    
    ' enviar comando de consulta
    str_cmd = "0830|0000|" & str_tipo_comprobante
    Error = EnviarComando(str_cmd)
    If Error <> ERROR_NINGUNO Then
      msg = MsgBox(Error, vbOKOnly, "Error: EnviarComando()")
    End If
    
    ' leer campo de respuesta número 5 (cinco) el cual contine el numero del ultimo comprobante
    Dim buffer(str_numero_comprobante_largo_maximo) As Byte
    Error = ObtenerRespuestaExtendida(5, VarPtr(buffer(1)), str_numero_comprobante_largo_maximo, VarPtr(str_numero_comprobante_largo_real))
    If Error <> ERROR_NINGUNO Then
      msg = MsgBox(Error, vbOKOnly, "Error: ObtenerRespuestaExtendida()")
    End If
    
    ' construyendo en buffer a un string
    str_numero_comprobante = ""
    For index = 1 To str_numero_comprobante_largo_real
        str_numero_comprobante = str_numero_comprobante + Chr$(buffer(index))
    Next
    
    ' mostrar número de comprobante formateado como string
    'msg = MsgBox(str_numero_comprobante, vbOKOnly, "Número de comprobante TM-T900FA")
    
    
    ' close port
    Error = Desconectar()

    Epson_ConsultarNumeroComprobanteUltimo = str_numero_comprobante
    
End Function

