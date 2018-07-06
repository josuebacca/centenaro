Attribute VB_Name = "ReadIni"
Option Explicit
Public DRIVE        As String   'Unidad donde est� mapeado
Public SERVIDOR     As String   'Servidor al cual conectarse
Public BASEDATO     As String   'Base de datos a la que te queres conectar
Public DirReport    As String
Public Impresora    As String   'PARA SABER QUE IMPRESORA USA PARA LA FACTURA Y REMITO
Public DirBkp       As String
Public FISCAL       As String
Public usuario As String
Public DirAFIP As String


Public Sub LeoIni()
Dim Pos     As Integer
Dim Largo   As Integer
Dim ValVar  As String
Dim NomVar  As String
Open "C:\WINDOWS\CENTENARO.INI" For Input As #1
Do While Not EOF(1)
    Line Input #1, ValVar
    Largo = Len(ValVar)
    If Largo > 3 Then
        Pos = IIf(InStr(1, ValVar, "=") = 0, Largo, InStr(1, ValVar, "="))
        NomVar = UCase(Trim(Left(ValVar, Pos - 2)))
        ValVar = Trim(Right(ValVar, Largo - (Pos)))
        Select Case NomVar
           Case "SERVIDOR"
                SERVIDOR = ValVar
        
           Case "BASEDATO"
                BASEDATO = ValVar
        
           Case "DRIVE"
                DRIVE = ValVar
                
           Case "DIR_REPORT"
                DirReport = ValVar
           
           Case "DIR_BKP"
                DirBkp = ValVar
           
           Case "IMPRESORA"
                Impresora = ValVar
           
           Case "DIR_AFIP"
                DirAFIP = ValVar
                
           Case "FISCAL"
                FISCAL = ValVar
             
        End Select
    End If
Loop
Close #1
End Sub
