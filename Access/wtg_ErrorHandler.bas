Attribute VB_Name = "wtg_ErrorHandler"
Option Compare Database
Option Explicit


'-------------------------------------------------------------------------------
' Method    : wtg_ErrorHandler
' Author    : Witigo
' Date      : 09/06/2013
' Version   : 1.0
' Purpose   : Captura los mensajes de error de Ms Access y los muestra ya
'             formateados en una ventana de información/error
'
' @Param    String   strSource
' @Param    Integer  intErrNumber
' @Param    String   strDescription
' @Param    String   strProcedure
' @Param    String   strModuleType
' @Param    String   strModuleName
'-------------------------------------------------------------------------------
Public Function wtg_ErrorHandler( _
                    strSource As String, _
                    intErrNumber As Integer, _
                    strDescription As String, _
                    strProcedure As String, _
                    strModuleType As String, _
                    strModuleName As String _
                    )

Dim strError as string

    ' Filtramos por tipo de error, para dar mas información...
    Select Case intErrNumber

        Case 2102
            strError = intErrNumber & " - Formulario no encontrado."

        Case 2103
            strError = intErrNumber & " - Informe no encontrado."

        Case 3192
            strError = intErrNumber & " - Tabla de datos no existe."

        Case Else
            strError = intErrNumber

    End Select

    ' Mostramos el mensaje de información del error...
    MsgBox "Ha ocurrido un error inesperado en la aplicación." _
        & vbCrLf & "" _
        & vbCrLf & "Fecha : " & Date _
        & vbCrLf & "" _
        & vbCrLf & "Disparador : " & strSource _
        & vbCrLf & "" _
        & vbCrLf & "Error : " & strError _
        & vbCrLf & "" _
        & vbCrLf & "Descripción : " & strDescription _
        & vbCrLf & "" _
        & vbCrLf & "Procedimiento : " & strProcedure _
        & vbCrLf & "" _
        & vbCrLf & "Módulo : " & strModuleType & " " & strModuleName _
        & vbCrLf & "" _
        & vbCrLf & "Si el problema persiste, pongase en contacto con el desarrollador de la aplicación, para buscar una solución al problema.", vbCritical, " Error"

    ' Hacemos log del error
    Call wtg_WriteErrorLog(strSource, strError, strDescription, strProcedure, strModuleType, strModuleName)

End Function
