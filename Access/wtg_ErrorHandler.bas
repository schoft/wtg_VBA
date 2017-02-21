Attribute VB_Name = "bas_ErrorHandler"
'---------------------------------------------------------------------------------------
' File     : bas_ErrorLogHandler
' Date     : 16/02/2017
' Version  : 1.0.1
' Author   : Witigo
'
' Purpose  : Código VBA necesario para manejar los errores de la aplicación y grabar la
'            información del error en una tabla de la base de datos, en un fichero de
'            log o mostrar una ventana de error al usuario...
'
' Requires : Requiere del módulo bas_AppGlobals para obtener el valor de las constantes
'            gcstrErrorLogTableName, gcstrLogsDir, gvarUserSession,
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit


' Enumeración para controlar donde se almacenará el log de error
Public Enum eLogDestination
    LogAll = 0
    LogToFile = 1
    LogToTable = 2
    LogToUser = 3
End Enum


'-------------------------------------------------------------------------------
' Procedure : wtg_CreateErrorLog_Table
' Date      : 09/06/2013
' Version   : 1.0.1
' Author    : Witigo
'
' Purpose   : Crea una tabla en la base de datos para almacenar los logs de
'             error de la aplicación.
'-------------------------------------------------------------------------------
Public Function wtg_CreateErrorLog_Table()

Dim dbs As dao.Database
Dim tbl As dao.TableDef
Dim fld As Field

    ' Comprobamos si existe la tabla
    If Not wtg_CheckIfTableExists(gcstrErrorLogTableName) Then

        ' Creamos la tabla en la base de datos local
        Set dbs = CurrentDb
        Set tbl = dbs.CreateTableDef(gcstrErrorLogTableName)

        ' Creamos los campos de la tabla de log de errores
        Set fld = tbl.CreateField("ErrDate", dbDate)
        tbl.Fields.Append fld

        Set fld = tbl.CreateField("ErrSource", dbText, 50)
        fld.AllowZeroLength = True
        tbl.Fields.Append fld

        Set fld = tbl.CreateField("ErrNumber", dbInteger)
        tbl.Fields.Append fld

        Set fld = tbl.CreateField("ErrDescription", dbText, 255)
        fld.AllowZeroLength = True
        tbl.Fields.Append fld

        Set fld = tbl.CreateField("ProcedureType", dbText, 50)
        fld.AllowZeroLength = True
        tbl.Fields.Append fld

        Set fld = tbl.CreateField("ProcedureName", dbText, 50)
        fld.AllowZeroLength = True
        tbl.Fields.Append fld

        Set fld = tbl.CreateField("AppUser", dbText, 50)
        tbl.Fields.Append fld

        Set fld = tbl.CreateField("WinUser", dbText, 50)
        tbl.Fields.Append fld

        Set fld = tbl.CreateField("Terminal", dbText, 50)
        tbl.Fields.Append fld

        dbs.TableDefs.Append tbl
        dbs.TableDefs.Refresh

    End If

End Function


'------------------------------------------------------------------------------
' Method  : wtg_ErrorLog
' Date    : 16/02/2017
' Version : 1.0.1
' Author  : Witigo
' Company : Proteksa Networks S.L.
'
' Purpose : Escribe un registro de LOG de error en la aplicación
'
' @Param    String    strSource           Origen del error
' @Param    String    intErrNumber        Número de error
' @Param    String    strDescription      Descripción del error
' @Param    String    strProcedureName    Nombre del procedimiento
' @Param    String    strProcedureType    Tipo de procedimiento
' @Param    String    strModuleName       Nombre del módulo
'------------------------------------------------------------------------------
Public Function wtg_ErrorLog( _
                    strSource As String, _
                    intErrNumber As Integer, _
                    strDescription As String, _
                    strProcedureName As String, _
                    strProcedureType As String, _
                    strModuleName As String _
                    )

    ' Manejamos el error
    Call wtg_ErrorHandler(strSource, intErrNumber, strDescription, strProcedureName, strProcedureType, strModuleName, LogToUser)

End Function


'-------------------------------------------------------------------------------
' Method  : wtg_ErrorHandler
' Date    : 16/02/2017
' Version : 1.0.1
' Author  : Witigo
' Company : Proteksa Networks S.L.
'
' Purpose : Captura el evento de error y lo redirige a su destino
'
' @Param    String    strSource
' @Param    String    intErrNumber
' @Param    String    strDescription
' @Param    String    strProcedureName
' @Param    String    strProcedureType
' @Param    String    strModuleName
' @Param    Enum      eLogDestination
'-------------------------------------------------------------------------------
Private Sub wtg_ErrorHandler( _
                    strSource As String, _
                    intErrNumber As Integer, _
                    strDescription As String, _
                    strProcedureName As String, _
                    strProcedureType As String, _
                    strModuleName As String, _
                    LogDestination As eLogDestination _
                    )

    Select Case LogDestination
    
        Case 0  ' LogAll
            
            ' Grabamos el log de erro al fichero
            Call wtg_WriteErrorLogFile(strSource, intErrNumber, strDescription, strProcedureName, strProcedureType, strModuleName)
        
            ' Grabamos el log de error a la tabla
            Call wtg_WriteErrorLogTable(strSource, intErrNumber, strDescription, strProcedureName, strProcedureType, strModuleName)
            
            ' Mostramos mensaje de error al usuario
            Call wtg_ShowErrorLog(strSource, intErrNumber, strDescription, strProcedureName, strProcedureType, strModuleName)

        Case 1  ' LogToFile
        
            ' Grabamos el log de error en un fichero
            Call wtg_WriteErrorLogFile(strSource, intErrNumber, strDescription, strProcedureName, strProcedureType, strModuleName)

        Case 2  ' LogToTable
        
            ' Grabamos el log de error en una tabla de la base de datos
            Call wtg_WriteErrorLogTable(strSource, intErrNumber, strDescription, strProcedureName, strProcedureType, strModuleName)

        Case 3  ' LogToUser
        
            ' Mostramos mensaje de error al usuario
            Call wtg_ShowErrorLog(strSource, intErrNumber, strDescription, strProcedureName, strProcedureType, strModuleName)
            
            ' Grabamos el log de error al fichero de logs de error
            Call wtg_WriteErrorLogFile(strSource, intErrNumber, strDescription, strProcedureName, strProcedureType, strModuleName)
        
    End Select
    
End Sub


'-------------------------------------------------------------------------------
' Method    : wtg_ShowErrorLog
' Date      : 09/06/2013
' Version   : 1.0.1
' Author    : Witigo
'
' Purpose   : Captura los mensajes de error de Ms Access y los muestra ya
'             formateados en una ventana de información/error
'
' @Param    String    strSource
' @Param    String    intErrNumber
' @Param    String    strDescription
' @Param    String    strProcedureName
' @Param    String    strProcedureType
' @Param    String    strModuleName
'-------------------------------------------------------------------------------
Public Function wtg_ShowErrorLog( _
                    strSource As String, _
                    intErrNumber As Integer, _
                    strDescription As String, _
                    strProcedureName As String, _
                    strProcedureType As String, _
                    strModuleName As String _
                    )

Dim strError As String

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
        & vbCrLf & "Módulo : " & strModuleName _
        & vbCrLf & "" _
        & vbCrLf & strProcedureType & " : " & strProcedureName _
        & vbCrLf & "" _
        & vbCrLf & "Error : " & CStr(intErrNumber) _
        & vbCrLf & "" _
        & vbCrLf & "Descripción : " & strDescription _
        & vbCrLf & "" _
        & vbCrLf & "Si el problema persiste, pongase en contacto con el desarrollador de la aplicación, para buscar una solución al problema.", vbCritical, " Error"

End Function


'-------------------------------------------------------------------------------
' Procedure : wtg_WriteErrorLogFile
' Date      : 09/06/2013
' Version   : 1.0.1
' Author    : Witigo
'
' Purpose   : Escribe en un fichero de log los errores de la aplicación.
'
' @Param    String    strSource
' @Param    String    intErrNumber
' @Param    String    strDescription
' @Param    String    strProcedureName
' @Param    String    strProcedureType
' @Param    String    strModuleName
'-------------------------------------------------------------------------------
Private Function wtg_WriteErrorLogFile( _
                    strSource As String, _
                    intErrNumber As Integer, _
                    strDescription As String, _
                    strProcedureName As String, _
                    strProcedureType As String, _
                    strModuleName As String _
                    )

Dim strFullPath As String
Dim strFileName As String

    strFileName = "ErrorLog_" & Format(Date, "dd-mm-yyyy") & ".log"
    
    ' gcstrLogsDir = Variable Global ....
    strFullPath = gcstrLogsDir & "\" & strFileName

    ' Abrimos el fichero para introducir datos
    Open strFullPath For Append Lock Write As #1

    '-- Write some basic info then close the file

    Print #1, Date & " " & Time & " - ERROR " & intErrNumber & " (" & wtg_GetOSComputerName() & "\" & wtg_GetOSUserName() & ")"
    Print #1, ""
    Print #1, "   Origen        : " & strSource
    Print #1, "   Modulo        : " & strModuleName
    Print #1, "   Procedimiento : " & strProcedureType & " " & strProcedureName & "()"
    Print #1, ""
    Print #1, "   " & strDescription
    Print #1, ""
    Print #1, ""

    ' Cerramos el fichero
    Close #1
    
End Function


'-------------------------------------------------------------------------------
' Procedure : wtg_WriteErrorLogTable
' Date      : 09/06/2013
' Version   : 1.0.1
' Author    : Witigo
'
' Purpose   : Escribe un registro de log en una tabla de la aplicación.
'
' @Param    String    strSource
' @Param    String    intErrNumber
' @Param    String    strDescription
' @Param    String    strProcedureName
' @Param    String    strProcedureType
' @Param    String    strModuleName
'-------------------------------------------------------------------------------
Private Function wtg_WriteErrorLogTable( _
                    strSource As String, _
                    intErrNumber As Integer, _
                    strDescription As String, _
                    strProcedureName As String, _
                    strProcedureType As String, _
                    strModuleName As String _
                    )

Dim dbs As dao.Database
Dim rst As dao.Recordset
    
Dim strAppUser As String
Dim strWinUser As String
Dim strTerminal As String
    
    ' Comprobamos si existe la tabla
    If Not wtg_CheckIfTableExists(gcstrErrorLogTableName) Then
    
        ' Creamos la tabla de errores
        Call wtg_CreateErrorLog_Table
    
    End If
    
    Set dbs = CurrentDb()
    Set rst = dbs.OpenRecordset(gcstrErrorLogTableName, dbOpenDynaset)

    ' Si no hay usuario autenticado, establecemos por defecto "Desconocido"
    strAppUser = wtg_GetAppUserName
    
    ' Obtenemos el nombre de sesión en windows
    strWinUser = wtg_GetOSUserName
    ' Obtenemos el nombre de la computadora
    strTerminal = wtg_GetOSComputerName
    
    With rst
    
        ' Agregamos un nuevo registro
        .AddNew
    
        ' Agregamos la información del registro
        .Fields("ErrDate").Value = Now()
        .Fields("ErrSource").Value = Nz(strSource, "")
        .Fields("ErrNumber").Value = Nz(intErrNumber, 0)
        .Fields("ErrDescription").Value = Nz(strDescription, "")
        .Fields("ProcedureType").Value = Nz(strProcedureType, "")
        .Fields("ProcedureName").Value = Nz(strProcedureName, "")
        .Fields("AppUser").Value = Nz(strAppUser, "")
        .Fields("WinUser").Value = Nz(strWinUser, "")
        .Fields("Terminal").Value = Nz(strTerminal, "")
    
        ' Guardamos los cambios
        .Update
    
    End With
    
    ' Cerramos el recordset
    rst.Close
    
    ' Borramos los objetos
    Set rst = Nothing
    Set dbs = Nothing

End Function