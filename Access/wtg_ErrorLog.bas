Attribute VB_Name = "wtg_ErrorLog"
Option Compare Database
Option Explicit


' Nombre por defecto para la tabla que almacenará el
' LOG de los errores de la aplicación.
Const cstrErrorTableName As String = "dbo_ErrorLog"


'-------------------------------------------------------------------------------
' Procedure : wtg_WriteErrorLog
' Author    : Witigo
' Date      : 09/06/2013
' Version   : 1.0
' Purpose   : Escribe un registro de log en una tabla de la aplicación.
'
' @Param    String   strSource
' @Param    String   strErrNumber
' @Param    String   strDescription
' @Param    String   strProcedure
' @Param    String   strModuleType
' @Param    String   strModuleName
'-------------------------------------------------------------------------------
Public Function wtg_WriteErrorLog( _
                    strSource As String, _
                    strErrNumber As String, _
                    strDescription As String, _
                    strProcedure As String, _
                    strModuleType As String, _
                    strModuleName As String _
                    )

Dim dbs As DAO.Database
Dim rst As DAO.Recordset

    Set dbs = CurrentDb()
    Set rst = dbs.OpenRecordset(cstrErrorTableName, dbOpenDynaset)

Dim strAppUser as String
Dim strWinUser As String
Dim strTerminal As String

    ' Si no hay usuario autenticado, establecemos por defecto "Desconocido"
    strAppUser = "Desconocido"
    ' Obtenemos el nombre de sesión en windows
    strWinUser = Environ("UserName")
    ' Obtenemos el nombre de la computadora
    strTerminal = Environ("ComputerName")

    With rst

        ' Agregamos un nuevo registro
        .AddNew

        ' Agregamos la información del registro
        .Fields("ErrDate").Value = Now()
        .Fields("ErrSource").Value = Nz(strSource, "")
        .Fields("ErrNumber").Value = Nz(strErrNumber, "")
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



'---------------------------------------------------------------------------------------
' Procedure : ErrorLog
' Author    : Witigo
' Date      : 09/06/2013
' Version   : 1.0
' Purpose   : Hacer Log de los errores de la aplicación.
'---------------------------------------------------------------------------------------
Public Function wtg_WriteErrorLogFile( _
                    strSource As String, _
                    strErrNumber As String, _
                    strDescription As String, _
                    strProcedure As String, _
                    strModuleType As String, _
                    strModuleName As String _
                    )

Dim strFichero As String

    strFichero = CurrentProject.Path & "\_err.log"

Dim strDBUser

    If Len(Trim(gvarSesion.strUsuario)) = 0 Then
        ' Si no hay usuario autenticado, establecemos por defecto "Desconocido"
        strDBUser = "Desconocido"
    Else
        ' Obtenemos el nombre del usuario de la aplicación
        strDBUser = gvarSesion.strUsuario
    End If

    ' Abrimos el fichero para introducir datos
    Open strFichero For Append Lock Write As #1

    '-- Write some basic info then close the file
    Print #1, "================================================================"
    Print #1, "   ERROR  (" & GetOSComputerName() & "\" & GetOSUserName() & ")  -   " & Date & " " & Time
    Print #1, "================================================================"
    Print #1, "   Número de error          : " & NumError
    Print #1, "   Origen del error         : " & Origen
    Print #1, "   Tipo de procedimiento    : " & TipoModulo & " >>> " & Modulo
    Print #1, "   Nombre del procedimiento : " & Procedimiento
    Print #1, "   Descripción del error    : " & Descripcion
    Print #1, ""

    ' Cerramos el fichero
    Close #1

End Function

'-------------------------------------------------------------------------------
' Procedure : wtg_CreateErrorLog_Table
' Author    : Witigo
' Date      : 09/06/2013
' Version   : 1.0
' Purpose   : Escribe un registro de log en una tabla de la aplicación.
'
' @Param    String   strSource
' @Param    String   strErrNumber
' @Param    String   strDescription
' @Param    String   strProcedure
' @Param    String   strModuleType
' @Param    String   strModuleName
'-------------------------------------------------------------------------------
Public Function wtg_CreateErrorLog_Table()

Dim dbs As dao.Database
Dim tbl As dao.TableDef
Dim fld As Field

    Set dbs = CurrentDb
    Set tbl = dbs.CreateTableDef(cstrErrorTableName)

    Set fld = tbl.CreateField("ErrDate", dbText, 50)
    tbl.Fields.Append fld

    Set fld = tbl.CreateField("ErrSource", dbText, 50)
    tbl.Fields.Append fld

    Set fld = tbl.CreateField("ErrNumber", dbInteger)
    tbl.Fields.Append fld

    Set fld = tbl.CreateField("ErrDescription", dbText, 255)
    tbl.Fields.Append fld

    Set fld = tbl.CreateField("ProcedureType", dbText, 50)
    tbl.Fields.Append fld

    Set fld = tbl.CreateField("ProcedureName", dbText, 50)
    tbl.Fields.Append fld

    Set fld = tbl.CreateField("AppUser", dbText, 50)
    tbl.Fields.Append fld

    Set fld = tbl.CreateField("WinUser", dbText, 50)
    tbl.Fields.Append fld

    Set fld = tbl.CreateField("Terminal", dbText, 50)
    tbl.Fields.Append fld

    dbs.TableDefs.Append tbl
    dbs.TableDefs.Refresh

End Function
