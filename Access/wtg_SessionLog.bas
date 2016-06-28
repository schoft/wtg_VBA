Attribute VB_Name = "wtg_SessionLog"
Option Compare Database
Option Explicit


' Nombre por defecto para la tabla que almacenará el
' LOG de las sesiones de los usuarios de la aplicación.
Const cstrSessionTableName As String = "dbo_SessionLog"


'-------------------------------------------------------------------------------
' Method    : wtg_WriteSesionLog
' Author    : Witigo (Angel Campos Muñoz)
' Date      : 22/09/2014
' Version   : 1.0
' Purpose   : Inserta un registro en la tabla de log de sesiones con los
'             datos de la sesión de usuario
'
' @Param    Integer   intUser
' @Param    String    strResult
' @Param    String    strPassword
'-------------------------------------------------------------------------------
Public Function wtg_WriteSesionLog( _
                    intUser As Integer, _
                    strResult As String, _
                    Optional strPassword As String = "" _
                    )

Dim dbs As DAO.Database
Dim rst As DAO.Recordset

    Set dbs = CurrentDb()
    Set rst = dbs.OpenRecordset(cstrSessionTableName)

    With rst

        ' Agregamos un registro
        .AddNew

        ' Especificamos los campos en los que vamos a insertar
        ' los datos relativos a la sesión de usuario.
        .Fields("SessionDate").Value = Now()
        .Fields("Terminal").Value = Environ("Computername")
        .Fields("User_ID").Value = intUser
        .Fields("Result").Value = strResult
        .Fields("ErrPassword").Value = strPassword

        ' Actualizamos el registro
        .Update

    End With

    ' Cerramos el recordset
    rst.Close

    ' Borramos los objetos
    Set rst = Nothing
    Set dbs = Nothing

End Function


'-------------------------------------------------------------------------------
' Method    : wtg_CreateSessionLog_Table
' Author    : Witigo
' Date      : 09/06/2013
' Version   : 1.0
' Purpose   : Crea la tabla para almacenar los registros del log de sesiones de
'             usuario de la aplicación
'-------------------------------------------------------------------------------
Public Function wtg_CreateSessionLog_Table()

Dim dbs As dao.Database
Dim tbl As dao.TableDef
Dim fld As Field

    Set dbs = CurrentDb
    Set tbl = dbs.CreateTableDef(cstrSessionTableName)

    Set fld = tbl.CreateField("SessionDate", dbDate)
    tbl.Fields.Append fld

    Set fld = tbl.CreateField("Terminal", dbText, 50)
    tbl.Fields.Append fld

    Set fld = tbl.CreateField("User_ID", dbInteger)
    tbl.Fields.Append fld

    Set fld = tbl.CreateField("Result", dbText, 100)
    tbl.Fields.Append fld

    Set fld = tbl.CreateField("ErrPassword", dbText, 50)
    tbl.Fields.Append fld

    dbs.TableDefs.Append tbl
    dbs.TableDefs.Refresh

End Function
