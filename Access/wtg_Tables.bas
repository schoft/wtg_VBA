Attribute VB_Name = "bas_Tables"
Option Compare Database
Option Explicit


'--------------------------------------------------------------------------------
' Method  : CheckRegistros
' Author  : witigo
' Date    : 03/11/2016
' Version : 1.0
' Purpose : Comprueba si el cliente tiene registros asociados en la tabla
'           dbo_clientes_equipos
'
' @Param    String    strTableName
' @Param    String    strFieldName
' @Param    Long      lngRegistry    Optional
' @Return   Boolean
'--------------------------------------------------------------------------------
Public Function wtg_CheckRegistro( _
					strTableName As String, _
                    lngRegistry As Long, _
                    Optional strFieldName As String = "id" _
                    ) As Boolean

Dim dbs As DAO.Database
Dim rst As DAO.Recordset

Dim strSQL As String

    strSQL = "SELECT " & strTableName & ".* " & _
                "FROM " & strTableName & " " & _
                "WHERE (((" & strTableName & "." & strFieldName & ")= " & lngRegistry & "));"

    Set dbs = CurrentDb()
    Set rst = dbs.OpenRecordset(strSQL, dbOpenDynaset)

    ' Comprobamos si hay registros asociados
    If rst.RecordCount <> 0 Then
        
        wtg_CheckRegistro = True
    
    Else
    
        wtg_CheckRegistro = False
        
    End If

    ' Cerramos el recordset
    rst.Close
    ' Borramos los objetos
    Set rst = Nothing
    Set dbs = Nothing

End Function


'--------------------------------------------------------------------------------
' Method  : EliminaRegistro
' Author  : witigo
' Date    : 03/11/2016
' Version : 1.0
' Purpose : Elimina el registro objetivo de la tabla seleccionada...
'
' @Param    String    strTableName
' @Param    String    strFieldName
' @Param    Long      lngRegistry    Optional
'--------------------------------------------------------------------------------
Public Function wtg_EliminaRegistro( _
					strTableName As String, _
                    lngRegistry As Long, _
                    Optional strFieldName As String = "id" _
                    )

Dim dbs As DAO.Database
Dim rst As DAO.Recordset

Dim strSQL As String

    strSQL = "SELECT " & strTableName & ".* " & _
                "FROM " & strTableName & " " & _
                "WHERE (((" & strTableName & "." & strFieldName & ")= " & lngRegistry & "));"

    Set dbs = CurrentDb()
    Set rst = dbs.OpenRecordset(strSQL, dbOpenDynaset)
    
    If rst.RecordCount <> 0 Then
        
        ' Eliminamos el registro
        rst.Delete

    Else

        ' Salimos de la funcion
        Exit Function

    End If

    ' Cerramos el recordset
    rst.Close
    ' Borramos los objetos
    Set rst = Nothing
    Set dbs = Nothing

End Function