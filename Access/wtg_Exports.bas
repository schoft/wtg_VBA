Attribute VB_Name = "wtg_Exports"
Option Compare Database
Option Explicit


'-------------------------------------------------------------------------------
' Method  : wtg_ExportTableToDatabase
' Author  : Witigo
' Date    : 01/06/2016
' Version : 1.0
' Purpose : Exporta la tabla indicada a otra base de datos.
'
' @Param    string     strTableName
' @return   string     strText
'-------------------------------------------------------------------------------
Public Function wtg_ExportTableToDatabase( _
                    strTableName As String, _
                    Optional strDestinationDatabase as string = "" _
                    ) As Boolean

Dim strDatabaseType As String

    strDatabaseType = "Microsoft Access"
    ' Mostramos FileDialog para seleccionar la base de datos donde guardar la
    ' tabla a exportar
    if strDestinationDatabase = "" then strDestinationDatabase = wtg_SelectFile

    DoCmd.TransferDatabase acExport, _
                           strDatabaseType, _
                           strDestinationDatabase, _
                           acTable, _
                           strTableName, _
                           strTableName, _
                           False, _
                           True

End Function
