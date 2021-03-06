Attribute VB_Name = "bas_Exports"
Option Compare Database
Option Explicit


'-------------------------------------------------------------------------------
' Method  : wtg_ExportTableToDatabase
' Date    : 01/06/2016
' Version : 1.0
' Author  : Witigo
'
' Purpose : Exporta la tabla indicada a otra base de datos.
'
' @Param    String    strTableName
' @Param    String    strDestinationDatabase
'
' @Return    Boolean
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


'-------------------------------------------------------------------------------
' Method  : wtg_ExportReportPDF
' Date    : 01/06/2016
' Version : 1.0
' Author  : Witigo
'
' Purpose : Exporta el informe a formato PDF
'
' @Param    string     strReportName
' @Param    string     strCriteria
'-------------------------------------------------------------------------------
' TODO Terminar este modulo, debería permitir elegir que directorio y nombre de fichero queremos utilizar...
Public Function wtg_ExportReportPDF( _
                    strReportName As String, _
                    strCriteria As String _
                    )

Dim strDate As String
Dim strOutputFile As String

    ' Formateamos la fecha
    strDate = Format(Date, "ddmmyy")
    ' Buscamos el valor en la tabla
    strOutputFile = CurrentProject.Path & "\Reports\" & strDate & "_Report" & ".pdf"

    ' Abrimos el informe
    DoCmd.OpenReport strReportName, acViewPreview, , strCriteria

    ' Establecemos el formato de salida como PDF
    DoCmd.OutputTo acOutputReport, strReportName, acFormatPDF, strOutputFile, False

    ' Cerramos el Informe
    DoCmd.Close acReport, strReportName, acSaveNo

End Function
