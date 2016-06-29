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


' TODO Terminar este modulo
Public Function wtg_ExportReportPDF( _
                    strReportName As String, _
                    strCriteria As String
                    )

Dim strDate As String
Dim strOutputFile As String

    ' Formateamos la fecha
    strDate = Format(Date, "ddmmyy")
    ' Buscamos el valor en la tabla
    strOutputFile = wtg_SelectFile
    strOutputFile = DLookup("[Dir_Reports]", "db_o_Settings", "[ID]=1")
    strOutputFile = strOutputFile & "\" & strDate & "_Campaña" & ".pdf"

    ' Abrimos el informe
    DoCmd.OpenReport strReprotName, acViewPreview, , strCriteria

    ' Establecemos el formato de salida como PDF
    DoCmd.OutputTo acOutputReport, strReportName, acFormatPDF, strOutputFile, False

    ' Cerramos el Informe
    DoCmd.Close acReport, strReprotName, acSaveNo

End Function
