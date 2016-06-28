Attribute VB_Name = "wtg_DatabaseObjects"
Option Compare Database
Option Explicit


'-------------------------------------------------------------------------------
' Method  : wtg_IsSunday
' Author  : Witigo
' Date    : 01/06/2016
' Version : 1.0
' Purpose : Muestra una ventana FileDialog de Microsoft Office, que devuelve la
'           ruta completa del directorio seleccionado.
'
' @Param    date     dtDate
' @return   boolean  strFolderPath
'-------------------------------------------------------------------------------
Public Function wtg_IsFormLoaded(ByVal strFormName As String) As Boolean

Dim objForm As Object

    ' Creamos una nueva instancia del objeto
    Set objForm = CurrentProject.AllForms(strFormName)

    ' Devolvemos el estado del formulario
    wtg_IsFormLoaded = objForm.IsLoaded

    ' Borramos el objeto
    Set objForm = Nothing

End Function
