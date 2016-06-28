Attribute VB_Name = "wtg_Dates"
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
Public Function wtg_IsSunday(dtDate As Date) As Boolean

Dim bolSunday As Boolean

    bolSunday = False

    If Weekday(dtDate) = 1 Then bolSunday = True

    wtg_IsSunday = bolSunday

End Function
