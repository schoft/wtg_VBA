Attribute VB_Name = "wtg_Dates"
Option Compare Database
Option Explicit


'-------------------------------------------------------------------------------
' Method  : wtg_IsSunday
' Author  : Witigo
' Date    : 01/06/2016
' Version : 1.0
' Purpose : Comprueba si la fecha pasada como argumento corresponde a Domingo.
'           Devuelve VERDADERO si es Domingo
'
' @Param    Date     dtDate
' @Return   Boolean  wtg_IsSunday
'-------------------------------------------------------------------------------
Public Function wtg_IsSunday( _
					dtDate As Date _
					) As Boolean

Dim bolSunday As Boolean

    bolSunday = False

    If Weekday(dtDate) = 1 Then bolSunday = True

    wtg_IsSunday = bolSunday

End Function
