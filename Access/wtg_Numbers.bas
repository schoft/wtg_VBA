Attribute VB_Name = "wtg_Numbers"
Option Compare Database
Option Explicit

'-------------------------------------------------------------------------------
' Method  : wtg_RandomNumber
' Author  : Witigo
' Date    : 4/2/2014
' Version : 1.0
' Purpose : Genera un número aleatorio comprendido entre dos números dados.
'
' @Param    integer     intMinimo
' @Param    integer     intMaximo
' @return   integer     intRandomNumber
'-------------------------------------------------------------------------------
Public Function wtg_RandomNumber( _
                    intMinNumber As integer, _
                    intMaxNumber as integer _
                    ) As integer

Dim intRandomNumber As integer

    intRandomNumber = intMaxNumber - intMinNumber + 1
    intRandomNumber = CInt(Int(intRandomNumber * Rnd())) + intMinNumber
    'intRandomNumber = CInt(Int((intMaxNumber - intMinNumber + 1) * Rnd())) + intMinNumber

    wtg_RandomNumber = intRandomNumber

End Function
