Attribute VB_Name = "bas_IPAddress"
Option Compare Database
Option Explicit

'-------------------------------------------------------------------------------
' NOTE :
' To calculate the decimal address from a dotted string, perform the following
' calculation.
'
' (1st octet  * 256³)+ (2nd octet * 256²)+ (3rd octet  * 256) + (4th octet)
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
' NOTE :
' Para almacenar los números de tipo long en la base de datos, utilizaremos
' en el campo de la tabla el tipo de dato numérico (double)
'-------------------------------------------------------------------------------


'-------------------------------------------------------------------------------
' Method  : wtg_ValidateIP
' Author  : Witigo
' Date    : 08/08/2016
' Version : 1.0
' Purpose : Realiza la validación de una IP
'
' @Param    Variant    varIPAddress
' @Return   Boolean
'-------------------------------------------------------------------------------
Public Function wtg_ValidateIP( _
                    varIPAddress As Variant _
                    ) As Boolean

Dim bolResult  As Boolean
Dim intLoop    As Integer
Dim intNumber  As Integer
Dim intPos     As Integer
Dim intPrevPos As Integer
Dim strMsg     As String

    strMsg = ""
    strMsg = strMsg & "La dirección IP introducida no es válida."
    strMsg = strMsg & vbCrLf & vbCrLf
    strMsg = strMsg & "El valor de cada uno de los octetos no puede ser superior a 255."

    bolResult = True

    If UBound(Split(varIPAddress, ".")) = 3 Then

        For intLoop = 1 To 4

            intPos = InStr(intPrevPos + 1, varIPAddress, ".", 1)

            If intLoop = 4 Then intPos = Len(varIPAddress) + 1

            intNumber = Int(Mid(varIPAddress, intPrevPos + 1, intPos - intPrevPos - 1))

            If intNumber > 255 Then

                bolResult = False

                MsgBox strMsg, vbOKOnly Or vbExclamation Or vbSystemModal, Application.Name

            End If

            intPrevPos = intPos

        Next

    End If

    wtg_ValidateIP = bolResult

End Function


'-------------------------------------------------------------------------------
' Method  : wtg_IP2Long
' Author  : Witigo
' Date    : 08/08/2016
' Version : 1.0
' Purpose : Realiza la conversión de una dirección IP válida a un número de tipo
'           Long para poder almacenarlo en la base de datos.
'
' @Param    variant    varIPAddress
' @Return   variant    wtg_IP2Long
'-------------------------------------------------------------------------------
Public Function wtg_IP2Long( _
                    varIPAddress As Variant _
                    ) As Variant

Dim IntLoop    As Integer
Dim intNumber  As Integer
Dim intPos     As Integer
Dim intPrevPos As Integer

    If UBound(Split(varIPAddress, ".")) = 3 Then

        ' Loop entre los octetos
        For IntLoop = 1 To 4

            intPos = InStr(intPrevPos + 1, varIPAddress, ".", 1)

            If IntLoop = 4 Then intPos = Len(varIPAddress) + 1

            intNumber = Int(Mid(varIPAddress, intPrevPos + 1, intPos - intPrevPos - 1))

            If intNumber > 255 Then

                wtg_IP2Long = "0"

                Exit Function

            End If

            intPrevPos = intPos

            wtg_IP2Long = ((intNumber Mod 256) * (256 ^ (4 - IntLoop))) + wtg_IP2Long

        Next

    End If

End Function


'-------------------------------------------------------------------------------
' Method  : wtg_Long2IP
' Author  : Witigo
' Date    : 08/08/2016
' Version : 1.0
' Purpose : Realiza la conversión de un número de tipo Long a una dirección IP.
'
' @Param    variant    varLong
' @Return   variant    wtg_Long2IP
'-------------------------------------------------------------------------------
Public Function wtg_Long2IP( _
                    varLong As Variant _
                    ) As Variant

Dim IntLoop    As Integer
Dim intNumber  As Integer
Dim intPos     As Integer
Dim intPrevPos As Integer

    If IsNumeric(varLong) Then

        wtg_Long2IP = "0.0.0.0"

        For IntLoop = 1 To 4

            intNumber = Int(varLong / 256 ^ (4 - IntLoop))

            varLong = varLong - (intNumber * 256 ^ (4 - IntLoop))

            If intNumber > 255 Then

                wtg_Long2IP = "0.0.0.0"

                Exit Function

            End If

            If IntLoop = 1 Then

                wtg_Long2IP = wtg_Zeros4IP(intNumber)

            Else

                wtg_Long2IP = wtg_Long2IP & "." & wtg_Zeros4IP(intNumber)

            End If

        Next

    End If

End Function

'-------------------------------------------------------------------------------
' Method  : wtg_Zeros4IP
' Author  : Witigo
' Date    : 08/08/2016
' Version : 1.0
' Purpose : Añade ceros delante del número, para que este tenga siempre 3 cifras.
'
' @Param    Integer   intNumber
' @Return   String    wtg_Zeros4IP
'-------------------------------------------------------------------------------
Public Function wtg_Zeros4IP( _
                    intNumber As Integer _
                    ) As String

Dim strNumber as string

    Select Case intNumber

        Case Is < 10

            strNumber = "00" & CStr(intNumber)

        Case Is < 100

            strNumber = "0" & CStr(intNumber)

        Case Else

            strNumber = CStr(intNumber)

    End Select

    wtg_Zeros4IP = strNumber

End Function
