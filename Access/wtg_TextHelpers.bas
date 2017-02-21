Attribute VB_Name = "bas_TextHelpers"
Option Compare Database
Option Explicit


' Enumeración para determinar la coplejidad de la cadena de texto aleatoria...
Enum eComplexity
    LettersOnly
    LettersAndSigns
    NumbersOnly
    NumbersAndLetters
    NumbersAndSigns
    NumbersLettersSigns
    SignsOnly
End Enum

'-------------------------------------------------------------------------------
' Method  : wtg_CutString
' Author  : Witigo
' Date    : 01/06/2016
' Version : 1.0
' Purpose: Corta una cadena de texto a una determinada longitud
'
' @Param    String     strMessage
' @Param    Integer    intLenght
'
' @Return    String
'-------------------------------------------------------------------------------
Public Function wtg_CutString( _
                    strText As String, _
                    intLenght As Integer _
                    ) As String

    wtg_CutString = Mid(strText, 1, intLenght)

End Function


'------------------------------------------------------------------------------
' Method  : wtg_RandomText
' Author  : witigo
' Date    : 21/02/2017
' Version : 1.0
'
' Purpose : Genera una cadena de texto aleatorio con una longitud y complejidad
'           determinada en los parámetros de la función...
'
' @Param    Integer        intLenght
' @Param    eComplexity    Complexity
'
' @Return   String
'------------------------------------------------------------------------------
Public Function wtg_RandomText( _
                    intLenght As Integer, _
                    Complexity As eComplexity _
                    ) As String

Dim i As Integer
Dim intRandom As Integer

Dim strChars As String
Dim strLetters As String
Dim strNumbers As String
Dim strSigns As String

Dim strRandomtext As String

    strLetters = "abcdefghijklmnñopqrstuvwxyz" & _
                 "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ"

    strNumbers = "0123456789"

    strSigns = "(){}[]<>/|\.:,;_-='^*¡!¿?Ç@#%&€$"

    Select Case Complexity

        Case 0  ' LettersOnly

            strChars = strLetters

        Case 1  ' LettersAndSigns

            strChars = strLetters & strSigns

        Case 2  ' NumbersOnly

            strChars = strNumbers

        Case 3  ' NumbersAndLetters

            strChars = strNumbers & strLetters

        Case 4  ' NumbersAndSigns

            strChars = strNumbers & strSigns

        Case 5 ' NumbersLettersSigns

            strChars = strLetters & strNumbers & strSigns

        Case 6  ' SignsOnly

            strChars = strSigns

    End Select

    ' Hacemos un bucle hasta el número máximo de carácteres especificado en
    ' el parámetro de la función...
    For i = 1 To intLenght

        ' Generamos un número aleatorio hasta un máximo de la longitud de la
        ' cadena de texto strChars...
        intRandom = Int((Len(strChars) * Rnd) + 1)

        ' Adjuntamos a la cadena de texto aleatorio el carácter situado en
        ' la posición del número generado de forma aleatoria...
        strRandomtext = strRandomtext & Mid(strChars, intRandom, 1)

    Next i

    wtg_RandomText = strRandomtext

End Function



'-------------------------------------------------------------------------------
' Method  : wtg_StripAccent
' Author  : Witigo
' Date    : 01/06/2016
' Version : 1.0
' Purpose : Quita los acentos, cedillas, dieresís de una cadena de texto.
'
' @Param    string     strText
'
' @Return   string     strText
'-------------------------------------------------------------------------------
Public Function wtg_StripAccent( _
                    strText As String _
                    ) as string

Dim strA As String * 1
Dim strB As String * 1
Dim i As Integer

Const AccChars = "áàäâéèëêíìïîóòöôúùüûÁÀÄÂÉÈËÊÍÌÏÎÓÒÖÔÚÙÜÛ"
Const RegChars = "aaaaeeeeiiiioooouuuuAAAAEEEEIIIIOOOOUUUU"

    For i = 1 To Len(AccChars)

        ' Obtenemos los carácteres'
        strA = Mid(AccChars, i, 1)
        strB = Mid(RegChars, i, 1)

        ' Reemplazamos el carácter con signos de puntuación por el carácter 
        ' sin signos...
        strText = Replace(strText, strA, strB, 1, , vbBinaryCompare)

    Next

    wtg_StripAccent = strText

End Function

'-------------------------------------------------------------------------------
' Method  : wtg_Tabs
' Author  : Witigo
' Date    : 01/06/2016
' Version : 1.0
' Purpose : Inserta la cantidad de espacios equivalentes al número de
'           tabulaciones definido por la varible intTabs.
'
' @Param    Integer   intTabs
'
' @Return   String    strBlankMessage
'-------------------------------------------------------------------------------
Public Function wtg_Tabs( _
                    intTabs as integer _
                    ) As String

Dim i As Integer
Dim intSpacesXTab As Integer
Dim intTotalSpaces As Integer

    ' Número de espacios en blanco que equivalen a una tabulacion
    intSpacesXTab = 4
    
    ' Número total de espacios en blanco
    intTotalSpaces = intTabs * intSpacesXTab

Dim strBlankMessage As String

    strBlankMessage = ""

    For i = 1 To intTotalSpaces

        strBlankMessage = strBlankMessage & " "

    Next i

    wtg_Tabs = strBlankMessage

End Function