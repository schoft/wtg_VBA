Attribute VB_Name = "wtg_Strings"
Option Compare Database
Option Explicit


'-------------------------------------------------------------------------------
' Method  : wtg_CutString
' Author  : Witigo
' Date    : 01/06/2016
' Version : 1.0
' Purpose: Corta una cadena de texto a una determinada longitud
'
' @Param    String    strMessage
' @Param    Integer   intLenght
' @Return   String
'-------------------------------------------------------------------------------
Public Function wtg_CutString( _
                    strText As String, _
                    intLenght As Integer _
                    ) As String

    wtg_CutString = Mid(strText, 1, intLenght)

End Function


'-------------------------------------------------------------------------------
' Method  : wtg_StripAccent
' Author  : Witigo
' Date    : 01/06/2016
' Version : 1.0
' Purpose : Quita los acentos, cedillas, dieresís de una cadena de texto.
'
' @Param    string     strText
' @Return   string     strText
'-------------------------------------------------------------------------------
Public Function wtg_StripAccent(strText As String) as string

Dim strA As String * 1
Dim strB As String * 1
Dim i As Integer

Const AccChars = "áàäâéèëêíìïîóòöôúùüûÁÀÄÂÉÈËÊÍÌÏÎÓÒÖÔÚÙÜÛ"
Const RegChars = "aaaaeeeeiiiioooouuuuAAAAEEEEIIIIOOOOUUUU"

    For i = 1 To Len(AccChars)

        ' Obtenemos los carácteres'
        strA = Mid(AccChars, i, 1)
        strB = Mid(RegChars, i, 1)

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
' @Return   String    strBlankMessage
'-------------------------------------------------------------------------------
Public Function wtg_Tabs(intTabs as integer) As String

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
