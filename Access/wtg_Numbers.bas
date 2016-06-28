Attribute VB_Name = "wtg_Numbers"
Option Compare Database
Option Explicit


' Enumeraciones para su uso en la
' función wtg_OrdinalNumbers.
Public Enum GenderType
    Masculino
    Femenino
End Enum


'-------------------------------------------------------------------------------
' Method  : wtg_RandomNumber
' Author  : Witigo
' Date    : 01/06/2016
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


'-------------------------------------------------------------------------------
' Method  : wtg_OrdinalNumber
' Author  : Witigo (Angel Campos Muñoz)
' Date    : 07/10/2015
' Version :
' Purpose : Devuelve como número ordinal el número cardinal pasado como argu-
'           mento, la función devuelve el número ordinal en género masculino
'           o femeníno dependiendo del valor seleccionado en la enumeración
'           eGenero.
'
'           Opcionalmente, si queremos que se muestren las apocopes de los nú-
'           meros primero y tercero, debemos seleccionar TRUE en el valor
'           boleano bolApocope...
'
' @Param    integer   intNumber
' @Param    enum      eGender
' @Param    boolean   bolApocope
' @return   string    strNumber
'-------------------------------------------------------------------------------
Public Function wtg_OrdinalNumber( _
                    intNumber As Integer, _
                    ByVal eGender As GenderType, _
                    Optional bolApocope As Boolean = False _
                    ) As String

Dim u, d, c As Integer
Dim Unidad()
Dim Decena()
Dim Centena()
Dim strNumber As String

    If eGender = Masculino Then

        ' Rellenamos los arrays para los números odinales en MASCULINO
        Unidad = Array("", _
                        "primero", "segundo", "tercero", _
                        "cuarto", "quinto", "sexto", _
                        "septimo", "octavo", "noveno")

        Decena = Array("", _
                        "decimo", "vigesimo", "trigesimo", _
                        "cuadragesimo", "quincuagesimo", "sexagesimo", _
                        "septuagesimo", "octogesimo", "nonagesimo")

        Centena = Array("", _
                        "centesimo", "ducentesimo", "tricentesimo", _
                        "cuadringentesimo", "quingentesimo", "sexcentesimo", _
                        "septingentesimo", "octingentesimo", "noningentesimo")

    Else

        ' Rellenamos los arrays para los números ordinales en FEMENINO
        Unidad = Array("", _
                        "primera", "segunda", "tercera", _
                        "cuarta", "quinta", "sexta", _
                        "septima", "octava", "novena")

        Decena = Array("", _
                        "decima", "vigesima", "trigesima", _
                        "cuadragesima", "quincuagesima", "sexagesima", _
                        "septuagesima", "octogesima", "nonagesima")

        Centena = Array("", _
                        "centesima", "ducentesima", "tricentesima", _
                        "cuadringentesima", "quingentesima", "sexcentesima", _
                        "septingentesima", "octingentesima", "noningentesima")

    End If

    Select Case Len(CStr(intNumber))

        ' UNIDADES
        ' ----------
        Case 1

            u = intNumber
            strNumber = Unidad(u)

            If bolApocope Then

                ' Si definimos que queremos utilizar el apocope (1er y 3er)
                ' sustituimos el valor de la cadena strNumber por el
                ' apocope adecuado (sólo en primero y tercero).
                If strNumber = "primero" Then strNumber = "primer"
                If strNumber = "tercero" Then strNumber = "tercer"

            End If

        ' DECENAS
        ' ----------
        Case 2

            d = intNumber \ 10
            u = intNumber - (d * 10)

            strNumber = Decena(d) & " " & Unidad(u)

        ' CENTENAS
        ' ----------
        Case 3

            c = intNumber \ 100
            d = (intNumber - (c * 100)) \ 10
            u = intNumber - ((c * 100) + (d * 10))

            strNumber = Centena(c) & " " & Decena(d) & " " & Unidad(u)

    End Select

    wtg_OrdinalNumber = strNumber

End Function


' TODO: arreglar esta función
'--------------------------------------------------------
'
' OnlyNumbers
'
' Código escrito originalmente por Juan M Afán de Ribera.
' Estás autorizado a utilizarlo dentro de una aplicación
' siempre que esta nota de autor permanezca inalterada.
' En el caso de querer publicarlo en una página Web,
' por favor, contactar con el autor en
'
'     accessvba@ya.com
'
' Este código se brinda por cortesía de
' Juan M. Afán de Ribera
'
Sub OnlyNumbers(KeyAscii As Integer)
    ' si no es un número entre el 0 y el 9
    ' o es un punto o una coma
    If Not Chr(KeyAscii) Like "[0-9]" Then
        Select Case KeyAscii
            ' si es un retroceso, enter o tabulación
            Case vbKeyBack, vbKeyReturn, vbKeyTab
            ' no se hace nada
            Case Else
                ' si no, se anula el caracter
                ' introducido
                KeyAscii = 0
                Beep
        End Select
    End If
End Sub
