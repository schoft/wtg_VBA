Attribute VB_Name = "wtg_ClassCreator"
Option Compare Database
Option Explicit


'-------------------------------------------------------------------------------
' Method  : wtg_ClassCreator
' Author  : Witigo
' Date    : 01/06/2016
' Version : 1.0
' Purpose : Genera toda la estructura de la clase a partir de una tabla.
'
' @Param    string     strText
'-------------------------------------------------------------------------------
Public Function wtg_ClassCreator(strTable As String)

Dim dbs As DAO.Database
Dim rst As DAO.Recordset
Dim fld As DAO.Field

    Set dbs = CurrentDb()
    Set rst = dbs.OpenRecordset(strTable, dbOpenDynaset)

Dim strPrefix As String
Dim strName As String
Dim strType As String
Dim strVarName As String
Dim strNull As String

Dim sNL As String
Dim sNP As String

    sNL = vbNewLine
    sNP = vbNewLine & vbNewLine

Dim strRstName As String

    strRstName = "mrstRecordset"

' Variable para la declaración de variables
Dim strDecl As String
' Variable para la declaración de propiedades
Dim strProp As String
' Variable para la declaración de funciones
Dim strFunc As String
' Variable para la todo el texto que compone la clase
Dim strClase As String


    ' -------------------------------------------------------------------------
    '   Declaraciones
    ' -------------------------------------------------------------------------
    strDecl = "Private " & strRstName & " As Recordset" & sNL
    strDecl = strDecl & "Private mbooLoaded As Boolean" & sNP

    For Each fld In rst.Fields

        ' Cargamos los valores y tipos
        strPrefix = wtg_GetPrefixType(fld.Type)
        strName = fld.Name
        strType = fGetType(fld.Type)

        ' Construimos el nombre de la variable
        strVarName = "m" & strPrefix & strName

        ' Declaramos las variables para las columnas de la tabla
        strDecl = strDecl & "Private " & strVarName & " As " & strType & sNL

    Next fld

    ' Insertamos una linea en blanco más
    strDecl = strDecl & sNL


    ' -------------------------------------------------------------------------
    '   Propiedades
    ' -------------------------------------------------------------------------
    For Each fld In rst.Fields

        ' Cargamos los valores y tipos
        strPrefix = fGetPrefixType(fld.Type)
        strName = fld.Name
        strType = fGetType(fld.Type)

        ' Construimos el nombre de la variable
        strVarName = "m" & strPrefix & strName

        ' PROPERTY LET
        strProp = strProp & "Public Property Let " & strName & "(" & strPrefix & strName & " As " & strType & ")" & sNL
        strProp = strProp & wtg_Tabs(1) & strVarName & " = " & strPrefix & strName & sNL
        strProp = strProp & "End Property" & sNP

        ' PROPERTY GET
        strProp = strProp & "Public Property Get " & strName & "() As " & strType & sNL
        strProp = strProp & wtg_Tabs(1) & strName & " = " & strVarName & sNL
        strProp = strProp & "End Property" & sNP

    Next fld


    ' -------------------------------------------------------------------------
    '   Funciones y Procedimientos
    ' -------------------------------------------------------------------------

    ' INICIALIZADOR
    strFunc = "Private Sub Class_Initialize()" & sNL
    strFunc = strFunc & wtg_Tabs(1) & "' Abrimos el recordset" & sNL
    strFunc = strFunc & wtg_Tabs(1) & "Set " & strRstName & " = CurrentDb.OpenRecordset(""" & strTable & """, dbOpenDynaset)" & sNL
    strFunc = strFunc & "End Sub" & sNP

    ' TERMINADOR
    strFunc = strFunc & "Private Sub Class_Terminate()" & sNL
    strFunc = strFunc & wtg_Tabs(1) & "' Cerramos el recordset" & sNL
    strFunc = strFunc & wtg_Tabs(1) & strRstName & ".Close" & sNL
    strFunc = strFunc & wtg_Tabs(1) & "' Borramos el objeto" & sNL
    strFunc = strFunc & wtg_Tabs(1) & "Set " & strRstName & " = Nothing" & sNL
    strFunc = strFunc & "End Sub" & sNP

    ' LOAD
    strFunc = strFunc & "Private Sub Load()" & sNP
    strFunc = strFunc & wtg_Tabs(1) & "With " & strRstName & sNL

    For Each fld In rst.Fields
        ' Cargamos el nombre del campo
        strName = fld.Name
        '
        strFunc = strFunc & wtg_Tabs(2) & "Me." & strName & " = Nz(.Fields(""" & strName & """).Value)" & sNL
    Next fld

    strFunc = strFunc & wtg_Tabs(1) & "End With" & sNP
    strFunc = strFunc & wtg_Tabs(1) & "mbooLoaded = True" & sNP
    strFunc = strFunc & "End Sub" & sNP

    ' FINDFIRST
    strFunc = strFunc & "Public Function FindFirst(Optional Criteria As Variant) As Boolean" & sNP
    strFunc = strFunc & wtg_Tabs(1) & "If IsMissing(Criteria) Then" & sNL
    strFunc = strFunc & wtg_Tabs(2) & strRstName & ".MoveFirst" & sNL
    strFunc = strFunc & wtg_Tabs(2) & "FindFirst = Not " & strRstName & ".EOF" & sNL
    strFunc = strFunc & wtg_Tabs(1) & "else" & sNL
    strFunc = strFunc & wtg_Tabs(2) & strRstName & ".FindFirst Criteria" & sNL
    strFunc = strFunc & wtg_Tabs(2) & "FindFirst = Not " & strRstName & ".NoMatch" & sNL
    strFunc = strFunc & wtg_Tabs(1) & "End If" & sNP
    strFunc = strFunc & wtg_Tabs(1) & "If FindFirst Then Load" & sNP
    strFunc = strFunc & "End Function" & sNP

    ' UPDATE
    strFunc = strFunc & "Public Sub Update()" & sNP
    strFunc = strFunc & wtg_Tabs(1) & "With " & strRstName & sNP
    strFunc = strFunc & wtg_Tabs(2) & "If mbooLoaded = True Then" & sNL
    strFunc = strFunc & wtg_Tabs(3) & "' Editamos el registro" & sNL
    strFunc = strFunc & wtg_Tabs(3) & ".Edit" & sNL
    strFunc = strFunc & wtg_Tabs(2) & "else" & sNL
    strFunc = strFunc & wtg_Tabs(3) & "' Agregamos un nuevo registro" & sNL
    strFunc = strFunc & wtg_Tabs(3) & ".AddNew" & sNP
    strFunc = strFunc & wtg_Tabs(2) & "End If" & sNP

    For Each fld In rst.Fields
        ' Cargamos el nombre del campo
        strName = fld.Name
        If fld.Type = 10 And fld.AllowZeroLength = False _
        Or fld.Type = 12 And fld.AllowZeroLength = False Then
            strFunc = strFunc & wtg_Tabs(2) & ".Fields(""" & strName & """).Value = Nz(Me." & strName & ",""NULL"")" & sNL
        Else
            strFunc = strFunc & wtg_Tabs(2) & ".Fields(""" & strName & """).Value = Nz(Me." & strName & ","""")" & sNL
        End If
    Next fld

    strFunc = strFunc & sNL
    strFunc = strFunc & wtg_Tabs(2) & "' Actualizamos el registro" & sNL
    strFunc = strFunc & wtg_Tabs(2) & ".update" & sNP
    strFunc = strFunc & wtg_Tabs(1) & "End With" & sNP
    strFunc = strFunc & wtg_Tabs(1) & "mbooLoaded = True" & sNP
    strFunc = strFunc & "End Sub"


    ' -------------------------------------------------------------------------
    '   Completamos la Clase
    ' -------------------------------------------------------------------------
    strClase = strDecl & strProp & strFunc

    'copy from immediate window and paste into class module
    'Debug.Print strDeclare & vbTab & vbNewLine & scode
    Call wtg_CreateCLS(strTable, strClase)

    ' Cerramos el recordset
    rst.Close
    ' Borramos los objetos
    Set rst = Nothing
    Set dbs = Nothing

End Function


'-------------------------------------------------------------------------------
' Method  : wtg_GetFieldType
' Author  : Witigo
' Date    : 01/06/2016
' Version : 1.0
' Purpose : Devuevle el nombre del tipo de campo de las tablas.
'
' @Param    integer   intTableFieldType
' @Return   string    strFieldType
'-------------------------------------------------------------------------------
Private Function wtg_GetFieldType(intTableFieldType As Integer) As String

Dim strFieldType As String

    Select Case intTableFieldType

        Case 1
            ' Boolean (True/False) data
            strFieldType = "boolean"
        Case 2
            ' Byte (8-bit) data
            strFieldType = "byte"
        Case 3
            ' Integer data
            strFieldType = "integer"
        Case 4
            ' Long Integer data
            strFieldType = "long"
        Case 5
            ' Currency data
            strFieldType = "currency"
        Case 6
            ' Single-precision floating-point data
            strFieldType = "single"
        Case 7
            ' Double-precision floating-point data
            strFieldType = "double"
        Case 8
            ' Date value data
            strFieldType = "date"
        Case 10
            ' Text data (variable width)
            strFieldType = "string"
        Case 12
            ' Memo data (extended text)
            strFieldType = "string"
        Case 20
            ' Decimal data (ODBCDirect only)
            strFieldType = "variant"
        Case 11
            ' Binary data (bitmap)
            strFieldType = "binary"
        Case 15
            ' GUID data
            strFieldType = ""
        Case 101
            ' Attachment data
            strFieldType = ""

    End Select

    wtg_GetFieldType = strFieldType

End Function


'-------------------------------------------------------------------------------
' Method  : wtg_GetPrefixType
' Author  : Witigo
' Date    : 01/06/2016
' Version : 1.0
' Purpose : Devuelve la abreviatura de un tipo de campo según la LNC
'           (Lezinsky Naming Convention)
'
' @Param    integer    intTableFieldType
' @Return   string     strLCN
'-------------------------------------------------------------------------------
Private Function wtg_GetPrefixType(intTableFieldType As Integer) As String

Dim strLCN As String

    Select Case intTableFieldType

        Case 1
            ' Boolean (True/False) data
            strLCN = "bol"
        Case 2
            ' Byte (8-bit) data
            strLCN = "byt"
        Case 3
            ' Integer data
            strLCN = "int"
        Case 4
            ' Long Integer data
            strLCN = "lng"
        Case 5
            ' Currency data
            strLCN = "cur"
        Case 6
            ' Single-precision floating-point data
            strLCN = "sng"
        Case 7
            ' Double-precision floating-point data
            strLCN = "dbl"
        Case 8
            ' Date value data
            strLCN = "dtm"
        Case 10
            ' Text data (variable width)
            strLCN = "str"
        Case 12
            ' Memo data (extended text)
            strLCN = "str"
        Case 20
            ' Decimal data (ODBCDirect only)
            strLCN = "var"
        Case 11
            ' Binary data (bitmap)
            strLCN = "bin"
        Case 15
            ' GUID data
            strLCN = ""
        Case 101
            ' Attachment data
            strLCN = ""

    End Select

    wtg_GetPrefixType = strLCN

End Function


'-------------------------------------------------------------------------------
' Method  : wtg_CreateCLS
' Author  : Witigo
' Date    : 01/06/2016
' Version : 1.0
' Purpose : Crea mediante programación un nuevo módulo de clase en la base de
'           datos.
'
' @Param    string   strClsName
' @Param    string   strClsBody
'-------------------------------------------------------------------------------
Public Function wtg_CreateCLS(strClsName As String, strClsBody As String)

Dim basClsModule As Module

    ' Creamos el nuevo modulo de clase
    DoCmd.RunCommand acCmdNewObjectClassModule

    ' Set MyModule to be the new Module Object.
    Set basClsModule = Application.Modules.Item(Application.Modules.count - 1)

    With basClsModule

        ' Agregamos la opción de declaración implicita de variables
        .InsertLines 2, "Option explicit"

        ' Agregamos el texto para la nueva clase
        .InsertText strClsBody

    End With

    ' Guardamos los cambios en el modulo de clase
    DoCmd.Save acModule, basClsModule

    ' Cerramos el modulo de clase
    DoCmd.Close acModule, basClsModule, acSaveYes

    ' Renombramos el modulo de clase
    DoCmd.Rename strClsName, acModule, basClsModule

End Function
