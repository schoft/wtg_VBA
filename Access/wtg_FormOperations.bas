Attribute VB_Name = "bas_FormOperations"
Option Compare Database
Option Explicit


'------------------------------------------------------------------------------
' Method  : wtg_MostrarRegistroEnForm
' Date    : 21/02/2017
' Version : 1.0.0
' Company : Proteksa Networks S.L.
' Author  : Witigo
'
' Purpose : Muestra en el formulario indicado, el registro indicado de la tabla
'
' Notes   : Para que la función no presente fallos, el formulario objetivo, ha
'           de tener los nombres de los controles iguales a los campos de la
'           tabla...
'
' @Param    String    strFormName
' @Param    String    strTableName
' @Param    Long      lngRegistro
' @Param    String    strFieldPreffix
'------------------------------------------------------------------------------
Public Function wtg_MostrarRegistroEnForm( _
                    strFormName As String, _
                    strTableName As String, _
                    lngRegistro As Long, _
                    Optional strFieldPreffix As String = vbNullString _
                    )

Dim dbs As DAO.Database
Dim rst As DAO.Recordset

Dim frm As Form
Dim ctl As control

Dim strPreffix As String

    ' Instanciamos los objetos
    Set dbs = CurrentDb()
    Set rst = dbs.OpenRecordset("dbo_Equipos", dbOpenDynaset)

    Set frm = Forms(strFormName)

    ' Comprobamos si strFieldPreffix es nula, si es nula, utilizamos como cadena de filtro la predeterminadad
    If strFieldPreffix = vbNullString Then

        strFieldPreffix = "txt|cbo|chk|lst"
    
    End If
    
    With rst
    
        ' Buscamos el registro
        .FindFirst "[ID] = " & lngRegistro
    
        If .NoMatch Then
        
            ' No existe
        
        Else
        
            ' Hacemos un bucle por cada uno de los controles del formulario y
            ' rellenaremos su valor con los campos de la tabla de la base de
            ' datos con igual nombre...
            For Each ctl In frm.Controls
            
                ' Seleccionamos únicamente los tipos de controles CheckBox,
                ' Combobox y Textbox...
                If ctl.ControlType = acTextBox _
                Or ctl.ControlType = acComboBox _
                Or ctl.ControlType = acCheckBox Then
                    
                    ' Obtenemos el prefijo del control
                    strPreffix = Left(ctl.Name, 3)
                    
                    ' Filtramos los controles, trabajaremos únicamente con los
                    ' controles con el mismo nombre que los campos de tabla en
                    ' la base de datos...
                    If Not (InStr(1, strFieldPreffix, strPreffix, vbBinaryCompare) > 0) Then

                        ' El valor del ctl es igual al valor almacenado en el campo
                        ' con el mismo nombre de la tabla...
                        frm.Controls(ctl.Name).Value = .Fields(ctl.Name).Value
                    
                    End If
                
                End If
                
            Next ctl
            
        End If
        
    End With

    ' Cerramos el recordset
    rst.Close
    ' Eliminamos los objetos
    Set rst = Nothing
    Set dbs = Nothing

End Function


'------------------------------------------------------------------------------
' Method  : wtg_GuardarRegistroDesdeForm
' Date    : 21/02/2017
' Version : 1.0.0
' Company : Proteksa Networks S.L.
' Author  : Witigo
'
' Purpose : Toma los valores de los controles del formulario especificado y los
'           guarda en la tabla de la base de datos indicada...
'
' Notes   : Para guardar el registro y que no de ningún fallo, los campos del
'           formulario, han de tener el mismo nombre que el campo de la tabla
'           en la base de datos...
'
' @Param    String     strFormName
' @Param    String     strTableName
' @Param    Boolean    bolNewRecord
' @Param    Long       lngRegistro
' @Param    String     strFieldPreffix
'------------------------------------------------------------------------------
Public Function wtg_GuardarRegistroDesdeForm( _
                    strFormName As String, _
                    strTableName As String, _
                    Optional bolNewRecord As Boolean = True, _
                    Optional lngRegistro As Long = 0, _
                    Optional strFieldPreffix As String = vbNullString _
                    )

Dim dbs As DAO.Database
Dim rst As DAO.Recordset

Dim frm As Form
Dim ctl As control

    ' Instanciamos los objetos
    Set dbs = CurrentDb()
    Set rst = dbs.OpenRecordset("dbo_Equipos", dbOpenDynaset)

    Set frm = Forms(strFormName)

Dim strPreffix As String
Dim strExcluded As String

    strExcluded = "ID|created_at|created_by"
    
    ' Comprobamos si se ha definido una cadena de filtro
    If strFieldPreffix = vbNullString Then
    
        ' Valores por defecto como prefijo de campo según sistema de
        ' nomenclatura Leszynski...
        strFieldPreffix = "cbo|chk|lst|opt|txt"
    
    End If
    
    With rst
    
        ' Comprobamos si se trata de un nuevo registro o no...
        If bolNewRecord Then
        
            ' Agregamos un nuevo registro
            .AddNew
        
            ' Rellenamos los campos de fecha de creación del registro y el
            ' usuario que lo ha creado...
            '
            ' El campo "created_by" está previsto ser usado en aplicaciones
            ' de bases de datos multi-usuario...
            .Fields("created_at").Value = Now()
            .Fields("created_by").Value = "Admin"
            
        Else
        
            ' Buscamos el registro
            .FindFirst "[ID] = " & lngRegistro
            
            ' Guardamos el registro
            .Edit
        
            ' Rellenamos los campos de fecha de modificación del registro y
            ' el usuario que lo ha modificado...
            '
            ' El campo "created_by" está previsto ser usado en aplicaciones
            ' de bases de datos multi-usuario...
            .Fields("updated_at").Value = Now()
            .Fields("updated_by").Value = "Admin"
            
        End If
        
        ' Recorremos todos los controles del formulario
        For Each ctl In frm.Controls
        
            ' Trabajaremos directamente con los controles del formulario de
            ' tipo Checkbox, Combobox y Textbox...
            If ctl.ControlType = acTextBox _
            Or ctl.ControlType = acComboBox _
            Or ctl.ControlType = acCheckBox Then
                    
                ' Obtenemos el prefijo del control
                strPreffix = Left(ctl.Name, 3)

                ' Filtramos los controles, trabajaremos únicamente con los
                ' controles con el mismo nombre que los campos de tabla en
                ' la base de datos...
                If Not (InStr(1, strFieldPreffix, strPreffix, vbBinaryCompare) > 0) Then
                
                    ' Comprobamos si el campo es el campo "ID", puesto que el
                    ' campo "ID" es el campo autonumérico, NO PODEMOS asignarle
                    ' un valor...
                    If Not (InStr(1, strExcluded, ctl.Name, vbBinaryCompare) > 0) Then
                    
                        .Fields(ctl.Name).Value = Nz(ctl.Value, Null Or 0)
                    
                    End If
                    
                End If
                
            End If
        
        Next ctl
        
        ' Guardamos los cambios
        .Update
        
        ' Mostramos el ID sólo para los nuevos registros
        'If bolNewRecord Then
        
            ' Vamos al registro recién creado
            'rst.Bookmark = rst.LastModified
            
            ' Mostramos el ID
            'Me.ID = rst.Fields("ID")
        
        'End If
        
    End With

    ' Cerramos el recordset
    rst.Close
    ' Eliminamos los objetos
    Set rst = Nothing
    Set dbs = Nothing

End Function