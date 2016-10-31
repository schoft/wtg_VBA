Attribute VB_Name = "wtg_DatabaseObjects"
Option Compare Database
Option Explicit


'-------------------------------------------------------------------------------
' Method  : wtg_IsFormLoaded
' Author  : Witigo
' Date    : 01/06/2016
' Version : 1.0
' Purpose : Muestra una ventana FileDialog de Microsoft Office, que devuelve la
'           ruta completa del directorio seleccionado.
'
' @Param    string   strFormName
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


'-------------------------------------------------------------------------------
' Procedure : wtg_IsRuntime
' Author    : Witigo
' Date      : 01/02/2014
' Version   : 1.0
' Purpose   : Comprueba el modo de funcionamiento de la aplicación y devuelve
'             VERDADERO si la aplicación se está ejecutando en modo RunTime...
'-------------------------------------------------------------------------------
Public Function wtg_IsRuntime() As Boolean

Dim bolRunTime As Boolean

    bolRunTime = Application.SysCmd(acSysCmdRuntime)

    wtg_IsRuntime = bolRunTime

End Function

' TODO Terminar esta función
'-------------------------------------------------------------------------------
' Procedure : StatusBar
' Author    : Witigo (Angel Campos Muñoz)
' Date      : 26/11/2014
' Version   : 1.0
' Purpose   :
'
' @Param    String   strMessage
'-------------------------------------------------------------------------------
Public Function StatusBar(Optional strMessage As String = vbNullString)

Dim Temp As Variant

    ' If the Msg variable is omitted or is empty,
    ' return the control of the status bar to Access
    If strMessage <> vbNullString Then

        ' Mostramos el mensaje
        SysCmd(acSysCmdSetStatus, strMessage)

    Else

        ' Borramos cualquier mensaje
        SysCmd(acSysCmdClearStatus)

    End If

End Function


'-------------------------------------------------------------------------------
' Procedure : wtg_OpenEditForm
' Author    : Witigo
' Date      : 26/11/2014
' Version   : 1.0
' Purpose   : Abre el formulario en modo EDICION.
'             Opcionalmente se puede añadir una clausula where.
'
' @Param    String   strFormName
' @Param    String   strWhere
'-------------------------------------------------------------------------------
Public Sub wtg_OpenEditForm( _
                strFormName As String, _
                Optional strWhere as string = "" _
                )

    ' Abrimos el formulario en modo edición
    DoCmd.OpenForm strFormName, acNormal, , strWhere, acFormEdit, acWindowNormal

End Sub


'-------------------------------------------------------------------------------
' Procedure : wtg_OpenReadOnlyForm
' Author    : Witigo
' Date      : 26/11/2014
' Version   : 1.0
' Purpose   : Abre el formulario en modo SOLO LECTURA.
'             Opcionalmente se puede añadir una clausula where.
'
' @Param    String   strFormName
' @Param    String   strWhere
'-------------------------------------------------------------------------------
Public Sub wtg_OpenReadOnlyForm( _
                strFormName As String, _
                Optional strWhere as string = "" _
                )

    ' Abrimos el formulario en modo edición
    DoCmd.OpenForm strFormName, acNormal, , strWhere, acFormEdit, acWindowNormal

End Sub


'-------------------------------------------------------------------------------
' Procedure : wtg_CloseForm
' Author    : Witigo
' Date      : 26/11/2014
' Version   : 1.0
' Purpose   : Cierra el formulario
'
' @Param    String   strFormName
'-------------------------------------------------------------------------------
Public Sub wtg_CloseForm(strFormName As String)

    ' Cerramos el formulario sin guardar cambios.
    DoCmd.Close acForm, strFormName, acSaveNo

End Sub
