Attribute VB_Name = "wtg_DatabaseObjects"
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


' TODO: Arreglar esta función
Public Function ExportaTabla(tabla As String) As Boolean

Dim tblName As String
Dim dbDestination As String
Dim dbType As String

tblName = tabla
dbDestination = "C:\Pacientes\_Dat.accdb"
dbType = "Microsoft Access"

DoCmd.TransferDatabase acExport, dbType, dbDestination, acTable, tblName, tblName, False, True

End Function

' TODO Arreglar esta función
Private Sub Form_Load()
DoCmd.RunCommand acCmdAppMaximize

End Sub

' TODO Arreglar esta función
Public Sub abrir(strForm As String)

DoCmd.OpenForm strForm, , , , , acWindowNormal
End Sub

' TODO Arreglar esta función
Public Sub cerrar_F(strForm As String)

DoCmd.Close acForm, strForm
End Sub
