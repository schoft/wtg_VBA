Attribute VB_Name = "wtg_Questions"
Option Compare Database
Option Explicit


'--------------------------------------------------------------------------------
' Method  : wtg_Eliminar
' Author  : witigo
' Date    : 03/11/2016
' Version : 1.0
' Purpose : Pregunta para eliminar un registro
'
' @Return	Boolean
'--------------------------------------------------------------------------------
Public Function wtg_Eliminar() As Boolean

Dim strCaption as String
Dim strMessage As String

	strCaption = "Eliminar"
    strMessage = "¿Desea eliminar el registro?"
    
    Select Case MsgBox(strMessage, vbYesNo Or vbQuestion, strCaption)
        
        Case vbYes
        
            wtg_Eliminar = True
    
        Case vbNo
        
            wtg_Eliminar = False
    
    End Select

End Function


'--------------------------------------------------------------------------------
' Method  : wtg_Guardar
' Author  : witigo
' Date    : 03/11/2016
' Version : 1.0
' Purpose : Pregunta para guardar los cambios a un registro
'
' @Return	Boolean
'--------------------------------------------------------------------------------
Public Function wtg_Guardar() As Boolean

Dim strCaption as String
Dim strMessage As String

	strCaption = "Guardar"
    strMessage = "¿Desea guardar los cambios realizados?"
    
    Select Case MsgBox(strMessage, vbYesNo Or vbQuestion, strCaption)
        
        Case vbYes
        
            wtg_Guardar = True
    
        Case vbNo
        
            wtg_Guardar = False
    
    End Select

End Function