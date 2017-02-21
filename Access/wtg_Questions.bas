Attribute VB_Name = "bas_Questions"
Option Compare Database
Option Explicit


'--------------------------------------------------------------------------------
' Method  : wtg_Eliminar
' Author  : witigo
' Date    : 03/11/2016
' Version : 1.0.1
'
' Purpose : Pregunta para eliminar un registro
'
' @Param    String    strMessage    Optional
'
' @Return	 Boolean
'--------------------------------------------------------------------------------
Public Function wtg_Eliminar( _
                    Optional strMessage as string = "¿Desea eliminar el registro?" _
                    ) As Boolean

Dim strCaption as String

	strCaption = "Eliminar"
    
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
' Version : 1.0.1
'
' Purpose : Pregunta para guardar los cambios a un registro
'
' @Param    String    strMessage    Optional
'
' @Return    Boolean
'--------------------------------------------------------------------------------
Public Function wtg_Guardar( _
                    Optional strMessage as string = "¿Desea eliminar el registro?" _
                    ) As Boolean

Dim strCaption as String

	strCaption = "Guardar"
    
    Select Case MsgBox(strMessage, vbYesNo Or vbQuestion, strCaption)
        
        Case vbYes
        
            wtg_Guardar = True
    
        Case vbNo
        
            wtg_Guardar = False
    
    End Select

End Function