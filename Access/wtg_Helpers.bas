Attribute VB_Name = "wtg_Helpers"
Option Compare Database
Option Explicit


'--------------------------------------------------------------------------------
' Method : wtg_DisablePag
' Author : witigo
' Date   : 28/10/2016
' Version: 1.0
' Purpose: Deshabilita el uso de las teclas AvPag y RePag para el cambio de
'          registro en formularios...
' Use    : In form_KeyDown method...
'          KeyCode = wtg_DisablePag(KeyCode)
'
' @Param    Integer     intKeyPress
'--------------------------------------------------------------------------------
Public Function wtg_DisablePag(intKeyPress As Integer) As Integer

    Select Case intKeyPress
    
        Case 33, 34
        
            ' Deshabilitamos la puslación de tecla
            wtg_DisablePag = 0
        
        Case Else
        
            ' Devolvemos el valor de la pulsación de tecla
            wtg_DisablePag = intKeyPress
    
    End Select
    
End Function


'--------------------------------------------------------------------------------
' Method  : wtg_FormCaption
' Author  : witigo
' Date    : 31/10/2016
' Version : 1.0
' Purpose : Establece el nombre del formulario dependiendo de la operatoria con 
' 			el formulario...
'
' @Param    String  StrFormName
' @Param    String  StrObjetivo
'--------------------------------------------------------------------------------
Public Sub wtg_FormCaption(strFormName As String, _
                           Optional strObjetivo As String = "")

    ' Comprobamos si es un nuevo registro
    If Forms(strFormName).Form.NewRecord Then
    
    	' Formulario abierto en modo agregar
        Forms(strFormName).Form.Caption = "Nuevo registro"
        
    Else
    
        ' Comprobamos si podemos editar
        If Forms(strFormName).Form.AllowEdits Then
            
            ' Formulario abierto en modo edición
            Forms(strFormName).Form.Caption = "Editar registro"
        
        Else
        
            ' Formulario abierto en modo sólo lectura
            Forms(strFormName).Form.Caption = "Detalles " & strObjetivo
        
        End If
        
    End If
    
End Function
