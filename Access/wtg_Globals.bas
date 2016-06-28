Attribute VB_Name = "wtg_Globals"
Option Compare Database
Option Explicit


Public Const AppName As String = "<AppName>"
Public Const AppVersion As String = "V.1.0"
Public Const AppDeveloper As String = "Angel Campos Muñoz"
Public Const AppDeveloperMail As String = "<developer@email.com>"


' Definimos los tipos personalizados
' para las variables globales...
Private Type typDetallesSesion
    intUsuario          As Integer
    strUsuario          As String
    strNombreUsuario    As String
    strApellidosUsuario As String
    intRol              As Integer
    strRolUsuario       As String
End Type


' Definimos las variables globales
Global gvarSesion As typDetallesSesion
