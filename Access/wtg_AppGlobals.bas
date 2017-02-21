Attribute VB_Name = "bas_AppGlobals"
Option Compare Database
Option Explicit


'------------------------------------------------------------------------------
' Constant : gcstrAppName
' Scope    : Public
'
' Purpose  : Nombre de la aplicación
'------------------------------------------------------------------------------
Public Const gcstrAppName As String = "Integración de equipos"


'------------------------------------------------------------------------------
' Constant : gcstrAppCompanyName
' Scope    : Public
'
' Purpose  : Nombre de la Compañía desarrolladora de la base de datos
'------------------------------------------------------------------------------
Public Const gcstrAppCompanyName As String = "Proteksa Networks S.L."


'------------------------------------------------------------------------------
' Constant : gcstrAppDeveloper
' Scope    : Public
'
' Purpose  : Nombre del desarrollador de la base de datos
'------------------------------------------------------------------------------
Public Const gcstrAppDeveloper As String = "Angel Campos Muñoz"


'------------------------------------------------------------------------------
' Constant : gcstrAppDeveloperMail
' Scope    : Public
'
' Purpose  : Correo electrónico del desarrollador
'------------------------------------------------------------------------------
Public Const gcstrAppDeveloperMail As String = "witigo@msn.com"


'------------------------------------------------------------------------------
' Constant : gcstrAppSupportMail
' Scope    : Public
'
' Purpose  : Correo electrónico de soporte para la aplicación
'------------------------------------------------------------------------------
Public Const gcstrAppSupportMail As String = gcstrAppDeveloperMail


'------------------------------------------------------------------------------
' Constant : gcstrAppVer
' Scope    : Public
'
' Purpose  : Versión de la aplicación
'------------------------------------------------------------------------------
Public Const gcstrAppVer As String = "1.0.0"


'------------------------------------------------------------------------------
' Constant : gcstrAppDir
' Scope    : Public
'
' Purpose  : Directorio base de la aplicación
'------------------------------------------------------------------------------
Public Const gcstrAppDir As String = "C:\Proteksa\Integraciones"


'------------------------------------------------------------------------------
' Constant : gcstrBackupsDir
' Scope    : Public
'
' Purpose  : Ruta del directorio Backups de la base de datos...
'------------------------------------------------------------------------------
Public Const gcstrBackupsDir As String = gcstrAppDir & "\Backups"


'------------------------------------------------------------------------------
' Constant : gcstrDocumentsDir
' Scope    : Public
'
' Purpose  : Ruta del directorio Documents de la base de datos...
'------------------------------------------------------------------------------
Public Const gcstrDocumentsDir As String = gcstrAppDir & "\Documents"


'------------------------------------------------------------------------------
' Constant : gcstrLogsDir
' Scope    : Public
'
' Purpose  : Ruta del directorio Logs de la base de datos...
'------------------------------------------------------------------------------
Public Const gcstrLogsDir As String = gcstrAppDir & "\Logs"


'------------------------------------------------------------------------------
' Constant : gcstrReportsDir
' Scope    : Public
'
' Purpose  : Ruta del directorio Reports de la base de datos...
'------------------------------------------------------------------------------
Public Const gcstrReportsDir As String = gcstrAppDir & "\Reports"


'------------------------------------------------------------------------------
' Constant : gcstrStorageDir
' Scope    : Public
'
' Purpose  : Ruta del directorio Storage de la base de datos....
'------------------------------------------------------------------------------
Public Const gcstrStorageDir As String = gcstrAppDir & "\Storage"


'------------------------------------------------------------------------------
' Constant : gcstrErrorLogTableName
' Scope    : Public
'
' Purpose  : Nombre de la tabla en la base de datos, donde se almacenarán los
'            logs de error de la base de datos...
'------------------------------------------------------------------------------
Public Const gcstrErrorLogTableName As String = "dbo_LogError"


'------------------------------------------------------------------------------
' Sesión de usuario
'------------------------------------------------------------------------------
Private Type typSessionDetails
    intUser            As Integer
    strUser            As String
    strUserFirstName   As String
    strUserLastName    As String
    intRole            As Integer
    strRole            As String
End Type

Public gvarUserSession As typSessionDetails
