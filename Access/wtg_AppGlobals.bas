Attribute VB_Name = "bas_AppGlobals"
Option Compare Database
Option Explicit


'------------------------------------------------------------------------------
' Constant : gcstrAppName
' Scope    : Public
'
' Purpose  : Nombre de la aplicaci�n
'------------------------------------------------------------------------------
Public Const gcstrAppName As String = "Integraci�n de equipos"


'------------------------------------------------------------------------------
' Constant : gcstrAppCompanyName
' Scope    : Public
'
' Purpose  : Nombre de la Compa��a desarrolladora de la base de datos
'------------------------------------------------------------------------------
Public Const gcstrAppCompanyName As String = "Proteksa Networks S.L."


'------------------------------------------------------------------------------
' Constant : gcstrAppDeveloper
' Scope    : Public
'
' Purpose  : Nombre del desarrollador de la base de datos
'------------------------------------------------------------------------------
Public Const gcstrAppDeveloper As String = "Angel Campos Mu�oz"


'------------------------------------------------------------------------------
' Constant : gcstrAppDeveloperMail
' Scope    : Public
'
' Purpose  : Correo electr�nico del desarrollador
'------------------------------------------------------------------------------
Public Const gcstrAppDeveloperMail As String = "witigo@msn.com"


'------------------------------------------------------------------------------
' Constant : gcstrAppSupportMail
' Scope    : Public
'
' Purpose  : Correo electr�nico de soporte para la aplicaci�n
'------------------------------------------------------------------------------
Public Const gcstrAppSupportMail As String = gcstrAppDeveloperMail


'------------------------------------------------------------------------------
' Constant : gcstrAppVer
' Scope    : Public
'
' Purpose  : Versi�n de la aplicaci�n
'------------------------------------------------------------------------------
Public Const gcstrAppVer As String = "1.0.0"


'------------------------------------------------------------------------------
' Constant : gcstrAppDir
' Scope    : Public
'
' Purpose  : Directorio base de la aplicaci�n
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
' Purpose  : Nombre de la tabla en la base de datos, donde se almacenar�n los
'            logs de error de la base de datos...
'------------------------------------------------------------------------------
Public Const gcstrErrorLogTableName As String = "dbo_LogError"


'------------------------------------------------------------------------------
' Sesi�n de usuario
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
