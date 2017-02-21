Attribute VB_Name = "bas_fileFolder_Checks"
Option Compare Database
Option Explicit


'------------------------------------------------------------------------------
' Method  : wtg_FolderExist
' Date    : 01/06/2016
' Version : 1.0.1
' Author  : Witigo
'
' Purpose : Comprueba la existencia de un directorio.
'
' @Param    String    strFolderPath
'
' @Return    Boolean
'------------------------------------------------------------------------------
Public Function wtg_FolderExist( _
                    strFolderPath As String _
                    ) As Boolean

    ' Comprobamos la existencia del directorio.
    wtg_FolderExist = (Dir(strFolderPath, vbDirectory) <> "")

End Function


'------------------------------------------------------------------------------
' Method  : wtg_FileExist
' Date    : 30/11/2016
' Version : 1.0.1
' Author  : witigo
'
' Purpose : Comprueba la existencia de un fichero
'
' @Param    String    strFilePath
'
' @Return    Boolean
'------------------------------------------------------------------------------
Public Function wtg_FileExist( _
                    strFilePath As String _
                    ) As Boolean

    ' Comprobamos la existencia del directorio.
    wtg_FileExist = (Dir(strFilePath, vbArchive) <> "")

End Function