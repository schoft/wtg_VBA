Attribute VB_Name = "bas_fileFolder_Operations"
Option Compare Database
Option Explicit


Private Type SHFILEOPSTRUCT
    hWnd As Long                    ' Type: HWND
    wFunc As Long                   ' Type: UINT
    pFrom As String                 ' Type: PCZZTSTR
    pTo As String                   ' Type: PCZZTSTR
    fFlags As Integer               ' Type: FILEOP_FLAGS
    fAnyOperationsAborted As Long   ' Type: BOOL
    hNameMappings As Long           ' Type: LPVOID
    lpszProgressTitle As Long       ' Type: PCTSTR
End Type


' Declaración de constantes para wFunc
'--------------------------------------------------------------------------------
Private Const FO_MOVE As Long = &H1     ' Move the files specified in pFrom to the 
                                        ' location specified in pTo.

Private Const FO_COPY As Long = &H2     ' Copy the files specified in the pFrom 
                                        ' member to the location specified in the 
                                        ' pTo member.

Private Const FO_DELETE As Long = &H3   ' Delete the files specified in pFrom.

Private Const FO_RENAME As Long = &H4   ' Rename the file specified in pFrom. You 
                                        ' cannot use this flag to rename multiple 
                                        ' files with a single function call.
                                        ' Use FO_MOVE instead.


' Declaración de constantes para fFlags
'--------------------------------------------------------------------------------
Private Const FOF_MULTIDESTFILES As Long = &H1
' The pTo member specifies multiple destination files (one for each source file in 
' pFrom) rather than one directory where all source files are to be deposited.

Private Const FOF_CONFIRMMOUSE As Long = &H2
' Not used.

Private Const FOF_SILENT As Long = &H4
' Don't create progress/report

Private Const FOF_RENAMEONCOLLISION As Long = &H8
' Give the file being operated on a new name in a move, copy, or rename operation 
' if a file with the target name already exists at the destination.

Private Const FOF_NOCONFIRMATION As Long = &H10
' Respond with Yes to All for any dialog box that is displayed.

Private Const FOF_WANTMAPPINGHANDLE As Long = &H20
' If FOF_RENAMEONCOLLISION is specified and any files were renamed, assign a name 
' mapping object that contains their old and new names to the hNameMappings member. 
' This object must be freed using SHFreeNameMappings when it is no longer needed.

Private Const FOF_ALLOWUNDO = &H40
' Preserve undo information, if possible.

Private Const FOF_FILESONLY = &H80
' Perform the operation only on files (not on folders) if a wildcard file name (*.*) 
' is specified.

Private Const FOF_SIMPLEPROGRESS = &H100
' Display a progress dialog box but do not show individual file names as they are 
' operated on.

Private Const FOF_NOCONFIRMMKDIR = &H200
' Do not ask the user to confirm the creation of a new directory if the operation 
' requires one to be created.

Private Const FOF_NOERRORUI = &H400
' Do not display a dialog to the user if an error occurs.

Private Const FOF_NOCOPYSECURITYATTRIBS = &H800
' Version 4.71. Do not copy the security attributes of the file. The destination 
' file receives the security attributes of its new folder.


Private Declare Function SHFileOperation _
                Lib "shell32.dll" _
                Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long


' Declaración de constantes propias
'--------------------------------------------------------------------------------
Private const cstrDoubleNullCharacter as string = vbNullChar & vbNullChar


'--------------------------------------------------------------------------------
' Method  : wtg_WinAPI_Copy
' Author  : witigo
' Date    : 29/11/2016
' Version : 1.0
'
' Purpose : Copia un fichero o directorio a una ruta determinada...
'
' @Param    String    strSourcePath
' @Param    String    strTargetPath
'
' @Return   Boolean
'--------------------------------------------------------------------------------
Public Function wtg_WinAPI_Copy( _
					ByRef strSourcePath As String, _
					ByRef strTargetPath As String _
					) As Boolean
    
Dim objShf As SHFILEOPSTRUCT
Dim intFlags As Integer
Dim lngResult As Long

    intFlags = FOF_SIMPLEPROGRESS

    With objShf
        
        ' Operación
        .wFunc = FO_COPY
        ' Directorio origen
        .pFrom = strSourcePath & cstrDoubleNullCharacter
        ' Directorio destino
        .pTo = strTargetPath & cstrDoubleNullCharacter
        ' Flags de la operación
        .fFlags = intFlags
        
    End With

    ' Almacenamos el resultado de la operación
    lngResult = SHFileOperation(objShf)
    
    ' Devolvemos VERDADERO o FALSO según resultado
    wtg_WinAPI_Copy = (lngResult = 0)
    
End Function


'--------------------------------------------------------------------------------
' Method  : wtg_WinAPI_Delete
' Author  : witigo
' Date    : 30/11/2016
' Version : 1.0
'
' Purpose : Borra el directorio especificado.
'
' @Param    String    strTargetPath
'
' @Return   Boolean
'--------------------------------------------------------------------------------
Public Function wtg_WinAPI_Delete( _
					ByRef strTargetPath As String _
					) As Boolean

Dim objShf As SHFILEOPSTRUCT
Dim intFlags As Integer
Dim lngResult As Long

    intFlags = FOF_SILENT + FOF_ALLOWUNDO + FOF_NOCONFIRMATION
    
    With objShf
    
        ' Operación
        .wFunc = FO_DELETE
        ' Directorio objeto
        .pFrom = strTargetPath & cstrDoubleNullCharacter
        ' Flags para la operación
        .fFlags = intFlags
        
    End With

    ' Almacenamos el resultado de la operación
    lngResult = SHFileOperation(objShf)
    
    ' Devolvemos VERDADERO o FALSO según resultado
    wtg_WinAPI_Delete = (lngResult = 0)
    
End Function