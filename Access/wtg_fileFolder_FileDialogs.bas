Attribute VB_Name = "wtg_fileFolder_FileDialogs"
Option Compare Database
Option Explicit
'-------------------------------------------------------------------------------
' ¡Atención!
'
' Recuerda que para que las funciones de este módulo funcionen correctamente,
' se ha de agregar la referencia a una de las siguientes librerias de objetos.
'
' - Microsoft Office 12.0 Object Library (Office 2007)
' - Microsoft Office 14.0 Object Library (Office 2010)
' - Microsoft Office 15.0 Object Library (Office 2013)
'
'-------------------------------------------------------------------------------


'-------------------------------------------------------------------------------
' Method  : wtg_SelectFolder
' Author  : Witigo
' Date    : 01/06/2016
' Version : 1.0
' Purpose : Muestra una ventana FileDialog de Microsoft Office, que devuelve la
'           ruta completa del directorio seleccionado.
'
' @Param    String    strTitle
' @Return   String    strFolderPath
'-------------------------------------------------------------------------------
Public Function wtg_SelectFolder( _
                    Optional strTitle As String = "Seleccione un directorio" _
                    ) As String

Dim fDialog As Office.FileDialog

Dim strFolderPath as string

    strFolderPath = ""

    ' Creamos una nueva instancia del objeto FileDialog
    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)

    With fDialog

        ' Titulo de la ventana
        .Title = strTitle

        ' Texto para el botón
        .ButtonName = "Seleccionar directorio"

        ' Mostramos la ventana
        ' Si el método .Show devuelve:
        '   True  >>>> El usuario a seleccionado un directorio.
        '   False >>>> El usuario a pulsado el botón Cancelar.
        If .Show = True Then strFolderPath = .SelectedItems(1)

    End With

    wtg_SelectFolder = strFolderPath

    ' Borramos el objeto
    Set fDialog = Nothing

End Function


'-------------------------------------------------------------------------------
' Method  : wtg_SaveFileAs
' Author  : Witigo
' Date    : 01/06/2016
' Version : 1.0
' Purpose : Muestra una ventana FileDialog de Microsoft Office, de tipo SaveAs
'           que devuelve la ruta completa del fichero que queremos guardar.
'
' @Return   String    strFullFilePath
'-------------------------------------------------------------------------------
Public Function wtg_SaveFileAs() As String

Dim fDialog As Office.FileDialog

Dim strFullFilePath As String

    strFullFilePath = ""

    ' Creamos una nueva instancia del objeto FileDialog
    Set fDialog = Application.FileDialog(msoFileDialogSaveAs)

    With fDialog

        ' Titulo de la ventana
        .Title = "Guardar"

        ' Texto del botón
        .ButtonName = "Guardar como"

        ' Mostramos la ventana
        ' Si el método .Show devuelve:
        '   True  >>>> El usuario a introducido un nombre de fichero y ha
        '              pulsado el botón Guardar.
        '   False >>>> El usuario a pulsado el botón Cancelar.
        If .Show = True Then strFullFilePath = .SelectedItems.Item(1)

    End With

    wtg_SaveFileAs = strFullFilePath

    ' Borramos el objeto
    Set fDialog = Nothing

End Function


'-------------------------------------------------------------------------------
' Method  : wtg_SelectFile
' Author  : Witigo
' Date    : 01/06/2016
' Version : 1.0
' Purpose : Muestra una ventana FileDialog de Microsoft Office, que devuelve la
'           ruta completa del fichero seleccionado.
'
' @Param    String    strTitle
' @Return   String    strFilePath
'-------------------------------------------------------------------------------
Public Function wtg_SelectFile( _
                    Optional strTitle As String = "Select a folder" _
                    ) As String

Dim fDialog As Office.FileDialog

Dim strFilePath as string

    strFilePath = ""

    ' Creamos una nueva instancia del objeto FileDialog
    Set fDialog = Application.FileDialog(msoFileDialogOpen)

    With fDialog

        ' Titulo de la ventana
        .Title = strTitle

        ' Texto para el botón
        .ButtonName = "Seleccionar fichero"

        ' Filtros de extensión
        With .Filters

            ' Primero borramos todos los filtros
            .Clear

            ' Para añadir un filtro, sustituimos los tres parámetros.
            ' (El tercer parámetro es opcional)
            ' Ejemplo: .Add <Description>, <Extension>, <[position/index]>

            ' Añadimos los filtros por defecto
            .Add "Microsoft Access files", "*.accdb, *.mdb"
            .Add "Microsoft Excel files", "*.xls, *.xlsx"
            .Add "Microsoft Word files", "*.doc, *.docx"
            .Add "CSV files", "*.csv"
            .Add "Adobe PDF files", "*.pdf"
            .Add "Text files", "*.txt"

        End With

        ' Si hay más de un filtro de extensión, podemos seleccionar cual será
        ' seleccionado por defecto
        ' .filterindex = <index>
        .filterindex = 1

        ' Mostramos la ventana
        ' Si el método .Show devuelve:
        '   True  >>>> El usuario a seleccionado un fichero.
        '   False >>>> El usuario a pulsado el botón Cancelar.
        If .Show = True Then strFilePath = .SelectedItems(1)

    End With

    wtg_SelectFile = strFilePath

    ' Borramos el objeto
    Set fDialog = Nothing

End Function


'-------------------------------------------------------------------------------
' Method  : wtg_FolderExist
' Author  : Witigo
' Date    : 01/06/2016
' Version : 1.0
' Purpose : Comprueba la existencia de un directorio.
'
' @Param    String    strFolderPath
' @Return   Boolean   bolFolderPath
'-------------------------------------------------------------------------------
Public Function wtg_FolderExist( _
                    strFolderPath As String _
                    ) As Boolean

Dim bolFolderPath As Boolean

    bolFolderPath = False

    ' Comprobamos la existencia del directorio.
    If Dir(strFolderPath, vbDirectory) <> "" Then bolFolderPath = True

    wtg_FolderExist = bolFolderPath

End Function
