# wtg_VBA


Witigos VBA helper functions for rapid VBA Application development.

## Módulos VBA

### wtg_ClassCreator

Con este módulo podemos generar de forma automática un módulo de clase para cualquier tabla de la base de datos...

[Ver...](https://github.com/witigo/wtg_VBA/blob/master/Access/wtg_ClassCreator.bas)

### wtg_DatabaseObjects


[Ver...](https://github.com/witigo/wtg_VBA/blob/master/Access/wtg_DatabaseObjects.bas)

### wtg_Dates

En este módulo hay procedimientos para manejar las fechas en Access.

[Ver...](https://github.com/witigo/wtg_VBA/blob/master/Access/wtg_Dates.bas)

### wtg_ErrorHandler

En este módulo hay funciones y procedimientos para crear un sistema de gestión de errores.

[Ver...](https://github.com/witigo/wtg_VBA/blob/master/Access/wtg_ErrorHandler.bas)

### wtg_ErrorLog

En este módulo hay funciones y procedimientos para crear un log de errores de nuestra aplicación.

[Ver...](https://github.com/witigo/wtg_VBA/blob/master/Access/wtg_ErrorLog.bas)

### wtg_Exports


[Ver...](https://github.com/witigo/wtg_VBA/blob/master/Access/wtg_Exports.bas)

### wtg_FileSystem

En este módulo, hay funciones para trabajar con el sistema de ficheros del sistema operativo.

	- Con **wtg_SelectFolder** mostramos un fileDialog para seleccionar un directorio.
	La función nos devuelve la ruta completa del directorio.
	- Con **wtg_SaveFileAs** mostramos un fileDialog para guardar un fichero en un directorio.
	La función nos devuelve la ruta completa del fichero a guardar.
	- Con **wtg_SelectFile** mostramos un fileDialog para seleccionar un fichero.
	La función nos devuelve la ruta completa del fichero a seleccionar.
	- Con **wtg_FolderExist** comprobamos la existencia de un directorio.

[Ver...](https://github.com/witigo/wtg_VBA/blob/master/Access/wtg_FileSystem.bas)

### wtg_Globals

Variables globales usadas para la gestión de las sesiones.

[Ver...](https://github.com/witigo/wtg_VBA/blob/master/Access/wtg_Globals.bas)

### wtg_Helpers


[Ver...](https://github.com/witigo/wtg_VBA/blob/master/Access/wtg_Helpers.bas)

### wtg_IPAddress

En este módulo, encontramos funciones para trabajar con direcciones IP

	- Con **wtg_ValidateIP** podemos validar una dirección IP.
	- Con **wtg_IP2Long** podemos convertir una dirección IP a un número de tipo Long.
	- Con **wtg_Long2IP** podemos convertir un número de tipo Long a una dirección IP.
	- Con **wtg_Zeros4IP** podemos mostrar una dirección IP con los ceros a la izquierda.
	(ejemplo: 192.168.001.001)


[Ver...](https://github.com/witigo/wtg_VBA/blob/master/Access/wtg_IPAddress.bas)

### wtg_Numbers

En este módulo, hay funciones para trabajar con números...

	- Con **wtg_RandomNumber** podemos generar un número aleatorio comprendido entre 2 números 
	(min y max)
	- Con **wtg_OrdinalNumber** podemos generar el texto/número ordinal a partir de un número 
	decimal.
	- Con **wtg_OnlyNumbers** podemos anular la introducción de cualquier caracter que no sea 
	un número.

> Quiero mostrar mi agradecimiento a Juan M. Afán de Ribera creador de la función wtg_OnlyNumbers 
(wtg_OnlyNumbers no es su nombre original)

[Ver...](https://github.com/witigo/wtg_VBA/blob/master/Access/wtg_Numbers.bas)

### wtg_Questions

En este módulo, hay funciones que generan las preguntas comunes en la base de datos.

	- Con **wtg_Eliminar** mostramos un messagebox con la pregunta "¿Desea eliminar el registro?"
	- Con **wtg_Guardar** mostramos un messagebox con la pregunta "¿Desea guardar los cambios?"

[Ver...](https://github.com/witigo/wtg_VBA/blob/master/Access/wtg_Questions.bas)

### wtg_SessionLog

En este módulo, hay funciones para trabajar con las sesiones de usuario en la aplicación.

> Para trabajar con sesiones de usuario en nuestra aplicación, es **muy importante** controlar todos los errores, ya que cualquier error no controlado, cierra/borra la sesión.

	- Con **wtg_WriteSesionLog** insertamos un registro con la información de la sesión en la tabla de 
	logs de sesión.
	- Con **wtg_CreateSessionLog_Table** podemos crear de forma automática la tabla para las sesiones 
	de usuario de la aplicación.

[Ver...](https://github.com/witigo/wtg_VBA/blob/master/Access/wtg_SessionLog.bas)

### wtg_Strings

En este módulo, hay funciones para trabajar con cadenas de texto...

	- Con **wtg_StripAccent** podemos quitar los acentos diacríticos, dieresis y comillas de un cadena de texto.
	- Con **wtg_CutString** podemos cortar una cadena de texto a una longitud determinada.
	- Con **wtg_Tabs** podemos insertar un número equivalente de espacios a una tabulación.

[Ver...](https://github.com/witigo/wtg_VBA/blob/master/Access/wtg_Strings.bas)

### wtg_Tables

En este módulo, hay funciones comunes para trabajar con las tablas de la base de datos.

	- Con **wtg_EliminarRegistro** podemos eliminar un registro determinado de la tabla objetivo.
	- Con **wtg_CheckRegistro** podemos comprobar si existe un registro determinado un la tabla objetivo.

[Ver...](https://github.com/witigo/wtg_VBA/blob/master/Access/wtg_Tables.bas)