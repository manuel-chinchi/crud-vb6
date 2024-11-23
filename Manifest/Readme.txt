app.manifest
===============================================================================

SOBRE EL ARCHIVO

Este archivo de manifiesto contiene un conjunto de metadatos en formato XML que
indican al sistema los ensamblados que van a enlazarse a la aplicaci�n en tiempo
de ejecuci�n. Esto aplica para el binario (.exe) una vez ya compilado ya que
que en tiempo de desarrollo el editor de Microsoft Visual Basic 6.0 no permite 
a�adir dicho manifiesto.

USO

Para incluir el manifiesto en el binario (.exe) se debe compilar fuera del edi
tor de Microsoft Visual Basic 6.0. de la siguiente forma:

1. crear el archivo de recursos (.RES).
`rc.exe app.rc`

2. abrir el proyecto y agregar el archivo de recursos creado recientemente.

3. compilar la aplicaci�n con el nuevo archivo de recursos (.RES)
`C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE /make crud_vb6.exe`

*NOTA*

Si no se va a cambiar la refencia al archivo de recursos ni la refencia dentro
de la aplicaci�n entonces se puede omitir el paso 2.


IMPORTANTE

El archivo de manifiesto debe ser de un tama�o que sea multiplo de 4. Por ej.
si el manifiesto a usar pesa 1255 bytes una vez compilado el .exe al querer
abrirlo saldra el siguiente mensaje:

```
No se pudo iniciar la aplicaci�n;la configuraci�n en paralelo no es correcta. 
Consulte el registro de eventos de la aplicaci�n o use la herramienta sxstrace
.exe de la l�nea de comandos para obtener m�s detalles.

```

RECURSOS
* https://learn.microsoft.com/es-es/windows/win32/sbscs/application-manifests