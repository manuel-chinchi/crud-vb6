install.iss
===============================================================================

SOBRE EL ARCHIVO

Este script de instalaci�n contiene las instrucciones para crear el instalador
de forma autom�tica usando la herramienta InnoSetup (versi�n 6.3.3)

USO 

Para poder usar el script primero se debe ejecutar el archivo build.bat que se
encuentra en ./Scripts/build.bat para que se cree el binario (.exe) ya que no
se incluye en el repositorio por seguridad.

Una vez hecho esto abrir el archivo ./Scripts/Install/install.iss con InnoSetup
e ir a la opci�n Build > Compile. Se creara una carpeta ./InnoSetup_Installer 
y un archivo setup.exe

*NOTA*

El script debe ejecutarse desde la ubicaci�n donde se encuentre una vez descar
gado el repositorio del proyecto crud_vb6. En caso de querer usarlo desde otro
lugar modificar la variable `RootPath` del mismo.