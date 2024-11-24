# Crud VB6                                                                     

Simple CRUD made with VB6, Crystal Reports 8.5 and SQLite.

## Overview

This project consists of a small CRUD (create, read, update, delete) system 
made in Visual Basic 6 that uses SQLite as a local database engine. It also 
uses Crystal Reports for reports and allows export to other formats.

The project consists of 3 modules: 

* Articles module with create/update/delete and search operations.
* Categories module with create/delete operations.
* Reports module with function to export to .pdf, .doc, and .xls formats.

It is not necessary to install the SQLite ODBC to run the application since
 it has a special class module for this that makes use of the internal 
 functions of sqlite.dll (included in the repository)

Also included is the script to create the installer using InnoSetup

## How to use this?

In order to modify the project, you only need to have Crystal Reports 8.5 
installed (it automatically installs the .OCX libraries used in the 
application) and then open the crud_vb6.vbp project since Microsoft Visual 
Basic 6.0. **It is not recommended to modify the cSQLiteConnection.cls 
file as it contains the functions necessary to interact with the SQLite 
database**

Then, to test your modified version (test the .exe) with an installer you 
only need to do the following:

1. Run the `build.bat` file from a terminal as administrator, 
	this will create a crud_vb6.exe in the root folder
2. Open the `install.iss` file with InnoSetup and option 
	 Build > Compile. This will create the installer in the `InnoSetup_Installer` folder

And voila, you now have your own customized version of the CRUD_VB6 system with its standalone installer.

## Development environment

- Microsoft Visual Basic 6.0 (VB6)
- Crystal Reports 8.5
- InnoSetup 6.3.3
- Greenshot 1.2.10 (for screenshots)

and the classics... Sublime Text, Notepad++, VS Code, etc.

## Platforms

Windows 7 or latter

## Screenshots

Installer

<p align="center">
	<img src=".readme/newscreenshots/2024-11-23%2020_45_11-Seleccione%20el%20Idioma%20de%20la%20Instalaci%C3%B3n.png" width="279">
</p>

<br>
<br>

<p align="center">
	<img src=".readme/newscreenshots/2024-11-23%2020_45_52-Instalar%20-%20crud%20vb6%20(by%20Manuel%20%20Chinchi)%20versi%C3%B3n%201.0.png" width="466">
</p>

Main menu

<p align="center">
	<img src=".readme/newscreenshots/2024-11-17 21_39_56-Main.png" width="243">
</p>

Articles module

<p align="center">
	<img src=".readme/newscreenshots/2024-11-17 21_41_06-ListArticles.png" width="629">
</p>

<br>
<br>

<p align="center">
	<img src=".readme/newscreenshots/2024-11-17 21_41_22-CreateArticle.png" width="243">
</p>

Reports module

<p align="center">
	<img src=".readme/newscreenshots/2024-11-17 21_41_51-Reports.png" width="658">
</p>

<br>

<!--
	Nota: No se porque pero para que las imagenes se vean con buena calidad tuve que
	poner el ancho real de la imagen * 0.8. Por ej. 304 * 0.8 ~= 243. Así las caputras
	se ven igual que cuando se abren desde algún visualizador de imagenes en Windows.
-->

## Errors with Crystal Reports

This is a list of some errors I encountered when working with Crystal Reports 8.5 and wanting 
to deploy the application to a clean machine. I hope it helps you.

| LIBRARY      | ERROR IF NOT FOUND                                                         |
|--------------|----------------------------------------------------------------------------|
| crviewer.dll | Runtime-Error '339' Component crviewer.dll or one of its dependences<br> not correctly registered: a file is missing or invalid |
| craxdrt.dll  | Runtime-Error '-2147024770 (8007007e)': Automation error                   |
| P2smon.dll*  | Physical database not found.                                               |
| crxf_pdf.dll | Runtime-Error '-2147190548 (800478ec)': Missing or out-of-date expor DLL.  |
| crtslv.dll   | Runtime-Error '-2147190908 (80047784)': Failed to export the report.       |
| EXPMOD.dll   | Runtime-Error '-2147190908 (80047784)': Failed to export the report.       |
| u2ddisk.dll  | Runtime-Error '-2147190548 (800478ec)': Missing or out-of-date export DLL. |
| u2fwordw.dll  | Runtime-Error '-2147190548 (800478ec)': Missing or out-of-date export DLL. |
| u2fxls.dll  | Runtime-Error '-2147190548 (800478ec)': Missing or out-of-date export DLL. |

luckily and thank God all these errors are resolved simply by registering these libraries. View the file dependencies.bat

(\*) It may not be necessary

## References

* [Report no showing data](https://stackoverflow.com/questions/67210371/vb6-crystal-report-8-5-not-refresh-data)
* [Export report to PDF](https://stackoverflow.com/questions/1356588/how-to-export-to-a-pdf-file-in-crystal-report)
* [Error: 'Pyhsical database not found'](https://www.tek-tips.com/threads/crystal-reports-8-0-ttx-quot-physical-database-not-found-quot.34935/)
* [Error: cr8.5 Export to PDF](https://stackoverflow.com/questions/18062033/vb-6-0-crystal-reports-export-to-pdf)
* [Error: cr8.5 export to PDF (bendito seas)](https://www.vbforums.com/showthread.php?196385-RESOLVED-gt-Error-while-exporting-a-crystal-report)
* [Working with Crystal Reports 8.5 and .ttx files](http://www.crystalreportsbook.com/forum/forum_posts.asp?TID=14087#:~:text=A%20Data%20Definition%20file%20is,one%20piece%20of%20sample%20data.)
* [Trust Center locked word file](https://learn.microsoft.com/es-es/office/troubleshoot/settings/file-blocked-in-office)
* [Use of Dictionary in VB6](https://www.codestack.net/visual-basic/data-sets/dictionary/)
* [Use of dictionary in VBA/VB6](https://vba846.wordpress.com/objeto-dictionary-para-vba/)
* [ListView column click order](https://www.vbforums.com/showthread.php?275658-ListView-Column-Click-(Sort)-Resolved!!)
* [Sort ListView column click](https://www.vbforums.com/showthread.php?301328-Vb6-Sort-Listview-By-Dates-numbers-text)
* [Embebed manifest into exe (vb6)](https://stackoverflow.com/questions/2182815/embedding-an-application-manifest-into-a-vb6-exe)
* [vb6 icons free](https://www.vbcorner.net/download_icons.htm)
* [Permission Windows](https://learn.microsoft.com/en-us/previous-versions/bb756929(v=msdn.10)?)