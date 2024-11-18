# Crud VB6                                                                     

Simple CRUD made with VB6 and Crystal Reports 8.5.

## Development environment

- Microsoft Visual Basic 6.0
- Crystal Reports 8.5

## Screenshots

Menu

<p align="center">
	<img src=".resources/screenshots/2024-10-25 20_43_20-Main.png" width="243">
</p>

Articles panel

<p align="center">
	<img src=".resources/screenshots/2024-10-25 20_52_32-ListArticles.png" width="629">
</p>

<br>
<br>

<p align="center">
	<img src=".resources/screenshots/2024-10-25 20_54_18-CreateArticle.png" width="243">
</p>

Reports

<p align="center">
	<img src=".resources/screenshots/2024-11-11 22_23_07-Reports.png" width="658">
</p>

<br>
<!--
	Nota: No se porque pero para que las imagenes se vean con buena calidad tuve que
	poner el ancho real de la imagen * 0.8. Por ej. 304 * 0.8 ~= 243. Así las caputras
	se ven igual que cuando se abren desde algún visualizador de imagenes en Windows.
-->

## Errors with CR 8.5

This is a list of some errors I encountered when working with CR 8.5 and wanting 
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