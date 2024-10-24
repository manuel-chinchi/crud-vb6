# Crud VB6                                                                     

Simple CRUD made with VB6 and Crystal Reports 8.5.

## Development environment

- Microsoft Visual Basic 6.0
- Crystal Reports 8.5

## Errors with CR 8.5

This is a list of some errors I encountered when working with CR 8.5 and wanting 
to deploy the application to a clean machine. I hope it helps you.

| LIBRARY      | ERROR IF NOT FOUND                                                         |
|--------------|----------------------------------------------------------------------------|
| crviewer.dll | Runtime-Error '339' Component crviewer.dll or one of its dependences<br> not correctly registered: a file is missing or invalid |
| craxdrt.dll  | Runtime-Error '-2147024770 (8007007e)': Automation error                   |
| P2smon.dll   | Physical database not found.                                               |
| crxf_pdf.dll | Runtime-Error '-2147190548 (800478ec)': Missing or out-of-date expor DLL.  |
| crtslv.dll   | Runtime-Error '-2147190908 (80047784)': Failed to export the report.       |
| EXPMOD.dll   | Runtime-Error '-2147190908 (80047784)': Failed to export the report.       |
| u2ddisk.dll  | Runtime-Error '-2147190548 (800478ec)': Missing or out-of-date export DLL. |

## References

* [Report no showing data](https://stackoverflow.com/questions/67210371/vb6-crystal-report-8-5-not-refresh-data)
* [Export report to PDF](https://stackoverflow.com/questions/1356588/how-to-export-to-a-pdf-file-in-crystal-report)
* [Error: 'Pyhsical database not found'](https://www.tek-tips.com/threads/crystal-reports-8-0-ttx-quot-physical-database-not-found-quot.34935/)
* [Error: cr8.5 Export to PDF](https://stackoverflow.com/questions/18062033/vb-6-0-crystal-reports-export-to-pdf)
* [Error: cr8.5 export to PDF (bendito seas)](https://www.vbforums.com/showthread.php?196385-RESOLVED-gt-Error-while-exporting-a-crystal-report)
* [Working with Crystal Reports 8.5 and .ttx files](http://www.crystalreportsbook.com/forum/forum_posts.asp?TID=14087#:~:text=A%20Data%20Definition%20file%20is,one%20piece%20of%20sample%20data.)
