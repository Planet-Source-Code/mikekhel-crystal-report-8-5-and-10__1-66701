<div align="center">

## Crystal Report 8\.5 and 10


</div>

### Description

crystal report 8.5 and 10..

dynamic server using setLogoninfo in sqlserver
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[mikekhel](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mikekhel.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mikekhel-crystal-report-8-5-and-10__1-66701/archive/master.zip)





### Source Code

```
'-------
'for crystal report 8.5
'Add this code to a form named form1 with a CrystalReports Viewer and reference
'the Crystal Reports 8.5 ActiveX Designer Run Time Library
'-------
'-------
'for crystal 10.0
'Add this code to a form named form1 and module with a CrystalReports Viewer and reference
'the Crystal Reports ActiveX Designer Run Time Library 10.0
'-----------
'form
Option Explicit
'Add this code to a form named form1 and module with a CrystalReports Viewer and reference
'the Crystal Reports ActiveX Designer Run Time Library 10.0
Private Sub Form_Load()
  strSelect = " your sqlquery"
  viewReport strSelect, App.Path & "\your report file"
End Sub
Private Sub Form_Resize()
CRViewer1.Width = ScaleWidth
CRViewer1.Height = ScaleHeight
End Sub
'--------------
'module
Option Explicit
Public crApp As New CRAXDRT.Application
Public crRep As CRAXDRT.Report
Public dbTable As CRAXDRT.DatabaseTable
Public strSelect As String
Function viewReport(ByVal strSql As String, ByVal strReportFile As String)
Set crRep = New CRAXDRT.Report
Set crApp = CreateObject("crystalruntime.application")
Set crRep = crApp.OpenReport(strReportFile)
For Each dbTable In crRep.Database.Tables
 dbTable.SetLogOnInfo "servername", "databasename", "", ""
Next dbTable
crRep.SQLQueryString = strSql
Form1.CRViewer1.ReportSource = crRep
Form1.CRViewer1.viewReport
Set crRep = Nothing
Set crApp = Nothing
End Function
```

