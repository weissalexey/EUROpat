Dim xlApp, xlWkb, SourceFolder,TargetFolder,file
Set xlApp = CreateObject("excel.application")
Set fs = CreateObject("Scripting.FileSystemObject")

Const xlNormal=1

SourceFolder="c:\xml-to-xls\xml"
TargetFolder="c:\xml-to-xls\xls"

xlApp.Visible = false

for each file in fs.GetFolder(SourceFolder).files
  Set xlWkb = xlApp.Workbooks.Open(file)
  BaseName= fs.getbasename(file)
  FullTargetPath=TargetFolder & "\" & BaseName & ".xls"
  xlWkb.SaveAs FullTargetPath, xlNormal
  xlWkb.close
next

fs.DeleteFile("C:\xml-to-xls\xml\*.xml")

Set xlWkb = Nothing
Set xlApp = Nothing
Set fs = Nothing
