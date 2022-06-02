'aw Xml to csv 01.06.2022




Dim oXML


Set oXML = CreateObject("Microsoft.XMLDOM")

'Load the XML file
oXML.Load("C:\Users\aw\Desktop\Neuer Ordner\euro_part\InFileBeispil.xml")

If oXML.parseError <> 0 Then
  WScript.Echo oXML.parseError.reason
  WScript.Quit 1
End If

'Loop through each nodes
For Each oChdNd In oXML.DocumentElement.ChildNodes
if oChdNd.nodeName = "InvoiceNumber" then InvoiceNumber = oChdNd.text
if oChdNd.nodeName = "InvoiceType" then InvoiceType = oChdNd.text
if oChdNd.nodeName = "InvoiceDate" then InvoiceDate = oChdNd.text   
if oChdNd.nodeName = "Netto" then Netto = oChdNd.text      
if oChdNd.nodeName = "VATPercent" then VATPercent = oChdNd.text   
if oChdNd.nodeName = "VATAmount" then VATAmount = oChdNd.text 
if oChdNd.nodeName = "FreightCosts" then FreightCosts = oChdNd.text 
Next

CSVSTR = "TDKOPF" & ";" & InvoiceNumber & ";" & InvoiceType & ";" & InvoiceDate & ";" & Netto & ";" & VATPercent & ";" & VATAmount & ";" & FreightCosts & ";" 

Set nodes = oXML.selectNodes("/Invoice/Delivery/OrderNumber")

For Each node In nodes
 OrderNumber = node.text
 CSVSTR = CSVSTR &  OrderNumber & ";" &vbCrLf
 WriteCSV CSVSTR
 
Next

Set nodes = oXML.selectNodes("/Invoice/Delivery/FreeText")

Const ForReading = 1
Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile("c:\Users\aw\Desktop\Neuer Ordner\euro_part\20220601.csv", ForReading)


For Each node In nodes
 FreeText = node.text
 myLine = f.ReadLine
 if left(myLine, 6) = "TDKOPF" then myLine = myLine & FreeText &vbCrLf else myLine = ";;;;;;;;" & myLine & "" & FreeText&vbCrLf
 WriteCSV myLine 

Next





Sub WriteCSV(LogMessage)
Const ForAppending = 8
A = year(Now)
B=Month(Now)
if len(B) < 2  then b="0" & B end if
C = day (Now)
if len(C)<2 then C ="0" & C end if
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objLogFile = objFSO.OpenTextFile(A & B & C & ".csv" , ForAppending, TRUE)
objLogFile.Write(LogMessage)
End Sub