
Dim doc 
Dim xmlString 
Dim nodes
'xmlString = "<?xml version='1.0'?><Result><PersonID>991166187</PersonID><AddressID>1303836</AddressID></Result>"
xmlString = "c:\Users\weiss\Desktop\123.xml"

Set doc = CreateObject("MSXML2.DOMDocument")
'Load the XML file
doc.Load(xmlString)
If Doc.parseError.errorCode <> 0 Then
   Dim myErr
   Set myErr = Doc.parseError
   MsgBox("You have error " & myErr.reason)
Else

   Set objNodeList = Doc.getElementsByTagName("ID")
   For i = 0 To (objNodeList.length - 1)
      MsgBox (objNodeList.Item(i).xml)
   Next
End If