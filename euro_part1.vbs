
Set doc = CreateObject("MSXML2.DOMDocument")

doc.Load("C:\Users\aw\Desktop\Neuer Ordner\euro_part\InFileBeispil.xml")

If doc.parseError <> 0 Then
  WScript.Echo doc.parseError.reason
  WScript.Quit 1
End If

Set nodes = doc.selectNodes("/Invoice/Delivery")



For Each node In nodes
 WriteCSV node.nodeName & ";" & node.text
  
Next

Sub WriteCSV(LogMessage)
Const ForAppending = 8
A = year(Now)
B=Month(Now)
if len(B) < 2  then b="0" & B end if
C = day (Now)
if len(C)<2 then C ="0" & C end if
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objLogFile = objFSO.OpenTextFile("" & A & B & C & ".txt" , ForAppending, TRUE)
objLogFile.Write(LogMessage)
End Sub