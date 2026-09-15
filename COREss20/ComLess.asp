<%

Dim oObj
dim oComLess 
Dim sStr
Dim nMemberID
Dim vRes

session("obj") = "obj"

Set oDllLoader = CreateObject("ComLess.DllLoader")
set oDllLoader.session = session
set oDllLoader.response = response

set oobj = oDllLoader.CreateDllObject("C:/Inetpub/wwwroot/Coressdll/Corwebserver.dll","FuncVarias")
set tobj = session("obj")
'set oobj.session
response.write "CLSID:" & oDllLoader.ClassID & "<br>" & "<br>"
set tobj.session = session
set tobj.response = response
set tobj.request = request
set tobj.server = server
tobj.uv = "SI"
tobj.browser()
response.write "xxx" & tobj.aagregar & " 2 " & tobj.aconsultar & " 3 " & tobj.explorador & "<br>"
response.write tobj.formheader("prueba de comless")

'Set oCallDispatch = CreateObject("ComLess.CallDispatch")
'nMemberID = &H60030001
'nMemberID = oCallDispatch.GetMemberID(oObj,"Version")
'response.write "MemberID:" & nMemberID & "<br>" & "<br>"
'vRes = oCallDispatch.CallInvoke(oObj, nMemberID, 1)
'response.write "About: " & vRes & "<br>"

oDllLoader.UnloadDll

%>