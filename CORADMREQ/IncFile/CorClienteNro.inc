<%
function GTCLX()

    sh = datepart("h",date())
    sm = datepart("n",date())
    ss = datepart("s",date())

XCodobJ = sh & sm & ss

GThms = datediff("d",date(),cdate("12/12/2050")) & formatdatetime(now(),3)
GTses = session.sessionid
GTsecu = cstr(XCodobJ & GThms & GTses)

GTsecuf = ""

for GTi=1 to len(GTsecu)
GTcs = mid(GTsecu,GTi,1)
if GTcs => "0" and GTcs <= "9" then
   GTsecuf = GTsecuf & GTcs
end if
next
 
GTCLX = trim(GTsecuf)

end function
%>