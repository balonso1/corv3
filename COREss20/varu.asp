<%
user=Request.ServerVariables("LOGON_USER") 
user=mid(user,InStrRev(user,"\")+1,len(user))
%>
<form>

<input type=text readonly value=<%=user%>>
</form>