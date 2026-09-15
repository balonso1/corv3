<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()


set oMen = server.createobject("CorFrontEXLR.MenuLR")

Set oMen.Request = Request
Set oMen.Server = Server
Set oMen.Response = Response
Set oMen.Session = Session
Set oMen.Ofv = Ofv

Session("Usuario") = Ucase(Session("Usuario"))

sHTML = replace(oMen.Generar,"<A ","<A class=tt ") 

Response.Write sHTML

%>