<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()


set oFoo = server.createobject("CorFrontEXLR.Footer")


Set oFoo.Request = Request
Set oFoo.Server = Server
Set oFoo.Response = Response
Set oFoo.Session = Session
Set oFoo.Ofv = Ofv

sHTML = oFoo.Generar

sHTML = replace(sHTML,"2.2","2.4")

Response.Write sHTML

%>
