<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()


set oSub = server.createobject("CorFrontEXLR.Submenus")

Set oSub.Request = Request
Set oSub.Server = Server
Set oSub.Response = Response
Set oSub.Session = Session
Set oSub.Ofv = Ofv

sHTML = oSub.Generar

Response.Write sHTML

%>
