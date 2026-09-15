<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<% 

ok=ofv.CheckUsuario()

set oApl = server.createobject("CorFrontSP.Apl")

Set oApl.Request = Request
Set oApl.Server = Server
Set oApl.Response = Response
Set oApl.Session = Session
Set oApl.Ofv = Ofv

sHTML = oApl.Generar



Response.Write sHTML

%>