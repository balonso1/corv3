<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  set oIni = server.createobject("CorSecurityTCdc.Inicio")
  Set oIni.Request = Request
  Set oIni.Server = Server
  Set oIni.Response = Response
  Set oIni.Session = Session
  Set oIni.Ofv = Ofv

  shtml = oIni.HabilitPass()

end if

response.write shtml

%>