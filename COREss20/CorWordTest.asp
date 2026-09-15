<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%



ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0

  Dim oWord
  Dim sMsg

  Set oWord = CreateObject("CorWord.WordCreator")

  Set oWord.Request = Request
  Set oWord.Server = Server
  Set oWord.Response = Response
  Set oWord.Session = Session

  sMsg = oWord.Generar(252, "NMA", "Esquemas Normativos", "OBJMAIN")
  If sMsg <> "" Then response.write sMsg
  Set oWord = Nothing
end if

%>