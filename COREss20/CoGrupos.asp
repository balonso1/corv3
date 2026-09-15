<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

    shtml = ofv.Repbody(Ucase("Grupos Habilitados:")) & ofv.RepHeader2(Ucase("Grupos Habilitados:"),"600")

  set cn = ofv.conectar(ofv.strconn0)
  Strsql = "Select * From grupos order by Descripcion " 
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     
        Celda = ofv.gencelda("","vt","","2%","","#99aadd","",Ucase("Grupo"),"si")
        Celda = Celda & ofv.gencelda("","vt","","50%","","#99aadd","",Ucase("descripcion"),"si")
        Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

     do until rs.eof
        Celda = ofv.gencelda("","ut","","5%","","","",rs(0),"si")
        Celda = Celda & ofv.gencelda("","ut","","60%","","","",rs(1),"si")
        Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

        sinput = "<hr size=1 color=#5588bb widht=100% >"
        Celda = ofv.gencelda("","vt colspan=2","","100%","","","",sinput,"si")
        Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

        rs.movenext
     loop
  sHTML = sHTML & "<br>"
  sHTML = sHTML & ofv.GenTabla("","aa  border=0","","600","","0","0","1","0","","",fila,"si")
  else
   sHtml = sHtml & "<font face=tahoma,arial size=2 color=#336699>"   
   sHtml = sHtml & "<p align=center>No existen Grupos habilitados</p>"
   sHtml = sHtml & "</font>"
  end if
  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</center></body></html>"
end if

response.write sHTML

%>