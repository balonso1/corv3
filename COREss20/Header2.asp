<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()


  response.expires=0

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From Clienteplan"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)  

  Session("Cliente") = rs(1)
  Session("CliLogo") = "../COREss20/Empresa/" & rs(5)
  SCliente = rs(5)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = ofv.MenuHeader("#336699","background=Imagenes/topban.jpg topmargin=3 leftmargin=3 style='border-style:outset;border-width:1;' ")

  sHTML = sHTML & "<table width=100% marginwidth=1 cellspacing=0 cellpadding=0 border=0>"
  sHTML = sHTML & "<tr><td  width=2% align=center valign=middle>"
  sHTML = sHTML & "<img src=Empresa/" & SCliente & " border=0  align=center></td>"
  sHTML = sHTML & "<td valign=middle class=at width=100% >&nbsp;&nbsp;" & Session("Cliente")  & "</td></tr></table></body></html>"


response.write  sHTML

%>