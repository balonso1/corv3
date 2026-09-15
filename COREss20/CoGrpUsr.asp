<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  ngrp = request.querystring("grupo")

  set cn = ofv.conectar(ofv.strconn0)

  Strsql = "Select * From grupos Where ID='" & ngrp & "' " 
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     snombre = rs(1)
  end if
  call ofv.cerrarconsulta(rs)

    shtml = ofv.Repbody(Ucase("Usuarios del Grupo:") & "&nbsp;" & snombre) & ofv.RepHeader2(Ucase("Usuarios del Grupo:") & "&nbsp;" & snombre,"600")

 ' Strsql = "Select ID,descripcion,nivel From usuarios Where grupo='" & ngrp & "' order by Descripcion " 
 ' set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
 
 call  getUsrIngroup(ngrp)
  if session("SGU_UIGuserid")<>"" then
 
 
 response.write session("SGU_UIGuserid")
 A_SGU_UIGuserid =split(session("SGU_UIGuserid"),chr(9))
 A_SGU_UIGdescri =split(session("SGU_UIGdescri") ,chr(9))
	                  
        Celda = ofv.gencelda("","vt","","2%","","#99aadd","",Ucase("Usuario"),"si")
        Celda = Celda & ofv.gencelda("","vt","","50%","","#99aadd","",Ucase("descripcion"),"si")
        Celda = Celda & ofv.gencelda("","vt colspan=2","","2%","","#99aadd","",Ucase("nivel"),"si")

        Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

     dim Aniv()
     redim preserve Aniv(0)
     Aniv(0) = "S/D"
     Strsql = "Select * from nivelconfiden order by NCcod " 
     set rx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     ix = 0
     if not rx.eof then
        do until rx.eof
           redim preserve Aniv(ix)
           Aniv(ix) = rx(1)
           ix = ix + 1
           rx.movenext
        loop
     end if
     call ofv.cerrarconsulta(rx)

     'do until rs.eof
     
     for a_i=0 to ubound(A_SGU_UIGdescri)
     
        Celda = ofv.gencelda("","ut","","5%","","","",A_SGU_UIGuserid(a_i),"si")
        Celda = Celda & ofv.gencelda("","ut","","60%","","","",A_SGU_UIGdescri(a_i),"si")
        sniv = ""
     '   for ix=0 to Ubound(Aniv)
     '       if cint(rs(2)) = ix then 
     '          sniv = Aniv(ix)
     '          exit for
     '       end if
     '   next
        'Celda = Celda & ofv.gencelda("","ut","","35%","","","",rs(2) & " - " & sniv,"si")

        Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

        sinput = "<hr size=1 color=#5588bb widht=100% >"
        Celda = ofv.gencelda("","vt colspan=3","","100%","","","",sinput,"si")
        Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

   '     rs.movenext
    ' loop
    next
  sHTML = sHTML & "<br>"
  sHTML = sHTML & ofv.GenTabla("","aa  border=0","","600","","0","0","1","0","","",fila,"si")
  else
   sHtml = sHtml & "<font face=tahoma,arial size=2 color=#336699>"   
   sHtml = sHtml & "<p align=center>No existen Usuarios asociados a este Grupo</p>"
   sHtml = sHtml & "</font>"
  end if
  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</center></body></html>"
end if

response.write sHTML

%>