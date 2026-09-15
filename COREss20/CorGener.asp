<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0

  sHTML = ofv.FormHeader("Generacion de Descargas")

  'Encabezados

  vuelta = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover)

  set cn = ofv.conectar(ofv.strconn1)



  ofv.ObtenerAtributos request.querystring("seg"),""
  Session("Form") = request.querystring("seg")
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  StrSql = "Select * From CorCrpt order by CRPdesc"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  formulario = ""
  do until rs.eof 

    sInput = "<img src=iconos/ICorAbre.jpg border=0 align=center>"
    celdas = ofv.GenCelda("","cc","","5%","","","",sinput,"si")
    sInput = rs(2)
    celdas = celdas & ofv.GenCelda("","cc","90%","","","","",sInput,"si")

    sInput = ofv.generarinput("","Abrir","submit","Menus","10","bt","")
    celdax = ofv.GenCelda("","cc","","5%","","","",sInput,"si")
    sAccion = "CorGener2.asp?MEN=" & rs(1)
    celdas = celdas & ofv.genform("Menus","",sAccion,"","post","",Celdax,"si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")



      formulario = formulario & filas
 

    rs.movenext
  loop

  sHTML = sHTML  & ofv.GenTabla("","","","80%","","0","","0","0","","",formulario,"si")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML

%>