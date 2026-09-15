<%@LCID = 11274%> 
<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0



  sopcion = request.querystring("opcion")

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)



  if Session("EMPMUN") <> "S" then

     response.redirect "CorSitesUn.asp?UN=1"

  end if





    sImagen4 = "<B><FONT FACE=WingDINGS COLOR=#aa0000 SIZE=2>&#0232;</FONT></B>"





''  sHTML = ofv.FormHeader(session("FormN"))

  sHTML = replace(ofv.FormHeader(session("FormN")),"<body ","<body onload='uno();' ")  & chr(13)
  shtml = shtml & "<script  language=javascript>" & chr(13)
  shtml = shtml & " var cnt = 0; " & chr(13)
  shtml = shtml & " function uno() {  " & chr(13)
  shtml = shtml & " if (document.all.odiv) { document.all.odiv.style.top=document.body.clientHeight - 30; " & chr(13)
  shtml = shtml & " } else {  if (cnt > 10) {} else { cnt++; " & chr(13)
  shtml = shtml & " setTimeout (" & chr(34) & "uno()" & chr(34) & ", 350); }}" & chr(13)
  shtml = shtml & " } " & chr(13)
  shtml = shtml & " </script>" & chr(13)

  'Encabezados
  Formulario = ofv.Encabezados("cod","descripcion","Sites/Sucursales")


  Strsql = "Select * From Clienteplan order by ID"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
 
  posi = 0
  Do While (Not rs.EOF)

    sForm = "form" & posi

    'Eliminar

      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")


    'Descripcion
    sInput = rs(0) & "."
    Celda = Celda & ofv.gencelda("","vt","right","1%","","","",sInput ,"si")


    'Descripcion
    sInput = rs(1)
    Celda = Celda & ofv.gencelda("","vt","","50%","","","",sInput,"si")

    sinput = "<a class=ut href=CorSitesUn.asp?UN=" & rs(0)
    sinput = sinput & " >" & sImagen4 & "</a>" 
    t1 = ofv.gencelda("","aaa ","center","","20%","","",sInput,"si")

    Celda = Celda & t1

'    Celda = Celda & ofv.gencelda("","","","","","","","&nbsp;","si")
    Fila = ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
 
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1
    rs.Movenext
  Loop


  ftab = ofv.GenTabla("","","","100% ","","0","","0","0","","",Formulario,"si")

    Celda = ofv.GenCelda("","","center","","","","",ftab,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aa  ","","90%","","0","","0","0","","",filas,"si")




  'Botonera
  Agregar = ""

  Buscar = ""

  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>