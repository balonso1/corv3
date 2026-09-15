<%@LCID = 11274%> 
<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else

 if session("empmun") = "S" then
  response.expires=0

  CantRegMostrarx = 10
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  sopcion = request.querystring("opcion")

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  if sopcion = "SI" then


        STRSQL = "INSERT INTO CLIENTEPLAN (CLIENTE,INLOGO,SEGURIDAD,TIPO) VALUES ("  
        STRSQL = STRSQL & "'NUEVA EMPRESA','N',1,'2')" 
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)



  end if

    sImagen4 = "<B><FONT FACE=WingDINGS COLOR=#aa0000 SIZE=2>&#0232;</FONT></B>"



  sHTML = ofv.FormHeader(session("FormN"))

''  sHTML = replace(ofv.FormHeader(session("FormN")),"<body ","<body onload='uno();' ")  & chr(13)
  shtml = shtml & "<script  language=javascript>" & chr(13)
  shtml = shtml & " var cnt = 0; " & chr(13)
  shtml = shtml & " function uno() {  " & chr(13)
  shtml = shtml & " if (document.all.odiv) { document.all.odiv.style.top=document.body.clientHeight - 30; " & chr(13)
  shtml = shtml & " } else {  if (cnt > 10) {} else { cnt++; " & chr(13)
  shtml = shtml & " setTimeout (" & chr(34) & "uno()" & chr(34) & ", 350); }}" & chr(13)
  shtml = shtml & " } " & chr(13)
  shtml = shtml & " </script>" & chr(13)

  'Encabezados
  Formulario = ofv.Encabezados("cod","descripcion","tipo","propiedades")


  Strsql = "Select * From Clienteplan order by ID"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    'Eliminar
    if bSuprimir and rs(6) <> "1" then
      sHref = "asp/EliminarRegistro2.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=Clienteplan&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(1),sHref)
     else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'Descripcion
    sInput = rs(0) & "."
    Celda = Celda & ofv.gencelda("","vt","right","1%","","","",sInput ,"si")


    'Descripcion
    sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"80","dt","")
    Celda = Celda & ofv.gencelda("","","","50%","","","",sInput,"si")

    'fecha
  xtipo = "UNIDAD&nbsp;DE&nbsp;NEGOCIOS"
  if not isnull(rs(6)) then
     if rs(6) = "1" then xtipo = "GRUPO"
  end if


    Celda = Celda & ofv.gencelda("","vt","","20%","","","",xtipo,"si")


  ''  sinput = "<a class=tt href=DatosdelClienteProp.asp?EMP=" & rs(0)
  ''  sinput = sinput & " ><b>propiedades</b></a>" 
  ''  t1 = ofv.gencelda("","tt style='border-style:outset;border-width:2;' ","center","","20%","","",sInput,"si")

    sinput = "<a class=ut href=DatosdelClienteProp.asp?EMP=" & rs(0)
    sinput = sinput & " >" & sImagen4 & "</a>" 
    t1 = ofv.gencelda("","aaa ","center","","20%","","",sInput,"si")

    Celda = Celda & t1

'    Celda = Celda & ofv.gencelda("","","","","","","","&nbsp;","si")
    Fila = ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
     SACCION = "asp/actualizard.asp?fname=" 
     SACCION = SACCION & rs(0) & "&tabla=Clienteplan&volver=../DatosDelCliente.asp"
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1
    rs.Movenext
  Loop


  ftab = ofv.GenTabla("","","","100% ","","0","","0","0","","",Formulario,"si")

    Celda = ofv.GenCelda("","","center","","","","",ftab,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aa  ","","90%","","0","","0","0","","",filas,"si")



 '' sHTML = sHTML & ofv.GenTabla("","","","100%","","0","","0","0","","",Formulario,"si")

  'Botonera
   Agregar = ""
   Buscar = ""

  IF SESSION("EMPMUN") = "S"  THEN
     if bAgregar then
        sAccion = "DatosdelCliente.asp?Opcion=SI"
        Agregar = ofv.BotonAgregar(sAccion)
     end if

     str1 = "Select cliente From Clienteplan order by ID "    ' orden pagina
     str2 = "Select cliente From Clienteplan order by cliente "    ' orden alfabetico
     Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

  end if




  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"

 else
  response.redirect "DatosdelClienteProp.asp?EMP=1"
 end if


end if

Response.Write sHTML

%>