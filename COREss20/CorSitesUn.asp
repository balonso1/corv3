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

  CantRegMostrarx = 10
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  sUN = request.querystring("UN")

  IF sUN <> "" THEN

     SESSION("SUN") = SUN

  ELSE

    SUN = SESSION("SUN")

  END IF

  sopcion = request.querystring("opcion")

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  if sopcion = "SI" then


        STRSQL = "INSERT INTO SITES (DESCRI,SITEUN) VALUES ("  
        STRSQL = STRSQL & "'NUEVO SITES/SUCURSAL/FILIAL','" & SUN & "')" 
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)



  end if


  sHTML = ofv.FormHeader(session("FormN"))

'  sHTML = replace(ofv.FormHeader(session("FormN")),"<body ","<body onload='uno();' ")  & chr(13)
'  shtml = shtml & "<script  language=javascript>" & chr(13)
'  shtml = shtml & " var cnt = 0; " & chr(13)
'  shtml = shtml & " function uno() {  " & chr(13)
'  shtml = shtml & " if (document.all.odiv) { document.all.odiv.style.top=document.body.clientHeight - 30; " & chr(13)
'  shtml = shtml & " } else {  if (cnt > 10) {} else { cnt++; " & chr(13)
'  shtml = shtml & " setTimeout (" & chr(34) & "uno()" & chr(34) & ", 350); }}" & chr(13)
'  shtml = shtml & " } " & chr(13)
'  shtml = shtml & " </script>" & chr(13)

  'Encabezados

  Strsql = "Select * From CLIENTEPLAN  WHERE ID = " & SUN 
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then 

     SNOMEMPRE = RS(1)

  END IF

    sImagen4 = "<B><FONT FACE=WEBDINGS COLOR=#aa0000 SIZE=2>4</FONT></B>"

  Celda = ofv.gencelda("","VT","","","","","","<BIG>" & sImagen4 & " " & SNOMEMPRE & "</BIG>","si")
  Fila = ofv.genrow("","","","","","","",Celda,"si")
  ftabH = ofv.GenTabla("","","","100% ","","0","","0","0","","",Fila,"si") & "<BR>"


  Formulario = ofv.Encabezados("cod","descripcion","NUM")


  Strsql = "Select * From SITES  WHERE SITEUN = " & SUN & " order by OBJID"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    'Eliminar
    if bSuprimir  then
      sHref = "asp/EliminarRegistro2.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=SITES&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(1),sHref)
     else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'Descripcion
    sInput = rs(0) & "."
    Celda = Celda & ofv.gencelda("","vt","right","1%","","","",sInput ,"si")


    'Descripcion
    sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"100","dt","")
    Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

    'CODIGO INTERNO
    sInput = ofv.GenerarInput(rs(3).name,rs(3),"text",sForm,"10","dt","")
    Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")



    Fila = ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
     SACCION = "asp/actualizard.asp?fname=" 
     SACCION = SACCION & rs(0) & "&tabla=SITES&volver=../CORSITESUN.asp"
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1
    rs.Movenext
  Loop


  ftab = ofv.GenTabla("","","","100% ","","0","","0","0","","",Formulario,"si")

    Celda = ofv.GenCelda("","","center","","","","",ftabh,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    Celda = ofv.GenCelda("","","center","","","","",ftab,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aa  ","","90%","","0","","0","0","","",filas,"si")


  retornar = ""
  if Session("EMPMUN") = "S" then
 '''    sInput = ofv.generarinput("","Volver","submit","Volver","10","bt","")
 '''    sInput = ofv.gencelda("","","center","","","","",sInput,"si")
     sAccion = "CorSites.asp"
 '''    retornar = ofv.genform("Volver","",sAccion,"","post","",sInput,"si")
     retornar = ofv.BotonVolver(sAccion)

  end if


  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "CORSITESUN.asp?Opcion=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select DESCRI From SITES WHERE SITEUN = " & SUN & " order by OBJID "    ' orden pagina
  str2 = "Select DESCRI From SITES WHERE SITEUN = " & SUN & " order by DESCRI "    ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))


    



  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,retornar)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>