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

  EMP = request.querystring("EMP")

  IF EMP <> "" THEN
     SESSION("VEMP") = EMP
  ELSE
     EMP = SESSION("VEMP")
  END IF 

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

    sImagen4 = "<B><FONT FACE=WingDINGS COLOR=#aa0000 SIZE=2>&#0232;</FONT></B>"

  sHTML = ofv.FormHeader(session("FormN"))


  shtml = shtml & "<script  language=javascript>" & chr(13)
  shtml = shtml & " var cnt = 0; " & chr(13)
  shtml = shtml & " function uno() {  " & chr(13)
  shtml = shtml & " if (document.all.odiv) { document.all.odiv.style.top=document.body.clientHeight - 30; " & chr(13)
  shtml = shtml & " } else {  if (cnt > 10) {} else { cnt++; " & chr(13)
  shtml = shtml & " setTimeout (" & chr(34) & "uno()" & chr(34) & ", 350); }}" & chr(13)
  shtml = shtml & " } " & chr(13)
  shtml = shtml & " </script>" & chr(13)

  'Encabezados
  Formulario = ofv.Encabezados("cod","descripcion","TOMA DEFAULT","ACTIVO","DETALLE")

  SIMAGE1 = "<B><FONT FACE=WINGDINGS SIZE=2 COLOR=#005500>&#0232;</FONT></B>"



  VDEF = 1

  Strsql = "Select CODSET From Clienteplan WHERE ID = " & EMP
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then 

     VDEF = RS(0)

  end if

  VIENE = "N"


  Strsql = "Select * From SYSDEFDSKTOP WHERE OBJUN = " & EMP & " order by OBJID"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    'Eliminar
    if bSuprimir and rs(2) <> "SI" AND RS(0) <> VDEF then
      sHref = "DEFESCRIDETDEL.ASP?CODSET=" & RS(0) & "&OPT=VA"
      Celda = ofv.BotonEliminar(rs(3),sHref)
     else
      Celda = ofv.GenCelda("","ut","","","","","","&nbsp;","si")
    end if

    'Descripcion
    sInput = rs(0) & "."
    Celda = Celda & ofv.gencelda("","ut","right","1%","","","",sInput ,"si")


    'Descripcion
    sInput = ofv.GenerarInput(rs(3).name,rs(3),"text",sForm,"80","dt","")
    Celda = Celda & ofv.gencelda("","","","50%","","","",sInput,"si")

    'DEFAULT
  xtipo = "SI"
  if rs(2) <> "SI" then
      xtipo = "NO"
  end if

    Celda = Celda & ofv.gencelda("","ut","","5%","","","",xtipo,"si")


    'ACTIVO
  xtipo = "<b><font color=#990000 >SI</font></b>"
  if rs(0) <> VDEF then
      xtipo = "NO"
  end if

    Celda = Celda & ofv.gencelda("","ut","","5%","","","",xtipo,"si")


  if rs(2) <> "SI" then

    sinput = "<a href=defescridetEXT.asp?CODSET=" & rs(0)
    sinput = sinput & " >" & SIMAGE1 & "</a>" 
    Celda = Celda & ofv.gencelda("","VT  ","center","","20%","","",sInput,"si")

  else
    Celda = Celda & ofv.gencelda("","vt","center","1%","","","","&nbsp;","si")
  end if




    Fila = ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
     SACCION = "asp/actualizard.asp?fname=" 
     SACCION = SACCION & rs(0) & "&tabla=SYSDEFDSKTOP&volver=../DEFESCRIDET.asp"
    end if
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
   Retornar = ""



     sAccion = "DEFESCRI.asp"
     retornar = ofv.BotonVolver(sAccion)

     if bAgregar then
        sAccion = "defescridetadd.asp"
        Agregar = ofv.BotonAgregar(sAccion)
     end if



     str1 = "Select DSCSET From SYSDEFDSKTOP WHERE OBJUN = " & EMP & " order by OBJID "    ' orden pagina
     str2 = "Select DSCSET From SYSDEFDSKTOP WHERE OBJUN = " & EMP & " order by DSCSET "    ' orden alfabetico
     Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))


  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,retornar)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"





end if

Response.Write sHTML

%>