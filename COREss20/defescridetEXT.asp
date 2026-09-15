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

  CODSET = request.querystring("CODSET")

  IF CODSET <> "" THEN
     SESSION("VSET") = CODSET
  ELSE
     CODSET = SESSION("VSET")
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

 '' sHTML = replace(ofv.FormHeader(session("FormN")),"<body ","<body onload='uno();' ")  & chr(13)
  shtml = shtml & "<script  language=javascript>" & chr(13)
  shtml = shtml & " var cnt = 0; " & chr(13)
  shtml = shtml & " function uno() {  " & chr(13)
  shtml = shtml & " if (document.all.odiv) { document.all.odiv.style.top=document.body.clientHeight - 30; " & chr(13)
  shtml = shtml & " } else {  if (cnt > 10) {} else { cnt++; " & chr(13)
  shtml = shtml & " setTimeout (" & chr(34) & "uno()" & chr(34) & ", 350); }}" & chr(13)
  shtml = shtml & " } " & chr(13)
  shtml = shtml & " </script>" & chr(13)

  'Encabezados
  Formulario = ofv.Encabezados("GRUPO DE CONFIGURACION","ABRIR")

  VDEF = 1

  Strsql = "Select * From SYSDEFDSKTOP WHERE OBJID = " & CODSET
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then 

     VDSC = RS(3)

  end if

  SIMAGE1 = "<B><FONT FACE=WINGDINGS SIZE=2 COLOR=#005500>&#0232;</FONT></B>"
  SIMAGE2 = "<B><FONT FACE=WINGDINGS SIZE=1 COLOR=#AA0000>l</FONT></B>"


  Strsql = "SELECT PRTY2 AS GRUPO, PRTY1 AS COD "
  Strsql = Strsql & " FROM SYSDEFDSKTOPDET"
  Strsql = Strsql & " WHERE PRTY5 = 'SI' AND CODSET=0"
  Strsql = Strsql & "  GROUP BY PRTY2, PRTY1"
  Strsql = Strsql & " ORDER BY PRTY2"


  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  Do While (Not rs.EOF)



      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")




    'Descripcion

    Celda = Celda & ofv.gencelda("","VT","","50%","","","",SIMAGE2 & "&nbsp;" & rs(0),"si")


  SIMAGE1 = "<B><FONT FACE=WINGDINGS SIZE=2 COLOR=#005500>&#0232;</FONT></B>"

    sinput = "<a href=defescridetEXTDET.asp?CODFAM=" & RS(1)
    sinput = sinput & " >" & SIMAGE1 & "</a>" 
    Celda = Celda & ofv.gencelda("","VT  ","center","","20%","","",sInput,"si")






    Fila = ofv.genrow("","","","","","","",Celda,"si")


    Formulario = Formulario & Fila

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



     sAccion = "DEFESCRIDET.asp"
     retornar = ofv.BotonVolver(sAccion)



  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,Retornar)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"





end if

Response.Write sHTML

%>