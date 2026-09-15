<%@LCID = 11274%> 
<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

     CODSET = request.querystring("CODSET")
     OPT    = request.querystring("OPT")


   IF OPT = "SI" THEN

     set cn = ofv.conectar(ofv.strconn2)



     strsql = "DELETE FROM SYSDEFDSKTOP WHERE OBJID = " & CODSET
     call ofv.crearconsultaEx(StrSql,cn,1,parametros)

     strsql = "DELETE FROM SYSDEFDSKTOPDET WHERE CODSET = " & CODSET
     call ofv.crearconsultaEx(StrSql,cn,1,parametros)


    call ofv.cerrarconn(cn)


  Response.redirect "defescridet.asp"



   ELSE

   sHTML = ofv.FormHeader(session("FormN"))

   sHTML = sHTML & "<BR><BR><BR><CENTER>"



    Celda = ofv.gencelda("","vt COLSPAN=3 STYLE='BORDER-STYLE:RIDGE;BORDER-WIDTH:2;BORDER-COLOR:#F7AA22;' ","CENTER","","","","","<BIG>CONFIRMA LA ELIMINACION DE SET DE ESCRITORIO COD:" & CODSET & "</BIG>","si")


    Fila = ofv.genrow("","","","","","","",Celda,"si")

    Celda = ofv.gencelda("","vt","","40%","","","","&nbsp;","si")
    Celda = Celda & ofv.gencelda("","vt","","1%","","","","&nbsp;","si")
    Celda = Celda & ofv.gencelda("","vt","","40%","","","","&nbsp;","si")


    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

  sInput = ofv.GenerarInput("boton1","CONFIRMA","submit","","20","bt","")
  sAccion = "DEFESCRIDETDEL.asp?CODSET=" & CODSET & "&OPT=SI"
  Formulario = ofv.gencelda("","vt","RIGHT","40%","","","",sInput,"si")
  Formulario1 = ofv.genform("CONF","",sAccion,"","post","",Formulario,"si")

    Celda = Formulario1
    Celda = Celda & ofv.gencelda("","vt","","1%","","","","&nbsp;","si")


  sInput = ofv.GenerarInput("boton2","CANCELA","submit","","20","bt","")
  sAccion = "DEFESCRIDET.asp"
  Formulario = ofv.gencelda("","vt","","40%","","","",sInput,"si")
  Formulario2 = ofv.genform("CANC","",sAccion,"","post","",Formulario,"si")



    Celda = Celda & Formulario2


    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


    sHTML = sHTML & ofv.GenTabla("","aa style='filter:shadow(color=" & Session("APLSHW") & ",direction=125);' ","","60%","","0","","0","0","","",fila,"si")



   sHTML = sHTML & "</body></html>"

   Response.Write sHTML



   END IF



%>