<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%



ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0



  sHTML = ofv.FormHeader(session("FormN"))


  ID = request.querystring("Opcion")
  t = request.querystring("t")
  vuelta = ofv.vueltaasp("../",CantRegAMover,CantRegAMostrar)
  vuelta = vuelta & "&Opcion=" & ID

  set cn = ofv.conectar(ofv.strconn1)

  Strsql = "Select * From CorAplic where enable<>0 Order By DescApli"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  redim preserve aApli(0)
  redim preserve aAplicd(0)
  ix = 0
  do until rs.eof
    redim preserve aApli(ix)
    redim preserve aAplicd(ix)
    aApli(ix) = ofv.convertircar(rs(2),"*"," ") & rs(4)
    aAplicd(ix) = rs(1)
    ix = ix + 1
    rs.movenext

  loop
  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
'xstrconn0 =  ofv.strconn0

  ofv.ObtenerAtributos Session("Form") ,""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar



  set cn = ofv.conectar(ofv.strconn0)

  strsql = "SELECT * From Atrm Where tipo = 'A' order by atributo"
  set rs0 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  xgrupo = " "
  xgrp = "S"
  xusr = "S" 

  if t <> "G" then
     strsql = "SELECT grupo,descripcion From usuarios Where ID = '" & ID & "' "
     set rsu = ofv.crearconsultaEx(StrSql,cn,1,parametros) 
     if not rsu.eof then
            xgrupo = rsu(0) 
            session("GRPUSRDSC") = "USUARIO:&nbsp;" & id & "-" & rsu(1)
     end if 
     call ofv.cerrarconsulta(rsu)

     strsql = "SELECT * From AtrmXusuario Where usuario = '" & ID & "' "
     strsql = strsql & "And Tipo = 'A' Order By atributo"
     set rsu = ofv.crearconsultaEx(StrSql,cn,1,parametros) 
     if rsu.eof then xusr = "N"
  else

      xgrupo = ID
      xusr = "N" 

     strsql = "SELECT id,descripcion From grupos Where ID = '" & ID & "' "
     set rsu = ofv.crearconsultaEx(StrSql,cn,1,parametros) 
     if not rsu.eof then
            session("GRPUSRDSC") = "GRUPO:&nbsp;" & id & "-" & rsu(1)
     end if 
     call ofv.cerrarconsulta(rsu)


  end if


     strsql = "SELECT * From AtrmXgrupo Where grupo = '" & xgrupo & "' "
     strsql = strsql & "And Tipo = 'A' Order By atributo"
     set rsg = ofv.crearconsultaEx(StrSql,cn,1,parametros) 
     if rsg.eof then xgrp = "N"


  celdas = ofv.GenCelda("","ttr colspan=4","center","","","","","APLICACIONES&nbsp;&nbsp;&nbsp;" & session("GRPUSRDSC"),"si")
  formulario = "<thead>" & ofv.GenRow("","","","","22","","",celdas,"si") & "</thead>"



    scomK = "0" & chr(9) & "1"
    scomD = "NO" & chr(9) & "SI"




  ix = 0
  ixx = 0
  do until rs0.eof
     xestado = 1
     xestadoD = 1
     xdef = "D"
     xcolor = ""
     for ix=0 to  Ubound(aApli)
         if rs0("Aplicacion") = aAplicd(ix) then
            ixx = ix
            exit for
         end if  
     next   
    if rs0("Aplicacion") = aAplicd(ixx) then
      ix = ixx
'       sAplicacion = rs0("Aplicacion") & " - " & UCase(aApli(ix))
      sAplicacion = UCase(aApli(ix))
    xdesatr = " <font color=#993333 face=tahoma size=1>(" & rs0("atributo") & ")</font>"
    celdas = ofv.GenCelda("","ut","","","","","",sAplicacion & " - " & xdesatr,"si")

    if xgrp <> "N" then
       if rsg("atributo") = rs0("atributo") then
          if not rsg("grpestado") then 
             xestado = 0  
             xestadoD = 0  
             xdef = "G"
             xcolor = "#993333"
          end if
          rsg.movenext
          if rsg.eof then xgrp = "N"
        end if
     end if

    if xusr <> "N" then
       if rsu("atributo") = rs0("atributo") then
             xestado = abs(cint(rsu("usrestado")))
             xcolor = "#6699cc"
          rsU.movenext
          if rsU.eof then xusr = "N"
        end if
     end if

 
       

    if bModificar then

         sInput = ofv.gencomboFX("Aplicaciones" & rs0("ID") & ix,"ESTADO",xestado,scomK,scomD," ")




    else
      if xestado = 1 then
         xestado = "SI"
      else
         xestado = "NO"
      end if   
      sInput = ofv.GenerarInput("",xestado,"readonly","","2","dt","")
    end if
      sInput2 = ofv.GenerarInput("default",xdef,"hidden","","2","dt","")
      sInput1 = ofv.GenerarInput("ESTADOD",xestadod,"hidden","","2","dt","")
    celdas = celdas & ofv.GenCelda("","","center","","",xcolor,"",sInput + sInput1 + sinput2,"si")

    if bModificar then
      sAccion = "asp/ActualizarAtributo.asp?volver=" & vuelta & "&ID=" & rs0("ID")
      sAccion = sAccion & "&GrupoUsuario=" & ID & "&Tipo=Menu&t=" & t
      fila = ofv.GenForm("Aplicaciones" & rs0("ID") & ix,"",sAccion,"","post","",celdas,"si")
    else
      fila = ofv.GenForm("","","","","post","",celdas,"si")
    end if

    if xestado = 1 or xestado = "SI" then
      sInput = ofv.generarinput("","Solapas","submit","Menus","10","bt","")
      celdax = ofv.GenCelda("","","","","","","",sInput,"si")
      sAccion = "Menues.asp?Aplicacion=" & ofv.convertircar(rs0("Aplicacion")," ","_")
      sAccion = sAccion & "&Opcion=" & ID & "&Nombre=" & ofv.convertircar(sAplicacion," ","_")
      sAccion = sAccion & "&t=" & t
      Celdas = ofv.genform("Menus","",sAccion,"","post","",Celdax,"si")
    else
      celdas = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if
    formulario = formulario & ofv.GenRow("","","","","10","","",fila & celdas,"si")

  end if
    rs0.movenext

  loop
  tabla = ofv.GenTabla("","","","80%","","0","0","0","0","","",formulario,"si")

  sHTML = sHTML & "<center>" & tabla & "</center>"

  Agregar = ""
  Buscar  = ""

  if  t = "G" then
      sAccion = "gruposx.asp?CantRegMover=0&cantregamover=" & session("CRAM")
  else
      sAccion = "usuariosx.asp"
  end if
  Celda = ofv.BotonVolver(sAccion)


  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)


  sHTML = sHTML & "</body></html>"

  if t <> "G" then call ofv.cerrarconsulta(rsU)
  call ofv.cerrarconsulta(rsG)

  call ofv.cerrarconsulta(rs0)
  call ofv.cerrarconn(cn)
end if

response.write sHTML

%>