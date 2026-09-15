<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  ID = request.querystring("Opcion")
  t = request.querystring("t")
  sMenu = ofv.convertircar(request.querystring("Menu"),"_"," ")
  Nombre = ofv.convertircar(request.querystring("Nombre"),"_"," ")

  sAplicacion = request.querystring("Aplicacion")
  vuelta = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover)
  vuelta = vuelta & "&Opcion=" & ID & "&Aplicacion=" & sAplicacion



  sHTML = ofv.FormHeader(session("FormN"))

  ofv.ObtenerAtributos Session("Form") ,""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cx =  ofv.conectar(ofv.strconn1)
  smenuD = smenu
  strsql = "SELECT Menu From Menux Where Menuatr = '" & smenu & "'  And mcodapli = '"
  strsql = strsql & sAplicacion & "' "
  set rx = ofv.crearconsultaEx(StrSql,cx,1,parametros)
  if not rx.eof then smenuD = Ucase(rx(0))
  call ofv.cerrarconsulta(rx)

  set cn = ofv.conectar(ofv.strconn0)

  strsql = "SELECT * From Atrm Where tipo = 'M'  And aplicacion = '"
  strsql = strsql & sAplicacion & "' and depen = '" & sMenu & "'  order by atributo"
  set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  xgrupo = " "
  xgrp = "S"
  xusr = "S" 

  if t <> "G" then
     strsql = "SELECT grupo From usuarios Where ID = '" & ID & "' "
     set rsu = ofv.crearconsultaEx(StrSql,cn,1,parametros) 
     if not rsu.eof then xgrupo = rsu(0) 
     call ofv.cerrarconsulta(rsu)

     strsql = "SELECT * From AtrmXusuario Where usuario = '" & ID & "' "
     strsql = strsql & "And Tipo = 'M'  and depen = '" & sMenu & "'  And aplicacion = '" & sAplicacion & "'  Order By atributo"
     set rsu = ofv.crearconsultaEx(StrSql,cn,1,parametros) 
     if rsu.eof then xusr = "N"
  else

      xgrupo = ID
      xusr = "N" 

  end if

     strsql = "SELECT * From AtrmXgrupo Where grupo = '" & xgrupo & "' "
     strsql = strsql & "And Tipo = 'M'  and depen = '" & sMenu & "'  And aplicacion = '" & sAplicacion & "' Order By atributo"
     set rsg = ofv.crearconsultaEx(StrSql,cn,1,parametros) 
     if rsg.eof then xgrp = "N"

  celdas = ofv.GenCelda("","ttr colspan=4","center","","","","","MENUES DE " & sMenuD & "&nbsp;&nbsp;&nbsp;" & session("GRPUSRDSC"),"si")
  formulario = "<thead>" & ofv.GenRow("","","","","22","","",celdas,"si") & "</thead>"

    scomK = "0" & chr(9) & "1"
    scomD = "NO" & chr(9) & "SI"


  do until rs2.eof
     xestado = 1
     xestadoD = 1
     xdef = "D"
     xcolor = ""
    sForm = "AtributosGrupos" & rs2("ID")
      xdescri = ""
      xdesatr = " <font color=#993333 face=tahoma size=1>(" & rs2("Alias") & ")</font>"
      strsql = "select menu from menux where menuatr = '" & trim(rs2("Alias")) & "' "
      set rsm = ofv.crearconsultaEx(StrSql,cx,1,parametros)
      if not rsm.eof then
         xdescri = Ucase(rsm(0)) & " - "
      end if 
      call ofv.cerrarconsulta(rsm)


    celdas = ofv.GenCelda("","ut","","","","","",xdescri & xdesatr,"si")

    if xgrp <> "N" then
       if rsg("atributo") = rs2("atributo") then
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
       if rsu("atributo") = rs2("atributo") then
             xestado = abs(cint(rsu("usrestado")))
             xcolor = "#6699cc"
          rsU.movenext
          if rsU.eof then xusr = "N"
        end if
     end if


    if bModificar then

  '  scomK = "0" & chr(9) & "1"
  '  scomD = "NO" & chr(9) & "SI"

         sInput = ofv.gencomboFX(sForm,"ESTADO",xestado,scomK,scomD," ")

  '    sInput = "<font class=at><select onchange=" & sForm & ".submit() "
  '    sInput = sInput & "name=ESTADO><option value=1>SI  "
  '    if  xestado = 1 then
  '      sInput = sInput & "<option value=0>NO  "
  '    else
  '      sInput = sInput & "<option value=0 selected>NO  "
  '    end if
      sInput = sInput & "</select></font>"
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
      sAccion = "asp/ActualizarAtributo.asp?volver=" & vuelta & "&ID=" & rs2("ID")
      sAccion = sAccion & "&GrupoUsuario=" & ID & "&Tipo=Menu&t=" & t & "&Nombre="
      sAccion = sAccion & ofv.convertircar(Nombre," ","_")
      sAccion = sAccion & "&menu=" & ofv.convertircar(smenu," ","_")
      fila = ofv.GenForm(sForm,"",sAccion,"","post","",celdas,"si")
    else
      fila = ofv.GenForm("","","","","post","",celdas,"si")
    end if

    if xestado = 1 or xestado = "SI" then
      Menu = "'Menu " & rs2("Alias") & "'"


      strsql = "SELECT * From Atrm Where tipo = 'M'  And aplicacion = '"
      strsql = strsql & sAplicacion & "' and depen = '" & rs2("Alias") & "'  order by alias"
      set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if not rs3.eof then
        sInput = ofv.generarinput("","Submenues","submit","Submenues","10","bt","")
        celdax = ofv.GenCelda("","","","","","","",sInput,"si")
        sAccion = "SubMenues.asp?Menu=" & ofv.convertircar(rs2("Alias")," ","_") & "&Opcion="
        sAccion = sAccion & ID & "&t=" & t & "&Aplicacion=" & ofv.convertircar(sAplicacion," ","_") & "&Nombre=" & ofv.convertircar(Nombre," ","_")
        Celdas = ofv.genform("SubMenus","",sAccion,"","post","",Celdax,"si")
        formulario = formulario & ofv.GenRow("","","","","10","","",fila & celdas,"si")
      else
        sInput = ofv.generarinput("","Pantallas","submit","Pantallas","10","bt","")
        celdax = ofv.GenCelda("","","","","","","",sInput,"si")
        sAccion = "Pantallas.asp?Menu=" & ofv.convertircar(rs2("Alias")," ","_") & "&Opcion="
        sAccion = sAccion & ID & "&t=" & t & "&Aplicacion=" & ofv.convertircar(sAplicacion," ","_")
        sAccion = sAccion & "&Submenu=" & ofv.convertircar(sMenu," ","_") & "&Nombre=" & ofv.convertircar(Nombre," ","_")
        Celdas = ofv.genform("Pantallas","",sAccion,"","post","",Celdax,"si")
        formulario = formulario & ofv.GenRow("","","","","10","","",fila & celdas,"si")
      end if
      call ofv.cerrarconsulta(rs3)
    else
      celdas = ofv.GenCelda("","","","","","","","&nbsp;","si")
      formulario = formulario & ofv.GenRow("","","","","10","","",fila & celdas,"si")
    end if
    rs2.movenext
  loop
  celdas = ofv.GenCelda("","abc colspan=3","center","","","","","&nbsp;","si")
  fila = ofv.GenRow("","","","","","","",celdas,"si")




  strsql = "SELECT * From Atrm Where tipo = 'M'  And aplicacion = '"
  strsql = strsql & sAplicacion & "' and alias = '" & smenu & "'  order by alias"
  set rs4 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs4.eof  then 



     svmenu = rs4("depen")  
     svuelve = "Submenues.asp?Menu=" & svmenu & "&"
  else
     svuelve = "Menues.asp?"
  end if 

  call ofv.cerrarconsulta(rs4)

  sInput = ofv.generarinput("","Volver","submit","Volver","10","bt","")
  celdas = ofv.GenCelda("","abc colspan=3","center","","","","",sInput,"si")
  fila = fila & ofv.GenRow("","","","","","","",celdas,"si")
  accionform = svuelve & "Opcion=" & ID & "&t=" & t 
  accionform =  accionform & "&Aplicacion=" & ofv.convertircar(sAplicacion," ","_")
  accionform =  accionform & "&Nombre=" & ofv.convertircar(Nombre," ","_")
 ' formulario = formulario & ofv.GenForm("Volver","",accionform,"Contenido","post","",fila,"si")

  tabla = ofv.GenTabla("","","","80%","","","0","0","0","","",formulario,"si")

''  sHTML = sHTML & "<center>" & tabla & "</center></body></html>"

  sHTML = sHTML & "<center>" & tabla & "</center>"

  Agregar = ""
  Buscar  = ""

  sAccion = accionform
  Celda = ofv.BotonVolver(sAccion)


  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)


  sHTML = sHTML & "</body></html>"


  if t <> "G" then call ofv.cerrarconsulta(rsU)
  call ofv.cerrarconsulta(rsG)
  call ofv.cerrarconsulta(rs2)
  call ofv.cerrarconn(cn)
end if

response.write sHTML

%>