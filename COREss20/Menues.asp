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
  Nombre = ofv.convertircar(request.querystring("Nombre"),"_"," ")
  sAplicacion = ofv.convertircar(request.querystring("Aplicacion"),"_"," ")
  vuelta = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover)
  vuelta = vuelta & "&Opcion=" & ID & "&Aplicacion=" & sAplicacion



  sHTML = ofv.FormHeader(session("FormN"))


  ofv.ObtenerAtributos Session("Form") ,""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cx =  ofv.conectar(ofv.strconn1)

  set cn = ofv.conectar(ofv.strconn0)

  strsql = "SELECT * From Atrm Where tipo = 'S'  And aplicacion = '" & sAplicacion & "'  order by alias"
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
     strsql = strsql & "And Tipo = 'S'  And aplicacion = '" & sAplicacion & "'  Order By Alias"
     set rsu = ofv.crearconsultaEx(StrSql,cn,1,parametros) 
     if rsu.eof then xusr = "N"
  else

      xgrupo = ID
      xusr = "N" 

  end if

     strsql = "SELECT * From AtrmXgrupo Where grupo = '" & xgrupo & "' "
     strsql = strsql & "And Tipo = 'S'  And aplicacion = '" & sAplicacion & "' Order By Alias"
     set rsg = ofv.crearconsultaEx(StrSql,cn,1,parametros) 
     if rsg.eof then xgrp = "N"

    scomK = "0" & chr(9) & "1"
    scomD = "NO" & chr(9) & "SI"


  celdas = ofv.GenCelda("","ttr colspan=4","center","","","","","SOLAPAS DE " & Nombre & "&nbsp;&nbsp;&nbsp;" & session("GRPUSRDSC"),"si")
  formulario = "<thead>" & ofv.GenRow("","","","","22","","",celdas,"si") & "</thead>"



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
       if rsg("Alias") = rs2("Alias") then
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
       if rsu("Alias") = rs2("Alias") then
             xestado = abs(cint(rsu("usrestado")))
             xcolor = "#6699cc"
          rsU.movenext
          if rsU.eof then xusr = "N"
        end if
     end if

    if bModificar then

         sInput = ofv.gencomboFX(sForm,"ESTADO",xestado,scomK,scomD," ")


   ''   sInput = "<font class=at><select onchange=" & sForm & ".submit() "
    ''  sInput = sInput & "name=ESTADO><option value=1>SI  "
    ''  if  xestado = 1 then
    ''    sInput = sInput & "<option value=0>NO  "
    ''  else
    ''    sInput = sInput & "<option value=0 selected>NO  "
    ''  end if
    ''  sInput = sInput & "</select></font>"

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
      fila = ofv.GenForm(sForm,"",sAccion,"","post","",celdas,"si")
    else
      fila = ofv.GenForm("","","","","post","",celdas,"si")
    end if

    if xestado = 1 or xestado = "SI" then
      sInput = ofv.generarinput("","Menues","submit","SubMenus","10","bt","")
      celdax = ofv.GenCelda("","","","","","","",sInput,"si")
      sAccion = "SubMenues.asp?Menu=" & rs2("Alias") & "&Opcion=" & ID & "&Aplicacion="
      sAccion = sAccion & sAplicacion & "&t=" & t & "&Nombre="
      sAccion = sAccion & ofv.convertircar(Nombre," ","_")
      Celdas = ofv.genform("SubMenus","",sAccion,"","post","",Celdax,"si")
    else
      celdas = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if
    formulario = formulario & ofv.GenRow("","","","","10","","",fila & celdas,"si")

    rs2.movenext
  loop
  celdas = ofv.GenCelda("","abc colspan=3","center","","","","","&nbsp;","si")
  fila = ofv.GenRow("","","","","","","",celdas,"si")

  sInput = ofv.generarinput("","Volver","submit","Volver","10","bt","")
  celdas = ofv.GenCelda("","abc colspan=3","center","","","","",sInput,"si")
  fila = fila & ofv.GenRow("","","","","","","",celdas,"si")
  sAccion = "Aplicaciones.asp?Opcion=" & ID & "&t=" & t
'  formulario = formulario & ofv.GenForm("Volver","",sAccion,"Contenido","post","",fila,"si")

  tabla = ofv.GenTabla("","","","80%","","","0","0","0","","",formulario,"si")

'''  sHTML = sHTML & "<center>" & tabla & "</center></body></html>"

  sHTML = sHTML & "<center>" & tabla & "</center>"

  Agregar = ""
  Buscar  = ""

  sAccion = "Aplicaciones.asp?Opcion=" & ID & "&t=" & t
  Celda = ofv.BotonVolver(sAccion)


  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)


  sHTML = sHTML & "</body></html>"


  if t <> "G" then call ofv.cerrarconsulta(rsU)
  call ofv.cerrarconsulta(rsG)
  call ofv.cerrarconsulta(rs2)
  call ofv.cerrarconn(cn)
  call ofv.cerrarconn(cx)
end if

response.write sHTML

%>