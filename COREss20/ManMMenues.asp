<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  Nombre = ofv.convertircar(request.querystring("Nombre"),"_"," ")
  sAplicacion = request.querystring("Apl")
  vuelta = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover)


'  if ofv.explorador = "MSIE" then
'    sHTML = ofv.MenuHeader("#cccccc","")
'  else
'    sHTML = ofv.MenuHeader("#ccbbaa","")
'  end if

  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))
  ofv.ObtenerAtributos Session("Form") ,""

  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn1)

  celdas = ofv.GenCelda("","tt colspan=4","center","","","","","Administracion de Menues","si")
  formulario = "<thead>" & ofv.GenRow("","","","","10","","",celdas,"si")
  celdas = ofv.GenCelda("","tt colspan=4","center","","","","","Solapas de " & Nombre,"si")
  formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si")
  celdas = ofv.GenCelda("","xx colspan=4","center","","","","","&nbsp;","si")
  formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","tt ","center","","","","","Orden","si")
  celdas = celdas & ofv.GenCelda("","tt ","center","","","","","Descripcion","si")
  celdas = celdas & ofv.GenCelda("","tt colspan=2","center","","","","","Alias de Seguridad","si")

  formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si") & "</thead>"

'  strsql = "SELECT * From menux Where IDMenu = 0 and Aplicacion = " & sAplicacion & " order by Morden "
  strsql = "SELECT * From menux Where mdepen = '0' and mcodapli = '"
  strsql = strsql & sAplicacion & "' order by Morden "

  set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  do until rs2.eof
    sForm = "AtributosGrupos" & rs2(0)

      xdesatr = " <font color=#993333 face=tahoma size=1>(" & rs2(6) & ")</font>"

    if bModificar then
      sInput = ofv.GenerarInput(rs2(7).name,rs2(7),"text",sForm,"5","dtn","SI")
    else
      sInput = ofv.GenerarInput("",rs2(7),"readonly","","5","dtn","")
    end if

    celdas = ofv.GenCelda("","","","","","","",sInput,"si")


    if bModificar then
      sInput = ofv.GenerarInput(rs2(1).name,rs2(1),"text",sForm,"50","dt","")
    else
      sInput = ofv.GenerarInput("",rs2(1),"readonly","","50","dt","")
    end if

    celdas = celdas & ofv.GenCelda("","ut","","","","","",sInput,"si")

    celdas = celdas & ofv.GenCelda("","ut","","","","","",xdesatr,"si")

    if bModificar then
      vuelta2 = vuelta & "&DD=0&DS=x&Apl=" & sAplicacion
      vuelta2 = vuelta2 & "&Nombre=" & ofv.convertircar(Nombre," ","*")
      vuelta2 = ofv.convertircar(vuelta2,"&","_")
      sAccion = "asp/ActualizarMenues.asp?volver=" & vuelta2 & "&ID=" & rs2(0) & "&tabla=menux"
      fila = ofv.GenForm(sForm,"",sAccion,"","post","",celdas,"si")
    else
      fila = ofv.GenForm("","","","","post","",celdas,"si")
    end if

    sInput = ofv.generarinput("","Menues","submit","SubMenus","10","bt","")
    celdax = ofv.GenCelda("","","","","","","",sInput,"si")
    sAccion = "ManMSubmen.asp?Menu=" & ofv.convertircar(rs2(1)," ","_") & "&DD=" & rs2(6)
    sAccion = sAccion & "&DS=0&Apl=" & sAplicacion & "&Nombre="
    sAccion = sAccion & ofv.convertircar(Nombre," ","_")
    Celdas = ofv.genform("SubMenus","",sAccion,"","post","",Celdax,"si")

    formulario = formulario & ofv.GenRow("","","","","10","","",fila & celdas,"si")

    rs2.movenext
  loop

  celdas = ofv.GenCelda("","abc colspan=4","center","","","","","&nbsp;","si")
  fila = ofv.GenRow("","","","","","","",celdas,"si")

  sInput = ofv.generarinput("","Volver","submit","Volver","10","bt","")
  celdas = ofv.GenCelda("","abc colspan=4","center","","","","",sInput,"si")
  fila = fila & ofv.GenRow("","","","","","","",celdas,"si")
  sAccion = "ManMAplic.asp?Apl=x&DD=x&DS=X"
  formulario = formulario & ofv.GenForm("Volver","",sAccion,"","post","",fila,"si")

  tabla = ofv.GenTabla("","","","80%","","","0","0","0","","",formulario,"si")

  sHTML = sHTML & "<center>" & tabla & "</center></body></html>"

  call ofv.cerrarconsulta(rs2)
  call ofv.cerrarconn(cn)

end if

response.write sHTML

%>