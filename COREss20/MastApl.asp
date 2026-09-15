<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%



ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0

'  if ofv.Explorador = "MSIE" then
'    sHTML = ofv.MenuHeader("#cccccc","")
'  else
'    sHTML = ofv.MenuHeader("#ccbbaa","")
'  end if

  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))

  CantRegMostrar = 11
  if request.querystring("CantRegMover") = "" then
     CantRegMover = 0
  else 
     CantRegMover = cdbl(request.querystring("CantRegMover"))
  end if


  vuelta = ofv.vueltaasp("../",CantRegMostrar,CantRegMover)


  ofv.ObtenerAtributos request.querystring("seg"),""
  Session("Form") = request.querystring("seg")

  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  celdas = ofv.GenCelda("","cc colspan=6","center","","","","","APLICACIONES DE COR E-SOLUTION SUITE","si")
  formulario = "<thead>" & ofv.GenRow("","","","","10","","",celdas,"si") & "</thead>"

  celdas = ofv.GenCelda("","aaa colspan=6","center","","","","","&nbsp;","si")
  formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si")

  set cn = ofv.conectar(ofv.strconn1)

  Strsql = "Select * From CorAplic Order By ordenapli, descapli"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
'  if Not rs.EOF then rs.move(CantRegMover)
'  posI = 0
'  Do While (Not rs.EOF) and (posi < CantRegMostrar)
  Do While Not rs.EOF 
     sAplicacion = rs(1)
     sform = "Apl" & rs(1)
    'Eliminar
    if bSuprimir and rs("modoapli") <> "F" then
      sHref = "asp/EliminarApl.asp?ID=" & rs(0) & "&Descripcion=" & ofv.convertircar(rs(2)," ","_")
      sHref = sHref & "&volver=" & vuelta & "&Eliminar=NSNC&alias=" & rs(1)
      sHref = sHref & "&aplic=" & rs(1)
      celdas =  ofv.BotonEliminar2(replace(rs(2),"*"," "),sHref,"")
    else
      celdas =  ofv.GenCelda("","cc","","5%","","","","&nbsp;","si")
    end if

    if bModificar then
      sInput = ofv.GenerarInput(rs(13).name,rs(13),"text",sForm,"5","dtn","SI")
    else
      sInput = ofv.GenerarInput("",rs(13),"readonly","","5","dtn","")
    end if

    celdas = celdas & ofv.GenCelda("","cc","","8%","","","",sInput,"si")

    if bModificar  and rs("modoapli") <> "F" then
      sInput = ofv.GenerarInput(rs(2).name,replace(rs(2),"*"," "),"text",sForm,"100","dt","")
    else
      sInput = replace(rs(2),"*"," ")
    end if
    celdas = celdas & ofv.GenCelda("","cc","","63%","","","",sInput,"si")
      

    if bModificar  and rs("modoapli") <> "F" then
      sInput = "<font class=at><select onchange=" & sForm & ".submit() "
      sInput = sInput & "name=enable ><option value=1>SI  "
      if  cint(rs("enable")) <> 0 then
        sInput = sInput & "<option value=0>NO  "
      else
        sInput = sInput & "<option value=0 selected>NO  "
      end if
      sInput = sInput & "</select></font>"
    else
      if cint(rs("enable")) = 0 then
         xestado = "NO"
      else
         xestado = "SI"
      end if   
      sInput = xestado
    end if
    celdas = celdas & ofv.GenCelda("","cc","","8%","","","",sInput,"si")

    if bModificar then
       sAccion = "asp/ActualizarMenues.asp?volver=" & vuelta & "&ID=" & rs(0) & "&tabla=Coraplic" 
      fila = ofv.GenForm(sform,"",sAccion,"","post","",celdas,"si")
    else
      fila = ofv.GenForm("","","","","post","",celdas,"si")
    end if

    if cint(rs("enable")) <> 0  and rs("modoapli") <> "" then
      sInput = ofv.generarinput("","Solapas","submit","Menus","10","bt","")
      celdax = ofv.GenCelda("","cc","","8%","","","",sInput,"si")
      sAccion = "MastAplMen.asp?codapli=" & sAplicacion & "&coddepen=0"
      Celdas = ofv.genform("Menus","",sAccion,"","post","",Celdax,"si")
    else
      celdas = ofv.GenCelda("","cc","","8%","","","","&nbsp;","si")
    end if

    if cint(rs("enable")) <> 0  and rs("modoapli") <> "X" then
      sInput = ofv.generarinput("","Propiedades","submit","Prop","10","bt","")
      celdax = ofv.GenCelda("","cc","","8%","","","",sInput,"si")
      sAccion = "MastAplProp.asp?codapli=" & sAplicacion
      celdas = celdas & ofv.genform("Prop","",sAccion,"","post","",Celdax,"si")
    else
      celdas = celdas & ofv.GenCelda("","cc","","8%","","","","&nbsp;","si")
    end if

    formulario = formulario & ofv.GenRow("","","","","10","","",fila & celdas,"si")



'    posi = posi + 1
    rs.movenext

  loop


  celdas = ofv.GenCelda("","aaa colspan=6","center","","","","","&nbsp;","si")
  formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si")

  'Botonera
  Agregar = "&nbsp;"
  if bAgregar then
     sInput = ofv.generarinput("","Agregar","submit","Agr","10","bt","")
     sAccion = "AgregarMastApl.asp?Regisvuelta=" & CantRegMover & "&tabla=CorAplic&volver=" & vuelta
     Agregar = ofv.genform("Agr","",sAccion,"","post","",sInput,"si")
  end if

  celdas = ofv.GenCelda("","aaa colspan=6","","","","","",Agregar,"si")
  formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si")


  tabla = ofv.GenTabla("","","","80%","","0","0","0","0","#6699cc","",formulario,"si")

  sHTML = sHTML & "<center><br>" & tabla

'  'Botonera
'  Agregar =  ofv.GenCelda("","","","","","","","&nbsp;","si")
'  if bAgregar then
'    sAccion = "AgregarMastApl.asp?Regisvuelta=" & CantRegMover & "&tabla=CorAplic&volver=" & vuelta
'    Agregar = ofv.BotonAgregar(sAccion)
'  end if

'  formulario = ofv.GenRow("","","","","10","","",Agregar,"si")
'  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","0","0","0","#6699cc","",formulario,"si")

  sHTML = sHTML & "</center></body></html>"

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)



end if

response.write sHTML

%>