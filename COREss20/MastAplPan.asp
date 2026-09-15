<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%



ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0

  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))

  CantRegMostrar = 11
  if request.querystring("CantRegMover") = "" then
     CantRegMover = 0
  else 
     CantRegMover = cdbl(request.querystring("CantRegMover"))
  end if

  codapli = request.querystring("Codapli")
  coddepen = request.querystring("Coddepen")

  vuelta = ofv.vueltaasp("../",CantRegMostrar,CantRegMover)
  vuelta = vuelta & "_codapli=" & codapli & "_coddepen=" & coddepen
  vuelta2 = replace(vuelta,"../","") 


  ofv.ObtenerAtributos Session("Form"),""


  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn1)




     Strsql = "Select * From Menux where Menuatr = '" & coddepen & "' and mcodapli = '" & codapli & "' "
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if Not rs.EOF then
        sdescMenu = replace(rs(1),"*"," ")
     end if
     sdepen = rs(8)
     call ofv.cerrarconsulta(rs)
     stipo = "X"
     stipoD = "PANTALLAS DE " & Ucase(sDescMenu)
     coddepenS = " and depen = '" & coddepen & "' and (tipo = 'F' or tipo = 'N') "
     call ofv.cerrarconn(cn)

     sAplicacion = codapli

  celdas = ofv.GenCelda("","TTT colspan=4","","","","","",stipoD,"si")
  formulario = "<thead>" & ofv.GenRow("","","","","10","","",celdas,"si") & "</thead>"

  celdas = ofv.GenCelda("","aaa colspan=4","center","","","","","&nbsp;","si")
  formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si")


  set cn = ofv.conectar(ofv.strconn0)

  Strsql = "Select * From Atrm "
  Strsql = Strsql & " where aplicacion = '" & codapli & "' "
  Strsql = Strsql & coddepenS
  Strsql = Strsql & " order by alias "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  Do While Not rs.EOF
     sform = "AplM" & rs(0)
    'Eliminar
    if bSuprimir  then
      sHref = "asp/EliminarApl.asp?ID=" & rs(0) & "&Descripcion=" & ofv.convertircar(rs(3)," ","_")
      sHref = sHref & "&volver=" & vuelta & "&Eliminar=NSNC&alias=" & rs("alias")
      sHref = sHref & "&aplic=" & codapli
      celdas =  ofv.BotonEliminar2(replace(rs(3),"*"," "),sHref,"")
    else
      celdas =  ofv.GenCelda("","UT","","5%","","","","&nbsp;","si")
    end if


    
    if rs("tipo") = "N" then
       sTipo = "N"
       sinput = "NODO DE " & Ucase(sdescmenu)
    else
       sTipo = "F"
       sinput = "FORM DE " & Ucase(sdescmenu)
    end if

    celdas = celdas & ofv.GenCelda("","UT","","45%","","","",sInput,"si")

    if bModificar   then
      sInput = ofv.GenerarInput(rs(3).name,replace(rs(3),"*"," "),"text",sForm,"50","dt","")
    else
      sInput = replace(rs(3),"*"," ")
    end if
    celdas = celdas & ofv.GenCelda("","UT","","40%","","","",sInput,"si")
      


    if bModificar then
       sAccion = "asp/ActualizarMenues.asp?volver=" & vuelta & "&ID=" & rs(0) & "&tabla=Atrm" 
      fila = ofv.GenForm(sform,"",sAccion,"","post","",celdas,"si")
    else
      fila = ofv.GenForm("","","","","post","",celdas,"si")
    end if


    sInput = ofv.generarinput("","Propiedades","submit","Menus","10","bt","")
    celdax = ofv.GenCelda("","UT","","10%","","","",sInput,"si")
    sAccion = "MastAplPropPan.asp?codapli=" & codapli & "&coddepen=" & coddepen & "&dato=" & rs(8) & "&valor=" & rs(9)
    celdas = ofv.genform("Menus","",sAccion,"","post","",Celdax,"si")


    formulario = formulario & ofv.GenRow("","","","","10","","",fila & celdas,"si")

    if stipo = "N" then

       sform = "AplM2" & rs(0)

       celdas =  ofv.GenCelda("","","","5%","","","","&nbsp;","si")

       if bModificar   then
          sinput = rs(8)
          if isnull(sinput) then sinput = "XXX"  
          sInput = ofv.GenerarInput(rs(8).name,sinput,"text",sForm,"50","dt","")
       else
          sInput = rs(8)
       end if
       celdas = celdas & ofv.GenCelda("","UT","","40%","","","","Dato: " & sInput,"si")

       if bModificar   then
          sinput = rs(9)
          if isnull(sinput) then sinput = "XXX"  
          sInput = ofv.GenerarInput(rs(9).name,sinput,"text",sForm,"50","dt","")
       else
          sInput = rs(9)
       end if
       celdas = celdas & ofv.GenCelda("","UT","","40%","","","","Valor: " & sInput,"si")

       celdas = celdas & ofv.GenCelda("","UT","","10%","","","","&nbsp;","si")

       if bModificar then
          sAccion = "asp/ActualizarMenues.asp?volver=" & vuelta & "&ID=" & rs(0) & "&tabla=Atrm" 
          fila = ofv.GenForm(sform,"",sAccion,"","post","",celdas,"si")
       else
          fila = ofv.GenForm("","","","","post","",celdas,"si")
       end if

       formulario = formulario & ofv.GenRow("","","","","10","","",fila,"si")


    end if

    celdas = ofv.GenCelda("","aaa colspan=4","center","","","","","&nbsp;","si")
    formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si")



    rs.movenext

  loop

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  celdas = ofv.GenCelda("","aaa colspan=4","center","","","","","&nbsp;","si")
  formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si")

  tabla = ofv.GenTabla("","","","80%","","0","0","0","0","","",formulario,"si")

  sHTML = sHTML & "<center><br>" & tabla & "</center>"

  'Botonera
  Agregar = ""


  'Botonera
  if bAgregar  and stipo <> "F" then
     sAccion = "AgregarMastAplPan.asp?Regisvuelta=" & CantRegMover & "&tabla=Atrm&volver="
     sAccion = sAccion & vuelta & "&stipo=" & stipo & "&sdepen=" & coddepen
     sAccion = sAccion & "&sapli=" & codapli & "&stipoD=" & stipoD
     Agregar = ofv.BotonAgregar(sAccion)
  end if

  sAccion = "MastAplMen.asp?codapli=" & codapli & "&coddepen=" & sdepen
  Celda   = ofv.BotonVolver(sAccion)

  Buscar = ""

  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","",Agregar,Buscar,Celda,"")

  sHTML = sHTML & "</body></html>"


end if

response.write sHTML

%>