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



  if coddepen = "0" then 
     Strsql = "Select * From CorAplic where codapli = '" & codapli & "' "
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if Not rs.EOF then
        sdescapli = replace(rs(2),"*"," ")
     end if
     call ofv.cerrarconsulta(rs)
     stipo = "S"
     stipoD = "SOLAPAS DE " & sDescapli
     coddepenS = " and mdepen = '" & coddepen & "' "
  else
     Strsql = "Select * From Menux where Menuatr = '" & coddepen & "' and mcodapli = '" & codapli & "' "
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if Not rs.EOF then
        sdescMenu = replace(rs(1),"*"," ")
     end if
     sdepen = rs(8)
     if rs(8) <> "0" then 
        smensub = "S"
     else
        smensub = "N"
     end if 
     call ofv.cerrarconsulta(rs)
     stipo = "M"
     stipoD = "MENUES DE " & sDescMenu
     if smensub = "S" then stipoD = "SUB" & stipoD 
     coddepenS = " and mdepen = '" & coddepen & "' "
  end if 

     sAplicacion = codapli

  celdas = ofv.GenCelda("","cc colspan=6","center","","","","",stipoD,"si")
  formulario = "<thead>" & ofv.GenRow("","","","","10","","",celdas,"si") & "</thead>"

  celdas = ofv.GenCelda("","aaa colspan=6","center","","","","","&nbsp;","si")
  formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si")



  Strsql = "Select * From Menux "
  Strsql = Strsql & " where mcodapli = '" & codapli & "' "
  Strsql = Strsql & coddepenS
  Strsql = Strsql & " order by morden, menu"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
'  if Not rs.EOF then rs.move(CantRegMover)
'  posI = 0

'  Do While (Not rs.EOF) and (posi < CantRegMostrar)
  Do While Not rs.EOF
     sform = "AplM" & rs(0)
    'Eliminar
    if bSuprimir  then
      sHref = "asp/EliminarApl.asp?ID=" & rs(0) & "&Descripcion=" & ofv.convertircar(rs(1)," ","_")
      sHref = sHref & "&volver=" & vuelta & "&Eliminar=NSNC&alias=" & rs("menuatr")
      sHref = sHref & "&aplic=" & codapli
      celdas =  ofv.BotonEliminar2(replace(rs(1),"*"," "),sHref,"")
    else
      celdas =  ofv.GenCelda("","cc","","5%","","","","&nbsp;","si")
    end if

    if bModificar then
      sInput = ofv.GenerarInput(rs(7).name,rs(7),"text",sForm,"2","dtn","SI")
    else
      sInput = ofv.GenerarInput("",rs(7),"readonly","","2","dtn","")
    end if

    celdas = celdas & ofv.GenCelda("","cc","","2%","","","",sInput,"si")

    if bModificar   then
      sInput = ofv.GenerarInput(rs(1).name,replace(rs(1),"*"," "),"text",sForm,"90","dt","")
    else
      sInput = replace(rs(1),"*"," ")
    end if
    celdas = celdas & ofv.GenCelda("","cc","","55%","","","",sInput,"si")
      


    if bModificar then
       sAccion = "asp/ActualizarMenues.asp?volver=" & vuelta & "&ID=" & rs(0) & "&tabla=Menux" 
      fila = ofv.GenForm(sform,"",sAccion,"","post","",celdas,"si")
    else
      fila = ofv.GenForm("","","","","post","",celdas,"si")
    end if

      if coddepen = "0" then
         sPref = "Menues" 
      else
         sPref = "Submenues" 
      end if 


      
      sTMen = "N"
      sTFrm = "N"
      sTNod = "N"

      if stipo <> "S" then
         set cx = ofv.conectar(ofv.strconn0)
         Strsql = "Select tipo From atrm where aplicacion = '" & codapli & "' "
         Strsql = Strsql & " and depen = '" & rs(6) & "' "
         set rx = ofv.crearconsultaEx(StrSql,cx,1,parametros)
         if Not rx.EOF then
            select case rx(0)
                   case "M"
                        sTMen = "S"

                   case "F"
                        sTFrm = "S"

                   case "N"
                        sTNod = "S"
            end select
         end if
         call ofv.cerrarconsulta(rx) 
         call ofv.cerrarconn(cx)
      end if       


      if stipo <> "S" and smensub = "S" then
         celdas =  ofv.GenCelda("","cc","","10%","","","","&nbsp;","si")
      else
         if sTFrm = "S" or sTNod = "S" then
            celdas =  ofv.GenCelda("","cc","","10%","","","","&nbsp;","si")
         else
            sInput = ofv.generarinput("",sPref,"submit","Menus","10","bt","")
            celdax = ofv.GenCelda("","cc","","10%","","","",sInput,"si")
            sAccion = "MastAplMen.asp?codapli=" & sAplicacion & "&coddepen=" & rs(6)
            Celdas = ofv.genform("Menus","",sAccion,"","post","",Celdax,"si")
         end if
      end if

      if stipo = "S" then
         celdas = celdas & ofv.GenCelda("","cc","","13%","","","","&nbsp;","si")
      else
         if sTMen = "S"  then
            celdas = celdas & ofv.GenCelda("","cc","","13%","","","","&nbsp;","si")
         else
            sInput = ofv.generarinput("","Pantallas","submit","Menus","10","bt","")
            celdax = ofv.GenCelda("","cc","","13%","","","",sInput,"si")
            sAccion = "MastAplPan.asp?codapli=" & sAplicacion & "&coddepen=" & rs(6)
            celdas = celdas & ofv.genform("Menus","",sAccion,"","post","",Celdax,"si")
         end if  
      end if


'      sInput = ofv.generarinput("","Propiedades","submit","Menus","10","bt","")
'      celdax = ofv.GenCelda("","cc","","10%","","","",sInput,"si")
'      sAccion = "Menues.asp?Aplicacion=" & sAplicacion
'      sAccion = sAccion & "&Opcion=" & ID & "&Nombre=" & ofv.convertircar(sAplicacion," ","_")
'      sAccion = sAccion & "&t=" & t
'      celdas = celdas & ofv.genform("Menus","",sAccion,"Contenido","post","",Celdax,"si")

       celdas = celdas & ofv.GenCelda("","cc","","10%","","","","&nbsp;","si")

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
     sAccion = "AgregarMastAplMen.asp?Regisvuelta=" & CantRegMover & "&tabla=Menux&volver="
     sAccion = sAccion & vuelta & "&stipo=" & stipo & "&sdepen=" & coddepen
     sAccion = sAccion & "&sapli=" & codapli & "&stipoD=" & stipoD & "&ssub=" & smensub
     Agregar = ofv.genform("Agr","",sAccion,"","post","",sInput,"si")
  end if

  sInput = ofv.generarinput("","Volver","submit","Volver","10","bt","")
  if coddepen = "0" then
      sAccion = "MastApl.asp?seg=" & session("Form")
  else
      sAccion = "MastAplMen.asp?codapli=" & codapli & "&coddepen=" & sdepen
  end if
 
  retornar = ofv.genform("Volver","",sAccion,"","post","",sInput,"si")


  celdas = ofv.GenCelda("","aaa colspan=2","","","","","",Agregar,"si")
  celdas = celdas & ofv.GenCelda("","aaa colspan=4","","","","","",retornar,"si")
  formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si")



  tabla = ofv.GenTabla("","","","80%","","0","0","0","0","#6699cc","",formulario,"si")

  sHTML = sHTML & "<center><br>" & tabla & ""

'  'Botonera
'  Agregar = ofv.GenCelda("","","","","","","","&nbsp;","si")
'  if bAgregar then
'    sAccion = "AgregarMastAplMen.asp?Regisvuelta=" & CantRegMover & "&tabla=Menux&volver="
'    sAccion = sAccion & vuelta & "&stipo=" & stipo & "&coddepen=" & coddepen
'    sAccion = sAccion & "&codapli=" & codapli
'    Agregar = ofv.BotonAgregar(sAccion)
'  end if

'  sInput = ofv.generarinput("","Volver","submit","Volver","10","bt","")
'  celdax = ofv.GenCelda("","","","","","","",sInput,"si")
'  if coddepen = "0" then
'      sAccion = "MastApl.asp?seg=" & session("Form")
'  else
'      sAccion = "MastAplMen.asp?codapli=" & codapli & "&coddepen=" & sdepen
'  end if
' 
'  retornar = ofv.genform("Volver","",sAccion,"","post","",Celdax,"si")


'  formulario = ofv.GenRow("","","","","10","","",Agregar & retornar,"si")
'  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","0","0","0","#6699cc","",formulario,"si")
  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","","","","")

  sHTML = sHTML & "</center></body></html>"

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)



end if

response.write sHTML

%>