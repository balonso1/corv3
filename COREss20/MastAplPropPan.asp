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




  sdato = request.querystring("dato")
  svalor = request.querystring("Valor")
  codapli = request.querystring("Codapli")
  coddepen = request.querystring("Coddepen")

  vuelta = ofv.vueltaasp("../",CantRegMostrar,CantRegMover) & "_codapli=" & codapli
  vuelta = vuelta & "_coddepen=" & coddepen & "_dato=" & sdato & "_valor=" & svalor

  ofv.ObtenerAtributos Session("Form"),""

  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn1)

  Strsql = "Select * From menux WHERE mcodapli = '" & codapli & "' and menuatr = '" & coddepen & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if Not rs.EOF then 
  celdas = ofv.GenCelda("","cc colspan=2","center","","","","","PROPIEDADES DE PANTALLA","si")
  formulario = "<thead>" & ofv.GenRow("","","","","10","","",celdas,"si") & "</thead>"

  celdas = ofv.GenCelda("","aaa  colspan=2","center","","","","","&nbsp;","si")
  formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si")


     sform = "Apl" & rs(0)
     filas = ""

    celdas = ofv.GenCelda("","tt","","30%","","","","Modulo: ","si")

    scx = 0
    if not isnull(rs(2)) then 
       scx = instr(rs(2),"?")
    end if

    if scx > 0 then 
       sCompo = mid(rs(2),1,cint(scx) - 1)
       sstr = mid(rs(2),cint(scx) + 1,len(rs(2)) - cint(scx) )  
       scx = instr(sstr,"&NDAT")
       if scx = 0 then 
          scx = instr(sstr,"NDAT")
          if scx > 0 then 
             sstr = ""
          end if 
       else
          sstr = mid(sstr,1,cint(scx) - 1 )              
       end if
    else
       sCompo = rs(2)
       sstr = ""  
    end if
    
    


    if bModificar then
      sInput = ofv.GenerarInput("scompo",scompo,"text",sForm,"90","dt","")
    else
      sInput = scompo
    end if
    celdas = celdas & ofv.GenCelda("","cc","","70%","","","",sInput,"si")

    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")
      


    celdas = ofv.GenCelda("","tt","","30%","","","","String adicional: ","si")

    if bModificar then
      sInput = ofv.GenerarInput("sstr",sstr,"textsb",sForm,"90","dt","")
    else
      sInput = sstr
    end if
    celdas = celdas & ofv.GenCelda("","cc","","70%","","","",sInput,"si")

    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    sInput1 = ofv.GenerarInput("sdato",sdato,"hidden","","1","dt","")
    sInput2 = ofv.GenerarInput("svalor",svalor,"hidden","","1","dt","")

    celdas = ofv.GenCelda("","aaa colspan=2","","100%","","","",sInput1 & sInput2,"si")

    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    if bModificar then
       sAccion = "asp/ActualizarPropPan.asp?volver=" & vuelta & "&codapli=" & codapli & "&coddepen=" & coddepen
      fila = ofv.GenForm(sform,"",sAccion,"","post","",filas,"si")
    else
      fila = ofv.GenForm("","","","","post","",filas,"si")
    end if


    formulario = formulario & fila 

   end if

 


  celdas = ofv.GenCelda("","aaa colspan=2","center","","","","","&nbsp;","si")
  formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si")

  sInput = ofv.generarinput("","Volver","submit","Volver","10","bt","")
  sAccion = "MastAplPan.asp?codapli=" & codapli & "&coddepen=" & coddepen
  retornar = ofv.genform("Volver","",sAccion,"","post","",sInput,"si")


  celdas = ofv.GenCelda("","aaa colspan=2","","","","","",retornar,"si")
  formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si")


  tabla = ofv.GenTabla("","","","80%","","0","0","0","0","#6699cc","",formulario,"si")

  sHTML = sHTML & "<center><br>" & tabla



  sHTML = sHTML & "</center></body></html>"

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)



end if

response.write sHTML

%>