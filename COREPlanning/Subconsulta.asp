<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->

<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

set orp = server.createobject("CorPlanRep.Reportes")

Set orp.Request = Request
Set orp.Server = Server
Set orp.Response = Response
Set orp.Session = Session
Set orp.Ofv = Ofv


  sTipo = UCase(request.querystring("Tipo"))
  sMenu = request.querystring("Menu")
  sNombre = request.querystring("Nombre")
  sNombre = "Consulta por " & sNombre
  vuelta = request.querystring("vuelta")


  sEntidadAnt = request.querystring("CodEntidad") 
  sTemaAnt = request.querystring("Codtema")
  sSubTemaAnt = request.querystring("Codprodserv")
  sEventoAnt = request.querystring("Codevento")
  sProcesoAnt = request.querystring("Codproceso")
  sContinAnt = request.querystring("Codcontin")
  sCodcir = request.querystring("Codcir")
  sAsoc = request.querystring("CodAsoc")

  if vuelta = 1 then
    sEntidad = sEntidadAnt
    sTema = sTemaAnt
    sSubtema = sSubtemaAnt
    sEvento = sEventoAnt
    sProceso = sProcesoAnt
    sContin = sContinAnt
    
  else
    sEntidad = request.form("entidad")
    sTema = request.form("tema")
    sSubtema = request.form("prodserv")
    sEvento = request.form("evento")
    sProceso = request.form("proceso")
    sContin = request.form("contin")
    if sTema <> sTemaAnt then sSubtema=0
    if sSubtema <> sSubtemaAnt then sEvento=0
    if sEvento <> sEventoAnt then sProceso=0
    if UCase(sNombre) <> "CONSULTA POR CONTINGENCIA" then
    if sProceso <> sProcesoAnt then sContin=0
    end if
  end if

  if sEntidad = "" or sEntidad = "nsnc" then sEntidad = 0
  if sTema = "" or sTema = "nsnc" then sTema = 0
  if sSubtema = "" or sSubtema = "nsnc" then sSubtema=0
  if sEvento = "" or sEvento = "nsnc" then sEvento = 0
  if sProceso = "" or sProceso = "nsnc" then sProceso = 0
  if sContin = "" or sContin = "nsnc" then sContin = 0

  if sTipo = "" then
    set cn = ofv.conectar(ofv.strconn4)

    sHTML = ofv.FormHeader(sNombre)

    sValida = ""
    if UCase(sNombre) = "CONSULTA POR ENTIDAD" then
      sValida = "sb"
      SQL0 = "Select * From ECO Order By Descri"

      Celda = ofv.GenCelda("","abc colspan=3","","","","","","&nbsp;","si")
      Fila = ofv.GenRow("","","","","","","",Celda,"si")
      Celda = ofv.GenCelda("","","","40%","","","","&nbsp;","si")
      Celda = Celda & ofv.GenCelda("","tt","right","10%","","","","Entidad:&nbsp;&nbsp;","si")
      sInput = ofv.GenerarCombo("Subconsulta","entidad",sEntidad,sql0,cn,0,1,"")
      Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")
      Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
    elseif UCase(sNombre) = "CONSULTA POR CONTINGENCIA" then
      sValida = "sb"
      SQL0 = "Select * From Tipocont Order By Descri"

      Celda = ofv.GenCelda("","abc colspan=3","","","","","","&nbsp;","si")
      Fila = ofv.GenRow("","","","","","","",Celda,"si")
      Celda = ofv.GenCelda("","","","40%","","","","&nbsp;","si")
      Celda = Celda & ofv.GenCelda("","tt","right","10%","","","","Contingencia:&nbsp;&nbsp;","si")
      sInput = ofv.GenerarCombo("Subconsulta","contin",sContin,sql0,cn,0,1,"")
      Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")
      Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
    end if

    SQL1 = "Select * From Tema Order By Descri"
    SQL2 = "Select * From PRODSERV Where codtema = " & sTema &  " Order By Descri"
    SQL3 = "Select * From Evento Where Codprodserv =" & sSubtema & " Order By Descri"
    SQL4 = "Select * From Proceso Where Codevento =" & sEvento & " Order By Descri "
    SQL5 = "SELECT CONTINGENCIA.CodContin, TIPOCONT.descri, CONTINGENCIA.codproceso "
    SQL5 = SQL5 & "FROM CONTINGENCIA INNER JOIN TIPOCONT ON "
    SQL5 = SQL5 & "CONTINGENCIA.tipocontin = TIPOCONT.tipocontin "
    SQL5 = SQL5 & "Where CONTINGENCIA.Codproceso = " & sProceso & " Order By TIPOCONT.Descri"

    Celda = ofv.GenCelda("","abc colspan=3","","","","","","&nbsp;","si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
    Celda = ofv.GenCelda("","","","40%","","","","&nbsp;","si")
    Celda = Celda & ofv.GenCelda("","tt","right","10%","","","","Tema:&nbsp;&nbsp;","si")
    sInput = ofv.GenerarCombo("Subconsulta","tema",sTema,sql1,cn,0,1,svalida)
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    Celda = ofv.GenCelda("","abc colspan=3","","","","","","&nbsp;","si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
    Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    Celda = Celda & ofv.GenCelda("","tt","right","","","","","Subtema:&nbsp;&nbsp;","si")
    sInput = ofv.GenerarCombo("Subconsulta","prodserv",sSubtema,sql2,cn,0,1,"sb")
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    Celda = ofv.GenCelda("","abc colspan=3","","","","","","&nbsp;","si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
    Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    Celda = Celda & ofv.GenCelda("","tt","right","","","","","Eventos:&nbsp;&nbsp;","si")
    sInput = ofv.GenerarCombo("Subconsulta","evento",sEvento,sql3,cn,0,1,"sb")
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    Celda = ofv.GenCelda("","abc colspan=3","","","","","","&nbsp;","si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
    Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    Celda = Celda & ofv.GenCelda("","tt","right","","","","","Procesos:&nbsp;&nbsp;","si")
    sInput = ofv.GenerarCombo("Subconsulta","proceso",sProceso,sql4,cn,0,1,"sb")
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    if UCase(sNombre) <> "CONSULTA POR CONTINGENCIA" then
      Celda = ofv.GenCelda("","abc colspan=3","","","","","","&nbsp;","si")
      Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
      Celda = Celda & ofv.GenCelda("","tt","right","","","","","Contingencias:&nbsp;&nbsp;","si")
      sInput = ofv.GenerarCombo("Subconsulta","contin",sContin,sql5,cn,0,1,"sb")
      Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")
      Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
    end if

    sAccion = "Subconsulta.asp?Menu=" & sMenu & "&Nombre=" & request.querystring("Nombre")
    sAccion = sAccion & "&Codtema=" & sTema & "&Codprodserv=" & sSubTema & "&Codevento="
    sAccion = sAccion & sEvento & "&Codproceso=" & sProceso & "&Codcontin=" & sContin
    sAccion = sAccion & "&CodEntidad=" & sEntidad & "&Codcir=" & sCodcir  
    Formulario = ofv.GenForm("Subconsulta","",sAccion,"","post","",Fila,"si")

    Celda = ofv.GenCelda("","abc colspan=3","","","","","","&nbsp;","si")
    Fila = ofv.GenRow("","","","","","","",Celda,"si")
    if UCase(sNombre) <> "CONSULTA POR ENTIDAD" then
      Celda = ofv.GenCelda("","abc colspan=3","","","","","","&nbsp;","si")
      Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
    end if

    sInput = ofv.GenerarInput("","Circuito&nbsp;Alternativo","submit","","10","bt","")
    if UCase(sNombre) = "CONSULTA POR ENTIDAD" then
      Celda = ofv.GenCelda("","","Center colspan=3","","","","",sInput,"si")
    else
      Celda = ofv.GenCelda("","","Center colspan=2","","","","",sInput,"si")
    end if
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"no")

    if UCase(sNombre) = "CONSULTA POR TEMA" then
      if sContin = 0 then
        sASP = "Subconsulta.asp?Tipo=SubtemaAlternativoVer"
      else
        sASP = "Subconsulta.asp?Tipo=SubtemaAlternativo"
      end if
      sFunc = "onsubmit=" & chr(34)
      sFunc = sFunc & "if (Subconsulta.tema.options[Subconsulta.tema.selectedIndex]."
      sFunc = sFunc & "value=='nsnc') { alert( 'Seleccione un Tema ');return false;}" & chr(34)
    elseif UCase(sNombre) = "CONSULTA POR ENTIDAD" then
      sASP = "Subconsulta.asp?Tipo=EntidadAlternativaVer"
      sFunc = "onsubmit=" & chr(34)
      sFunc = sFunc & "if (Subconsulta.entidad.options[Subconsulta.entidad.selectedIndex]."
      sFunc = sFunc & "value=='nsnc') { alert( 'Seleccione una Entidad ');return false;}"
      sFunc = sFunc & chr(34)
    elseif UCase(sNombre) = "CONSULTA POR CONTINGENCIA" then
      sASP = "Subconsulta.asp?Tipo=ContingenciaAlternativaVer"
      sFunc = "onsubmit=" & chr(34)
      sFunc = sFunc & "if (Subconsulta.contin.options[Subconsulta.contin.selectedIndex]."
      sFunc = sFunc & "value=='nsnc') { alert( 'Seleccione una Contingencia ');return false;}"
      sFunc = sFunc & chr(34)
    end if 

    sAccion = sASP & "&CodTema=" & sTema & "&CodProdserv=" & sSubTema & "&CodEvento=" & sEvento
    sAccion = sAccion & "&CodProceso=" & sProceso & "&CodContin=" & sContin & "&Menu=" & sMenu
    sAccion = sAccion & "&Nombre=" & request.querystring("Nombre") & "&CodEntidad=" & sEntidad
    sAccion = sAccion &  "&Codcir=" & sCodcir & "&Vuelta=1"
    Formulario = Formulario & ofv.GenForm("","",sAccion,"","post",sFunc,Fila,"si")

    if UCase(sNombre) <> "CONSULTA POR ENTIDAD" then
      sInput = ofv.GenerarInput("","Circuito&nbsp;Correctivo","submit","","10","bt","")
      Celda = ofv.GenCelda("","","Center colspan=3","","","","",sInput,"si")

      if UCase(sNombre) = "CONSULTA POR CONTINGENCIA" then
        sASP = "Subconsulta.asp?Tipo=ContingenciaCorrectivaVer"
      else
        if sContin = 0 then
          sASP = "Subconsulta.asp?Tipo=SubtemaCorrectivoVer"
        else
          sASP = "Subconsulta.asp?Tipo=SubtemaCorrectivo"
        end if
      end if

      sAccion = sASP & "&CodTema=" & sTema & "&CodProdserv=" & sSubTema & "&CodEvento=" & sEvento
      sAccion = sAccion & "&CodProceso=" & sProceso & "&CodContin=" & sContin & "&Codcir=" & sCodcir & "&Menu=" & sMenu
      sAccion = sAccion & "&Nombre=" & request.querystring("Nombre") & "&Vuelta=1"
      Formulario = Formulario & ofv.GenForm("","",sAccion,"","post",sFunc,Celda,"si")
      Formulario = Formulario & ofv.Genfrow()
    end if

    sHTML = sHTML & "<center>"
    sHTML = sHTML & ofv.GenTabla("","","","100%","","0","","0","0","","",Formulario,"si")
    sHTML = sHTML & "</center></body></html>"

    call ofv.cerrarconn(cn)
  else
    set cn = ofv.conectar(ofv.strconn4)

    select case sTipo
      case "SUBTEMAALTERNATIVOVER"
'response.write "a " & sTema & " b " & sSubtema & " c " & sEvento & " d " & sProceso & " e " & sContin & " f " & sEntidad & " g " & sCodcir
        orp.GenerarReporteDetallado "CONSULTA CIRCUITO ALTERNATIVO", sTema, sSubtema, sEvento, sProceso, sContin, sEntidad, sCodcir, cn, sMenu, request.querystring("Nombre")
      case "SUBTEMACORRECTIVOVER"
        orp.GenerarReporteDetallado "CONSULTA CIRCUITO CORRECTIVO", sTema, sSubtema, sEvento, sProceso, sContin, sEntidad, sCodcir, cn, sMenu, request.querystring("Nombre")
      case "CONTINGENCIAALTERNATIVAVER"
        orp.GenerarReporteDetallado "CONSULTA CIRCUITO ALTERNATIVO", sTema, sSubtema, sEvento, sProceso, sContin, sEntidad, "0", cn, sMenu, request.querystring("Nombre")
      case "CONTINGENCIACORRECTIVAVER"
        orp.GenerarReporteDetallado "CONSULTA CIRCUITO CORRECTIVO", sTema, sSubtema, sEvento, sProceso, sContin, sEntidad, "0", cn, sMenu, request.querystring("Nombre")
      case "ENTIDADALTERNATIVAVER"
        orp.GenerarReporteDetallado "CONSULTA CIRCUITO ALTERNATIVO", sTema, sSubtema, sEvento, sProceso, sContin, sEntidad, "0", cn, sMenu, request.querystring("Nombre")
      case "SUBTEMAALTERNATIVO"
        orp.GenerarReporteTotal "CONSULTA CIRCUITO ALTERNATIVO",  sTema, sSubtema, sEvento, sProceso, sContin, "", cn, sMenu, request.querystring("Nombre")
      case "SUBTEMACORRECTIVO"
        orp.GenerarReporteTotal "CONSULTA CIRCUITO CORRECTIVO", sTema, sSubtema, sEvento, sProceso, sContin, "", cn, sMenu, request.querystring("Nombre")
    end select

    call ofv.cerrarconn(cn)
  end if
end if

response.write sHTML

%>