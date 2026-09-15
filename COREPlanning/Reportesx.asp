<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
set orp = server.createobject("CorPlanRep.Reportes")

Set orp.Request = Request
Set orp.Server = Server
Set orp.Response = Response
Set orp.Session = Session
Set orp.Ofv = Ofv


'  response.expires=0

  dim aAsp()
  dim aTitulo()
  dim aHeaders()
  dim aCampos()
  dim aAnchos()

  sSubReporte = ofv.convertircar(request.querystring("SubReporte"),"_"," ")
  xReporte = request.querystring("Reporte")
  sReporte = ofv.convertircar(xReporte,"_"," ")
  sTipo = request.querystring("Tipo")

  if sSubReporte <> "" then
    select case UCase(sSubReporte)
      case "ASOCIACIONES"
        redim aHeaders(0)
        redim aCampos(0)
        redim aAnchos(0)

        aHeaders(0) = "Descripcion de la Asociacion"
        aCampos(0) = 1
        aAnchos(0) = "100%"
        Strsql = "Select * From Asociacion Order By Descri"
      case "NIVELES DE ASOCIACION"
        redim aHeaders(2)
        redim aCampos(2)
        redim aAnchos(2)

        aHeaders(0) = "Asociacion"
        aHeaders(1) = "Sec"
        aHeaders(2) = "Nivel"
        aCampos(0) = 1
        aCampos(1) = 4
        aCampos(2) = 3
        aAnchos(0) = "35%"
        aAnchos(1) = "1%"
        aAnchos(2) = "54%"
        Strsql = "Select Asociacion.*, Niveles.* from Asociacion, Niveles "
        Strsql = Strsql & "Where Asociacion.codasoc = Niveles.codasoc "
        Strsql = Strsql & "Order by Niveles.codasoc, Niveles.secuen"
      case "ENTIDADES DE CONTINGENCIA"
        redim aHeaders(0)
        redim aCampos(0)
        redim aAnchos(0)

        aHeaders(0) = "Descripcion de la Entidad"
        aCampos(0) = 1
        aAnchos(0) = "100%"
        Strsql = "Select * from Eco Order By Descri"
      case "CASOS DE CONTINGENCIA"
        redim aHeaders(2)
        redim aCampos(2)
        redim aAnchos(2)

        aHeaders(0) = "&nbsp;"
        aHeaders(1) = "Contingencia"
        aHeaders(2) = "Asociacion"
        aCampos(0) = 0
        aCampos(1) = 1
        aCampos(2) = 4
        aAnchos(0) = "1%"
        aAnchos(1) = "35%"
        aAnchos(2) = "54%"
        Strsql = "Select Tipocont.*, Asociacion.* from Tipocont, Asociacion "
        Strsql = Strsql & "Where Tipocont.codasoc = Asociacion.codasoc"
      case "ACCIONES DE CIRCUITO ALTERNATIVO"
        redim aHeaders(0)
        redim aCampos(0)
        redim aAnchos(0)

        aHeaders(0) = "Descripcion de las Acciones"
        aCampos(0) = 1
        aAnchos(0) = "100%"
        Strsql = "Select * From Tipoaccalt Order By Descri"
      case "ACCIONES DE CIRCUITO CORRECTIVO"
        redim aHeaders(0)
        redim aCampos(0)
        redim aAnchos(0)

        aHeaders(0) = "Descripcion de las Acciones"
        aCampos(0) = 1
        aAnchos(0) = "100%"
        Strsql = "Select * From Tipoacccorr Order By Descri"
      case "INFORMACION DE APOYO"
        redim aHeaders(1)
        redim aCampos(1)
        redim aAnchos(1)

        aHeaders(0) = "Descripcion de la Informacion de Apoyo"
        aHeaders(1) = "Existe"
        aCampos(0) = 1
        aCampos(1) = 2
        aAnchos(0) = "90%"
        aAnchos(1) = "10%"
        Strsql = "Select * from Iar Order by Descri"
      case "INFORMACION DE BASE"
        redim aHeaders(1)
        redim aCampos(1)
        redim aAnchos(1)

        aHeaders(0) = "Descripcion de la Informacion de Base"
        aHeaders(1) = "Existe"
        aCampos(0) = 1
        aCampos(1) = 2
        aAnchos(0) = "90%"
        aAnchos(1) = "10%"
        Strsql = "Select * from Infrespcorr Order by Descri"
      case "CENTROS DE PROCESAMIENTO"
        redim aHeaders(1)
        redim aCampos(1)
        redim aAnchos(1)

        aHeaders(0) = "Descripcion del Centro de Procesamiento"
        aHeaders(1) = "Existe"
        aCampos(0) = 1
        aCampos(1) = 2
        aAnchos(0) = "90%"
        aAnchos(1) = "10%"
        Strsql = "Select * from Cenproccorr Order by Descri"
      case "PROVEEDORES"
        redim aHeaders(1)
        redim aCampos(1)
        redim aAnchos(1)

        aHeaders(0) = "Descripcion del Proveedor"
        aHeaders(1) = "Existe"
        aCampos(0) = 1
        aCampos(1) = 2
        aAnchos(0) = "90%"
        aAnchos(1) = "10%"
        Strsql = "Select * from Provcorr Order by Descri"
      case "EQUIPOS"
        redim aHeaders(1)
        redim aCampos(1)
        redim aAnchos(1)

        aHeaders(0) = "Descripcion del Equipo"
        aHeaders(1) = "Existe"
        aCampos(0) = 1
        aCampos(1) = 2
        aAnchos(0) = "90%"
        aAnchos(1) = "10%"
        Strsql = "Select * from Equipcorr Order by Descri"
      case "TEMAS Y SUBTEMAS"
        redim aHeaders(1)
        redim aCampos(1)
        redim aAnchos(1)

        aHeaders(0) = "Temas"
        aHeaders(1) = "Subtemas"
        aCampos(0) = 1
        aCampos(1) = 4
        aAnchos(0) = "40%"
        aAnchos(1) = "60%"
        Strsql = "Select TEMA.*, PRODSERV.* From Tema, Prodserv "
        Strsql = Strsql & "Where Tema.codtema = prodserv.codtema "
        Strsql = Strsql & "Order by Tema.descri, prodserv.descri"
    end select

    orp.GenerarReporteParametros sReporte,aHeaders,Strsql,aCampos,aAnchos
  elseif sTipo <> "" then
    strwhere = ""
    if UCase(sReporte) = "MANUAL CIRCUITO ALTERNATIVO PARCIAL" then
      Tomo = Request.form("SelTomo")
      Capitulo = Request.form("Selcapitulo")
      SiNoTomo = Request.form("Tomo")
      SiNoCapitulo = Request.form("Capitulo")

      If SiNoTomo then
        if SiNoCapitulo  Then 
          strwhere = "Where ECO =" & Tomo & " And Codprodserv=" & Capitulo 
        else
          strwhere = "Where ECO =" & Tomo  
        End If
      elseIf SiNoCapitulo Then 
        strwhere = "Where Codprodserv=" & Capitulo 
      End If
    elseif UCase(sReporte) = "MANUAL CIRCUITO CORRECTIVO PARCIAL" then
      Tomo = Request.form("SelTomo")
      Capitulo = Request.form("Selcapitulo")
      SiNoTomo = Request.form("Tomo")
      SiNoCapitulo = Request.form("Capitulo")

      If SiNoTomo then
        if SiNoCapitulo  Then 
          strwhere = "Where codnivel=" & Tomo & " And Codprodserv=" & Capitulo 
        else
          strwhere = "Where codnivel=" & Tomo  
        End If
      elseIf SiNoCapitulo Then 
        strwhere = "Where Codprodserv=" & Capitulo 
      End If
    end if

    select case UCase(sTipo)
      case "PROCESOS"
        strsql = "Select * From CsProcCri"
        sTitulo = "DEFINICION DE CRITICIDAD DE PROCESOS - Hoja de Trabajo"
      case "CRITICOS"
        strsql = "Select * From CsProcCri"
        sTitulo = "DETALLE DE PROCESOS CRITICOS"
      case "NOCRITICOS"
        strsql = "Select * From CsProcNoCri"
        sTitulo = "DETALLE DE PROCESOS NO CRITICOS"
      case "PSCONTINGENCIAS"
        strsql = "Select * From CsProcSinCont"
        sTitulo = "DETALLE DE PROCESOS CRITICOS SIN CONTINGENCIA"
      case "MATRIZALTERNATIVO"
        strsql = "Select * From CsManalt Order By inda,indb,indc,"
        strsql = strsql & "prosecuen,Procdescri,tipocontin,contdescri,"
        strsql = strsql & "Ecodescri, Accprior"
        sTitulo = "MATRIZ DE CONTINGENCIAS - Circuito Alternativo"
      case "MATRIZCORRECTIVO"
        strsql = "Select * From CsManCorr Order By inda,indb,indc,"
        strsql = strsql & "prosecuen,Procdescri,tipocontin,contdescri,"
        strsql = strsql & "Asocdescri,secuen,Accprior,"
        strsql = strsql & "codaccioncorr"
        sTitulo = "MATRIZ DE CONTINGENCIAS - Circuito Correctivo"
      case "MANUALALTERNATIVO"
        strsql = "Select * From CsManalt " & strwhere & " Order By Ecodescri,inda,"
        strsql = strsql & "indb,indc,prosecuen,Procdescri,"
        strsql = strsql & "tipocontin,contdescri,Accprior"
        sTitulo = "MANUAL DE CONTINGENCIAS - Circuito Alternativo"
'response.write strsql
      case "MANUALCORRECTIVO"
        strsql = "Select * From CsManCorr " & strwhere & " Order By Asocdescri,"
        strsql = strsql & "secuen,inda,indb,indc,"
        strsql = strsql & "prosecuen,Procdescri,tipocontin,"
        strsql = strsql & "contdescri,Accprior,codaccioncorr"
        sTitulo = "MANUAL DE CONTINGENCIAS - Circuito Correctivo"

    end select

    select case UCase(sReporte)
      case "REPORTES DE PROCESOS"
        orp.GenerarReporteProcesos sTitulo, strsql
      case "MATRIZ CIRCUITO ALTERNATIVO","MATRIZ CIRCUITO CORRECTIVO"
        orp.GenerarReporteMatriz sTitulo, strsql
      case "MANUAL CIRCUITO ALTERNATIVO","MANUAL CIRCUITO CORRECTIVO","MANUAL CIRCUITO ALTERNATIVO PARCIAL","MANUAL CIRCUITO CORRECTIVO PARCIAL"
        orp.GenerarReporteManual sTitulo, strsql
    end select
  else
    select case UCase(sReporte)
      case "REPORTES DE PROCESOS"
        redim aAsp(3)
        redim aTitulo(3)

        aAsp(0) = "Reportesx.asp?Tipo=Procesos&Reporte=" & xReporte
        aAsp(1) = "Reportesx.asp?Tipo=Criticos&Reporte=" & xReporte
        aAsp(2) = "Reportesx.asp?Tipo=NoCriticos&Reporte=" & xReporte
        aAsp(3) = "Reportesx.asp?Tipo=PSContingencias&Reporte=" & xReporte
        aTitulo(0) = "Planilla de Procesos"
        aTitulo(1) = "Procesos Criticos"
        aTitulo(2) = "Procesos no Criticos"
        aTitulo(3) = "Procesos sin Contingencias"
      case "MATRIZ CIRCUITO ALTERNATIVO"
        redim aAsp(0)
        redim aTitulo(0)

        aAsp(0) = "Reportesx.asp?Tipo=MatrizAlternativo&Reporte=" & xReporte
        aTitulo(0) = "Matriz Circuito Alternativo"
      case "MATRIZ CIRCUITO CORRECTIVO"
        redim aAsp(0)
        redim aTitulo(0)

        aAsp(0) = "Reportesx.asp?Tipo=MatrizCorrectivo&Reporte=" & xReporte
        aTitulo(0) = "Matriz Circuito Correctivo"
      case "REPORTES DE PARAMETROS"
        redim aAsp(11)
        redim aTitulo(11)

        aAsp(0) = "Reportesx.asp?SubReporte=Asociaciones"
        aAsp(1) = "Reportesx.asp?SubReporte=Niveles_de_Asociacion"
        aAsp(2) = "Reportesx.asp?SubReporte=Entidades_de_Contingencia"
        aAsp(3) = "Reportesx.asp?SubReporte=Casos_de_Contingencia"
        aAsp(4) = "Reportesx.asp?SubReporte=Acciones_de_circuito_alternativo"
        aAsp(5) = "Reportesx.asp?SubReporte=Acciones_de_circuito_correctivo"
        aAsp(6) = "Reportesx.asp?SubReporte=Informacion_de_apoyo"
        aAsp(7) = "Reportesx.asp?SubReporte=Informacion_de_base"
        aAsp(8) = "Reportesx.asp?SubReporte=Centros_de_procesamiento"
        aAsp(9) = "Reportesx.asp?SubReporte=Proveedores"
        aAsp(10) = "Reportesx.asp?SubReporte=Equipos"
        aAsp(11) = "Reportesx.asp?SubReporte=Temas_y_subtemas"
        for n = 0 to 11
          aAsp(n) = aAsp(n) & "&Tipo=Parametros&Reporte=" & xReporte
        next

        aTitulo(0) = "Asociaciones"
        aTitulo(1) = "Niveles"
        aTitulo(2) = "Entidades de Contingencia"
        aTitulo(3) = "Casos de Contingencias"
        aTitulo(4) = "Acciones del Circuito Alternativo"
        aTitulo(5) = "Acciones del Circuito Correctivo"
        aTitulo(6) = "Informacion de Apoyo Circuito Alternativo"
        aTitulo(7) = "Informacion de Base Circuito Correctivo"
        aTitulo(8) = "Centros de Procesamiento"
        aTitulo(9) = "Proveedores"
        aTitulo(10) = "Equipos"
        aTitulo(11) = "Temas y Subtemas"
      case "MANUAL CIRCUITO ALTERNATIVO"
        redim aAsp(0)
        redim aTitulo(0)

        aAsp(0) = "Reportesx.asp?Tipo=ManualAlternativo&Reporte=" & xReporte
        aTitulo(0) = "Manual Circuito Alternativo"
      case "MANUAL CIRCUITO CORRECTIVO"
        redim aAsp(0)
        redim aTitulo(0)

        aAsp(0) = "Reportesx.asp?Tipo=ManualCorrectivo&Reporte=" & xReporte
        aTitulo(0) = "Manual Circuito Correctivo"
    end select

    sHTML = orp.GenerarFormReporte(sReporte,aAsp,aTitulo,"","_blank")
  end if
end if

Response.Write sHTML

%>