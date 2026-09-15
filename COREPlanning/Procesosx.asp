<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  sTipo = UCase(request.querystring("Tipo"))
  sFiltrar = request.querystring("Filtrar")
  sTema = request.querystring("Tema")
  sSubtema = request.querystring("Subtema")
  sEvento = request.querystring("Evento")
  sProceso = request.querystring("Proceso")
  sContingencia = request.querystring("Contingencia")
  sAlternativo = request.querystring("Alternativo")
  sCorrectivo = request.querystring("Correctivo")
  sAsociacion = request.querystring("Asociacion")
  xTema = ofv.convertircar(request.querystring("xTema"),"*"," ")
  xSubtema = ofv.convertircar(request.querystring("xSubtema"),"*"," ")
  xEvento = ofv.convertircar(request.querystring("xEvento"),"*"," ")
  xProceso = ofv.convertircar(request.querystring("xProceso"),"*"," ")
  xContingencia = ofv.convertircar(request.querystring("xContingencia"),"*", " ")
  xAlternativo = ofv.convertircar(request.querystring("xAlternativo"),"*", " ")
  xCorrectivo = ofv.convertircar(request.querystring("xCorrectivo"),"*", " ")
  xAsociacion = ofv.convertircar(request.querystring("xAsociacion"),"*", " ")
  nextTipo = ""
  prevTipo = ""
  CantRegMostrarx = 10
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))

  sVuelta = "Procesosx.asp?Tipo=" & sTipo & "_Filtrar=" & sFiltrar & "_Tema=" & sTema
  sVuelta = sVuelta & "_Subtema=" & sSubtema & "_Evento=" & sEvento & "_Proceso=" & sProceso
  sVuelta = sVuelta & "_Contingencia=" & sContingencia & "_Alternativo=" & sAlternativo
  sVuelta = sVuelta & "_Correctivo=" & sCorrectivo & "_Asociacion=" & sAsociacion
  sVuelta = sVuelta & "_xTema=" & xTema & "_xSubtema=" & xSubtema & "_xEvento=" & xEvento
  sVuelta = sVuelta & "_xProceso=" & xProceso & "_xContingencia=" & xContingencia
  sVuelta = sVuelta & "_xAlternativo=" & xAlternativo & "_xCorrectivo=" & xCorrectivo
  sVuelta = sVuelta & "_xAsociacion=" & xAsociacion
  sVuelta2 = "../" & sVuelta & "_CantRegMover=" & CantRegMoverx

  dim aInput()
  dim aInput2()
  dim aInput3()
  dim aInput4()
  dim aInput5()
  dim aEncabezado0,aEncabezado1,aEncabezado2,aEncabezado3,aEncabezado4,aEncabezado5,aEncabezado6

  sBoton2 = ""
  nextTipo2 = ""
  select case sTipo
    case "CIRCUITOS"
      aEncabezado0 = "Tema"
      aEncabezado1 = "Orden"
      sTabla = "Tema"
      sBoton = "Subtemas"
      sASPAgregar = "AgregarCircuito.asp?"
      Strsql = "Select * From Tema Order By Inda"
      nClave = 2
      nextTipo = "Subtemas"
      Str1 = "Select descri From Tema Order By inda"
      Str2 = "Select descri From Tema Order By descri"
    case "SUBTEMAS"
      aEncabezado0 = "Subtema"
      aEncabezado1 = "Orden"
      sTabla = "Prodserv"
      sBoton = "Eventos"
      sASPAgregar = "AgregarSubtemas.asp?ID=" & sTema & "&IDx=" & ofv.convertircar(xTema," ","*")
      sASPAgregar = sASPAgregar & "&"
      Strsql = "Select * From Prodserv Where CodTema=" & sTema & " Order By Indb"
      nClave = 3
      prevTipo = "Circuitos"
      nextTipo = "Eventos"
      Str1 = "Select descri From Prodserv Where Codtema =" & sTema & " Order By Indb"
      Str2 = "Select descri From Prodserv Where Codtema =" & sTema & " Order By descri" 
    case "EVENTOS"
      aEncabezado0 = "Evento"
      aEncabezado1 = "Secuencia"
      aEncabezado2 = "Orden"
      sTabla = "Evento"
      sBoton = "Procesos"
      sASPAgregar = "AgregarEventos.asp?ID=" & sSubtema & "&IDx="
      sASPAgregar = sASPAgregar & ofv.convertircar(xSubtema," ","*")
      sASPAgregar = sASPAgregar & "&Tema=" & sTema & "&xTema=" & ofv.convertircar(xTema," ","*")
      sASPAgregar = sASPAgregar & "&"
      Strsql = "Select * From Evento Where Codprodserv =" & sSubtema & " Order By Indc"
      nClave = ""
      prevTipo = "Subtemas"
      nextTipo = "Procesos"
      Str1 = "Select descri from Evento Where Codprodserv = " & sSubtema & " Order By Indc"
      Str2 = "Select descri from Evento Where Codprodserv = " & sSubtema & " Order By descri"
    case "PROCESOS"
      aEncabezado0 = "Proceso"
      aEncabezado1 = "Sec."
      aEncabezado2 = "Cri."
      aEncabezado3 = "Pri."
      sTabla = "Proceso"
      sBoton = "Contingencias"
      sASPAgregar = "AgregarProcesos.asp?ID=" & sEvento & "&IDx="
      sASPAgregar = sASPAgregar & ofv.convertircar(xEvento," ","*")
      sASPAgregar = sASPAgregar & "&Tema=" & sTema & "&xTema=" & ofv.convertircar(xTema," ","*")
      sASPAgregar = sASPAgregar & "&Subtema=" & sSubtema & "&xSubtema="
      sASPAgregar = sASPAgregar & ofv.convertircar(xSubtema," ","*") & "&"
      nucleo= " From Proceso Where Codevento =" & sEvento
      If sFiltrar = "cri" then
        Strsql = "Select * " & nucleo & " and critico= 's' Order By Prosecuen "
        Str1 = "Select descri " & nucleo & " and critico='s' Order By Prosecuen "
        Str2 = "Select descri " & nucleo & " and critico='s' Order By Descri " 
      ElseIf sFiltrar = "nocri" then
        Strsql = "Select * " & nucleo & " and critico='n' Order By Prosecuen "
        Str1 = "Select descri " & nucleo & " and critico='n' Order By Prosecuen "
        Str2 = "Select descri " & nucleo & " and critico='n' Order By Descri " 
      else
        Strsql = "Select * " & nucleo & " Order By Prosecuen "
        Str1 = "Select descri " & nucleo & " Order By Prosecuen "
        Str2 = "Select descri " & nucleo & " Order By Descri "
      End If
      nClave = ""
      prevTipo = "Eventos"
      nextTipo = "CircContingencia"

      Redim Perfiles(2)
      Perfiles(1) = "N"
      Perfiles(2) = "S"
      Redim Perfiles2(3)
      Perfiles2(1) = "A"
      Perfiles2(2) = "M"
      Perfiles2(3) = "B"
    case "CIRCCONTINGENCIA"
      aEncabezado0 = "Contingencia"
      aEncabezado1 = ""
      sTabla = "Contingencia"
      sBoton = "Alternativo"
      sBoton2 = "Correctivo"
      sASPAgregar = "AgregarCircContingencia.asp?ID=" & sProceso & "&IDx="
      sASPAgregar = sASPAgregar & ofv.convertircar(xProceso," ","*")
      sASPAgregar = sASPAgregar & "&Tema=" & sTema & "&xTema=" & ofv.convertircar(xTema," ","*")
      sASPAgregar = sASPAgregar & "&Subtema=" & sSubtema & "&xSubtema="
      sASPAgregar = sASPAgregar & ofv.convertircar(xSubtema," ","*")
      sASPAgregar = sASPAgregar & "&Evento=" & sEvento & "&xEvento="
      sASPAgregar = sASPAgregar & ofv.convertircar(xEvento," ","*") & "&"
      Strsql = "Select Contingencia.*, tipocont.descri From Contingencia,tipocont "
      Strsql = Strsql & "Where Contingencia.tipocontin=Tipocont.tipocontin "
      Strsql = Strsql & "And Contingencia.Codproceso = " & sProceso
      nClave = ""
      prevTipo = "Procesos"
      nextTipo = "Alternativo"
      nextTipo2 = "Correctivo"
      Str1 = "Select TIPOCONT.Descri From CONTINGENCIA,TIPOCONT "
      Str1 = Str1 & "Where CONTINGENCIA.Codproceso =" & sProceso
      Str1 = Str1 & " And CONTINGENCIA.Tipocontin = TIPOCONT.Tipocontin"
      Str2 = Str1 & " Order By descri"
    case "ALTERNATIVO"
      aEncabezado0 = "Entidad de Contingencia"
      sTabla = "CircAlt"
      sBoton = "Procedimiento"
      sASPAgregar = "AgregarAlternativo.asp?ID=" & sContingencia & "&IDx="
      sASPAgregar = sASPAgregar & ofv.convertircar(xContingencia," ","*")
      sASPAgregar = sASPAgregar & "&Tema=" & sTema & "&xTema=" & ofv.convertircar(xTema," ","*")
      sASPAgregar = sASPAgregar & "&Subtema=" & sSubtema & "&xSubtema="
      sASPAgregar = sASPAgregar & ofv.convertircar(xSubtema," ","*")
      sASPAgregar = sASPAgregar & "&Evento=" & sEvento & "&xEvento="
      sASPAgregar = sASPAgregar & ofv.convertircar(xEvento," ","*")
      sASPAgregar = sASPAgregar & "&Proceso=" & sProceso & "&xProceso="
      sASPAgregar = sASPAgregar & ofv.convertircar(xProceso," ","*") & "&"
      Strsql = "Select Circalt.*, Eco.descri From Circalt, Eco Where Circalt.eco = Eco.eco "
      Strsql = Strsql & "And Circalt.Codcontin=" & sContingencia
      nClave = ""
      prevTipo = "CircContingencia"
      nextTipo = "Procedimiento"
      Str1 = "Select eco.descri From Circalt,eco Where Codcontin=" & sContingencia
      Str1 = Str1 & " And circalt.eco = eco.eco"
      Str2 = Str1 & " Order By eco.descri"
    case "CORRECTIVO"
      aEncabezado0 = "Nivel"
      aEncabezado1 = "Centro Procesamiento"
      sTabla = "Circorr"
      sBoton = "Procedimiento"
      sASPAgregar = "AgregarCorrectivo.asp?ID=" & sContingencia & "&IDx="
      sASPAgregar = sASPAgregar & ofv.convertircar(xContingencia," ","*")
      sASPAgregar = sASPAgregar & "&Tema=" & sTema & "&xTema=" & ofv.convertircar(xTema," ","*")
      sASPAgregar = sASPAgregar & "&Subtema=" & sSubtema & "&xSubtema="
      sASPAgregar = sASPAgregar & ofv.convertircar(xSubtema," ","*")
      sASPAgregar = sASPAgregar & "&Evento=" & sEvento & "&xEvento="
      sASPAgregar = sASPAgregar & ofv.convertircar(xEvento," ","*")
      sASPAgregar = sASPAgregar & "&Proceso=" & sProceso & "&xProceso="
      sASPAgregar = sASPAgregar & ofv.convertircar(xProceso," ","*")
      sASPAgregar = sASPAgregar & "&Asociacion=" & sAsociacion & "&xAsociacion="
      sASPAgregar = sASPAgregar & ofv.convertircar(xAsociacion," ","*") & "&"
      Strsql = "Select * From Circorr Where Codcontin = " & sContingencia
      nClave = ""
      prevTipo = "CircContingencia"
      nextTipo = "Procedimiento"
      Str1 = "Select NIVELES.Descri From CIRCORR,NIVELES Where CIRCORR.Codcontin = "
      Str1 = Str1 & sContingencia & " And CIRCORR.Codnivel = NIVELES.Codnivel"
      Str2 = Str1 & " Order By NIVELES.Descri"
  end select

  sTitulo = sTipo
  if sTitulo = "CIRCCONTINGENCIA" then sTitulo = "CIRCUITO DE CONTINGENCIA"
  if sTitulo = "ALTERNATIVO" then sTitulo = "CIRCUITO ALTERNATIVO"
  if sTitulo = "CORRECTIVO" then sTitulo = "CIRCUITO CORRECTIVO"
  sHTML = ofv.FormHeader(sTitulo)

  if request.querystring("seg") <> "" then
    Session("Form") = ofv.convertircar(request.querystring("seg"),"_"," ")
  end if
  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn4)

  if nClave <> "" then
    Clave = ofv.vectorclave(Strsql,nClave)
    sHTML = sHTML & ofv.ValClaveRep(Clave)
  end if

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then
    sHTML = sHTML & "<p align=center><font face=tahoma size=1 color=#990000><b>"
    sHTML = sHTML & "NO HAY ELEMENTOS ASOCIADOS</b><br><br></FONT></P>"
  Else 
    Fila = ""
    if xTema <> "" then
      nAncho = "90%"
      if xEvento <> "" then nAncho = "40%"
      Celda = ofv.GenCelda("","tx","","10%","10","","","Tema","si")
      Celda = Celda & ofv.GenCelda("","cc","",nAncho,"10","","",xTema,"si")
      if nAncho = "90%" then
        Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")
        Celda = ""
      end if
      if xSubtema <> "" then
        Celda = Celda & ofv.GenCelda("","tx","","10%","10","","","Subtema","si")
        Celda = Celda & ofv.GenCelda("","cc","",nAncho,"10","","",xSubtema,"si")
        Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")
      end if
      if xEvento <> "" then
        Celda = ofv.GenCelda("","tx","","10%","10","","","Evento","si")
        Celda = Celda & ofv.GenCelda("","cc","",nAncho,"10","","",xEvento,"si")
        if xProceso <> "" then
          Celda = Celda & ofv.GenCelda("","tx","","10%","10","","","Proceso","si")
          Celda = Celda & ofv.GenCelda("","cc","",nAncho,"10","","",xProceso,"si")
        end if
        Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")
      end if
      if xContingencia <> "" then
        Celda = ofv.GenCelda("","tx","","10%","10","","","Contingencia","si")
        Celda = Celda & ofv.GenCelda("","cc","",nAncho,"10","","",xContingencia,"si")
        if xAsociacion <> "" And sTipo <> "ALTERNATIVO" then
          Celda = Celda & ofv.GenCelda("","tx","","10%","10","","","Asociacion","si")
          Celda = Celda & ofv.GenCelda("","cc","",nAncho,"10","","",xAsociacion,"si")
        end if
        Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")
      end if
      sHTML = sHTML & ofv.GenTabla("","","","93%","","0","","0","0","","",Fila,"si")
    end if

    'Encabezados
    if aEncabezado1 = "" then
      Formulario = ofv.Encabezados(aEncabezado0)
    elseif aEncabezado2 = "" then
      Formulario = ofv.Encabezados(aEncabezado0,aEncabezado1)
    elseif aEncabezado3 = "" then
      Formulario = ofv.Encabezados(aEncabezado0,aEncabezado1,aEncabezado2)
    elseif aEncabezado4 = "" then
      Formulario = ofv.Encabezados(aEncabezado0,aEncabezado1,aEncabezado2,aEncabezado3)
    elseif aEncabezado5 = "" then
      Formulario = ofv.Encabezados(aEncabezado0,aEncabezado1,aEncabezado2,aEncabezado3,aEncabezado4)
    elseif aEncabezado6 = "" then
      Formulario = ofv.Encabezados(aEncabezado0,aEncabezado1,aEncabezado2,aEncabezado3,aEncabezado4,aEncabezado5)
    else
      Formulario = ofv.Encabezados(aEncabezado0,aEncabezado1,aEncabezado2,aEncabezado3,aEncabezado4,aEncabezado5,aEncabezado6)
    end if

    rs.move(CantRegMoverx)
    posi=0
    Do While Not rs.EOF and posi < CantRegMostrarx
      sForm = "form" & posi
      sForm1 = "boton" & posi
      sForm1b = "botonb" & posi

      'Eliminar
      if bSuprimir then
        sHref = "asp/Eliminar.asp?ID=" & ofv.convertircar(rs(0)," ","_")
        sHref = sHref & "&Tabla=" & sTabla & "&volver=" & sVuelta2 & "&Eliminar=NSNC"
        nCampDesc = 1
        if sTipo = "CIRCCONTINGENCIA" or sTipo = "ALTERNATIVO" then nCampDesc = 3
        Celda = ofv.BotonEliminar(rs(nCampDesc),sHref)
      else
        Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
      end if

      select case sTipo
        case "CIRCUITOS"
          sTema = rs(0)
          xTema = ofv.ConvertirCar(rs(1)," ","*")
          redim aInput(1)
          aInput(0) = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"110","dt","")
          aInput(1) = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"10","dt","SIR")
        case "SUBTEMAS"
          sSubTema = rs(0)
          xSubtema = ofv.ConvertirCar(rs(1)," ","*") 
          redim aInput(1)
          aInput(0) = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"110","dt","")
          aInput(1) = ofv.GenerarInput(rs(3).name,rs(3),"text",sForm,"10","dt","SIR")
        case "EVENTOS"
          sEvento = rs(0)
          xEvento = ofv.ConvertirCar(rs(1)," ","*") 
          redim aInput(2)
          aInput(0) = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"100","dt","")
          aInput(1) = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"10","dt","SI")
          aInput(2) = ofv.GenerarInput(rs(4).name,rs(4),"text",sForm,"10","dt","SI")
        case "PROCESOS"
          sProceso = rs(0)
          xProceso = ofv.ConvertirCar(rs(1)," ","*") 
          redim aInput(3)
          aInput(0) = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"100","dt","")
          aInput(1) = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"5","dt","SI")
          Combo = "<select onchange=document." & sForm & ".submit() name=" & rs(3).name & ">"
          For n = 1 To Ubound(Perfiles)
            If Perfiles(n) = rs(3) Then
              Combo = Combo & "<option value=" & Perfiles(n) & " selected>" & Perfiles(n)
            Else
              Combo = Combo & "<option value=" & Perfiles(n) & ">" & Perfiles(n)
            End If
          Next
          Combo = Combo & "</select></font>"
          aInput(2) = Combo
          Combo = "<select onchange=document." & sForm & ".submit() name=" & rs(4).name & ">"
          For n = 1 To Ubound(Perfiles2)
            If Perfiles2(n) = rs(4) Then
              Combo = Combo & "<option value=" & Perfiles2(n) & " selected>" & Perfiles2(n)
            Else
              Combo = Combo & "<option value=" & Perfiles2(n) & ">" & Perfiles2(n)
            End If
          Next
          Combo = Combo & "</select></font>"
          aInput(3) = Combo
        case "CIRCCONTINGENCIA"
          Strsqlx = "Select Tipocont.*, Asociacion.descri From Tipocont, Asociacion" 
          Strsqlx = Strsqlx & " Where Asociacion.codasoc = Tipocont.codasoc And "
          Strsqlx = Strsqlx & "Tipocont.Tipocontin = " & rs(1) & " Order By Tipocont.descri"
          set rsx = ofv.crearconsultaEx(StrSqlx,cn,1,parametros)
          if not rsx.eof then
            sContingencia = rs(0)
            xContingencia = ofv.ConvertirCar(rsx(1)," ","*") 
            sAsociacion = rsx(2)
            xAsociacion = ofv.ConvertirCar(rsx(3)," ","*") 
          end if
          call ofv.cerrarconsulta(rsx)
          redim aInput(0)
          Strsql = "Select * From Tipocont" 
          aInput(0) = ofv.GenerarCombo(sForm,rs(1).name,rs(1),Strsql,cn,0,1,"")
        case "ALTERNATIVO"
          sAlternativo = rs(0)
          xAlternativo = ofv.ConvertirCar(rs(1)," ","*") 
          redim aInput(0)
          Strsql = "Select * From Eco Order By descri" 
          aInput(0) = ofv.GenerarCombo(sForm,rs(1).name,rs(1),Strsql,cn,0,1,"")
        case "CORRECTIVO"
          sCorrectivo = rs(0)
          xCorrectivo = ofv.ConvertirCar(rs(1)," ","*") 
          redim aInput(1)
          Strsql = "Select * From Niveles Where codasoc = " & sAsociacion & " Order By secuen"
          aInput(0) = ofv.GenerarCombo(sForm,rs(1).name,rs(1),Strsql,cn,0,1,"")
          Strsql = "Select * From CENPROCCORR Order By descri"
          aInput(1) = ofv.GenerarCombo(sForm,rs(7).name,rs(7),Strsql,cn,0,1,"sb0")
          redim aInput2(2)
          aInput2(0) = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"2","dt","SI")
          Strsql = "Select * From Provcorr Order By descri"
          aInput2(1) = ofv.GenerarCombo(sForm,rs(3).name,rs(3),Strsql,cn,0,1,"sb0")
          Strsql = "Select * From Provcorr Order By descri"
          aInput2(2) = ofv.GenerarCombo(sForm,rs(4).name,rs(4),Strsql,cn,0,1,"sb0")
          redim aInput3(2)
          aInput3(0) = "&nbsp;"
          Strsql = "Select * From Equipcorr Order By descri"
          aInput3(1) = ofv.GenerarCombo(sForm,rs(5).name,rs(5),Strsql,cn,0,1,"sb0")
          Strsql = "Select * From Equipcorr Order By descri"
          aInput3(2) = ofv.GenerarCombo(sForm,rs(6).name,rs(6),Strsql,cn,0,1,"sb0")
          redim aInput4(2)
          aInput4(0) = "&nbsp;"
          Strsql = "Select * From INFRESPCORR Order By descri"
          aInput4(1) = ofv.GenerarCombo(sForm,rs(8).name,rs(8),Strsql,cn,0,1,"sb0")
          Strsql = "Select * From INFRESPCORR Order By descri"
          aInput4(2) = ofv.GenerarCombo(sForm,rs(9).name,rs(9),Strsql,cn,0,1,"sb0")
          redim aInput5(2)
          aInput5(0) = "&nbsp;"
          Strsql = "Select * From INFRESPCORR Order By descri"
          aInput5(1) = ofv.GenerarCombo(sForm,rs(10).name,rs(10),Strsql,cn,0,1,"sb0")
          Strsql = "Select * From INFRESPCORR Order By descri"
          aInput5(2) = ofv.GenerarCombo(sForm,rs(11).name,rs(11),Strsql,cn,0,1,"sb0")
      end select

      for i = 0 to ubound(aInput)
        Celda = Celda & ofv.GenCelda("","","","","","","",aInput(i),"si")
      next
      Fila = Celda
      Celda = ""
      on error resume next
      for i = 0 to ubound(aInput2)
        if i = 0 then Celda = Celda & ofv.genfrow()
        Celda = Celda & ofv.GenCelda("","","","","","","",aInput2(i),"si")
      next
      Fila = Fila & Celda
      Celda = ""
      for i = 0 to ubound(aInput3)
        if i = 0 then Celda = Celda & ofv.genfrow()
        Celda = Celda & ofv.GenCelda("","","","","","","",aInput3(i),"si")
      next
      Fila = Fila & Celda
      Celda = ""
      for i = 0 to ubound(aInput4)
        if i = 0 then Celda = Celda & ofv.genfrow()
        Celda = Celda & ofv.GenCelda("","","","","","","",aInput4(i),"si")
      next
      Fila = Fila & Celda
      Celda = ""
      for i = 0 to ubound(aInput5)
        if i = 0 then Celda = Celda & ofv.genfrow()
        Celda = Celda & ofv.GenCelda("","","","","","","",aInput5(i),"si")
      next
      Fila = Fila & Celda

      sAccion = ""
      if bModificar then
        sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
        sAccion = sAccion & "&tabla=" & sTabla & "&volver=" & sVuelta2
      end if
      Formulario = Formulario & ofv.GenForm(SForm,"",sAccion,"","post","",Fila,"si")

      sInput = ofv.GenerarInput("",sBoton,"submit",sForm1,"10","bt","")
      Celda = ofv.GenCelda("","","","","","","",sInput,"si")
      if sBoton2 = "" then
        Celda = Celda & ofv.gencelda("","","","50%","","","","&nbsp;","si")
        Celda = Celda & ofv.genfrow()
      end if

      sAccion = "Procesosx.asp?Tipo=" & nextTipo & "&Tema=" & sTema & "&Subtema=" & sSubtema
      sAccion = sAccion & "&Evento=" & sEvento & "&Proceso=" & sProceso & "&Contingencia="
      sAccion = sAccion & sContingencia & "&Alternativo=" & sAlternativo & "&Correctivo="
      sAccion = sAccion & sCorrectivo & "&Asociacion=" & sAsociacion & "&CantRegMover="
      sAccion = sAccion & CantRegMoverx & "&xTema=" & xTema & "&xSubtema=" & xSubtema
      sAccion = sAccion & "&xEvento=" & xEvento & "&xProceso=" & xProceso & "&xContingencia="
      sAccion = sAccion & xContingencia & "&xAlternativo=" & xAlternativo & "&xCorrectivo="
      sAccion = sAccion & xCorrectivo & "&xAsociacion=" & xAsociacion
      Formulario = Formulario & ofv.GenForm(sForm1,"",sAccion,"","post","",celda,"si")

      if sBoton2 <> "" then
        sInput = ofv.GenerarInput("",sBoton2,"submit",sForm1b,"10","bt","")
        Celda = ofv.GenCelda("","","","","","","",sInput,"si")

        Celda = Celda & ofv.gencelda("","","","50%","","","","&nbsp;","si")
        Celda = Celda & ofv.genfrow()

        sAccion = "Procesosx.asp?Tipo=" & nextTipo2 & "&Tema=" & sTema & "&Subtema=" & sSubtema
        sAccion = sAccion & "&Evento=" & sEvento & "&Proceso=" & sProceso & "&Contingencia="
        sAccion = sAccion & sContingencia & "&Alternativo=" & sAlternativo & "&Correctivo="
        sAccion = sAccion & sCorrectivo & "&Asociacion=" & sAsociacion & "&CantRegMover="
        sAccion = sAccion & CantRegMoverx & "&xTema=" & xTema & "&xSubtema=" & xSubtema
        sAccion = sAccion & "&xEvento=" & xEvento & "&xProceso=" & xProceso & "&xContingencia="
        sAccion = sAccion & xContingencia & "&xAlternativo=" & xAlternativo & "&xCorrectivo="
        sAccion = sAccion & xCorrectivo & "&xAsociacion=" & xAsociacion
        Formulario = Formulario & ofv.GenForm(sForm1b,"",sAccion,"","post","",celda,"si")
      end if
      posi=posi+1
      rs.Movenext
    Loop

    sHTML = sHTML & ofv.GenTabla("","","","80% border","","0","","0","0","","",Formulario,"si")
  End If

  'Botonera
  Filtro = ""
  if sTipo = "PROCESOS" then
    vuelta1 = "Procesosx.asp?Tipo=Procesos&Tema=" & sTema & "&Subtema=" & sSubtema
    vuelta1 = vuelta1 & "&Evento=" & sEvento & "&Proceso=" & sProceso & "&Contingencia="
    vuelta1 = vuelta1 & sContingencia & "&Alternativo=" & sAlternativo & "&Correctivo="
    vuelta1 = vuelta1 & sCorrectivo & "&Asociacion=" & sAsociacion & "&CantRegMover="
    vuelta1 = vuelta1 & CantRegMoverx & "&xTema=" & ofv.convertircar(xTema," ","*")
    vuelta1 = vuelta1 & "&xSubtema=" & ofv.convertircar(xSubtema," ","*") & "&xEvento="
    vuelta1 = vuelta1 & ofv.convertircar(xEvento," ","*") & "&xProceso="
    vuelta1 = vuelta1 & ofv.convertircar(xProceso," ","*") & "&xContingencia="
    vuelta1 = vuelta1 & ofv.convertircar(xContingencia," ","*") & "&xAlternativo="
    vuelta1 = vuelta1 & ofv.convertircar(xAlternativo," ","*") & "&xCorrectivo="
    vuelta1 = vuelta1 & ofv.convertircar(xCorrectivo," ","*") & "&xAsociacion="
    vuelta1 = vuelta1 & ofv.convertircar(xAsociacion," ","*")
    Filtro = ofv.Combofiltrar(cantregmostrarx,CantRegMoverx,vuelta1,sFiltrar)
  end if

  svolv = ""
  if prevTipo <> "" then
    select case sTipo
      case "SUBTEMAS"
        xTema = ""
        xSubtema = ""
        xEvento = ""
        xProceso = ""
        xContingencia = ""
        xAlternativo = ""
        xCorrectivo = ""
        xAsociacion = ""
      case "EVENTOS"
        xSubtema = ""
        xEvento = ""
        xProceso = ""
        xContingencia = ""
        xAlternativo = ""
        xCorrectivo = ""
        xAsociacion = ""
      case "PROCESOS"
        xEvento = ""
        xProceso = ""
        xContingencia = ""
        xAlternativo = ""
        xCorrectivo = ""
        xAsociacion = ""
      case "CIRCCONTINGENCIA"
        xProceso = ""
        xContingencia = ""
        xAlternativo = ""
        xCorrectivo = ""
        xAsociacion = ""
      case "ALTERNATIVO", "CORRECTIVO"
        xContingencia = ""
        xAlternativo = ""
        xCorrectivo = ""
        xAsociacion = ""
    end select

    sInput = ofv.GenerarInput("","Volver","submit","Volver","10","bt","")
    Celda =  ofv.GenCelda("","","","","","","",sInput,"si")
    sAccion = "Procesosx.asp?Tipo=" & prevTipo & "&Tema=" & sTema & "&Subtema=" & sSubtema
    sAccion = sAccion & "&Evento=" & sEvento & "&Proceso=" & sProceso & "&Contingencia="
    sAccion = sAccion & sContingencia & "&Alternativo=" & sAlternativo & "&Correctivo="
    sAccion = sAccion & sCorrectivo & "&Asociacion=" & sAsociacion & "&CantRegMover="
    sAccion = sAccion & CantRegMoverx & "&xTema=" & xTema & "&xSubtema=" & xSubtema
    sAccion = sAccion & "&xEvento=" & xEvento & "&xProceso=" & xProceso & "&xContingencia="
    sAccion = sAccion & xContingencia & "&xAlternativo=" & xAlternativo & "&xCorrectivo="
    sAccion = sAccion & xCorrectivo & "&xAsociacion=" & xAsociacion
    svolv = ofv.GenForm("Volver","",sAccion,"","post","",celda,"si")
  end if

  Agregar = ""
  if bAgregar then
    sAccion = sASPAgregar & "Regisvuelta=" & CantRegMoverx & "&Opcion=si&tabla=" & sTabla
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  sVuelta = ofv.convertircar(sVuelta,"_","&")
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,sVuelta)

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","",Agregar,svolv,Buscar,Filtro)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>