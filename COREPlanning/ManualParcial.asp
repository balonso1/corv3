<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  sTipo = UCase(ofv.convertircar(request.querystring("Tipo"),"_"," "))

  set cn = ofv.conectar(ofv.strconn4)

  sHTML = ofv.FormHeader(sTipo)

  if request.querystring("seg") <> "" then
    Session("Form") = ofv.convertircar(request.querystring("seg"),"_"," ")
  end if
  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  Celda = ofv.GenCelda("","","","","60","","","","si")
  Fila = ofv.genrow("","","","","","","",Celda,"si")

  Celda = ofv.GenCelda("","tt colspan=4","center","2%","10","#993333","","TOMO","si")
  Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

  Celda = ofv.GenCelda("","lt","","2%","10","","","Activar","si")
  if bModificar then
    sInput = ofv.comboSiNOmin("Tomo")
  else
    sInput = ofv.GenerarInput("Tomo",rs("Tomo"),"text","manual","5","dt","")
  end if
  Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")
  Celda = Celda & ofv.GenCelda("","lt","","2%","10","","","Titulo","si")
  if bModificar then
    if sTipo = "MANUAL CIRCUITO ALTERNATIVO PARCIAL" then
      Strsql = "Select * From ECO Order By Descri"
      sInput = ofv.GenerarCombo("manual","SelTomo","",Strsql,cn,0,1,"no")
    else
      Strsql="Select * From CSSASOCNIVEL Order By Titulo"
      sInput = ofv.GenerarCombo("manual","SelTomo","",Strsql,cn,1,0,"no")
    end if
  else
    sInput = ofv.GenerarInput("SelTomo",rs("SelTomo"),"text","manual","60","dt","")
  end if
  Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")
  Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

  Celda = ofv.GenCelda("","","","","40","","","","si")
  Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

  Celda = ofv.GenCelda("","tt colspan=4","center","2%","10","#993333","","CAPITULO","si")
  Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

  Celda = ofv.GenCelda("","lt","","2%","10","","","Activar","si")
  if bModificar then
    sInput = ofv.comboSiNOmin("Capitulo")
  else
    sInput = ofv.GenerarInput("Capitulo",rs("Capitulo"),"text","manual","5","dt","")
  end if
  Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")
  Celda = Celda & ofv.GenCelda("","lt","","2%","10","","","Titulo","si")
  if bModificar then
    Strsql = "Select * From csstemasubtema Order By titulo"
    sInput = ofv.GenerarCombo("manual","Selcapitulo","",Strsql,cn,1,0,"no")
  else
    sInput = ofv.GenerarInput("SelCapitulo",rs("SelCapitulo"),"text","manual","80","dt","")
  end if
  Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")
  Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

  'Emitir Manual
  Celda = ofv.GenCelda("","","","","30","","","","si")
  Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

  sInput = ofv.GenerarInput("","Emitir&nbsp;Manual","submit","","10","bt","")
  Celda = ofv.GenCelda("","","center colspan=4","","","","",sInput,"si")
  Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

  sAccion = ""
  if bModificar then
    if sTipo = "MANUAL CIRCUITO ALTERNATIVO PARCIAL" then
      xTipo = "ManualAlternativo"
    else
      xTipo = "ManualCorrectivo"
    end if
    sAccion = "Reportesx.asp?Tipo=" & xTipo & "&Reporte=" & request.querystring("Tipo")
    sFunc = "onsubmit=" & chr(34) & "if (this.Tomo.options[this.Tomo.selectedIndex].value=="
    sFunc = sFunc & "'True' && this.SelTomo.options[this.SelTomo.selectedIndex].value=='nsnc')"
    sFunc = sFunc & "{ alert( 'Seleccione el titulo del Tomo');return false;}"
    sFunc = sFunc & " if (this.Capitulo.options[this.Capitulo.selectedIndex].value=="
    sFunc = sFunc & "'True' && this.Selcapitulo.options[this.Selcapitulo.selectedIndex].value=="
    sFunc = sFunc & "'nsnc') { alert( 'Seleccione el titulo del Capitulo');return false;}"
    sFunc = sFunc & " if (this.Tomo.options[this.Tomo.selectedIndex].value=='False' && "
    sFunc = sFunc & "this.Capitulo.options[this.Capitulo.selectedIndex].value=='False') "
    sFunc = sFunc & "{ alert( 'Debe Activar al menos una Opcion');return false;}" & chr(34)
  end if
  Formulario = Formulario & ofv.GenForm("manual","",sAccion,"_blank","post",sFunc,Fila,"si")

  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>