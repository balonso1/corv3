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

  vuelta = ofv.vueltaasp("../",0,0)

  ofv.ObtenerAtributos request.querystring("seg"),""
  Session("Form") = request.querystring("seg")

  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  celdas = ofv.GenCelda("","tt colspan=3","center","","","","","Administracion de Menues","si")
  formulario = "<thead>" & ofv.GenRow("","","","","10","","",celdas,"si")
  celdas = ofv.GenCelda("","tt colspan=3","center","","","","","Aplicaciones","si")
  formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si") & "</thead>"

  set cn = ofv.conectar(ofv.strconn1)

  Strsql = "Select * From CorAplic where enable<>0 Order By OrdenApli, descapli"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  ix = 0

  do until rs.eof

    sAplicacion = UCase(rs(1) & " - " & ofv.convertircar(rs(2),"*"," ") & rs(4))

    sForm = "AtributosA" & rs(0)

    if bModificar then
      sInput = ofv.GenerarInput(rs(13).name,rs(13),"text",sForm,"5","dtn","SI")
    else
      sInput = ofv.GenerarInput("",rs(13),"readonly","","5","dtn","")
    end if

    celdas = ofv.GenCelda("","","","","","","",sInput,"si")



    if bModificar then
       sAccion = "asp/ActualizarMenues.asp?volver=" & vuelta & "&ID=" & rs(0) & "&tabla=Coraplic" 
       celdas = ofv.GenForm(sForm,"",sAccion,"","post","",celdas,"si")
    else
       celdas = ofv.GenForm("","","","","post","",celdas,"si")
    end if

    Celdas = Celdas & ofv.GenCelda("","ut","","90%","","","",sAplicacion,"si")


    sInput = ofv.generarinput("","Solapas","submit","Menus","10","bt","")
    celdax = ofv.GenCelda("","","","","","","",sInput,"si")

    sAccion = "ManMMenues.asp?Apl=" & rs(1)
    sAccion = sAccion & "&Nombre=" & ofv.convertircar(sAplicacion," ","_")
    sAccion = sAccion & "&DD=x&DS=x"
    Celdas = Celdas & ofv.genform("Menus","",sAccion,"","post","",Celdax,"si")

    formulario = formulario & ofv.GenRow("","","","","10","","",celdas,"si")


  rs.movenext
  loop

  tabla = ofv.GenTabla("","","","80%","","0","0","0","0","","",formulario,"si")

  sHTML = sHTML & "<center>" & tabla & "</center></body></html>"

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
end if

response.write sHTML

%>