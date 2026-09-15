<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))

  'Encabezados
  formulario =  ofv.Encabezados("Nivel","Descripcion")

  vuelta = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover)
  set cn = ofv.conectar(ofv.strconn0)



  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  StrSql = "Select * From NivelConfiden order by NCCod"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


  do until rs.eof 
    sInput = ofv.GenerarInput(rs(0).name,rs(0),"hidden","niveles" & rs(0),"10","dt","")
    celdas = ofv.GenCelda("","tt colspan=2","","","","","",sinput & rs(0),"si")

    if bModificar then
      sInput = ofv.GenerarInput(rs(1).name,rs(1),"text","niveles" & rs(0),"90","dt","")
    else
      sInput = ofv.GenerarInput("",rs(1),"readonly","","90","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")

    if bModificar then
      sAccion = "asp/Actualizar.asp?ID=" & rs(0) & "&volver=" & vuelta & "&Tabla=NivelConfiden"
      formulario = formulario & ofv.GenForm("niveles" & rs(0),"",sAccion,"","post","",filas,"si")
    else
      formulario = formulario & ofv.GenForm("","","","","post","",filas,"si")
    end if

    rs.movenext
  loop

  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",formulario,"si")

  'Botonera
   Agregar = ""
   Buscar = ""
   Retornar = ""

  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,Retornar)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML

%>