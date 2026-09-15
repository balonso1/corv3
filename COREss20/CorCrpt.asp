<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0

  sHTML = ofv.FormHeader("Carpetas de Descarga")

  'Encabezados
  formulario =  ofv.Encabezados("&nbsp;","Codigo","Descripcion","Activo")

  vuelta = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover)

  set cn = ofv.conectar(ofv.strconn1)



  ofv.ObtenerAtributos request.querystring("seg"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  StrSql = "Select * From CorCrpt order by CRPdesc"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


  do until rs.eof 
'    sInput = ofv.GenerarInput(rs(0).name,rs(0),"hidden","corcrpt" & rs(0),"10","dt","")
'    celdas = ofv.GenCelda("","tt colspan=2","","","","","",sinput & rs(0),"si")

    sInput = ofv.GenerarInput(rs(0).name,rs(0),"hidden","corcrpt" & rs(0),"10","dt","")
    celdas = ofv.GenCelda("","tt colspan=2","","","","","",sinput,"si")

    if bModificar then
      sInput = ofv.GenerarInput(rs(1).name,rs(1),"text","corcrpt" & rs(0),"5","dt","")
    else
      sInput = ofv.GenerarInput("",rs(1),"readonly","","5","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")

    if bModificar then
      sInput = ofv.GenerarInput(rs(2).name,rs(2),"text","corcrpt" & rs(0),"90","dt","")
    else
      sInput = ofv.GenerarInput("",rs(2),"readonly","","90","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")

    if bModificar then
      sInput = "<font class=at><select onchange=corcrpt" & rs(0) & ".submit() "
      sInput = sInput & "name=CRPEnable><option value=1>SI"
      if  rs(3)  then
        sInput = sInput & "<option value=0>NO"
      else
        sInput = sInput & "<option value=0 selected>NO"
      end if
      sInput = sInput & "</select></font>"
    else
      if rs(3)  then
         xestado = "SI"
      else
         xestado = "NO"
      end if   
      sInput = ofv.GenerarInput("",xestado,"readonly","","2","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")


     if bModificar then
      sAccion = "asp/Actualizar.asp?ID=" & rs(0) & "&volver=" & vuelta & "&Tabla=CorCrpt"
      formulario = formulario & ofv.GenForm("corcrpt" & rs(0),"",sAccion,"","post","",filas,"si")
    else
      formulario = formulario & ofv.GenForm("","","","","post","",filas,"si")
    end if

    rs.movenext
  loop

  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",formulario,"si")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML

%>