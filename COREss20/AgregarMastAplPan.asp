<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  regisvuelta = request.querystring("regisvuelta")
  Tabla = request.querystring("Tabla")
  sTipo = request.querystring("sTipo")
  sTipoD = request.querystring("sTipoD")
  sdepen = request.querystring("sdepen")
  sapli = request.querystring("sapli")

  Volver = request.querystring("Volver")
  Volver2 = replace(volver,"../","")

  set cn = ofv.conectar(ofv.strconn0)

  sid = 0
  sid2 = 0
  sor = 0
  strsql = "Select max(ID) From " & tabla 
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
  if not rs.eof then
     if not isnull(rs(0)) then sid = clng(rs(0))
     sid = gtkey(sid + 1)
  end if
  call ofv.cerrarconsulta(rs)

  sid2 = "CLI-" & stipo & cstr(sid)
  sid = "CLI" & stipo & cstr(sid)





  strsql = "Select * From " & tabla 
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

'  if ofv.Explorador = "MSIE" then
'    sHTML = ofv.MenuHeader("#cccccc","")
'  else
'    sHTML = ofv.MenuHeader("#ccbbaa","")
'  end if

  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))

sHTML = sHTML & "<script language=JavaScript>" &  chr(13) 
sHTML = sHTML & "function Gener() " &  chr(13) 
sHTML = sHTML & "{ "  &  chr(13) 
sHTML = sHTML &  "  document.all.AgregarMastApl.submit();" & chr(13)
sHTML = sHTML &  " }" & chr(13)
sHTML = sHTML & "</script>" &  chr(13) 



  f1a = "asp/AgregarMastAplPan2.asp?Tabla=" & Tabla
  f1a = f1a & "&volver=" & Volver & "&codapli=" & sapli & "&alias=" & sid 
  f1a = f1a & "&ATRM=" & sid2 & "&depen=" & sdepen & "&stipo=" & stipo

  sHTML = sHTML & "<center>" & ofv.GenForm("AgregarMastApl","",f1a,"","post","","","no")


  sHTML = sHTML & ofv.GenTabla("panel","","","80%","",0,0,0,0,"","","","no")

  celdas = ofv.GenCelda("","cc colspan=2","","","","","","Agregar " & stipoD,"si")

  sHTML = sHTML & ofv.GenRow("","","","","","","",celdas,"si")

  celdas = ofv.GenCelda("","aaa colspan=2","","","","","","&nbsp;","si")

  sHTML = sHTML & ofv.GenRow("","","","","","","",celdas,"si")

  'Encabezado Accion
    celdas = ofv.GenCelda("","tt","","2%","","","","Tipo: ","si")
    if stipo = "X" then
       sInput = "<font class=at><select  "
       sInput = sInput & "name=tipo ><option value=F>FORM"
       sInput = sInput & "<option value=N>NODO"
       sInput = sInput & "</select></font>"       
    else
      sInput = "NODO" & ofv.GenerarInput("tipo","N","hidden","","1","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","cc","","","","","",sInput,"si")

  sHTML = sHTML & ofv.GenRow("","","","","","","",celdas,"si")


  'Encabezado Accion
    celdas = ofv.GenCelda("","tt","","2%","","","","Descripcion: ","si")
    sInput = ofv.GenerarInput("observaciones","","agregar","","100","dt","")
    celdas = celdas & ofv.GenCelda("","cc","","","","","",sInput,"si")

  sHTML = sHTML & ofv.GenRow("","","","","","","",celdas,"si")

    sInput1 = ofv.GenerarInput("atributo",sid2,"hidden","","1","dt","")
    sInput2 = ofv.GenerarInput("alias",sid,"hidden","","1","dt","")
    sInput3 = ofv.GenerarInput("depen",sdepen,"hidden","","1","dt","")
    sInput4 = ofv.GenerarInput("comprobar","1","hidden","","1","dt","")
    sInput5 = ofv.GenerarInput("aplicacion",sapli,"hidden","","1","dt","")
    celdas =  ofv.GenCelda("","aaa colspan=2","","","","","",sInput1 & sInput2 & sInput3 & sInput4 & sInput5,"si")

  sHTML = sHTML & ofv.GenRow("","","","","","","",celdas,"si")



  sHTML = sHTML & ofv.GenFTabla()

  sHTML = sHTML & ofv.GenFForm()


    sHTML = sHTML & OFV.GenForm("Inises", "", replace(volver2,"_","&"), "", "post", "", "", "no")
    sHTML = sHTML & ofv.GenTabla("panel","","","40%","",0,0,0,0,"","","","no")
    Celda = OFV.GenCelda("", "at style='color:#336699;'", "center", "80%", "20", "", "", "<input class=bt type=button onclick='Gener();' size=" & OFV.wsize & " name=boton value=Aceptar>", "si")
    Celda = celda & OFV.GenCelda("", "at style='color:#336699;'", "center", "80%", "20", "", "", "<input class=bt type=submit size=" & OFV.wsize & " name=botonC value=Cancelar>", "si")
    sHTML = sHTML & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
    sHTML = sHTML & ofv.GenFTabla()

    sHTML = sHTML & ofv.GenFForm()


  sHTML = sHTML & "</center></body></html>"

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
end if

response.write sHTML

%>