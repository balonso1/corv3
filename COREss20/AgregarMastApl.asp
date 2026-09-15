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
  Volver = request.querystring("Volver")
  Volver2 = replace(volver,"../","")

  set cn = ofv.conectar(ofv.strconn1)

  sid = 0
  sid2 = 0
  sor = 0
  strsql = "Select max(ID) From " & tabla 
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
  if not rs.eof then sid = gtkey(clng(rs(0)) + 1)
  call ofv.cerrarconsulta(rs)

  sid2 = "CLI-" & cstr(sid)
  sid = "CLI" & cstr(sid)


  strsql = "Select max(ordenapli) From " & tabla 
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
  if not rs.eof then sor = clng(rs(0)) + 1
  call ofv.cerrarconsulta(rs)


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
'sHTML = sHTML &  "  document.all.AgregarMastatrm.submit();" & chr(13)
sHTML = sHTML &  " }" & chr(13)
sHTML = sHTML & "</script>" &  chr(13) 



    f1a = "asp/AgregarMastApl2.asp?Tabla=" & Tabla
    f1a = f1a & "&volver=" & Volver & "&codapli=" & sid & "&alias=" & sid  & "&ATRM=" & sid2

  sHTML = sHTML & "<center>" & ofv.GenForm("AgregarMastApl","",f1a,"","post","","","no")


  sHTML = sHTML & ofv.GenTabla("panel","","","80%","",0,0,0,0,"","","","no")

  celdas = ofv.GenCelda("","cc colspan=2","","","","","","Agregar Aplicacion","si")

  sHTML = sHTML & ofv.GenRow("","","","","","","",celdas,"si")

  celdas = ofv.GenCelda("","aaa colspan=2","","","","","","&nbsp;","si")

  sHTML = sHTML & ofv.GenRow("","","","","","","",celdas,"si")

  'Encabezado Accion
    celdas = ofv.GenCelda("","tt","","2%","","","","Aplicacion: ","si")
    sInput = ofv.GenerarInput("descapli","","agregar","","100","dt","")
    celdas = celdas & ofv.GenCelda("","cc","","","","","",sInput,"si")

  sHTML = sHTML & ofv.GenRow("","","","","","","",celdas,"si")

    sInput1 = ofv.GenerarInput("codapli",sid,"hidden","","1","dt","")
    sInput2 = ofv.GenerarInput("enable","1","hidden","","1","dt","")
    sInput3 = ofv.GenerarInput("modoapli","C","hidden","","1","dt","")
    sInput4 = ofv.GenerarInput("Iconoapli","Icoress.jpg","hidden","","1","dt","")
    sInput5 = ofv.GenerarInput("dirapli","CorClApl","hidden","","1","dt","")
    sInput6 = ofv.GenerarInput("imagenapli","Coressuite20.jpg","hidden","","1","dt","")
    sInput7 = ofv.GenerarInput("Corpapli","0","hidden","","1","dt","")
    sInput8 = ofv.GenerarInput("targapli","Menus","hidden","","1","dt","")
    sInput9 = ofv.GenerarInput("Formapli","Menus.asp","hidden","","1","dt","")
    sInputA = ofv.GenerarInput("titapli",sid,"hidden","","1","dt","")
    sInputB = ofv.GenerarInput("Ordenapli",sor,"hidden","","1","dt","")
    celdas =  ofv.GenCelda("","aaa colspan=2","","","","","",sInput1 & sInput2 & sInput3 & sInput4 & sInput5 & sInput6 & sInput7 & sInput8 & sInput9 & sInputA & sInputB,"si")

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