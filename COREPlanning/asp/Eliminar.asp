<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

response.expires=0

Tabla = request.querystring("Tabla")
nCampo = request.querystring("Campo")
if nCampo = "" then nCampo = 0
nCampo = CDbl(nCampo)
sID = request.querystring("ID")
sID = ofv.convertircar(sID,"_"," ")
sEliminar = request.querystring("Eliminar")
sNombre = request.querystring("Nombre")
Volver = request.querystring("Volver")
Volver = ofv.convertircar(Volver,"_","&")
if sEliminar = "NSNC" then session("volver2") = volver

set cn = ofv.conectar(ofv.strconn4)

sMsg = ""
'Obtengo tipo de Clave
strsql = "Select * From " & tabla
set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
clave = rs(nCampo).name
s = ofv.QueEsx(rs(nCampo).type)
sID2 = s & ofv.convertircar(sID,"_"," ") & s
call ofv.cerrarconsulta(rs)

'Validaciones
select case Tabla
  case "Asociacion"
    strsql = "select * from NIVELES Where codasoc =" & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    if not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)
    if sMsg <> "NO" then
      strsql = "select * from TIPOCONT Where codasoc =" & sID2
      set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
      if not rs.eof then sMsg = "NO"
      call ofv.cerrarconsulta(rs)
    end if
  case "Niveles"
    strsql = "select * from CIRCORR Where codnivel =" & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    if not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)
  case "Tipocont"
    strsql = "select * from CONTINGENCIA Where Tipocontin =" & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    if not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)
    if sMsg <> "NO" then
      strsql = "select * from ACCIONALT Where disparador =" & sID2
      set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
      if not rs.eof then sMsg = "NO"
      call ofv.cerrarconsulta(rs)
    end if
  case "Eco"
    strsql = "select * from CIRCALT Where eco =" & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    if not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)
    if sMsg <> "NO" then
      strsql = "select * from ACCIONALT Where ecc =" & sID2
      set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
      if not rs.eof then sMsg = "NO"
      call ofv.cerrarconsulta(rs)
    end if
  case "Tipoaccalt"
    strsql = "select * from ACCIONALT Where tipoaccionalt =" & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    if not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)
  case "Tipoacccorr"
    strsql = "select * from ACCIONCORR Where Tipoaccioncorr =" & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    if not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)
  case "IAR"
    strsql = "select * from ACCIONALT Where CodIar =" & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    if not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)
  case "Infrespcorr"
    strsql = "select * from CIRCORR Where (Codinfo1 =" & sID2
    For i = 2 to 4
      Strsql = Strsql & ") or (Codinfo" & i & "=" & sID2
    next
    Strsql = Strsql & ")"
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    if not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)
  case "Provcorr"
    strsql = "select * from CIRCORR Where Codprovprinc =" & sID2 & " or Codprovalt =" & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    if not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)
  case "Equipcorr"
    strsql = "select * from CIRCORR Where Codequiprin =" & sID2 & " or Codequialt =" & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    if not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)
  case "Cenproccorr"
    strsql = "select * from CIRCORR Where codcentalt =" & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    if not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)
  case "Circorr"
    strsql="select * from ACCIONCORR Where codcircorr =" & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    if not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)
  case "Circalt"
    strsql="select * from ACCIONALT Where codcircalt =" & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    if not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)
  case "Contingencia"
    strsql="select * from CIRCALT Where codcontin =" & sID2 
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    if not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)
  case "Proceso"
    strsql="select * from CONTINGENCIA Where codproceso =" & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    if not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)
  case "Evento"
    strsql="select * from PROCESO Where codevento =" & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    if not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)
  case "Prodserv"
    strsql="select * from EVENTO Where codprodserv =" & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    if not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)
  case "Tema"
    strsql = "select * from PRODSERV Where codtema =" & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    if not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)
end select
if sMsg = "NO" then sMsg = "No puede eliminar " & sNombre & " porque esta siendo utilizado"

if sEliminar <> "NO" then
  if sMsg = "" then
    if sEliminar = "NSNC" then
      sMensaje = "Esta seguro que desea eliminar " & sNombre
      dato = "../Mensajes3.asp?msg=" & ofv.convertircar(sMensaje," ","_") & "&volver="
      dato = dato & "asp/Eliminar.asp?ID=" & ofv.convertircar(sID," ","_")
      dato = dato & "&Botones=2" & "&Tabla=" & Tabla & "&Nombre=" & sNombre
      dato = dato & "&Campo=" & nCampo & chr(13)
      response.redirect dato
    elseif sEliminar = "SI" then
      strsql = "Delete From " & Tabla & " Where " & clave & "=" & sID2
      call ofv.crearconsultaEx(StrSql,cn,1,volver)
    end if
  else
    dato = "../Mensajes3.asp?msg=" & ofv.convertircar(sMsg," ","_") & "&volver="
    dato = dato & "asp/Eliminar.asp?ID=" & ofv.convertircar(sID," ","_")
    dato = dato & "&Botones=1" & "&Tabla=" & Tabla & "&Nombre=" & sNombre
    dato = dato & "&Campo=" & nCampo & chr(13)
    response.redirect dato
  end if
end if

Volver = session("volver2")
Volver = ofv.convertircar(Volver," ","*")
set errores = cn.errors

if errores.count = 0 then
  response.redirect volver
else
  response.write ofv.ErrorDB(cn,volver)
end if 

call ofv.cerrarconn(cn)

%>