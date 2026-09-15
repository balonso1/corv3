<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

response.expires=0

' response.write "aaaaaaaaaa"

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

on error resume next

set cn = ofv.conectar(ofv.strconn2)

sMsg = ""
'Obtengo tipo de Clave
strsql = "Select * From " & tabla
set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
clave = rs(nCampo).name
s = ofv.QueEsx(rs(nCampo).type)
if rs(nCampo).type = 200 then s = "'"
sID2 = s & ofv.convertircar(sID,"_"," ") & s
call ofv.cerrarconsulta(rs)

'Validaciones
select case ucase(Tabla)
  case ucase("clienteplan")
    strsql = "select * from objetos Where objun = " & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    If not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)



end select

if sMsg = "NO" then sMsg = "No puede eliminar " & sNombre & " porque esta siendo utilizado"

if sEliminar <> "NO" then
  if sMsg = "" then
    if sEliminar = "NSNC" then
      sMensaje = "Esta seguro que desea eliminar " & sNombre
      dato = "../Mensajes3.asp?msg=" & ofv.convertircar(sMensaje," ","_") & "&volver="
      dato = dato & "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(sID," ","_")
      dato = dato & "&Botones=2" & "&Tabla=" & Tabla & "&Nombre=" & sNombre
      dato = dato & "&Campo=" & nCampo & chr(13)
      response.redirect dato
    elseif sEliminar = "SI" then
      strsql = "Delete From " & Tabla & " Where " & clave & "=" & sID2
      call ofv.crearconsultaEx(StrSql,cn,1,volver)
    end if
  else
    dato = "../Mensajes3.asp?msg=" & ofv.convertircar(sMsg," ","_") & "&volver="
    dato = dato & "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(sID," ","_")
    dato = dato & "&Botones=1" & "&Tabla=" & Tabla & "&Nombre=" & sNombre
    dato = dato & "&Campo=" & nCampo & chr(13)
    response.redirect dato
  end if
end if

Volver = session("volver2")
set errores = cn.errors
if errores.count = 0 then
  response.redirect volver
else
  call ofv.ErrorDB(cn,volver)
end if 

call ofv.cerrarconn(cn)

%>