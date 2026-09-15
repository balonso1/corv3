<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

response.expires=0

Tabla = request.querystring("Tabla")
nCampo = request.querystring("Campo")
if nCampo = "" then nCampo = 0
nCampo = CDbl(nCampo)
sID = request.querystring("ID")

sEliminar = request.querystring("Eliminar")
sNombre = request.querystring("Nombre")
Volver = request.querystring("Volver")
Volver = ofv.convertircar(Volver,"_","&")
if sEliminar = "NSNC" then session("volver2") = volver

on error resume next

set cn = ofv.conectar(ofv.strconn6)

sMsg = ""
'Obtengo tipo de Clave
strsql = "Select * From " & tabla
set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
clave = rs(nCampo).name
call ofv.cerrarconsulta(rs)

sID2 = sID

'Validaciones
select case Tabla
  case "PROYFUNCIONES"
    strsql = "select * from PROYRESP Where FUNCION=" & sID2
    
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    If not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)


  case "PROYSUBTIPOTSK"
    strsql = "Select * From PROYTAREAS WHERE TIPO = '" & sID2 & "' "
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    if not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)

  case "PROYTIPOTSK"
    strsql = "Select * From PROYTAREAS WHERE PROYTSKTIPO=" & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    if not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)
    

  case "TIPOS"
    strsql = "Select * From REQUERIMIENTO Where TIPO=" & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    if not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)

    strsql = "Select * From PROYECTO Where TIPO=" & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    if not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)

  case "TIPOREQ"

    strsql = "Select * From TIPOREQ Where ID="  & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    if not rs.eof then SIDNN = RS(1)
    call ofv.cerrarconsulta(rs)

    strsql = "Select * From REQUERIMIENTO Where TIPOREQ="  & SIDNN

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
      dato = dato & "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(sID," ","_")
      dato = dato & "&Botones=2" & "&Tabla=" & Tabla & "&Nombre=" & ofv.convertircar(sNombre," ","_")
      dato = dato & "&Campo=" & nCampo & chr(13)
      response.redirect dato
    elseif sEliminar = "SI" then
           select case Ucase(tabla)
                  case "PROYTIPOTSK"

                       strsql = "Delete From PROYSUBTIPOTSK Where TIPO = " & sID2
                       call ofv.crearconsultaEx(StrSql,cn,1,volver)

                  case "TIPOREQ"

                       strsql = "Delete From TIPOS Where TIPOREQ = " & sID2
                       call ofv.crearconsultaEx(StrSql,cn,1,volver)

           END SELECT
      strsql = "Delete From " & Tabla & " Where " & clave & "=" & sID2
   '   RESPONSE.WRITE STRSQL & "<BR>"
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