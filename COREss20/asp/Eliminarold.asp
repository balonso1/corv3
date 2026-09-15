<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

response.expires=0

ID = request.querystring("ID")
Descripcion = request.querystring("Descripcion")
Tabla = request.querystring("Tabla")
volver = request.querystring("volver")
sEliminar = request.querystring("Eliminar")
CantRegAMover = cdbl(request.querystring("CantRegAMover"))

set cn = ofv.conectar(ofv.strconn0)
'response.write  request.form

 
'sEliminar="SI"
if sEliminar = "NSNC" then
  if ID = "DEF" and Tabla = "Grupos" then
    sMensaje = "No puede eliminar " & Descripcion & " porque es el grupo predeterminado"
    sVolver = "../gruposx.asp?seg=grupos"
    sHTML = ofv.Mensajes(sMensaje,svolver,"1")
    
  else
    sMsg = ""
    

    if Tabla = "Grupos" then
       vvuelta = "../gruposx.asp?seg=grupos"
       ' response.write vvuelta
       strsql = "Select * From Usuarios Where grupo = '" & ID & "'"
       set rsx = ofv.crearconsultaEx(strsql,cn,1,volver)
       if not rsx.eof then
         sMsg = "No puede eliminar " & ID & " porque tiene usuarios dependientes"
       end if
       call ofv.cerrarconsulta(rsx)
    else
       vvuelta = "../usuariosx.asp?seg=usuarios"
    end if
    sMensaje = "Esta seguro que desea eliminar " & Descripcion
    sVolver = "Eliminar.asp?ID=" & ID & "&Descripcion=" & Descripcion & "&CantRegAMover="
    sVolver = sVolver & CantRegAMover & "&Tabla=" & Tabla & "&volver=" & vvuelta
    
    if sMsg = "" then
      sHTML = ofv.Mensajes(sMensaje,svolver,"2")
    else
      sVolver = "../gruposx.asp?seg=grupos"
      sHTML = ofv.Mensajes(sMsg,svolver,"1")
    end if
    response.write sHTML
  end if

else
 
 if sEliminar = "SI" then
  select case Ucase(tabla)
         case "GRUPOS"
              strsql = "Delete From atrmGRUPOs Where GRUPO = '" & ID & "'"
               response.write strsql
              call ofv.crearconsultaEx(StrSql,cn,1,volver)
              set errores = cn.errors
              call ofv.cerrarconsulta(rs)
         CASE "USUARIOS"
              strsql = "Delete From AtrmUSUARIOS Where USUARIO = '" & ID & "'"
              call ofv.crearconsultaEx(StrSql,cn,1,volver)
              set errores = cn.errors
              call ofv.cerrarconsulta(rs)

              strsql = "Delete From Histpwd Where USUARIO = '" & ID & "'"
              call ofv.crearconsultaEx(StrSql,cn,1,volver)
              set errores = cn.errors
              call ofv.cerrarconsulta(rs)
  end select

  if errores.count <> 0 then
     Response.write ofv.ErrorDB(cn,volver)
  ELSE

  strsql = "Delete From " & Tabla & " Where ID = '" & ID & "'"
  
  'response.write strsql
  call ofv.crearconsultaEx(StrSql,cn,1,volver)
  set errores = cn.errors
  if errores.count = 0 then
    response.redirect volver
  else
    Response.write ofv.ErrorDB(cn,volver)
  end if 
  call ofv.cerrarconsulta(rs)
  END IF
elseif sEliminar = "NO" then
  response.redirect Volver
end if
end if
'response.write tabla
call ofv.cerrarconn(cn)

%>