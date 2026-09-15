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

sID2=sid
on error resume next

set cn = ofv.conectar(ofv.strconn2)

sMsg = ""
'Obtengo tipo de Clave
strsql = "Select * From " & tabla
response.write strsql
set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
clave = rs(nCampo).name
s = ofv.QueEsx(rs(nCampo).type)
if rs(nCampo).type = 200 then s = "'"
sID2 = s & ofv.convertircar(sID,"_"," ") & s
call ofv.cerrarconsulta(rs)

'Validaciones
Tabla=UCASE(TABLA)

select case Tabla
  case "GRUPOS"
   strsql = "Select * From Usuarios Where grupo = '" & ID & "'"
          set rsx = ofv.crearconsultaEx(strsql,cn,1,volver)
          if not rsx.eof then
            sMsg = "NO"
          end if
       call ofv.cerrarconsulta(rsx)
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
    
     select case tabla
       case "GRUPOS"
        
         
        strsql = "Delete From atrmGRUPOs Where GRUPO = '" & SID & "'"
	response.write strsql
	 call ofv.crearconsultaEx(StrSql,cn,1,volver)
	 set errores = cn.errors
         call ofv.cerrarconsulta(rs)
         
         strsql = "Delete From GRUPOS Where id = '" & SID & "'"
	 	'response.write strsql	
	 call ofv.crearconsultaEx(StrSql,cn,1,volver)
	 set errores = cn.errors
         call ofv.cerrarconsulta(rs)
         
         Volver = "../gruposx.asp?seg=grupos"
      CASE "USUARIOS"
      
      strsql = "Delete From AtrmUSUARIOS Where USUARIO = '" & SID & "'"
      call ofv.crearconsultaEx(StrSql,cn,1,volver)
      set errores = cn.errors
      call ofv.cerrarconsulta(rs)

      strsql = "Delete From Histpwd Where USUARIO = '" & SID & "'"
      call ofv.crearconsultaEx(StrSql,cn,1,volver)
      set errores = cn.errors
      call ofv.cerrarconsulta(rs)
      
      strsql = "Delete From usuarios Where id = '" & SID & "'"
       call ofv.crearconsultaEx(StrSql,cn,1,volver)
       set errores = cn.errors
      call ofv.cerrarconsulta(rs)
      Volver = "../usuariosx.asp?seg=usuarios"
 
              
     end select
      'strsql = "Delete From " & Tabla & " Where " & clave & "=" & sID2
      'RESPONSE.WRITE STRSQL
      'call ofv.crearconsultaEx(StrSql,cn,1,volver)
    end if
  else
    dato = "../Mensajes3.asp?msg=" & ofv.convertircar(sMsg," ","_") & "&volver="
    dato = dato & "asp/Eliminar.asp?ID=" & ofv.convertircar(sID," ","_")
    dato = dato & "&Botones=1" & "&Tabla=" & Tabla & "&Nombre=" & sNombre
    dato = dato & "&Campo=" & nCampo & chr(13)
    response.redirect dato
  end if
end if

'Volver = session("volver2")
set errores = cn.errors
if errores.count = 0 then
  response.redirect volver
else
  call ofv.ErrorDB(cn,volver)
end if 

call ofv.cerrarconn(cn)


%>