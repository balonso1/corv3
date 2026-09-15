<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
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



'RESPONSE.WRITE request.form & " " & request.form.count
if sEliminar = "NSNC" then
  if ID = "DEF" and Tabla = "Grupos" then
    sMensaje = "No puede eliminar " & Descripcion & " porque es el grupo predeterminado"
    sVolver = "gruposx.asp?seg=grupos"
    sHTML = Mensajes(sMensaje,svolver,"1")
    response.write sHTML
  else
    sMsg = ""
    if Tabla = "Grupos" then
       vvuelta = "gruposx.asp?seg=grupos"

       strsql = "Select * From Usuarios Where grupo = '" & ID & "'"
       set rsx = ofv.crearconsultaEx(strsql,cn,1,volver)
       if not rsx.eof then
         sMsg = "No puede eliminar " & ID & " porque tiene usuarios dependientes"
       end if
       call ofv.cerrarconsulta(rsx)
    else
       vvuelta = "usuariosx.asp?seg=usuarios"
    end if
    sMensaje = "Esta seguro que desea eliminar " & Descripcion
    sVolver = "Eliminarusrgrp.asp?ID=" & ID & "&Descripcion=" & Descripcion & "&CantRegAMover="
    sVolver = sVolver & CantRegAMover & "&Tabla=" & Tabla & "&volver=" & vvuelta
    
    if sMsg = "" then
      sHTML = Mensajes(sMensaje,svolver,"2")
    else
    if Tabla = "Grupos" then
      sVolver = "gruposx.asp?seg=grupos"
    else
      sVolver = "usuariosx.asp?seg=usuarios"
    end if  

      sHTML = Mensajes(sMsg,svolver,"1")
    end if
    response.write sHTML
  end if

elseif sEliminar = "SI" then
  select case Ucase(tabla)
         case "GRUPOS"
              strsql = "Delete From atrmGRUPOs Where GRUPO = '" & ID & "'"
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

call ofv.cerrarconn(cn)


 Function Mensajes(sMsg, sVolver, nBotones)

''RESPONSE.WRITE request.form & " " & request.form.count
    
'    Dim s$, sInput$, Celda$, Fila$, Tabla$, Formulario$
    
     nBotones = CInt(nBotones)

   
     s = ofv.FormHeader("Atencion")

     sInput = "<font color=#336699 face=verdana size=1><b>" & sMsg & "</b></font>"
     Celda = ofv.GenCelda("", "", "center", "100%", "", "", "", sInput, "si")
    Fila = ofv.GenRow("", "", "", "", "", "", "", Celda, "si")
     Tabla = s & ofv.GenTabla("", "", "", "100%", "", "0", "", "0", "0", "", "", Fila, "si")
     Tabla = Tabla & "<hr size=1 color=#000099>"



     If nBotones = 1 Then
         '  sInput = GenerarInput("", "Volver", "submit", "", "10", "bt", "")
         '  Celda = GenCelda("", "", "center", "50%", "", "", "", sInput, "si")
         '  Formulario = GenForm("Mensajes", "", sVolver, "", "post", "", Celda, "si")

        Formulario = ofv.GenBoton(7, "MenBot", "BT", "S", "", "S", "S", "N", "MenFrm", sVolver & "")

  Else
          ' sInput = GenerarInput("", "Eliminar", "submit", "", "10", "bt", "")
          ' Celda = GenCelda("", "", "center", "50%", "", "", "", sInput, "si")
          ' Formulario = GenForm("Mensajes", "", sVolver & "&Eliminar=SI", "", "post", "", Celda, "si")

       Formulario = ofv.GenBoton(15, "MensBot", "BT", "S", "", "S", "S", "N", "MensFrm", sVolver & "&Eliminar=SI")
    
          ' sInput = GenerarInput("", "Cancelar", "submit", "", "10", "bt", "")
          ' Celda = GenCelda("", "", "center", "50%", "", "", "", sInput, "si")
          ' Formulario = Formulario & GenForm("Mensajesb", "", sVolver & "&Eliminar=NO", "", "post", "", Celda, "si")

       Formulario = Formulario & ofv.GenBoton(14, "MennBot", "BT", "S", "", "S", "S", "N", "MennFrm", sVolver & "&Eliminar=NO")

    End If
    Fila = ofv.GenRow("", "", "", "", "", "", "", Formulario, "si")

    Tabla = Tabla & "<br><br><center>" & ofv.GenTabla("", "", "", "", "", "0", "", "0", "0", "", "", Fila, "si") & "</center></body></html>"
    
    Mensajes = Tabla

 end Function

%>