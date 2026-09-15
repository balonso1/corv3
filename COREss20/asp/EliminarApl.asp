<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

response.expires=0

ID = request.querystring("ID")
alias = request.querystring("alias")
aplic = request.querystring("aplic")
Descripcion = request.querystring("Descripcion")
volver = request.querystring("volver")
sEliminar = request.querystring("Eliminar")

descripcionD = replace(descripcion,"_"," ")

volver2 = volver

volver = replace(volver,"_","&")

'if request.form.count = 11 then
'  sEliminar = request.form(1)
'  if sEliminar = "Eliminar" then
'    sEliminar = "SI"
'  elseif sEliminar = "Cancelar" then
'    sEliminar = "NO"
'  end if  
'end if

if sEliminar = "NSNC" then

    sMensaje = "Esta seguro que desea eliminar " & DescripcionD & " y sus dependientes? "
    sVolver = "EliminarApl.asp?ID=" & ID & "&Descripcion=" & Descripcion & "&alias="
    sVolver = sVolver & alias & "&aplic=" & aplic & "&volver=" & volver2
    sHTML = ofv.Mensajes(sMensaje,svolver,"2")
    response.write sHTML
elseif sEliminar = "SI" then

        
       if alias = aplic then
' elimina toda la aplicacion
          set cn = ofv.conectar(ofv.strconn0)
              strsql = "Delete From atrm Where aplicacion = '" & aplic & "'"
              call ofv.crearconsultaEx(StrSql,cn,1,volver)
              set errores = cn.errors
          call ofv.cerrarconn(cn)
          set cn = ofv.conectar(ofv.strconn1)
              strsql = "Delete From coraplic Where codapli = '" & aplic & "'"
              call ofv.crearconsultaEx(StrSql,cn,1,volver)
              set errores = cn.errors

              strsql = "Delete From menux Where mcodapli = '" & aplic & "'"
              call ofv.crearconsultaEx(StrSql,cn,1,volver)
              set errores = cn.errors
          call ofv.cerrarconn(cn)
       else

              dim malias()
              dim aalias()

              redim malias(0)
              redim aalias(0)
              aalias(0) = alias
              malias(0) = alias

              xxi = 0
              xxia = 0
              set cn = ofv.conectar(ofv.strconn1)
              strsql = "select * From menux Where mdepen = '" & alias & "' and mcodapli = '" & aplic & "'" 
              set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
              do until rs.eof
                 xxi = xxi + 1 
                 xxia = xxia + 1 
                 redim preserve aalias(xxia)
                 redim preserve malias(xxi)
                 aalias(xxia) = rs("menuatr")
                 malias(xxi) = rs("menuatr")


                 set cn2 = ofv.conectar(ofv.strconn1)
                 strsql = "select * From menux Where mdepen = '" & alias & "' and mcodapli = '" & aplic & "'" 
                 set rs2 = ofv.crearconsultaEx(StrSql,cn2,1,volver)
                 do until rs2.eof
                    xxi = xxi + 1 
                    xxia = xxia + 1 
                    redim preserve aalias(xxia)
                    redim preserve malias(xxi)
                    aalias(xxia) = rs2("menuatr")
                    malias(xxi) = rs2("menuatr")

                    set cn3 = ofv.conectar(ofv.strconn0)
                    strsql = "select * From atrm Where depen = '" & alias & "' and aplicacion = '" & aplic & "'" 
                    set rs3 = ofv.crearconsultaEx(StrSql,cn3,1,volver)
                    do until rs3.eof
                       xxia = xxia + 1 
                       redim preserve aalias(xxia)
                       aalias(xxia) = rs3("alias")
                       rs3.movenext
                    loop
                    call ofv.cerrarconsulta(rs3) 
                    call ofv.cerrarconn(cn3)



                    rs2.movenext
                 loop
                 call ofv.cerrarconsulta(rs2) 
                 call ofv.cerrarconn(cn2)

                 set cn2 = ofv.conectar(ofv.strconn0)
                 strsql = "select * From atrm Where depen = '" & alias & "' and aplicacion = '" & aplic & "'" 
                 set rs2 = ofv.crearconsultaEx(StrSql,cn2,1,volver)
                 do until rs2.eof
                    xxia = xxia + 1 
                    redim preserve aalias(xxia)
                    aalias(xxia) = rs2("alias")
                    rs2.movenext
                 loop
                 call ofv.cerrarconsulta(rs2) 
                 call ofv.cerrarconn(cn2)


                 rs.movenext
              loop
              call ofv.cerrarconsulta(rs) 
              call ofv.cerrarconn(cn)

              set cn = ofv.conectar(ofv.strconn0)
              for xxi=0 to ubound(aalias)
                  strsql = "Delete From atrm Where alias = '" & aalias(xxi) & "' and aplicacion = '" & aplic & "'" 
                  call ofv.crearconsultaEx(StrSql,cn,1,volver)
                  set errores = cn.errors

              next

          
              call ofv.cerrarconn(cn)

              set cn = ofv.conectar(ofv.strconn1)
              for xxi=0 to ubound(malias)
                  strsql = "Delete From menux Where menuatr = '" & malias(xxi) & "' and mcodapli = '" & aplic & "'" 
                  call ofv.crearconsultaEx(StrSql,cn,1,volver)
                  set errores = cn.errors

              next

              call ofv.cerrarconn(cn)
     end if  
 

  if errores.count <> 0 then
     Response.write ofv.ErrorDB(cn,volver)
  ELSE
     response.redirect volver
  END IF
elseif sEliminar = "NO" then
  response.redirect Volver
end if



%>