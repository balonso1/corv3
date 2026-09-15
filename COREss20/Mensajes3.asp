<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

response.expires=0

sMsg = ofv.convertircar(request.querystring("msg"),"_"," ")
'''sMsg = request.querystring("msg")
sVolver = request.querystring("volver")
nCampo = request.querystring("Campo")
sNombre = request.querystring("Nombre")
nAccion = request.querystring("Accion")
nProced = request.querystring("Codproced")
vuelta = request.querystring("vuelta")
nCodigo = request.querystring("Codobjeto")
sTipo = request.querystring("Tipobjeto")
Tabla = request.querystring("Tabla")
nBotones = CInt(request.querystring("Botones"))

'''sHTML = ofv.FormHeader("Atencion")

'sMsg=request.querystring
'''response.write "<br><br>" & sMsg & "<br>"  & sVolver & "<br>" & ncodigo & "-" & stipo & "<br>"
select case Tabla
  case "Tipoperfil", "Eco", "Tipoaccalt", "Sisgen", "Tipogen", "Origen", "Instrumen", "Objcuenta", "Niveljerar", "Segmento", "Jerarquia"
    sVolver = sVolver & "&Nombre=" & sNombre & "&Tabla=" & Tabla & "&Campo=" & nCampo
  case "Objaccioncontrol"
    sVolver = sVolver & "&Codproced=" & nProced & "&Codobjeto=" & nCodigo & "&Accion="
    sVolver = sVolver & nAccion & "&Tipobjeto=" & sTipo & "&vuelta=" & vuelta & "&Tabla="
    sVolver = sVolver & Tabla
  case else
    if right(sVolver,3) = "asp" then
      sVolver = sVolver & "?Tipobjeto=" & sTipo & "&Codobjeto=" & nCodigo & "&Tabla=" & Tabla
    else
      sVolver = sVolver & "&Tipobjeto=" & sTipo & "&Codobjeto=" & nCodigo & "&Tabla=" & Tabla
    end if
end select




 tabla = ofv.mensajes(sMsg,sVolver,nBotones)

 tablan = instr(tabla,"&Eliminar=")

 if tablan = 0 then
    tabla = replace(tabla,"asp?","asp?Eliminar=NO&")
 end if

 tabla = replace(tabla,"<body ","<body onunload='shwhdemenu();' ")
 tabla = replace(tabla,"</body></html>","")

sHTML = sHTML & Tabla & "<script language=javascript>shwhdemenu();</script>"

sHTML = sHTML & "</body>"
sHTML = sHTML & "</html>"

response.write sHTML

%>