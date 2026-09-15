<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

response.expires=0

id = request.querystring("fname")
id = ofv.convertircar(id,"*"," ")  
tabla = request.querystring("tabla")
volver = request.querystring("volver")
volver = ofv.convertircar(volver,"_","&")
volver = ofv.convertircar(volver," ","*")

set cn = ofv.conectar(ofv.strconn4)

strsql = "Select * From " & tabla
set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

strsql = "Update " & tabla & " Set "

For Each x In Request.form
  vNuevoValor = UCase(CStr(ofv.convertircar(Request.Form(x),"_"," ")))
  vNuevoValor = replace(vNuevoValor,"'","")
  vNuevoValor = replace(vNuevoValor,chr(34),"")
  s = ofv.QueEsx(rs(x).type)
  strsql = strsql & x  & " = " & s & vNuevoValor & s & ","
Next 
 
s = ofv.QueEsx(rs(0).type)
strsql = mid(strsql,1,len(strsql)-1) & " Where " & rs(0).name & "="
strsql = strsql & s & id & s

call ofv.crearconsultaEx(StrSql,cn,1,volver)

if cn.errors.count = 0 then
  if volver <> "" then response.redirect volver
else
  response.write ofv.ErrorDB(cn,volver)
end if 

call ofv.cerrarconsulta(rs)
call ofv.cerrarconn(cn)

%>